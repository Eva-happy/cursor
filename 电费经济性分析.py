import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from matplotlib.widgets import Button
import time
import numpy as np
import os
import traceback
from matplotlib import font_manager
import glob
# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
# 定义变压器容量(kVA)
transformer_capacity = 8050
# 客户选择的变压器基本容量费用计算方法(1或2)
# 1:按容量收取 2:按需收取
method_basic_capacity_cost_transformer = 1
# 存储累积统计数据的全局变量
monthly_statistics = {}
current_month_start_index = 0
# 储能系统参数
max_power_per_system = 125  # kW
storage_capacity_per_system = 261  # kWh
efficiency_bess = 0.97  # 储能系统的充放电效率
num_storage_systems = 1
initial_storage_capacity = 30.00  # kWh
# 定义时段类型常量 (便于阅读)
SHARP = 1
PEAK = 2
FLAT = 3
VALLEY = 4
DEEP_VALLEY = 5
# 辅助函数：查找连续窗口
def find_continuous_window(periods, target_types, min_duration, start_hour=0):
    """查找第一个满足条件的连续时间窗口"""
    n = len(periods)
    current_start = -1
    count = 0
    for i in range(start_hour, n):
        if periods[i] in target_types:
            if count == 0:
                current_start = i
            count += 1
            if count >= min_duration:
                # 找到满足最短持续时间的窗口
                # 现在检查这个窗口后面连续同类型时段的总长度
                total_duration = count
                for j in range(i + 1, n):
                    if periods[j] in target_types:
                        total_duration += 1
                    else:
                        break
                return current_start, min_duration, total_duration # 返回开始时间, 要求的最短时间, 实际连续总时间
        else:
            count = 0
            current_start = -1
    return None, None, None # 未找到
# 辅助函数：获取时段名称
def get_period_name(period_type):
    if period_type == SHARP: return "尖"
    if period_type == PEAK: return "峰"
    if period_type == FLAT: return "平"
    if period_type == VALLEY or period_type == DEEP_VALLEY: return "谷"
    return "未知"
def load_data(load_file, price_config_file):
    """
    加载负荷和电价数据
    load_file: 负载数据Excel文件路径（包含date, min, load三列）
    price_config_file: 电价配置Excel文件路径
    """
    try:
        # 读取负载数据
        load_df = pd.read_excel(load_file)
        
        # 确保列名正确
        if 'date' not in load_df.columns or 'load' not in load_df.columns:
            raise ValueError("负载数据文件必须包含'date', 'load'列")
        
        # 检查数据是否包含NaN值
        if load_df['load'].isna().any():
            print("警告：负载数据中包含无效值(NaN)，将使用0替换。")
            load_df['load'] = load_df['load'].fillna(0)
        
        # 处理时间列：支持'hour'列或'min'列(时:分格式)
        has_time_info = False
        
        if 'hour' in load_df.columns:
            # 原来的处理方式
            load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['hour'], unit='h')
            load_df['hour'] = load_df['datetime'].dt.hour
            has_time_info = True
            
        elif 'min' in load_df.columns:
            # 处理"时:分"格式
            # 检查min列的格式
            first_value = str(load_df['min'].iloc[0])
            if ':' in first_value:  # 如果是"时:分"格式
                print("检测到时:分格式的数据，正在处理...")
                
                # 将"时:分"拆分为小时和分钟
                def extract_hour_minute(time_str):
                    try:
                        time_str = str(time_str)
                        if ':' not in time_str:
                            return 0, 0
                        hour, minute = map(int, time_str.split(':'))
                        return hour, minute
                    except Exception:
                        print(f"警告：无法解析时间 '{time_str}'，使用0:00替代")
                        return 0, 0
                
                # 应用函数提取小时和分钟
                load_df[['hour', 'minute']] = load_df['min'].apply(lambda x: pd.Series(extract_hour_minute(x)))
                
                # 创建datetime列
                load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['hour'], unit='h') + pd.to_timedelta(load_df['minute'], unit='m')
                
                # 检查是否为15分钟间隔的数据
                unique_minutes = sorted(load_df['minute'].unique())
                if set(unique_minutes) == {0, 15, 30, 45} or len(unique_minutes) > 1:
                    print("检测到15分钟间隔数据，将转换为小时数据...")
                    
                    # 按日期和小时分组，计算每小时的平均负载
                    hourly_load = load_df.groupby([load_df['datetime'].dt.date, load_df['hour']])['load'].mean().reset_index()
                    hourly_load.columns = ['date', 'hour', 'load']
                    
                    # 重新创建datetime列
                    hourly_load['datetime'] = pd.to_datetime(hourly_load['date']) + pd.to_timedelta(hourly_load['hour'], unit='h')
                    
                    # 替换原始数据框
                    load_df = hourly_load
                    
                has_time_info = True
            else:  # 如果是数字格式(假设为小时)
                load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['min'], unit='h')
                load_df['hour'] = load_df['datetime'].dt.hour
                has_time_info = True
        
        # 如果没有时间信息，尝试从date列推断
        if not has_time_info:
            # 检查是否已经是datetime格式
            if pd.api.types.is_datetime64_any_dtype(load_df['date']):
                load_df['datetime'] = load_df['date']
                load_df['hour'] = load_df['datetime'].dt.hour
            else:
                # 尝试将date转换为datetime
                try:
                    load_df['datetime'] = pd.to_datetime(load_df['date'])
                    load_df['hour'] = load_df['datetime'].dt.hour
                except:
                    raise ValueError("无法从'date'列推断时间信息，请提供'hour'或'min'列")
        
        # 读取电价配置
        price_config = pd.read_excel(price_config_file, sheet_name=['时段配置', '电价配置'])
        
        # 获取时段配置
        time_periods = price_config['时段配置']
        price_rules = price_config['电价配置']
        
        # 添加月份、小时信息
        load_df['month'] = load_df['datetime'].dt.month
        
        # 重新赋值小时，确保使用正确的小时值
        load_df['hour'] = load_df['datetime'].dt.hour
        
        # 根据时段配置和电价规则计算每个时间点的电价和时段类型
        def get_price_and_period(row):
            month = row['month']
            hour = row['hour']
            try:
                # 从时段配置中获取时段类型（数字1-5）
                period_type = time_periods.loc[
                    (time_periods['月份'] == month),
                    f'{hour}-{hour+1}'
                ].values[0]
                
                # 根据时段类型查找对应电价
                price = price_rules.loc[
                    price_rules['时段类型'] == period_type,
                    '电价'
                ].values[0]
                
                return pd.Series({'price': price, 'period_type': period_type})
            except Exception as e:
                print(f"警告：无法获取月份{month}小时{hour}的电价和时段类型，使用默认值。")
                return pd.Series({'price': 0, 'period_type': 3})  # 使用默认值
        
        # 应用函数添加电价和时段类型列
        load_df[['price', 'period_type']] = load_df.apply(get_price_and_period, axis=1)
        
        # 检查电价是否包含NaN值
        if load_df['price'].isna().any():
            print("警告：电价数据中包含无效值(NaN)，将使用0替换。")
            load_df['price'] = load_df['price'].fillna(0)
        
        # 只保留需要的列
        result_df = load_df[['datetime', 'load', 'price', 'period_type', 'month', 'hour']]
        
        # 填补缺失的小时数据 - 确保每天都有24小时的数据
        # 获取数据中的所有日期
        all_dates = pd.to_datetime(result_df['datetime']).dt.date.unique()
        
        # 创建完整的时间索引（每个日期24小时）
        all_hours = range(24)
        complete_index = []
        for date in all_dates:
            for hour in all_hours:
                complete_index.append(pd.Timestamp(date) + pd.Timedelta(hours=hour))
        
        # 创建完整的DataFrame
        complete_df = pd.DataFrame(index=pd.DatetimeIndex(complete_index))
        complete_df.index.name = 'datetime'
        complete_df = complete_df.reset_index()
        
        # 合并原始数据与完整索引
        merged_df = pd.merge(complete_df, result_df, on='datetime', how='left')
        
        # 填充缺失值
        merged_df['load'] = merged_df['load'].fillna(0)
        
        # 填充其他列
        for col in ['month', 'hour']:
            if col in merged_df.columns and merged_df[col].isna().any():
                merged_df[col] = merged_df['datetime'].dt.__getattribute__(col[0] + col[1:])
        
        # 重新计算缺失的电价和时段类型
        missing_price = merged_df['price'].isna()
        if missing_price.any():
            merged_df.loc[missing_price, ['price', 'period_type']] = merged_df[missing_price].apply(get_price_and_period, axis=1)
        
        print(f"数据处理完成！共有 {len(all_dates)} 天的数据，每天24小时。")
        
        # 返回填补后的数据
        return merged_df
        
    except Exception as e:
        print(f"加载数据时出错: {e}")
        traceback.print_exc()
        # 返回一个空的DataFrame，避免程序崩溃
        return pd.DataFrame(columns=['datetime', 'load', 'price', 'period_type', 'month', 'hour'])
def plot_price_curve(data, month=None):
    """
    绘制电价曲线
    data: 包含datetime, price和period_type的DataFrame
    month: 指定要绘制哪个月份的数据，如果为None则绘制全年
    """
    # 定义时段类型对应的颜色和标签
    period_colors = {
        1: 'orange',    # 尖峰
        2: 'red',       # 高峰
        3: 'green',     # 平段
        4: 'blue',      # 低谷
        5: 'darkblue'   # 深谷
    }
    
    period_labels = {
        1: '尖峰时段',
        2: '高峰时段',
        3: '平段时段',
        4: '低谷时段',
        5: '深谷时段'
    }
    
    if month is not None:
        # 筛选指定月份的数据
        month_data = data[data['datetime'].dt.month == month]
        # 获取该月第一天的数据作为24小时电价曲线
        day_data = month_data[month_data['datetime'].dt.day == 1]
        
        # 创建24小时时间序列
        hours = range(24)
        prices = day_data['price'].values[:24]
        period_types = day_data['period_type'].values[:24]
        
        # 创建柱状图
        plt.figure(figsize=(12, 6))
        bars = plt.bar(hours, prices)
        
        # 根据时段类型设置不同的颜色
        for hour, bar, period_type in zip(hours, bars, period_types):
            bar.set_color(period_colors.get(period_type, 'gray'))
        
        # 创建图例元素
        legend_elements = [plt.Rectangle((0,0),1,1, color=color, label=label) 
                          for period_type, color in period_colors.items() 
                          for label in [period_labels.get(period_type, f'未知时段{period_type}')] 
                          if period_type in period_types]
        
        plt.legend(handles=legend_elements)
        plt.title(f'电价曲线，月份: {month}')
        plt.xlabel('小时')
        plt.ylabel('电价 (元/kWh)')
        plt.grid(True, axis='y')
        plt.xticks(hours)
    else:
        # 绘制全年电价趋势，按时段标记颜色
        fig, ax = plt.subplots(figsize=(15, 6))
        
        # 获取所有存在的时段类型
        period_types = data['period_type'].unique()
        
        # 为每种时段类型绘制散点
        for period_type in period_types:
            mask = data['period_type'] == period_type
            label = period_labels.get(period_type, f'未知时段{period_type}')
            color = period_colors.get(period_type, 'gray')
            ax.scatter(data.loc[mask, 'datetime'], data.loc[mask, 'price'], 
                      color=color, s=2, label=label)
        
        # 添加图例
        ax.legend()
        
        plt.title('全年电价曲线')
        plt.xlabel('日期')
        plt.ylabel('电价 (元/kWh)')
        plt.grid(True)
    
    plt.show()
def plot_load_and_price(data, date=None, price_config_file=None, auto_scroll=False):
    """
    绘制负载和电价曲线
    data: 包含datetime、load、price和period_type的DataFrame
    date: 指定要绘制哪一天的数据，格式为'YYYY-MM-DD'
    price_config_file: 电价配置文件路径
    auto_scroll: 是否为自动滚动模式
    """
    # 定义时段类型对应的颜色和标签
    period_colors = {
        1: 'orange',    # 尖峰
        2: 'pink',      # 高峰
        3: 'lightblue', # 平段
        4: 'lightgreen',# 低谷
        5: 'blue'       # 深谷
    }
    
    period_labels = {
        1: '尖峰时段',
        2: '高峰时段',
        3: '平段时段',
        4: '低谷时段',
        5: '深谷时段'
    }
    
    def update_figure(fig, ax1, ax2, current_date, data, price_config_file):
        """更新图表数据"""
        specific_date = pd.to_datetime(current_date).date()
        day_data = data[data['datetime'].dt.date == specific_date]
        
        if day_data.empty:
            print(f"日期 {current_date} 没有数据，跳过。")
            return False
        
        # 获取当前日期的月份
        month = pd.to_datetime(current_date).month
        
        # 读取电价配置
        if price_config_file:
            try:
                price_config = pd.read_excel(price_config_file, sheet_name=['时段配置', '电价配置'])
                
                # 获取该月的时段配置
                time_periods = price_config['时段配置']
                price_rules = price_config['电价配置']
                
                # 创建24小时的时段类型和电价数据
                period_types = []
                prices = []
                
                for hour in range(24):
                    # 获取时段类型
                    period_type = time_periods.loc[
                        (time_periods['月份'] == month),
                        f'{hour}-{hour+1}'
                    ].values[0]
                    
                    # 获取电价
                    price = price_rules.loc[
                        price_rules['时段类型'] == period_type,
                        '电价'
                    ].values[0]
                    
                    period_types.append(period_type)
                    prices.append(price)
            except Exception as e:
                print(f"读取电价配置失败，使用数据中的电价信息: {e}")
                prices = day_data['price'].values[:24]
                period_types = day_data['period_type'].values[:24]
        else:
            # 使用数据中的电价和时段类型
            prices = day_data['price'].values[:24]
            period_types = day_data['period_type'].values[:24]
        
        # 创建横坐标的时间区间标签
        hours = range(24)
        hour_labels = [f"{h}" for h in hours]
        
        # 清空当前图表
        ax1.clear()
        ax2.clear()
        
        # 获取负载数据
        loads = day_data['load'].values[:24]
        
        # 更新负载柱状图
        bars1 = ax1.bar(hours, loads, color='lightblue')
        ax1.set_title(f'Hourly Load Profile - {current_date}')
        ax1.set_xlabel('小时')
        ax1.set_ylabel('Load (kW)')
        ax1.set_xticks(hours)
        ax1.set_xticklabels(hour_labels)
        ax1.grid(True)
        
        # 更新电价柱状图
        bars2 = ax2.bar(hours, prices)
        
        # 根据时段类型设置不同的颜色
        for hour, bar, period_type in zip(hours, bars2, period_types):
            bar.set_color(period_colors.get(period_type, 'gray'))
        
        # 创建图例元素
        unique_period_types = set(period_types)
        legend_elements = [plt.Rectangle((0,0),1,1, color=period_colors[pt], label=period_labels[pt]) 
                          for pt in sorted(unique_period_types) if pt in period_colors]
        
        ax2.legend(handles=legend_elements)
        ax2.set_title(f'Price Profile - 月份: {month}')
        ax2.set_xlabel('小时')
        ax2.set_ylabel('电价 (元/kWh)')
        ax2.set_xticks(hours)
        ax2.set_xticklabels(hour_labels)
        ax2.grid(True)
        
        fig.tight_layout()
        fig.canvas.draw_idle()  # 重绘canvas
        
        # 计算并输出信息
        # 1. 总负载
        total_load = sum(loads)
        
        # 2. 各时段电费
        cost_by_period = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}  # 初始化各时段电费
        for hour, load, price, period_type in zip(hours, loads, prices, period_types):
            if period_type in cost_by_period:
                cost_by_period[period_type] += load * price
        
        # 3. 总电费
        total_cost = sum(cost_by_period.values())
        
        # 输出信息
        print(f"Date: {current_date}")
        print(f"Total Load: {total_load:.2f} kWh")
        print("Daily cost: ", end="")
        if 1 in cost_by_period:
            print(f"尖: {cost_by_period[1]:.2f}元", end=", ")
        if 2 in cost_by_period:
            print(f"峰: {cost_by_period[2]:.2f}元", end=", ")
        if 3 in cost_by_period:
            print(f"平: {cost_by_period[3]:.2f}元", end=", ")
        if 4 in cost_by_period:
            print(f"谷: {cost_by_period[4]:.2f}元", end=", ")
        if 5 in cost_by_period:
            print(f"深谷: {cost_by_period[5]:.2f}元", end="")
        print()  # 换行
        print(f"Total cost: {total_cost:.2f}元")
        print("-" * 50)  # 分隔线
        
        return True
    
    # 单日模式
    if date is not None:
        # 创建图形窗口
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
        update_figure(fig, ax1, ax2, date, data, price_config_file)
        plt.show()
        return
    
    # 自动滚动模式
    else:
        # 获取数据中的所有日期
        unique_dates = sorted(data['datetime'].dt.date.unique())
        
        if not unique_dates:
            print("没有找到任何日期数据，无法绘图。")
            return
        
        total_days = len(unique_dates)
        print(f"共找到 {total_days} 天的数据，即将开始自动滚动显示...")
        print("按空格键暂停/继续滚动，按Esc键退出")
        
        # 创建图形和按钮
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
        
        # 添加控制变量
        scroll_paused = [False]  # 使用列表以便在回调函数中修改
        should_exit = [False]
        
        # 键盘事件回调
        def on_key(event):
            if event.key == ' ':  # 空格键
                scroll_paused[0] = not scroll_paused[0]
                status = "暂停" if scroll_paused[0] else "继续"
                print(f"滚动已{status}")
            elif event.key == 'escape':  # Esc键
                should_exit[0] = True
                print("退出滚动")
        
        # 注册键盘事件
        fig.canvas.mpl_connect('key_press_event', on_key)
        
        # 启用交互模式
        plt.ion()
        
        try:
            i = 0
            while i < len(unique_dates) and not should_exit[0]:
                current_date = unique_dates[i]
                date_str = current_date.strftime('%Y-%m-%d')
                print(f"显示第 {i+1}/{total_days} 天: {date_str}")
                
                # 更新图表
                update_success = update_figure(fig, ax1, ax2, date_str, data, price_config_file)
                
                # 暂停查看
                plt.pause(0.01)  # 短暂暂停以刷新界面
                
                # 等待一段时间，同时检查暂停状态
                wait_start = time.time()
                while time.time() - wait_start < 1.5 and not should_exit[0]:  # 等待1.5秒
                    plt.pause(0.1)  # 短暂暂停以处理事件
                    if scroll_paused[0]:
                        plt.pause(0.1)  # 暂停时继续处理事件
                        wait_start = time.time()  # 重置等待时间
                
                # 如果用户按了Esc，则退出循环
                if should_exit[0]:
                    break
                
                # 下一天
                if update_success:
                    i += 1
        
        finally:
            # 禁用交互模式
            plt.ioff()
            
            # 如果用户没有按Esc退出，则保持图形显示
            if not should_exit[0]:
                print("所有日期的数据已显示完毕！")
                plt.show()
def calculate_annual_cost(data):
    """计算年度电费总成本"""
    # 声明全局变量
    global method_basic_capacity_cost_transformer
    
    # 检查数据是否包含NaN值
    if data['load'].isna().any() or data['price'].isna().any():
        print("警告：数据中包含无效值(NaN)，这可能导致计算结果不准确。")
        print("正在尝试处理无效值...")
        
        # 使用0替换NaN值
        data_clean = data.copy()
        data_clean['load'] = data_clean['load'].fillna(0)
        data_clean['price'] = data_clean['price'].fillna(0)
        data = data_clean
    
    # 计算年度总耗电量（全年的load相加）
    annual_total_load = sum(data['load'])
    print(f'年度总耗电量: {annual_total_load:.2f} KWh')
    
    # 计算各时段电费总和
    total_electricity_cost = sum(data['load'] * data['price'])
    
    # 询问用户变压器基本容量费用计算方法
    print("\n变压器基本容量费用计算方式:")
    print("1. 按容量收取 (容量单价×变压器容量×12)")
    print("2. 按需收取 (需量单价×每月最大负荷，12个月累计)")
    print("0. 不计算变压器基本容量费用")
    
    # 获取用户选择
    valid_choice = False
    while not valid_choice:
        try:
            method_choice = int(input("请选择计算方式 (0/1/2): "))
            if method_choice in [0, 1, 2]:
                valid_choice = True
            else:
                print("请输入有效的选项 (0/1/2)")
        except ValueError:
            print("请输入有效的数字")
    
    # 根据用户选择的方法计算变压器基本电费
    if method_choice == 1:
        # 按容量收取
        print('变压器基本电费，按变压器容量收取！')
        
        # 获取容量单价
        capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
        
        # 计算变压器基本电费（容量单价乘以变压器容量）
        transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        total_cost = transformer_basic_cost + total_electricity_cost
        
        print(f'变压器基本电费: {transformer_basic_cost:.2f} 元 (容量单价×变压器容量×12)')
        print(f'电量电费: {total_electricity_cost:.2f} 元')
        print(f'年度总电费（含变压器基本电费和电量电费）: {total_cost:.2f} 元')
        
        # 更新全局变量，用于其他函数的计算
        method_basic_capacity_cost_transformer = 1
        globals()['capacity_price'] = capacity_price
    
    elif method_choice == 2:
        # 按需收取
        print('变压器基本电费，按需收取！')
        
        # 获取需量单价
        demand_price = float(input('请输入需量单价（元/kW·月）: '))
        
        # 计算每个月最大负载
        monthly_max_loads = data.groupby('month')['load'].max()
        
        # 显示每月最大负载
        print("\n每月最大负载:")
        for month, max_load in monthly_max_loads.items():
            print(f"月份 {month}: {max_load:.2f} kW")
        
        # 计算变压器基本电费（需量单价乘以每个月中的最大功率）
        transformer_basic_cost = sum(monthly_max_loads) * demand_price
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        total_cost = transformer_basic_cost + total_electricity_cost
        
        print(f'变压器基本电费: {transformer_basic_cost:.2f} 元 (需量单价×每月最大负荷累计)')
        print(f'电量电费: {total_electricity_cost:.2f} 元')
        print(f'年度总电费（含变压器基本电费和电量电费）: {total_cost:.2f} 元')
        
        # 更新全局变量，用于其他函数的计算
        method_basic_capacity_cost_transformer = 2
        globals()['demand_price'] = demand_price
    
    else:
        # 不计算变压器基本容量费用
        print('不计算变压器基本容量费用')
        total_cost = total_electricity_cost
        print(f'年度电费总成本: {total_electricity_cost:.2f} 元')
        
        # 更新全局变量
        method_basic_capacity_cost_transformer = 0
    
    # 保存总电费用于后续比较
    save_cost = input("\n是否保存此总电费用于与储能系统节省比较？(y/n): ")
    if save_cost.lower() == 'y':
        return total_cost
    
    return None
def simulate_storage_system(data, storage_capacity):
    """模拟储能系统运行"""
    storage_level = initial_storage_capacity
    modified_load = data['load'].copy()
    
    # 确保数据按时间排序
    data_sorted = data.sort_values('datetime')
    
    # 获取数据中所有的日期
    all_dates = data_sorted['datetime'].dt.date.unique()
    
    # 逐日处理数据
    for date in all_dates:
        # 获取当前日期的数据
        day_data = data_sorted[data_sorted['datetime'].dt.date == date]
        
        # 如果当天没有数据，继续下一天
        if day_data.empty:
            continue
            
        # 确保当天24小时都有数据，如果某些小时缺失，补充为0
        day_hours = day_data['datetime'].dt.hour
        missing_hours = set(range(24)) - set(day_hours)
        
        for hour in day_data.index:
            current_load = day_data.loc[hour, 'load']
            current_price = day_data.loc[hour, 'price']
            
            # 简单的充放电策略
            if current_price > np.mean(data['price']):  # 电价高于平均值时放电
                discharge = min(storage_level, max_power_per_system)
                storage_level -= discharge / efficiency_bess
                modified_load[hour] -= discharge
            else:  # 电价低于平均值时充电
                charge = min(max_power_per_system, storage_capacity - storage_level)
                storage_level += charge * efficiency_bess
                modified_load[hour] += charge
            
    return modified_load, storage_level
def plot_storage_system(data, initial_date_index=0, auto_analyze=False):
    """绘制增加储能系统后的曲线（两充两放策略）并分析套利模式"""
    
    # 获取所有可用日期
    unique_dates = sorted(data['datetime'].dt.date.unique())
    
    if len(unique_dates) == 0:
        print("数据中没有有效日期，无法执行分析和绘图。")
        return
    
    if initial_date_index >= len(unique_dates):
        initial_date_index = 0
    
    current_date_idx = initial_date_index
    current_date = unique_dates[current_date_idx]
    
    # 用于存储每个月电费的字典
    monthly_costs = {}
    
    # 分析并绘制当前日期的数据
    def analyze_and_plot_date(date, fig=None, axs=None, store_results=False):
        # 获取当天数据
        day_data = data[data['datetime'].dt.date == date].copy()
        
        # 检查一天数据是否完整
        if len(day_data) == 0:
            print(f"日期 {date} 没有任何负载数据，将视为全天负载为0。")
            # 创建一个全为0的24小时数据
            hours = range(24)
            date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
            
            # 获取这一天应有的时间索引
            date_indices = [pd.Timestamp(date_str) + pd.Timedelta(hours=h) for h in hours]
            
            # 获取该日期月份
            month = pd.Timestamp(date).month
            
            # 创建默认的时段和电价信息（根据月份）
            default_periods = []
            default_prices = []
            
            # 从数据中获取该月份的时段和电价信息
            month_data = data[data['datetime'].dt.month == month]
            if not month_data.empty:
                for hour in hours:
                    hour_data = month_data[month_data['datetime'].dt.hour == hour]
                    if not hour_data.empty:
                        default_periods.append(hour_data['period_type'].iloc[0])
                        default_prices.append(hour_data['price'].iloc[0])
                    else:
                        # 如果没有该小时的数据，使用默认值
                        default_periods.append(FLAT)  # 默认为平段
                        default_prices.append(0.5)    # 默认电价
            else:
                # 如果没有该月的数据，使用默认值
                default_periods = [FLAT] * 24  # 所有时段都默认为平段
                default_prices = [0.5] * 24    # 所有电价都默认为0.5
            
            # 创建临时DataFrame
            temp_data = {
                'datetime': date_indices,
                'load': [0] * 24,  # 负载全为0
                'price': default_prices,
                'period_type': default_periods,
                'month': [month] * 24,
                'hour': hours
            }
            day_data = pd.DataFrame(temp_data)
        elif len(day_data) < 24:
            print(f"日期 {date} 数据不足24小时，将补充缺失小时的数据（负载为0）。")
            
            # 获取当前存在的小时
            existing_hours = day_data['hour'].unique()
            
            # 找出缺失的小时
            missing_hours = set(range(24)) - set(existing_hours)
            
            # 如果有缺失的小时，补充数据
            if missing_hours:
                date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
                month = pd.Timestamp(date).month
                
                # 创建缺失小时的数据
                missing_data = []
                
                for hour in missing_hours:
                    # 根据月份和小时查找相应的时段和电价
                    hour_data = data[(data['datetime'].dt.month == month) & (data['datetime'].dt.hour == hour)]
                    
                    if not hour_data.empty:
                        period_type = hour_data['period_type'].iloc[0]
                        price = hour_data['price'].iloc[0]
                    else:
                        # 如果没有该小时的参考数据，使用默认值
                        period_type = FLAT  # 默认为平段
                        price = 0.5        # 默认电价
                    
                    # 创建该小时的数据行
                    missing_data.append({
                        'datetime': pd.Timestamp(date_str) + pd.Timedelta(hours=hour),
                        'load': 0,  # 缺失小时的负载设为0
                        'price': price,
                        'period_type': period_type,
                        'month': month,
                        'hour': hour
                    })
                
                # 添加缺失小时的数据
                missing_df = pd.DataFrame(missing_data)
                day_data = pd.concat([day_data, missing_df])
                
                # 确保按小时排序
                day_data = day_data.sort_values('hour')
            
        # 重置索引，保证能按位置访问
        day_data = day_data.reset_index(drop=True)
            
        day_str = date.strftime('%Y-%m-%d')
        month = date.month
        periods = day_data['period_type'].values
        
        # 检查periods是否有缺失或者无效值
        if np.isnan(periods).any():
            print(f"警告：{day_str} 的时段数据存在缺失，将使用默认值替换。")
            # 将NaN替换为默认时段（平段）
            periods = np.where(np.isnan(periods), FLAT, periods)
        
        arbitrage_results = []
        charge1_start, charge1_duration, charge1_total_len = None, None, None
        discharge1_start, discharge1_duration, discharge1_type_code = None, None, None
        charge2_start, charge2_duration, charge2_total_len = None, None, None
        discharge2_start, discharge2_duration, discharge2_type_code = None, None, None
        
        # --- 第一次套利识别 ---
        # 1. 查找充电窗口 (优先谷，后平)
        charge1_start_valley, _, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
        charge1_start_flat, _, _ = find_continuous_window(periods, [FLAT], 2, 0)
        charge1_type = None
        if charge1_start_valley is not None:
            # 如果平段窗口在谷段窗口之前，但谷段存在，则优先谷段
            if charge1_start_flat is not None and charge1_start_flat < charge1_start_valley:
                 print(f"[{day_str}] 注意：发现平段充电窗口早于谷段，但优先选择谷段充电。")
            charge1_start = charge1_start_valley
            charge1_duration = 2
            charge1_type = "谷"
        elif charge1_start_flat is not None:
            charge1_start = charge1_start_flat
            charge1_duration = 2
            charge1_type = "平"
            print(f"[{day_str}] 未找到连续2小时谷段，使用平段充电。")
        
        # 2. 如果找到充电窗口，查找放电窗口
        if charge1_start is not None:
            search_start_discharge1 = charge1_start + charge1_duration
            
            # 查找尖峰和高峰窗口（至少2小时）
            sharp1_start, _, sharp1_total_len = find_continuous_window(periods, [SHARP], 2, search_start_discharge1)
            peak1_start, _, peak1_total_len = find_continuous_window(periods, [PEAK], 2, search_start_discharge1)
            # 决定放电类型 (比较总时长)
            discharge1_type_name = None
            if sharp1_start is not None and peak1_start is not None:
                if sharp1_total_len >= peak1_total_len: # 尖峰优先
                    discharge1_start = sharp1_start
                    discharge1_duration = 2
                    discharge1_type_code = SHARP
                    discharge1_type_name = "尖"
                else: # 高峰优先
                    discharge1_start = peak1_start
                    discharge1_duration = 2
                    discharge1_type_code = PEAK
                    discharge1_type_name = "峰"
            elif sharp1_start is not None:
                discharge1_start = sharp1_start
                discharge1_duration = 2
                discharge1_type_code = SHARP
                discharge1_type_name = "尖"
            elif peak1_start is not None:
                discharge1_start = peak1_start
                discharge1_duration = 2
                discharge1_type_code = PEAK
                discharge1_type_name = "峰"
            else:
                 # 检查特殊平谷套利: 是否只有平/谷，无尖/峰
                 has_sharp_peak = any(p in [SHARP, PEAK] for p in periods)
                 if not has_sharp_peak and charge1_type == "谷":
                     # 找充电后的2小时平段
                     flat_discharge_start, _, _ = find_continuous_window(periods, [FLAT], 2, search_start_discharge1)
                     if flat_discharge_start is not None:
                         discharge1_start = flat_discharge_start
                         discharge1_duration = 2
                         discharge1_type_code = FLAT # 特殊情况
                         discharge1_type_name = "平"
                         arbitrage_results.append(f"第一次套利：平谷套利 (充电时段: {charge1_start}-{charge1_start+charge1_duration-1}, 放电时段: {discharge1_start}-{discharge1_start+discharge1_duration-1})")
            if discharge1_type_name and discharge1_type_name != "平": # 非特殊平谷
                arbitrage_results.append(f"第一次套利：{discharge1_type_name}{charge1_type}套利 (充电时段: {charge1_start}-{charge1_start+charge1_duration-1}, 放电时段: {discharge1_start}-{discharge1_start+discharge1_duration-1})")
        # --- 第二次套利识别 ---
        if discharge1_start is not None: # 必须在第一次套利完成后
            search_start_charge2 = discharge1_start + discharge1_duration
            
            # 1. 查找充电窗口 (优先谷，后平, >= 2小时)
            charge2_start_valley, _, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, search_start_charge2)
            charge2_start_flat, _, _ = find_continuous_window(periods, [FLAT], 2, search_start_charge2)
            charge2_type = None
            if charge2_start_valley is not None:
                if charge2_start_flat is not None and charge2_start_flat < charge2_start_valley:
                     print(f"[{day_str}] 注意：第二次套利发现平段充电窗口早于谷段，但优先选择谷段充电。")
                charge2_start = charge2_start_valley
                charge2_duration = 2
                charge2_type = "谷"
            elif charge2_start_flat is not None:
                charge2_start = charge2_start_flat
                charge2_duration = 2
                charge2_type = "平"
                print(f"[{day_str}] 第二次套利未找到连续2小时谷段，使用平段充电。")
            # 2. 如果找到充电窗口，查找放电窗口 (>= 1小时)
            if charge2_start is not None:
                search_start_discharge2 = charge2_start + charge2_duration
                
                # 修改：同样查找连续2小时的放电窗口，而不是1小时
                sharp2_start, _, sharp2_total_len = find_continuous_window(periods, [SHARP], 2, search_start_discharge2)
                peak2_start, _, peak2_total_len = find_continuous_window(periods, [PEAK], 2, search_start_discharge2)
                discharge2_type_name = None
                if sharp2_start is not None and peak2_start is not None:
                    if sharp2_total_len >= peak2_total_len:
                        discharge2_start = sharp2_start
                        discharge2_duration = 2  # 修改：第二次放电改为2小时
                        discharge2_type_code = SHARP
                        discharge2_type_name = "尖"
                    else:
                        discharge2_start = peak2_start
                        discharge2_duration = 2  # 修改：第二次放电改为2小时
                        discharge2_type_code = PEAK
                        discharge2_type_name = "峰"
                elif sharp2_start is not None:
                    discharge2_start = sharp2_start
                    discharge2_duration = 2  # 修改：第二次放电改为2小时
                    discharge2_type_code = SHARP
                    discharge2_type_name = "尖"
                elif peak2_start is not None:
                    discharge2_start = peak2_start
                    discharge2_duration = 2  # 修改：第二次放电改为2小时
                    discharge2_type_code = PEAK
                    discharge2_type_name = "峰"
                # 不再检查第二次的平谷套利
                # 如果找不到连续2小时，则尝试1小时（兼容原有逻辑）
                if discharge2_type_name is None:
                    sharp2_start, _, sharp2_total_len = find_continuous_window(periods, [SHARP], 1, search_start_discharge2)
                    peak2_start, _, peak2_total_len = find_continuous_window(periods, [PEAK], 1, search_start_discharge2)
                    
                    if sharp2_start is not None and peak2_start is not None:
                        if sharp2_total_len >= peak2_total_len:
                            discharge2_start = sharp2_start
                            discharge2_duration = 1  # 只有1小时可用
                            discharge2_type_code = SHARP
                            discharge2_type_name = "尖"
                        else:
                            discharge2_start = peak2_start
                            discharge2_duration = 1  # 只有1小时可用
                            discharge2_type_code = PEAK
                            discharge2_type_name = "峰"
                    elif sharp2_start is not None:
                        discharge2_start = sharp2_start
                        discharge2_duration = 1  # 只有1小时可用
                        discharge2_type_code = SHARP
                        discharge2_type_name = "尖"
                    elif peak2_start is not None:
                        discharge2_start = peak2_start
                        discharge2_duration = 1  # 只有1小时可用
                        discharge2_type_code = PEAK
                        discharge2_type_name = "峰"
                if discharge2_type_name:
                     arbitrage_results.append(f"第二次套利：{discharge2_type_name}{charge2_type}套利 (充电时段: {charge2_start}-{charge2_start+charge2_duration-1}, 放电时段: {discharge2_start}-{discharge2_start+discharge2_duration-1})")
        # --- 打印套利结果 ---
        if not auto_analyze or not store_results:
            print(f"\n--- {day_str} 套利分析结果 ---")
            if not arbitrage_results:
                print("当天未识别到满足条件的套利模式。")
            else:
                for result in arbitrage_results:
                    print(result)
            print("--------------------------")
        # --- 执行储能模拟 ---
        storage_level = initial_storage_capacity
        # 使用 .to_numpy() 提高性能
        original_load = day_data['load'].to_numpy()
        modified_load = original_load.copy()
        storage_power = np.zeros(24) # 初始化为0
        
        # 确保数组长度为24
        if len(original_load) < 24:
            print(f"警告：{day_str} 负载数据长度不足24小时，已自动补充为0")
            # 扩展数组至24小时
            pad_length = 24 - len(original_load)
            original_load = np.pad(original_load, (0, pad_length), 'constant')
            modified_load = original_load.copy()
        elif len(original_load) > 24:
            print(f"警告：{day_str} 负载数据超过24小时，将截断多余数据")
            original_load = original_load[:24]
            modified_load = original_load.copy()
        
        # 标记第一次和第二次套利的数组
        first_arbitrage_mask = np.zeros(24, dtype=bool)
        second_arbitrage_mask = np.zeros(24, dtype=bool)
        
        # 检查是否所有负载都为0，如果是，则跳过套利模拟
        if np.all(original_load == 0):
            print(f"[{day_str}] 当天所有负载为0，跳过套利模拟")
        else:
            # 第一次套利模拟
            if charge1_start is not None and discharge1_start is not None:
                # 检查是否越界
                if charge1_start >= 24 or discharge1_start >= 24:
                    print(f"[{day_str}] 套利时段超出范围，跳过第一次套利")
                else:
                    # 确保不会越界
                    charge1_end = min(charge1_start + charge1_duration, 24)
                    discharge1_end = min(discharge1_start + discharge1_duration, 24)
                    
                    # 充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge1 = min(max_power_per_system, charge_needed / ((charge1_end - charge1_start) * efficiency_bess)) # 考虑效率
                    for h in range(charge1_start, charge1_end):
                        actual_charge = min(power_charge1, max_power_per_system) # 再次确认不超过最大功率
                        # 确保不超过电池容量
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess) 
                        if actual_charge <= 0: break # 电池已满或计算错误
                        storage_power[h] = actual_charge
                        modified_load[h] += actual_charge # 充电增加电网负荷
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system) # 确保不超过上限
                        first_arbitrage_mask[h] = True # 标记为第一次套利
                    
                    # 放电
                    discharge_available = storage_level
                    power_discharge1 = min(max_power_per_system, discharge_available * efficiency_bess / (discharge1_end - discharge1_start)) # 考虑效率
                    for h in range(discharge1_start, discharge1_end):
                        actual_discharge = min(power_discharge1, max_power_per_system)
                        # 确保不超过可用电量
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess) 
                        # 确保不超过企业用电负荷（避免向电网反向输电）
                        actual_discharge = min(actual_discharge, original_load[h])
                        if actual_discharge <= 0: break # 电池已空或计算错误
                        storage_power[h] = -actual_discharge
                        modified_load[h] -= actual_discharge # 放电减少电网负荷
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0) # 确保不低于0
                        first_arbitrage_mask[h] = True # 标记为第一次套利
            
            # 第二次套利模拟
            if charge2_start is not None and discharge2_start is not None:
                # 检查是否越界
                if charge2_start >= 24 or discharge2_start >= 24:
                    print(f"[{day_str}] 套利时段超出范围，跳过第二次套利")
                else:
                    # 确保不会越界
                    charge2_end = min(charge2_start + charge2_duration, 24)
                    discharge2_end = min(discharge2_start + discharge2_duration, 24)
                    
                    # 充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge2 = min(max_power_per_system, charge_needed / ((charge2_end - charge2_start) * efficiency_bess))
                    for h in range(charge2_start, charge2_end):
                        actual_charge = min(power_charge2, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        storage_power[h] = actual_charge
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                        second_arbitrage_mask[h] = True # 标记为第二次套利
                    
                    # 放电 (支持1-2小时)
                    discharge_available = storage_level
                    power_discharge2 = min(max_power_per_system, discharge_available * efficiency_bess / (discharge2_end - discharge2_start))
                    for h in range(discharge2_start, discharge2_end):
                        actual_discharge = min(power_discharge2, max_power_per_system)
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        # 确保不超过企业用电负荷（避免向电网反向输电）
                        actual_discharge = min(actual_discharge, original_load[h])
                        if actual_discharge <= 0: break  # 电池已空或计算错误
                        storage_power[h] = -actual_discharge
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
                        second_arbitrage_mask[h] = True  # 标记为第二次套利
        
        final_storage_level = storage_level
        # 计算各时段电费（按最终负载曲线）
        costs_by_period = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
        period_names = {SHARP: "尖", PEAK: "峰", FLAT: "平", VALLEY: "谷", DEEP_VALLEY: "深谷"}
        
        # 确保日数据长度与负载数组长度一致
        price_array = day_data['price'].values
        period_array = day_data['period_type'].values
        
        # 如果数据不足24小时，进行填充
        if len(price_array) < 24:
            price_array = np.pad(price_array, (0, 24 - len(price_array)), 'constant')
        if len(period_array) < 24:
            period_array = np.pad(period_array, (0, 24 - len(period_array)), 'constant', constant_values=FLAT)
            
        # 确保只使用前24小时
        price_array = price_array[:24]
        period_array = period_array[:24]
        
        for h in range(24):
            period = period_array[h]
            price = price_array[h]
            load = modified_load[h]
            costs_by_period[period] += load * price
        
        # 如果是自动分析模式，将每日电费统计添加到月度统计中
        if store_results:
            # 初始化当月统计
            if month not in monthly_costs:
                monthly_costs[month] = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
            
            # 累加每日电费到月度统计
            for period, cost in costs_by_period.items():
                monthly_costs[month][period] += cost
        # --- 绘图 ---
        if fig is None or axs is None:
            return True
            
        for ax in axs:
            ax.clear()
                
        hours = range(24)
        # 创建统一的横坐标标签格式（与第四个图表一致）
        hour_labels = [f"{h}" for h in hours]
        
        # 确保所有数据都是24小时长度
        if len(original_load) < 24:
            original_load = np.pad(original_load, (0, 24 - len(original_load)), 'constant')
        elif len(original_load) > 24:
            original_load = original_load[:24]
            
        if len(modified_load) < 24:
            modified_load = np.pad(modified_load, (0, 24 - len(modified_load)), 'constant')
        elif len(modified_load) > 24:
            modified_load = modified_load[:24]
        
        # 1. 原始负载
        axs[0].bar(hours, original_load, color='lightblue')
        axs[0].set_title(f'Hourly Load Profile - {day_str}')
        axs[0].set_ylabel('Load (kW)')
        axs[0].grid(True)
        axs[0].set_xticks(hours)
        axs[0].set_xticklabels(hour_labels)
        
        # 2. 电价
        period_colors = { SHARP: 'orange', PEAK: 'pink', FLAT: 'lightblue', VALLEY: 'lightgreen', DEEP_VALLEY: 'blue' }
        period_labels = { SHARP: '尖峰', PEAK: '高峰', FLAT: '平段', VALLEY: '低谷', DEEP_VALLEY: '深谷' }
        
        # 确保电价和时段数据是24小时长度
        price_array = day_data['price'].values
        if len(price_array) < 24:
            price_array = np.pad(price_array, (0, 24 - len(price_array)), 'constant')
        elif len(price_array) > 24:
            price_array = price_array[:24]
            
        period_types_day = day_data['period_type'].values
        if len(period_types_day) < 24:
            period_types_day = np.pad(period_types_day, (0, 24 - len(period_types_day)), 'constant', constant_values=FLAT)
        elif len(period_types_day) > 24:
            period_types_day = period_types_day[:24]
            
        bars_price = axs[1].bar(hours, price_array)
        
        unique_period_types = np.unique(period_types_day)
        unique_period_types = unique_period_types[~np.isnan(unique_period_types)].astype(int)  # 移除NaN并转为整数
        
        legend_elements = [plt.Rectangle((0,0),1,1, color=period_colors.get(pt, 'gray'), label=period_labels.get(pt, f'未知{pt}'))
                          for pt in unique_period_types if pt in period_colors]
                          
        for hour, bar, period_type in zip(hours, bars_price, period_types_day):
            if np.isnan(period_type):
                # 如果时段类型是NaN，使用默认颜色
                bar.set_color('gray')
            else:
                bar.set_color(period_colors.get(int(period_type), 'gray'))
                
        axs[1].legend(handles=legend_elements, loc='upper right')
        axs[1].set_title('Price Profile')
        axs[1].set_ylabel('电价 (元/kWh)')
        axs[1].grid(True)
        axs[1].set_xticks(hours)
        axs[1].set_xticklabels(hour_labels)
        
        # 3. 储能充放电 (使用不同颜色区分第一次和第二次套利)
        # 第一次套利用绿色
        first_arb_power = np.where(first_arbitrage_mask, storage_power, 0)
        # 第二次套利用紫色
        second_arb_power = np.where(second_arbitrage_mask, storage_power, 0)
        # 无套利用灰色
        no_arb_power = np.where(~first_arbitrage_mask & ~second_arbitrage_mask, storage_power, 0)
        
        # 先绘制静止状态（灰色）
        if np.any(no_arb_power != 0):
            axs[2].bar(hours, no_arb_power, color='gray', label='静止')
        
        # 绘制第一次套利（绿色）
        if np.any(first_arb_power != 0):
            axs[2].bar(hours, first_arb_power, color='green', label='第一次套利')
            
        # 绘制第二次套利（紫色）
        if np.any(second_arb_power != 0):
            axs[2].bar(hours, second_arb_power, color='purple', label='第二次套利')
            
        axs[2].set_title(f'Storage Charging (Remain capacity: {final_storage_level:.2f} kWh)')
        axs[2].set_ylabel('Power (kW)')
        axs[2].axhline(0, color='black', linewidth=0.5)
        axs[2].grid(True)
        axs[2].legend()
        axs[2].set_xticks(hours)
        axs[2].set_xticklabels(hour_labels)
        
        # 4. 最终负载
        axs[3].bar(hours, modified_load, color='blue')
        axs[3].set_title('Total Power Consumption After Storage')
        axs[3].set_ylabel('Load (kW)')
        axs[3].grid(True)
        
        # 设置X轴标签
        axs[3].set_xlabel('小时')
        axs[3].set_xticks(hours)
        axs[3].set_xticklabels(hour_labels)
        
        # 打印当天各时段电费
        if not auto_analyze:
            daily_total = sum(costs_by_period.values())
            print(f"\n{day_str} 加装储能系统后电费统计:")
            for period, cost in costs_by_period.items():
                if cost > 0:
                    print(f"{period_names.get(period, f'未知{period}')}：{cost:.2f}元")
            print(f"总计：{daily_total:.2f}元")
            print("--------------------------")
        
        return True
    # 自动分析所有日期
    def analyze_all_dates():
        nonlocal monthly_costs
        monthly_costs = {}  # 重置月度统计
        
        print("\n开始分析所有日期的套利情况...")
        for date in unique_dates:
            analyze_and_plot_date(date, store_results=True)
        
        # 打印每月电费统计
        print("\n\n=== 每月加装储能系统后电费统计 ===")
        for month in sorted(monthly_costs.keys()):
            month_data = monthly_costs[month]
            month_total = sum(month_data.values())
            
            print(f"\n{month}月电费统计:")
            for period, cost in month_data.items():
                if cost > 0:
                    period_name = {SHARP: "尖", PEAK: "峰", FLAT: "平", VALLEY: "谷", DEEP_VALLEY: "深谷"}.get(period, f"未知{period}")
                    print(f"{period_name}：{cost:.2f}元")
            print(f"月总计：{month_total:.2f}元")
        
        # 计算年度总电费
        annual_total = sum(sum(month_data.values()) for month_data in monthly_costs.values())
        print(f"\n年度总电费：{annual_total:.2f}元")
        print("=" * 40)
    # 如果是自动分析模式，直接分析所有日期并返回
    if auto_analyze:
        analyze_all_dates()
        return
    # 创建主图形
    fig = plt.figure(figsize=(12, 16))
    
    # 手动设置子图位置，留出底部空间给按钮
    ax1 = fig.add_axes([0.125, 0.77, 0.775, 0.18])  # 上方20%
    ax2 = fig.add_axes([0.125, 0.54, 0.775, 0.18])  # 中上20%
    ax3 = fig.add_axes([0.125, 0.31, 0.775, 0.18])  # 中下20% 
    ax4 = fig.add_axes([0.125, 0.08, 0.775, 0.18])  # 下方20%，留出底部空间
    axs = [ax1, ax2, ax3, ax4]
    
    # 添加导航按钮的回调函数
    def on_prev(event):
        nonlocal current_date_idx
        current_date_idx = (current_date_idx - 1) % len(unique_dates)
        current_date = unique_dates[current_date_idx]
        analyze_and_plot_date(current_date, fig, axs)
        fig.canvas.draw_idle()
        
    def on_next(event):
        nonlocal current_date_idx
        current_date_idx = (current_date_idx + 1) % len(unique_dates)
        current_date = unique_dates[current_date_idx]
        analyze_and_plot_date(current_date, fig, axs)
        fig.canvas.draw_idle()
    
    def on_stop(event):
        plt.close(fig)
        print("\n图表已关闭。")
    
    def on_analyze_all(event):
        plt.close(fig)
        analyze_all_dates()
    # 添加导航按钮
    button_ax_prev = fig.add_axes([0.3, 0.01, 0.15, 0.04])
    button_ax_next = fig.add_axes([0.5, 0.01, 0.15, 0.04])
    button_ax_stop = fig.add_axes([0.7, 0.01, 0.15, 0.04])
    button_ax_analyze_all = fig.add_axes([0.1, 0.01, 0.15, 0.04])
    
    btn_prev = Button(button_ax_prev, '上一天')
    btn_next = Button(button_ax_next, '下一天')
    btn_stop = Button(button_ax_stop, '停止')
    btn_analyze_all = Button(button_ax_analyze_all, '分析全部')
    
    btn_prev.on_clicked(on_prev)
    btn_next.on_clicked(on_next)
    btn_stop.on_clicked(on_stop)
    btn_analyze_all.on_clicked(on_analyze_all)
    
    # 显示初始日期数据
    success = analyze_and_plot_date(unique_dates[current_date_idx], fig, axs)
    if not success:
        # 如果初始日期无法显示，尝试下一个日期
        for i in range(1, len(unique_dates)):
            idx = (current_date_idx + i) % len(unique_dates)
            if analyze_and_plot_date(unique_dates[idx], fig, axs):
                current_date_idx = idx
                break
    
    plt.show()
def find_optimal_storage_capacity(data):
    """寻找最佳储能系统容量"""
    # 声明全局变量，必须放在函数开头
    global storage_capacity_per_system, max_power_per_system
    
    # 测试范围：从100kW到10000kW的功率
    power_range = list(range(100, 10001, 100))  # 以100kW为步长
    costs = []
    cost_breakdown = {}
    
    # 储能系统容量与功率的比例
    capacity_power_ratio = storage_capacity_per_system / max_power_per_system  # 约等于261/125
    
    # 获取储能系统单位造价和使用寿命
    system_unit_cost = float(input('请输入储能系统单位造价（元/Wh）: '))
    system_lifetime = float(input('请输入储能系统使用寿命（年）: '))
    
    print("正在分析不同储能容量下的总成本...\n")
    print("功率(kW)\t容量(kWh)\t年度总电费(元)\t储能系统造价(元)\t年度摊销成本(元)\t年度总成本(元)")
    print("-" * 120)
    
    # 保存原始参数值
    original_capacity = storage_capacity_per_system
    original_power = max_power_per_system
    
    # 保存用户输入的价格，避免重复输入
    capacity_price = 0
    demand_price = 0
    if method_basic_capacity_cost_transformer == 1:  # 按容量收取
        capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
    elif method_basic_capacity_cost_transformer == 2:  # 按需收取
        demand_price = float(input('请输入需量单价（元/kW·月）: '))
    
    # 遍历所有容量进行分析
    for power in power_range:
        # 计算对应的储能容量
        capacity = power * capacity_power_ratio
        
        # 设置新的容量和功率
        storage_capacity_per_system = capacity
        max_power_per_system = power
        
        # 进行所有日期的模拟计算
        monthly_costs = {}
        monthly_max_loads = {}  # 存储每月最大负载，用于计算按需收费
        
        # 确保数据按时间排序
        data_sorted = data.sort_values('datetime')
        
        # 按日期进行分析
        for date in sorted(data_sorted['datetime'].dt.date.unique()):
            # 获取当天数据
            day_data = data_sorted[data_sorted['datetime'].dt.date == date].copy()
            
            # 如果当天没有数据，创建一个全零的数据集
            if day_data.empty:
                # 创建当天24小时的全零数据
                month = pd.Timestamp(date).month
                date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
                hours = range(24)
                
                # 获取这一天应有的时间索引
                date_indices = [pd.Timestamp(date_str) + pd.Timedelta(hours=h) for h in hours]
                
                # 从数据中获取当月电价和时段信息
                default_periods = []
                default_prices = []
                for hour in hours:
                    month_hour_data = data_sorted[(data_sorted['datetime'].dt.month == month) & 
                                             (data_sorted['datetime'].dt.hour == hour)]
                    if not month_hour_data.empty:
                        default_periods.append(month_hour_data['period_type'].iloc[0])
                        default_prices.append(month_hour_data['price'].iloc[0])
                    else:
                        default_periods.append(FLAT)
                        default_prices.append(0.5)
                
                # 创建临时DataFrame
                temp_data = {
                    'datetime': date_indices,
                    'load': [0] * 24,
                    'price': default_prices,
                    'period_type': default_periods,
                    'month': [month] * 24,
                    'hour': hours
                }
                day_data = pd.DataFrame(temp_data)
            
            # 如果数据不足24小时，补充缺失的小时
            elif len(day_data) < 24:
                # 找出当前有哪些小时
                existing_hours = day_data['hour'].unique()
                missing_hours = set(range(24)) - set(existing_hours)
                
                if missing_hours:
                    month = day_data['month'].iloc[0]
                    date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
                    
                    # 为缺失的小时创建数据
                    missing_data = []
                    for hour in missing_hours:
                        # 查找同月同小时的数据作为参考
                        ref_data = data_sorted[(data_sorted['datetime'].dt.month == month) & 
                                          (data_sorted['datetime'].dt.hour == hour)]
                        
                        if not ref_data.empty:
                            period_type = ref_data['period_type'].iloc[0]
                            price = ref_data['price'].iloc[0]
                        else:
                            period_type = FLAT
                            price = 0.5
                        
                        missing_data.append({
                            'datetime': pd.Timestamp(date_str) + pd.Timedelta(hours=hour),
                            'load': 0,  # 缺失小时负载为0
                            'price': price,
                            'period_type': period_type,
                            'month': month,
                            'hour': hour
                        })
                    
                    # 添加缺失小时数据
                    missing_df = pd.DataFrame(missing_data)
                    day_data = pd.concat([day_data, missing_df])
                    
                # 确保按小时排序
                day_data = day_data.sort_values('hour').reset_index(drop=True)
            
            month = day_data['month'].iloc[0]
            periods = day_data['period_type'].values
            
            # 处理可能的NaN值
            if np.isnan(periods).any():
                periods = np.where(np.isnan(periods), FLAT, periods)
                
            # 执行套利分析和模拟
            # 第一次套利（简化版）
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None:
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
                
                if discharge1_start is not None:
                    # 确保索引不会超出范围
                    charge1_end = min(charge1_start + charge1_duration, 24)
                    discharge1_end = min(discharge1_start + discharge1_duration, 24)
                    
                    # 模拟充电
                    storage_level = initial_storage_capacity
                    original_load = day_data['load'].values
                    modified_load = original_load.copy()
                    
                    # 确保负载数组长度为24
                    if len(original_load) < 24:
                        original_load = np.pad(original_load, (0, 24 - len(original_load)), 'constant')
                        modified_load = original_load.copy()
                    elif len(original_load) > 24:
                        original_load = original_load[:24]
                        modified_load = original_load.copy()
                    
                    # 充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge = min(max_power_per_system, charge_needed / ((charge1_end - charge1_start) * efficiency_bess))
                    for h in range(charge1_start, charge1_end):
                        actual_charge = min(power_charge, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                    
                    # 放电
                    discharge_available = storage_level
                    power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / (discharge1_end - discharge1_start))
                    for h in range(discharge1_start, discharge1_end):
                        # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                        actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        if actual_discharge <= 0: break
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
            
            # 第二次套利（简化版）
            if discharge1_start is not None:
                charge2_start, charge2_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, discharge1_start + discharge1_duration)
                if charge2_start is None:
                    charge2_start, charge2_duration, _ = find_continuous_window(periods, [FLAT], 2, discharge1_start + discharge1_duration)
                
                if charge2_start is not None:
                    discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:  # 尝试找1小时窗口
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 1, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 1, charge2_start + charge2_duration)
                    
                    if discharge2_start is not None:
                        # 确保索引不会超出范围
                        charge2_end = min(charge2_start + charge2_duration, 24)
                        discharge2_end = min(discharge2_start + discharge2_duration, 24)
                        
                        # 充电
                        charge_needed = storage_capacity_per_system - storage_level
                        power_charge = min(max_power_per_system, charge_needed / ((charge2_end - charge2_start) * efficiency_bess))
                        for h in range(charge2_start, charge2_end):
                            actual_charge = min(power_charge, max_power_per_system)
                            actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                            if actual_charge <= 0: break
                            modified_load[h] += actual_charge
                            storage_level += actual_charge * efficiency_bess
                            storage_level = min(storage_level, storage_capacity_per_system)
                        
                        # 放电
                        discharge_available = storage_level
                        power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / (discharge2_end - discharge2_start))
                        for h in range(discharge2_start, discharge2_end):
                            # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                            actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                            actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                            if actual_discharge <= 0: break
                            modified_load[h] -= actual_discharge
                            storage_level -= actual_discharge / efficiency_bess
                            storage_level = max(storage_level, 0)
            
            # 计算当天各时段电费
            day_costs = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
            
            # 确保价格和时段数组长度为24
            price_array = day_data['price'].values
            period_array = day_data['period_type'].values
            
            if len(price_array) < 24:
                price_array = np.pad(price_array, (0, 24 - len(price_array)), 'constant')
            elif len(price_array) > 24:
                price_array = price_array[:24]
                
            if len(period_array) < 24:
                period_array = np.pad(period_array, (0, 24 - len(period_array)), 'constant', constant_values=FLAT)
            elif len(period_array) > 24:
                period_array = period_array[:24]
            
            for h in range(24):
                period = period_array[h]
                if np.isnan(period):
                    period = FLAT  # 处理可能的NaN值
                else:
                    period = int(period)  # 确保是整数
                    
                price = price_array[h]
                load = modified_load[h]
                day_costs[period] += load * price
            
            # 记录该天的最大负载，用于按需收费计算
            max_load_of_day = np.max(modified_load)
            
            # 累加到月度统计
            if month not in monthly_costs:
                monthly_costs[month] = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
                monthly_max_loads[month] = 0
            
            for period, cost in day_costs.items():
                monthly_costs[month][period] += cost
            
            # 更新月最大负载
            monthly_max_loads[month] = max(monthly_max_loads[month], max_load_of_day)
        
        # 计算电量电费总额
        electricity_cost = sum(sum(month_data.values()) for month_data in monthly_costs.values())
        
        # 计算变压器基本电费
        transformer_basic_cost = 0
        if method_basic_capacity_cost_transformer == 1:  # 按容量收取
            transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
        elif method_basic_capacity_cost_transformer == 2:  # 按需收取
            transformer_basic_cost = sum(monthly_max_loads.values()) * demand_price
        
        # 计算年度总电费
        annual_electricity_cost = electricity_cost + transformer_basic_cost
        
        # 计算储能系统总造价
        system_total_cost = system_unit_cost * capacity * 1000  # 转换为Wh
        
        # 计算储能系统年度摊销成本
        annual_system_cost = system_total_cost / system_lifetime
        
        # 计算年度总成本（电费+系统摊销）
        annual_total_cost = annual_electricity_cost + annual_system_cost
        
        costs.append(annual_total_cost)
        
        # 存储详细电费数据
        cost_breakdown[power] = {
            'monthly': monthly_costs.copy(),
            'electricity_cost': electricity_cost,
            'transformer_basic_cost': transformer_basic_cost,
            'annual_electricity_cost': annual_electricity_cost,
            'system_total_cost': system_total_cost,
            'annual_system_cost': annual_system_cost,
            'annual_total_cost': annual_total_cost
        }
        
        # 打印当前容量的分析结果
        print(f"{power}\t{capacity:.2f}\t{annual_electricity_cost:.2f}\t{system_total_cost:.2f}\t{annual_system_cost:.2f}\t{annual_total_cost:.2f}")
    
    # 还原原始参数
    storage_capacity_per_system = original_capacity
    max_power_per_system = original_power
    
    # 找出最优容量
    optimal_power_idx = np.argmin(costs)
    optimal_power = power_range[optimal_power_idx]
    optimal_capacity = optimal_power * capacity_power_ratio
    optimal_cost = costs[optimal_power_idx]
    
    print("\n" + "=" * 100)
    print(f'最佳储能系统功率: {optimal_power} kW')
    print(f'最佳储能系统容量: {optimal_capacity:.2f} kWh')
    print(f'年度电费(含变压器基本电费): {cost_breakdown[optimal_power]["annual_electricity_cost"]:.2f} 元')
    print(f'储能系统总造价: {cost_breakdown[optimal_power]["system_total_cost"]:.2f} 元')
    print(f'储能系统年摊销成本: {cost_breakdown[optimal_power]["annual_system_cost"]:.2f} 元')
    print(f'年度总成本(电费+系统摊销): {optimal_cost:.2f} 元')
    print("=" * 100)
    
    # 显示最优容量的详细电费分析
    print("\n最优容量的月度电费明细:")
    print("-" * 80)
    optimal_monthly = cost_breakdown[optimal_power]['monthly']
    period_names = {SHARP: "尖", PEAK: "峰", FLAT: "平", VALLEY: "谷", DEEP_VALLEY: "深谷"}
    
    for month in sorted(optimal_monthly.keys()):
        month_data = optimal_monthly[month]
        month_total = sum(month_data.values())
        
        print(f"\n{month}月电费统计:")
        for period, cost in month_data.items():
            if cost > 0:
                print(f"{period_names.get(period, f'未知{period}')}：{cost:.2f}元")
        print(f"月总电量电费：{month_total:.2f}元")
    
    # 绘制容量-成本曲线
    plt.figure(figsize=(12, 8))
    
    # 创建三条曲线
    plt.plot(power_range, [cost_breakdown[p]['annual_electricity_cost'] for p in power_range], 'g-o', label='年度电费')
    plt.plot(power_range, [cost_breakdown[p]['annual_system_cost'] for p in power_range], 'r-o', label='储能系统年摊销成本')
    plt.plot(power_range, costs, 'b-o', label='年度总成本')
    
    plt.title('储能系统功率-成本关系曲线')
    plt.xlabel('储能系统功率 (kW)')
    plt.ylabel('成本 (元)')
    plt.grid(True)
    plt.legend()
    
    # 标记最优点
    plt.plot(optimal_power, optimal_cost, 'mo', markersize=10)
    plt.annotate(f'最优功率: {optimal_power} kW\n最优容量: {optimal_capacity:.2f} kWh\n最低年度总成本: {optimal_cost:.2f} 元', 
                 xy=(optimal_power, optimal_cost), 
                 xytext=(optimal_power + 200, optimal_cost - 50000 if optimal_power < 1600 else optimal_cost + 50000),
                 arrowprops=dict(facecolor='black', shrink=0.05, width=1.5),
                 bbox=dict(boxstyle="round,pad=0.5", fc="yellow", alpha=0.8))
    
    # 添加第二个图表：展示最优功率下的成本构成
    plt.figure(figsize=(10, 6))
    
    # 获取最优功率下的各项成本
    opt_data = cost_breakdown[optimal_power]
    
    # 准备饼图数据
    labels = ['电量电费', '变压器基本电费', '储能系统摊销成本']
    sizes = [opt_data['electricity_cost'], opt_data['transformer_basic_cost'], opt_data['annual_system_cost']]
    colors = ['lightgreen', 'lightblue', 'coral']
    explode = (0.1, 0.1, 0.1)  # 突出显示所有部分
    
    # 绘制饼图
    plt.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')  # 确保饼图是圆的
    plt.title(f'最优容量({optimal_capacity:.2f} kWh)下的年度成本构成')
    
    plt.show()
    
    # 返回最优容量和功率，以便后续使用
    return {
        'optimal_power': optimal_power,
        'optimal_capacity': optimal_capacity,
        'annual_electricity_cost': cost_breakdown[optimal_power]["annual_electricity_cost"],
        'system_total_cost': cost_breakdown[optimal_power]["system_total_cost"],
        'annual_system_cost': cost_breakdown[optimal_power]["annual_system_cost"],
        'annual_total_cost': optimal_cost
    }
def get_latest_excel_files(directory):
    """获取目录中最新的Excel配置文件"""
    # 获取目录中所有Excel文件
    price_files = glob.glob(os.path.join(directory, '*price_config_*.xlsx'))
    period_files = glob.glob(os.path.join(directory, '*load_data_*.xlsx'))
    
    if not price_files or not period_files:
        print("未找到配置文件！")
        return None, None
    
    # 显示所有可用的配置文件
    print("\n可用的配置文件：")
    all_files = list(set([os.path.basename(f) for f in price_files + period_files]))
    for i, file in enumerate(all_files, 1):
        print(f"{i}. {file}")
    
    # 让用户选择文件
    while True:
        try:
            choice = input("\n请输入要使用的文件编号（用逗号分隔，或直接回车使用最新版本）: ").strip()
            if not choice:  # 如果用户直接回车，使用最新版本
                latest_price = max(price_files, key=os.path.basename)
                latest_period = max(period_files, key=os.path.basename)
                print(f"\n使用最新版本：{os.path.basename(latest_price)}, {os.path.basename(latest_period)}")
                return latest_price, latest_period
            
            choices = [int(c.strip()) for c in choice.split(',')]
            if len(choices) != 2:
                print("请确保选择两个文件编号！")
                continue
            
            selected_files = [all_files[i-1] for i in choices]
            selected_price = [f for f in price_files if os.path.basename(f) in selected_files]
            selected_period = [f for f in period_files if os.path.basename(f) in selected_files]
            
            if len(selected_price) == 1 and len(selected_period) == 1:
                print(f"\n已选择：{selected_price[0]}, {selected_period[0]}")
                return selected_price[0], selected_period[0]
            else:
                print("请确保选择一个价格配置文件和一个负载数据文件！")
        except ValueError:
            print("请输入有效的数字！")
            return None, None
def plot_monthly_price_curves(data):
    """
    绘制每个月的电价柱状图，支持前后翻页
    """
    period_colors = {
        1: 'orange',    # 尖峰
        2: 'pink',      # 高峰
        3: 'lightblue', # 平段
        4: 'lightgreen',# 低谷
        5: 'blue'       # 深谷
    }
    
    period_labels = {
        1: '尖峰时段',
        2: '高峰时段',
        3: '平段时段',
        4: '低谷时段',
        5: '深谷时段'
    }
    
    # 创建一个月份字典，存储每个月的数据
    monthly_data = {}
    
    # 从电价配置文件中直接读取时段和电价配置
    price_config_file = None
    for f in glob.glob(os.path.join("D:/小工具/cursor/装机容量测算", "*price_config_*.xlsx")):
        price_config_file = f
        break
    
    if price_config_file:
        try:
            price_config = pd.read_excel(price_config_file, sheet_name=['时段配置', '电价配置'])
            time_periods = price_config['时段配置']
            price_rules = price_config['电价配置']
            
            print(f"发现电价配置文件: {price_config_file}")
            print("正在从电价配置文件生成月度电价数据...")
            
            # 获取配置文件中的所有月份
            available_months = sorted(time_periods['月份'].unique())
            
            # 为每个月生成电价数据
            for month in available_months:
                periods_by_hour = {}
                prices_by_hour = {}
                
                # 从配置文件提取每小时的时段类型
                for hour in range(24):
                    column = f"{hour}-{hour+1}"
                    if column in time_periods.columns:
                        period_type = time_periods.loc[time_periods['月份'] == month, column].values[0]
                        periods_by_hour[hour] = period_type
                        
                        # 根据时段类型查找电价
                        price = price_rules.loc[price_rules['时段类型'] == period_type, '电价'].values[0]
                        prices_by_hour[hour] = price
                
                # 确保有完整的24小时数据
                if len(periods_by_hour) == 24:
                    monthly_data[month] = {
                        'hours': range(24),
                        'prices': [prices_by_hour[h] for h in range(24)],
                        'period_types': [periods_by_hour[h] for h in range(24)]
                    }
                else:
                    print(f"警告: 月份 {month} 的配置数据不完整，只有 {len(periods_by_hour)} 小时的数据")
        except Exception as e:
            print(f"从配置文件生成电价数据失败: {e}")
            print("将尝试从负载数据中提取电价信息...")
    
    # 如果从配置文件没有成功生成数据，则尝试从负载数据中提取
    if not monthly_data:
        # 首先生成所有月份的数据
        for month in range(1, 13):
            # 创建一个完整的日期范围，确保每个月都有完整的24小时数据
            month_start = pd.Timestamp(f'2024-{month:02d}-01')
            if month == 12:
                next_month_start = pd.Timestamp('2025-01-01')
            else:
                next_month_start = pd.Timestamp(f'2024-{month+1:02d}-01')
            
            # 获取这个月的所有数据
            month_data = data[(data['datetime'] >= month_start) & (data['datetime'] < next_month_start)]
            
            # 如果这个月没有数据，则继续下一个月
            if month_data.empty:
                print(f"月份 {month} 没有数据，跳过绘图。")
                continue
            
            # 尝试构建完整的24小时数据
            hours = range(24)
            prices = []
            period_types = []
            
            # 为每小时提取数据
            for hour in hours:
                hour_data = month_data[month_data['datetime'].dt.hour == hour]
                if not hour_data.empty:
                    prices.append(hour_data['price'].iloc[0])
                    period_types.append(hour_data['period_type'].iloc[0])
                else:
                    # 如果该小时没有数据，尝试从其他日期的同一小时获取
                    # 或者使用默认值
                    other_hour_data = data[data['datetime'].dt.hour == hour]
                    if not other_hour_data.empty:
                        prices.append(other_hour_data['price'].iloc[0])
                        period_types.append(other_hour_data['period_type'].iloc[0])
                    else:
                        prices.append(0.5)  # 默认电价
                        period_types.append(3)  # 默认为平段
            
            # 存储月度数据
            monthly_data[month] = {
                'hours': hours,
                'prices': prices,
                'period_types': period_types
            }
    
    # 如果没有任何月份数据，则直接返回
    if not monthly_data:
        print("没有找到任何月份的完整数据，无法绘图。")
        return
    
    # 当前显示的月份索引
    current_month_idx = 0
    available_months = sorted(monthly_data.keys())
    
    # 创建图形
    fig, ax = plt.subplots(figsize=(12, 6))
    
    # 绘制指定月份的图表
    def plot_month(month):
        ax.clear()
        data = monthly_data[month]
        hours = data['hours']
        prices = data['prices']
        period_types = data['period_types']
        
        # 创建与第四个图表一致的横坐标标签
        hour_labels = [f"{h}" for h in hours]
        
        bars = ax.bar(hours, prices)
        
        for hour, bar, period_type in zip(hours, bars, period_types):
            if np.isnan(period_type):
                bar.set_color('gray')
            else:
                bar.set_color(period_colors.get(int(period_type), 'gray'))
        
        # 添加图例
        unique_period_types = np.unique(period_types)
        if hasattr(unique_period_types, '__iter__'):  # 确保是可迭代对象
            # 移除NaN值
            unique_period_types = [pt for pt in unique_period_types if not np.isnan(pt)]
            # 转换为整数
            unique_period_types = [int(pt) for pt in unique_period_types]
        
        legend_elements = [plt.Rectangle((0,0),1,1, color=period_colors.get(pt, 'gray'), label=period_labels.get(pt, f'未知{pt}')) 
                          for pt in sorted(unique_period_types) if pt in period_colors]
        
        ax.legend(handles=legend_elements)
        ax.set_title(f'电价曲线，月份: {month}')
        ax.set_xlabel('小时')
        ax.set_ylabel('电价 (元/kWh)')
        ax.set_xticks(hours)
        ax.set_xticklabels(hour_labels)
        ax.grid(True)
        fig.canvas.draw()
    
    # 添加按钮的回调函数
    def on_prev(event):
        nonlocal current_month_idx
        current_month_idx = (current_month_idx - 1) % len(available_months)
        plot_month(available_months[current_month_idx])
    
    def on_next(event):
        nonlocal current_month_idx
        current_month_idx = (current_month_idx + 1) % len(available_months)
        plot_month(available_months[current_month_idx])
    
    # 添加按钮
    plt.subplots_adjust(bottom=0.2)
    ax_prev = plt.axes([0.7, 0.05, 0.1, 0.075])
    ax_next = plt.axes([0.81, 0.05, 0.1, 0.075])
    btn_prev = Button(ax_prev, '上一月')
    btn_next = Button(ax_next, '下一月')
    btn_prev.on_clicked(on_prev)
    btn_next.on_clicked(on_next)
    
    # 显示第一个月份的图表
    plot_month(available_months[current_month_idx])
    plt.show()
def main():
    # 修改默认的文件路径
    default_directory = "D:/小工具/cursor/装机容量测算"
    
    # 获取配置文件路径
    price_file, period_file = get_latest_excel_files(default_directory)
    if not price_file or not period_file:
        print("无法继续执行，请确保配置文件存在！")
        return
    try:
        # 加载数据
        print("\n正在加载实际数据...")
        data = load_data(period_file, price_file)  # 确保参数顺序正确
        print("数据加载成功！")
        
        # 保存最优容量分析结果
        optimal_results = None
        original_annual_cost = None
        
        while True:
            print("\n请选择要执行的操作:")
            print("1. 绘制每个月的电价柱状图")
            print("2. 绘制负载和电价曲线")
            print("3. 计算不安装储能系统的年度总电费")
            print("4. 储能系统分析")
            print("5. 寻找最佳储能系统容量")
            if optimal_results is not None:
                print("6. 使用最佳容量进行储能系统分析和节省计算")
                print("7. 最佳容量与自定义容量对比分析")
            print("0. 退出")
            
            choice = input("请输入选项的编号: ")
            
            if choice == '1':
                plot_monthly_price_curves(data)
            elif choice == '2':
                date = input("请输入要绘制的日期(格式: YYYY-MM-DD)，直接回车自动滚动显示所有日期: ")
                if date.strip():
                    try:
                        plot_load_and_price(data, date, price_file)
                    except Exception as e:
                        print(f"绘制失败: {e}")
                        traceback.print_exc()
                else:
                    plot_load_and_price(data, None, price_file)
            elif choice == '3':
                try:
                    # 计算原始年度总电费（不使用储能系统）
                    original_cost = calculate_annual_cost(data)
                    
                    # 保存原始总电费用于后续比较
                    if original_cost is not None:
                        original_annual_cost = original_cost
                        print(f"已保存总电费: {original_annual_cost:.2f} 元")
                except Exception as e:
                    print(f"计算失败: {e}")
                    traceback.print_exc()
            elif choice == '4':
                print("\n储能系统分析选项:")
                print("1. 图形界面分析（可逐日查看）")
                print("2. 直接输出所有日期的分析结果和月度统计")
                sub_choice = input("请选择: ")
                
                try:
                    if sub_choice == '1':
                        plot_storage_system(data)
                        print("储能系统图形分析完成！")
                    elif sub_choice == '2':
                        plot_storage_system(data, auto_analyze=True)
                        print("储能系统批量分析完成！")
                    else:
                        print("无效的选项，返回主菜单")
                except Exception as e:
                    print(f"储能系统分析失败: {e}")
                    traceback.print_exc()
            elif choice == '5':
                try:
                    # 寻找最佳容量并保存结果
                    optimal_results = find_optimal_storage_capacity(data)
                    print("最佳储能系统容量分析完成！")
                    if original_annual_cost is None:
                        print("提示: 您尚未保存原始年度总电费。如需计算节省的电费，请先运行选项3计算并保存原始总电费。")
                    print("您现在可以选择选项 6 使用此最佳容量进行进一步分析和电费节省计算。")
                except Exception as e:
                    print(f"寻找最佳容量失败: {e}")
                    traceback.print_exc()
            elif choice == '6' and optimal_results is not None:
                if original_annual_cost is None:
                    save_original_cost = input("尚未保存原始年度总电费，是否现在计算并保存？(y/n): ")
                    if save_original_cost.lower() == 'y':
                        try:
                            # 计算原始年度总电费（不使用储能系统）
                            original_cost = calculate_annual_cost(data)
                            if original_cost is not None:
                                original_annual_cost = original_cost
                                print(f"已保存原始年度总电费: {original_annual_cost:.2f} 元")
                        except Exception as e:
                            print(f"计算失败: {e}")
                            traceback.print_exc()
                
                try:
                    # 使用最佳容量进行分析和节省计算
                    analyze_with_optimal_capacity(data, optimal_results, original_annual_cost)
                except Exception as e:
                    print(f"使用最佳容量分析失败: {e}")
                    traceback.print_exc()
            elif choice == '7' and optimal_results is not None:
                if original_annual_cost is None:
                    save_original_cost = input("尚未保存原始年度总电费，是否现在计算并保存？(y/n): ")
                    if save_original_cost.lower() == 'y':
                        try:
                            # 计算原始年度总电费（不使用储能系统）
                            original_cost = calculate_annual_cost(data)
                            if original_cost is not None:
                                original_annual_cost = original_cost
                                print(f"已保存原始年度总电费: {original_annual_cost:.2f} 元")
                        except Exception as e:
                            print(f"计算失败: {e}")
                            traceback.print_exc()
                
                try:
                    # 调用对比分析功能
                    compare_storage_capacities(data, optimal_results, original_annual_cost)
                except Exception as e:
                    print(f"对比分析失败: {e}")
                    traceback.print_exc()
            elif choice == '0':
                break
            else:
                print("无效的选项，请重新选择")
    except Exception as e:
        print(f"执行过程中发生错误: {e}")
        print("详细错误信息:")
        print(traceback.format_exc())
def analyze_with_optimal_capacity(data, optimal_results, original_annual_cost=None):
    """使用最佳储能容量分析系统运行情况并计算节省的电费
    
    Parameters:
    -----------
    data : DataFrame
        负载和电价数据
    optimal_results : dict
        包含最佳功率和容量的字典，由find_optimal_storage_capacity函数返回
    original_annual_cost : float, optional
        不装储能系统时的年度总电费，默认为None
    """
    # 提取最佳功率和容量
    optimal_power = optimal_results['optimal_power']
    optimal_capacity = optimal_results['optimal_capacity']
    optimal_annual_cost = optimal_results['annual_electricity_cost']  # 只有电费部分
    
    # 暂时更改全局变量以使用最佳容量
    global storage_capacity_per_system, max_power_per_system
    original_capacity = storage_capacity_per_system
    original_power = max_power_per_system
    
    # 设置为最佳值
    storage_capacity_per_system = optimal_capacity
    max_power_per_system = optimal_power
    
    # 使用最佳容量运行储能系统分析（自动分析模式）
    print("\n" + "=" * 80)
    print(f"使用最佳储能系统进行分析：")
    print(f"功率: {optimal_power} kW")
    print(f"容量: {optimal_capacity:.2f} kWh")
    
    # 先进行自动分析以获取详细结果
    print("\n开始以最佳容量运行储能系统分析...")
    try:
        plot_storage_system(data, auto_analyze=True)
    except Exception as e:
        print(f"分析过程中出错: {e}")
        traceback.print_exc()
    
    # 输出电费节省信息
    print("\n" + "=" * 80)
    print("储能系统电费节省分析结果：")
    print("-" * 50)
    
    # 如果提供了原始年度总电费，计算节省的电费
    if original_annual_cost is not None:
        annual_savings = original_annual_cost - optimal_annual_cost
        savings_percentage = (annual_savings / original_annual_cost) * 100
        
        print(f"原始年度总电费（不安装储能系统）: {original_annual_cost:.2f} 元")
        print(f"使用最佳储能系统后年度总电费: {optimal_annual_cost:.2f} 元")
        print(f"年度节省电费: {annual_savings:.2f} 元")
        print(f"电费节省比例: {savings_percentage:.2f}%")
    else:
        print("未提供原始年度总电费，无法计算节省比例。")
        print(f"使用最佳储能系统后年度总电费: {optimal_annual_cost:.2f} 元")
    
    print("-" * 50)
    print("备注: 上述节省金额仅考虑电费部分，未包含储能系统投资和维护成本。")
    print("=" * 80)
    
    # 使用最佳容量进行图形化分析（可视化模式）
    print("\n正在绘制储能系统图形分析...")
    try:
        plot_storage_system(data)
    except Exception as e:
        print(f"绘图过程中出错: {e}")
        traceback.print_exc()
    
    # 恢复原始值
    storage_capacity_per_system = original_capacity
    max_power_per_system = original_power
    
    print("\n分析完成！已恢复原始储能系统参数。")

def compare_storage_capacities(data, optimal_results, original_annual_cost=None):
    """比较最佳容量与用户指定容量的效果
    
    Parameters:
    -----------
    data : DataFrame
        负载和电价数据
    optimal_results : dict
        包含最佳功率和容量的字典，由find_optimal_storage_capacity函数返回
    original_annual_cost : float, optional
        不装储能系统时的年度总电费，默认为None
    """
    if optimal_results is None:
        print("错误：需要先运行选项5查找最佳容量才能进行比较分析。")
        return
        
    # 提取最佳功率和容量
    optimal_power = optimal_results['optimal_power']
    optimal_capacity = optimal_results['optimal_capacity']
    optimal_annual_cost = optimal_results['annual_electricity_cost']
    
    # 用户输入需要比较的容量
    try:
        print("\n" + "=" * 80)
        print("最佳容量与自定义容量比较分析")
        print("-" * 50)
        print(f"最佳容量：{optimal_capacity:.2f} kWh (功率：{optimal_power} kW)")
        
        # 储能系统容量与功率的比例
        capacity_power_ratio = optimal_capacity / optimal_power
        
        custom_capacity = float(input("请输入需要比较的储能容量(kWh): "))
        # 根据比例计算对应功率
        custom_power = round(custom_capacity / capacity_power_ratio)
        print(f"根据比例计算的对应功率：{custom_power} kW")
        
        # 暂存原始容量设置
        global storage_capacity_per_system, max_power_per_system
        original_capacity = storage_capacity_per_system
        original_power = max_power_per_system
        
        # 分析结果储存
        results = {
            "optimal": {"capacity": optimal_capacity, "power": optimal_power, "annual_cost": optimal_annual_cost},
            "custom": {}
        }
        
        # 1. 对最佳容量进行分析
        storage_capacity_per_system = optimal_capacity
        max_power_per_system = optimal_power
        
        print("\n正在分析最佳容量的效果...")
        
        # 为最佳容量运行相同的分析
        optimal_analysis = analyze_storage_capacity(data)
        results["optimal"].update(optimal_analysis)
        
        # 2. 对自定义容量进行分析
        storage_capacity_per_system = custom_capacity
        max_power_per_system = custom_power
        
        print("\n正在分析自定义容量的效果...")
        
        # 运行相同的分析逻辑
        custom_analysis = analyze_storage_capacity(data)
        results["custom"] = {
            "capacity": custom_capacity,
            "power": custom_power,
            "annual_cost": custom_analysis["annual_cost"]
        }
        results["custom"].update(custom_analysis)
        
        # 恢复原始值
        storage_capacity_per_system = original_capacity
        max_power_per_system = original_power
        
        # 输出比较结果
        print("\n" + "=" * 80)
        print("最佳容量与自定义容量比较分析结果")
        print("-" * 50)
        
        # 容量和功率比较
        print(f"最佳容量: {optimal_capacity:.2f} kWh (功率: {optimal_power} kW)")
        print(f"自定义容量: {custom_capacity:.2f} kWh (功率: {custom_power} kW)")
        print("-" * 50)
        
        # 电费比较
        if original_annual_cost is not None:
            optimal_savings = original_annual_cost - optimal_annual_cost
            optimal_savings_percentage = (optimal_savings / original_annual_cost) * 100
            
            custom_savings = original_annual_cost - results["custom"]["annual_cost"]
            custom_savings_percentage = (custom_savings / original_annual_cost) * 100
            
            print(f"原始年度总电费（不安装储能系统）: {original_annual_cost:.2f} 元")
            print(f"最佳容量年度总电费: {optimal_annual_cost:.2f} 元")
            print(f"最佳容量节省电费: {optimal_savings:.2f} 元 ({optimal_savings_percentage:.2f}%)")
            print(f"自定义容量年度总电费: {results['custom']['annual_cost']:.2f} 元")
            print(f"自定义容量节省电费: {custom_savings:.2f} 元 ({custom_savings_percentage:.2f}%)")
        else:
            print("未提供原始年度总电费，无法计算节省比例。")
            print(f"最佳容量年度总电费: {optimal_annual_cost:.2f} 元")
            print(f"自定义容量年度总电费: {results['custom']['annual_cost']:.2f} 元")
        
        print("-" * 50)
        
        # 充放电策略有效性分析对比
        print("充放电策略有效性分析对比:")
        print(f"【最佳容量 {optimal_capacity:.2f} kWh】:")
        print(f"  总计分析天数: {results['optimal']['total_days']} 天")
        print(f"  有效放电天数: {results['optimal']['days_with_discharge']} 天 ({results['optimal']['discharge_rate']*100:.2f}%)")
        print(f"  有效放电总时长: {results['optimal']['total_discharge_hours']} 小时")
        print(f"  平均每天放电时长: {results['optimal']['avg_discharge_hours']:.2f} 小时")
        print(f"  容量浪费天数: {results['optimal']['wasted_capacity_days']} 天 ({results['optimal']['wasted_capacity_rate']*100:.2f}%)")
        print(f"  容量不足天数: {results['optimal']['insufficient_capacity_days']} 天 ({results['optimal']['insufficient_capacity_rate']*100:.2f}%)")
        
        # 添加最佳容量结论
        if results['optimal']['wasted_capacity_rate'] > 0.3 and results['optimal']['insufficient_capacity_rate'] < 0.1:
            optimal_conclusion = "容量可能过大，存在浪费"
        elif results['optimal']['wasted_capacity_rate'] < 0.1 and results['optimal']['insufficient_capacity_rate'] > 0.3:
            optimal_conclusion = "容量可能不足，无法满足放电需求"
        elif results['optimal']['wasted_capacity_rate'] < 0.2 and results['optimal']['insufficient_capacity_rate'] < 0.2:
            optimal_conclusion = "容量匹配良好，符合企业用电需求"
        else:
            optimal_conclusion = "容量利用率一般，建议进一步优化"
        print(f"  结论: {optimal_conclusion}")
        
        print(f"\n【自定义容量 {custom_capacity:.2f} kWh】:")
        print(f"  总计分析天数: {results['custom']['total_days']} 天")
        print(f"  有效放电天数: {results['custom']['days_with_discharge']} 天 ({results['custom']['discharge_rate']*100:.2f}%)")
        print(f"  有效放电总时长: {results['custom']['total_discharge_hours']} 小时")
        print(f"  平均每天放电时长: {results['custom']['avg_discharge_hours']:.2f} 小时")
        print(f"  容量浪费天数: {results['custom']['wasted_capacity_days']} 天 ({results['custom']['wasted_capacity_rate']*100:.2f}%)")
        print(f"  容量不足天数: {results['custom']['insufficient_capacity_days']} 天 ({results['custom']['insufficient_capacity_rate']*100:.2f}%)")
        
        # 添加自定义容量结论
        if results['custom']['wasted_capacity_rate'] > 0.3 and results['custom']['insufficient_capacity_rate'] < 0.1:
            custom_conclusion = "容量可能过大，存在浪费"
        elif results['custom']['wasted_capacity_rate'] < 0.1 and results['custom']['insufficient_capacity_rate'] > 0.3:
            custom_conclusion = "容量可能不足，无法满足放电需求"
        elif results['custom']['wasted_capacity_rate'] < 0.2 and results['custom']['insufficient_capacity_rate'] < 0.2:
            custom_conclusion = "容量匹配良好，符合企业用电需求"
        else:
            custom_conclusion = "容量利用率一般，建议进一步优化"
        print(f"  结论: {custom_conclusion}")
        
        # 综合推荐
        print("\n" + "=" * 80)
        print("综合分析推荐:")
        
        # 综合考虑经济性和容量浪费率，推荐最佳容量
        if original_annual_cost is not None:
            if custom_savings > optimal_savings and results['custom']['wasted_capacity_rate'] < 0.3:
                print(f"推荐选择自定义容量 {custom_capacity:.2f} kWh，可节省更多电费且容量浪费率可接受。")
                print(f"相比最佳容量，额外节省电费: {custom_savings - optimal_savings:.2f} 元")
            elif optimal_savings > custom_savings and results['custom']['insufficient_capacity_rate'] > 0.3:
                print(f"推荐选择最佳容量 {optimal_capacity:.2f} kWh，尽管自定义容量节省电费较少，但可避免容量不足问题。")
                print(f"相比自定义容量，额外节省电费: {optimal_savings - custom_savings:.2f} 元")
            elif abs(custom_savings - optimal_savings) < original_annual_cost * 0.01:  # 差异小于1%
                if results['custom']['wasted_capacity_rate'] < results['custom']['insufficient_capacity_rate']:
                    print(f"两种容量节省电费相近，但自定义容量 {custom_capacity:.2f} kWh 容量利用率更高，推荐使用。")
                else:
                    print(f"两种容量节省电费相近，但最佳容量 {optimal_capacity:.2f} kWh 更能满足放电需求，推荐使用。")
            else:
                # 选择节省电费最多的方案
                if custom_savings > optimal_savings:
                    print(f"从经济角度考虑，推荐选择自定义容量 {custom_capacity:.2f} kWh，可节省更多电费。")
                else:
                    print(f"从经济角度考虑，推荐选择最佳容量 {optimal_capacity:.2f} kWh，可节省更多电费。")
        else:
            if results["custom"]["annual_cost"] < optimal_annual_cost:
                print(f"从经济角度考虑，推荐选择自定义容量 {custom_capacity:.2f} kWh，电费更低。")
            else:
                print(f"从经济角度考虑，推荐选择最佳容量 {optimal_capacity:.2f} kWh，电费更低。")
        
        print("-" * 50)
        print("备注: 上述节省金额仅考虑电费部分，未包含储能系统投资和维护成本。")
        print("=" * 80)
        
    except Exception as e:
        print(f"比较分析过程中出错: {e}")
        traceback.print_exc()
        
        # 确保恢复原始值
        storage_capacity_per_system = original_capacity
        max_power_per_system = original_power

# 将原有的分析逻辑提取为单独的函数，以便对两种容量进行相同的分析
def analyze_storage_capacity(data):
    """分析指定容量的储能系统效果
    
    Parameters:
    -----------
    data : DataFrame
        负载和电价数据
        
    Returns:
    --------
    dict
        包含分析结果的字典
    """
    # 定义变量用于存储分析结果
    total_days = 0
    days_with_discharge = 0
    total_discharge_hours = 0
    total_hours = 0
    wasted_capacity_days = 0
    insufficient_capacity_days = 0
    
    # 确保数据按时间排序
    data_sorted = data.sort_values('datetime')
    
    # 计算所有日期的电费
    monthly_costs = {}
    monthly_max_loads = {}
    
    # 月度价格计算所需的变量
    monthly_period_costs = {}
    
    # 按日期进行分析
    for date in sorted(data_sorted['datetime'].dt.date.unique()):
        # 获取当天数据
        day_data = data_sorted[data_sorted['datetime'].dt.date == date].copy()
        
        # 如果当天没有数据，创建一个全零的数据集
        if day_data.empty:
            continue
        
        # 如果数据不足24小时，补充缺失的小时
        if len(day_data) < 24:
            # 找出当前有哪些小时
            existing_hours = day_data['hour'].unique()
            missing_hours = set(range(24)) - set(existing_hours)
            
            if missing_hours:
                month = day_data['month'].iloc[0]
                date_str = pd.Timestamp(date).strftime('%Y-%m-%d')
                
                # 为缺失的小时创建数据
                missing_data = []
                for hour in missing_hours:
                    # 查找同月同小时的数据作为参考
                    ref_data = data_sorted[(data_sorted['datetime'].dt.month == month) & 
                                      (data_sorted['datetime'].dt.hour == hour)]
                    
                    if not ref_data.empty:
                        period_type = ref_data['period_type'].iloc[0]
                        price = ref_data['price'].iloc[0]
                    else:
                        period_type = FLAT
                        price = 0.5
                    
                    missing_data.append({
                        'datetime': pd.Timestamp(date_str) + pd.Timedelta(hours=hour),
                        'load': 0,  # 缺失小时负载为0
                        'price': price,
                        'period_type': period_type,
                        'month': month,
                        'hour': hour
                    })
                
                # 添加缺失小时数据
                missing_df = pd.DataFrame(missing_data)
                day_data = pd.concat([day_data, missing_df])
                
            # 确保按小时排序
            day_data = day_data.sort_values('hour').reset_index(drop=True)
        
        month = day_data['month'].iloc[0]
        periods = day_data['period_type'].values
        
        # 处理可能的NaN值
        if np.isnan(periods).any():
            periods = np.where(np.isnan(periods), FLAT, periods)
            
        # 执行套利分析和模拟
        # 第一次套利（简化版）
        charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
        if charge1_start is None:
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
        
        if charge1_start is not None:
            discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
            if discharge1_start is None:
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
            
            if discharge1_start is not None:
                # 增加天数计数
                total_days += 1
                total_hours += len(day_data)
                
                # 确保索引不会超出范围
                charge1_end = min(charge1_start + charge1_duration, 24)
                discharge1_end = min(discharge1_start + discharge1_duration, 24)
                
                # 模拟充电
                storage_level = initial_storage_capacity
                original_load = day_data['load'].values
                modified_load = original_load.copy()
                
                # 确保负载数组长度为24
                if len(original_load) < 24:
                    original_load = np.pad(original_load, (0, 24 - len(original_load)), 'constant')
                    modified_load = original_load.copy()
                elif len(original_load) > 24:
                    original_load = original_load[:24]
                    modified_load = original_load.copy()
                
                # 充电
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / ((charge1_end - charge1_start) * efficiency_bess))
                for h in range(charge1_start, charge1_end):
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
                
                # 放电
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / (discharge1_end - discharge1_start))
                
                discharge_happened = False
                discharge_hours = 0
                
                for h in range(discharge1_start, discharge1_end):
                    # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
                    
                    if actual_discharge > 0:
                        discharge_happened = True
                        discharge_hours += 1
                
                if discharge_happened:
                    days_with_discharge += 1
                    total_discharge_hours += discharge_hours
                
                # 检查容量利用情况
                if storage_level > 0.3 * storage_capacity_per_system:  # 容量利用率低于70%
                    wasted_capacity_days += 1
                
                # 检查是否有未满足的放电需求
                if discharge_happened and discharge_hours < (discharge1_end - discharge1_start):
                    insufficient_capacity_days += 1
                
                # 计算电费
                costs = day_data['price'].values * modified_load
                daily_cost = sum(costs)
                month_key = f"{month:02d}"
                
                if month_key not in monthly_costs:
                    monthly_costs[month_key] = 0
                    monthly_max_loads[month_key] = 0
                
                monthly_costs[month_key] += daily_cost
                # 更新月度最大负载（用于计算变压器基本电费）
                month_max_load = max(modified_load)
                if month_max_load > monthly_max_loads[month_key]:
                    monthly_max_loads[month_key] = month_max_load
    
    # 计算总电费
    total_electricity_cost = sum(monthly_costs.values())
    
    # 如果使用变压器基本电费计算方法
    if method_basic_capacity_cost_transformer == 1:  # 按容量收取
        # 获取全局变量中的容量单价
        if 'capacity_price' in globals():
            capacity_price = globals()['capacity_price']
        else:
            capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
            globals()['capacity_price'] = capacity_price
        
        # 计算变压器基本电费（容量单价乘以变压器容量）
        transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        annual_cost = transformer_basic_cost + total_electricity_cost
        
    elif method_basic_capacity_cost_transformer == 2:  # 按需收取
        # 获取全局变量中的需量单价
        if 'demand_price' in globals():
            demand_price = globals()['demand_price']
        else:
            demand_price = float(input('请输入需量单价（元/kW·月）: '))
            globals()['demand_price'] = demand_price
        
        # 计算变压器基本电费
        transformer_basic_cost = sum(monthly_max_loads.values()) * demand_price
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        annual_cost = transformer_basic_cost + total_electricity_cost
        
    else:  # 不计算变压器基本容量费用
        annual_cost = total_electricity_cost
    
    # 返回分析结果
    return {
        "annual_cost": annual_cost,
        "total_days": total_days,
        "days_with_discharge": days_with_discharge,
        "discharge_rate": days_with_discharge / total_days if total_days > 0 else 0,
        "total_discharge_hours": total_discharge_hours,
        "avg_discharge_hours": total_discharge_hours / days_with_discharge if days_with_discharge > 0 else 0,
        "wasted_capacity_days": wasted_capacity_days,
        "wasted_capacity_rate": wasted_capacity_days / total_days if total_days > 0 else 0,
        "insufficient_capacity_days": insufficient_capacity_days,
        "insufficient_capacity_rate": insufficient_capacity_days / total_days if total_days > 0 else 0
    }

if __name__ == "__main__":
    main() 
