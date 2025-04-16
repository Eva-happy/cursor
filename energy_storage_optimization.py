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
    load_file: 负载数据Excel文件路径（包含date, hour/min, load三列）
    price_config_file: 电价配置Excel文件路径
    """
    try:
        # 读取负载数据
        load_df = pd.read_excel(load_file)
        
        # 确保列名正确
        required_columns = ['date', 'load']
        if not all(col in load_df.columns for col in required_columns):
            raise ValueError("负载数据文件必须包含'date'和'load'列，以及'hour'或'min'列")
        
        # 检查是否是小时级别还是分钟级别的数据
        is_minute_data = 'min' in load_df.columns
        time_column = 'min' if is_minute_data else 'hour'
        
        if time_column not in load_df.columns:
            raise ValueError(f"负载数据文件缺少时间列，需要'hour'或'min'列")
        
        # 检查数据是否包含NaN值
        if load_df['load'].isna().any():
            print("警告：负载数据中包含无效值(NaN)，将使用0替换。")
            load_df['load'] = load_df['load'].fillna(0)
        
        # 将date和hour/min合并成datetime
        if is_minute_data:
            # 对于分钟数据，处理可能的"时:分"格式
            if isinstance(load_df['min'].iloc[0], str) and ":" in str(load_df['min'].iloc[0]):
                # 处理"时:分"格式的数据
                print("检测到'时:分'格式的分钟数据...")
                
                # 提取小时和分钟
                def extract_hour_min(time_str):
                    if pd.isna(time_str):
                        return 0, 0
                    parts = str(time_str).split(':')
                    hour = int(parts[0])
                    minute = int(parts[1]) if len(parts) > 1 else 0
                    return hour, minute
                
                # 应用函数提取小时和分钟
                hour_min_pairs = load_df['min'].apply(extract_hour_min)
                load_df['hour'] = hour_min_pairs.apply(lambda x: x[0])
                load_df['minute'] = hour_min_pairs.apply(lambda x: x[1])
                
                # 创建datetime对象
                load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['hour'], unit='h') + pd.to_timedelta(load_df['minute'], unit='m')
            else:
                # 处理纯分钟数的数据
                load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['min'], unit='m')
                # 添加小时列用于后续处理
                load_df['hour'] = load_df['datetime'].dt.hour
                load_df['minute'] = load_df['datetime'].dt.minute
        else:
            # 对于小时数据，保持原有的处理方式
            load_df['datetime'] = pd.to_datetime(load_df['date']) + pd.to_timedelta(load_df['hour'], unit='h')
            load_df['minute'] = 0  # 添加分钟列，默认为0
        
        # 读取电价配置
        price_config = pd.read_excel(price_config_file, sheet_name=['时段配置', '电价配置'])
        
        # 获取时段配置和电价规则
        time_periods = price_config['时段配置']
        price_rules = price_config['电价配置']
        
        # 添加月份信息
        load_df['month'] = load_df['datetime'].dt.month
        
        # 为不同粒度的数据添加小时信息（确保有hour列）
        if 'hour' not in load_df.columns:
            load_df['hour'] = load_df['datetime'].dt.hour
            
        # 添加15分钟时段的索引（0-3表示每小时内的四个15分钟）
        load_df['quarter_idx'] = (load_df['datetime'].dt.minute // 15)
        
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
        
        # 添加数据粒度标记
        load_df['is_minute_data'] = is_minute_data
        
        # 只保留需要的列
        if is_minute_data:
            result_df = load_df[['datetime', 'load', 'price', 'period_type', 'month', 'hour', 'minute', 'quarter_idx', 'is_minute_data']]
        else:
            result_df = load_df[['datetime', 'load', 'price', 'period_type', 'month', 'hour', 'is_minute_data']]
        
        return result_df
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
    
    # 检查数据是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
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
                if is_minute_data:
                    # 对于分钟数据，按小时聚合电价和时段类型（取众数）
                    hour_data = day_data.groupby('hour').agg({
                        'price': 'mean',
                        'period_type': lambda x: x.mode()[0] if not x.mode().empty else None
                    })
                    prices = hour_data['price'].values
                    period_types = hour_data['period_type'].values
                else:
                    # 使用数据中的电价和时段类型
                    prices = day_data['price'].values[:24]
                    period_types = day_data['period_type'].values[:24]
        else:
            # 使用数据中的电价和时段类型
            if is_minute_data:
                # 对于分钟数据，按小时聚合电价和时段类型
                hour_data = day_data.groupby('hour').agg({
                    'price': 'mean',
                    'period_type': lambda x: x.mode()[0] if not x.mode().empty else None
                })
                prices = hour_data['price'].values
                period_types = hour_data['period_type'].values
            else:
                prices = day_data['price'].values[:24]
                period_types = day_data['period_type'].values[:24]
        
        # 创建横坐标的时间区间标签
        hours = range(24)
        hour_labels = [f"{h}-{h+1}" for h in hours]
        
        # 清空当前图表
        ax1.clear()
        ax2.clear()
        
        # 获取负载数据
        if is_minute_data:
            # 对于分钟数据，按小时聚合负载
            hour_loads = day_data.groupby('hour')['load'].mean().values
            
            # 创建15分钟粒度的横坐标
            quarter_times = []
            quarter_loads = []
            quarter_x = []
            
            for i, row in day_data.iterrows():
                hour = row['hour']
                # 考虑可能的"时:分"格式
                if 'minute' in row:
                    minute = row['minute']
                else:
                    minute = row['datetime'].minute
                quarter_times.append(f"{hour}:{minute:02d}")
                quarter_loads.append(row['load'])
                quarter_x.append(hour + minute/60)
            
            # 绘制小时级别的柱状图
            bars1 = ax1.bar(hours, hour_loads, color='lightblue', alpha=0.5)
            
            # 叠加15分钟粒度的线图
            ax1_2 = ax1.twinx()
            ax1_2.plot(quarter_x, quarter_loads, 'r-', linewidth=1.5)
            ax1_2.set_ylabel('15分钟负载 (kW)', color='r')
            ax1_2.tick_params(axis='y', labelcolor='r')
            ax1_2.set_ylim(0, max(quarter_loads) * 1.1)
            
            # 隐藏右侧的y轴刻度，只保留左侧的小时级别刻度
            ax1_2.set_yticks([])
        else:
            # 对于小时数据，直接使用
            loads = day_data['load'].values[:24]
            
            # 绘制负载柱状图
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
        if is_minute_data:
            # 对于分钟数据，计算每个15分钟段的电费，并按时段类型分组
            total_load = day_data['load'].sum()
            cost_by_period = {}
            for i, row in day_data.iterrows():
                period_type = row['period_type']
                if period_type not in cost_by_period:
                    cost_by_period[period_type] = 0
                cost_by_period[period_type] += row['load'] * row['price']
        else:
            # 对于小时数据，保持原有的计算方式
            total_load = sum(loads)
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
    # 检查数据是否包含NaN值
    if data['load'].isna().any() or data['price'].isna().any():
        print("警告：数据中包含无效值(NaN)，这可能导致计算结果不准确。")
        print("正在尝试处理无效值...")
        
        # 使用0替换NaN值
        data_clean = data.copy()
        data_clean['load'] = data_clean['load'].fillna(0)
        data_clean['price'] = data_clean['price'].fillna(0)
        data = data_clean
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    # 计算年度总耗电量（全年的load相加）
    annual_total_load = sum(data['load'])
    print(f'年度总耗电量: {annual_total_load:.2f} KWh')
    
    # 计算各时段电费总和
    total_electricity_cost = sum(data['load'] * data['price'])
    
    # 根据客户选择的变压器基本容量费用计算方法计算
    if method_basic_capacity_cost_transformer == 1:
        # 按容量收取
        print('变压器基本电费，按变压器容量收取！')
        
        # 获取容量单价
        capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
        
        # 计算变压器基本电费（容量单价乘以变压器容量）
        transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        total_cost = transformer_basic_cost + total_electricity_cost
        
        print(f'变压器基本电费: {transformer_basic_cost:.2f} 元')
        print(f'电量电费: {total_electricity_cost:.2f} 元')
        print(f'年度总电费（含变压器基本电费和电量电费）: {total_cost:.2f} 元')
    
    elif method_basic_capacity_cost_transformer == 2:
        # 按需收取
        print('变压器基本电费，按需收取！')
        
        # 获取需量单价
        demand_price = float(input('请输入需量单价（元/kW·月）: '))
        
        # 计算每个月最大负载
        if is_minute_data:
            # 对于分钟数据，首先按照15分钟为单位找到每个月的最大负载
            monthly_max_loads = data.groupby('month')['load'].max()
        else:
            # 对于小时数据，保持原有的处理方式
            monthly_max_loads = data.groupby('month')['load'].max()
        
        # 计算变压器基本电费（需量单价乘以每个月中的最大功率）
        transformer_basic_cost = sum(monthly_max_loads) * demand_price
        
        # 计算年度总电费（含变压器基本电费和电量电费）
        total_cost = transformer_basic_cost + total_electricity_cost
        
        print(f'变压器基本电费: {transformer_basic_cost:.2f} 元')
        print(f'电量电费: {total_electricity_cost:.2f} 元')
        print(f'年度总电费（含变压器基本电费和电量电费）: {total_cost:.2f} 元')
    
    else:
        print(f'年度电费总成本: {total_electricity_cost:.2f} 元')

def simulate_storage_system(data, storage_capacity=None):
    """
    模拟储能系统运行，支持小时或15分钟粒度的数据
    
    参数:
    data: 负载和电价数据
    storage_capacity: 储能系统容量，如果为None则使用全局变量
    
    返回:
    modified_load: 修改后的负载曲线
    storage_level_history: 储能系统电量变化历史
    """
    # 使用全局变量
    global storage_capacity_per_system, max_power_per_system
    
    # 如果未指定容量，使用全局变量
    if storage_capacity is None:
        storage_capacity = storage_capacity_per_system
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    # 预处理数据，确保数据按时间排序
    if 'datetime' in data.columns:
        sorted_data = data.sort_values('datetime').copy()
    else:
        # 如果没有datetime列，尝试按date和hour排序
        if 'date' in data.columns and 'hour' in data.columns:
            sorted_data = data.sort_values(['date', 'hour']).copy()
        else:
            # 如果无法排序，直接使用原始数据
            sorted_data = data.copy()
    
    # 初始化储能系统
    storage_level = initial_storage_capacity
    storage_level_history = []
    modified_load = sorted_data['load'].copy()
    
    # 按日期进行分析
    unique_dates = sorted_data['datetime'].dt.date.unique() if 'datetime' in sorted_data.columns else []
    
    if len(unique_dates) > 0:
        # 按日期循环处理，适用于有datetime列的数据
        for date in unique_dates:
            day_data = sorted_data[sorted_data['datetime'].dt.date == date]
            
            # 识别价格分段
            prices = day_data['price'].values
            periods = day_data['period_type'].values if 'period_type' in day_data.columns else []
            
            # 找出最适合充放电的时段
            # 先按价格排序
            price_indices = np.argsort(prices)
            low_price_indices = price_indices[:len(price_indices)//3]  # 最低1/3价格的时段
            high_price_indices = price_indices[-len(price_indices)//3:]  # 最高1/3价格的时段
            
            low_price_indices = sorted(low_price_indices)  # 按时间顺序排列
            high_price_indices = sorted(high_price_indices)  # 按时间顺序排列
            
            # 处理日内所有时间点
            day_indices = day_data.index.tolist()
            
            # 充电策略：在低谷时段充电
            for idx in low_price_indices:
                if idx < len(day_indices):
                    data_idx = day_indices[idx]
                    # 考虑时间粒度，计算充电时间
                    charge_time = 0.25 if is_minute_data else 1.0  # 15分钟或1小时
                    
                    # 计算充电功率
                    charge_power = min(max_power_per_system, 
                                       (storage_capacity - storage_level) / (efficiency_bess * charge_time))
                    
                    if charge_power > 0:
                        storage_level += charge_power * efficiency_bess * charge_time
                        storage_level = min(storage_level, storage_capacity)
                        modified_load.loc[data_idx] += charge_power
            
            # 放电策略：在高峰时段放电
            for idx in high_price_indices:
                if idx < len(day_indices):
                    data_idx = day_indices[idx]
                    # 考虑时间粒度，计算放电时间
                    discharge_time = 0.25 if is_minute_data else 1.0  # 15分钟或1小时
                    
                    # 确保放电不超过当前负载，避免向电网反向供电
                    max_discharge = min(storage_level * efficiency_bess / discharge_time, 
                                        max_power_per_system,
                                        sorted_data.loc[data_idx, 'load'])
                    
                    if max_discharge > 0:
                        storage_level -= max_discharge * discharge_time / efficiency_bess
                        storage_level = max(storage_level, 0)
                        modified_load.loc[data_idx] -= max_discharge
            
            # 记录每个时点的储能电量
            for i in day_data.index:
                storage_level_history.append(storage_level)
    else:
        # 如果没有按日期分组，则直接按价格高低处理
        prices = sorted_data['price'].values
        price_threshold_high = np.percentile(prices, 70)  # 高于70%的价格认为是高价
        price_threshold_low = np.percentile(prices, 30)  # 低于30%的价格认为是低价
        
        # 模拟每个时间点
        for i, row in sorted_data.iterrows():
            current_price = row['price']
            current_load = row['load']
            
            # 考虑时间粒度
            time_factor = 0.25 if is_minute_data else 1.0  # 15分钟或1小时
            
            if current_price <= price_threshold_low:  # 低谷价格时充电
                charge_power = min(max_power_per_system, 
                                   (storage_capacity - storage_level) / (efficiency_bess * time_factor))
                
                if charge_power > 0:
                    storage_level += charge_power * efficiency_bess * time_factor
                    storage_level = min(storage_level, storage_capacity)
                    modified_load.loc[i] += charge_power
            
            elif current_price >= price_threshold_high:  # 高峰价格时放电
                # 确保放电不超过当前负载
                max_discharge = min(storage_level * efficiency_bess / time_factor, 
                                    max_power_per_system,
                                    current_load)
                
                if max_discharge > 0:
                    storage_level -= max_discharge * time_factor / efficiency_bess
                    storage_level = max(storage_level, 0)
                    modified_load.loc[i] -= max_discharge
            
            # 记录储能电量
            storage_level_history.append(storage_level)
    
    return modified_load, storage_level_history

def plot_storage_system(data, initial_date_index=0, auto_analyze=False):
    """绘制增加储能系统后的曲线（两充两放策略）并分析套利模式"""
    
    # 声明全局变量，必须放在函数开头
    global storage_capacity_per_system, max_power_per_system
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    if is_minute_data:
        # 对于15分钟数据，需要转换为小时级别的数据进行套利分析
        print("检测到15分钟粒度的数据，将按小时聚合进行套利分析...")
        # 按日期和小时聚合数据
        hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
            'load': 'mean',  # 使用平均值，处理不足4条/小时的情况
            'price': 'mean',
            'period_type': lambda x: x.mode()[0] if not x.mode().empty else None,
            'month': 'first',
            'datetime': 'first'  # 这里可能会导致列重复
        })
        
        # 重置索引前先删除datetime列，以避免重复
        if 'datetime' in hourly_data.columns:
            hourly_data = hourly_data.drop(columns=['datetime'])
            
        # 重置索引
        hourly_data = hourly_data.reset_index()
        
        # 检查并重命名索引列名
        if 'level_0' in hourly_data.columns:
            hourly_data = hourly_data.rename(columns={'level_0': 'date'})
        else:
            # 如果列名不是level_0，则假设第一列是日期
            date_column_name = hourly_data.columns[0]
            hourly_data = hourly_data.rename(columns={date_column_name: 'date'})
        
        # 检查是否有缺失的小时
        all_dates = hourly_data['date'].unique()
        complete_hourly_data = []
        
        for date in all_dates:
            date_data = hourly_data[hourly_data['date'] == date]
            complete_hourly_data.append(date_data)
        
        # 合并处理后的数据
        hourly_data = pd.concat(complete_hourly_data, ignore_index=True)
        
        # 确保必要的列都存在
        hourly_data['is_minute_data'] = False
        
        # 创建datetime列（如果需要）
        if 'datetime' not in hourly_data.columns:
            try:
                # 确保date列是datetime类型
                if not pd.api.types.is_datetime64_any_dtype(hourly_data['date']):
                    hourly_data['date'] = pd.to_datetime(hourly_data['date'])
                # 创建datetime列
                hourly_data['datetime'] = hourly_data['date'] + pd.to_timedelta(hourly_data['hour'], unit='h')
            except Exception as e:
                print(f"创建datetime列失败: {e}")
                # 这里不会终止，因为我们可以使用date列
        
        # 使用聚合后的小时级数据
        analysis_data = hourly_data
    else:
        # 对于小时数据，直接使用
        analysis_data = data
    
    # 确保有日期列可用于分析
    if 'datetime' not in analysis_data.columns and 'date' not in analysis_data.columns:
        print("错误: 数据中缺少日期信息，无法继续分析")
        return
    
    # 获取唯一日期列表
    try:
        # 尝试使用datetime列
        if 'datetime' in analysis_data.columns:
            unique_dates = sorted(analysis_data['datetime'].dt.date.unique())
        # 如果没有datetime列，尝试使用date列
        elif 'date' in analysis_data.columns:
            if pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                unique_dates = sorted(analysis_data['date'].dt.date.unique())
            else:
                unique_dates = sorted(analysis_data['date'].unique())
    except Exception as e:
        print(f"获取日期失败: {e}")
        return
        
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
    
    # 计算总天数
    total_days = len(unique_dates)
    analyzed_days = 0
        
    # 使用分析数据进行套利计算
    for power in power_range:
        # 计算对应的储能容量
        capacity = power * capacity_power_ratio
        
        # 设置新的容量和功率
        storage_capacity_per_system = capacity
        max_power_per_system = power
        
        # 进行所有日期的模拟计算
        monthly_costs = {}
        monthly_max_loads = {}  # 存储每月最大负载，用于计算按需收费
        
        # 按日期进行分析
        days_processed = 0
        for date in unique_dates:
            # 根据日期获取当天数据
            try:
                if 'datetime' in analysis_data.columns:
                    day_data = analysis_data[analysis_data['datetime'].dt.date == date].copy()
                elif pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                    day_data = analysis_data[analysis_data['date'].dt.date == date].copy()
                else:
                    day_data = analysis_data[analysis_data['date'] == date].copy()
            except Exception as e:
                continue
            
            # 接受任何非空的数据
            if day_data.empty:
                continue
                
            days_processed += 1
            
            # 获取月份
            month = day_data['month'].iloc[0]
            
            # 执行储能模拟
            storage_level = initial_storage_capacity
            original_load = day_data['load'].to_numpy()
            modified_load = original_load.copy()
            periods = day_data['period_type'].values if 'period_type' in day_data.columns else [FLAT] * len(original_load)
            
            # 执行套利分析和模拟
            # 第一次套利（简化版）
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None and charge1_start < len(original_load):
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
                
                if discharge1_start is not None and discharge1_start < len(original_load):
                    # 模拟充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess) if charge1_duration > 0 else 0)
                    for h in range(charge1_start, min(charge1_start + charge1_duration, len(original_load))):
                        actual_charge = min(power_charge, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                    
                    # 模拟放电
                    discharge_available = storage_level
                    power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration if discharge1_duration > 0 else 0)
                    for h in range(discharge1_start, min(discharge1_start + discharge1_duration, len(original_load))):
                        # 特殊处理原始负载为0的情况
                        if h >= len(original_load) or original_load[h] <= 0:
                            continue
                            
                        # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                        actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        if actual_discharge <= 0: break
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
            
            # 第二次套利（简化版）
            if discharge1_start is not None and discharge1_start < len(original_load):
                # 第二次套利代码保持不变...
                charge2_start, charge2_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, discharge1_start + discharge1_duration)
                if charge2_start is None:
                    charge2_start, charge2_duration, _ = find_continuous_window(periods, [FLAT], 2, discharge1_start + discharge1_duration)
                
                if charge2_start is not None and charge2_start < len(original_load):
                    discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:  # 尝试找1小时窗口
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 1, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 1, charge2_start + charge2_duration)
                    
                    if discharge2_start is not None and discharge2_start < len(original_load):
                        # 模拟充电
                        charge_needed = storage_capacity_per_system - storage_level
                        power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess) if charge2_duration > 0 else 0)
                        for h in range(charge2_start, min(charge2_start + charge2_duration, len(original_load))):
                            actual_charge = min(power_charge, max_power_per_system)
                            actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                            if actual_charge <= 0: break
                            modified_load[h] += actual_charge
                            storage_level += actual_charge * efficiency_bess
                            storage_level = min(storage_level, storage_capacity_per_system)
                        
                        # 模拟放电
                        discharge_available = storage_level
                        power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration if discharge2_duration > 0 else 0)
                        for h in range(discharge2_start, min(discharge2_start + discharge2_duration, len(original_load))):
                            # 特殊处理原始负载为0的情况
                            if h >= len(original_load) or original_load[h] <= 0:
                                continue
                                
                            # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                            actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                            actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                            if actual_discharge <= 0: break
                            modified_load[h] -= actual_discharge
                            storage_level -= actual_discharge / efficiency_bess
                            storage_level = max(storage_level, 0)
            
            # 计算当天各时段电费
            day_costs = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
            for h in range(len(modified_load)):
                if h >= len(day_data):
                    continue
                period = day_data['period_type'].iloc[h]
                price = day_data['price'].iloc[h]
                load = modified_load[h]
                day_costs[period] += load * price
            
            # 记录该天的最大负载，用于按需收费计算
            if len(modified_load) > 0:
                max_load_of_day = np.max(modified_load)
                
                # 累加到月度统计
                if month not in monthly_costs:
                    monthly_costs[month] = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
                    monthly_max_loads[month] = 0
                
                for period, cost in day_costs.items():
                    monthly_costs[month][period] += cost
                
                # 更新月最大负载
                monthly_max_loads[month] = max(monthly_max_loads[month], max_load_of_day)
        
        # 只在第一次循环更新分析天数
        if power == power_range[0]:
            analyzed_days = days_processed
            if auto_analyze:  # 只有在批量分析时才打印
                print(f"共分析了 {analyzed_days}/{total_days} 天的数据")
        
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
    
    # 添加菜单返回提示
    print("\n分析完成！按回车键返回主菜单...", end="")
    input()
    print()

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
        
        # 获取该月的24小时数据
        day_data = month_data[month_data['datetime'].dt.day == 1]
        
        # 检查数据是否完整
        if day_data.empty or len(day_data) < 24:
            print(f"月份 {month} 的数据不完整，跳过绘图。")
            continue
        
        # 确保有24小时数据
        monthly_data[month] = {
            'hours': range(24),
            'prices': day_data['price'].values[:24],
            'period_types': day_data['period_type'].values[:24]
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
        
        bars = ax.bar(hours, prices)
        
        for hour, bar, period_type in zip(hours, bars, period_types):
            bar.set_color(period_colors.get(period_type, 'gray'))
        
        # 添加图例
        unique_period_types = set(period_types)
        legend_elements = [plt.Rectangle((0,0),1,1, color=period_colors[pt], label=period_labels[pt]) 
                          for pt in sorted(unique_period_types) if pt in period_colors]
        
        ax.legend(handles=legend_elements)
        ax.set_title(f'电价曲线，月份: {month}')
        ax.set_xlabel('小时')
        ax.set_ylabel('电价 (元/kWh)')
        ax.set_xticks(hours)
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

def find_optimal_storage_capacity(data):
    """寻找最佳储能系统容量"""
    # 声明全局变量，必须放在函数开头
    global storage_capacity_per_system, max_power_per_system, annual_electricity_cost_without_storage
    
    # 首先计算不使用储能系统时的总电费
    print("首先计算不安装储能系统的年度总电费...\n")
    # 计算原始电费
    if 'annual_electricity_cost_without_storage' not in globals() or annual_electricity_cost_without_storage == 0:
        # 计算原始电费
        monthly_costs_original = {}
        monthly_max_loads_original = {}
        
        # 检查是否为分钟粒度的数据
        is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
        
        if is_minute_data:
            # 对于15分钟数据，需要转换为小时级别的数据进行套利分析
            print("检测到15分钟粒度的数据，将按小时聚合进行套利分析...")
            # 按日期和小时聚合数据
            hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
                'load': 'mean',  # 使用平均值，处理不足4条/小时的情况
                'price': 'mean',
                'period_type': lambda x: x.mode()[0] if not x.mode().empty else None,
                'month': 'first',
                'datetime': 'first'  # 这里可能会导致列重复
            })
            
            # 重置索引前先删除datetime列，以避免重复
            if 'datetime' in hourly_data.columns:
                hourly_data = hourly_data.drop(columns=['datetime'])
                
            # 重置索引
            hourly_data = hourly_data.reset_index()
            
            # 检查并重命名索引列名
            if 'level_0' in hourly_data.columns:
                hourly_data = hourly_data.rename(columns={'level_0': 'date'})
            else:
                # 如果列名不是level_0，则假设第一列是日期
                date_column_name = hourly_data.columns[0]
                hourly_data = hourly_data.rename(columns={date_column_name: 'date'})
            
            # 不再打印缺少小时的警告，只添加可用数据
            complete_hourly_data = []
            
            for date in hourly_data['date'].unique():
                date_data = hourly_data[hourly_data['date'] == date]
                complete_hourly_data.append(date_data)
            
            # 合并处理后的数据
            hourly_data = pd.concat(complete_hourly_data, ignore_index=True)
            
            # 确保必要的列都存在
            hourly_data['is_minute_data'] = False
            
            # 创建datetime列（如果需要）
            if 'datetime' not in hourly_data.columns:
                try:
                    # 确保date列是datetime类型
                    if not pd.api.types.is_datetime64_any_dtype(hourly_data['date']):
                        hourly_data['date'] = pd.to_datetime(hourly_data['date'])
                    # 创建datetime列
                    hourly_data['datetime'] = hourly_data['date'] + pd.to_timedelta(hourly_data['hour'], unit='h')
                except Exception as e:
                    print(f"创建datetime列失败: {e}")
                    # 这里不会终止，因为我们可以使用date列
                
                # 使用聚合后的小时级数据
                analysis_data = hourly_data
            else:
                # 对于小时数据，直接使用
                analysis_data = data
            
            # 确保有日期列可用于分析
            if 'datetime' not in analysis_data.columns and 'date' not in analysis_data.columns:
                print("错误: 数据中缺少日期信息，无法继续分析")
                return
            
            # 获取唯一日期列表
            try:
                # 尝试使用datetime列
                if 'datetime' in analysis_data.columns:
                    unique_dates = sorted(analysis_data['datetime'].dt.date.unique())
                # 如果没有datetime列，尝试使用date列
                elif 'date' in analysis_data.columns:
                    if pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                        unique_dates = sorted(analysis_data['date'].dt.date.unique())
                    else:
                        unique_dates = sorted(analysis_data['date'].unique())
            except Exception as e:
                print(f"获取日期失败: {e}")
                return
            
            # 计算变压器基本电费前需要用户输入
            if method_basic_capacity_cost_transformer == 1:  # 按容量收取
                capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
            elif method_basic_capacity_cost_transformer == 2:  # 按需收取
                demand_price = float(input('请输入需量单价（元/kW·月）: '))
            
            # 按日期计算原始电费
            for date in unique_dates:
                try:
                    if 'datetime' in analysis_data.columns:
                        day_data = analysis_data[analysis_data['datetime'].dt.date == date].copy()
                    elif pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                        day_data = analysis_data[analysis_data['date'].dt.date == date].copy()
                    else:
                        day_data = analysis_data[analysis_data['date'] == date].copy()
                except Exception as e:
                    continue
                
                # 接受任何非空的数据
                if day_data.empty:
                    continue
                
                # 获取月份
                month = day_data['month'].iloc[0]
                
                # 计算当天各时段电费
                day_costs = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
                for i, row in day_data.iterrows():
                    period = row['period_type']
                    price = row['price']
                    load = row['load']
                    # 计算电费时考虑时间粒度
                    time_factor = 0.25 if is_minute_data else 1.0
                    day_costs[period] += load * price * time_factor
                
                # 记录最大负载（用于需量计费）
                max_load = day_data['load'].max()
                
                # 累加到月度统计
                if month not in monthly_costs_original:
                    monthly_costs_original[month] = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
                    monthly_max_loads_original[month] = 0
                
                for period, cost in day_costs.items():
                    monthly_costs_original[month][period] += cost
                
                # 更新月最大负载
                monthly_max_loads_original[month] = max(monthly_max_loads_original[month], max_load)
            
            # 计算原始总电费
            electricity_cost_original = sum(sum(month_data.values()) for month_data in monthly_costs_original.values())
            
            # 计算变压器基本电费
            transformer_basic_cost_original = 0
            if method_basic_capacity_cost_transformer == 1:  # 按容量收取
                transformer_basic_cost_original = capacity_price * transformer_capacity * 12  # 12个月
            elif method_basic_capacity_cost_transformer == 2:  # 按需收取
                transformer_basic_cost_original = sum(monthly_max_loads_original.values()) * demand_price
            
            # 计算年度总电费（无储能系统）
            annual_electricity_cost_without_storage = electricity_cost_original + transformer_basic_cost_original
            
            print(f"不安装储能系统的年度总电费: {annual_electricity_cost_without_storage:.2f} 元")
        else:
            print(f"使用已计算的不安装储能系统的年度总电费: {annual_electricity_cost_without_storage:.2f} 元")
        
        print("\n开始寻找最佳储能系统容量...")
        
        # 检查是否为分钟粒度的数据
        is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
        
        if is_minute_data:
            # 对于15分钟数据，需要转换为小时级别的数据进行套利分析
            print("检测到15分钟粒度的数据，将按小时聚合进行套利分析...")
            # 按日期和小时聚合数据
            hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
                'load': 'mean',  # 使用平均值，处理不足4条/小时的情况
                'price': 'mean',
                'period_type': lambda x: x.mode()[0] if not x.mode().empty else None,
                'month': 'first',
                'datetime': 'first'  # 这里可能会导致列重复
            })
            
            # 重置索引前先删除datetime列，以避免重复
            if 'datetime' in hourly_data.columns:
                hourly_data = hourly_data.drop(columns=['datetime'])
                
            # 重置索引
            hourly_data = hourly_data.reset_index()
            
            # 检查并重命名索引列名
            if 'level_0' in hourly_data.columns:
                hourly_data = hourly_data.rename(columns={'level_0': 'date'})
            else:
                # 如果列名不是level_0，则假设第一列是日期
                date_column_name = hourly_data.columns[0]
                hourly_data = hourly_data.rename(columns={date_column_name: 'date'})
            
            # 检查是否有缺失的小时
            all_dates = hourly_data['date'].unique()
            complete_hourly_data = []
            
            for date in all_dates:
                date_data = hourly_data[hourly_data['date'] == date]
                # 不再打印缺少小时的警告，只添加可用数据
                complete_hourly_data.append(date_data)
            
            # 合并处理后的数据
            hourly_data = pd.concat(complete_hourly_data, ignore_index=True)
            
            # 确保必要的列都存在
            hourly_data['is_minute_data'] = False
            
            # 创建datetime列（如果需要）
            if 'datetime' not in hourly_data.columns:
                try:
                    # 确保date列是datetime类型
                    if not pd.api.types.is_datetime64_any_dtype(hourly_data['date']):
                        hourly_data['date'] = pd.to_datetime(hourly_data['date'])
                    # 创建datetime列
                    hourly_data['datetime'] = hourly_data['date'] + pd.to_timedelta(hourly_data['hour'], unit='h')
                except Exception as e:
                    print(f"创建datetime列失败: {e}")
                    # 这里不会终止，因为我们可以使用date列
                
                # 使用聚合后的小时级数据
                analysis_data = hourly_data
            else:
                # 对于小时数据，直接使用
                analysis_data = data
    
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
    
    # 计算总天数
    total_days = len(unique_dates)
    total_processed_days = 0
        
    # 使用分析数据进行套利计算
    for power in power_range:
        # 计算对应的储能容量
        capacity = power * capacity_power_ratio
        
        # 设置新的容量和功率
        storage_capacity_per_system = capacity
        max_power_per_system = power
        
        # 进行所有日期的模拟计算
        monthly_costs = {}
        monthly_max_loads = {}  # 存储每月最大负载，用于计算按需收费
        
        # 按日期进行分析
        days_processed = 0
        for date in unique_dates:
            # 根据日期获取当天数据
            try:
                if 'datetime' in analysis_data.columns:
                    day_data = analysis_data[analysis_data['datetime'].dt.date == date].copy()
                elif pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                    day_data = analysis_data[analysis_data['date'].dt.date == date].copy()
                else:
                    day_data = analysis_data[analysis_data['date'] == date].copy()
            except Exception as e:
                continue
            
            # 接受任何非空的数据，不再要求必须是24小时
            if day_data.empty:
                continue
                
            days_processed += 1
            
            # 获取月份
            month = day_data['month'].iloc[0]
            
            # 执行储能模拟
            storage_level = initial_storage_capacity
            original_load = day_data['load'].to_numpy()
            modified_load = original_load.copy()
            periods = day_data['period_type'].values if 'period_type' in day_data.columns else [FLAT] * len(original_load)
            
            # 执行套利分析和模拟，但不打印详细日志
            # 第一次套利（简化版）
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None and charge1_start < len(original_load):
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
                
                if discharge1_start is not None and discharge1_start < len(original_load):
                    # 模拟充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess) if charge1_duration > 0 else 0)
                    for h in range(charge1_start, min(charge1_start + charge1_duration, len(original_load))):
                        actual_charge = min(power_charge, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                    
                    # 模拟放电
                    discharge_available = storage_level
                    power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration if discharge1_duration > 0 else 0)
                    for h in range(discharge1_start, min(discharge1_start + discharge1_duration, len(original_load))):
                        # 特殊处理原始负载为0的情况
                        if h >= len(original_load) or original_load[h] <= 0:
                            continue
                            
                        # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                        actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        if actual_discharge <= 0: break
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
            
            # 第二次套利（简化版）
            if discharge1_start is not None and discharge1_start < len(original_load):
                # 第二次套利代码保持不变...
                charge2_start, charge2_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, discharge1_start + discharge1_duration)
                if charge2_start is None:
                    charge2_start, charge2_duration, _ = find_continuous_window(periods, [FLAT], 2, discharge1_start + discharge1_duration)
                
                if charge2_start is not None and charge2_start < len(original_load):
                    discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:  # 尝试找1小时窗口
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 1, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 1, charge2_start + charge2_duration)
                    
                    if discharge2_start is not None and discharge2_start < len(original_load):
                        # 模拟充电
                        charge_needed = storage_capacity_per_system - storage_level
                        power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess) if charge2_duration > 0 else 0)
                        for h in range(charge2_start, min(charge2_start + charge2_duration, len(original_load))):
                            actual_charge = min(power_charge, max_power_per_system)
                            actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                            if actual_charge <= 0: break
                            modified_load[h] += actual_charge
                            storage_level += actual_charge * efficiency_bess
                            storage_level = min(storage_level, storage_capacity_per_system)
                        
                        # 模拟放电
                        discharge_available = storage_level
                        power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration if discharge2_duration > 0 else 0)
                        for h in range(discharge2_start, min(discharge2_start + discharge2_duration, len(original_load))):
                            # 特殊处理原始负载为0的情况
                            if h >= len(original_load) or original_load[h] <= 0:
                                continue
                                
                            # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                            actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                            actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                            if actual_discharge <= 0: break
                            modified_load[h] -= actual_discharge
                            storage_level -= actual_discharge / efficiency_bess
                            storage_level = max(storage_level, 0)
            
            # 计算当天各时段电费
            day_costs = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
            for h in range(len(modified_load)):
                if h >= len(day_data):
                    continue
                period = day_data['period_type'].iloc[h]
                price = day_data['price'].iloc[h]
                load = modified_load[h]
                day_costs[period] += load * price
            
            # 记录该天的最大负载，用于按需收费计算
            if len(modified_load) > 0:
                max_load_of_day = np.max(modified_load)
                
                # 累加到月度统计
                if month not in monthly_costs:
                    monthly_costs[month] = {SHARP: 0, PEAK: 0, FLAT: 0, VALLEY: 0, DEEP_VALLEY: 0}
                    monthly_max_loads[month] = 0
                
                for period, cost in day_costs.items():
                    monthly_costs[month][period] += cost
                
                # 更新月最大负载
                monthly_max_loads[month] = max(monthly_max_loads[month], max_load_of_day)
        
        # 更新总分析天数计数，仅第一次循环时更新
        if power == power_range[0]:
            total_processed_days = days_processed
            print(f"分析完成：共处理了 {total_processed_days} 天的数据")
        
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
    
    # 添加菜单返回提示
    print("\n分析完成！按回车键返回主菜单...", end="")
    input()
    print()
    
    # 在结尾添加电费节省分析
    print("\n" + "=" * 100)
    print(f'不安装储能系统的年度总电费: {annual_electricity_cost_without_storage:.2f} 元')
    print(f'最佳储能系统功率: {optimal_power} kW')
    print(f'最佳储能系统容量: {optimal_capacity:.2f} kWh')
    print(f'安装储能系统后的年度电费: {cost_breakdown[optimal_power]["annual_electricity_cost"]:.2f} 元')
    print(f'储能系统总造价: {cost_breakdown[optimal_power]["system_total_cost"]:.2f} 元')
    print(f'储能系统年摊销成本: {cost_breakdown[optimal_power]["annual_system_cost"]:.2f} 元')
    print(f'年度总成本(电费+系统摊销): {optimal_cost:.2f} 元')
    
    # 计算节省金额和百分比
    savings = annual_electricity_cost_without_storage - cost_breakdown[optimal_power]["annual_electricity_cost"]
    savings_percent = savings / annual_electricity_cost_without_storage * 100
    
    print(f'年度电费节省金额: {savings:.2f} 元')
    print(f'年度电费节省比例: {savings_percent:.2f}%')
    print("=" * 100)
    
    # 绘制容量-成本曲线与节省比例
    plt.figure(figsize=(14, 10))
    
    # 电费和成本曲线
    plt.subplot(2, 1, 1)
    plt.plot(power_range, [cost_breakdown[p]['annual_electricity_cost'] for p in power_range], 'g-o', label='年度电费')
    plt.plot(power_range, [cost_breakdown[p]['annual_system_cost'] for p in power_range], 'r-o', label='储能系统年摊销成本')
    plt.plot(power_range, costs, 'b-o', label='年度总成本')
    plt.axhline(y=annual_electricity_cost_without_storage, color='black', linestyle='--', label='不安装储能系统的年度电费')
    
    plt.title('储能系统功率-成本关系曲线')
    plt.xlabel('储能系统功率 (kW)')
    plt.ylabel('成本 (元)')
    plt.grid(True)
    plt.legend()
    
    # 标记最优点
    plt.plot(optimal_power, optimal_cost, 'mo', markersize=10)
    plt.annotate(f'最优功率: {optimal_power} kW\n节省电费: {savings:.2f} 元 ({savings_percent:.2f}%)', 
                 xy=(optimal_power, optimal_cost), 
                 xytext=(optimal_power + 200, optimal_cost - 50000 if optimal_power < 1600 else optimal_cost + 50000),
                 arrowprops=dict(facecolor='black', shrink=0.05, width=1.5),
                 bbox=dict(boxstyle="round,pad=0.5", fc="yellow", alpha=0.8))
    
    # 电费节省百分比曲线
    plt.subplot(2, 1, 2)
    savings_percents = [(annual_electricity_cost_without_storage - cost_breakdown[p]["annual_electricity_cost"]) / annual_electricity_cost_without_storage * 100 for p in power_range]
    plt.plot(power_range, savings_percents, 'g-o')
    plt.title('储能系统功率-电费节省百分比关系曲线')
    plt.xlabel('储能系统功率 (kW)')
    plt.ylabel('电费节省比例 (%)')
    plt.grid(True)
    
    # 标记最优点
    plt.plot(optimal_power, savings_percent, 'mo', markersize=10)
    plt.annotate(f'最优功率: {optimal_power} kW\n节省比例: {savings_percent:.2f}%', 
                 xy=(optimal_power, savings_percent), 
                 xytext=(optimal_power + 200, savings_percent + 2),
                 arrowprops=dict(facecolor='black', shrink=0.05, width=1.5),
                 bbox=dict(boxstyle="round,pad=0.5", fc="yellow", alpha=0.8))
    
    plt.tight_layout()
    plt.show()
    
    # 添加饼图表示成本构成
    plt.figure(figsize=(12, 6))
    
    # 获取最优功率下的各项成本
    opt_data = cost_breakdown[optimal_power]
    
    # 准备饼图数据
    labels = ['电量电费', '变压器基本电费', '储能系统摊销成本']
    sizes = [opt_data['electricity_cost'], opt_data['transformer_basic_cost'], opt_data['annual_system_cost']]
    colors = ['lightgreen', 'lightblue', 'coral']
    explode = (0.1, 0.1, 0.1)  # 突出显示所有部分
    
    # 绘制饼图
    plt.subplot(1, 2, 1)
    plt.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', shadow=True, startangle=140)
    plt.axis('equal')  # 确保饼图是圆的
    plt.title(f'最优容量({optimal_capacity:.2f} kWh)下的年度成本构成')
    
    # 添加电费节省对比图
    plt.subplot(1, 2, 2)
    labels = ['不安装储能系统', '安装最优储能系统']
    sizes = [annual_electricity_cost_without_storage, opt_data['annual_electricity_cost']]
    colors = ['red', 'green']
    
    plt.bar(labels, sizes, color=colors)
    plt.title('安装储能系统前后年度电费对比')
    plt.ylabel('年度电费 (元)')
    plt.grid(axis='y')
    
    # 添加节省金额和百分比文本
    plt.text(0.5, 0.5, f'节省: {savings:.2f} 元\n({savings_percent:.2f}%)',
            horizontalalignment='center',
            verticalalignment='center',
            transform=plt.gca().transAxes,
            bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.5))
    
    plt.tight_layout()
    plt.show()

def compare_storage_capacities(data):
    """
    比较最佳容量点和其他输入容量的性能指标
    
    包括以下分析:
    - 年度总成本对比
    - 充放电策略评估
    - 与历史负载匹配度分析
    - 关键性能指标计算:
      - 平均日充放电量
      - 等效满充次数
      - 容量利用率
      - 系统利用率
      - 负载匹配度
      - 容量浪费/不足风险评估
    """
    # 声明全局变量，必须放在函数开头
    global storage_capacity_per_system, max_power_per_system, annual_electricity_cost_without_storage
    
    # 首先确保已计算原始电费
    if 'annual_electricity_cost_without_storage' not in globals() or annual_electricity_cost_without_storage == 0:
        print("需要先计算不安装储能系统的年度总电费，请先运行选项5...")
        return
    else:
        print(f"不安装储能系统的年度总电费: {annual_electricity_cost_without_storage:.2f} 元")
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    if is_minute_data:
        # 对于15分钟数据，按小时聚合
        print("检测到15分钟粒度的数据，将按小时聚合进行分析...")
        # 按日期和小时聚合数据
        hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
            'load': 'mean',  # 使用平均值
            'price': 'mean',
            'period_type': lambda x: x.mode()[0] if not x.mode().empty else None,
            'month': 'first',
            'datetime': 'first'
        })
        
        # 重置索引前先删除datetime列，以避免重复
        if 'datetime' in hourly_data.columns:
            hourly_data = hourly_data.drop(columns=['datetime'])
            
        # 重置索引
        hourly_data = hourly_data.reset_index()
        
        # 检查并重命名索引列名
        if 'level_0' in hourly_data.columns:
            hourly_data = hourly_data.rename(columns={'level_0': 'date'})
        else:
            # 如果列名不是level_0，则假设第一列是日期
            date_column_name = hourly_data.columns[0]
            hourly_data = hourly_data.rename(columns={date_column_name: 'date'})
        
        # 确保必要的列都存在
        hourly_data['is_minute_data'] = False
        
        # 创建datetime列（如果需要）
        if 'datetime' not in hourly_data.columns:
            try:
                if not pd.api.types.is_datetime64_any_dtype(hourly_data['date']):
                    hourly_data['date'] = pd.to_datetime(hourly_data['date'])
                hourly_data['datetime'] = hourly_data['date'] + pd.to_timedelta(hourly_data['hour'], unit='h')
            except Exception as e:
                print(f"创建datetime列失败: {e}")
        
        # 使用聚合后的小时级数据
        analysis_data = hourly_data
    else:
        # 对于小时数据，直接使用
        analysis_data = data
    
    # 确保有日期列可用于分析
    if 'datetime' not in analysis_data.columns and 'date' not in analysis_data.columns:
        print("错误: 数据中缺少日期信息，无法继续分析")
        return
    
    # 获取唯一日期列表
    try:
        if 'datetime' in analysis_data.columns:
            unique_dates = sorted(analysis_data['datetime'].dt.date.unique())
        elif 'date' in analysis_data.columns:
            if pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                unique_dates = sorted(analysis_data['date'].dt.date.unique())
            else:
                unique_dates = sorted(analysis_data['date'].unique())
    except Exception as e:
        print(f"获取日期失败: {e}")
        return
    
    # 统计所有日期的总数
    total_dates = len(unique_dates)
    print(f"数据集包含 {total_dates} 天的数据")
    
    # 储能系统容量与功率的比例
    capacity_power_ratio = storage_capacity_per_system / max_power_per_system
    
    # 保存用户输入的价格，避免重复输入
    system_unit_cost = float(input('请输入储能系统单位造价（元/Wh）: '))
    system_lifetime = float(input('请输入储能系统使用寿命（年）: '))
    
    capacity_price = 0
    demand_price = 0
    if method_basic_capacity_cost_transformer == 1:  # 按容量收取
        capacity_price = float(input('请输入容量单价（元/kVA·月）: '))
    elif method_basic_capacity_cost_transformer == 2:  # 按需收取
        demand_price = float(input('请输入需量单价（元/kW·月）: '))
    
    # 保存原始参数值
    original_capacity = storage_capacity_per_system
    original_power = max_power_per_system
    
    # 首先，找出最佳容量点
    # 测试范围：从100kW到10000kW的功率
    power_range = list(range(100, 10001, 100))  # 以100kW为步长
    costs = []
    
    print("正在计算最佳容量点...")
    for power in power_range:
        # 计算对应的储能容量
        capacity = power * capacity_power_ratio
        
        # 设置新的容量和功率
        storage_capacity_per_system = capacity
        max_power_per_system = power
        
        # 执行简化版的成本计算
        monthly_max_loads = {}  # 存储每月最大负载
        monthly_costs = {}  # 存储每月电费
        
        # 分析所有日期，不再跳过任何日期数据
        date_count = 0
        for date in unique_dates[:min(30, len(unique_dates))]:  # 仅使用前30天加速计算
            try:
                if 'datetime' in analysis_data.columns:
                    day_data = analysis_data[analysis_data['datetime'].dt.date == date].copy()
                elif pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                    day_data = analysis_data[analysis_data['date'].dt.date == date].copy()
                else:
                    day_data = analysis_data[analysis_data['date'] == date].copy()
            except Exception as e:
                continue
            
            # 接受任何非空的数据，不再要求24小时
            if day_data.empty:
                continue
                
            date_count += 1
            month = day_data['month'].iloc[0]
            
            # 模拟充放电
            storage_level = initial_storage_capacity
            original_load = day_data['load'].to_numpy()
            modified_load = original_load.copy()
            periods = day_data['period_type'].values if 'period_type' in day_data.columns else [FLAT] * len(original_load)
            
            # 简化版套利分析
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None and charge1_start < len(original_load):
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
                
                if discharge1_start is not None and discharge1_start < len(original_load):
                    # 模拟充电和放电（保留实际实现）
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess) if charge1_duration > 0 else 0)
                    for h in range(charge1_start, min(charge1_start + charge1_duration, len(original_load))):
                        actual_charge = min(power_charge, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                    
                    # 模拟放电
                    discharge_available = storage_level
                    power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration if discharge1_duration > 0 else 0)
                    for h in range(discharge1_start, min(discharge1_start + discharge1_duration, len(original_load))):
                        if h >= len(original_load) or original_load[h] <= 0:
                            continue
                        actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        if actual_discharge <= 0: break
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
            
            # 计算当天最大负载
            if len(modified_load) > 0:
                max_load_of_day = np.max(modified_load)
                
                # 累加到月度统计
                if month not in monthly_costs:
                    monthly_costs[month] = 0
                    monthly_max_loads[month] = 0
                
                monthly_max_loads[month] = max(monthly_max_loads[month], max_load_of_day)
        
        if date_count == 0:
            print("警告：没有找到有效的日期数据进行分析")
            return
            
        # 计算变压器基本电费
        transformer_basic_cost = 0
        if method_basic_capacity_cost_transformer == 1:  # 按容量收取
            transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
        elif method_basic_capacity_cost_transformer == 2:  # 按需收取
            transformer_basic_cost = sum(monthly_max_loads.values()) * demand_price
        
        # 计算系统总造价和年度摊销
        system_total_cost = system_unit_cost * capacity * 1000  # 转换为Wh
        annual_system_cost = system_total_cost / system_lifetime
        
        # 年度总成本（粗略估计）
        annual_total_cost = transformer_basic_cost + annual_system_cost
        costs.append(annual_total_cost)
    
    # 找出最优容量
    optimal_power_idx = np.argmin(costs)
    optimal_power = power_range[optimal_power_idx]
    optimal_capacity = optimal_power * capacity_power_ratio
    
    print(f"\n初步估计的最佳储能系统功率: {optimal_power} kW")
    print(f"对应容量: {optimal_capacity:.2f} kWh\n")
    
    # 用户输入其他要比较的容量
    other_capacity = 0
    try:
        other_input = input("请输入要比较的其他储能容量(kWh)，或直接回车使用默认值(原始容量): ")
        if other_input.strip():
            other_capacity = float(other_input)
    except ValueError:
        print("输入无效，将使用默认容量")
        
    if other_capacity <= 0:
        other_capacity = original_capacity
        print(f"使用默认容量: {other_capacity} kWh")
    
    # 计算对应的功率
    other_power = other_capacity / capacity_power_ratio
    
    # 准备比较的容量列表
    capacities_to_compare = [
        {"power": optimal_power, "capacity": optimal_capacity, "name": "最佳容量点"},
        {"power": other_power, "capacity": other_capacity, "name": "指定容量点"}
    ]
    
    # 存储比较结果的字典
    comparison_results = {}
    
    print("\n正在进行容量对比分析...")
    print("=" * 100)
    
    # 对每个容量点执行详细分析
    for cap_info in capacities_to_compare:
        power = cap_info["power"]
        capacity = cap_info["capacity"]
        name = cap_info["name"]
        
        print(f"\n分析 {name}（功率: {power:.2f} kW, 容量: {capacity:.2f} kWh）...")
        
        # 设置容量和功率
        storage_capacity_per_system = capacity
        max_power_per_system = power
        
        # 初始化统计数据
        total_charge = 0  # 总充电量
        total_discharge = 0  # 总放电量
        days_with_full_charge = 0  # 达到满充的天数
        days_with_analysis = 0  # 有效分析天数
        daily_charges = []  # 每日充电量
        daily_discharges = []  # 每日放电量
        max_loads = []  # 每日最大负载
        original_max_loads = []  # 原始最大负载
        
        # 对所有日期进行详细分析，不再跳过任何日期
        for date in unique_dates:
            try:
                if 'datetime' in analysis_data.columns:
                    day_data = analysis_data[analysis_data['datetime'].dt.date == date].copy()
                elif pd.api.types.is_datetime64_any_dtype(analysis_data['date']):
                    day_data = analysis_data[analysis_data['date'].dt.date == date].copy()
                else:
                    day_data = analysis_data[analysis_data['date'] == date].copy()
            except Exception as e:
                continue
            
            # 接受任何数量的小时数据，只要不是空的
            if day_data.empty:
                continue
                
            days_with_analysis += 1
            
            # 执行储能模拟
            storage_level = initial_storage_capacity
            original_load = day_data['load'].to_numpy()
            modified_load = original_load.copy()
            periods = day_data['period_type'].values if 'period_type' in day_data.columns else [FLAT] * len(original_load)
            
            # 记录该日的充放电量
            day_charge = 0
            day_discharge = 0
            max_storage_level = storage_level
            
            # 执行套利分析和模拟
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None and charge1_start < len(original_load):
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
                
                if discharge1_start is not None and discharge1_start < len(original_load):
                    # 模拟充电
                    charge_needed = storage_capacity_per_system - storage_level
                    power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess) if charge1_duration > 0 else 0)
                    for h in range(charge1_start, min(charge1_start + charge1_duration, len(original_load))):
                        actual_charge = min(power_charge, max_power_per_system)
                        actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                        if actual_charge <= 0: break
                        modified_load[h] += actual_charge
                        storage_level += actual_charge * efficiency_bess
                        storage_level = min(storage_level, storage_capacity_per_system)
                        day_charge += actual_charge
                        max_storage_level = max(max_storage_level, storage_level)
                    
                    # 模拟放电
                    discharge_available = storage_level
                    power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration if discharge1_duration > 0 else 0)
                    for h in range(discharge1_start, min(discharge1_start + discharge1_duration, len(original_load))):
                        if original_load[h] <= 0:
                            continue
                        actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                        actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                        if actual_discharge <= 0: break
                        modified_load[h] -= actual_discharge
                        storage_level -= actual_discharge / efficiency_bess
                        storage_level = max(storage_level, 0)
                        day_discharge += actual_discharge
            
            # 第二次套利尝试
            if discharge1_start is not None and discharge1_start < len(original_load):
                charge2_start, charge2_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, discharge1_start + discharge1_duration)
                if charge2_start is None:
                    charge2_start, charge2_duration, _ = find_continuous_window(periods, [FLAT], 2, discharge1_start + discharge1_duration)
                
                if charge2_start is not None and charge2_start < len(original_load):
                    discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 2, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [SHARP], 1, charge2_start + charge2_duration)
                    if discharge2_start is None:
                        discharge2_start, discharge2_duration, _ = find_continuous_window(periods, [PEAK], 1, charge2_start + charge2_duration)
                    
                    if discharge2_start is not None and discharge2_start < len(original_load):
                        # 模拟充电
                        charge_needed = storage_capacity_per_system - storage_level
                        power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess) if charge2_duration > 0 else 0)
                        for h in range(charge2_start, min(charge2_start + charge2_duration, len(original_load))):
                            actual_charge = min(power_charge, max_power_per_system)
                            actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                            if actual_charge <= 0: break
                            modified_load[h] += actual_charge
                            storage_level += actual_charge * efficiency_bess
                            storage_level = min(storage_level, storage_capacity_per_system)
                            day_charge += actual_charge
                            max_storage_level = max(max_storage_level, storage_level)
                        
                        # 模拟放电
                        discharge_available = storage_level
                        power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration if discharge2_duration > 0 else 0)
                        for h in range(discharge2_start, min(discharge2_start + discharge2_duration, len(original_load))):
                            if original_load[h] <= 0:
                                continue
                            actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                            actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                            if actual_discharge <= 0: break
                            modified_load[h] -= actual_discharge
                            storage_level -= actual_discharge / efficiency_bess
                            storage_level = max(storage_level, 0)
                            day_discharge += actual_discharge
            
            # 统计信息
            total_charge += day_charge
            total_discharge += day_discharge
            daily_charges.append(day_charge)
            daily_discharges.append(day_discharge)
            
            # 检查是否达到满充
            if max_storage_level >= 0.95 * storage_capacity_per_system:
                days_with_full_charge += 1
            
            # 记录负载峰值
            if len(original_load) > 0:
                original_max_loads.append(np.max(original_load))
                max_loads.append(np.max(modified_load))
        
        # 计算关键指标
        avg_daily_charge = total_charge / days_with_analysis if days_with_analysis > 0 else 0
        avg_daily_discharge = total_discharge / days_with_analysis if days_with_analysis > 0 else 0
        equivalent_full_cycles = total_discharge / capacity if capacity > 0 else 0
        capacity_utilization = days_with_full_charge / days_with_analysis * 100 if days_with_analysis > 0 else 0
        system_utilization = (total_discharge / days_with_analysis) / (power * 24) * 100 if days_with_analysis > 0 else 0
        
        # 负载匹配度计算
        load_match_ability = 0
        if len(original_max_loads) > 0:
            peak_reduction = np.mean([(o - m) / o * 100 if o > 0 else 0 for o, m in zip(original_max_loads, max_loads)])
            load_match_ability = peak_reduction
        
        # 容量评估
        capacity_assessment = ""
        if capacity_utilization < 60:
            capacity_assessment = "可能容量过大，存在浪费风险"
        elif system_utilization > 90:
            capacity_assessment = "容量可能不足，无法完全满足峰值需求"
        else:
            capacity_assessment = "容量合适"
        
        # 保存结果
        comparison_results[name] = {
            "power": power,
            "capacity": capacity,
            "avg_daily_charge": avg_daily_charge,
            "avg_daily_discharge": avg_daily_discharge,
            "equivalent_full_cycles": equivalent_full_cycles,
            "capacity_utilization": capacity_utilization,
            "system_utilization": system_utilization,
            "load_match_ability": load_match_ability,
            "capacity_assessment": capacity_assessment
        }
    
    # 比较结果
    print("\n" + "=" * 100)
    print("容量点对比结果:")
    print("=" * 100)
    
    # 创建比较表格
    comparison_table = []
    
    # 添加不安装储能系统的信息作为对比基准
    comparison_table.append([
        "不安装储能系统",
        "0 kW / 0 kWh",
        "0 kWh",
        "0 kWh",
        "0",
        "0%",
        "0%",
        "0%",
        "基准值"
    ])
    
    # 添加各容量点的信息
    for name, results in comparison_results.items():
        # 计算节省电费
        electricity_cost = results.get('annual_electricity_cost', 0)  # 添加默认值避免报错
        if electricity_cost == 0 and 'annual_electricity_cost' not in results:
            # 如果没有计算年电费，需要从monthly_costs计算
            monthly_costs = results.get('monthly_costs', {})
            electricity_cost = sum(sum(month_data.values()) for month_data in monthly_costs.values()) if monthly_costs else 0
            electricity_cost += results.get('transformer_basic_cost', 0)
            
        savings = annual_electricity_cost_without_storage - electricity_cost
        savings_percent = (savings / annual_electricity_cost_without_storage * 100) if annual_electricity_cost_without_storage > 0 else 0
        
        comparison_table.append([
            name,
            f"{results['power']:.2f} kW / {results['capacity']:.2f} kWh",
            f"{results['avg_daily_charge']:.2f} kWh",
            f"{results['avg_daily_discharge']:.2f} kWh",
            f"{results['equivalent_full_cycles']:.2f}",
            f"{results['capacity_utilization']:.2f}%",
            f"{results['system_utilization']:.2f}%",
            f"{results['load_match_ability']:.2f}%",
            f"{results['capacity_assessment']} (节省电费:{savings:.2f}元,{savings_percent:.2f}%)"
        ])
    
    # 输出表格
    headers = ["容量点", "功率/容量", "平均日充电量", "平均日放电量", "等效满充次数", "容量利用率", "系统利用率", "负载匹配度", "容量评估"]
    print("\n比较结果表格:")
    
    # 计算每列宽度
    col_widths = [max(len(str(row[i])) for row in comparison_table + [headers]) for i in range(len(headers))]
    
    # 打印表头
    header_line = " | ".join(h.ljust(col_widths[i]) for i, h in enumerate(headers))
    print(header_line)
    print("-" * len(header_line))
    
    # 打印数据行
    for row in comparison_table:
        print(" | ".join(str(cell).ljust(col_widths[i]) for i, cell in enumerate(row)))
    
    # 结论和建议
    print("\n" + "=" * 100)
    print("结论和建议:")
    optimal_result = comparison_results["最佳容量点"]
    other_result = comparison_results["指定容量点"]
    
    # 比较两种容量的电费节省能力
    if optimal_result["load_match_ability"] > other_result["load_match_ability"]:
        print(f"1. 最佳容量点({optimal_result['capacity']:.2f} kWh)在负载峰值削减方面表现更好，可能带来更多电费节省。")
    else:
        print(f"1. 指定容量点({other_result['capacity']:.2f} kWh)在负载峰值削减方面表现较好。")
    
    # 比较系统利用效率
    if optimal_result["system_utilization"] > other_result["system_utilization"]:
        print(f"2. 最佳容量点系统利用率更高({optimal_result['system_utilization']:.2f}%)，投资回报可能更好。")
    else:
        print(f"2. 指定容量点系统利用率较高({other_result['system_utilization']:.2f}%)。")
    
    # 比较容量利用率
    if optimal_result["capacity_utilization"] > other_result["capacity_utilization"]:
        print(f"3. 最佳容量点更充分利用了电池容量({optimal_result['capacity_utilization']:.2f}%)。")
    else:
        print(f"3. 指定容量点更充分利用了电池容量({other_result['capacity_utilization']:.2f}%)。")
    
    # 最终建议
    print("\n最终建议:")
    if optimal_result["system_utilization"] > other_result["system_utilization"] and \
       optimal_result["load_match_ability"] > other_result["load_match_ability"]:
        print(f"推荐采用最佳容量点 {optimal_result['capacity']:.2f} kWh，预计将实现更好的经济效益和系统性能。")
    elif other_result["capacity_assessment"] == "容量合适" and \
         other_result["system_utilization"] > 70:
        print(f"指定容量点 {other_result['capacity']:.2f} kWh 也是可行的选择，特别是如果有其他非经济因素考虑。")
    else:
        print(f"综合考虑，推荐采用最佳容量点 {optimal_result['capacity']:.2f} kWh，以获得更好的经济回报。")
    
    # 绘制对比图表
    plt.figure(figsize=(14, 10))
    
    # 创建子图
    plt.subplot(2, 2, 1)
    names = list(comparison_results.keys())
    values = [results["avg_daily_charge"] for results in comparison_results.values()]
    plt.bar(names, values, color=['blue', 'green'])
    plt.title('平均日充电量对比 (kWh)')
    plt.grid(axis='y')
    
    plt.subplot(2, 2, 2)
    values = [results["capacity_utilization"] for results in comparison_results.values()]
    plt.bar(names, values, color=['blue', 'green'])
    plt.title('容量利用率对比 (%)')
    plt.grid(axis='y')
    
    plt.subplot(2, 2, 3)
    values = [results["system_utilization"] for results in comparison_results.values()]
    plt.bar(names, values, color=['blue', 'green'])
    plt.title('系统利用率对比 (%)')
    plt.grid(axis='y')
    
    plt.subplot(2, 2, 4)
    values = [results["load_match_ability"] for results in comparison_results.values()]
    plt.bar(names, values, color=['blue', 'green'])
    plt.title('负载匹配度对比 (%)')
    plt.grid(axis='y')
    
    plt.tight_layout()
    plt.show()
    
    # 添加电费节省对比图
    plt.figure(figsize=(12, 8))
    
    # 创建电费对比图
    plt.subplot(2, 1, 1)
    
    # 准备数据
    labels = ['不安装储能系统'] + names
    
    # 计算每种方案的电费（如果没有年电费，则计算）
    electricity_costs = []
    electricity_costs.append(annual_electricity_cost_without_storage)  # 不安装储能系统的电费
    
    for results in comparison_results.values():
        # 计算年电费
        electricity_cost = results.get('annual_electricity_cost', 0)
        if electricity_cost == 0 and 'annual_electricity_cost' not in results:
            # 如果没有计算年电费，需要从monthly_costs计算
            monthly_costs = results.get('monthly_costs', {})
            electricity_cost = sum(sum(month_data.values()) for month_data in monthly_costs.values()) if monthly_costs else 0
            electricity_cost += results.get('transformer_basic_cost', 0)
        
        electricity_costs.append(electricity_cost)
    
    # 创建柱状图
    colors = ['red'] + ['green'] * len(names)
    plt.bar(labels, electricity_costs, color=colors)
    plt.title('各方案年度电费对比')
    plt.ylabel('年度电费 (元)')
    plt.grid(axis='y')
    
    # 显示节省金额
    for i, cost in enumerate(electricity_costs[1:], 1):
        savings = electricity_costs[0] - cost
        savings_percent = (savings / electricity_costs[0] * 100) if electricity_costs[0] > 0 else 0
        plt.text(i, cost, f'节省: {savings:.2f}元\n({savings_percent:.2f}%)', 
                ha='center', va='bottom', fontsize=9,
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.6))
    
    # 创建年度总成本对比图（电费+系统摊销）
    plt.subplot(2, 1, 2)
    
    # 准备数据
    total_costs = []
    total_costs.append(annual_electricity_cost_without_storage)  # 不安装储能系统的总成本
    
    for results in comparison_results.values():
        electricity_cost = results.get('annual_electricity_cost', 0)
        if electricity_cost == 0 and 'annual_electricity_cost' not in results:
            monthly_costs = results.get('monthly_costs', {})
            electricity_cost = sum(sum(month_data.values()) for month_data in monthly_costs.values()) if monthly_costs else 0
            electricity_cost += results.get('transformer_basic_cost', 0)
        
        # 计算系统摊销成本
        power = results['power']
        capacity = results['capacity']
        system_total_cost = system_unit_cost * capacity * 1000  # 转换为Wh
        annual_system_cost = system_total_cost / system_lifetime
        
        # 计算年度总成本（电费+系统摊销）
        total_cost = electricity_cost + annual_system_cost
        total_costs.append(total_cost)
    
    # 创建柱状图
    plt.bar(labels, total_costs, color=['red'] + ['blue'] * len(names))
    plt.title('各方案年度总成本对比 (电费+系统摊销)')
    plt.ylabel('年度总成本 (元)')
    plt.grid(axis='y')
    
    # 显示总成本差异
    for i, cost in enumerate(total_costs[1:], 1):
        cost_diff = total_costs[0] - cost
        plt.text(i, cost, f'差异: {cost_diff:.2f}元', 
                ha='center', va='bottom', fontsize=9,
                bbox=dict(boxstyle='round', facecolor='yellow', alpha=0.6))
    
    plt.tight_layout()
    plt.show()
    
    # 还原原始参数
    storage_capacity_per_system = original_capacity
    max_power_per_system = original_power
    
    # 添加菜单返回提示
    print("\n分析完成！按回车键返回主菜单...", end="")
    input()
    print()

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
        
        # 检查数据粒度
        is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
        if is_minute_data:
            print("已检测到15分钟粒度的负载数据。系统将自动处理此类数据并支持所有分析功能。")
            print("注意：套利分析和优化将基于小时级聚合数据，但可视化会显示原始15分钟数据。")
        
        print("数据加载成功！")

        while True:
            print("\n请选择要执行的操作:")
            print("1. 绘制每个月的电价柱状图")
            print("2. 绘制负载和电价曲线")
            print("3. 计算年度电费总成本")
            print("4. 储能系统分析")
            print("5. 寻找最佳储能系统容量")
            print("6. 最佳容量点和其他输入容量的对比")
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
                calculate_annual_cost(data)
            elif choice == '4':
                print("\n储能系统分析选项:")
                print("1. 单日图形化分析（可选择日期）")
                print("2. 按日批量分析")
                print("3. 15分钟数据分析（更精细粒度）")
                sub_choice = input("请选择: ")
                
                try:
                    if sub_choice == '1':
                        # 选择日期进行分析
                        date_str = input("请输入要分析的日期(格式: YYYY-MM-DD)，直接回车使用第一个日期: ")
                        try:
                            # 获取所有可用日期
                            if 'datetime' in data.columns:
                                all_dates = sorted(data['datetime'].dt.date.unique())
                            else:
                                all_dates = sorted(data['date'].dt.date.unique())
                                
                            # 选择日期
                            if date_str.strip():
                                analysis_date = pd.to_datetime(date_str).date()
                                if analysis_date not in all_dates:
                                    print(f"未找到日期 {date_str}，将使用第一个可用日期")
                                    analysis_date = all_dates[0]
                            else:
                                analysis_date = all_dates[0]
                                
                            # 过滤出当日数据
                            if 'datetime' in data.columns:
                                day_data = data[data['datetime'].dt.date == analysis_date].copy()
                            else:
                                day_data = data[data['date'].dt.date == analysis_date].copy()
                                
                            print(f"\n分析日期: {analysis_date}")
                            
                            # 使用模拟函数计算带储能系统的负载
                            modified_load, storage_levels = simulate_storage_system(day_data)
                            
                            # 计算节省的电费
                            original_cost = 0
                            modified_cost = 0
                            
                            is_minute_data = 'is_minute_data' in day_data.columns and day_data['is_minute_data'].iloc[0]
                            time_factor = 0.25 if is_minute_data else 1.0  # 15分钟数据的时间因子
                            
                            for i, row in day_data.iterrows():
                                original_cost += row['load'] * row['price'] * time_factor
                                modified_cost += modified_load.loc[i] * row['price'] * time_factor
                                
                            savings = original_cost - modified_cost
                            savings_percent = (savings / original_cost * 100) if original_cost > 0 else 0
                            
                            # 绘制图表
                            plt.figure(figsize=(12, 10))
                            
                            # 子图1：负载对比
                            plt.subplot(3, 1, 1)
                            time_labels = [f"{h:02d}:{m:02d}" for h, m in zip(day_data['hour'], day_data['minute'] if 'minute' in day_data.columns else [0]*len(day_data))]
                            
                            plt.plot(time_labels, day_data['load'], 'b-', label='原始负载')
                            plt.plot(time_labels, modified_load, 'r-', label='带储能系统负载')
                            plt.title(f'日期: {analysis_date} 负载对比')
                            plt.ylabel('负载 (kW)')
                            plt.legend()
                            plt.grid(True)
                            
                            # 子图2：电价和时段
                            plt.subplot(3, 1, 2)
                            
                            # 绘制电价曲线
                            plt.plot(time_labels, day_data['price'], 'g-', label='电价')
                            plt.ylabel('电价 (元/kWh)')
                            plt.legend(loc='upper left')
                            plt.grid(True)
                            
                            # 子图3：储能电量变化
                            plt.subplot(3, 1, 3)
                            plt.plot(time_labels, storage_levels, 'c-')
                            plt.title('储能系统电量变化')
                            plt.ylabel('电量 (kWh)')
                            plt.xlabel('时间')
                            plt.grid(True)
                            
                            # 添加电费节省信息
                            plt.figtext(0.5, 0.01, 
                                      f'原始电费: {original_cost:.2f}元 | 优化后电费: {modified_cost:.2f}元 | 节省: {savings:.2f}元 ({savings_percent:.2f}%)',
                                      ha='center', bbox={'facecolor':'yellow', 'alpha':0.5, 'pad':5})
                            
                            plt.tight_layout(rect=[0, 0.03, 1, 0.97])
                            plt.show()
                            
                        except Exception as e:
                            print(f"日期分析失败: {e}")
                            traceback.print_exc()
                            
                    elif sub_choice == '2':
                        # 调用批量分析功能 
                        plot_storage_system(data, auto_analyze=True)
                        print("储能系统批量分析完成！")
                        
                    elif sub_choice == '3':
                        # 15分钟数据分析
                        if 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]:
                            # 选择日期进行分析
                            date_str = input("请输入要分析的日期(格式: YYYY-MM-DD)，直接回车使用第一个日期: ")
                            
                            try:
                                # 获取所有可用日期
                                all_dates = sorted(data['datetime'].dt.date.unique())
                                
                                # 选择日期
                                if date_str.strip():
                                    analysis_date = pd.to_datetime(date_str).date()
                                    if analysis_date not in all_dates:
                                        print(f"未找到日期 {date_str}，将使用第一个可用日期")
                                        analysis_date = all_dates[0]
                                else:
                                    analysis_date = all_dates[0]
                                    
                                # 过滤出当日数据
                                day_data = data[data['datetime'].dt.date == analysis_date].copy()
                                
                                print(f"\n分析日期: {analysis_date}，使用15分钟粒度数据")
                                
                                # 使用模拟函数计算带储能系统的负载
                                modified_load, storage_levels = simulate_storage_system(day_data)
                                
                                # 计算节省的电费
                                original_cost = 0
                                modified_cost = 0
                                
                                # 15分钟数据的时间因子
                                time_factor = 0.25
                                
                                for i, row in day_data.iterrows():
                                    original_cost += row['load'] * row['price'] * time_factor
                                    modified_cost += modified_load.loc[i] * row['price'] * time_factor
                                    
                                savings = original_cost - modified_cost
                                savings_percent = (savings / original_cost * 100) if original_cost > 0 else 0
                                
                                # 绘制图表
                                plt.figure(figsize=(15, 12))
                                
                                # 子图1：负载对比
                                plt.subplot(3, 1, 1)
                                time_labels = [f"{h:02d}:{m:02d}" for h, m in zip(day_data['hour'], day_data['minute'])]
                                
                                plt.plot(time_labels, day_data['load'], 'b-', label='原始负载')
                                plt.plot(time_labels, modified_load, 'r-', label='带储能系统负载')
                                plt.title(f'日期: {analysis_date} 负载对比 (15分钟粒度)')
                                plt.ylabel('负载 (kW)')
                                plt.legend()
                                plt.grid(True)
                                plt.xticks(rotation=90)
                                
                                # 子图2：电价和时段
                                plt.subplot(3, 1, 2)
                                
                                # 绘制电价曲线
                                plt.plot(time_labels, day_data['price'], 'g-', label='电价')
                                plt.ylabel('电价 (元/kWh)')
                                plt.legend(loc='upper left')
                                plt.grid(True)
                                plt.xticks(rotation=90)
                                
                                # 子图3：储能电量变化
                                plt.subplot(3, 1, 3)
                                plt.plot(time_labels, storage_levels, 'c-')
                                plt.title('储能系统电量变化')
                                plt.ylabel('电量 (kWh)')
                                plt.xlabel('时间')
                                plt.grid(True)
                                plt.xticks(rotation=90)
                                
                                # 添加电费节省信息
                                plt.figtext(0.5, 0.01, 
                                          f'原始电费: {original_cost:.2f}元 | 优化后电费: {modified_cost:.2f}元 | 节省: {savings:.2f}元 ({savings_percent:.2f}%)',
                                          ha='center', bbox={'facecolor':'yellow', 'alpha':0.5, 'pad':5})
                                
                                plt.tight_layout(rect=[0, 0.03, 1, 0.97])
                                plt.show()
                                
                            except Exception as e:
                                print(f"15分钟数据分析失败: {e}")
                                traceback.print_exc()
                        else:
                            print("当前数据不是15分钟粒度，无法进行15分钟精度分析")
                    else:
                        print("无效的选项，返回主菜单")
                except Exception as e:
                    print(f"储能系统分析失败: {e}")
                    traceback.print_exc()
            elif choice == '5':
                find_optimal_storage_capacity(data)
            elif choice == '6':
                try:
                    compare_storage_capacities(data)
                    print("容量对比分析完成！")
                except Exception as e:
                    print(f"容量对比分析失败: {e}")
                    traceback.print_exc()
            elif choice == '0':
                break
            else:
                print("无效的选项，请重新选择")
    except Exception as e:
        print(f"执行过程中发生错误: {e}")
        print("详细错误信息:")
        print(traceback.format_exc())

if __name__ == "__main__":
    main() 
