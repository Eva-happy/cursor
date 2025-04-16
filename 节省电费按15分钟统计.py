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

def simulate_storage_system(data, storage_capacity):
    """
    模拟储能系统运行并显示可视化结果
    data: 包含datetime, load, price, period_type的DataFrame
    storage_capacity: 储能系统容量 (kWh)
    
    返回: 修改后的负载和最终的储能电量
    """
    # 检查是否为15分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    if is_minute_data:
        # 对于15分钟粒度数据，首先按小时聚合
        print("检测到15分钟粒度数据，将按小时聚合进行分析和可视化...")
        
        # 按日期和小时聚合数据
        hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
            'load': 'mean',
            'price': 'mean',
            'period_type': lambda x: x.mode()[0] if not x.mode().empty else None,
            'month': 'first',
            'datetime': 'first'
        }).reset_index()
        
        # 确保索引正确
        if 'level_0' in hourly_data.columns:
            hourly_data = hourly_data.rename(columns={'level_0': 'date'})
        
        # 创建datetime列（如果需要）
        if 'datetime' not in hourly_data.columns:
            try:
                hourly_data['datetime'] = pd.to_datetime(hourly_data['date']) + pd.to_timedelta(hourly_data['hour'], unit='h')
            except Exception as e:
                print(f"创建datetime列失败: {e}")
        
        # 使用聚合后的小时数据
        analysis_data = hourly_data
        analysis_data['is_minute_data'] = False
    else:
        # 使用原始小时数据
        analysis_data = data
    
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
        return None, None
    
    # 计算全年模拟结果
    monthly_original_cost = {}
    monthly_modified_cost = {}
    monthly_savings = {}
    
    # 储能初始电量
    storage_level = initial_storage_capacity
    modified_loads = {}
    
    # 创建图形
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
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
    
    # 按日期处理数据
    for current_date in unique_dates:
        # 获取当日数据
        day_data = analysis_data[
            (analysis_data['date'] == current_date) if 'date' in analysis_data.columns 
            else (analysis_data['datetime'].dt.date == current_date)
        ].copy()
        
        # 对于不完整的数据，我们不再跳过，而是填充缺失的小时
        if len(day_data) < 24:
            print(f"日期 {current_date} 数据不完整，有 {len(day_data)} 小时的数据。将以负载为0填充缺失小时。")
            
            # 获取当前存在的小时
            existing_hours = set(day_data['hour'])
            
            # 找出缺失的小时
            missing_hours = set(range(24)) - existing_hours
            
            # 对于缺失的小时，我们使用0负载值填充
            if not day_data.empty:
                avg_price = day_data['price'].mean()  # 仍需要获取价格的平均值
                common_period = day_data['period_type'].mode()[0] if not day_data['period_type'].mode().empty else 3
                month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
                
                # 创建填充数据
                fill_rows = []
                for hour in missing_hours:
                    # 使用0负载填充，不使用平均值
                    fill_data = {
                        'hour': hour,
                        'load': 0,  # 设置为0，而不是平均值
                        'price': avg_price,
                        'period_type': common_period,
                        'month': month
                    }
                    
                    # 添加日期列
                    if 'date' in day_data.columns:
                        fill_data['date'] = current_date
                    
                    # 添加datetime列
                    if 'datetime' in day_data.columns:
                        try:
                            fill_data['datetime'] = pd.to_datetime(current_date) + pd.to_timedelta(hour, unit='h')
                        except:
                            pass
                    
                    fill_rows.append(fill_data)
                
                # 添加填充行
                if fill_rows:
                    fill_df = pd.DataFrame(fill_rows)
                    day_data = pd.concat([day_data, fill_df], ignore_index=True)
                    
                    # 按小时排序
                    day_data = day_data.sort_values('hour')
            
        # 获取月份
        month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
        
        # 获取原始负载和价格
        original_load = day_data['load'].values
        prices = day_data['price'].values
        periods = day_data['period_type'].values
        
        # 执行简单的充放电策略
        modified_load = original_load.copy()
        
        for i in range(len(day_data)):
            current_load = original_load[i]
            current_price = prices[i]
            
            # 充放电策略
            if current_price > np.mean(prices):  # 电价高于平均值时放电
                discharge = min(storage_level, max_power_per_system)
                # 确保放电量不超过当前企业用电负荷
                discharge = min(discharge, current_load)
                
                if discharge > 0:
                    storage_level -= discharge / efficiency_bess
                    modified_load[i] -= discharge
            else:  # 电价低于平均值时充电
                charge = min(max_power_per_system, storage_capacity - storage_level)
                if charge > 0:
                    storage_level += charge * efficiency_bess
                    modified_load[i] += charge
        
        # 存储修改后的负载
        modified_loads[current_date] = modified_load
        
        # 计算当日电费
        original_cost = sum(original_load * prices)
        modified_cost = sum(modified_load * prices)
        daily_savings = original_cost - modified_cost
        
        # 累加到月度统计
        if month not in monthly_savings:
            monthly_savings[month] = 0
            monthly_original_cost[month] = 0
            monthly_modified_cost[month] = 0
        
        monthly_savings[month] += daily_savings
        monthly_original_cost[month] += original_cost
        monthly_modified_cost[month] += modified_cost
    
    # 清空图表
    ax1.clear()
    ax2.clear()
    
    # 绘制月度电费节省图表
    months = sorted(monthly_savings.keys())
    ax1.bar(months, [monthly_original_cost[m] for m in months], width=0.4, label='原始电费', color='lightblue')
    ax1.bar([m+0.4 for m in months], [monthly_modified_cost[m] for m in months], width=0.4, label='优化后电费', color='lightcoral')
    ax1.set_title('月度电费对比')
    ax1.set_xlabel('月份')
    ax1.set_ylabel('电费 (元)')
    ax1.legend()
    ax1.grid(True, axis='y')
    
    # 为每个月标注节省金额和比例
    for i, month in enumerate(months):
        savings = monthly_savings[month]
        percent = (savings / monthly_original_cost[month] * 100) if monthly_original_cost[month] > 0 else 0
        ax1.text(month, monthly_original_cost[month] + 2000, 
                f'节省: {savings:.0f}元\n({percent:.1f}%)', 
                ha='center', fontsize=9)
    
    # 绘制随机一天的负载曲线
    random_date = unique_dates[len(unique_dates) // 2]  # 选取中间的日期作为样例
    
    # 获取该日数据
    day_data = analysis_data[
        (analysis_data['date'] == random_date) if 'date' in analysis_data.columns 
        else (analysis_data['datetime'].dt.date == random_date)
    ].copy().sort_values('hour')
    
    hours = day_data['hour'].values
    original_load = day_data['load'].values
    modified_load = modified_loads[random_date]
    periods = day_data['period_type'].values
    
    # 绘制负载曲线
    width = 0.35
    ax2.bar(hours - width/2, original_load, width, label='原始负载', color='lightblue')
    ax2.bar(hours + width/2, modified_load, width, label='加入储能系统后负载', color='lightcoral')
    
    # 显示时段背景
    for i in range(24):
        if i < len(periods):
            period_type = periods[i]
            ax2.axvspan(i-0.5, i+0.5, alpha=0.1, color=period_colors.get(period_type, 'gray'))
    
    # 完善图表
    ax2.legend(loc='upper right')
    ax2.set_title(f'负载曲线示例 - {random_date}')
    ax2.set_xlabel('小时')
    ax2.set_ylabel('负载 (kW)')
    ax2.set_xticks(range(24))
    ax2.grid(True)
    
    # 计算并显示总体统计信息
    total_original = sum(monthly_original_cost.values())
    total_modified = sum(monthly_modified_cost.values())
    total_savings = sum(monthly_savings.values())
    total_percent = (total_savings / total_original * 100) if total_original > 0 else 0
    
    system_cost = storage_capacity * 1000 * 2.5  # 假设储能系统单价2.5元/Wh
    roi_percent = (total_savings / system_cost * 100) if system_cost > 0 else 0
    payback_years = system_cost / total_savings if total_savings > 0 else float('inf')
    
    fig.suptitle(f'储能系统分析（容量: {storage_capacity} kWh, 功率: {max_power_per_system} kW）\n'
                f'年度节省: {total_savings:.2f}元 ({total_percent:.2f}%), 投资回收期: {payback_years:.2f}年', 
                fontsize=12)
    
    fig.tight_layout()
    plt.show()
    
    # 输出分析结果
    print(f"\n储能系统经济性分析:")
    print(f"系统容量: {storage_capacity} kWh")
    print(f"系统功率: {max_power_per_system} kW")
    print(f"系统造价: {system_cost:.2f} 元 (假设单价2.5元/Wh)")
    print(f"年度节省: {total_savings:.2f} 元")
    print(f"年收益率: {roi_percent:.2f}%")
    print(f"投资回收期: {payback_years:.2f} 年")
    
    return modified_loads, storage_level

def plot_storage_system(data, initial_date_index=0, auto_analyze=False):
    """绘制增加储能系统后的曲线（两充两放策略）并分析套利模式"""
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    if is_minute_data:
        # 对于15分钟数据，我们按小时聚合进行分析和可视化
        print("检测到15分钟粒度的数据，将按小时聚合进行分析和可视化...")
        
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
    
    if len(unique_dates) == 0:
        print("数据中没有有效的日期，无法分析")
        return
    
    # 初始化月度统计变量（无论是自动模式还是交互式模式）
    monthly_savings = {}
    monthly_original_cost = {}
    monthly_modified_cost = {}
        
    # 如果是自动分析模式，则直接执行所有日期分析
    if auto_analyze:
        print("\n开始全年储能系统套利分析...")
        print("日期\t\t原始电费(元)\t优化后电费(元)\t节省金额(元)\t节省比例(%)")
        print("-" * 80)
        
        # 储能系统参数信息
        print(f"储能系统容量: {storage_capacity_per_system} kWh, 功率: {max_power_per_system} kW")
        
        total_days = len(unique_dates)
        for day_idx, current_date in enumerate(unique_dates):
            # 获取当天数据用于分析
            day_data = analysis_data[
                (analysis_data['date'] == current_date) if 'date' in analysis_data.columns 
                else (analysis_data['datetime'].dt.date == current_date)
            ].copy()
            
            # 对于不完整的数据，我们不再跳过，而是填充缺失的小时
            if len(day_data) < 24:
                print(f"日期 {current_date} 数据不完整，有 {len(day_data)} 小时的数据。将以负载为0填充缺失小时。")
                
                # 获取当前存在的小时
                existing_hours = set(day_data['hour'])
                
                # 找出缺失的小时
                missing_hours = set(range(24)) - existing_hours
                
                # 对于缺失的小时，我们使用0负载值填充
                if not day_data.empty:
                    avg_price = day_data['price'].mean()  # 仍需要获取价格的平均值
                    common_period = day_data['period_type'].mode()[0] if not day_data['period_type'].mode().empty else 3
                    month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
                    
                    # 创建填充数据
                    fill_rows = []
                    for hour in missing_hours:
                        # 使用0负载填充，不使用平均值
                        fill_data = {
                            'hour': hour,
                            'load': 0,  # 设置为0，而不是平均值
                            'price': avg_price,
                            'period_type': common_period,
                            'month': month
                        }
                        
                        # 添加日期列
                        if 'date' in day_data.columns:
                            fill_data['date'] = current_date
                        
                        # 添加datetime列
                        if 'datetime' in day_data.columns:
                            try:
                                fill_data['datetime'] = pd.to_datetime(current_date) + pd.to_timedelta(hour, unit='h')
                            except:
                                pass
                        
                        fill_rows.append(fill_data)
                    
                    # 添加填充行
                    if fill_rows:
                        fill_df = pd.DataFrame(fill_rows)
                        day_data = pd.concat([day_data, fill_df], ignore_index=True)
                        
                        # 按小时排序
                        day_data = day_data.sort_values('hour')
                
            # 获取当月
            if 'month' in day_data.columns:
                month = day_data['month'].iloc[0]
            else:
                # 如果没有month列，从日期获取
                if isinstance(current_date, pd.Timestamp):
                    month = current_date.month
                else:
                    try:
                        month = pd.to_datetime(current_date).month
                    except:
                        month = -1  # 无法获取月份
            
            # 执行储能模拟
            storage_level = initial_storage_capacity
            original_load = day_data['load'].values
            periods = day_data['period_type'].values
            
            # 执行套利分析 - 使用两充两放策略
            # 第一次套利
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None:
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
            else:
                discharge1_start = None
            
            # 第二次套利
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
                else:
                    discharge2_start = None
            else:
                charge2_start = None
                discharge2_start = None
            
            # 执行充放电模拟，获取修改后的负载
            modified_load = original_load.copy()
            
            # 模拟第一次充电
            if charge1_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess))
                for h in range(charge1_start, charge1_start + charge1_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
            
            # 模拟第一次放电
            if discharge1_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration)
                for h in range(discharge1_start, discharge1_start + discharge1_duration):
                    if h >= len(modified_load): break
                    # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
            
            # 模拟第二次充电
            if charge2_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess))
                for h in range(charge2_start, charge2_start + charge2_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
            
            # 模拟第二次放电
            if discharge2_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration)
                for h in range(discharge2_start, discharge2_start + discharge2_duration):
                    if h >= len(modified_load): break
                    # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
            
            # 计算电费
            try:
                prices = day_data['price'].values
                original_cost = sum(original_load * prices)
                modified_cost = sum(modified_load * prices)
                daily_savings = original_cost - modified_cost
                savings_percent = (daily_savings / original_cost * 100) if original_cost > 0 else 0
                
                # 输出当天结果
                date_str = current_date.strftime('%Y-%m-%d') if hasattr(current_date, 'strftime') else str(current_date)
                print(f"{date_str}\t{original_cost:.2f}\t\t{modified_cost:.2f}\t\t{daily_savings:.2f}\t\t{savings_percent:.2f}")
                
                # 累计到月度统计
                if month not in monthly_savings:
                    monthly_savings[month] = 0
                    monthly_original_cost[month] = 0
                    monthly_modified_cost[month] = 0
                
                monthly_savings[month] += daily_savings
                monthly_original_cost[month] += original_cost
                monthly_modified_cost[month] += modified_cost
                
                # 显示进度
                if (day_idx + 1) % 10 == 0 or day_idx == total_days - 1:
                    print(f"已分析 {day_idx + 1}/{total_days} 天 ({(day_idx + 1)/total_days*100:.1f}%)...")
            
            except Exception as e:
                print(f"分析日期 {current_date} 时出错: {e}")
    else:
        # 交互式模式下的绘图和分析
        
        # 定义状态变量和控制变量
        current_date_index = initial_date_index
        current_date = unique_dates[current_date_index]
        
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
        
        # 创建图形
        fig = plt.figure(figsize=(14, 10))
        gs = fig.add_gridspec(3, 1, height_ratios=[4, 4, 1])
        ax1 = fig.add_subplot(gs[0])  # 负载曲线
        ax2 = fig.add_subplot(gs[1])  # 储能状态曲线
        ax3 = fig.add_subplot(gs[2])  # 控制按钮区域
        
        # 添加交互按钮
        prev_ax = plt.axes([0.7, 0.05, 0.1, 0.04])
        next_ax = plt.axes([0.81, 0.05, 0.1, 0.04])
        prev_btn = Button(prev_ax, '前一天')
        next_btn = Button(next_ax, '后一天')
        
        def update_plot(date_index):
            nonlocal current_date_index
            current_date_index = date_index % len(unique_dates)
            current_date = unique_dates[current_date_index]
            
            # 清空图表
            ax1.clear()
            ax2.clear()
            
            # 获取当天数据
            day_data = analysis_data[
                (analysis_data['date'] == current_date) if 'date' in analysis_data.columns 
                else (analysis_data['datetime'].dt.date == current_date)
            ].copy()
            
            # 对于不完整的数据，我们不再跳过，而是填充缺失的小时
            if len(day_data) < 24:
                print(f"日期 {current_date} 数据不完整，有 {len(day_data)} 小时的数据。将以负载为0填充缺失小时。")
                
                # 获取当前存在的小时
                existing_hours = set(day_data['hour'])
                
                # 找出缺失的小时
                missing_hours = set(range(24)) - existing_hours
                
                # 对于缺失的小时，我们使用0负载值填充
                if not day_data.empty:
                    avg_price = day_data['price'].mean()  # 仍需要获取价格的平均值
                    common_period = day_data['period_type'].mode()[0] if not day_data['period_type'].mode().empty else 3
                    month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
                    
                    # 创建填充数据
                    fill_rows = []
                    for hour in missing_hours:
                        # 使用0负载填充，不使用平均值
                        fill_data = {
                            'hour': hour,
                            'load': 0,  # 设置为0，而不是平均值
                            'price': avg_price,
                            'period_type': common_period,
                            'month': month
                        }
                        
                        # 添加日期列
                        if 'date' in day_data.columns:
                            fill_data['date'] = current_date
                        
                        # 添加datetime列
                        if 'datetime' in day_data.columns:
                            try:
                                fill_data['datetime'] = pd.to_datetime(current_date) + pd.to_timedelta(hour, unit='h')
                            except:
                                pass
                        
                        fill_rows.append(fill_data)
                    
                    # 添加填充行
                    if fill_rows:
                        fill_df = pd.DataFrame(fill_rows)
                        day_data = pd.concat([day_data, fill_df], ignore_index=True)
                        
                        # 按小时排序
                        day_data = day_data.sort_values('hour')
                
            # 准备横坐标
            x_hours = range(24)  # 整点小时
            
            # 执行储能模拟分析
            storage_level = initial_storage_capacity
            original_load = day_data['load'].values
            periods = day_data['period_type'].values
            
            # 执行套利分析 - 使用两充两放策略
            # 第一次套利
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None:
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
            else:
                discharge1_start = None
            
            # 第二次套利
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
                else:
                    discharge2_start = None
            else:
                charge2_start = None
                discharge2_start = None
            
            # 执行充放电逻辑模拟，获取修改后的负载
            modified_load = original_load.copy()
            storage_levels = [storage_level]  # 记录每小时的储能电量
            
            # 模拟第一次充电
            if charge1_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess))
                for h in range(charge1_start, charge1_start + charge1_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
                    storage_levels.append(storage_level)
            
            # 模拟第一次放电
            if discharge1_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration)
                for h in range(discharge1_start, discharge1_start + discharge1_duration):
                    if h >= len(modified_load): break
                    # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
                    storage_levels.append(storage_level)
            
            # 模拟第二次充电
            if charge2_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess))
                for h in range(charge2_start, charge2_start + charge2_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
                    storage_levels.append(storage_level)
            
            # 模拟第二次放电
            if discharge2_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration)
                for h in range(discharge2_start, discharge2_start + discharge2_duration):
                    if h >= len(modified_load): break
                    # 确保放电量不超过当前企业用电负荷，避免向电网反向输电
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
                    storage_levels.append(storage_level)
                    
            # 在剩余时间段内填充储能状态数据
            while len(storage_levels) < 25:  # 确保有24小时+初始状态的数据
                storage_levels.append(storage_level)
            
            # 计算当天各时段电费
            original_cost = sum(original_load * day_data['price'].values)
            modified_cost = sum(modified_load * day_data['price'].values)
            daily_savings = original_cost - modified_cost
            
            # 绘制负载曲线 - 使用小时级别的柱状图
            width = 0.35
            # 确保小时索引按顺序排列
            hours = sorted(day_data['hour'].values)
            
            if len(hours) < 24:
                print(f"警告：小时数据不连续，可能影响显示")
                
            # 绘制柱状图
            ax1.bar(hours - width/2, original_load, width, label='原始负载', color='lightblue')
            ax1.bar(hours + width/2, modified_load, width, label='加入储能系统后负载', color='lightcoral')
            
            # 标记充放电时段
            if charge1_start is not None:
                ax1.axvspan(charge1_start, charge1_start + charge1_duration, alpha=0.2, color='green', label='第一次充电')
            if discharge1_start is not None:
                ax1.axvspan(discharge1_start, discharge1_start + discharge1_duration, alpha=0.2, color='red', label='第一次放电')
            if charge2_start is not None:
                ax1.axvspan(charge2_start, charge2_start + charge2_duration, alpha=0.2, color='green')
            if discharge2_start is not None:
                ax1.axvspan(discharge2_start, discharge2_start + discharge2_duration, alpha=0.2, color='red')
            
            # 绘制时段背景
            for i in range(24):
                if i < len(periods):
                    period_type = periods[i]
                    ax1.axvspan(i-0.5, i+0.5, alpha=0.1, color=period_colors.get(period_type, 'gray'))
            
            # 完善图表
            ax1.legend()
            ax1.set_title(f'负载曲线 - {current_date} (节省: {daily_savings:.2f}元)')
            ax1.set_xlabel('小时')
            ax1.set_ylabel('负载 (kW)')
            ax1.set_xticks(range(24))
            ax1.grid(True)
            
            # 绘制储能电量曲线
            ax2.plot(range(len(storage_levels)), storage_levels, 'g-', linewidth=2)
            ax2.set_title('储能系统电量')
            ax2.set_xlabel('小时')
            ax2.set_ylabel('储能电量 (kWh)')
            ax2.set_xlim(0, 24)
            ax2.set_xticks(range(0, 25, 2))
            ax2.grid(True)
            
            # 计算并显示今天的统计信息
            ax2.text(0.02, 0.95, f'储能容量: {storage_capacity_per_system:.1f} kWh\n'
                               f'最大功率: {max_power_per_system:.1f} kW\n'
                               f'今日节省: {daily_savings:.2f} 元',
                     transform=ax2.transAxes, fontsize=9,
                     bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.8))
            
            # 刷新图形
            fig.tight_layout()
            fig.canvas.draw_idle()
            
            # 累加到月度统计
            if month not in monthly_savings:
                monthly_savings[month] = 0
                monthly_original_cost[month] = 0
                monthly_modified_cost[month] = 0
            
            monthly_savings[month] += daily_savings
            monthly_original_cost[month] += original_cost
            monthly_modified_cost[month] += modified_cost
            
            # 显示进度
            if (day_idx + 1) % 10 == 0 or day_idx == total_days - 1:
                print(f"已分析 {day_idx + 1}/{total_days} 天 ({(day_idx + 1)/total_days*100:.1f}%)...")
        
        # 输出月度统计
        print("\n月度电费节省统计:")
        print("月份\t原始电费(元)\t优化后电费(元)\t节省金额(元)\t节省比例(%)")
        print("-" * 80)
        
        total_original = 0
        total_modified = 0
        total_savings = 0
        
        for month in sorted(monthly_savings.keys()):
            original = monthly_original_cost[month]
            modified = monthly_modified_cost[month]
            savings = monthly_savings[month]
            percent = (savings / original * 100) if original > 0 else 0
            
            print(f"{month}\t{original:.2f}\t\t{modified:.2f}\t\t{savings:.2f}\t\t{percent:.2f}")
            
            total_original += original
            total_modified += modified
            total_savings += savings
        
        # 输出年度总结
        total_percent = (total_savings / total_original * 100) if total_original > 0 else 0
        print("-" * 80)
        print(f"全年\t{total_original:.2f}\t\t{total_modified:.2f}\t\t{total_savings:.2f}\t\t{total_percent:.2f}")
        
        # 储能系统年度收益率计算
        system_cost = storage_capacity_per_system * 1000 * 2.5  # 假设储能系统单价2.5元/Wh
        roi_percent = (total_savings / system_cost * 100) if system_cost > 0 else 0
        payback_years = system_cost / total_savings if total_savings > 0 else float('inf')
        
        print("\n储能系统经济性分析:")
        print(f"系统容量: {storage_capacity_per_system} kWh")
        print(f"系统功率: {max_power_per_system} kW")
        print(f"系统造价: {system_cost:.2f} 元 (假设单价2.5元/Wh)")
        print(f"年度节省: {total_savings:.2f} 元")
        print(f"年收益率: {roi_percent:.2f}%")
        print(f"投资回收期: {payback_years:.2f} 年")
    
    # 还原原始参数
    storage_capacity_per_system = original_capacity
    max_power_per_system = original_power
    
    # 删除以下部分，因为在此函数中没有定义costs、power_range和cost_breakdown变量
    # # 找出最优容量
    # optimal_power_idx = np.argmin(costs)
    # optimal_power = power_range[optimal_power_idx]
    # optimal_capacity = optimal_power * capacity_power_ratio
    # optimal_cost = costs[optimal_power_idx]
    # 
    # print("\n" + "=" * 100)
    # print(f'最佳储能系统功率: {optimal_power} kW')
    # print(f'最佳储能系统容量: {optimal_capacity:.2f} kWh')
    # print(f'年度电费(含变压器基本电费): {cost_breakdown[optimal_power]["annual_electricity_cost"]:.2f} 元')
    # print(f'储能系统总造价: {cost_breakdown[optimal_power]["system_total_cost"]:.2f} 元')
    # print(f'储能系统年摊销成本: {cost_breakdown[optimal_power]["annual_system_cost"]:.2f} 元')
    # print(f'年度总成本(电费+系统摊销): {optimal_cost:.2f} 元')
    # print("=" * 100)
    # 
    # # 绘制容量-成本曲线
    # plt.figure(figsize=(12, 8))
    # 
    # # 创建三条曲线
    # plt.plot(power_range, [cost_breakdown[p]['annual_electricity_cost'] for p in power_range], 'g-o', label='年度电费')
    # plt.plot(power_range, [cost_breakdown[p]['annual_system_cost'] for p in power_range], 'r-o', label='储能系统年摊销成本')
    # plt.plot(power_range, costs, 'b-o', label='年度总成本')
    # 
    # plt.title('储能系统功率-成本关系曲线')
    # plt.xlabel('储能系统功率 (kW)')
    # plt.ylabel('成本 (元)')
    # plt.grid(True)
    # plt.legend()
    # 
    # # 标记最优点
    # plt.plot(optimal_power, optimal_cost, 'mo', markersize=10)
    # plt.annotate(f'最优功率: {optimal_power} kW\n最优容量: {optimal_capacity:.2f} kWh\n最低年度总成本: {optimal_cost:.2f} 元', 
    #              xy=(optimal_power, optimal_cost), 
    #              xytext=(optimal_power + 200, optimal_cost - 50000 if optimal_power < 1600 else optimal_cost + 50000),
    #              arrowprops=dict(facecolor='black', shrink=0.05, width=1.5),
    #              bbox=dict(boxstyle="round,pad=0.5", fc="yellow", alpha=0.8))
    # 
    # # 添加第二个图表：展示最优功率下的成本构成
    # plt.figure(figsize=(10, 6))
    # 
    # # 获取最优功率下的各项成本
    # opt_data = cost_breakdown[optimal_power]
    # 
    # # 准备饼图数据
    # labels = ['电量电费', '变压器基本电费', '储能系统摊销成本']
    # sizes = [opt_data['electricity_cost'], opt_data['transformer_basic_cost'], opt_data['annual_system_cost']]
    # colors = ['lightgreen', 'lightblue', 'coral']
    # explode = (0.1, 0.1, 0.1)  # 突出显示所有部分
    # 
    # # 绘制饼图
    # plt.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', shadow=True, startangle=140)
    # plt.axis('equal')  # 确保饼图是圆的
    # plt.title(f'最优容量({optimal_capacity:.2f} kWh)下的年度成本构成')
    # 
    # plt.show()

def find_optimal_storage_capacity(data):
    """找到最佳的储能系统容量"""
    print("\n开始寻找最佳储能系统容量...")
    
    # 保存原始参数以便后续恢复
    global storage_capacity_per_system, max_power_per_system
    original_capacity = storage_capacity_per_system
    original_power = max_power_per_system
    
    # 设置分析参数范围
    min_power = 50  # kW
    max_power = 2000  # kW
    step = 50  # kW
    capacity_power_ratio = 2.1  # 容量与功率比例
    
    # 更新用户输入的参数范围
    try:
        user_min = float(input(f"请输入最小功率(kW)[默认{min_power}]: ") or min_power)
        user_max = float(input(f"请输入最大功率(kW)[默认{max_power}]: ") or max_power)
        user_step = float(input(f"请输入步长(kW)[默认{step}]: ") or step)
        user_ratio = float(input(f"请输入容量与功率比例[默认{capacity_power_ratio}]: ") or capacity_power_ratio)
        
        min_power = user_min
        max_power = user_max
        step = user_step
        capacity_power_ratio = user_ratio
    except ValueError:
        print("输入格式错误，使用默认值")
    
    # 创建功率范围
    power_range = list(range(int(min_power), int(max_power) + int(step), int(step)))
    
    # 检查是否为分钟粒度的数据
    is_minute_data = 'is_minute_data' in data.columns and data['is_minute_data'].iloc[0]
    
    if is_minute_data:
        # 对于15分钟数据，按小时聚合进行分析
        print("检测到15分钟粒度的数据，将按小时聚合进行分析...")
        
        # 按日期和小时聚合数据
        hourly_data = data.groupby([data['datetime'].dt.date, data['hour']]).agg({
            'load': 'mean',
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
        
        # 使用聚合后的小时数据
        analysis_data = hourly_data
        analysis_data['is_minute_data'] = False
    else:
        # 使用原始小时数据
        analysis_data = data
    
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
    
    # 变压器基本电费参数
    if method_basic_capacity_cost_transformer == 1:
        # 按容量收取
        print('变压器基本电费，按变压器容量收取')
        capacity_price = float(input('请输入容量单价（元/kVA·月）[默认42]: ') or 42)
        transformer_basic_cost = capacity_price * transformer_capacity * 12  # 12个月
    else:
        # 按需收取
        print('变压器基本电费，按需收取')
        demand_price = float(input('请输入需量单价（元/kW·月）[默认42]: ') or 42)
        
        # 计算每个月最大负载
        monthly_max_loads = analysis_data.groupby('month')['load'].max()
        transformer_basic_cost = sum(monthly_max_loads) * demand_price
    
    # 储能系统参数设置
    storage_life_years = float(input('请输入储能系统使用寿命(年)[默认10]: ') or 10)
    storage_price_per_kwh = float(input('请输入储能系统每kWh的价格(元)[默认2000]: ') or 2000)
    
    print(f"\n开始分析不同功率储能系统的经济性，这可能需要一些时间...")
    print(f"分析范围: {min_power}kW - {max_power}kW, 步长: {step}kW")
    
    # 存储不同功率下的成本
    costs = []
    cost_breakdown = {}
    
    # 逐个分析不同的功率
    for idx, power in enumerate(power_range):
        # 计算对应的储能容量
        capacity = power * capacity_power_ratio
        
        # 更新全局变量以供模拟使用
        max_power_per_system = power
        storage_capacity_per_system = capacity
        
        # 模拟结果
        monthly_savings = {}
        monthly_original_cost = {}
        monthly_modified_cost = {}
        
        print(f"分析进度: {idx+1}/{len(power_range)} - 功率: {power}kW, 容量: {capacity:.2f}kWh")
        
        # 按日期处理数据
        for current_date in unique_dates:
            # 获取当日数据
            day_data = analysis_data[
                (analysis_data['date'] == current_date) if 'date' in analysis_data.columns 
                else (analysis_data['datetime'].dt.date == current_date)
            ].copy()
            
            # 对于不完整的数据，填充缺失的小时
            if len(day_data) < 24:
                # 获取当前存在的小时
                existing_hours = set(day_data['hour'])
                
                # 找出缺失的小时
                missing_hours = set(range(24)) - existing_hours
                
                # 对于缺失的小时，使用0负载值填充
                if not day_data.empty:
                    avg_price = day_data['price'].mean()
                    common_period = day_data['period_type'].mode()[0] if not day_data['period_type'].mode().empty else 3
                    month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
                    
                    # 创建填充数据
                    fill_rows = []
                    for hour in missing_hours:
                        fill_data = {
                            'hour': hour,
                            'load': 0,  # 使用0而不是平均值
                            'price': avg_price,
                            'period_type': common_period,
                            'month': month
                        }
                        
                        if 'date' in day_data.columns:
                            fill_data['date'] = current_date
                        
                        if 'datetime' in day_data.columns:
                            try:
                                fill_data['datetime'] = pd.to_datetime(current_date) + pd.to_timedelta(hour, unit='h')
                            except:
                                pass
                        
                        fill_rows.append(fill_data)
                    
                    if fill_rows:
                        fill_df = pd.DataFrame(fill_rows)
                        day_data = pd.concat([day_data, fill_df], ignore_index=True)
                        day_data = day_data.sort_values('hour')
            
            # 获取月份
            month = day_data['month'].iloc[0] if 'month' in day_data.columns else pd.to_datetime(current_date).month
            
            # 获取原始负载、价格和时段类型
            original_load = day_data['load'].values
            prices = day_data['price'].values
            periods = day_data['period_type'].values
            
            # 执行储能套利策略
            storage_level = initial_storage_capacity
            modified_load = original_load.copy()
            
            # 执行套利分析 - 使用两充两放策略
            # 第一次套利
            charge1_start, charge1_duration, _ = find_continuous_window(periods, [VALLEY, DEEP_VALLEY], 2, 0)
            if charge1_start is None:
                charge1_start, charge1_duration, _ = find_continuous_window(periods, [FLAT], 2, 0)
            
            if charge1_start is not None:
                discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [SHARP], 2, charge1_start + charge1_duration)
                if discharge1_start is None:
                    discharge1_start, discharge1_duration, _ = find_continuous_window(periods, [PEAK], 2, charge1_start + charge1_duration)
            else:
                discharge1_start = None
            
            # 第二次套利
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
                else:
                    discharge2_start = None
            else:
                charge2_start = None
                discharge2_start = None
            
            # 模拟第一次充电
            if charge1_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge1_duration * efficiency_bess))
                for h in range(charge1_start, charge1_start + charge1_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
            
            # 模拟第一次放电
            if discharge1_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge1_duration)
                for h in range(discharge1_start, discharge1_start + discharge1_duration):
                    if h >= len(modified_load): break
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
            
            # 模拟第二次充电
            if charge2_start is not None:
                charge_needed = storage_capacity_per_system - storage_level
                power_charge = min(max_power_per_system, charge_needed / (charge2_duration * efficiency_bess))
                for h in range(charge2_start, charge2_start + charge2_duration):
                    if h >= len(modified_load): break
                    actual_charge = min(power_charge, max_power_per_system)
                    actual_charge = min(actual_charge, (storage_capacity_per_system - storage_level) / efficiency_bess)
                    if actual_charge <= 0: break
                    modified_load[h] += actual_charge
                    storage_level += actual_charge * efficiency_bess
                    storage_level = min(storage_level, storage_capacity_per_system)
            
            # 模拟第二次放电
            if discharge2_start is not None:
                discharge_available = storage_level
                power_discharge = min(max_power_per_system, discharge_available * efficiency_bess / discharge2_duration)
                for h in range(discharge2_start, discharge2_start + discharge2_duration):
                    if h >= len(modified_load): break
                    actual_discharge = min(power_discharge, max_power_per_system, original_load[h])
                    actual_discharge = min(actual_discharge, storage_level * efficiency_bess)
                    if actual_discharge <= 0: break
                    modified_load[h] -= actual_discharge
                    storage_level -= actual_discharge / efficiency_bess
                    storage_level = max(storage_level, 0)
            
            # 计算电费
            try:
                original_cost = sum(original_load * prices)
                modified_cost = sum(modified_load * prices)
                daily_savings = original_cost - modified_cost
                
                # 累计到月度统计
                if month not in monthly_savings:
                    monthly_savings[month] = 0
                    monthly_original_cost[month] = 0
                    monthly_modified_cost[month] = 0
                
                monthly_savings[month] += daily_savings
                monthly_original_cost[month] += original_cost
                monthly_modified_cost[month] += modified_cost
                
            except Exception as e:
                print(f"分析日期 {current_date} 时出错: {e}")
        
        # 计算年度总成本
        annual_electricity_cost = sum(monthly_modified_cost.values())
        system_total_cost = capacity * 1000 * (storage_price_per_kwh / 1000)  # 系统总造价
        annual_system_cost = system_total_cost / storage_life_years  # 系统年摊销成本
        
        # 计算总成本（包括电费和系统成本）
        total_annual_cost = annual_electricity_cost + transformer_basic_cost + annual_system_cost
        
        # 存储结果
        costs.append(total_annual_cost)
        cost_breakdown[power] = {
            'electricity_cost': sum(monthly_modified_cost.values()) - transformer_basic_cost,
            'transformer_basic_cost': transformer_basic_cost,
            'annual_electricity_cost': annual_electricity_cost,
            'system_total_cost': system_total_cost,
            'annual_system_cost': annual_system_cost,
            'total_annual_cost': total_annual_cost
        }
    
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

def get_latest_excel_files(directory_path):
    """
    获取指定目录下最新的两个Excel文件（用于价格和时段配置）
    返回: (price_file_path, period_file_path)
    """
    try:
        # 检查目录是否存在
        if not os.path.exists(directory_path):
            print(f"目录 {directory_path} 不存在！")
            return None, None
        
        # 获取目录中所有的Excel文件
        excel_files = []
        for file in glob.glob(os.path.join(directory_path, "*.xlsx")):
            excel_files.append(file)
        
        for file in glob.glob(os.path.join(directory_path, "*.xls")):
            excel_files.append(file)
        
        if not excel_files:
            print(f"目录 {directory_path} 中没有找到Excel文件！")
            print("请确保有电价配置文件和负载数据文件。")
            return None, None
        
        # 按修改时间排序文件
        excel_files.sort(key=os.path.getmtime, reverse=True)
        
        # 用户选择
        print("找到以下Excel文件:")
        for i, file in enumerate(excel_files):
            print(f"{i+1}. {os.path.basename(file)}")
        
        # 让用户选择电价配置文件和负载数据文件
        price_idx = int(input("\n请选择电价配置文件编号: ")) - 1
        if price_idx < 0 or price_idx >= len(excel_files):
            print("无效的选择！")
            return None, None
        
        period_idx = int(input("请选择负载数据文件编号: ")) - 1
        if period_idx < 0 or period_idx >= len(excel_files):
            print("无效的选择！")
            return None, None
        
        return excel_files[price_idx], excel_files[period_idx]
    
    except Exception as e:
        print(f"获取Excel文件时发生错误: {e}")
        return None, None

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
                find_optimal_storage_capacity(data)
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
