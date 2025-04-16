import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider
import matplotlib.gridspec as gridspec

# 设置中文字体 ------------------------------------------------------
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 成本计算函数 ------------------------------------------------------
def linear_cost(q, base=100, step=5, rate=0.05):
    return base * (1 + (q // step) * rate)

def progressive_cost(q, base=150, rate=0.03):
    return base * (1 + rate * q)

def stepwise_cost(q, base=200, threshold=10):
    return base * np.ceil(q / threshold)

def control_center_cost(q, base=80):
    return np.full_like(q, base)

# 网格布局设置 ----------------------------------------------------
fig = plt.figure(figsize=(16, 12))
gs = gridspec.GridSpec(3, 2, height_ratios=[1, 1, 1.5], hspace=0.5)

# 子图分配
ax1 = fig.add_subplot(gs[0, 0])
ax2 = fig.add_subplot(gs[0, 1])
ax3 = fig.add_subplot(gs[1, 0])
ax4 = fig.add_subplot(gs[1, 1])
ax5 = fig.add_subplot(gs[2, :])

# 数据生成 --------------------------------------------------------
q = np.linspace(1, 20, 100)
l_cost = linear_cost(q)
p_cost = progressive_cost(q)
s_cost = stepwise_cost(q)
cc_cost = control_center_cost(q)
total = l_cost + p_cost + s_cost + cc_cost

# 绘制各模块曲线 --------------------------------------------------
line1, = ax1.plot(q, l_cost, 'b-', label='接待区（阶梯）')
line2, = ax2.plot(q, p_cost, 'g--', label='展示厅（线性）')
line3, = ax3.plot(q, s_cost, 'r:', label='活动中心（递增）')
line4, = ax4.plot(q, cc_cost, 'c-', label='调度中心（固定）')
line5, = ax5.plot(q, total, 'm-', label='总成本')

# 可视化增强（修复标注逻辑）-----------------------------------------
def find_nearest_idx(array, value):
    """ 找到最接近值的索引 """
    return np.argmin(np.abs(array - value))

target_q = 10
nearest_idx = find_nearest_idx(q, target_q)

titles = ['阶梯增长模式', '线性增长模式', '递增增长模式', '固定成本模式', '综合成本趋势']
for i, (ax, title) in enumerate(zip([ax1,ax2,ax3,ax4,ax5], titles)):
    ax.set(xlabel='飞机数量（架）', ylabel='成本（万元）', title=title)
    ax.grid(True)
    ax.legend(loc='upper left')

      

# 交互控件设置 ----------------------------------------------------
ax_slider = plt.axes([0.2, 0.05, 0.6, 0.03])
q_slider = Slider(ax_slider, '飞机数量', 1, 20, valinit=10, color='#0099CC')

# 修改交互更新部分代码
def update(val):
    current_q = q_slider.val
    # 更新所有曲线数据
    line1.set_ydata(linear_cost(q))
    line2.set_ydata(progressive_cost(q)) 
    line3.set_ydata(stepwise_cost(q))
    line4.set_ydata(control_center_cost(q))  # 直接使用函数返回的数组
    
    # 动态标记当前值（修复lambda表达式）
    for ax, marker, func in zip(
        [ax1, ax2, ax3, ax4], 
        ['bo', 'go', 'ro', 'co'],
        [linear_cost, progressive_cost, stepwise_cost, control_center_cost]  # 移除[0]索引
    ):
        ax.lines[-1].remove() if len(ax.lines) > 1 else None
        ax.plot(current_q, func(current_q), marker, markersize=8)
    
    fig.canvas.draw_idle()


q_slider.on_changed(update)
update(10)  # 初始触发

plt.show()
