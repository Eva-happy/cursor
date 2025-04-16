import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider

# 设置中文字体 ------------------------------------------------------
plt.rcParams['font.sans-serif'] = ['SimHei']  # Windows系统用这个
# plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']  # Linux系统用这个
# plt.rcParams['font.sans-serif'] = ['STHeiti']  # macOS系统用这个
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 成本计算函数 ------------------------------------------------------
def linear_cost(q, base=100, step=5, rate=0.05):
    """线性增长：每step架增加rate%"""
    return base * (1 + (q // step) * rate)

def progressive_cost(q, base=150, rate=0.03):
    """递增增长：每架增加rate%"""
    return base * (1 + rate * q)

def stepwise_cost(q, base=200, threshold=10):
    """阶梯增长：每threshold架翻倍"""
    return base * np.ceil(q / threshold)

# 可视化设置 --------------------------------------------------------
fig, axs = plt.subplots(2, 2, figsize=(12, 8))
# 关键调整：增加子图间距 (hspace=行间距, wspace=列间距)
plt.subplots_adjust(left=0.08, right=0.92, 
                    top=0.88, bottom=0.25,
                    hspace=0.4, wspace=0.3)  # 原间距0.3 → 0.4

# 生成模拟数据
q = np.linspace(1, 20, 100)
l_cost = linear_cost(q)
p_cost = progressive_cost(q)
s_cost = stepwise_cost(q)
total = l_cost + p_cost + s_cost

# 绘制初始曲线（更新标签为中文）
line1, = axs[0,0].plot(q, l_cost, 'b-', label='接待区（阶梯）')
line2, = axs[0,1].plot(q, p_cost, 'g--', label='展示厅（线性）')
line3, = axs[1,0].plot(q, s_cost, 'r:', label='活动中心（递增）')
line4, = axs[1,1].plot(q, total, 'm-', label='总成本')

# 可视化增强（更新标题为中文）
titles = ['阶梯增长模式', '线性增长模式', '递增增长模式', '综合成本趋势']
for i, ax in enumerate(axs.flat):
    ax.set_xlabel('飞机数量（架）')  # 添加单位
    ax.set_ylabel('成本（万元）')   # 添加单位
    ax.set_title(titles[i], fontsize=12)  # 设置标题字号
    ax.grid(True)
    ax.legend(loc='upper left')  # 调整图例位置
    if i == 3:
        ax.annotate('成本跃升点', 
                   xy=(10, stepwise_cost(10)+linear_cost(10)), 
                   xytext=(8, 3000), 
                   arrowprops=dict(arrowstyle='->', color='darkred'),
                   fontsize=10)

# 添加交互控件 ----------------------------------------------------
ax_q = plt.axes([0.2, 0.1, 0.6, 0.03])
q_slider = Slider(ax=ax_q, label='飞机数量', valmin=1, valmax=20, valinit=10)

def update(val):
    current_q = q_slider.val
    # 更新各曲线
    line1.set_ydata(linear_cost(q))
    line2.set_ydata(progressive_cost(q))
    line3.set_ydata(stepwise_cost(q))
    line4.set_ydata(linear_cost(q) + progressive_cost(q) + stepwise_cost(q))
    
    # 标记当前值（添加中文标注）
    for ax in axs.flat[:-1]:
        for artist in ax.lines[1:]:
            artist.remove()
        ax.axvline(current_q, color='k', linestyle='--', alpha=0.5)
        ax.plot(current_q, linear_cost(current_q), 'bo') if ax == axs[0,0] else None
        ax.plot(current_q, progressive_cost(current_q), 'go') if ax == axs[0,1] else None
        ax.plot(current_q, stepwise_cost(current_q), 'ro') if ax == axs[1,0] else None
    
    fig.canvas.draw_idle()

q_slider.on_changed(update)

# 5. 智能标注系统 ----------------------------------------------
def smart_annotate(ax, x, text, y_ratio=0.8):
    """自动避让标注工具"""
    y_pos = ax.get_ylim()[1] * y_ratio
    ax.text(x + 0.3, y_pos, text, 
            color='#666666', fontsize=9, 
            bbox=dict(facecolor='white', alpha=0.8, edgecolor='none'))

# 添加阈值线标注（自动垂直避让）
for ax in [axs[0,0], axs[0,1], axs[1,0]]:
    for i, x in enumerate([5,10,15]):
        ax.axvline(x, color='gray', linestyle=':', alpha=0.3)
        smart_annotate(ax, x, f'扩容阈值{i+1}', y_ratio=0.75 - i*0.1)
def update(val):
    # ...（原有更新逻辑不变，需添加动态标注避让）
    # 在update函数中添加标注位置动态调整
    for ax in axs.flat[:-1]:
        for text in ax.texts:  # 清除旧标注
            text.set_visible(False)
        for i, x in enumerate([5,10,15]):
            smart_annotate(ax, x, f'扩容阈值{i+1}', y_ratio=0.75 - i*0.1)
    fig.canvas.draw_idle()
plt.show()
