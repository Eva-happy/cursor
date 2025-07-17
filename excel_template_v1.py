#!/usr/bin/env python3
"""
能源项目经济分析Excel模板生成器 V1
使用openpyxl创建完整格式化的Excel文件
"""

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("openpyxl未安装，将生成简化版本")

def create_excel_template():
    """创建Excel模板"""
    
    if not OPENPYXL_AVAILABLE:
        create_csv_template()
        return
        
    # 创建工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "能源项目经济分析"
    
    # 定义颜色方案
    colors = {
        'header': 'D7E4BD',      # 淡绿色 - 主标题
        'section': 'E2EFDA',     # 更淡绿色 - 分区标题
        'input': 'F2F2F2',       # 浅灰色 - 输入区
        'result': 'FCE4D6',      # 淡橙色 - 结果区
        'table': 'E7E6E6',       # 淡灰色 - 表格区
        'highlight': 'FFEB9C',   # 淡黄色 - 重点突出
    }
    
    # 设置列宽
    col_widths = [18, 15, 15, 15, 18, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    
    # 创建样式
    header_font = Font(name='微软雅黑', size=16, bold=True, color='2F4F4F')
    section_font = Font(name='微软雅黑', size=12, bold=True, color='2F4F4F')
    normal_font = Font(name='微软雅黑', size=10, color='2F4F4F')
    bold_font = Font(name='微软雅黑', size=10, bold=True, color='2F4F4F')
    
    thin_border = Border(
        left=Side(style='thin', color='366092'),
        right=Side(style='thin', color='366092'),
        top=Side(style='thin', color='366092'),
        bottom=Side(style='thin', color='366092')
    )
    
    # 1. 主标题
    ws.merge_cells('A1:O1')
    ws['A1'] = '能源项目经济分析模板'
    ws['A1'].font = header_font
    ws['A1'].fill = PatternFill(start_color=colors['header'], end_color=colors['header'], fill_type='solid')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].border = thin_border
    ws.row_dimensions[1].height = 30
    
    # 2. 区域标题
    ws.merge_cells('A3:D3')
    ws['A3'] = '1. 装机参数 & 设备参数'
    ws['A3'].font = section_font
    ws['A3'].fill = PatternFill(start_color=colors['section'], end_color=colors['section'], fill_type='solid')
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A3'].border = thin_border
    
    ws.merge_cells('E3:I3')
    ws['E3'] = '项目经济指标 (核心结果)'
    ws['E3'].font = section_font
    ws['E3'].fill = PatternFill(start_color=colors['result'], end_color=colors['result'], fill_type='solid')
    ws['E3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['E3'].border = thin_border
    
    ws.merge_cells('J3:O3')
    ws['J3'] = '电价表及参数配置'
    ws['J3'].font = section_font
    ws['J3'].fill = PatternFill(start_color=colors['table'], end_color=colors['table'], fill_type='solid')
    ws['J3'].alignment = Alignment(horizontal='center', vertical='center')
    ws['J3'].border = thin_border
    
    # 3. 填充数据
    fill_template_data(ws, colors, normal_font, bold_font, thin_border)
    
    # 保存文件
    wb.save('能源项目分析模板.xlsx')
    print("✅ Excel模板已生成：能源项目分析模板.xlsx")

def fill_template_data(ws, colors, normal_font, bold_font, thin_border):
    """填充模板数据"""
    
    # 左侧输入参数数据
    input_data = [
        ('装机功率 (MW)', '100'),
        ('装机容量 (MWh)', '200'),
        ('投资成本', '5'),
        ('', ''),
        ('年利用小时数', '1,500'),
        ('光伏容量配置效率', '98.50%'),
        ('逆变器效率', '99.30%'),
        ('线路损耗', '99.50%'),
        ('年平均利用小时', '99.00%'),
        ('REG-AGC容量统计体系', '92.00%'),
        ('负载跟踪', '95.00%'),
        ('调峰AGC机组最小出力', '1.60%'),
        ('系统AGC机组最小出力', '1.60%'),
        ('调峰AGC机组辅助服务', '2.00%'),
        ('', ''),
        ('2. 设备参数', ''),
        ('光伏发电设备投资', ''),
        ('逆变器功率', '330'),
        ('逆变器单价(万元/MW)', '0.00769'),
        ('光伏组件单价(万元/MW)', '0.0021325'),
        ('设备器材等其他费用', '0.3'),
        ('设备器材运输费率', '0.00705'),
        ('电池储能功率', '0.001014444'),
        ('线路接入费用', '0.001345556'),
        ('整体设备投资费用', '0.010648444'),
        ('发电综合效率', '0')
    ]
    
    # 中间经济指标数据
    economic_data = [
        ('投资数据', '', '', ''),
        ('投资总额 (万元)', '12,000.00', '所需投资', '20.00'),
        ('建设期利率 (万元)', '5,000.00', '建设期利率 (%)', '20.00'),
        ('工程费占投资比重', '7,000.00', '建设期年数 (%)', '6%'),
        ('设备费占投资比重', '33.00%', '建设投资回收', '2'),
        ('', '', '', ''),
        ('工程费用 (13%) (万元)', '1,560.00', '', ''),
        ('安装费用 (6%) (万元)', '720.00', '', ''),
        ('其他 (6%) (万元)', '720.00', '', ''),
        ('建设期监督费', '0.00', '', ''),
        ('设计费用 (%)', '20.00', '', ''),
        ('环保监测费', '65.00%', '', ''),
        ('工程监理费', '100.00%', '', ''),
        ('', '', '', ''),
        ('运营参数·成本参数', '', '', ''),
        ('运营维护费 (投资比例)', '0.00%', '固定成本费用1', '13.00%'),
        ('管理费用', '100.00', '固定成本费用2', '9.00%'),
        ('电力消费费', '500.00', '固定成本费用3', '7.00%'),
        ('', '', '', ''),
        ('运营参数·市场参数', '', '', ''),
        ('电力系统维护管理费率', '83.33%', '运营管理费调整', '1.00%'),
        ('电力系统维护总费用 (万元)', '10000.00', '运营管理过渡调整费', '1.00%'),
        ('', '', 'CO2排放', ''),
        ('', '', '综合投资', '20.00%'),
        ('', '', '税费 (%)', '15.00'),
        ('', '', '年费', '-4.00%'),
        ('', '', '投资方式', '现金流+分期+现金方式')
    ]
    
    # 右侧电价表数据
    price_data = [
        ('1: 购电价格 (可调整)', '', '', '', '', ''),
        ('电价分档', '8.39%', '第1档', '0.03450', '330', '第1-20年'),
        ('综合发电电费 (万元)', '9.39', '第2档', '0.00140', '第1-20年', '0.0313'),
        ('全额上网电费(20年滚转)', '21.58%', '', '', '', ''),
        ('新能源配额费 (%)', '4.10', '', '', '', ''),
        ('', '', '', '', '', ''),
        ('EPC合同费', '', '', '', '', ''),
        ('EPC含税 (万元)', '12,000.00', '', '', '', ''),
        ('EPC含税本 (万元)', '20,000.00', '', '', '', ''),
        ('EPC成本 (万元)', '8,000.00', '', '', '', ''),
        ('', '', '', '', '', ''),
        ('运营管理投标费 (万元/年)', '108.00', '', '', '', ''),
        ('运营管理投标费 (万元/年)', '100.00', '', '', '', ''),
        ('运营项目费率 (万元/年)', '-72.00', '', '', '', ''),
        ('运营管理投标费', '102.00', '', '', '', ''),
        ('', '', '', '', '', ''),
        ('2: 容量电价标准 (可调整)', '', '', '', '', ''),
        ('容量电价综合', '100', '200', '第1-20年', '', ''),
        ('', '', '', '', '', ''),
        ('3: 辅助服务电价表 (可调整)', '', '', '', '', ''),
        ('综合辅助服务', '1', '6', '70%', '330', '第1-20年'),
        ('', '', '', '', '', ''),
        ('4: 实时能源管理电价表 (万元/年)', '', '', '', '', ''),
        ('分时电价', '1', '2,500', '12', '第1-20年', ''),
        ('', '', '', '', '', ''),
        ('5: 购售电标准电价表 (万元/年)', '', '', '', '', ''),
        ('电费标准', '2', '60', '330', '第21年', ''),
        ('', '', '3', '60', '330', '第3-21年')
    ]
    
    # 填充左侧输入参数
    for i, (param, value) in enumerate(input_data):
        row = 4 + i
        ws[f'A{row}'] = param
        ws[f'B{row}'] = value
        
        # 设置样式
        ws[f'A{row}'].font = bold_font if param.startswith('2.') else normal_font
        ws[f'B{row}'].font = normal_font
        
        if param.startswith('2.'):
            ws[f'A{row}'].fill = PatternFill(start_color=colors['highlight'], end_color=colors['highlight'], fill_type='solid')
        else:
            ws[f'A{row}'].fill = PatternFill(start_color=colors['input'], end_color=colors['input'], fill_type='solid')
        
        ws[f'B{row}'].fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        ws[f'A{row}'].border = thin_border
        ws[f'B{row}'].border = thin_border
        ws[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
        ws[f'B{row}'].alignment = Alignment(horizontal='right', vertical='center')
    
    # 填充中间经济指标
    for i, (indicator, value1, indicator2, value2) in enumerate(economic_data):
        row = 4 + i
        ws[f'E{row}'] = indicator
        ws[f'F{row}'] = value1
        ws[f'G{row}'] = indicator2
        ws[f'H{row}'] = value2
        
        # 设置样式
        is_header = indicator and not value1 and indicator.endswith('参数')
        font = bold_font if is_header else normal_font
        bg_color = colors['highlight'] if is_header else colors['result']
        
        for col in ['E', 'F', 'G', 'H']:
            cell = ws[f'{col}{row}']
            cell.font = font
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if col in ['E', 'G'] else 'right', vertical='center')
    
    # 填充右侧电价表
    for i, row_data in enumerate(price_data):
        row = 4 + i
        for j, value in enumerate(row_data):
            col = chr(ord('J') + j)
            ws[f'{col}{row}'] = value
            
            # 设置样式
            is_header = row_data[0] and ':' in str(row_data[0])
            font = bold_font if is_header else normal_font
            bg_color = colors['highlight'] if is_header else colors['table']
            
            cell = ws[f'{col}{row}']
            cell.font = font
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if j == 0 else 'right', vertical='center')

def create_csv_template():
    """创建CSV模板（备用方案）"""
    import csv
    
    template_data = []
    
    # CSV数据
    csv_data = [
        ['能源项目经济分析模板'] + [''] * 14,
        [''] * 15,
        ['1. 装机参数 & 设备参数', '', '', '', '项目经济指标 (核心结果)', '', '', '', '', '电价表及参数配置', '', '', '', '', ''],
        ['装机功率 (MW)', '100', '', '', '投资数据', '', '', '', '', '1: 购电价格 (可调整)', '', '', '', '', ''],
        ['装机容量 (MWh)', '200', '', '', '投资总额 (万元)', '12,000.00', '所需投资', '20.00', '', '电价分档', '8.39%', '第1档', '0.03450', '330', '第1-20年']
    ]
    
    with open('能源项目分析模板.csv', 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerows(csv_data)
    
    print("✅ CSV模板已生成：能源项目分析模板.csv")

if __name__ == "__main__":
    create_excel_template()