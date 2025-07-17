import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

def create_energy_analysis_template():
    """创建能源项目分析模板Excel文件"""
    
    # 创建工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "能源项目经济分析"
    
    # 定义颜色方案
    colors = {
        'header_bg': 'D7E4BD',      # 淡绿色 - 主标题背景
        'section_bg': 'E2EFDA',     # 更淡的绿色 - 分区背景
        'input_bg': 'F2F2F2',       # 浅灰色 - 输入区背景
        'result_bg': 'FCE4D6',      # 淡橙色 - 结果区背景
        'table_bg': 'E7E6E6',       # 淡灰色 - 表格背景
        'highlight_bg': 'FFEB9C',   # 淡黄色 - 重点突出
        'border_color': '366092'     # 深蓝色 - 边框颜色
    }
    
    # 定义样式
    header_font = Font(name='微软雅黑', size=14, bold=True, color='2F4F4F')
    section_font = Font(name='微软雅黑', size=12, bold=True, color='2F4F4F')
    normal_font = Font(name='微软雅黑', size=10, color='2F4F4F')
    bold_font = Font(name='微软雅黑', size=10, bold=True, color='2F4F4F')
    
    thin_border = Border(
        left=Side(style='thin', color=colors['border_color']),
        right=Side(style='thin', color=colors['border_color']),
        top=Side(style='thin', color=colors['border_color']),
        bottom=Side(style='thin', color=colors['border_color'])
    )
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center')
    right_alignment = Alignment(horizontal='right', vertical='center')
    
    # 设置列宽
    column_widths = {
        'A': 18, 'B': 15, 'C': 15, 'D': 15, 'E': 18, 'F': 15, 'G': 15, 'H': 15,
        'I': 18, 'J': 15, 'K': 15, 'L': 15, 'M': 15, 'N': 15, 'O': 15
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # 1. 创建主标题
    ws.merge_cells('A1:O1')
    ws['A1'] = '能源项目经济分析模板'
    ws['A1'].font = Font(name='微软雅黑', size=16, bold=True, color='2F4F4F')
    ws['A1'].fill = PatternFill(start_color=colors['header_bg'], end_color=colors['header_bg'], fill_type='solid')
    ws['A1'].alignment = center_alignment
    ws['A1'].border = thin_border
    ws.row_dimensions[1].height = 30
    
    # 2. 左侧输入参数区域 (A3:D25)
    create_input_section(ws, colors, header_font, section_font, normal_font, bold_font, 
                        thin_border, center_alignment, left_alignment, right_alignment)
    
    # 3. 中间经济指标区域 (E3:I25)
    create_economic_indicators_section(ws, colors, header_font, section_font, normal_font, bold_font,
                                     thin_border, center_alignment, left_alignment, right_alignment)
    
    # 4. 右侧电价表区域 (J3:O25)
    create_electricity_price_section(ws, colors, header_font, section_font, normal_font, bold_font,
                                   thin_border, center_alignment, left_alignment, right_alignment)
    
    # 保存文件
    wb.save('能源项目分析模板.xlsx')
    print("Excel模板已生成：能源项目分析模板.xlsx")

def create_input_section(ws, colors, header_font, section_font, normal_font, bold_font,
                        thin_border, center_alignment, left_alignment, right_alignment):
    """创建输入参数区域"""
    
    # 区域标题
    ws.merge_cells('A3:D3')
    ws['A3'] = '1. 装机参数'
    ws['A3'].font = section_font
    ws['A3'].fill = PatternFill(start_color=colors['section_bg'], end_color=colors['section_bg'], fill_type='solid')
    ws['A3'].alignment = center_alignment
    ws['A3'].border = thin_border
    
    # 装机参数数据
    equipment_params = [
        ('装机功率 (MW)', '100'),
        ('装机容量 (MWh)', '200'),
        ('投资成本', '5'),
        ('', ''),
        ('年利用小时数', '1,500'),
        ('光伏容量配置效率', '98.50%'),
        ('逆变器效率', '99.30%'),
        ('线路损耗', '99.50%'),
        ('年平均利用小时', '99.00%'),
        ('REG-AFGP容量统计体系', '92.00%'),
        ('负载跟踪', '95.00%'),
        ('两班制AGC机组最小出力', '1.60%'),
        ('系统停电AGC机组最小出力', '1.60%'),
        ('两班制AGC机组补偿服务', '2.00%')
    ]
    
    start_row = 4
    for i, (param, value) in enumerate(equipment_params):
        row = start_row + i
        ws[f'A{row}'] = param
        ws[f'B{row}'] = value
        
        # 设置样式
        ws[f'A{row}'].font = normal_font
        ws[f'B{row}'].font = normal_font
        ws[f'A{row}'].fill = PatternFill(start_color=colors['input_bg'], end_color=colors['input_bg'], fill_type='solid')
        ws[f'B{row}'].fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        ws[f'A{row}'].border = thin_border
        ws[f'B{row}'].border = thin_border
        ws[f'A{row}'].alignment = left_alignment
        ws[f'B{row}'].alignment = right_alignment
    
    # 设备参数部分
    ws.merge_cells('A19:D19')
    ws['A19'] = '2. 设备参数'
    ws['A19'].font = section_font
    ws['A19'].fill = PatternFill(start_color=colors['section_bg'], end_color=colors['section_bg'], fill_type='solid')
    ws['A19'].alignment = center_alignment
    ws['A19'].border = thin_border
    
    # 设备参数数据
    equipment_data = [
        ('光伏发电改造设备', ''),
        ('逆变器大功率', '330'),
        ('逆变器功率万元/MW)', '0.00769'),
        ('光伏直流发电材万元/MW)', '0.0021325'),
        ('中位设备器材需台份中等改建', '0.3'),
        ('正常设备器材台份发电率', '0.00705'),
        ('万元功率电池功率', '0.001014444'),
        ('线路功率用电', '0.001345556'),
        ('万台万元功率改建整备投入', '0.010648444'),
        ('发电综合功率指', '0')
    ]
    
    start_row = 20
    for i, (param, value) in enumerate(equipment_data):
        row = start_row + i
        ws[f'A{row}'] = param
        ws[f'B{row}'] = value
        
        # 设置样式
        ws[f'A{row}'].font = normal_font
        ws[f'B{row}'].font = normal_font
        ws[f'A{row}'].fill = PatternFill(start_color=colors['input_bg'], end_color=colors['input_bg'], fill_type='solid')
        ws[f'B{row}'].fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        ws[f'A{row}'].border = thin_border
        ws[f'B{row}'].border = thin_border
        ws[f'A{row}'].alignment = left_alignment
        ws[f'B{row}'].alignment = right_alignment

def create_economic_indicators_section(ws, colors, header_font, section_font, normal_font, bold_font,
                                     thin_border, center_alignment, left_alignment, right_alignment):
    """创建经济指标区域"""
    
    # 区域标题
    ws.merge_cells('E3:I3')
    ws['E3'] = '项目经济指标 (核心结果)'
    ws['E3'].font = section_font
    ws['E3'].fill = PatternFill(start_color=colors['result_bg'], end_color=colors['result_bg'], fill_type='solid')
    ws['E3'].alignment = center_alignment
    ws['E3'].border = thin_border
    
    # 经济指标数据
    economic_indicators = [
        ('投资数据', '', '', ''),
        ('投资总额 (万元)', '12,000.00', '所需投资', '20.00'),
        ('建设期利率 (万元)', '5,000.00', '建设期利率 (%)', '20.00'),
        ('工程费费占投资比中', '7,000.00', '建设期年数 (%)', '6%'),
        ('设备费率占投资所比中', '33.00%', '建设投资回收', '2'),
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
        ('运营维护费 (规模投资万)', '0.00%', '固定成本费用 1', '13.00%'),
        ('管理费用', '100.00', '固定成本费用 2', '9.00%'),
        ('电力消费费 (在电电费用)', '500.00', '固定成本费用 3', '7.00%'),
        ('', '', '', ''),
        ('运营参数·市场参数', '', '', ''),
        ('电力系统维护管理期数', '83.33%', '运营管理期调整费用', '1.00%'),
        ('电力系统维修总费用 (万元)', '10000.00', '连运管理期过渡调整费', '1.00%'),
        ('', '', 'CO2排放', ''),
        ('', '', '综合投资', '20.00%'),
        ('', '', '生活费 (%)', '15.00'),
        ('', '', '年费', '-4.00%'),
        ('', '', '投货方式', '现金流+各配套+现金方式')
    ]
    
    start_row = 4
    for i, (indicator, value1, indicator2, value2) in enumerate(economic_indicators):
        row = start_row + i
        ws[f'E{row}'] = indicator
        ws[f'F{row}'] = value1
        ws[f'G{row}'] = indicator2
        ws[f'H{row}'] = value2
        
        # 设置样式
        for col in ['E', 'F', 'G', 'H']:
            cell = ws[f'{col}{row}']
            cell.font = bold_font if indicator and not value1 else normal_font
            cell.fill = PatternFill(start_color=colors['highlight_bg'] if indicator and not value1 else colors['result_bg'], 
                                  end_color=colors['highlight_bg'] if indicator and not value1 else colors['result_bg'], 
                                  fill_type='solid')
            cell.border = thin_border
            cell.alignment = center_alignment if col in ['E', 'G'] else right_alignment

def create_electricity_price_section(ws, colors, header_font, section_font, normal_font, bold_font,
                                   thin_border, center_alignment, left_alignment, right_alignment):
    """创建电价表区域"""
    
    # 区域标题
    ws.merge_cells('J3:O3')
    ws['J3'] = '电价表及参数配置'
    ws['J3'].font = section_font
    ws['J3'].fill = PatternFill(start_color=colors['table_bg'], end_color=colors['table_bg'], fill_type='solid')
    ws['J3'].alignment = center_alignment
    ws['J3'].border = thin_border
    
    # 第一个电价表
    price_table_1 = [
        ('1: 购电价格 (可打印)', '', '', '', '', ''),
        ('电价分别', '8.39%', '第1月', '0.03450', '330', '第1-20年'),
        ('综合发电电费 (万元)', '9.39', '第2月', '0.00140', '第1-20年', '0.0313'),
        ('全额费电电费(20年滚转)', '21.58%', '', '', '', ''),
        ('新能源配电费 (%)', '4.10', '', '', '', ''),
        ('', '', '', '', '', ''),
        ('EPC含税费', '', '', '', '', ''),
        ('EPC含税 (万元)', '12,000.00', '', '', '', ''),
        ('EPC含税本 (万元)', '20,000.00', '', '', '', ''),
        ('EPC含本 (万元)', '8,000.00', '', '', '', ''),
        ('', '', '', '', '', ''),
        ('运营管理投标费 (万元/年', '108.00', '', '', '', ''),
        ('运营管理投标费 (万元/年', '100.00', '', '', '', ''),
        ('运营项目费率 (万元/年', '-72.00', '', '', '', ''),
        ('运营管理投标费', '102.00', '', '', '', '')
    ]
    
    # 第二个电价表
    price_table_2 = [
        ('2: 容电价标准 (可打印)', '', '', '', '', ''),
        ('容电价综合', '100', '200', '第1-20年', '', ''),
        ('', '', '', '', '', ''),
        ('3: 市储电价利润电价表 (可打印)', '', '', '', '', ''),
        ('综合市储', '1', '6', '70%', '330', '第1-20年'),
        ('', '', '', '', '', ''),
        ('4: 实收能源普及电价表 (万元/年) (可打印)', '', '', '', '', ''),
        ('分段实收', '1', '2,500', '12', '第1-20年', ''),
        ('', '', '', '', '', ''),
        ('5: 购电标准市储电价表 (万元/年)', '', '', '', '', ''),
        ('电费市储', '2', '60', '330', '第21年', ''),
        ('', '', '3', '60', '330', '第3-21年')
    ]
    
    # 创建表格
    start_row = 4
    
    # 第一个表格
    for i, row_data in enumerate(price_table_1):
        row = start_row + i
        for j, value in enumerate(row_data):
            col = chr(ord('J') + j)
            ws[f'{col}{row}'] = value
            
            # 设置样式
            cell = ws[f'{col}{row}']
            cell.font = bold_font if row_data[0] and ':' in str(row_data[0]) else normal_font
            cell.fill = PatternFill(start_color=colors['highlight_bg'] if row_data[0] and ':' in str(row_data[0]) else colors['table_bg'],
                                  end_color=colors['highlight_bg'] if row_data[0] and ':' in str(row_data[0]) else colors['table_bg'],
                                  fill_type='solid')
            cell.border = thin_border
            cell.alignment = center_alignment if j == 0 else right_alignment
    
    # 第二个表格
    start_row2 = start_row + len(price_table_1) + 1
    for i, row_data in enumerate(price_table_2):
        row = start_row2 + i
        for j, value in enumerate(row_data):
            col = chr(ord('J') + j)
            ws[f'{col}{row}'] = value
            
            # 设置样式
            cell = ws[f'{col}{row}']
            cell.font = bold_font if row_data[0] and ':' in str(row_data[0]) else normal_font
            cell.fill = PatternFill(start_color=colors['highlight_bg'] if row_data[0] and ':' in str(row_data[0]) else colors['table_bg'],
                                  end_color=colors['highlight_bg'] if row_data[0] and ':' in str(row_data[0]) else colors['table_bg'],
                                  fill_type='solid')
            cell.border = thin_border
            cell.alignment = center_alignment if j == 0 else right_alignment

if __name__ == "__main__":
    create_energy_analysis_template()