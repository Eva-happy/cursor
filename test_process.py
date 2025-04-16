import pandas as pd
import numpy as np
import re

def is_header_row(row_values, row_data):
    """判断一行是否为表头行
    Args:
        row_values: 行中非空的值列表
        row_data: 原始行数据
    Returns:
        bool: 是否为表头行
    """
    # 检查是否包含关键字
    header_keywords = ['用电分类', '用电类别', '电压等级', '基本电价', '分类', '类别', '项目']
    has_keywords = any(keyword in val for val in row_values for keyword in header_keywords)
    
    # 检查是否包含单位信息（包括括号内的内容）
    has_units = any('元' in val or '（' in val or '(' in val for val in row_values)
    
    # 检查是否全为空或None
    is_empty_or_none = all(val in ['', 'None', None] or pd.isna(val) for val in row_data)
    
    # 检查是否包含独立数字
    has_numbers = any(re.match(r'^-?\d+\.?\d*$', str(val)) for val in row_values)
    
    # 如果行全为空或None，且不包含数字，则视为表头的一部分
    if is_empty_or_none and not has_numbers:
        return True
    
    # 如果包含单位信息或关键字，且不包含独立数字，则视为表头
    if (has_units or has_keywords) and not has_numbers:
        return True
    
    return False

def find_header_and_data_rows(df):
    """查找表头行和数据行
    Args:
        df: DataFrame对象
    Returns:
        tuple: (header_rows列表, data_rows列表)
    """
    header_rows = []
    data_rows = []
    
    # 从下往上扫描找到包含数字的行（数据行）
    for i in range(len(df) - 1, -1, -1):
        row = df.iloc[i]
        row_values = [str(val).strip() for val in row if pd.notna(val)]
        
        # 检查该行是否包含独立的数字
        has_independent_number = False
        for val in row_values:
            try:
                float_val = float(val.replace(',', ''))
                has_independent_number = True
                break
            except ValueError:
                continue
        
        if has_independent_number:
            data_rows.append(i)
    
    if not data_rows:
        return [], []
    
    # 获取最上面的数据行
    first_data_row = min(data_rows)
    
    # 从数据行向上查找表头行
    for i in range(first_data_row - 1, -1, -1):
        row = df.iloc[i]
        row_values = [str(val).strip() for val in row if pd.notna(val)]
        
        if is_header_row(row_values, row):
            header_rows.append(i)
        elif len(header_rows) > 0:
            # 如果已经找到了表头行，且当前行不是表头，则停止搜索
            break
    
    return header_rows, data_rows

def process_first_page(page_data):
    """处理第一页数据（主电价表）"""
    try:
        if not page_data or 'table' not in page_data:
            print("没有找到表格数据")
            return None, [], None
        
        table_data = page_data['table']
        if not table_data:
            print("表格数据为空")
            return None, [], None
        
        # 创建DataFrame
        df = pd.DataFrame(table_data)
        df = df.dropna(axis=1, how='all')
        
        # 查找表头行和数据行
        header_rows, data_rows = find_header_and_data_rows(df)
        
        if not header_rows or not data_rows:
            print("未找到表头行或数据行")
            return None, [], None
        
        # 获取表头范围
        first_header_row = min(header_rows)
        last_header_row = max(header_rows)
        first_data_row = min(data_rows)
        
        # 提取并合并表头
        headers = df.iloc[first_header_row:last_header_row + 1]
        final_columns = []
        for col in range(len(headers.columns)):
            column_values = [str(val).strip() for val in headers.iloc[:, col] if pd.notna(val) and str(val).strip() not in ['', 'None']]
            if column_values:
                final_columns.append('\n'.join(column_values))
            else:
                final_columns.append(f"Column_{col+1}")
        
        # 提取数据部分
        data = df.iloc[first_data_row:]
        data.columns = final_columns[:len(data.columns)]
        
        # 清理数据
        data = data.replace('', pd.NA)
        data = data.dropna(how='all')
        
        # 处理数值列
        numeric_columns = [col for col in data.columns if any(keyword in str(col) for keyword in ['价格', '费用', '电价', '金额', '附加', '数值'])]
        for col in numeric_columns:
            if col in data.columns:
                data[col] = data[col].astype(str).apply(lambda x: x.replace('\n', '').strip() if pd.notnull(x) else x)
                data[col] = pd.to_numeric(data[col], errors='coerce').round(6)
        
        # 获取注释和计算说明
        notes = page_data.get('notes', [])
        calc_df = None
        
        # 查找计算说明表
        calc_start_row = None
        for i in range(len(df)):
            row_values = [str(val).strip() for val in df.iloc[i] if pd.notna(val)]
            if any(keyword in val for val in row_values for keyword in ['计算说明', '计算方法', '计算公式']):
                calc_start_row = i
                break
        
        if calc_start_row is not None:
            calc_data = df.iloc[calc_start_row:]
            calc_df = pd.DataFrame(calc_data)
        
        return data, notes, calc_df
        
    except Exception as e:
        print(f"处理第一页数据失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, [], None

def process_second_page(page_data):
    """处理第二页数据（使用说明）"""
    try:
        if not page_data or 'table' not in page_data:
            print("没有找到表格数据")
            return None
        
        table_data = page_data['table']
        if not table_data:
            print("表格数据为空")
            return None
        
        # 创建DataFrame
        df = pd.DataFrame(table_data)
        df = df.dropna(axis=1, how='all')
        
        # 查找表头行和数据行
        header_rows, data_rows = find_header_and_data_rows(df)
        
        # 如果没有找到数据行，将所有非空行作为数据行
        if not data_rows:
            data_rows = [i for i in range(len(df)) if not df.iloc[i].isna().all()]
        
        if not data_rows:
            print("未找到有效行")
            return None
        
        first_data_row = min(data_rows)
        
        # 如果找到了表头行，使用表头作为列名
        if header_rows:
            first_header_row = min(header_rows)
            last_header_row = max(header_rows)
            
            headers = df.iloc[first_header_row:last_header_row + 1]
            final_columns = []
            for col in range(len(headers.columns)):
                column_values = [str(val).strip() for val in headers.iloc[:, col] if pd.notna(val) and str(val).strip() not in ['', 'None']]
                if column_values:
                    final_columns.append('\n'.join(column_values))
                else:
                    final_columns.append(f"Column_{col+1}")
            
            data = df.iloc[first_data_row:]
            data.columns = final_columns[:len(data.columns)]
        else:
            data = df.iloc[data_rows]
            data.columns = [f"Column_{i+1}" for i in range(len(data.columns))]
        
        # 清理数据
        data = data.replace('', pd.NA)
        data = data.dropna(how='all')
        
        return data
        
    except Exception as e:
        print(f"处理第二页数据失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

# 测试数据
test_data = {
    'table': [
        ['标题1', '标题2', '标题3', ''],
        ['用电分类', '电压等级', '基本电价', ''],
        ['None', 'None', '（元/千瓦·月）', '（元/千伏安·月）'],
        ['工商业用电', '1-10kV', '45.00', ''],
        ['工商业用电', '35-110kV', '42.50', ''],
        ['工商业用电', '220kV及以上', '40.00', ''],
        ['', '', '', ''],
        ['计算说明：', '', '', ''],
        ['1. 电价执行标准', '', '', ''],
        ['2. 计算方法说明', '', '', '']
    ],
    'notes': ['注1：以上电价为含税价格。', '注2：执行时间以通知为准。']
}

# 运行测试
print("=== 测试第一页处理 ===")
data, notes, calc_df = process_first_page(test_data)

if data is not None:
    print("\n处理后的数据:")
    print(data)
    
print("\n注释:")
print(notes)

if calc_df is not None:
    print("\n计算说明:")
    print(calc_df)

# 测试第二页处理
print("\n=== 测试第二页处理 ===")
test_data_page2 = {
    'table': [
        ['项目', '说明', '', ''],
        ['None', 'None', '（示例）', ''],
        ['类别1', '说明1', '', ''],
        ['类别2', '说明2', '', '']
    ]
}

data2 = process_second_page(test_data_page2)
if data2 is not None:
    print("\n第二页处理后的数据:")
    print(data2) 