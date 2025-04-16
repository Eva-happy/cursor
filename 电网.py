import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import pdfplumber
import warnings
import re
warnings.filterwarnings('ignore')
def get_pdf_url(page_url):
    """获取PDF文件链接"""
    driver = None
    try:
        print("正在配置Edge浏览器选项...")
        # 配置Edge选项
        edge_options = webdriver.EdgeOptions()
        edge_options.add_argument('--headless')  # 无头模式
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--disable-software-rasterizer')
        edge_options.add_argument('--disable-extensions')  # 禁用扩展
        edge_options.add_argument('--disable-logging')  # 禁用日志
        edge_options.add_argument('--disable-notifications')  # 禁用通知
        edge_options.add_argument('--ignore-certificate-errors')  # 忽略证书错误
        edge_options.add_argument('--log-level=3')  # 只显示致命错误
        edge_options.add_argument('--silent')  # 静默模式
        edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 禁用 DevTools 日志
        
        # 设置Edge浏览器路径
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
        
        edge_binary = None
        for path in edge_paths:
            if os.path.exists(path):
                edge_binary = path
                break
        
        if edge_binary:
            edge_options.binary_location = edge_binary
        
        # 创建Edge浏览器实例
        print("正在创建Edge浏览器实例...")
        try:
            webdriver_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "msedgedriver.exe")
            if not os.path.exists(webdriver_path):
                print(f"未找到WebDriver: {webdriver_path}")
                return None
                
            service = Service(executable_path=webdriver_path, log_output=os.devnull)  # 指定WebDriver路径并禁止日志输出
            driver = webdriver.Edge(service=service, options=edge_options)
            driver.set_page_load_timeout(3)  # 设置页面加载超时时间为3
            driver.set_script_timeout(3)  # 设置脚本执行超时时间为3秒
        except Exception as e:
            print("创建Edge浏览器实例失败，请确保已安装最新版本的Edge浏览器和WebDriver")
            return None
        
        # 访问网页
        print("正在访问网页...")
        try:
            driver.get(page_url)
        except Exception as e:
            print("访问网页失败，可能是网络问题")
            return None
        
        # 等待PDF链接元素出现
        print("正在等待PDF元素加载...")
        try:
            wait = WebDriverWait(driver, 5)  # 设置等待时间为5秒
            pdf_element = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'downloadFiles')]")))
            pdf_url = pdf_element.get_attribute('src')
            print(f"找到PDF链接: {pdf_url}")
            return pdf_url
        except Exception as e:
            print("等待PDF元素超时")
            return None
            
    except Exception as e:
        print("获取PDF链接失败")
        return None
    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass
def download_pdf(pdf_url, save_dir="downloads"):
    """下载PDF文件"""
    try:
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 生成文件名
        filename = f"电价表_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
        file_path = os.path.join(save_dir, filename)
        
        # 设置请求头
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Referer': 'https://95598.csg.cn/'
        }
        
        # 下载PDF文件
        response = requests.get(pdf_url, headers=headers, verify=False)
        
        # 检查响应状态和内容类型
        if response.status_code == 200 and response.headers.get('content-type', '').lower().startswith('application/pdf'):
            with open(file_path, 'wb') as f:
                f.write(response.content)
            print(f"PDF文件已保存到: {file_path}\n")
            return file_path
        else:
            print(f"下载失败，状态码: {response.status_code}, 内容类型: {response.headers.get('content-type')}")
            return None
    except Exception as e:
        print(f"下载PDF失败: {str(e)}")
        return None

def extract_table_from_pdf(pdf_path):
    """从PDF中提取表格数据"""
    try:
        all_pages_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_data = {}
                
                # 提取文本
                text = page.extract_text()
                lines = text.split('\n')
                
                # 提取标题（第一行通常是标题）
                title_line = None
                if lines:
                    title_line = lines[0].strip()
                    page_data['title'] = title_line
                    # 提取执行时间（通常在第二行，包含"执行时间"字样）
                    for line in lines[1:3]:  # 只查找前几行
                        if '执行时间' in line:
                            page_data['subtitle'] = line.strip()
                            break
                
                # 提取表格
                tables = page.extract_tables()
                if tables:
                    print(f"表格列数: {len(tables[0][0])}")
                    print(f"表头内容: {tables[0][0]}")
                    page_data['table'] = tables[0]
                
                # 提取注释
                notes = []
                current_note = ""
                in_notes = False
                
                for line in lines:
                    line = line.strip()
                    # 检查是否进入注释部分
                    if '注：' in line or '注:' in line:
                        in_notes = True
                        if current_note:
                            notes.append(current_note.strip())
                        current_note = line
                    # 如果在注释部分，继续收集注释内容
                    elif current_note and line and not line.startswith('执行时间'):
                        current_note += " " + line
                
                # 保存最后一条注释
                if current_note:
                    notes.append(current_note.strip())
                
                page_data['notes'] = notes
                
                # 提取单位信息
                for line in lines:
                    if '单位' in line:
                        page_data['unit'] = line.strip()
                        break
                
                all_pages_data.append(page_data)
        return all_pages_data
    except Exception as e:
        print(f"提取表格失败: {str(e)}")
        return None
def process_first_page(page_data):
    """处理第一页数据（主电价表）"""
    try:
        if not page_data or 'table' not in page_data:
            return None, None, None
            
        table_data = page_data['table']
        text = page_data['text']
        
        # 提取注释（在表格之后的文本）
        notes = []
        current_note = ""
        
        if text:
            # 分行处理文本
            lines = text.split('\n')
            for line in lines:
                line = line.strip()
                # 如果是新的注释项
                if line.startswith('注:') or line.startswith('注：') or any(line.startswith(f"{i}.") for i in range(1, 10)):
                    # 保存之前的注释（如果有）
                    if current_note:
                        notes.append(current_note.strip())
                    current_note = line
                # 如果是注释的继续内容
                elif current_note and line and not line.startswith('执行时间'):
                    current_note += " " + line
            
            # 保存最后一条注释
            if current_note:
                notes.append(current_note.strip())
        
        # 处理表格数据
        if not table_data:
            return None, notes, None
        
        # 创建DataFrame
        df = pd.DataFrame(table_data)
        
        # 删除全为空的列
        df = df.dropna(axis=1, how='all')
        
        # 找到表头行（包含"用电分类"的行）
        header_row = None
        for i in range(len(df)):
            if any('用电分类' in str(cell) for cell in df.iloc[i]):
                header_row = i
                break
        
        if header_row is None:
            return None, notes, None
        
        # 提取表头
        headers = df.iloc[header_row:header_row+2]
        
        # 合并表头（处理多行表头的情况）
        final_columns = []
        for col in range(len(headers.columns)):
            col_name = headers.iloc[0, col]
            sub_name = headers.iloc[1, col] if len(headers) > 1 else None
            
            if pd.isna(col_name) and not pd.isna(sub_name):
                final_columns.append(sub_name)
            elif not pd.isna(col_name):
                if not pd.isna(sub_name) and col_name != sub_name:
                    final_columns.append(f"{col_name}\n{sub_name}")
                else:
                    final_columns.append(col_name)
            else:
                final_columns.append(f"Column_{col+1}")
        
        # 提取数据部分
        data = df.iloc[header_row+2:]
        data.columns = final_columns[:len(data.columns)]
        
        # 清理数据
        data = data.replace('', pd.NA)
        data = data.dropna(how='all')  # 删除全为空的行
        
        # 识别数值列
        numeric_columns = []
        for col in data.columns:
            if any(keyword in str(col) for keyword in ['价格', '费用', '电价', '金额', '附加', '数值']):
                numeric_columns.append(col)
        
        # 处理数值列
        for col in numeric_columns:
            if col in data.columns:
                # 清理数值数据（移除换行符和空格）
                data[col] = data[col].astype(str).apply(lambda x: x.replace('\n', '').strip() if pd.notnull(x) else x)
                # 转换为数值并保持6位小数
                data[col] = pd.to_numeric(data[col], errors='coerce').round(6)
        
        # 提取计算说明表格
        calc_df = None
        calc_start = None
        
        # 在原始表格数据中查找计算说明表格的起始位置
        for i, row in enumerate(table_data):
            if any(str(cell).strip() == '名称' for cell in row):
                calc_start = i
                break
        
        if calc_start is not None:
            # 提取计算说明表格数据
            calc_data = table_data[calc_start:]
            # 创建DataFrame
            calc_df = pd.DataFrame(calc_data)
            # 使用第一行作为列名
            calc_df.columns = calc_df.iloc[0]
            # 删除作为列名的行
            calc_df = calc_df.iloc[1:]
            # 清理数据
            calc_df = calc_df.replace('', pd.NA)
            calc_df = calc_df.dropna(how='all')
            # 重置索引
            calc_df = calc_df.reset_index(drop=True)
            
            # 处理数值列
            if '数值' in calc_df.columns:
                calc_df['数值'] = calc_df['数值'].apply(lambda x: str(x).replace('\n', '').strip() if pd.notnull(x) else x)
                calc_df['数值'] = pd.to_numeric(calc_df['数值'], errors='coerce')
        
        return data, notes, calc_df
    except Exception as e:
        print(f"处理数据失败: {str(e)}")
        return None, None, None
def process_second_page(page_data):
    """处理第二页数据（使用说明）"""
    try:
        if not page_data or 'table' not in page_data:
            return None
            
        table_data = page_data['table']
        if not table_data:
            return None
            
        # 创建DataFrame
        df = pd.DataFrame(table_data)
        
        # 删除全为空的列和行
        df = df.dropna(axis=1, how='all')
        df = df.dropna(how='all')
        
        # 重置索引
        df = df.reset_index(drop=True)
        
        return df
    except Exception as e:
        print(f"处理第二页数据失败: {str(e)}")
        return None
def calculate_text_height(text, width, font_size=10):
    """计算文本需要的行高"""
    # 假设每个中文字符宽度为font_size，英文字符宽度为font_size/2
    # 每行能容纳的字符数
    chars_per_line = width * 2 // font_size  # 乘2是因为一个英文字符占半个中文字符宽度
    
    # 计算需要的行数
    text_length = sum(2 if ord(c) > 127 else 1 for c in text)  # 中文字符计2，英文字符计1
    lines = (text_length + chars_per_line - 1) // chars_per_line
    
    # 每行基础高度为font_size + 4（上下padding），最少一行
    return max(1, lines) * (font_size + 4)
def save_to_excel(df1, notes, calc_df, usage_df, page_contents, save_dir="downloads"):
    """保存为Excel文件"""
    try:
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 从第一页提取文件名前缀
        title = page_contents[0].get('title', '电价表')
        company_name = title.split('代理购电价格表')[0] if '代理购电价格表' in title else ''
        filename_prefix = f"{company_name}电价表" if company_name else "电价表"
        
        # 生成文件名
        filename = f"{filename_prefix}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(save_dir, filename)
        
        # 创建Excel写入器
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        workbook = writer.book
        
        # 设置格式
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        unit_format = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        number_format = workbook.add_format({
            'num_format': '0.000000',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        text_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        
        note_title_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'left',
            'font_size': 10,
            'bg_color': '#D9D9D9',  # 灰色背景
            'border': 1,
            'bold': True  # 加粗注释说明标题
        })
        
        note_format = workbook.add_format({
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'left',
            'font_size': 10,
            'border': 1
        })
        def write_sheet_with_title(worksheet, title, subtitle, col_count):
            """写入标题和副标题"""
            # 使用从PDF提取的实际标题
            actual_title = title if title else "电网公司代理购电价格表"
            worksheet.merge_range(f'A1:{chr(65+col_count-1)}1', actual_title, title_format)
            worksheet.merge_range(f'A2:{chr(65+col_count-1)}2', subtitle, subtitle_format)
            worksheet.set_row(0, 30)  # 设置标题行高
            worksheet.set_row(1, 25)  # 设置副标题行高
            return 3  # 返回下一个可用行号
        def write_notes(worksheet, notes, start_row, col_count, cell_width):
            """写入注释说明"""
            if not notes:
                return
            
            # 写入注释标题行
            worksheet.merge_range(start_row, 0, start_row, col_count-1, "注释说明：", note_title_format)
            worksheet.set_row(start_row, 20)  # 减小注释标题行高
            
            # 写入注释内容
            for i, note in enumerate(notes, start=1):
                # 设置固定行高为70
                worksheet.set_row(start_row + i, 70)
                # 合并单元格到整行，保持与上方表格同宽
                worksheet.merge_range(
                    start_row + i, 0,
                    start_row + i, col_count-1,
                    f"{i}. {note}", note_format
                )
        # 写入电价表sheet（第一页）
        if df1 is not None and page_contents:
            worksheet = writer.book.add_worksheet('第一页')
            
            # 从第一页提取标题和副标题
            title = page_contents[0].get('title', '')
            subtitle = page_contents[0].get('subtitle', f"（执行时间：{time.strftime('%Y')}年{time.strftime('%m')}月）")
            # 设置列宽
            col_widths = {
                'A': 15,  # 用电分类
                'B': 20,  # Column_2
                'C': 12,  # 电压等级
            }
            total_width = 0
            
            # 设置固定列宽
            for col, width in col_widths.items():
                worksheet.set_column(f'{col}:{col}', width)
                total_width += width
            
            # 其余列平均分配宽度
            remaining_cols = len(df1.columns) - len(col_widths)
            if remaining_cols > 0:
                remaining_width = 10  # 进一步减小其他列的统一宽度
                for i in range(len(col_widths), len(df1.columns)):
                    col_letter = chr(65 + i)
                    worksheet.set_column(f'{col_letter}:{col_letter}', remaining_width)
                    total_width += remaining_width
            
            worksheet.set_default_row(30)
            
            # 写入标题和副标题
            next_row = write_sheet_with_title(worksheet, title, subtitle, len(df1.columns))
            
            # 写入单位信息（从第一页提取）
            if page_contents and len(page_contents) > 0:
                first_page_unit = page_contents[0].get('unit', '')
                if first_page_unit:
                    worksheet.write(1, len(df1.columns)-2, first_page_unit, unit_format)
            # 写入表头
            for col, value in enumerate(df1.columns.values):
                worksheet.write(next_row, col, value, header_format)
            worksheet.set_row(next_row, 45)  # 增加表头行高
            next_row += 1
            
            # 写入数据并合并单元格
            current_name = None
            current_col2 = None
            merge_start_name = next_row
            merge_start_col2 = next_row
            
            # 写入数据
            for row, (_, data_row) in enumerate(df1.iterrows(), start=next_row):
                # 处理用电分类列的合并
                if pd.notnull(data_row['用电分类']):
                    if current_name != data_row['用电分类']:
                        if current_name is not None:
                            worksheet.merge_range(merge_start_name, 0, row-1, 0, current_name, text_format)
                        current_name = data_row['用电分类']
                        merge_start_name = row
                
                # 处理Column_2列的合并
                if pd.notnull(data_row['Column_2']):
                    if current_col2 != data_row['Column_2']:
                        if current_col2 is not None:
                            worksheet.merge_range(merge_start_col2, 1, row-1, 1, current_col2, text_format)
                        current_col2 = data_row['Column_2']
                        merge_start_col2 = row
                
                # 写入其他列
                for col, value in enumerate(data_row):
                    if col not in [0, 1] and pd.notnull(value):  # 跳过已处理的合并列
                        if isinstance(value, (int, float)):  # 数值列
                            worksheet.write(row, col, float(value), number_format)
                        else:  # 文本列
                            worksheet.write(row, col, value, text_format)
            
            # 处理最后一行的合并
            if current_name is not None:
                worksheet.merge_range(merge_start_name, 0, row, 0, current_name, text_format)
            if current_col2 is not None:
                worksheet.merge_range(merge_start_col2, 1, row, 1, current_col2, text_format)
            
            # 写入注释
            note_start_row = row + 2
            write_notes(worksheet, notes, note_start_row, len(df1.columns), total_width/len(df1.columns))
            
            # 冻结窗格
            worksheet.freeze_panes(next_row, 0)
        # 写入计算说明sheet
        if calc_df is not None:
            worksheet = writer.book.add_worksheet('计算说明')
            
            # 设置列宽
            col_widths = {
                'A': 20,  # 名称
                'B': 10,  # 序号
                'C': 50,  # 明细
                'D': 25,  # 计算关系
                'E': 15   # 数值
            }
            total_width = sum(col_widths.values())
            
            for col, width in col_widths.items():
                worksheet.set_column(f'{col}:{col}', width)
            
            worksheet.set_default_row(30)
            
            # 从第一页提取标题和副标题
            title = page_contents[0].get('title', '')
            subtitle = page_contents[0].get('subtitle', f"（执行时间：{time.strftime('%Y')}年{time.strftime('%m')}月）")
            next_row = write_sheet_with_title(worksheet, title, subtitle, len(col_widths))
            
            # 写入表头
            headers = ['名称', '序号', '明细', '计算关系', '数值']
            for col, header in enumerate(headers):
                worksheet.write(next_row, col, header, header_format)
            next_row += 1
            
            # 写入数据并合并单元格
            current_name = None
            merge_start = next_row
            
            # 写入数据
            for row, (_, data_row) in enumerate(calc_df.iterrows(), start=next_row):
                # 处理名称列的合并
                if pd.notnull(data_row['名称']):
                    if current_name != data_row['名称']:
                        if current_name is not None:
                            worksheet.merge_range(merge_start, 0, row-1, 0, current_name, text_format)
                        current_name = data_row['名称']
                        merge_start = row
                
                # 写入其他列
                for col, value in enumerate(data_row):
                    if col == 4:  # 数值列
                        if pd.notnull(value):
                            worksheet.write(row, col, float(value), number_format)
                    else:
                        if pd.notnull(value):
                            worksheet.write(row, col, value, text_format)
            
            # 处理最后一个名称的合并
            if current_name is not None:
                worksheet.merge_range(merge_start, 0, row, 0, current_name, text_format)
            
            # 写入注释
            note_start_row = row + 2
            write_notes(worksheet, notes, note_start_row, len(col_widths), total_width/len(col_widths))
            
            # 冻结窗格
            worksheet.freeze_panes(next_row, 0)
        # 写入使用说明sheet（第二页）
        if usage_df is not None and len(page_contents) > 1:
            worksheet = writer.book.add_worksheet('第二页')
            
            # 从第二页提取标题和副标题
            title = page_contents[1].get('title', '')
            subtitle = page_contents[1].get('subtitle', f"（执行时间：{time.strftime('%Y')}年{time.strftime('%m')}月）")
            
            # 获取第二页的注释
            second_page_notes = page_contents[1].get('notes', [])
            
            # 设置列宽（使用5列，与使用说明表格一致）
            col_widths = {
                'A': 25,  # 名称
                'B': 15,  # 序号
                'C': 35,  # 明细
                'D': 25,  # 计算关系
                'E': 25   # 数值
            }
            
            # 设置列宽
            for col, width in col_widths.items():
                worksheet.set_column(f'{col}:{col}', width)
            
            worksheet.set_default_row(25)  # 减小默认行高
            
            # 写入标题和副标题
            next_row = write_sheet_with_title(worksheet, title, subtitle, 5)  # 使用5列
            
            # 写入单位信息到右上角（第2行，E列）
            worksheet.write(1, 4, "单位：亿千瓦时、元/千瓦时", unit_format)  # 使用索引4代表E列
            
            # 写入数据并合并单元格
            current_content = None
            merge_start = next_row
            last_col = 0  # 记录最后一列的索引
            
            # 写入数据
            for row, (_, data_row) in enumerate(usage_df.iterrows(), start=next_row):
                for col, value in enumerate(data_row):
                    if pd.notnull(value):
                        last_col = max(last_col, col)  # 更新最后一列的索引
                    if col == 0:  # 第一列需要合并相同内容
                        if pd.notnull(value):
                            if current_content != value:
                                if current_content is not None:
                                    worksheet.merge_range(merge_start, col, row-1, col, current_content, text_format)
                                current_content = value
                                merge_start = row
                            worksheet.write(row, col, value, text_format)
                    else:
                        if pd.notnull(value):
                            worksheet.write(row, col, value, text_format)
            
            # 处理最后一个合并
            if current_content is not None:
                worksheet.merge_range(merge_start, 0, row, 0, current_content, text_format)
            
            # 写入第二页的注释（使用与上方表格相同的列数）
            note_start_row = row + 2
            actual_col_count = last_col + 1  # 实际使用的列数
            write_notes(worksheet, second_page_notes, note_start_row, actual_col_count, sum([col_widths[chr(65+i)] for i in range(actual_col_count)])/actual_col_count)
            
            # 冻结窗格
            worksheet.freeze_panes(next_row, 0)
                    # 保存文件
        writer.close()
        print(f"Excel文件已保存到: {file_path}")
        return file_path
    except Exception as e:
        print(f"保存Excel失败: {str(e)}")
        return None
def main():
    # 网页URL
    page_url = "https://95598.csg.cn/#/hn/serviceInquire/information/detail/?infoId=785a045dc4f547b6bf7c734b3a9fb75a"
    
    try:
        # 1. 获取PDF链接
        print("正在获取PDF链接...")
        pdf_url = get_pdf_url(page_url)
        if not pdf_url:
            return
            
        # 2. 下载PDF
        print("\n正在下载PDF文件...")
        pdf_path = download_pdf(pdf_url)
        if not pdf_path:
            return
        
        # 3. 提取内容
        print("\n正在提取数据...")
        page_contents = extract_table_from_pdf(pdf_path)
        if not page_contents:
            return
        
        # 4. 处理数据
        print("\n正在处理数据...")
        # 处理第一页数据
        df1, notes, calc_df = process_first_page(page_contents[0])
        if df1 is None:
            return
            
        # 处理第二页数据（使用说明）
        usage_df = None
        if len(page_contents) > 1:
            usage_df = process_second_page(page_contents[1])
        
        # 5. 保存Excel
        print("\n正在保存Excel文件...")
        # 从第一页提取文件名前缀
        title = page_contents[0].get('title', '电价表')
        company_name = title.split('代理购电价格表')[0] if '代理购电价格表' in title else ''
        filename_prefix = f"{company_name}电价表" if company_name else "电价表"
        
        # 生成文件名
        filename = f"{filename_prefix}_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_dir = "downloads"
        file_path = os.path.join(save_dir, filename)
        
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        excel_path = save_to_excel(df1, notes, calc_df, usage_df, page_contents, save_dir=save_dir)
        if excel_path:
            print("\n处理完成！")
            
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
if __name__ == "__main__":
    main() 
