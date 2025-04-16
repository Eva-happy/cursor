import requests
import os
import pdfplumber
import pandas as pd
from urllib.parse import unquote
import time

def download_pdf(url, save_dir="downloads"):
    """下载PDF文件"""
    try:
        # 创建下载目录
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 获取PDF文件
        response = requests.get(url)
        if response.status_code == 200:
            # 从URL中提取文件名
            filename = f"电价表_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
            file_path = os.path.join(save_dir, filename)
            
            # 保存PDF文件
            with open(file_path, 'wb') as f:
                f.write(response.content)
            print(f"PDF文件已保存到: {file_path}")
            return file_path
        else:
            print(f"下载失败，状态码: {response.status_code}")
            return None
    except Exception as e:
        print(f"下载出错: {str(e)}")
        return None

def extract_table_from_pdf(pdf_path):
    """从PDF中提取表格"""
    try:
        tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # 提取表格
                table = page.extract_table()
                if table:
                    tables.extend(table)
        return tables
    except Exception as e:
        print(f"提取表格失败: {str(e)}")
        return None

def process_price_table(table_data):
    """处理表格数据"""
    try:
        # 定义列名
        columns = [
            "用电分类", "电压等级", "电度用电价格",
            "代理购电价格", "电度输配电价", "系统运行费用",
            "上网环节线损电价", "政府性基金及附加",
            "高峰时段", "平时段", "低谷时段",
            "需量电价", "容量电价"
        ]
        
        # 创建DataFrame
        df = pd.DataFrame(table_data[1:], columns=columns)
        
        # 清理数据
        df = df.replace('', None)
        df = df.fillna(method='ffill')  # 向下填充合并单元格的值
        
        # 处理数值列
        numeric_columns = [
            "电度用电价格", "代理购电价格", "电度输配电价", 
            "系统运行费用", "上网环节线损电价", "政府性基金及附加",
            "高峰时段", "平时段", "低谷时段", 
            "需量电价", "容量电价"
        ]
        
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        return df
    except Exception as e:
        print(f"处理表格数据失败: {str(e)}")
        return None

def save_to_excel(df, save_dir="downloads"):
    """保存为Excel文件"""
    try:
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)
        
        # 生成文件名
        filename = f"电价表_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = os.path.join(save_dir, filename)
        
        # 保存Excel
        df.to_excel(file_path, index=False, encoding='utf-8')
        print(f"Excel文件已保存到: {file_path}")
        return file_path
    except Exception as e:
        print(f"保存Excel失败: {str(e)}")
        return None

def main():
    # PDF文件URL
    pdf_url = "https://95598.csg.cn/#/hn/serviceInquire/information/detail/?infoId=785a045dc4f547b6bf7c734b3a9fb75a"
    
    try:
        # 1. 下载PDF
        print("正在下载PDF文件...")
        pdf_path = download_pdf(pdf_url)
        if not pdf_path:
            return
        
        # 2. 提取表格
        print("\n正在提取表格数据...")
        table_data = extract_table_from_pdf(pdf_path)
        if not table_data:
            return
        
        # 3. 处理数据
        print("\n正在处理表格数据...")
        df = process_price_table(table_data)
        if df is None:
            return
        
        # 4. 保存Excel
        print("\n正在保存Excel文件...")
        excel_path = save_to_excel(df)
        if excel_path:
            print("\n处理完成！")
            
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")

if __name__ == "__main__":
    main() 