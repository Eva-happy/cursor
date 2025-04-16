import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
import re

def create_templates(version=None):
    try:
        # 确保输出目录存在
        output_dir = r"D:\小工具\cursor\装机容量测算"
        print(f"正在创建目录: {output_dir}")
        os.makedirs(output_dir, exist_ok=True)
        print(f"目录创建/确认成功")
        
        # 确定版本号
        if version is None:
            # 自动检测现有文件的版本号并递增
            version = get_next_version(output_dir)
        
        version_str = f"V{version}"
        print(f"正在创建版本 {version_str} 的Excel模板文件...")

        # 1. 创建负载数据模板
        print("\n开始创建负载数据模板...")
        start_date = datetime(2024, 1, 1)
        dates = []
        hours = []
        loads = []

        for day in range(365):
            current_date = start_date + timedelta(days=day)
            for hour in range(24):
                dates.append(current_date.strftime('%Y-%m-%d'))
                hours.append(hour)
                # 模拟负载数据
                if 8 <= hour <= 18:  # 工作时间负载高一些
                    loads.append(np.random.normal(2000, 300))
                else:
                    loads.append(np.random.normal(1500, 200))

        load_df = pd.DataFrame({
            'date': dates,
            'hour': hours,
            'load': loads
        })

        load_file_path = os.path.join(output_dir, f'项目名称_load_data_{version_str}.xlsx')
        print(f"正在保存负载数据到: {load_file_path}")
        load_df.to_excel(load_file_path, index=False)
        print("负载数据保存成功")

        # 2. 创建电价配置模板
        print("\n开始创建电价配置模板...")
        # 时段配置 - 使用数字代表时段类型
        # 1=尖峰,2=高峰,3=平段,4=低谷,5=深谷
        period_columns = ['月份'] + [f'{i}-{i+1}' for i in range(24)]
        period_data = []

        for month in range(1, 13):
            row = [month]
            for hour in range(24):
                if 0 <= hour < 7:  # 0-7点为低谷时段
                    row.append(4)
                elif 7 <= hour < 10 or 15 <= hour < 18:  # 7-10点、15-18点为平时段
                    row.append(3)
                elif 10 <= hour < 15:  # 10-15点为尖峰时段
                    row.append(1)
                else:  # 18-24点为高峰时段
                    row.append(2)
            period_data.append(row)

        period_df = pd.DataFrame(period_data, columns=period_columns)

        # 电价配置
        price_data = {
            '月份': list(range(1, 13)),
            '时段类型': [1] * 12,  # 尖峰时段
            '电价': [1.2] * 12,
            '说明': ['尖峰时段'] * 12
        }
        price_df1 = pd.DataFrame(price_data)
        
        price_data = {
            '月份': list(range(1, 13)),
            '时段类型': [2] * 12,  # 高峰时段
            '电价': [0.9] * 12,
            '说明': ['高峰时段'] * 12
        }
        price_df2 = pd.DataFrame(price_data)
        
        price_data = {
            '月份': list(range(1, 13)),
            '时段类型': [3] * 12,  # 平段时段
            '电价': [0.6] * 12,
            '说明': ['平段时段'] * 12
        }
        price_df3 = pd.DataFrame(price_data)
        
        price_data = {
            '月份': list(range(1, 13)),
            '时段类型': [4] * 12,  # 低谷时段
            '电价': [0.3] * 12,
            '说明': ['低谷时段'] * 12
        }
        price_df4 = pd.DataFrame(price_data)
        
        price_data = {
            '月份': list(range(1, 13)),
            '时段类型': [5] * 12,  # 深谷时段
            '电价': [0.2] * 12,
            '说明': ['深谷时段'] * 12
        }
        price_df5 = pd.DataFrame(price_data)
        
        # 合并所有电价数据
        price_df = pd.concat([price_df1, price_df2, price_df3, price_df4, price_df5], ignore_index=True)

        # 创建一个Excel写入器，用于写入多个表
        price_file_path = os.path.join(output_dir, f'项目名称_price_config_{version_str}.xlsx')
        print(f"正在保存电价配置到: {price_file_path}")
        with pd.ExcelWriter(price_file_path) as writer:
            period_df.to_excel(writer, sheet_name='时段配置', index=False)
            price_df.to_excel(writer, sheet_name='电价配置', index=False)
        print("电价配置保存成功")

        # 也更新主程序使用的默认文件路径
        update_main_program_paths(load_file_path, price_file_path)

        print(f"\nExcel模板文件 {version_str} 创建成功！")
        print(f"- 负载数据文件: {load_file_path}")
        print(f"- 电价配置文件: {price_file_path}")
        
        # 验证文件是否确实创建
        print("\n验证文件创建结果：")
        if os.path.exists(load_file_path):
            print(f"- 负载数据文件创建成功，大小: {os.path.getsize(load_file_path)} 字节")
        else:
            print("- 警告：负载数据文件未找到！")
            
        if os.path.exists(price_file_path):
            print(f"- 电价配置文件创建成功，大小: {os.path.getsize(price_file_path)} 字节")
        else:
            print("- 警告：电价配置文件未找到！")

    except Exception as e:
        print(f"\n错误: {str(e)}")
        print(f"错误类型: {type(e).__name__}")
        import traceback
        print("\n详细错误信息:")
        print(traceback.format_exc())
        sys.exit(1)

def get_next_version(directory):
    """
    检查目录中已有的版本号，返回下一个版本号
    """
    # 查找所有负载数据文件
    max_version = 0
    
    try:
        for filename in os.listdir(directory):
            # 匹配带有版本号的文件
            match = re.search(r'项目名称_(?:load_data|price_config)_V(\d+)\.xlsx', filename)
            if match:
                version = int(match.group(1))
                max_version = max(max_version, version)
    except Exception as e:
        print(f"获取版本号时出错: {e}，使用默认版本V1")
        return 1
    
    # 返回下一个版本号
    return max_version + 1

def update_main_program_paths(load_file_path, price_file_path):
    """
    更新主程序中的文件路径
    """
    try:
        main_program = r"D:\小工具\cursor\energy_storage_optimization.py"
        if not os.path.exists(main_program):
            print(f"警告: 未找到主程序文件 {main_program}，跳过更新路径")
            return
        
        with open(main_program, 'r', encoding='utf-8') as file:
            content = file.read()
        
        # 更新文件路径
        content = re.sub(
            r'default_load_file = r".*?"', 
            f'default_load_file = r"{load_file_path.replace("\\", "\\\\")}"', 
            content
        )
        content = re.sub(
            r'default_price_config_file = r".*?"', 
            f'default_price_config_file = r"{price_file_path.replace("\\", "\\\\")}"', 
            content
        )
        
        with open(main_program, 'w', encoding='utf-8') as file:
            file.write(content)
        
        print(f"已更新主程序中的文件路径指向新版本")
    except Exception as e:
        print(f"更新主程序路径时出错: {e}")

if __name__ == "__main__":
    # 可以手动指定版本号，例如 create_templates(3) 创建V3版本
    # 如果不指定，则自动检测并递增
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        version = int(sys.argv[1])
        create_templates(version)
    else:
        create_templates() 