def write_to_excel(all_pages_data, project_name=None, is_yunnan=False):
    """将数据写入Excel文件
    
    Args:
        all_pages_data: 所有页面的数据
        project_name: 项目名称，用于文件命名
        is_yunnan: 是否使用云南省特殊处理
    """
    try:
        if not all_pages_data:
            print("没有数据需要保存")
            return None
            
        print(f"共有 {len(all_pages_data)} 页数据需要处理")
        
        # 创建downloads目录（如果不存在）
        if not os.path.exists('downloads'):
            os.makedirs('downloads')
        
        # 生成基础文件名
        base_filename = project_name if project_name else f'电价表_{datetime.now().strftime("%Y%m%d_%H%M%S")}'
        
        # 检查是否存在同名文件，并添加版本号
        version = 1
        while True:
            excel_file = os.path.join('downloads', f'{base_filename}V{version}.xlsx')
            if not os.path.exists(excel_file):
                break
            version += 1
        
        # 创建Excel写入器
        writer = pd.ExcelWriter(excel_file, engine='openpyxl')
        
        # 处理所有电价表页面（除最后一页）
        print("处理电价表页面...")
        for page_index in range(len(all_pages_data) - 1):
            data, notes, calc_df = process_first_page(all_pages_data[page_index])
            
            # 为每个电价表页面创建不同的sheet名
            sheet_name = f'电价表{page_index + 1}' if page_index > 0 else '电价表'
            
            # 如果有表格数据，写入主表数据
            if data is not None:
                if is_yunnan:
                    # 云南省特殊处理：调整表格格式
                    data = adjust_yunnan_table_format(data)
                data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)
            
            # 获取工作表对象
            worksheet = writer.sheets[sheet_name]
            
            # 设置标题
            title = all_pages_data[page_index].get('title', '')
            subtitle = all_pages_data[page_index].get('subtitle', '')
            unit = all_pages_data[page_index].get('unit', '')
            
            # 写入标题并合并单元格
            worksheet['A1'] = title
            worksheet.merge_cells(f'A1:{openpyxl.utils.get_column_letter(worksheet.max_column)}1')
            worksheet['A2'] = subtitle
            worksheet.merge_cells(f'A2:{openpyxl.utils.get_column_letter(worksheet.max_column)}2')
            worksheet['A3'] = unit
            worksheet.merge_cells(f'A3:{openpyxl.utils.get_column_letter(worksheet.max_column)}3')
            
            # 设置云南省特有的样式
            if is_yunnan:
                apply_yunnan_styles(worksheet)
            
            # 处理图片
            if 'images' in all_pages_data[page_index]:
                print(f"正在处理第 {page_index + 1} 页的图片...")
                images = all_pages_data[page_index]['images']
                
                # 计算图片插入的起始行
                start_row = worksheet.max_row + 2  # 在表格数据后留两行空白
                
                for i, img_info in enumerate(images):
                    try:
                        # 添加图片标题
                        title_row = start_row + i * 20  # 每张图片预留20行空间
                        worksheet.cell(row=title_row, column=1, value=f"图片 {i+1}")
                        
                        # 插入图片
                        img = openpyxl.drawing.image.Image(img_info['path'])
                        
                        # 设置图片大小（根据原始尺寸等比例缩放）
                        max_width = 800  # 最大宽度（像素）
                        max_height = 600  # 最大高度（像素）
                        width = img_info['width']
                        height = img_info['height']
                        
                        # 计算缩放比例
                        width_ratio = max_width / width if width > max_width else 1
                        height_ratio = max_height / height if height > max_height else 1
                        scale_ratio = min(width_ratio, height_ratio)
                        
                        # 应用缩放
                        img.width = width * scale_ratio
                        img.height = height * scale_ratio
                        
                        # 设置图片位置
                        img.anchor = f'A{title_row + 1}'
                        
                        # 添加图片到工作表
                        worksheet.add_image(img)
                        print(f"已插入图片 {i+1}")
                        
                    except Exception as e:
                        print(f"插入图片 {i+1} 失败: {str(e)}")
                        continue
            
            # 合并空白单元格，标题行（第4行）不合并
            if data is not None:
                header_row = 4
                if is_yunnan:
                    merge_yunnan_empty_cells(worksheet, header_row + 1, worksheet.max_row, 1, worksheet.max_column, header_row=header_row)
                else:
                    merge_empty_cells(worksheet, header_row + 1, worksheet.max_row, 1, worksheet.max_column, header_row=header_row)
            
            # 写入注释
            if notes:
                if is_yunnan:
                    add_yunnan_notes_with_title(worksheet, notes, worksheet.max_row, worksheet.max_column)
                else:
                    add_notes_with_title(worksheet, notes, worksheet.max_row, worksheet.max_column)
            
            # 应用单元格格式
            if is_yunnan:
                apply_yunnan_cell_format(worksheet)
            else:
                apply_cell_format(worksheet)
        
        # 处理最后一页（计算说明）
        print("处理计算说明页面...")
        if len(all_pages_data) > 0:
            last_page = all_pages_data[-1]
            # 创建计算说明工作表
            calc_sheet = writer.book.create_sheet('计算说明')
            
            # 设置标题
            title = last_page.get('title', '')
            subtitle = last_page.get('subtitle', '')
            unit = last_page.get('unit', '')
            
            # 写入标题
            calc_sheet['A1'] = title
            calc_sheet['A2'] = subtitle
            calc_sheet['A3'] = unit
            
            # 获取表格数据
            if 'table' in last_page:
                table_data = last_page['table']
                # 写入表格数据
                for i, row in enumerate(table_data, start=4):
                    for j, value in enumerate(row, start=1):
                        calc_sheet.cell(row=i, column=j, value=value)
                
                # 合并标题行
                max_col = max(len(row) for row in table_data)
                for row in range(1, 4):
                    calc_sheet.merge_cells(f'A{row}:{openpyxl.utils.get_column_letter(max_col)}{row}')
                
                # 合并空白单元格，标题行（第4行）不合并
                header_row = 4
                if is_yunnan:
                    merge_yunnan_empty_cells(calc_sheet, header_row + 1, calc_sheet.max_row, 1, max_col, header_row=header_row)
                else:
                    merge_empty_cells(calc_sheet, header_row + 1, calc_sheet.max_row, 1, max_col, header_row=header_row)
            
            # 处理图片
            if 'images' in last_page:
                print("正在处理计算说明页面的图片...")
                images = last_page['images']
                
                # 计算图片插入的起始行
                start_row = calc_sheet.max_row + 2  # 在表格数据后留两行空白
                
                for i, img_info in enumerate(images):
                    try:
                        # 添加图片标题
                        title_row = start_row + i * 20  # 每张图片预留20行空间
                        calc_sheet.cell(row=title_row, column=1, value=f"图片 {i+1}")
                        
                        # 插入图片
                        img = openpyxl.drawing.image.Image(img_info['path'])
                        
                        # 设置图片大小（根据原始尺寸等比例缩放）
                        max_width = 800  # 最大宽度（像素）
                        max_height = 600  # 最大高度（像素）
                        width = img_info['width']
                        height = img_info['height']
                        
                        # 计算缩放比例
                        width_ratio = max_width / width if width > max_width else 1
                        height_ratio = max_height / height if height > max_height else 1
                        scale_ratio = min(width_ratio, height_ratio)
                        
                        # 应用缩放
                        img.width = width * scale_ratio
                        img.height = height * scale_ratio
                        
                        # 设置图片位置
                        img.anchor = f'A{title_row + 1}'
                        
                        # 添加图片到工作表
                        calc_sheet.add_image(img)
                        print(f"已插入图片 {i+1}")
                        
                    except Exception as e:
                        print(f"插入图片 {i+1} 失败: {str(e)}")
                        continue
            
            # 写入注释
            if 'notes' in last_page:
                notes = last_page['notes']
                if notes:
                    if is_yunnan:
                        add_yunnan_notes_with_title(calc_sheet, notes, calc_sheet.max_row, calc_sheet.max_column)
                    else:
                        add_notes_with_title(calc_sheet, notes, calc_sheet.max_row, calc_sheet.max_column)
            
            # 应用单元格格式
            if is_yunnan:
                apply_yunnan_cell_format(calc_sheet)
            else:
                apply_cell_format(calc_sheet)
        
        # 保存Excel文件
        writer.close()
        
        print(f"Excel文件已保存到: {excel_file}")
        return excel_file
        
    except Exception as e:
        print(f"保存Excel文件时出错: {str(e)}")
        import traceback
        traceback.print_exc()
        return None 