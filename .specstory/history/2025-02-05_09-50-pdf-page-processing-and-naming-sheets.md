<!-- Generated by SpecStory -->

# PDF Page Processing and Naming Sheets (2025-02-05 17:50:25)

_**User**_

请在这个文件的代码上修改，识别的pdf页数，有些pdf文档页数有6页，则是确保最后一页是计算说明，其他页都是电价表,这部分sheet的命名可以是电价表1，电价表2，电价表3.

---

_**Assistant**_

我将修改代码以支持多个电价表页面，并确保最后一页是计算说明。我会修改 `write_to_excel` 函数来实现这个功能。

---

```diff
- 
  import os
  import time
          writer = pd.ExcelWriter(excel_file, engine='openpyxl')
          
-         # 处理第一页数据（主电价表）
-         print("处理第一页数据...")
-         if len(all_pages_data) > 0:
-             data, notes, calc_df = process_first_page(all_pages_data[0])
+         # 处理所有电价表页面（除最后一页）
+         print("处理电价表页面...")
+         for page_index in range(len(all_pages_data) - 1):
+             data, notes, calc_df = process_first_page(all_pages_data[page_index])
              if data is not None:
+                 # 为每个电价表页面创建不同的sheet名
+                 sheet_name = f'电价表{page_index + 1}' if page_index > 0 else '电价表'
+                 
                  # 写入主表数据
-                 data.to_excel(writer, sheet_name='电价表', index=False, startrow=3)
+                 data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)
                  
                  # 获取工作表对象
-                 worksheet = writer.sheets['电价表']
+                 worksheet = writer.sheets[sheet_name]
                  
                  # 设置标题
-                 title = all_pages_data[0].get('title', '')
-                 subtitle = all_pages_data[0].get('subtitle', '')
-                 unit = all_pages_data[0].get('unit', '')
+                 title = all_pages_data[page_index].get('title', '')
+                 subtitle = all_pages_data[page_index].get('subtitle', '')
+                 unit = all_pages_data[page_index].get('unit', '')
                  
                  # 写入标题并合并单元格
                  apply_cell_format(worksheet)
          
-         # 处理第二页数据（使用说明）
-         print("处理第二页数据...")
-         if len(all_pages_data) > 1:
-             data = process_second_page(all_pages_data[1])
-             if data is not None:
-                 data.to_excel(writer, sheet_name='使用说明', index=False, startrow=3)
-                 
-                 # 获取工作表对象
-                 worksheet = writer.sheets['使用说明']
-                 
-                 # 设置标题
-                 title = all_pages_data[1].get('title', '')
-                 subtitle = all_pages_data[1].get('subtitle', '')
-                 unit = all_pages_data[1].get('unit', '')
-                 
-                 # 写入标题并合并单元格
-                 max_col = worksheet.max_column
-                 worksheet['A1'] = title
-                 worksheet.merge_cells(f'A1:{openpyxl.utils.get_column_letter(max_col)}1')
-                 worksheet['A2'] = subtitle
-                 worksheet.merge_cells(f'A2:{openpyxl.utils.get_column_letter(max_col)}2')
-                 worksheet['A3'] = unit
-                 worksheet.merge_cells(f'A3:{openpyxl.utils.get_column_letter(max_col)}3')
-                 
-                 # 合并空白单元格，标题行（第4行）不合并
-                 header_row = 4
-                 merge_empty_cells(worksheet, header_row + 1, worksheet.max_row, 1, max_col, header_row=header_row)
-                 
-                 # 写入注释
-                 if 'notes' in all_pages_data[1]:
-                     notes = all_pages_data[1]['notes']
-                     if notes:
-                         add_notes_with_title(worksheet, notes, worksheet.max_row, max_col)
-                 
-                 # 应用单元格格式
-                 apply_cell_format(worksheet, header_row=header_row)
-         
-         # 处理第三页数据（计算说明）
-         print("处理第三页数据...")
-         if len(all_pages_data) > 2:
+         # 处理最后一页（计算说明）
+         print("处理计算说明页面...")
+         if len(all_pages_data) > 0:
+             last_page = all_pages_data[-1]
              # 创建计算说明工作表
              calc_sheet = writer.book.create_sheet('计算说明')
              
              # 设置标题
-             title = all_pages_data[2].get('title', '')
-             subtitle = all_pages_data[2].get('subtitle', '')
-             unit = all_pages_data[2].get('unit', '')
+             title = last_page.get('title', '')
+             subtitle = last_page.get('subtitle', '')
+             unit = last_page.get('unit', '')
              
              # 写入标题
              
              # 获取表格数据
-             if 'table' in all_pages_data[2]:
-                 table_data = all_pages_data[2]['table']
+             if 'table' in last_page:
+                 table_data = last_page['table']
                  # 写入表格数据
                  for i, row in enumerate(table_data, start=4):
                  
                  # 写入注释
-                 if 'notes' in all_pages_data[2]:
-                     notes = all_pages_data[2]['notes']
+                 if 'notes' in last_page:
+                     notes = last_page['notes']
                      if notes:
                          add_notes_with_title(calc_sheet, notes, len(table_data) + 4, max_col)
```

---



我已经修改了 `write_to_excel` 函数，主要做了以下改动：

1. 移除了原来固定处理三页的逻辑（电价表、使用说明、计算说明）
2. 改为动态处理所有页面：
   - 除最后一页外的所有页面都作为电价表处理
   - 为每个电价表页面创建不同的sheet名（电价表、电价表1、电价表2等）
   - 最后一页固定作为计算说明处理

现在这个函数可以：
- 处理任意页数的PDF文档
- 自动将除最后一页外的所有页面作为电价表处理，并分别命名为"电价表"、"电价表1"、"电价表2"等
- 将最后一页作为计算说明处理
- 保持了原有的格式处理、单元格合并等功能

这样修改后，当处理6页的PDF文档时：
- 前5页会分别生成名为"电价表"、"电价表1"、"电价表2"、"电价表3"、"电价表4"的sheet
- 最后一页（第6页）会生成名为"计算说明"的sheet

是否还需要其他调整？

---



---

