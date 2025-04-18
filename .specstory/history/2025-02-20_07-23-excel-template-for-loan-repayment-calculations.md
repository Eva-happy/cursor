<!-- Generated by SpecStory -->

# Excel Template for Loan Repayment Calculations (2025-02-20 15:23:41)

_**User**_

请提供一份excel含有公式模版，分别计算等额本金，等额本息，先息后本，以及动态还款（动态还款的模式，比如期初借入本金，每年还利息，然而本金偿还要看当年年末是否有剩余的现金流，如果有剩余的现金流则拿去还款，第二年的利息就是基于期初借款余额乘以利率，以此类推。你可以帮我搭建个模版吗？列明各个参数值，和模版计算公式要有），分别在不同sheet列出以上模版。并且列出本金 利率 年限 每年现金余额等关键参数，计算公式要链接这些参数

---

_**Assistant**_

我来帮您创建一个包含多种还款方式计算的Excel模板。我会使用Python和openpyxl库来生成这个Excel文件。

首先，让我创建一个Python脚本来生成所需的Excel模板。

---

```diff
- 
+ import openpyxl
+ from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
+ from openpyxl.utils import get_column_letter
+ from openpyxl.worksheet.datavalidation import DataValidation
+ 
+ def create_loan_template():
+     wb = openpyxl.Workbook()
+     
+     # 创建参数表
+     params = wb.active
+     params.title = "参数设置"
+     setup_params_sheet(params)
+     
+     # 创建等额本息表
+     equal_payment = wb.create_sheet("等额本息")
+     setup_equal_payment_sheet(equal_payment)
+     
+     # 创建等额本金表
+     equal_principal = wb.create_sheet("等额本金")
+     setup_equal_principal_sheet(equal_principal)
+     
+     # 创建先息后本表
+     interest_first = wb.create_sheet("先息后本")
+     setup_interest_first_sheet(interest_first)
+     
+     # 创建动态还款表
+     dynamic = wb.create_sheet("动态还款")
+     setup_dynamic_sheet(dynamic)
+     
+     wb.save('loan_calculation_template.xlsx')
+ 
+ def setup_params_sheet(ws):
+     # 设置标题
+     ws['A1'] = '贷款计算基础参数'
+     ws['A1'].font = Font(bold=True, size=14)
+     ws.merge_cells('A1:D1')
+     
+     # 设置参数
+     params = [
+         ('贷款本金', 1000000),
+         ('年利率', 0.045),
+         ('贷款年限', 5),
+         ('还款频率', '年'),
+     ]
+     
+     for i, (param, value) in enumerate(params, start=2):
+         ws[f'A{i}'] = param
+         ws[f'B{i}'] = value
+         ws[f'A{i}'].font = Font(bold=True)
+     
+     # 设置现金流表
+     ws['E2'] = '年份'
+     ws['F2'] = '可用现金流'
+     ws['E2:F2'].font = Font(bold=True)
+     
+     for i in range(5):
+         ws[f'E{i+3}'] = f'第{i+1}年'
+         ws[f'F{i+3}'] = 300000  # 示例现金流
+ 
+ def setup_equal_payment_sheet(ws):
+     # 设置标题和公式说明
+     ws['A1'] = '等额本息还款计划'
+     ws['A1'].font = Font(bold=True, size=14)
+     ws.merge_cells('A1:G1')
+     
+     # 设置表头
+     headers = ['期数', '期初余额', '还款额', '本金', '利息', '期末余额', '累计还款']
+     for i, header in enumerate(headers, start=1):
+         ws[f'{get_column_letter(i)}3'] = header
+         ws[f'{get_column_letter(i)}3'].font = Font(bold=True)
+     
+     # 设置公式
+     for i in range(1, 6):  # 假设5年
+         row = i + 3
+         ws[f'A{row}'] = i  # 期数
+         if i == 1:
+             ws[f'B{row}'] = '=参数设置!B2'  # 期初余额等于贷款本金
+         else:
+             ws[f'B{row}'] = f'=G{row-1}'  # 期初余额等于上期期末余额
+         
+         # PMT公式计算还款额
+         ws[f'C{row}'] = '=PMT(参数设置!B3,参数设置!B4,-参数设置!B2)'
+         ws[f'D{row}'] = f'=B{row}*参数设置!B3'  # 本金
+         ws[f'E{row}'] = f'=C{row}-D{row}'  # 利息
+         ws[f'F{row}'] = f'=B{row}-D{row}'  # 期末余额
+         ws[f'G{row}'] = f'=IF(ROW()=4,C4,G{row-1}+C{row})'  # 累计还款
+ 
+ def setup_equal_principal_sheet(ws):
+     # 设置标题和公式说明
+     ws['A1'] = '等额本金还款计划'
+     ws['A1'].font = Font(bold=True, size=14)
+     ws.merge_cells('A1:G1')
+     
+     # 设置表头
+     headers = ['期数', '期初余额', '还款额', '本金', '利息', '期末余额', '累计还款']
+     for i, header in enumerate(headers, start=1):
+         ws[f'{get_column_letter(i)}3'] = header
+         ws[f'{get_column_letter(i)}3'].font = Font(bold=True)
+     
+     # 设置公式
+     for i in range(1, 6):  # 假设5年
+         row = i + 3
+         ws[f'A{row}'] = i  # 期数
+         if i == 1:
+             ws[f'B{row}'] = '=参数设置!B2'  # 期初余额等于贷款本金
+         else:
+             ws[f'B{row}'] = f'=F{row-1}'  # 期初余额等于上期期末余额
+         
+         ws[f'D{row}'] = '=参数设置!B2/参数设置!B4'  # 每期本金金额
+         ws[f'E{row}'] = f'=B{row}*参数设置!B3'  # 利息
+         ws[f'C{row}'] = f'=D{row}+E{row}'  # 还款额
+         ws[f'F{row}'] = f'=B{row}-D{row}'  # 期末余额
+         ws[f'G{row}'] = f'=IF(ROW()=4,C4,G{row-1}+C{row})'  # 累计还款
+ 
+ def setup_interest_first_sheet(ws):
+     # 设置标题和公式说明
+     ws['A1'] = '先息后本还款计划'
+     ws['A1'].font = Font(bold=True, size=14)
+     ws.merge_cells('A1:G1')
+     
+     # 设置表头
+     headers = ['期数', '期初余额', '还款额', '本金', '利息', '期末余额', '累计还款']
+     for i, header in enumerate(headers, start=1):
+         ws[f'{get_column_letter(i)}3'] = header
+         ws[f'{get_column_letter(i)}3'].font = Font(bold=True)
+     
+     # 设置公式
+     for i in range(1, 6):  # 假设5年
+         row = i + 3
+         ws[f'A{row}'] = i  # 期数
+         if i == 1:
+             ws[f'B{row}'] = '=参数设置!B2'  # 期初余额等于贷款本金
+         else:
+             ws[f'B{row}'] = f'=F{row-1}'  # 期初余额等于上期期末余额
+         
+         ws[f'E{row}'] = f'=B{row}*参数设置!B3'  # 利息
+         ws[f'D{row}'] = f'=IF(A{row}=参数设置!B4,B{row},0)'  # 本金（最后一期还清）
+         ws[f'C{row}'] = f'=D{row}+E{row}'  # 还款额
+         ws[f'F{row}'] = f'=B{row}-D{row}'  # 期末余额
+         ws[f'G{row}'] = f'=IF(ROW()=4,C4,G{row-1}+C{row})'  # 累计还款
+ 
+ def setup_dynamic_sheet(ws):
+     # 设置标题和公式说明
+     ws['A1'] = '动态还款计划'
+     ws['A1'].font = Font(bold=True, size=14)
+     ws.merge_cells('A1:I1')
+     
+     # 设置表头
+     headers = ['期数', '期初余额', '可用现金流', '还款额', '本金', '利息', '期末余额', '累计还款', '现金流结余']
+     for i, header in enumerate(headers, start=1):
+         ws[f'{get_column_letter(i)}3'] = header
+         ws[f'{get_column_letter(i)}3'].font = Font(bold=True)
+     
+     # 设置公式
+     for i in range(1, 6):  # 假设5年
+         row = i + 3
+         ws[f'A{row}'] = i  # 期数
+         if i == 1:
+             ws[f'B{row}'] = '=参数设置!B2'  # 期初余额等于贷款本金
+         else:
+             ws[f'B{row}'] = f'=G{row-1}'  # 期初余额等于上期期末余额
+         
+         ws[f'C{row}'] = f'=参数设置!F{i+2}'  # 可用现金流
+         ws[f'F{row}'] = f'=B{row}*参数设置!B3'  # 利息
+         ws[f'E{row}'] = f'=MIN(B{row},MAX(0,C{row}-F{row}))'  # 本金还款
+         ws[f'D{row}'] = f'=E{row}+F{row}'  # 还款额
+         ws[f'G{row}'] = f'=B{row}-E{row}'  # 期末余额
+         ws[f'H{row}'] = f'=IF(ROW()=4,D4,H{row-1}+D{row})'  # 累计还款
+         ws[f'I{row}'] = f'=C{row}-D{row}'  # 现金流结余
+ 
+ if __name__ == "__main__":
+     create_loan_template() 
```

---



现在让我们创建一个requirements.txt文件来指定所需的依赖：

---

```diff
- 
+ openpyxl==3.1.2 
```

---



让我们创建一个README文件来说明如何使用这个模板：

---

```diff
- 
+ # 贷款计算Excel模板生成器
+ 
+ 这个项目可以生成一个包含多种还款方式计算的Excel模板。
+ 
+ ## 功能特点
+ 
+ 该模板包含以下还款方式的计算：
+ 1. 等额本息
+ 2. 等额本金
+ 3. 先息后本
+ 4. 动态还款（根据现金流情况）
+ 
+ ## 安装依赖
+ 
+ ```bash
+ pip install -r requirements.txt
+ ```
+ 
+ ## 使用方法
+ 
+ 1. 运行Python脚本生成模板：
+ ```bash
+ python create_loan_template.py
+ ```
+ 
+ 2. 运行后会生成`loan_calculation_template.xlsx`文件，包含以下sheet页：
+    - 参数设置：设置基本参数（贷款本金、利率、期限等）
+    - 等额本息：等额本息还款计划
+    - 等额本金：等额本金还款计划
+    - 先息后本：先息后本还款计划
+    - 动态还款：根据现金流情况的动态还款计划
+ 
+ 3. 在"参数设置"页面中修改基本参数：
+    - 贷款本金
+    - 年利率
+    - 贷款年限
+    - 还款频率
+    - 每年可用现金流（用于动态还款计算）
+ 
+ ## 注意事项
+ 
+ - 所有计算公式都已设置好，修改参数后会自动更新计算结果
+ - 动态还款模式下，系统会根据可用现金流自动计算最大可还款金额
+ - 所有金额单位为元
+ - 利率请输入小数形式（如4.5%请输入0.045） 
```

---



现在让我们运行这个脚本来生成Excel模板：

---

```bash
python create_loan_template.py
```

```
PS D:\小工具\cursor> python create_loan_template.py
PS D:\小工具\cursor> python create_loan_template.py




























Traceback (most recent call last):
  File "D:\小工具\cursor\create_loan_template.py", line 172, in <module>
    create_loan_template() 
    ^^^^^^^^^^^^^^^^^^^^^^
  File "D:\小工具\cursor\create_loan_template.py", line 12, in create_loan_template
    setup_params_sheet(params)
  File "D:\小工具\cursor\create_loan_template.py", line 54, in setup_params_sheet
    ws['E2:F2'].font = Font(bold=True)
    ^^^^^^^^^^^^^^^^
AttributeError: 'tuple' object has no attribute 'font'
PS D:\小工具\cursor>
```

---



看起来代码中有一个小错误，让我修复一下：

---

```diff
      ws['E2'] = '年份'
      ws['F2'] = '可用现金流'
-     ws['E2:F2'].font = Font(bold=True)
+     ws['E2'].font = Font(bold=True)
+     ws['F2'].font = Font(bold=True)
      
      for i in range(5):
```

---



现在让我们再次运行脚本：

---

```bash
python create_loan_template.py
```

```
PS D:\小工具\cursor>
```

---



太好了！Excel模板已经成功生成。我来为您解释一下这个模板的主要特点：

1. **参数设置页面**：
   - 包含基本参数：贷款本金（默认100万）、年利率（默认4.5%）、贷款年限（默认5年）
   - 包含每年可用现金流设置（默认每年30万，用于动态还款计算）

2. **等额本息页面**：
   - 使用Excel的PMT函数计算每期还款额
   - 显示每期的本金、利息分配
   - 计算累计还款金额
   - 所有计算都链接到参数设置页面

3. **等额本金页面**：
   - 每期还款本金相同
   - 利息随本金减少而减少
   - 总还款额随期数递减
   - 所有计算都链接到参数设置页面

4. **先息后本页面**：
   - 每期只还利息
   - 最后一期还清所有本金
   - 所有计算都链接到参数设置页面

5. **动态还款页面**：
   - 根据每年的可用现金流情况计算还款额
   - 优先支付利息
   - 剩余现金用于还本
   - 显示每期的现金流结余
   - 所有计算都链接到参数设置页面

使用方法：
1. 打开生成的`loan_calculation_template.xlsx`文件
2. 在"参数设置"页面修改您的实际参数
3. 其他页面的计算结果会自动更新
4. 对于动态还款，您可以在参数设置页面修改每年的可用现金流

需要注意的是：
- 所有金额单位为元
- 利率请输入小数形式（如4.5%请输入0.045）
- 可以随时修改参数，所有计算都会自动更新
- 动态还款模式会自动计算最优的还款方案，确保优先支付利息

您可以查看生成的Excel文件，如果有任何具体的问题或需要调整的地方，请随时告诉我。

---



---

