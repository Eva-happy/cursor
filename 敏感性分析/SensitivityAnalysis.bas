' 计算单一参数的敏感性分析结果
Public Sub CalculateSingleParameterAnalysis(ByVal paramName As String)
    Dim ws As Worksheet
    Dim baseSheet As Worksheet
    Dim i As Long
    Dim baseValue As Double
    Dim row As Long
    Dim lastRow As Long
    Dim startRow As Long
    
    Set ws = ThisWorkbook.Sheets("单一参数敏感性分析")
    Set baseSheet = ThisWorkbook.Sheets("基础参数及输出结果表")
    
    ' 检查是否是第一次选择参数（通过检查E3单元格是否为空）
    If IsEmpty(ws.Range("E3").Value) Then
        ' 第一次选择参数，只填充数据
        startRow = 3
        
        ' 设置参数名称
        ws.Range("E3").Value = paramName
        
        ' 获取基准值和设置公式
        Select Case paramName
            Case "发电小时数（单位：小时）"
                baseValue = baseSheet.Range("B23").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$B$23"
            Case "电价（单位：元/kWh）"
                baseValue = baseSheet.Range("B42").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$B$42"
            Case "消纳率（比率）"
                baseValue = baseSheet.Range("B37").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$B$37"
            Case "初始总投资（万元）"
                baseValue = baseSheet.Range("F7").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$F$7"
            Case "消交流侧装机容量（备案容量）"
                baseValue = baseSheet.Range("B12").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$B$12"
            Case "技改费（元/W）"
                baseValue = baseSheet.Range("F25").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$F$25"
            Case "股权资本金占比"
                baseValue = baseSheet.Range("J23").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$J$23"
            Case "还款年限（年）"
                baseValue = baseSheet.Range("J26").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$J$26"
            Case "利率"
                baseValue = baseSheet.Range("J27").Value
                ws.Range("E4").Formula = "=基础参数及输出结果表!$J$27"
        End Select
        
        ' 设置其他原值公式
        ws.Range("E5").Formula = "=基础参数及输出结果表!$N$8"  ' 原全投资IRR
        ws.Range("E6").Formula = "=基础参数及输出结果表!$N$10" ' 原资本金IRR
        ws.Range("E7").Formula = "=基础参数及输出结果表!$N$9"  ' 原全投资回收期
        ws.Range("E8").Formula = "=基础参数及输出结果表!$N$11" ' 原资本金回收期
        
    Else
        ' 不是第一次选择参数，在原表格下方生成新表格
        ' 找到最后一个表格的位置
        On Error Resume Next
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        
        ' 找到最后一个"参数名称"单元格
        For i = lastRow To 1 Step -1
            If ws.Cells(i, 2).Value = "参数名称" Then
                startRow = i + 13  ' 参数名称行 + 7行表格结构 + 6行输入区域
                Exit For
            End If
        Next i
        
        ' 如果没找到，从第3行开始
        If startRow = 0 Then
            startRow = 3
        End If
        
        ' 在上一个表格后空两行开始新表格
        startRow = startRow + 2
        
        ' 设置表格结构
        With ws
            ' 合并单元格
            .Range("B" & startRow & ":D" & startRow).Merge
            .Range("B" & (startRow + 1) & ":D" & (startRow + 1)).Merge
            .Range("B" & (startRow + 2) & ":D" & (startRow + 2)).Merge
            .Range("B" & (startRow + 3) & ":D" & (startRow + 3)).Merge
            .Range("B" & (startRow + 4) & ":D" & (startRow + 4)).Merge
            .Range("B" & (startRow + 5) & ":D" & (startRow + 5)).Merge
            .Range("E" & startRow & ":G" & startRow).Merge
            .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Merge
            .Range("E" & (startRow + 2) & ":G" & (startRow + 2)).Merge
            .Range("E" & (startRow + 3) & ":G" & (startRow + 3)).Merge
            .Range("E" & (startRow + 4) & ":G" & (startRow + 4)).Merge
            .Range("E" & (startRow + 5) & ":G" & (startRow + 5)).Merge
            .Range("B" & (startRow + 6) & ":G" & (startRow + 6)).Merge
            
            ' 设置标题和参数名称
            .Range("B" & startRow & ":D" & startRow).Value = "参数名称"
            .Range("E" & startRow & ":G" & startRow).Value = paramName
            
            ' 设置标题
            .Range("B" & (startRow + 1) & ":D" & (startRow + 1)).Value = "参数原值"
            .Range("B" & (startRow + 2) & ":D" & (startRow + 2)).Value = "原全投资IRR"
            .Range("B" & (startRow + 3) & ":D" & (startRow + 3)).Value = "原资本金IRR"
            .Range("B" & (startRow + 4) & ":D" & (startRow + 4)).Value = "原全投资回收期（年）"
            .Range("B" & (startRow + 5) & ":D" & (startRow + 5)).Value = "原资本金回收期（年）"
            .Range("B" & (startRow + 6) & ":G" & (startRow + 6)).Value = "结果分析"
            
            ' 设置原值和公式
            Select Case paramName
                Case "发电小时数（单位：小时）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$B$23"
                Case "电价（单位：元/kWh）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$B$42"
                Case "消纳率（比率）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$B$37"
                Case "初始总投资（万元）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$F$7"
                Case "消交流侧装机容量（备案容量）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$B$12"
                Case "技改费（元/W）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$F$25"
                Case "股权资本金占比"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$J$23"
                Case "还款年限（年）"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$J$26"
                Case "利率"
                    .Range("E" & (startRow + 1) & ":G" & (startRow + 1)).Formula = "=基础参数及输出结果表!$J$27"
            End Select
            
            ' 设置其他原值公式
            .Range("E" & (startRow + 2) & ":G" & (startRow + 2)).Formula = "=基础参数及输出结果表!$N$8"  ' 原全投资IRR
            .Range("E" & (startRow + 3) & ":G" & (startRow + 3)).Formula = "=基础参数及输出结果表!$N$10" ' 原资本金IRR
            .Range("E" & (startRow + 4) & ":G" & (startRow + 4)).Formula = "=基础参数及输出结果表!$N$9"  ' 原全投资回收期
            .Range("E" & (startRow + 5) & ":G" & (startRow + 5)).Formula = "=基础参数及输出结果表!$N$11" ' 原资本金回收期
            
            ' 设置表头
            .Range("B" & (startRow + 7) & ":G" & (startRow + 7)) = Array("变动后的值", "变动率", "变动后全投资IRR", "变动后资本金IRR", "变动后全投资回收期", "变动后资本金回收期")
            
            ' 预设6行用于输入和计算
            For row = startRow + 8 To startRow + 13
                ' 设置变动率公式
                .Cells(row, 3).Formula = "=IF(B" & row & "<>"""",(" & "B" & row & "-$E$" & (startRow + 1) & ")/$E$" & (startRow + 1) & ","""")"
                
                ' 设置变动后全投资IRR公式
                .Cells(row, 4).Formula = "=IF(B" & row & "<>"""",IFERROR(CalculateParameterValue(""" & paramName & """,B" & row & ",""全投资IRR""),""""),"""")"
                
                ' 设置变动后资本金IRR公式
                .Cells(row, 5).Formula = "=IF(B" & row & "<>"""",IFERROR(CalculateParameterValue(""" & paramName & """,B" & row & ",""资本金IRR""),""""),"""")"
                
                ' 设置变动后全投资回收期公式
                .Cells(row, 6).Formula = "=IF(B" & row & "<>"""",IFERROR(CalculateParameterValue(""" & paramName & """,B" & row & ",""全投回收期""),""""),"""")"
                
                ' 设置变动后资本金回收期公式
                .Cells(row, 7).Formula = "=IF(B" & row & "<>"""",IFERROR(CalculateParameterValue(""" & paramName & """,B" & row & ",""资本金回收期""),""""),"""")"
            Next row
        End With
        
        ' 设置格式
        FormatParameterTable ws, startRow
    End If
End Sub

' 新增函数：格式化单个参数表格
Private Sub FormatParameterTable(ByVal ws As Worksheet, ByVal startRow As Long)
    With ws
        ' 设置内部细边框
        .Range("B" & startRow & ":G" & (startRow + 13)).Borders.LineStyle = xlContinuous
        .Range("B" & startRow & ":G" & (startRow + 13)).Borders.Weight = xlThin
        
        ' 设置外框粗边框
        With .Range("B" & startRow & ":G" & (startRow + 13)).Borders
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeLeft).Weight = xlMedium
            
            .Item(xlEdgeTop).LineStyle = xlContinuous
            .Item(xlEdgeTop).Weight = xlMedium
            
            .Item(xlEdgeBottom).LineStyle = xlContinuous
            .Item(xlEdgeBottom).Weight = xlMedium
            
            .Item(xlEdgeRight).LineStyle = xlContinuous
            .Item(xlEdgeRight).Weight = xlMedium
        End With
        
        ' 设置颜色填充
        ' 参数名称到原资本金回收期的行填充灰色
        .Range("B" & startRow & ":G" & (startRow + 5)).Interior.Color = RGB(217, 217, 217)
        
        ' 结果分析行填充蓝色
        .Range("B" & (startRow + 6) & ":G" & (startRow + 6)).Interior.Color = RGB(189, 215, 238)
        
        ' 变动后的值行（包括表头）填充浅绿色
        .Range("B" & (startRow + 7) & ":G" & (startRow + 13)).Interior.Color = RGB(226, 239, 218)
        
        ' 设置对齐方式
        .Range("B" & startRow & ":G" & (startRow + 13)).HorizontalAlignment = xlCenter
        .Range("B" & startRow & ":G" & (startRow + 13)).VerticalAlignment = xlCenter
        
        ' 设置数值格式
        .Range("B" & (startRow + 8) & ":B" & (startRow + 13)).NumberFormat = "0.00"        ' 变动后的值：保留两位小数
        .Range("C" & (startRow + 8) & ":E" & (startRow + 13)).NumberFormat = "0.00%"       ' 变动率和IRR：百分比格式
        .Range("F" & (startRow + 8) & ":G" & (startRow + 13)).NumberFormat = "0.00"        ' 回收期：保留两位小数
        .Range("B" & (startRow + 2) & ":D" & (startRow + 3)).NumberFormat = "0.00%"        ' 原IRR值：百分比格式
        .Range("B" & (startRow + 4) & ":D" & (startRow + 5)).NumberFormat = "0.00"         ' 原回收期：保留两位小数
        
        ' 设置标题格式
        .Range("B" & startRow & ":G" & startRow).Font.Bold = True
        .Range("B" & (startRow + 6) & ":G" & (startRow + 6)).Font.Bold = True
        .Range("B" & (startRow + 7) & ":G" & (startRow + 7)).Font.Bold = True
        
        ' 调整列宽
        .Columns("B:G").AutoFit
    End With
End Sub

' 计算新的IRR值
Public Function CalculateParameterValue(ByVal paramType As String, ByVal newValue As Double, ByVal targetType As String) As Double
    Dim baseSheet As Worksheet
    Dim calcSheet As Worksheet
    Dim baseValue As Double
    Dim oldCalculation As XlCalculation
    
    ' 保存当前计算模式
    oldCalculation = Application.calculation
    
    ' 设置手动计算模式
    Application.calculation = xlCalculationManual
    
    Set baseSheet = ThisWorkbook.Sheets("基础参数及输出结果表")
    Set calcSheet = ThisWorkbook.Sheets("光伏收益测算表")
    
    ' 保存原始值
    Select Case paramType
        Case "发电小时数（单位：小时）", "发电小时数"
            baseValue = baseSheet.Range("B23").Value
            baseSheet.Range("B23").Value = newValue
        Case "电价（单位：元/kWh）", "电价"
            baseValue = baseSheet.Range("B42").Value
            baseSheet.Range("B42").Value = newValue
        Case "消纳率（比率）", "消纳率"
            baseValue = baseSheet.Range("B37").Value
            baseSheet.Range("B37").Value = newValue
        Case "初始总投资（万元）", "初始总投资"
            baseValue = baseSheet.Range("F7").Value
            baseSheet.Range("F7").Value = newValue
        Case "消交流侧装机容量（备案容量）", "消交流侧装机容量"
            baseValue = baseSheet.Range("B12").Value
            baseSheet.Range("B12").Value = newValue
        Case "技改费（元/W）"
            baseValue = baseSheet.Range("F25").Value
            baseSheet.Range("F25").Value = newValue
        Case "股权资本金占比"
            baseValue = baseSheet.Range("J23").Value
            baseSheet.Range("J23").Value = newValue
        Case "还款年限（年）"
            baseValue = baseSheet.Range("J26").Value
            baseSheet.Range("J26").Value = newValue
        Case "利率"
            baseValue = baseSheet.Range("J27").Value
            baseSheet.Range("J27").Value = newValue
    End Select
    
    ' 计算
    calcSheet.Calculate
    baseSheet.Calculate
    
    ' 获取结果
    Select Case targetType
        Case "全投资IRR"
            CalculateParameterValue = baseSheet.Range("N8").Value
        Case "资本金IRR"
            CalculateParameterValue = baseSheet.Range("N10").Value
        Case "全投回收期"
            CalculateParameterValue = baseSheet.Range("N9").Value
        Case "资本金回收期"
            CalculateParameterValue = baseSheet.Range("N11").Value
    End Select
    
    ' 恢复原始值
    Select Case paramType
        Case "发电小时数（单位：小时）", "发电小时数"
            baseSheet.Range("B23").Value = baseValue
        Case "电价（单位：元/kWh）", "电价"
            baseSheet.Range("B42").Value = baseValue
        Case "消纳率（比率）", "消纳率"
            baseSheet.Range("B37").Value = baseValue
        Case "初始总投资（万元）", "初始总投资"
            baseSheet.Range("F7").Value = baseValue
        Case "消交流侧装机容量（备案容量）", "消交流侧装机容量"
            baseSheet.Range("B12").Value = baseValue
        Case "技改费（元/W）"
            baseSheet.Range("F25").Value = baseValue
        Case "股权资本金占比"
            baseSheet.Range("J23").Value = baseValue
        Case "还款年限（年）"
            baseSheet.Range("J26").Value = baseValue
        Case "利率"
            baseSheet.Range("J27").Value = baseValue
    End Select
    
    ' 重新计算以恢复原始状态
    calcSheet.Calculate
    baseSheet.Calculate
    
    ' 恢复计算模式
    Application.calculation = oldCalculation
End Function

' 显示参数选择窗口
Public Sub ShowParameterSelector()
    ' 显示参数选择窗口
    ParameterSelectorForm.Show
End Sub


Sub AddFormulas(ByVal ws As Worksheet)
    Dim i As Long, j As Long
    Dim 参数类型 As String
    Dim 变化率 As String
    
    With ws
        ' 引用基准值
        .Range("B4").Formula = "='基础参数及输出结果表'!B23"  ' 发电小时数
        .Range("B5").Formula = "='基础参数及输出结果表'!B42"  ' 电价
        .Range("B6").Formula = "='基础参数及输出结果表'!B37"  ' 消纳率
        .Range("B7").Formula = "='基础参数及输出结果表'!F7"  ' 初始总投资
        .Range("B8").Formula = "='基础参数及输出结果表'!B12"  ' 消交流侧装机容量
        .Range("B9").Formula = "='基础参数及输出结果表'!F25"  ' 技改费
        .Range("B10").Formula = "='基础参数及输出结果表'!J23"  ' 股权资本金占比
        .Range("B11").Formula = "='基础参数及输出结果表'!J26"  ' 还款年限
        .Range("B12").Formula = "='基础参数及输出结果表'!J27"  ' 利率
        ' 全投资IRR计算
        For i = 11 To 19
            For j = 2 To 8
                Select Case i
                    Case 11: 参数类型 = "发电小时数"
                    Case 12: 参数类型 = "电价"
                    Case 13: 参数类型 = "消纳率"
                    Case 14: 参数类型 = "初始总投资"
                    Case 15: 参数类型 = "消交流侧装机容量"
                    Case 16: 参数类型 = "技改费"
                    Case 17: 参数类型 = "股权资本金占比"
                    Case 18: 参数类型 = "还款年限"
                    Case 19: 参数类型 = "利率"
                End Select
                
                变化率 = Format((j - 4) * 0.05, "0.00")  ' 转换为字符串格式
                
                .Cells(i, j).Formula = "=CalculateParameterValue(""" & 参数类型 & """, " & 变化率 & ", ""全投资"")"
            Next j
        Next i
        
        ' 资本金IRR计算
        For i = 18 To 26
            For j = 2 To 8
                Select Case i
                    Case 18: 参数类型 = "发电小时数"
                    Case 19: 参数类型 = "电价"
                    Case 20: 参数类型 = "消纳率"
                    Case 21: 参数类型 = "初始总投资"
                    Case 22: 参数类型 = "消交流侧装机容量"
                    Case 23: 参数类型 = "技改费"
                    Case 24: 参数类型 = "股权资本金占比"
                    Case 25: 参数类型 = "还款年限"
                    Case 26: 参数类型 = "利率"
                End Select
                
                变化率 = Format((j - 4) * 0.05, "0.00")  ' 转换为字符串格式
                
                .Cells(i, j).Formula = "=CalculateParameterValue(""" & 参数类型 & """, " & 变化率 & ", ""资本金"")"
            Next j
        Next i
        
        ' 敏感性系数计算
        For i = 27 To 35
            For j = 2 To 8
                Select Case i
                    Case 27: 参数类型 = "发电小时数"
                    Case 28: 参数类型 = "电价"
                    Case 29: 参数类型 = "消纳率"
                    Case 30: 参数类型 = "初始总投资"
                    Case 31: 参数类型 = "消交流侧装机容量"
                    Case 32: 参数类型 = "技改费"
                    Case 33: 参数类型 = "股权资本金占比"
                    Case 34: 参数类型 = "还款年限"
                    Case 35: 参数类型 = "利率"
                End Select
                
                变化率 = Format((j - 4) * 0.05, "0.00")  ' 转换为字符串格式

                .Cells(i, j).Formula = "=IF(B" & (i - 13) & "<>0,((F" & (i - 13) & "-B" & (i - 13) & ")/B" & (i - 13) & ")/(0.1*2),0)"  ' 全投资IRR敏感性系数
            .Cells(i, 3).Formula = "=IF(B" & (i - 6) & "<>0,((F" & (i - 6) & "-B" & (i - 6) & ")/B" & (i - 6) & ")/(0.1*2),0)"    ' 资本金IRR敏感性系数
        Next i
    End With
End Sub

Public Sub CreateSingleParameterAnalysis()
    Dim ws As Worksheet
    Dim baseSheet As Worksheet
    
    ' 检查是否已存在单一参数敏感性分析表，如果存在则删除
    On Error Resume Next
    Application.displayAlerts = False
    ThisWorkbook.Sheets("单一参数敏感性分析").Delete
    Application.displayAlerts = True
    On Error GoTo 0
    
    ' 创建新的工作表
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "单一参数敏感性分析"
    
    ' 设置基本结构
    With ws
        ' 设置标题说明
        .Range("B2").Value = "请在B11单元格开始输入变动后的值"
        .Range("B2:G2").Merge
        .Range("B2").HorizontalAlignment = xlCenter
        .Range("B2").Font.Bold = True
        
        ' 合并单元格
        .Range("B3:D3").Merge
        .Range("B4:D4").Merge
        .Range("B5:D5").Merge
        .Range("B6:D6").Merge
        .Range("B7:D7").Merge
        .Range("B8:D8").Merge
        .Range("B9:G9").Merge
        .Range("E3:G3").Merge
        .Range("E4:G4").Merge
        .Range("E5:G5").Merge
        .Range("E6:G6").Merge
        .Range("E7:G7").Merge
        .Range("E8:G8").Merge
        .Range("E9:G9").Merge
        
        ' 设置标题
        .Range("B3:D3").Value = "参数名称"
        .Range("B4:D4").Value = "参数原值"
        .Range("B5:D5").Value = "原全投资IRR"
        .Range("B6:D6").Value = "原资本金IRR"
        .Range("B7:D7").Value = "原全投资回收期（年）"
        .Range("B8:D8").Value = "原资本金回收期（年）"
        .Range("B9:G9").Value = "结果分析"

        
        ' 设置表头
        .Range("B10:G10") = Array("变动后的值", "变动率", "变动后全投资IRR", "变动后资本金IRR", "变动后全投资回收期", "变动后资本金回收期")
        
        ' 在I4单元格插入运行结果按钮
        Dim btnResult As Object
        Set btnResult = .Shapes.AddFormControl(xlButtonControl, .Range("I4").Left, .Range("I4").Top, .Range("I4").Width * 1.5, .Range("I4").Height)
        With btnResult
            .OnAction = "CalculateResults"
            .TextFrame.Characters.Text = "运行结果"
        End With
        
        ' 在I2单元格插入选择参数按钮
        Dim btnSelect As Object
        Set btnSelect = .Shapes.AddFormControl(xlButtonControl, .Range("I2").Left, .Range("I2").Top, .Range("I2").Width * 1.5, .Range("I2").Height)
        With btnSelect
            .OnAction = "ShowParameterSelector"
            .TextFrame.Characters.Text = "选择参数"
        End With
    End With
    
    ' 设置格式
    FormatParameterTable ws, 3
End Sub

' 计算结果
Public Sub CalculateResults()
    Dim ws As Worksheet
    Dim baseSheet As Worksheet
    Dim calcSheet As Worksheet
    Dim i As Long, j As Long
    Dim paramName As String
    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim currentRow As Long
    Dim baseValue As Double
    Dim originalValue As Double
    
    ' 保存所有可能会改变的参数的原始值和公式
    Dim orig_发电小时数 As Variant
    Dim orig_电价 As Variant
    Dim orig_消纳率 As Variant
    Dim orig_初始总投资 As Variant
    Dim orig_装机容量 As Variant
    Dim orig_技改费 As Variant
    Dim orig_资本金占比 As Variant
    Dim orig_还款年限 As Variant
    Dim orig_利率 As Variant
    
    Dim orig_发电小时数_Formula As String
    Dim orig_电价_Formula As String
    Dim orig_消纳率_Formula As String
    Dim orig_初始总投资_Formula As String
    Dim orig_装机容量_Formula As String
    Dim orig_技改费_Formula As String
    Dim orig_资本金占比_Formula As String
    Dim orig_还款年限_Formula As String
    Dim orig_利率_Formula As String

    ' 保存Excel设置
    oldCalculation = Application.calculation
    oldScreenUpdating = Application.screenUpdating
    
    ' 优化性能设置
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    
    Set ws = ThisWorkbook.Sheets("单一参数敏感性分析")
    Set baseSheet = ThisWorkbook.Sheets("基础参数及输出结果表")
    Set calcSheet = ThisWorkbook.Sheets("光伏收益测算表")

    ' 保存所有参数的原始值和公式
    With baseSheet
        orig_发电小时数 = .Range("B23").Value
        orig_发电小时数_Formula = .Range("B23").Formula
        
        orig_电价 = .Range("B42").Value
        orig_电价_Formula = .Range("B42").Formula
        
        orig_消纳率 = .Range("B37").Value
        orig_消纳率_Formula = .Range("B37").Formula
        
        orig_初始总投资 = .Range("F7").Value
        orig_初始总投资_Formula = .Range("F7").Formula
        
        orig_装机容量 = .Range("B12").Value
        orig_装机容量_Formula = .Range("B12").Formula
        
        orig_技改费 = .Range("F25").Value
        orig_技改费_Formula = .Range("F25").Formula
        
        orig_资本金占比 = .Range("J23").Value
        orig_资本金占比_Formula = .Range("J23").Formula
        
        orig_还款年限 = .Range("J26").Value
        orig_还款年限_Formula = .Range("J26").Formula
        
        orig_利率 = .Range("J27").Value
        orig_利率_Formula = .Range("J27").Formula
    End With
    
    ' 从第3行开始查找所有参数表格
    currentRow = 3
    
    On Error GoTo ErrorHandler
    
    Do While currentRow <= ws.Cells(ws.Rows.Count, "B").End(xlUp).row
        ' 检查是否是参数表格的开始（通过检查"参数名称"单元格）
        If ws.Cells(currentRow, 2).Value = "参数名称" Then
            ' 获取参数名称
            paramName = ws.Range("E" & currentRow).Value
            ' 获取基准值
            baseValue = ws.Range("E" & (currentRow + 1)).Value
            
            ' 计算该表格的6行输入值
            For i = currentRow + 8 To currentRow + 13
                If Not IsEmpty(ws.Cells(i, 2).Value) And IsNumeric(ws.Cells(i, 2).Value) Then
                    Dim newValue As Double
                    newValue = CDbl(ws.Cells(i, 2).Value)
                    
                    ' 计算变动率
                    ws.Cells(i, 3).Value = (newValue - baseValue) / baseValue
                    
                    ' 修改参数值
                    Select Case paramName
                        Case "发电小时数（单位：小时）"
                            baseSheet.Range("B23").Value = newValue
                        Case "电价（单位：元/kWh）"
                            baseSheet.Range("B42").Value = newValue
                        Case "消纳率（比率）"
                            baseSheet.Range("B37").Value = newValue
                        Case "初始总投资（万元）"
                            baseSheet.Range("F7").Value = newValue
                        Case "消交流侧装机容量（备案容量）"
                            baseSheet.Range("B12").Value = newValue
                        Case "技改费（元/W）"
                            baseSheet.Range("F25").Value = newValue
                        Case "股权资本金占比"
                            baseSheet.Range("J23").Value = newValue
                        Case "还款年限（年）"
                            baseSheet.Range("J26").Value = newValue
                        Case "利率"
                            baseSheet.Range("J27").Value = newValue
                    End Select
                    
                    ' 强制计算所有工作表
                    Application.Calculate
                    
                    ' 获取新的IRR和回收期值
                    ws.Cells(i, 4).Value = baseSheet.Range("N8").Value  ' 变动后全投资IRR
                    ws.Cells(i, 5).Value = baseSheet.Range("N10").Value ' 变动后资本金IRR
                    ws.Cells(i, 6).Value = baseSheet.Range("N9").Value  ' 变动后全投资回收期
                    ws.Cells(i, 7).Value = baseSheet.Range("N11").Value ' 变动后资本金回收期
                    
                    ' 每次计算完一个值后就恢复原始参数
                    With baseSheet
                        If orig_发电小时数_Formula <> "" Then
                            .Range("B23").Formula = orig_发电小时数_Formula
                        Else
                            .Range("B23").Value = orig_发电小时数
                        End If
                        
                        If orig_电价_Formula <> "" Then
                            .Range("B42").Formula = orig_电价_Formula
                        Else
                            .Range("B42").Value = orig_电价
                        End If
                        
                        If orig_消纳率_Formula <> "" Then
                            .Range("B37").Formula = orig_消纳率_Formula
                        Else
                            .Range("B37").Value = orig_消纳率
                        End If
                        
                        If orig_初始总投资_Formula <> "" Then
                            .Range("F7").Formula = orig_初始总投资_Formula
                        Else
                            .Range("F7").Value = orig_初始总投资
                        End If
                        
                        If orig_装机容量_Formula <> "" Then
                            .Range("B12").Formula = orig_装机容量_Formula
                        Else
                            .Range("B12").Value = orig_装机容量
                        End If
                        
                        If orig_技改费_Formula <> "" Then
                            .Range("F25").Formula = orig_技改费_Formula
                        Else
                            .Range("F25").Value = orig_技改费
                        End If
                        
                        If orig_资本金占比_Formula <> "" Then
                            .Range("J23").Formula = orig_资本金占比_Formula
                        Else
                            .Range("J23").Value = orig_资本金占比
                        End If
                        
                        If orig_还款年限_Formula <> "" Then
                            .Range("J26").Formula = orig_还款年限_Formula
                        Else
                            .Range("J26").Value = orig_还款年限
                        End If
                        
                        If orig_利率_Formula <> "" Then
                            .Range("J27").Formula = orig_利率_Formula
                        Else
                            .Range("J27").Value = orig_利率
                        End If
                    End With
                    
                    ' 恢复后重新计算
                    Application.Calculate
                End If
            Next i
            
            ' 设置格式
            With ws
                .Range("B" & (currentRow + 8) & ":B" & (currentRow + 13)).NumberFormat = "0.00"        ' 变动后的值：保留两位小数
                .Range("C" & (currentRow + 8) & ":C" & (currentRow + 13)).NumberFormat = "0.00%"       ' 变动率：百分比格式
                .Range("D" & (currentRow + 8) & ":E" & (currentRow + 13)).NumberFormat = "0.00%"       ' IRR值：百分比格式
                .Range("F" & (currentRow + 8) & ":G" & (currentRow + 13)).NumberFormat = "0.00"        ' 回收期：保留两位小数
            End With
            
            ' 移动到下一个表格（当前表格13行 + 2行间隔）
            currentRow = currentRow + 15
        Else
            ' 如果不是参数表格的开始，移动到下一行
            currentRow = currentRow + 1
        End If
    Loop

ExitSub:
    ' 最后再次确保所有参数都恢复到原始值
    With baseSheet
        If orig_发电小时数_Formula <> "" Then
            .Range("B23").Formula = orig_发电小时数_Formula
        Else
            .Range("B23").Value = orig_发电小时数
        End If
        
        If orig_电价_Formula <> "" Then
            .Range("B42").Formula = orig_电价_Formula
        Else
            .Range("B42").Value = orig_电价
        End If
        
        If orig_消纳率_Formula <> "" Then
            .Range("B37").Formula = orig_消纳率_Formula
        Else
            .Range("B37").Value = orig_消纳率
        End If
        
        If orig_初始总投资_Formula <> "" Then
            .Range("F7").Formula = orig_初始总投资_Formula
        Else
            .Range("F7").Value = orig_初始总投资
        End If
        
        If orig_装机容量_Formula <> "" Then
            .Range("B12").Formula = orig_装机容量_Formula
        Else
            .Range("B12").Value = orig_装机容量
        End If
        
        If orig_技改费_Formula <> "" Then
            .Range("F25").Formula = orig_技改费_Formula
        Else
            .Range("F25").Value = orig_技改费
        End If
        
        If orig_资本金占比_Formula <> "" Then
            .Range("J23").Formula = orig_资本金占比_Formula
        Else
            .Range("J23").Value = orig_资本金占比
        End If
        
        If orig_还款年限_Formula <> "" Then
            .Range("J26").Formula = orig_还款年限_Formula
        Else
            .Range("J26").Value = orig_还款年限
        End If
        
        If orig_利率_Formula <> "" Then
            .Range("J27").Formula = orig_利率_Formula
        Else
            .Range("J27").Value = orig_利率
        End If
    End With
    
    ' 最后计算一次以确保所有值都正确
    Application.Calculate
    
    ' 恢复Excel设置
    Application.calculation = oldCalculation
    Application.screenUpdating = oldScreenUpdating
    
    MsgBox "计算完成！", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "计算过程中出现错误！" & vbNewLine & "错误描述: " & Err.Description, vbCritical
    Resume ExitSub
End Sub







