Sub TransferParametersAndRetrieveResults()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long, i As Long
    Dim rngParams As Range
    Dim destCells(1 To 13) As Range
    
     ' 设置工作表对象
    Set ws1 = ThisWorkbook.Sheets("测算汇总-运算结果")
    Set ws2 = ThisWorkbook.Sheets("1.小储项目运营测算")
    
    ' 保存ws2工作表的所有公式
    Dim saveRange As Range
    Set saveRange = ws2.UsedRange
    Dim savedFormulas As Variant
    savedFormulas = saveRange.Formula
    
    ' 初始化目标单元格的Range对象数组
    Set destCells(1) = ws2.Range("F2") ' 地区
    Set destCells(2) = ws2.Range("B5") ' 项目规模
    Set destCells(3) = ws2.Range("B4") ' 运行期限
    Set destCells(4) = ws2.Range("B13") ' 年充放天数
    Set destCells(5) = ws2.Range("D155") ' 峰平-放电折算次数
    Set destCells(6) = ws2.Range("D157") ' 峰平-放电折算次数
    Set destCells(7) = ws2.Range("B12") ' 充放循环次数
    Set destCells(8) = ws2.Range("B8") ' 资方分成比例
    Set destCells(9) = ws2.Range("D3") ' EPC单价：元/Wh
    Set destCells(10) = ws2.Range("B6") ' 运营运维费率
    Set destCells(11) = ws2.Range("D7") ' 居间成本
    Set destCells(12) = ws2.Range("I6") ' 增值税率（运营期）
    Set destCells(13) = ws2.Range("I7") ' 所得税率
    
    ' 确定要处理的数据行数
    lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    
    ' 设置计算模式
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    ' 遍历每一行
    For i = 8 To lastRow
        ' 复制数据到ws2
        Set rngParams = ws1.Range(ws1.Cells(i, 1), ws1.Cells(i, 13))
        For j = 1 To 13
            destCells(j).Value = rngParams.Cells(1, j).Value
        Next j
        
        ' 强制计算当前工作表
        ws2.Calculate
        Application.CalculateFullRebuild
        
        ' 先获取计算结果到变量中
        Dim irr1 As Variant, irr2 As Variant
        Dim payback1 As Variant, payback2 As Variant
        Dim diff1 As Variant, diff2 As Variant
        
        On Error Resume Next
        
        ' 获取计算结果
        irr1 = ws2.Range("D27").Value
        irr2 = ws2.Range("D26").Value
        payback1 = ws2.Range("E27").Value
        payback2 = ws2.Range("E26").Value
        diff1 = ws2.Range("F15").Value
        diff2 = ws2.Range("F17").Value
        
        ' 设置格式并写入结果
        With ws1
            ' IRR结果
            .Cells(i, 14).NumberFormat = "0.00%"
            If IsNumeric(irr1) Then
                .Cells(i, 14).Value = irr1
            Else
                .Cells(i, 14).Value = "N/A"
            End If
            
            .Cells(i, 15).NumberFormat = "0.00%"
            If IsNumeric(irr2) Then
                .Cells(i, 15).Value = irr2
            Else
                .Cells(i, 15).Value = "N/A"
            End If
            
            ' 回收期结果
            .Cells(i, 16).NumberFormat = "0.00"
            If IsNumeric(payback1) Then
                .Cells(i, 16).Value = payback1
            Else
                .Cells(i, 16).Value = "N/A"
            End If
            
            .Cells(i, 17).NumberFormat = "0.00"
            If IsNumeric(payback2) Then
                .Cells(i, 17).Value = payback2
            Else
                .Cells(i, 17).Value = "N/A"
            End If
            
            ' 价差结果
            .Cells(i, 18).NumberFormat = "0.00000"
            If IsNumeric(diff1) Then
                .Cells(i, 18).Value = diff1
            Else
                .Cells(i, 18).Value = "N/A"
            End If
            
            .Cells(i, 19).NumberFormat = "0.00000"
            If IsNumeric(diff2) Then
                .Cells(i, 19).Value = diff2
            Else
                .Cells(i, 19).Value = "N/A"
            End If
        End With
        
        On Error GoTo 0
    Next i
    
    ' 恢复ws2的所有公式
    saveRange.Formula = savedFormulas
    
    ' 恢复Excel设置
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    ' 最后计算一次
    ws2.Calculate
End Sub


