' 添加生成柱状图按钮
On Error Resume Next
Dim btnChart As Object
Set btnChart = ws.Buttons.Add(ws.Range("AA12").Left, ws.Range("AA12").Top, 80, 25)
If Not btnChart Is Nothing Then
    With btnChart
        .OnAction = ThisWorkbook.Name & "!CreateTimeOfUsePricingCharts"
        .Caption = "生成分时电价柱状图"
    End With
End If
On Error GoTo 0

' 新增函数：创建分时电价柱状图
Sub CreateTimeOfUsePricingCharts()
    On Error Resume Next
    
    ' 获取当前工作表
    Dim wsSingle As Worksheet
    Set wsSingle = ActiveSheet
    If wsSingle Is Nothing Then
        MsgBox "无法获取当前工作表", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的地区
    Dim selectedRegion As String
    selectedRegion = wsSingle.Range("B1").Value
    If selectedRegion = "" Then
        MsgBox "请先选择地区", vbExclamation
        Exit Sub
    End If
    
    ' 获取源数据工作表
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("解析结果")
    If wsSource Is Nothing Then
        MsgBox "未找到解析结果工作表", vbExclamation
        Exit Sub
    End If
    
    ' 获取最后一行
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' 创建或激活分时电价时段柱状图工作表
    Dim wsChart As Worksheet
    On Error Resume Next
    Set wsChart = ThisWorkbook.Sheets("分时电价时段柱状图")
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    If wsChart Is Nothing Then
        Set wsChart = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsChart.Name = "分时电价时段柱状图"
    Else
        wsChart.Cells.Clear
        wsChart.ChartObjects.Delete
    End If
    
    ' 在A1单元格显示地区名称
    wsChart.Range("A1").Value = "地区：" & selectedRegion
    
    ' 定义时段高度
    Const DEEP_VALLEY_HEIGHT As Double = 0.2
    Const VALLEY_HEIGHT As Double = 0.4
    Const FLAT_HEIGHT As Double = 0.6
    Const PEAK_HEIGHT As Double = 0.8
    Const SHARP_PEAK_HEIGHT As Double = 1.0
    
    ' 循环处理12个月份
    Dim monthNum As Integer
    Dim chartTop As Long
    chartTop = 50
    
    ' 处理每个月份
    For monthNum = 1 To 12
        ' 处理所有时段类型
        For i = 2 To lastRow
            If Trim(wsSource.Cells(i, 1).value) = selectedRegion Then
                Dim currentMonth As Integer
                currentMonth = CInt(Replace(wsSource.Cells(i, 3).value, "月", ""))
                
                ' 直接检查月份是否匹配
                If currentMonth = monthNum Then
                    ' 获取时段类型
                    Dim timeValue As Integer
                    Select Case Trim(wsSource.Cells(i, 4).value)
                        Case "尖峰": timeValue = PEAK
                        Case "高峰": timeValue = HIGH
                        Case "平段": timeValue = NORMAL
                        Case "低谷": timeValue = LOW
                        Case "深谷": timeValue = DEEP_LOW
                    End Select
                    
                    ' 创建数据区域
                    Dim dataRange As Range
                    Set dataRange = wsChart.Range(wsChart.Cells(chartTop, 1), wsChart.Cells(chartTop + 1, 25))
                    
                    ' 填充时间标签
                    Dim hour As Integer
                    For hour = 0 To 23
                        dataRange.Cells(1, hour + 2).Value = hour & "-" & (hour + 1)
                    Next hour
                    
                    ' 获取该月的时段数据
                    Dim timeSlotRange As Range
                    Set timeSlotRange = wsSingle.Range(wsSingle.Cells(monthNum + 3, 2), wsSingle.Cells(monthNum + 3, 25))
                    
                    ' 填充高度数据
                    Dim col As Integer
                    For col = 1 To 24
                        Select Case timeSlotRange.Cells(1, col).Interior.Color
                            Case RGB(0, 176, 80)    ' 深谷
                                dataRange.Cells(2, col + 1).Value = DEEP_VALLEY_HEIGHT
                            Case RGB(146, 208, 80)   ' 低谷
                                dataRange.Cells(2, col + 1).Value = VALLEY_HEIGHT
                            Case RGB(255, 255, 0)    ' 平段
                                dataRange.Cells(2, col + 1).Value = FLAT_HEIGHT
                            Case RGB(255, 192, 0)    ' 高峰
                                dataRange.Cells(2, col + 1).Value = PEAK_HEIGHT
                            Case RGB(255, 0, 0)      ' 尖峰
                                dataRange.Cells(2, col + 1).Value = SHARP_PEAK_HEIGHT
                        End Select
                    Next col
                    
                    ' 创建柱状图
                    On Error Resume Next
                    Dim chartObj As ChartObject
                    Set chartObj = wsChart.ChartObjects.Add(Left:=50, Top:=chartTop, Width:=600, Height:=300)
                    
                    If Not chartObj Is Nothing Then
                        With chartObj.Chart
                            .ChartType = xlColumnClustered
                            .SetSourceData Source:=dataRange
                            .HasTitle = True
                            .ChartTitle.Text = monthNum & "月分时电价时段柱状图"
                            
                            ' 设置柱状图颜色
                            Dim i As Integer
                            For i = 1 To .SeriesCollection(1).Points.Count
                                Select Case timeSlotRange.Cells(1, i).Interior.Color
                                    Case RGB(0, 176, 80)    ' 深谷
                                        .SeriesCollection(1).Points(i).Interior.Color = RGB(0, 176, 80)
                                    Case RGB(146, 208, 80)   ' 低谷
                                        .SeriesCollection(1).Points(i).Interior.Color = RGB(146, 208, 80)
                                    Case RGB(255, 255, 0)    ' 平段
                                        .SeriesCollection(1).Points(i).Interior.Color = RGB(255, 255, 0)
                                    Case RGB(255, 192, 0)    ' 高峰
                                        .SeriesCollection(1).Points(i).Interior.Color = RGB(255, 192, 0)
                                    Case RGB(255, 0, 0)      ' 尖峰
                                        .SeriesCollection(1).Points(i).Interior.Color = RGB(255, 0, 0)
                                End Select
                            Next i
                            
                            ' 添加图例
                            .HasLegend = True
                            With .Legend
                                .Position = xlBottom
                                
                                ' 创建自定义图例
                                Dim legendRange As Range
                                Set legendRange = wsChart.Range(wsChart.Cells(chartTop + 2, 1), wsChart.Cells(chartTop + 2, 5))
                                legendRange.Interior.ColorIndex = xlNone
                                
                                ' 填充图例数据
                                legendRange.Cells(1, 1).Value = "尖峰"
                                legendRange.Cells(1, 2).Value = "高峰"
                                legendRange.Cells(1, 3).Value = "平段"
                                legendRange.Cells(1, 4).Value = "低谷"
                                legendRange.Cells(1, 5).Value = "深谷"
                                
                                ' 设置图例颜色
                                legendRange.Cells(1, 1).Interior.Color = RGB(255, 0, 0)
                                legendRange.Cells(1, 2).Interior.Color = RGB(255, 192, 0)
                                legendRange.Cells(1, 3).Interior.Color = RGB(255, 255, 0)
                                legendRange.Cells(1, 4).Interior.Color = RGB(146, 208, 80)
                                legendRange.Cells(1, 5).Interior.Color = RGB(0, 176, 80)
                            End With
                        End With
                    End If
                    On Error GoTo 0
                    
                    ' 分析并添加收益模式说明
                    Dim profitMode As String
                    profitMode = AnalyzeProfitMode(timeSlotRange)
                    wsChart.Cells(chartTop - 20, 1).Value = profitMode
                    
                    chartTop = chartTop + 350
                End If
            End If
        Next i
    Next monthNum

    ' 调整工作表视图
    wsChart.Activate
    ActiveWindow.Zoom = 70

    Application.ScreenUpdating = True
End Sub

' 新增函数：分析收益模式
Function AnalyzeProfitMode(timeSlotRange As Range) As String
    Dim profitMode As String
    Dim firstProfit As String
    Dim secondProfit As String
    
    ' 分析第一次套利机会
    Dim i As Integer
    Dim j As Integer  ' 添加j变量声明
    Dim chargeTime As Integer
    Dim dischargeTime As Integer
    
    For i = 1 To 24
        ' 寻找充电时段（低谷或平段）
        If timeSlotRange.Cells(1, i).Interior.Color = RGB(146, 208, 80) Or _
           timeSlotRange.Cells(1, i).Interior.Color = RGB(255, 255, 0) Then
            chargeTime = i
            ' 寻找放电时段（高峰或尖峰）
            For j = i + 1 To 24
                If timeSlotRange.Cells(1, j).Interior.Color = RGB(255, 192, 0) Or _
                   timeSlotRange.Cells(1, j).Interior.Color = RGB(255, 0, 0) Then
                    dischargeTime = j
                    ' 确定第一次套利类型
                    If timeSlotRange.Cells(1, chargeTime).Interior.Color = RGB(146, 208, 80) Then
                        firstProfit = "第一次：峰谷套利"
                    Else
                        firstProfit = "第一次：峰平套利"
                    End If
                    Exit For
                End If
            Next j
            If dischargeTime > 0 Then Exit For
        End If
    Next i
    
    ' 分析第二次套利机会
    chargeTime = 0
    dischargeTime = 0
    For i = dischargeTime + 1 To 24
        ' 寻找充电时段（低谷或平段）
        If timeSlotRange.Cells(1, i).Interior.Color = RGB(146, 208, 80) Or _
           timeSlotRange.Cells(1, i).Interior.Color = RGB(255, 255, 0) Then
            chargeTime = i
            ' 寻找放电时段（高峰或尖峰）
            For j = i + 1 To 24
                If timeSlotRange.Cells(1, j).Interior.Color = RGB(255, 192, 0) Or _
                   timeSlotRange.Cells(1, j).Interior.Color = RGB(255, 0, 0) Then
                    dischargeTime = j
                    ' 确定第二次套利类型
                    If timeSlotRange.Cells(1, j).Interior.Color = RGB(255, 0, 0) Then
                        secondProfit = "第二次：尖平套利"
                    Else
                        secondProfit = "第二次：峰平套利"
                    End If
                    Exit For
                End If
            Next j
            If dischargeTime > 0 Then Exit For
        End If
    Next i
    
    ' 组合收益模式说明
    profitMode = firstProfit
    If secondProfit <> "" Then
        profitMode = profitMode & vbNewLine & secondProfit
    End If
    
    AnalyzeProfitMode = profitMode
End Function

