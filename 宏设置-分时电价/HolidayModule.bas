Option Explicit

' 时段类型常量
Private Const PEAK As Integer = 1       ' 尖峰
Private Const HIGH As Integer = 2       ' 高峰
Private Const NORMAL As Integer = 3     ' 平段
Private Const LOW As Integer = 4        ' 低谷
Private Const DEEP_LOW As Integer = 5   ' 深谷

' 高度值常量
Private Const SHARP_PEAK_HEIGHT As Double = 1#    ' 尖峰高度
Private Const PEAK_HEIGHT As Double = 0.9         ' 高峰高度
Private Const FLAT_HEIGHT As Double = 0.6         ' 平段高度
Private Const VALLEY_HEIGHT As Double = 0.3       ' 低谷高度
Private Const DEEP_VALLEY_HEIGHT As Double = 0.1  ' 深谷高度

' 选择地区功能
Public Sub SelectHolidayRegion()
    ' 创建并显示节假日地区选择窗体
    HolidayRegionSelectorForm.Show
End Sub

' 更新选定地区的数据
Public Sub UpdateSelectedRegion(ByVal selectedRegion As String)
    ' 获取工作表
    Dim wsHoliday As Worksheet
    Set wsHoliday = ThisWorkbook.Worksheets("国家节假日时段分析")
    
    ' 获取源数据工作表
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("解析国家节假日文字")
    
    ' 更新选定地区
    wsHoliday.Range("B1").value = selectedRegion
    
    ' 检查节日名称是否正确填充
    Dim holidays As Variant
    holidays = Array("元旦", "春节", "清明节", "劳动节", "端午节", "中秋节", "国庆节")
    
    ' 检查A列的节日名称
    Dim i As Long
    For i = 0 To UBound(holidays)
        If wsHoliday.Cells(i + 2, 1).value <> holidays(i) Then
            wsHoliday.Cells(i + 2, 1).value = holidays(i)
            Debug.Print "修正节日名称: " & holidays(i)
        End If
    Next i
    
    ' 填充时段数据
    Call FillHolidayTimeData(selectedRegion, wsHoliday, wsSource)
    
    ' 更新时段表颜色
    Call UpdateHolidayTimeTable
    
    ' 生成柱状图
    Call CreateHolidayTimeCharts
End Sub

' 填充节假日时段数据
Public Sub FillHolidayTimeData(ByVal selectedRegion As String, ByVal wsHoliday As Worksheet, ByVal wsSource As Worksheet)
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim timeConfigs() As Integer
    ReDim timeConfigs(1 To 7, 0 To 23)  ' 7个节日 x 24小时
    
    ' 获取最后一行
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    Debug.Print "总行数: " & lastRow
    
    ' 清除现有数据
    wsHoliday.Range("B17:Y23").ClearContents
    
    ' 处理每一行数据
    For i = 2 To lastRow
        If wsSource.Cells(i, 1).value = selectedRegion Then
            ' 获取节日名称和时段类型
            Dim holidayName As String
            Dim timeValue As Integer
            
            holidayName = Trim(wsSource.Cells(i, 2).value)  ' 获取节日名称
            
            ' 获取时段类型并处理格式
            Dim timeType As String
            timeType = wsSource.Cells(i, 3).value
            Debug.Print "原始时段类型: [" & timeType & "]"
            
            ' 移除所有可能的特殊字符
            timeType = Replace(timeType, Chr(160), "") ' 移除 NO-BREAK SPACE
            timeType = Replace(timeType, vbCr, "")     ' 移除回车
            timeType = Replace(timeType, vbLf, "")     ' 移除换行
            timeType = Replace(timeType, vbTab, "")    ' 移除制表符
            timeType = Trim(timeType)                  ' 移除前后空格
            timeType = LCase(timeType)                 ' 转换为小写
            
            Debug.Print "处理后时段类型: [" & timeType & "]"
            Debug.Print "ASCII码: " & ShowAsciiCodes(timeType)
            
            Select Case timeType
                Case "尖峰", "尖峰时段"
                    timeValue = PEAK
                    Debug.Print "设置为尖峰时段: " & timeValue
                Case "高峰", "高峰时段"
                    timeValue = HIGH
                    Debug.Print "设置为高峰时段: " & timeValue
                Case "平段", "平段时段"
                    timeValue = NORMAL
                    Debug.Print "设置为平段时段: " & timeValue
                Case "低谷", "低谷时段"
                    timeValue = LOW
                    Debug.Print "设置为低谷时段: " & timeValue
                Case "深谷", "深谷时段"
                    timeValue = DEEP_LOW
                    Debug.Print "设置为深谷时段: " & timeValue
                Case Else
                    Debug.Print "未知时段类型: [" & timeType & "]"
                    timeValue = 0
            End Select
            
            ' 如果是未知时段类型，跳过处理
            If timeValue = 0 Then
                Debug.Print "跳过未知时段类型的处理"
                GoTo NextIteration
            End If
            
            ' 获取时间范围
            Dim startTime As String, endTime As String
            startTime = Trim(wsSource.Cells(i, 4).value)
            endTime = Trim(wsSource.Cells(i, 5).value)
            Debug.Print "原始时间范围: [" & startTime & "] - [" & endTime & "]"
            
            ' 处理时间值
            Dim startHour As Long, endHour As Long
            
            ' 处理开始时间
            If InStr(startTime, ":") > 0 Then
                startHour = CLng(Split(startTime, ":")(0))
                Debug.Print "解析开始时间: " & startTime & " -> " & startHour & "时"
            ElseIf IsNumeric(startTime) Then
                ' 将小数时间转换为小时，四舍五入到最接近的小时
                startHour = Round(CDbl(startTime) * 24, 0)
                Debug.Print "解析开始时间(小数): " & startTime & " -> " & startHour & "时 (原始值=" & CDbl(startTime) * 24 & ")"
            Else
                Debug.Print "无法解析开始时间: " & startTime
                GoTo NextIteration
            End If
            
            ' 处理结束时间
            If InStr(endTime, ":") > 0 Then
                endHour = CLng(Split(endTime, ":")(0))
                Debug.Print "解析结束时间: " & endTime & " -> " & endHour & "时"
            ElseIf IsNumeric(endTime) Then
                ' 将小数时间转换为小时，四舍五入到最接近的小时
                endHour = Round(CDbl(endTime) * 24, 0)
                Debug.Print "解析结束时间(小数): " & endTime & " -> " & endHour & "时 (原始值=" & CDbl(endTime) * 24 & ")"
            Else
                Debug.Print "无法解析结束时间: " & endTime
                GoTo NextIteration
            End If
            
            ' 处理特殊的24:00情况
            If endTime = "24:00" Or endTime = "24:00:00" Or CDbl(endTime) >= 1 Then
                If CDbl(endTime) >= 1 Then
                    endHour = 24
                    Debug.Print "处理特殊结束时间 1.0 -> 24时"
                End If
            End If
            
            Debug.Print "最终时间范围: " & startHour & "时 到 " & endHour & "时"
            
            ' 处理跨天的情况
            If endHour < startHour And endHour <> 24 Then
                endHour = endHour + 24
                Debug.Print "处理跨天情况，调整后结束时间: " & endHour & "时"
            End If
            
            ' 验证时间范围的有效性
            If startHour < 0 Or startHour >= 24 Or endHour < 0 Or endHour > 24 Then
                Debug.Print "无效的时间范围: " & startHour & "时 到 " & endHour & "时"
                GoTo NextIteration
            End If
            
            Debug.Print "时段类型: " & timeType & ", 值: " & timeValue & ", 时间范围: " & startHour & "-" & endHour
            
            ' 找到对应的节日行
            Dim holidayRow As Long
            For j = 2 To 8  ' 7个节日
                If wsHoliday.Cells(j, 1).value = holidayName Then  ' 匹配节日名称
                    holidayRow = j  ' 保存当前节日行号
                    Debug.Print "找到节日行: " & holidayRow
                    
                    ' 填充时段配置
                    Dim k As Long
                    For k = startHour To endHour - 1  ' 修改为不包含结束小时
                        Dim targetHour As Long
                        targetHour = k
                        If targetHour >= 24 Then
                            targetHour = targetHour - 24
                        End If
                        
                        ' 检查单元格是否已有值，如果是深谷时段，则优先保留
                        Dim currentValue As Variant
                        currentValue = wsHoliday.Cells(holidayRow + 15, targetHour + 2).value
                        
                        ' 如果单元格为空或当前是深谷时段，则写入新值
                        If IsEmpty(currentValue) Or timeValue = DEEP_LOW Then
                            ' 直接写入到工作表
                            wsHoliday.Cells(holidayRow + 15, targetHour + 2).value = timeValue
                            Debug.Print "填充单元格: " & "行=" & (holidayRow + 15) & ", 列=" & (targetHour + 2) & ", 值=" & timeValue
                            
                            ' 同时更新内存中的数组
                            timeConfigs(holidayRow - 1, targetHour) = timeValue
                        Else
                            Debug.Print "跳过已有值的单元格: " & "行=" & (holidayRow + 15) & ", 列=" & (targetHour + 2) & ", 当前值=" & currentValue
                        End If
                    Next k
                    Exit For  ' 找到匹配的节日后就退出循环
                End If
            Next j
NextIteration:
        End If
    Next i
    
    Debug.Print "数据填充完成"
End Sub

' 辅助函数：显示字符串的ASCII码
Private Function ShowAsciiCodes(ByVal text As String) As String
    Dim i As Long
    Dim result As String
    
    For i = 1 To Len(text)
        If result <> "" Then result = result & " "
        result = result & Asc(Mid(text, i, 1))
    Next i
    
    ShowAsciiCodes = result
End Function

' 更新节假日时段表
Public Sub UpdateHolidayTimeTable()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("国家节假日时段分析")
    
    ' 清除现有颜色
    ws.Range("B2:Y13").Interior.ColorIndex = xlNone
    
    ' 遍历配置区域并更新颜色
    Dim i As Long, j As Long
    For i = 1 To 12  ' 假设最多12个节假日
        For j = 0 To 23  ' 24小时
            If Not IsEmpty(ws.Cells(i + 16, j + 2)) Then
                Select Case ws.Cells(i + 16, j + 2).value
                    Case PEAK ' 尖峰
                        ws.Cells(i + 1, j + 2).Interior.color = RGB(255, 192, 0)  ' 橙色
                    Case HIGH ' 高峰
                        ws.Cells(i + 1, j + 2).Interior.color = RGB(255, 192, 203)  ' 粉红色
                    Case NORMAL ' 平段
                        ws.Cells(i + 1, j + 2).Interior.color = RGB(189, 215, 238) ' 浅蓝色
                    Case LOW ' 低谷
                        ws.Cells(i + 1, j + 2).Interior.color = RGB(198, 239, 206) ' 浅绿色
                    Case DEEP_LOW ' 深谷
                        ws.Cells(i + 1, j + 2).Interior.color = RGB(0, 112, 192)   ' 深蓝色
                End Select
            End If
        Next j
    Next i
    
    MsgBox "节假日时段表更新完成！", vbInformation
End Sub

' 新增：创建节假日时段柱状图
Sub CreateHolidayTimeCharts()
    ' 获取当前工作表
    Dim wsHoliday As Worksheet
    Set wsHoliday = ThisWorkbook.Worksheets("国家节假日时段分析")
    If wsHoliday Is Nothing Then
        MsgBox "无法获取节假日时段分析表", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的地区
    Dim selectedRegion As String
    selectedRegion = wsHoliday.Range("B1").value
    If selectedRegion = "" Then
        MsgBox "请先选择地区！", vbExclamation
        Exit Sub
    End If
    
    ' 获取源数据工作表
    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("解析国家节假日文字")
    If wsSource Is Nothing Then
        MsgBox "无法获取源数据工作表", vbExclamation
        Exit Sub
    End If
    
    ' 获取最后一行
    Dim lastRow As Long
    lastRow = wsHoliday.Cells(wsHoliday.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' 创建或激活节假日时段柱状图工作表
    Dim wsChart As Worksheet
    On Error Resume Next
    Set wsChart = ThisWorkbook.Sheets("节假日时段柱状图")
    If wsChart Is Nothing Then
        Set wsChart = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsChart.Name = "节假日时段柱状图"
    End If
    wsChart.Tab.ColorIndex = 50 ' 设置为青绿色
    
    Application.ScreenUpdating = False
    
    ' 清除现有内容
    wsChart.Cells.Clear
    wsChart.ChartObjects.Delete
    
    ' 定义时段高度
    Const DEEP_VALLEY_HEIGHT As Double = 0.1  ' 深谷
    Const VALLEY_HEIGHT As Double = 0.3      ' 低谷
    Const FLAT_HEIGHT As Double = 0.6        ' 平段
    Const PEAK_HEIGHT As Double = 0.9        ' 高峰
    Const SHARP_PEAK_HEIGHT As Double = 1.2   ' 尖峰
    
    ' 清除高度值表区域（从第32行开始）
    wsHoliday.Range("B32:Y38").ClearContents
    
    ' 循环处理每个节日
    Dim holidayIndex As Long
    Dim chartTop As Long
    chartTop = 50
    
    For holidayIndex = 2 To lastRow
        Dim holidayName As String
        holidayName = wsHoliday.Cells(holidayIndex, 1).value
        If holidayName = "" Then Exit For
        
        ' 获取该节日的时段配置数据（从时段状态配置区域）
        Dim configRange As Range
        Set configRange = wsHoliday.Range(wsHoliday.Cells(holidayIndex + 15, 2), wsHoliday.Cells(holidayIndex + 15, 25))
        
        ' 填充高度数据到高度值表（从第32行开始）
        Dim colIndex As Integer
        For colIndex = 1 To 24
            Select Case configRange.Cells(1, colIndex).value
                Case PEAK  ' 尖峰
                    wsHoliday.Cells(holidayIndex + 30, colIndex + 1).value = SHARP_PEAK_HEIGHT
                Case HIGH  ' 高峰
                    wsHoliday.Cells(holidayIndex + 30, colIndex + 1).value = PEAK_HEIGHT
                Case NORMAL  ' 平段
                    wsHoliday.Cells(holidayIndex + 30, colIndex + 1).value = FLAT_HEIGHT
                Case LOW  ' 低谷
                    wsHoliday.Cells(holidayIndex + 30, colIndex + 1).value = VALLEY_HEIGHT
                Case DEEP_LOW  ' 深谷
                    wsHoliday.Cells(holidayIndex + 30, colIndex + 1).value = DEEP_VALLEY_HEIGHT
            End Select
        Next colIndex
        
        ' 创建柱状图
        Dim chartObj As ChartObject
        Set chartObj = wsChart.ChartObjects.Add(Left:=50, Top:=chartTop, Width:=800, Height:=300)
        
        ' 分析收益模式
        Dim profitMode As String
        profitMode = AnalyzeProfitMode(configRange)
        
        With chartObj.Chart
            .ChartType = xlColumnStacked
            
            ' 设置数据源为高度值表
            .SetSourceData Source:=wsHoliday.Range(wsHoliday.Cells(holidayIndex + 30, 2), wsHoliday.Cells(holidayIndex + 30, 25))
            
            ' 设置系列名称为空
            With .SeriesCollection(1)
                .Name = ""
                .XValues = wsHoliday.Range(wsHoliday.Cells(31, 2), wsHoliday.Cells(31, 25))
            End With
            
            ' 禁用图例
            .HasLegend = False
            
            ' 设置图表标题
            If profitMode <> "" Then
                If InStr(profitMode, "第二次") > 0 Then
                    ' 将第二次套利信息合并到第一行
                    Dim firstProfit As String, secondProfit As String
                    firstProfit = Left(profitMode, InStr(profitMode, vbNewLine) - 1)
                    secondProfit = Mid(profitMode, InStr(profitMode, vbNewLine) + 2)
                    .ChartTitle.text = selectedRegion & " - " & holidayName & "分时电价时段柱状图" & vbNewLine & _
                                     firstProfit & "，" & secondProfit
                Else
                    .ChartTitle.text = selectedRegion & " - " & holidayName & "分时电价时段柱状图" & vbNewLine & profitMode
                End If
            Else
                .ChartTitle.text = selectedRegion & " - " & holidayName & "分时电价时段柱状图"
            End If
            
            ' 设置标题格式
            With .ChartTitle
                .Font.Size = 14
                .Font.Bold = True
            End With
            
            ' 设置柱状图颜色
            Dim pointIndex As Integer
            For pointIndex = 1 To .SeriesCollection(1).Points.Count
                Select Case configRange.Cells(1, pointIndex).value
                    Case PEAK  ' 尖峰
                        .SeriesCollection(1).Points(pointIndex).Interior.color = RGB(255, 192, 0)
                    Case HIGH  ' 高峰
                        .SeriesCollection(1).Points(pointIndex).Interior.color = RGB(255, 192, 203)
                    Case NORMAL  ' 平段
                        .SeriesCollection(1).Points(pointIndex).Interior.color = RGB(189, 215, 238)
                    Case LOW  ' 低谷
                        .SeriesCollection(1).Points(pointIndex).Interior.color = RGB(198, 239, 206)
                    Case DEEP_LOW  ' 深谷
                        .SeriesCollection(1).Points(pointIndex).Interior.color = RGB(0, 112, 192)
                End Select
            Next pointIndex
            
            ' 添加图例并设置位置
            .HasLegend = True
            With .Legend
                .Position = xlTop
                .IncludeInLayout = False
            End With
            
            ' 添加新的图例项
            Dim legendSeries As series
            
            Set legendSeries = .SeriesCollection.NewSeries
            With legendSeries
                .Name = "尖峰段"
                .Border.ColorIndex = xlNone
                .Interior.color = RGB(255, 192, 0)    ' 橙色
                .ChartType = xlColumnStacked
                .Values = Array(0)
            End With
            
            Set legendSeries = .SeriesCollection.NewSeries
            With legendSeries
                .Name = "高峰时段"
                .Border.ColorIndex = xlNone
                .Interior.color = RGB(255, 192, 203)    ' 粉红色
                .ChartType = xlColumnStacked
                .Values = Array(0)
            End With
            
            Set legendSeries = .SeriesCollection.NewSeries
            With legendSeries
                .Name = "平时段"
                .Border.ColorIndex = xlNone
                .Interior.color = RGB(189, 215, 238)    ' 浅蓝色
                .ChartType = xlColumnStacked
                .Values = Array(0)
            End With
            
            Set legendSeries = .SeriesCollection.NewSeries
            With legendSeries
                .Name = "低谷时段"
                .Border.ColorIndex = xlNone
                .Interior.color = RGB(198, 239, 206)    ' 浅绿色
                .ChartType = xlColumnStacked
                .Values = Array(0)
            End With
            
            Set legendSeries = .SeriesCollection.NewSeries
            With legendSeries
                .Name = "深谷时段"
                .Border.ColorIndex = xlNone
                .Interior.color = RGB(0, 112, 192)      ' 深蓝色
                .ChartType = xlColumnStacked
                .Values = Array(0)
            End With
        End With
        
        chartTop = chartTop + 350
    Next holidayIndex
    
    ' 清除节假日时段图表工作表中的所有单元格内容
    wsChart.Cells.Clear
    
    ' 调整工作表视图
    wsChart.Activate
    ActiveWindow.Zoom = 100
    
    Application.ScreenUpdating = True
    ' 删除所有图表的第一个图例项
    Call DeleteChartLegend
    
    MsgBox "节假日时段柱状图创建完成！", vbInformation
End Sub

' 新增函数：分析收益模式
Function AnalyzeProfitMode(configRange As Range) As String
    Dim profitMode As String
    Dim firstProfit As String
    Dim secondProfit As String
    Dim value1 As Integer, value2 As Integer, value As Integer
    Dim chargeType As String, dischargeType As String
    Dim legendSeries As series
    
    ' 分析第一次套利机会
    Dim i As Integer, j As Integer, k As Integer
    Dim chargeStartTime As Integer, chargeEndTime As Integer
    Dim dischargeStartTime As Integer, dischargeEndTime As Integer
    Dim peakLength As Integer, sharpPeakLength As Integer
    Dim foundFirstProfit As Boolean
    foundFirstProfit = False
    
    ' 第一次套利：寻找充电时段（需要连续2小时）
    Dim bestChargeStartTime As Integer, bestChargeEndTime As Integer
    Dim bestChargeType As String
    bestChargeStartTime = -1
    bestChargeEndTime = -1
    
    ' 先寻找连续2小时以上的低谷时段
    For i = 1 To 23
        value1 = configRange.Cells(1, i).value
        value2 = configRange.Cells(1, i + 1).value
        
        ' 检查是否是低谷时段（4=低谷、5=深谷）
        If (value1 = LOW Or value1 = DEEP_LOW) And (value2 = LOW Or value2 = DEEP_LOW) Then
            bestChargeStartTime = i
            bestChargeEndTime = i + 1
            bestChargeType = "谷"
            Exit For
        End If
    Next i
    
    ' 如果没有找到连续2小时的低谷时段，再寻找平段
    If bestChargeStartTime = -1 Then
        For i = 1 To 23
            value1 = configRange.Cells(1, i).value
            value2 = configRange.Cells(1, i + 1).value
            
            ' 检查是否是平段（3=平段）
            If value1 = NORMAL And value2 = NORMAL Then
                bestChargeStartTime = i
                bestChargeEndTime = i + 1
                bestChargeType = "平"
                Exit For
            End If
        Next i
    End If
    
    ' 如果找到了充电时段
    If bestChargeStartTime <> -1 Then
        chargeStartTime = bestChargeStartTime
        chargeEndTime = bestChargeEndTime
        
        ' 寻找放电时段（需要连续2小时）
        For j = chargeEndTime + 1 To 23
            ' 统计后续高峰和尖峰时段的长度
            peakLength = 0
            sharpPeakLength = 0
            For k = j To 24
                value = configRange.Cells(1, k).value
                If value = 2 Then ' 高峰
                    peakLength = peakLength + 1
                ElseIf value = 1 Then ' 尖峰
                    sharpPeakLength = sharpPeakLength + 1
                End If
            Next k
            
            ' 检查是否有连续2小时的放电时段
            value1 = configRange.Cells(1, j).value
            value2 = configRange.Cells(1, j + 1).value
            
            If IsDischargeTimeValue(value1) And IsDischargeTimeValue(value2) Then
                dischargeStartTime = j
                dischargeEndTime = j + 1
                
                ' 根据时段长度决定放电类型
                If sharpPeakLength >= peakLength Then
                    dischargeType = "尖"
                Else
                    dischargeType = "峰"
                End If
                
                firstProfit = "第一次：" & dischargeType & bestChargeType & "套利"
                foundFirstProfit = True
                Exit For
            End If
        Next j
    End If
        
    ' 第二次套利：如果找到第一次套利，继续寻找第二次套利机会
    If foundFirstProfit Then
        Dim foundSecondProfit As Boolean
        foundSecondProfit = False
        
        ' 从第一次放电结束后开始寻找第二次充电时段（需要连续2小时）
        For i = dischargeEndTime + 1 To 23
            value1 = configRange.Cells(1, i).value
            value2 = configRange.Cells(1, i + 1).value
            
            If IsChargeTimeValue(value1) And IsChargeTimeValue(value2) Then
                chargeStartTime = i
                chargeEndTime = i + 1
                
                ' 确定充电类型
                chargeType = GetChargeType(value1)
                
                ' 寻找第二次放电时段（只需要1小时）
                For j = chargeEndTime + 1 To 24
                    ' 统计后续高峰和尖峰时段的长度
                    peakLength = 0
                    sharpPeakLength = 0
                    For k = j To 24
                        value = configRange.Cells(1, k).value
                        If value = 2 Then ' 高峰
                            peakLength = peakLength + 1
                        ElseIf value = 1 Then ' 尖峰
                            sharpPeakLength = sharpPeakLength + 1
                        End If
                    Next k
                    
                    ' 检查是否有放电时段
                    value = configRange.Cells(1, j).value
                    If IsDischargeTimeValue(value) Then
                        ' 根据时段长度决定放电类型
                        If sharpPeakLength >= peakLength Then
                            dischargeType = "尖"
                        Else
                            dischargeType = "峰"
                        End If
                        
                        secondProfit = vbNewLine & "第二次：" & dischargeType & chargeType & "套利"
                        foundSecondProfit = True
                        Exit For
                    End If
                Next j
                If foundSecondProfit Then Exit For
            End If
        Next i
    End If
    
    ' 组合收益模式说明
    profitMode = ""
    If foundFirstProfit Then
        profitMode = firstProfit
        If foundSecondProfit Then
            profitMode = profitMode & secondProfit
        End If
    End If
    
    AnalyzeProfitMode = profitMode
End Function

' 辅助函数：判断是否是充电时段
Private Function IsChargeTimeValue(ByVal value As Integer) As Boolean
    ' 判断是否是充电时段（低谷=4、深谷=5或平段=3）
    Select Case value
        Case 3, 4, 5  ' 平段、低谷、深谷
            IsChargeTimeValue = True
        Case Else
            IsChargeTimeValue = False
    End Select
End Function

' 辅助函数：判断是否是放电时段
Private Function IsDischargeTimeValue(ByVal value As Integer) As Boolean
    ' 判断是否是放电时段（高峰=2或尖峰=1）
    Select Case value
        Case 1, 2  ' 尖峰、高峰
            IsDischargeTimeValue = True
        Case Else
            IsDischargeTimeValue = False
    End Select
End Function

' 辅助函数：获取充电类型
Private Function GetChargeType(ByVal value As Integer) As String
    ' 获取充电类型（"平"或"谷"）
    Select Case value
        Case 3  ' 平段
            GetChargeType = "平"
        Case Else  ' 低谷或深谷
            GetChargeType = "谷"
    End Select
End Function

' 填充柱状图数据源
Public Sub FillChartDataSource(ByVal wsHoliday As Worksheet)
    Dim i As Long, j As Long
    Dim timeType As Integer
    
    ' 清除现有数据
    wsHoliday.Range("B32:Y38").ClearContents
    
    ' 遍历每个节日
    For i = 1 To 7
        ' 遍历每个小时
        For j = 0 To 23
            timeType = wsHoliday.Cells(i + 16, j + 2).value
            
            ' 根据时段类型设置高度值
            Select Case timeType
                Case PEAK  ' 尖峰
                    wsHoliday.Cells(i + 31, j + 2).value = SHARP_PEAK_HEIGHT
                Case HIGH  ' 高峰
                    wsHoliday.Cells(i + 31, j + 2).value = PEAK_HEIGHT
                Case NORMAL  ' 平段
                    wsHoliday.Cells(i + 31, j + 2).value = FLAT_HEIGHT
                Case LOW  ' 低谷
                    wsHoliday.Cells(i + 31, j + 2).value = VALLEY_HEIGHT
                Case DEEP_LOW  ' 深谷
                    wsHoliday.Cells(i + 31, j + 2).value = DEEP_VALLEY_HEIGHT
                Case Else  ' 默认值
                    wsHoliday.Cells(i + 31, j + 2).value = 0
            End Select
            
            ' 输出调试信息
            Debug.Print "处理节日: " & wsHoliday.Cells(i + 1, 1).value
            Debug.Print "  小时 " & j & " 的时段类型: " & timeType
            Debug.Print "    设置" & GetTimeTypeName(timeType) & "高度: " & wsHoliday.Cells(i + 31, j + 2).value
        Next j
    Next i
    
    Debug.Print "完成柱状图数据源填充"
End Sub

' 获取时段类型名称
Private Function GetTimeTypeName(ByVal timeType As Integer) As String
    Select Case timeType
        Case PEAK
            GetTimeTypeName = "尖峰"
        Case HIGH
            GetTimeTypeName = "高峰"
        Case NORMAL
            GetTimeTypeName = "平段"
        Case LOW
            GetTimeTypeName = "低谷"
        Case DEEP_LOW
            GetTimeTypeName = "深谷"
        Case Else
            GetTimeTypeName = "默认"
    End Select
End Function

' 删除图表的第一个图例项
Private Sub DeleteChartLegend()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("国家节假日时段分析")
    
    ' 遍历所有图表
    Dim chartObj As ChartObject
    For Each chartObj In ws.ChartObjects
        ' 检查图表是否是柱状图
        If chartObj.Chart.ChartType = xlColumnStacked Then
            ' 检查图表是否有图例
            If chartObj.Chart.HasLegend Then
                ' 获取图例系列
                Dim series As series
                For Each series In chartObj.Chart.SeriesCollection
                    ' 检查系列是否是第一个图例系列
                    If series.index = 1 Then
                        ' 删除系列
                        series.Delete
                        Exit For
                    End If
                Next series
            End If
        End If
    Next chartObj
End Sub








