Sub ProcessTimeSlotWithDuplication()
    ' 创建集合来存储每种类型的时段
    Dim peakSlots As New Collection
    Dim highSlots As New Collection
    Dim normalSlots As New Collection
    Dim lowSlots As New Collection
    
    ' 首先收集所有时段信息
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("时段输入")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 收集所有时段信息
    Dim i As Long
    For i = 2 To lastRow
        Dim monthStr As String
        Dim timeType As String
        Dim timeRange As String
        
        monthStr = ws.Cells(i, 1).Value
        timeType = ws.Cells(i, 2).Value
        timeRange = ws.Cells(i, 3).Value
        
        If timeRange <> "" Then
            Dim slotInfo As String
            slotInfo = monthStr & "|" & timeRange
            
            Select Case timeType
                Case "低谷"
                    lowSlots.Add slotInfo
                Case "平段"
                    normalSlots.Add slotInfo
                Case "高峰"
                    highSlots.Add slotInfo
                Case "尖峰"
                    peakSlots.Add slotInfo
            End Select
        End If
    Next i
    
    ' 清除现有数据
    Dim configWs As Worksheet
    Set configWs = ThisWorkbook.Sheets("时段配置")
    configWs.Range("A2:D25").ClearContents
    
    ' 按优先级处理时段（低谷->平段->高峰->尖峰）
    ProcessTimeSlotCollection lowSlots, "低谷"
    ProcessTimeSlotCollection normalSlots, "平段"
    ProcessTimeSlotCollection highSlots, "高峰"
    ProcessTimeSlotCollection peakSlots, "尖峰"
End Sub

Sub ProcessTimeSlotCollection(ByVal slots As Collection, ByVal timeType As String)
    Dim configWs As Worksheet
    Set configWs = ThisWorkbook.Sheets("时段配置")
    
    Dim slot As Variant
    For Each slot In slots
        Dim parts As Variant
        parts = Split(slot, "|")
        
        Dim monthStr As String
        Dim timeRange As String
        monthStr = parts(0)
        timeRange = parts(1)
        
        ' 提取时间对
        Dim timePairs As Variant
        timePairs = ExtractTimePairs(timeRange)
        
        ' 获取月份对应的列
        Dim col As Long
        Select Case monthStr
            Case "七月"
                col = 2
            Case "八月"
                col = 3
            Case "九月"
                col = 4
        End Select
        
        ' 处理每个时间对
        Dim j As Long
        For j = 0 To UBound(timePairs) Step 2
            Dim startTime As String
            Dim endTime As String
            startTime = timePairs(j)
            endTime = timePairs(j + 1)
            
            ' 转换时间为小时
            Dim startHour As Integer
            Dim endHour As Integer
            startHour = CInt(Split(startTime, ":")(0))
            endHour = CInt(Split(endTime, ":")(0))
            
            ' 写入配置
            Dim row As Long
            row = startHour + 2
            
            ' 根据时段类型设置值
            Dim value As Integer
            Select Case timeType
                Case "低谷"
                    value = 4
                Case "平段"
                    value = 3
                Case "高峰"
                    value = 2
                Case "尖峰"
                    value = 1
            End Select
            
            ' 记录调试信息
            Debug.Print "处理时段: " & monthStr & " " & timeType & " " & startTime & "-" & endTime
            
            ' 写入值
            configWs.Cells(row, col).value = value
        Next j
    Next slot
End Sub 