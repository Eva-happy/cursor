Option Explicit

Private Sub UserForm_Initialize()
    ' 获取工作表
    Dim ws As Worksheet
    
    ' 获取解析结果工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("解析国家节假日文字")
    If ws Is Nothing Then
        MsgBox "未找到'解析国家节假日文字'工作表！", vbExclamation
        Me.Hide
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 获取最后一行
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 清除现有列表
    HolidayRegionList.Clear
    
    ' 使用Collection来存储唯一的地区名称
    Dim uniqueRegions As Collection
    Set uniqueRegions = New Collection
    
    ' 获取A列所有数据并去重
    On Error Resume Next
    Dim i As Long
    For i = 2 To lastRow
        Dim regionName As String
        regionName = Trim(ws.Cells(i, 1).value)
        If regionName <> "" Then
            uniqueRegions.Add regionName, regionName
        End If
    Next i
    On Error GoTo 0
    
    ' 将去重后的地区名称添加到ListBox
    Dim region As Variant
    For Each region In uniqueRegions
        HolidayRegionList.AddItem region
    Next region
    
    ' 获取当前选中的地区
    Dim wsHoliday As Worksheet
    Set wsHoliday = ThisWorkbook.Worksheets("国家节假日时段分析")
    
    ' 如果有当前选中的地区，则在列表中选中它
    If Not wsHoliday Is Nothing Then
        Dim currentRegion As String
        currentRegion = wsHoliday.Range("B1").value
        If currentRegion <> "" Then
            Dim j As Long
            For j = 0 To HolidayRegionList.ListCount - 1
                If HolidayRegionList.List(j) = currentRegion Then
                    HolidayRegionList.Selected(j) = True
                    Exit For
                End If
            Next j
        End If
    End If
End Sub

Private Sub OKButton_Click()
    ' 如果没有选择地区，显示提示并退出
    If HolidayRegionList.ListIndex = -1 Then
        MsgBox "请选择一个地区！", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的地区
    Dim selectedRegion As String
    selectedRegion = HolidayRegionList.value
    Debug.Print "选择的地区: " & selectedRegion
    
    ' 获取必要的工作表
    Dim wsHoliday As Worksheet
    Dim wsSource As Worksheet
    Set wsHoliday = ThisWorkbook.Worksheets("国家节假日时段分析")
    Set wsSource = ThisWorkbook.Worksheets("解析国家节假日文字")
    
    ' 检查工作表是否存在
    If wsHoliday Is Nothing Then
        MsgBox "未找到'国家节假日时段分析'工作表！", vbExclamation
        Exit Sub
    End If
    If wsSource Is Nothing Then
        MsgBox "未找到'解析国家节假日文字'工作表！", vbExclamation
        Exit Sub
    End If
    
    ' 更新标题
    wsHoliday.Cells(1, 1).value = selectedRegion
    
    ' 清除现有数据
    wsHoliday.Range("B2:Y8").ClearContents
    wsHoliday.Range("B17:Y23").ClearContents
    wsHoliday.Range("B2:Y8").Interior.ColorIndex = xlNone
    wsHoliday.Range("B32:Y38").ClearContents ' 清除柱状图数据源行
    
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
        ' 同时确保柱状图数据源区域也有正确的节日名称
        If wsHoliday.Cells(i + 32, 1).value <> holidays(i) Then
            wsHoliday.Cells(i + 32, 1).value = holidays(i)
            Debug.Print "修正柱状图数据源节日名称: " & holidays(i)
        End If
    Next i
    
    ' 从解析结果中获取数据并填充
    Call HolidayModule.FillHolidayTimeData(selectedRegion, wsHoliday, wsSource)
    
    ' 更新颜色
    Call HolidayModule.UpdateHolidayTimeTable
    
    ' 填充柱状图数据源表
    Debug.Print "开始填充柱状图数据源"
    
    ' 遍历每个节日
    For i = 0 To 6 ' 7个节日
        ' 获取节日名称
        Dim holidayName As String
        holidayName = wsHoliday.Cells(i + 32, 1).value
        Debug.Print "处理节日: " & holidayName
        
        ' 遍历每个小时
        Dim j As Long
        For j = 0 To 23
            Dim timeValue As Integer
            timeValue = wsHoliday.Cells(i + 17, j + 2).value ' 从配置区域获取时段类型
            Debug.Print "  小时 " & j & " 的时段类型: " & timeValue
            
            ' 设置高度值
            Select Case timeValue
                Case 1 ' 尖峰
                    wsHoliday.Cells(i + 32, j + 2).value = SHARP_PEAK_HEIGHT ' 1.2
                    Debug.Print "    设置尖峰高度: " & SHARP_PEAK_HEIGHT
                Case 2 ' 高峰
                    wsHoliday.Cells(i + 32, j + 2).value = PEAK_HEIGHT ' 0.9
                    Debug.Print "    设置高峰高度: " & PEAK_HEIGHT
                Case 3 ' 平段
                    wsHoliday.Cells(i + 32, j + 2).value = FLAT_HEIGHT ' 0.6
                    Debug.Print "    设置平段高度: " & FLAT_HEIGHT
                Case 4 ' 低谷
                    wsHoliday.Cells(i + 32, j + 2).value = VALLEY_HEIGHT ' 0.3
                    Debug.Print "    设置低谷高度: " & VALLEY_HEIGHT
                Case 5 ' 深谷
                    wsHoliday.Cells(i + 32, j + 2).value = DEEP_VALLEY_HEIGHT ' 0.1
                    Debug.Print "    设置深谷高度: " & DEEP_VALLEY_HEIGHT
                Case Else
                    wsHoliday.Cells(i + 32, j + 2).value = 0
                    Debug.Print "    设置默认高度: 0"
            End Select
            
            ' 设置数值格式为一位小数
            wsHoliday.Cells(i + 32, j + 2).NumberFormat = "0.0"
        Next j
    Next i
    
    Debug.Print "完成柱状图数据源填充"
    Unload Me
End Sub

Private Sub CancelButton_Click()
    ' 点击取消按钮时关闭窗体
    Unload Me
End Sub

Private Sub HolidayRegionList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' 双击列表项时执行确认操作
    Call OKButton_Click
End Sub

