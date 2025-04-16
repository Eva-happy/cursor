Sub CreateDynamicNamedRangesForCities()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentProvince As String
    Dim currentCity As String
    Dim prevProvince As String
    Dim prevCity As String
    Dim provinceStartRow As Long
    Dim cityStartRow As Long
    Dim i As Long
    
    ' 设置工作表
    Set ws = ThisWorkbook.Sheets("中国所有省份城市县区")
    
    ' 获取最后一行
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 首先创建所有省份的列表
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="省份列表", _
        RefersTo:="=" & ws.Name & "!$A$2:$A$" & lastRow
    On Error GoTo 0
    
    ' 初始化变量
    provinceStartRow = 2
    cityStartRow = 2
    prevProvince = ""
    prevCity = ""
    
    ' 遍历每一行
    For i = 2 To lastRow + 1  ' 加1是为了处理最后一个省份
        currentProvince = ws.Cells(i, 1).Value
        currentCity = ws.Cells(i, 2).Value
        
        ' 如果省份变化了或到达最后一行
        If currentProvince <> prevProvince And prevProvince <> "" Or i = lastRow + 1 Then
            Dim endRow As Long
            endRow = i - 1
            
            ' 创建上一个省份对应的城市列表
            On Error Resume Next
            ThisWorkbook.Names.Add Name:=prevProvince, _
                RefersTo:="=" & ws.Name & "!$B$" & provinceStartRow & ":$B$" & endRow
            On Error GoTo 0
            
            provinceStartRow = i
            cityStartRow = i
        End If
        
        ' 如果城市变化了或到达最后一行
        If (currentCity <> prevCity And prevCity <> "") Or i = lastRow + 1 Then
            ' 创建上一个城市对应的县区列表
            On Error Resume Next
            ThisWorkbook.Names.Add Name:=prevProvince & "_" & prevCity, _
                RefersTo:="=" & ws.Name & "!$C$" & cityStartRow & ":$C$" & (i - 1)
            On Error GoTo 0
            
            cityStartRow = i
        End If
        
        prevProvince = currentProvince
        prevCity = currentCity
    Next i
    
    MsgBox "动态命名范围已创建完成！"
End Sub


