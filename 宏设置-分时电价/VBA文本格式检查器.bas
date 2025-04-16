Option Explicit

Public Function CheckTextFormat() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentRegion As String
    Dim line As String
    Dim hasError As Boolean
    Dim errorMsg As String
    Dim allErrorMsgs As String  ' 新增：存储所有错误信息
    Dim monthPattern As String
    Dim timeSlotPattern As String
    Dim timePattern As String
    Dim regEx As Object
    Dim hasTimeSlot As Boolean
    Dim lastRegion As String
    Dim regionTimeSlots As Collection
    
    ' 初始化正则表达式对象
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' 设置月份格式的正则表达式
    monthPattern = "(?:月份：\s*)?(?:[(（\[【])?([0-9]+(?:-[0-9]+)?(?:[,，、][0-9]+(?:-[0-9]+)?)*)[月](?:[)）\]】])?|([0-9]+(?:-[0-9]+)?(?:[,，、][0-9]+(?:-[0-9]+)?)*)[月]"
    
    ' 设置时段类型的正则表达式
    timeSlotPattern = "^(尖峰时段：|高峰时段：|平段时段：|深谷时段：|低谷时段：|平时段：|平段：).*$"
    
    ' 设置时间格式的正则表达式（支持更多分隔符和秒）
    timePattern = "(?:[0-2]?[0-9]|24)[:：][0-5][0-9](?:[:：][0-5][0-9])?[-—－–─━~至到](?:次日)?(?:[0-2]?[0-9]|24)[:：][0-5][0-9](?:[:：][0-5][0-9])?"
    
    ' 获取工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("分时电价政策表")
    If ws Is Nothing Then
        MsgBox "未找到'分时电价政策表'工作表！", vbCritical
        CheckTextFormat = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' 获取最后一行
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "工作表中没有数据！", vbExclamation
        CheckTextFormat = False
        Exit Function
    End If
    
    ' 初始化变量
    hasError = False
    allErrorMsgs = ""  ' 初始化错误信息集合
    Set regionTimeSlots = New Collection
    
    ' 遍历每一行
    For i = 2 To lastRow
        currentRegion = Trim(ws.Cells(i, "A").value)
        line = Trim(ws.Cells(i, "B").value)
        errorMsg = ""  ' 重置每行的错误信息
        
        ' 检查是否填写了地区
        If currentRegion = "" And line <> "" Then
            AddError allErrorMsgs, "第 " & i & " 行：未填写地区名称"
            hasError = True
            GoTo NextLine
        End If
        
        ' 如果是新的地区，检查上一个地区的时段完整性
        If currentRegion <> lastRegion And lastRegion <> "" Then
            If Not CheckTimeSlotCompleteness(regionTimeSlots, lastRegion, errorMsg) Then
                AddError allErrorMsgs, errorMsg
                hasError = True
            End If
            Set regionTimeSlots = New Collection
        End If
        
        ' 跳过空行
        If line = "" Then
            GoTo NextLine
        End If
        
        ' 标准化格式
        line = StandardizeFormat(line)
        
        ' 检查月份格式
        regEx.Pattern = monthPattern
        If InStr(line, "月") > 0 And Not InStr(line, "时段") > 0 Then
            If Not regEx.Test(line) Then
                AddError allErrorMsgs, "第 " & i & " 行：月份格式错误。请使用以下格式之一：" & vbCrLf & _
                                     "1. 月份：1-3月" & vbCrLf & _
                                     "2. (1-3)月" & vbCrLf & _
                                     "3. 1、2、3月"
                hasError = True
            End If
            
            ' 检查月份数值范围
            If Not CheckMonthRange(line) Then
                AddError allErrorMsgs, "第 " & i & " 行：月份数值超出范围（1-12）"
                hasError = True
            End If
        End If
        
        ' 检查时段格式
        regEx.Pattern = timeSlotPattern
        ' 检查是否包含任何可能的时段相关关键词
        If InStr(line, "峰") > 0 Or InStr(line, "谷") > 0 Or InStr(line, "平") > 0 Then
            If Not regEx.Test(line) Then
                AddError allErrorMsgs, "第 " & i & " 行：时段格式错误。必须严格使用以下标准格式之一（不允许使用其他变体，如'高峰段'、'高峰电价'、'高峰'等）：" & vbCrLf & _
                                     "1. 尖峰时段：" & vbCrLf & _
                                     "2. 高峰时段：" & vbCrLf & _
                                     "3. 平段时段：" & vbCrLf & _
                                     "4. 深谷时段：" & vbCrLf & _
                                     "5. 低谷时段：" & vbCrLf & _
                                     "6. 平时段：" & vbCrLf & _
                                     "7. 平段："
                hasError = True
            Else
                ' 记录时段类型
                Dim timeSlotType As String
                If InStr(line, "时段") > 0 Then
                    timeSlotType = Left(line, InStr(line, "时段") - 1)
                Else
                    timeSlotType = Left(line, InStr(line, "：") - 1)
                End If
                timeSlotType = StandardizeTimeSlotType(timeSlotType)
                On Error Resume Next
                regionTimeSlots.Add timeSlotType, timeSlotType
                On Error GoTo 0
                
                ' 检查时间格式
                regEx.Pattern = timePattern
                If Not regEx.Test(line) Then
                    AddError allErrorMsgs, "第 " & i & " 行：时间格式错误。应使用24小时制，如：8:00-11:00"
                    hasError = True
                End If
                
                ' 检查时间顺序
                If Not CheckTimeOrder(line) Then
                    AddError allErrorMsgs, "第 " & i & " 行：时间顺序错误，早的时间应在前面"
                    hasError = True
                End If
            End If
        End If
        
        lastRegion = currentRegion
NextLine:
    Next i
    
    ' 检查最后一个地区的时段完整性
    If Not hasError And lastRegion <> "" Then
        If Not CheckTimeSlotCompleteness(regionTimeSlots, lastRegion, errorMsg) Then
            AddError allErrorMsgs, errorMsg
            hasError = True
        End If
    End If
    
    ' 显示结果
    If hasError Then
        MsgBox "发现以下格式错误：" & vbCrLf & vbCrLf & allErrorMsgs, vbCritical, "格式检查失败"
        CheckTextFormat = False
    Else
        MsgBox "文本格式检查通过！请返回【分时电价操作步骤表】并点击'解析文字'按钮继续。", vbInformation, "格式检查成功"
        CheckTextFormat = True
    End If
End Function

Private Sub AddError(ByRef allErrorMsgs As String, ByVal newError As String)
    If allErrorMsgs <> "" Then allErrorMsgs = allErrorMsgs & vbCrLf & vbCrLf
    allErrorMsgs = allErrorMsgs & newError
End Sub

Private Function StandardizeFormat(ByVal text As String) As String
    ' 标准化分隔符和连接符
    
    ' 1. 统一分隔符为顿号（、）
    text = Replace(text, ",", "、")  ' 英文逗号
    text = Replace(text, "，", "、") ' 中文逗号
    text = Replace(text, ";", "、")  ' 英文分号
    text = Replace(text, "；", "、") ' 中文分号
    
    ' 2. 统一冒号为中文冒号（：）
    text = Replace(text, ":", "：")  ' 英文冒号转中文冒号
    
    ' 3. 统一连接符为短横线（-）
    text = Replace(text, "—", "-")   ' 中文破折号
    text = Replace(text, "－", "-")  ' 全角减号
    text = Replace(text, "–", "-")   ' 短破折号
    text = Replace(text, "─", "-")   ' 水平线
    text = Replace(text, "━", "-")   ' 粗水平线
    text = Replace(text, "~", "-")   ' 波浪号
    text = Replace(text, "至", "-")  ' "至"字
    text = Replace(text, "到", "-")  ' "到"字
    
    ' 4. 移除多余的空格
    text = Replace(text, " ", "")    ' 移除所有空格
    
    StandardizeFormat = text
End Function

Private Function CheckMonthRange(ByVal text As String) As Boolean
    Dim numbers() As String
    Dim num As Variant
    Dim monthText As String
    
    ' 提取数字部分
    monthText = text
    monthText = Replace(monthText, "月份：", "")
    monthText = Replace(monthText, "月", "")
    monthText = Replace(monthText, "(", "")
    monthText = Replace(monthText, ")", "")
    monthText = Replace(monthText, "（", "")
    monthText = Replace(monthText, "）", "")
    monthText = Replace(monthText, "[", "")
    monthText = Replace(monthText, "]", "")
    monthText = Replace(monthText, "【", "")
    monthText = Replace(monthText, "】", "")
    
    ' 分割数字
    numbers = Split(monthText, "、")
    
    ' 检查每个数字
    For Each num In numbers
        If InStr(num, "-") > 0 Then
            Dim rangeParts() As String
            rangeParts = Split(num, "-")
            If UBound(rangeParts) = 1 Then
                If Not (IsNumeric(rangeParts(0)) And IsNumeric(rangeParts(1))) Then
                    CheckMonthRange = False
                    Exit Function
                End If
                If CLng(rangeParts(0)) < 1 Or CLng(rangeParts(0)) > 12 Or _
                   CLng(rangeParts(1)) < 1 Or CLng(rangeParts(1)) > 12 Then
                    CheckMonthRange = False
                    Exit Function
                End If
            End If
        Else
            If Not IsNumeric(num) Then
                CheckMonthRange = False
                Exit Function
            End If
            If CLng(num) < 1 Or CLng(num) > 12 Then
                CheckMonthRange = False
                Exit Function
            End If
        End If
    Next num
    
    CheckMonthRange = True
End Function

Private Function CheckTimeOrder(ByVal text As String) As Boolean
    Dim timeRanges() As String
    Dim timeParts() As String
    Dim colonPos As Long
    Dim startTime As Double
    Dim endTime As Double
    Dim timeRange As Variant
    Dim hourPart As String
    Dim minutePart As String
    Dim lastEndTime As Double
    Dim compareStartTime As Double  ' 新增：用于比较的开始时间
    
    ' 提取时间部分
    text = Mid(text, InStr(text, "：") + 1)
    timeRanges = Split(text, "、")
    
    ' 初始化上一个时间段的结束时间
    lastEndTime = -1
    
    For Each timeRange In timeRanges
        ' 标准化连接符为短横线
        Dim standardizedTime As String
        standardizedTime = timeRange
        standardizedTime = Replace(standardizedTime, "—", "-")
        standardizedTime = Replace(standardizedTime, "－", "-")
        standardizedTime = Replace(standardizedTime, "–", "-")
        standardizedTime = Replace(standardizedTime, "─", "-")
        standardizedTime = Replace(standardizedTime, "━", "-")
        standardizedTime = Replace(standardizedTime, "~", "-")
        standardizedTime = Replace(standardizedTime, "至", "-")
        standardizedTime = Replace(standardizedTime, "到", "-")
        
        timeParts = Split(standardizedTime, "-")
        
        If UBound(timeParts) = 1 Then
            ' 处理开始时间
            colonPos = InStr(timeParts(0), ":")
            If colonPos = 0 Then colonPos = InStr(timeParts(0), "：")
            
            hourPart = Trim(Left(timeParts(0), colonPos - 1))
            minutePart = Trim(Mid(timeParts(0), colonPos + 1))
            ' 如果包含秒，忽略秒的部分
            If InStr(minutePart, ":") > 0 Or InStr(minutePart, "：") > 0 Then
                minutePart = Left(minutePart, InStr(minutePart, IIf(InStr(minutePart, ":") > 0, ":", "：")) - 1)
            End If
            
            If IsNumeric(hourPart) And IsNumeric(minutePart) Then
                ' 处理24:00的情况，将其视为0:00
                If Val(hourPart) = 24 And Val(minutePart) = 0 Then
                    startTime = 0
                    compareStartTime = 0  ' 用于比较的时间也设为0
                Else
                    startTime = Val(hourPart) + Val(minutePart) / 60
                    compareStartTime = startTime
                End If
            Else
                CheckTimeOrder = False
                Exit Function
            End If
            
            ' 处理结束时间
            Dim endTimeStr As String
            endTimeStr = timeParts(1)
            
            If InStr(endTimeStr, "次日") > 0 Then
                endTimeStr = Replace(endTimeStr, "次日", "")
                colonPos = InStr(endTimeStr, ":")
                If colonPos = 0 Then colonPos = InStr(endTimeStr, "：")
                
                hourPart = Trim(Left(endTimeStr, colonPos - 1))
                minutePart = Trim(Mid(endTimeStr, colonPos + 1))
                ' 如果包含秒，忽略秒的部分
                If InStr(minutePart, ":") > 0 Or InStr(minutePart, "：") > 0 Then
                    minutePart = Left(minutePart, InStr(minutePart, IIf(InStr(minutePart, ":") > 0, ":", "：")) - 1)
                End If
                
                If IsNumeric(hourPart) And IsNumeric(minutePart) Then
                    ' 处理24:00的情况
                    If Val(hourPart) = 24 And Val(minutePart) = 0 Then
                        endTime = 48  ' 次日24:00等于48:00
                    Else
                        endTime = Val(hourPart) + 24 + Val(minutePart) / 60
                    End If
                Else
                    CheckTimeOrder = False
                    Exit Function
                End If
            Else
                colonPos = InStr(endTimeStr, ":")
                If colonPos = 0 Then colonPos = InStr(endTimeStr, "：")
                
                hourPart = Trim(Left(endTimeStr, colonPos - 1))
                minutePart = Trim(Mid(endTimeStr, colonPos + 1))
                ' 如果包含秒，忽略秒的部分
                If InStr(minutePart, ":") > 0 Or InStr(minutePart, "：") > 0 Then
                    minutePart = Left(minutePart, InStr(minutePart, IIf(InStr(minutePart, ":") > 0, ":", "：")) - 1)
                End If
                
                If IsNumeric(hourPart) And IsNumeric(minutePart) Then
                    ' 处理24:00的情况
                    If Val(hourPart) = 24 And Val(minutePart) = 0 Then
                        endTime = 24
                    Else
                        endTime = Val(hourPart) + Val(minutePart) / 60
                    End If
                    ' 如果结束时间小于开始时间，自动视为次日
                    If endTime < startTime Then
                        endTime = endTime + 24
                    End If
                Else
                    CheckTimeOrder = False
                    Exit Function
                End If
            End If
            
            ' 检查单个时间段内的开始时间和结束时间顺序
            If startTime >= endTime Then
                CheckTimeOrder = False
                Exit Function
            End If
            
            ' 检查与上一个时间段的顺序
            If lastEndTime <> -1 Then
                ' 如果当前时间段以24:00开始，需要特殊处理
                If InStr(timeParts(0), "24") > 0 Then
                    compareStartTime = 24  ' 将24:00视为一天的结束
                End If
                
                If compareStartTime <= lastEndTime Then
                    CheckTimeOrder = False
                    Exit Function
                End If
            End If
            
            ' 更新上一个时间段的结束时间
            lastEndTime = endTime
        End If
    Next
    
    CheckTimeOrder = True
End Function

Private Function StandardizeTimeSlotType(ByVal timeSlotType As String) As String
    Select Case timeSlotType
        Case "平时段", "平段"
            StandardizeTimeSlotType = "平段时段"
        Case Else
            StandardizeTimeSlotType = timeSlotType & "时段"
    End Select
End Function

Private Function CheckTimeSlotCompleteness(ByVal timeSlots As Collection, ByVal region As String, ByRef errorMsg As String) As Boolean
    ' 只要有一个时段类型就可以了
    If timeSlots.Count > 0 Then
        CheckTimeSlotCompleteness = True
    Else
        errorMsg = "地区 """ & region & """ 没有任何时段类型"
        CheckTimeSlotCompleteness = False
    End If
End Function







