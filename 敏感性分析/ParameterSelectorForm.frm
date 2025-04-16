Option Explicit

' 添加对SensitivityAnalysis模块的引用
Private Sub UserForm_Initialize()
    ' 初始化参数列表
    Dim baseSheet As Worksheet
    Dim cell As Range
    Dim lastRow As Long
    
    ' 获取基础参数表
    Set baseSheet = ThisWorkbook.Sheets("基础参数及输出结果表")
    
    ' 清空列表框
    lstParameters.Clear
    
    ' 从Q4开始读取参数列表
    lastRow = baseSheet.Cells(baseSheet.Rows.Count, "Q").End(xlUp).row
    For Each cell In baseSheet.Range("Q4:Q4" & lastRow)
        If cell.Value <> "" Then
            lstParameters.AddItem cell.Value
        End If
    Next cell
End Sub

Private Sub cmdOK_Click()
    ' 确认按钮点击事件
    If lstParameters.ListIndex = -1 Then
        MsgBox "请选择一个参数！", vbExclamation
        Exit Sub
    End If
    
    ' 调用计算函数
    Call CalculateSingleParameterAnalysis(lstParameters.Value)
    
    ' 提示用户下一步操作
    MsgBox "请在B11单元格开始输入变动后的值，然后点击'运行结果'按钮进行计算。", vbInformation
    
    ' 关闭窗体
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    ' 取消按钮点击事件
    Unload Me
End Sub


