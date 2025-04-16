Option Explicit

Function IsSolverAvailable() As Boolean
    ' 检查Solver加载项是否已安装
    On Error Resume Next
    Application.Run "SOLVER.XLAM!SolverReset"
    If Err.Number = 0 Then
        IsSolverAvailable = True
        Exit Function
    End If
    
    ' 显示详细的启用指导
    MsgBox "请确保已正确添加SOLVER.XLAM：" & vbNewLine & _
           "1. 在VBA编辑器中展开'引用'文件夹" & vbNewLine & _
           "2. 确认已显示'引用 SOLVER.XLAM'" & vbNewLine & _
           "3. 如果没有显示，请先保存并关闭Excel" & vbNewLine & _
           "4. 重新打开Excel后再试", vbInformation
           
    IsSolverAvailable = False
End Function

Function GetTimeFactors() As Double()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim timeFactors() As Double
    Dim i As Long
    Dim startYear As Long
    
    Set ws1 = ThisWorkbook.Worksheets("光伏收益测算表")
    Set ws2 = ThisWorkbook.Worksheets("基础参数及输出结果表")
    
    ' 计算起始年份（第几年末）
    startYear = ws2.Range("B20").Value - ws2.Range("J26").Value
    
    ReDim timeFactors(1 To 20)
    
    ' 从光伏收益测算表137行读取时间系数，从I列开始（第6年）
    For i = 1 To 20
        timeFactors(i) = ws1.Range("I137").Offset(0, i - 1).Value
    Next i
    
    GetTimeFactors = timeFactors
End Function

Function FindNegativeValueInRow26() As Double
    Dim ws1 As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    Set ws1 = ThisWorkbook.Worksheets("光伏收益测算表")
    ' 从第26行E列开始查找
    Set rng = ws1.Range("E26:AB26")  ' 假设最多到AB列
    
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value < 0 Then
                FindNegativeValueInRow26 = cell.Value
                Exit Function
            End If
        End If
    Next cell
    
    ' 如果没有找到负数，返回0
    FindNegativeValueInRow26 = 0
End Function

Sub CalcIRRWithSolver()
    Dim ws1 As Worksheet, ws2 As Worksheet, wsSolver As Worksheet
    Dim i As Long
    Dim targetValue As Double
    Dim cashFlows(1 To 20) As Double
    
    ' 设置工作表引用
    Set ws1 = ThisWorkbook.Worksheets("光伏收益测算表")
    Set ws2 = ThisWorkbook.Worksheets("基础参数及输出结果表")
    
    ' 获取目标值（第26行E列开始的负数值）
    targetValue = Abs(FindNegativeValueInRow26())  ' 取绝对值，因为我们要用正数计算
    
    If targetValue = 0 Then
        MsgBox "未在第26行E列之后找到负数值！请检查数据。"
        Exit Sub
    End If
    
    ' 创建或获取求解工作表
    On Error Resume Next
    Set wsSolver = ThisWorkbook.Worksheets("求解临时表")
    If wsSolver Is Nothing Then
        Set wsSolver = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSolver.Name = "求解临时表"
    End If
    On Error GoTo 0
    
    ' 清除求解工作表内容
    wsSolver.Cells.Clear
    
    ' 设置标题和初始值
    With wsSolver
        .Range("A1").Value = "IRR"  ' 折现率
        .Range("A2").Value = 0.1    ' 初始值10%
        .Range("A2").NumberFormat = "0.00%"
        
        .Range("B1").Value = "计算结果"
        .Range("C1").Value = "目标值"
        .Range("C2").Value = targetValue
    End With
    
    ' 获取现金流数据
    For i = 1 To 20
        cashFlows(i) = ws1.Range("I26").Offset(0, i - 1).Value
    Next i
    
    ' 在工作表中写入现金流数据
    For i = 1 To 20
        wsSolver.Cells(i + 2, 1).Value = cashFlows(i)
        wsSolver.Cells(i + 2, 2).Formula = "=A" & (i + 2) & "/(1+$A$2)^" & i
    Next i
    
    ' 设置求和公式
    wsSolver.Range("B2").Formula = "=SUM(B3:B22)"
    
    ' 重置Solver
    SolverReset
    
    ' 设置求解参数
    SolverOk SetCell:=wsSolver.Range("B2"), _
             MaxMinVal:=3, _
             ValueOf:=targetValue, _
             ByChange:=wsSolver.Range("A2")
    
    ' 添加约束条件
    SolverAdd CellRef:=wsSolver.Range("A2"), Relation:=1, FormulaText:="0"  ' IRR > 0
    
    ' 设置求解选项
    SolverOptions Precision:=0.000001, _
                  Iterations:=1000, _
                  Convergence:=0.0001
    
    ' 执行求解
    Dim solverResult As Integer
    solverResult = SolverSolve(True)
    
    ' 保存求解结果
    If solverResult = 0 Then  ' 0表示找到解
        SolverFinish KeepFinal:=1
        
        ' 显示结果
        MsgBox "IRR计算完成！" & vbNewLine & _
               "IRR = " & Format(wsSolver.Range("A2").Value, "0.00%") & vbNewLine & _
               "计算结果 = " & wsSolver.Range("B2").Value & vbNewLine & _
               "目标值 = " & targetValue
    Else
        MsgBox "求解失败，请检查数据和约束条件！", vbExclamation
    End If
    
    ' 激活求解工作表
    wsSolver.Activate
    
End Sub


