Option Explicit

Sub 规划求解()
    ' 声明变量
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim changeCell As Range
    Dim absoluteValueCell As Range
    Dim sumRange As Range
    
    ' 设置工作表引用
    Set ws = ThisWorkbook.Sheets("光伏收益测算表")
    Set targetCell = ws.Range("C153")
    Set sumRange = ws.Range("I150:AB150")  ' 添加求和范围的引用
    Set changeCell = ws.Range("C152")
    Set absoluteValueCell = ws.Range("C154")
    
    ' 确保目标单元格包含正确的求和公式
    targetCell.Formula = "=SUM(" & sumRange.Address & ")"
    
    ' 保存原始值
    Dim originalValue As Double
    originalValue = changeCell.Value
    
    ' 获取目标值（绝对值）
    Dim targetValue As Double
    targetValue = Abs(absoluteValueCell.Value)
    
    ' 开启调试信息
    Debug.Print "开始规划求解"
    Debug.Print "----------------------------------------"
    Debug.Print "目标单元格信息："
    Debug.Print "地址: " & targetCell.Address
    Debug.Print "公式: " & targetCell.Formula
    Debug.Print "当前值: " & targetCell.Value
    Debug.Print "目标值: " & targetValue
    
    ' 设置自动计算
    Application.Calculation = xlCalculationAutomatic
    
    ' 清除现有的Solver设置
    SolverReset
    
    ' 配置Solver参数
    On Error Resume Next
    
    ' 设置目标单元格和目标值
    SolverOk SetCell:=targetCell.Address(False, False), _
             MaxMinVal:=3, _
             ValueOf:=targetValue, _
             ByChange:=changeCell.Address(False, False)
             
    If Err.Number <> 0 Then
        MsgBox "设置Solver参数时出错：" & Err.Description, vbCritical
        Exit Sub
    End If
    
    ' 添加约束条件
    SolverAdd CellRef:=changeCell.Address(False, False), _
              Relation:=3, _
              FormulaText:="0"  ' 确保变化单元格大于0
              
    ' 设置Solver选项
    SolverOptions MaxTime:=100, _
                  Iterations:=100, _
                  Precision:=0.000001, _
                  Convergence:=0.0001, _
                  StepThru:=False, _
                  Scaling:=True, _
                  AssumeNonNeg:=True
    
    ' 运行Solver
    Dim solverResult As Integer
    solverResult = SolverSolve(UserFinish:=False)
    
    ' 如果求解成功，保存结果
    If solverResult = 0 Then
        ' 获取最终值
        Dim finalValue As Double
        finalValue = changeCell.Value
        
        ' 显示结果
        MsgBox "规划求解完成！" & vbCrLf & _
               "目标单元格当前值: " & Format(targetCell.Value, "#,##0.00") & vbCrLf & _
               "找到的折现率: " & Format(finalValue * 100, "0.00") & "%", vbInformation
               
        ' 生成报告
        On Error Resume Next
        SolverSolve UserFinish:=True
        SolverSave SaveArea:=Range("A1")  ' 保存结果到工作表
        On Error GoTo 0
        
        ' 打印详细信息到即时窗口
        Debug.Print "----------------------------------------"
        Debug.Print "求解成功："
        Debug.Print "目标单元格值: " & targetCell.Value
        Debug.Print "找到的折现率: " & finalValue * 100 & "%"
        Debug.Print "原始目标值: " & targetValue
    Else
        ' 处理错误情况
        Select Case solverResult
            Case 1
                MsgBox "所有变量都在约束范围内。", vbInformation
            Case 2
                MsgBox "达到最大迭代次数限制。", vbExclamation
            Case 3
                MsgBox "求解器无法改进当前解。", vbExclamation
            Case 4
                MsgBox "求解器无法找到可行解。", vbCritical
            Case 5
                MsgBox "用户取消了求解过程。", vbInformation
            Case 6
                MsgBox "达到了时间限制。", vbExclamation
            Case 7
                MsgBox "求解器引擎内存不足。", vbCritical
            Case 8
                MsgBox "用户终止了求解过程。", vbInformation
            Case 9
                MsgBox "其他求解器错误。", vbCritical
        End Select
    End If
    
    ' 恢复原始数据（如果需要）
    'changeCell.Value = originalValue
    
    ' 确保保持自动计算
    Application.Calculation = xlCalculationAutomatic
    
    Debug.Print "----------------------------------------"
    Debug.Print "规划求解完成"
End Sub 