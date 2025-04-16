VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegionSelectorForm 
   Caption         =   "地区选择"
   ClientHeight    =   4800
   ClientWidth     =   4800
   OleObjectBlob   =   "RegionSelectorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegionSelectorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' 设置窗体标题
    Me.Caption = "选择要显示的地区"
    
    ' 设置列表框样式
    With RegionList
        .MultiSelect = False
        .Width = 200
        .Height = 200
    End With
    
    ' 获取解析结果工作表
    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("多个地区时段表")
    If wsSource Is Nothing Then
        MsgBox "未找到多个地区时段表！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 收集所有唯一的地区
    Dim uniqueRegions As Collection
    Set uniqueRegions = New Collection
    Dim cell As Range
    On Error Resume Next
    For Each cell In wsSource.Range("A2:A" & wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row)
        If cell.value <> "" Then
            uniqueRegions.Add cell.value, CStr(cell.value)
        End If
    Next cell
    On Error GoTo 0
    
    ' 将地区添加到列表框
    Dim item As Variant
    For Each item In uniqueRegions
        RegionList.AddItem item
    Next item
    
    ' 如果列表为空，显示提示
    If RegionList.ListCount = 0 Then
        MsgBox "未找到任何地区数据！", vbExclamation
        Unload Me
    End If
End Sub

Private Sub RegionList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' 双击列表项时应用筛选
    ApplyFilter
End Sub

Private Sub OKButton_Click()
    ' 如果没有选择地区，显示提示并退出
    If RegionList.ListIndex = -1 Then
        MsgBox "请选择一个地区！", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的地区
    Dim selectedRegion As String
    selectedRegion = RegionList.List(RegionList.ListIndex)
    
    ' 更新单个地区时段表
    Call UpdateSelectedRegion(selectedRegion)
    
    ' 关闭窗体
    Unload Me
End Sub

Private Sub CancelButton_Click()
    ' 点击取消按钮时关闭窗体
    Unload Me
End Sub

Private Sub ApplyFilter()
    ' 如果没有选择地区，显示提示并退出
    If RegionList.ListIndex = -1 Then
        MsgBox "请选择一个地区！", vbExclamation
        Exit Sub
    End If
    
    ' 获取选中的地区
    Dim selectedRegion As String
    selectedRegion = RegionList.List(RegionList.ListIndex)
    
    ' 更新单个地区时段表
    Call UpdateSelectedRegion(selectedRegion)
    
    ' 关闭窗体
    Unload Me
End Sub 