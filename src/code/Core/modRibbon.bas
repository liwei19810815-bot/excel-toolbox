Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - 全部 Ribbon 回调集中在此
'
' 约定：回调只负责解析控件参数，然后转派给 modAction.RunAction。
'       任何业务逻辑都不许写在这里，也不许让 Ribbon 直接调业务过程——
'       否则该操作就绕过了高速模式与撤销框架。
'
' IRibbonUI 指针持久化：
'   VBA 工程一旦被重置（出错后点"结束"、改代码等），模块级对象变量就被清空，
'   此后再调用 gRibbon.Invalidate 会直接崩 Excel。标准解法是在 OnLoad 时把
'   ObjPtr 存进工作簿的 Name 里，用的时候再 CopyMemory 还原成对象引用。
'==============================================================================
Option Explicit
' 注意：本模块【不能】加 Option Private Module。
' Ribbon 回调由 Excel 通过与 Application.Run 相同的机制按名字查找，
' 私有模块里的过程找不到，表现为按钮点了没反应或提示"找不到宏"。

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const RIBBON_PTR_NAME As String = "ExcelToolbox_RibbonPtr"

Private mRibbon As IRibbonUI

'==============================================================================
' 回调
'==============================================================================

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    StoreRibbonPointer ribbon
End Sub

Public Sub Ribbon_OnAction(control As IRibbonControl)
    modAction.RunAction control.Tag
End Sub

Public Sub Ribbon_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modAction.IsActionEnabled(control.Tag)
End Sub

Public Sub Ribbon_GetLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modAction.ActionLabel(control.Tag)
End Sub

Public Sub Ribbon_GetScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modAction.ActionScreentip(control.Tag)
End Sub

' 开关型按钮（聚光灯）。pressed 参数是 Ribbon 自己算好的新状态，这里不用它——
' 真实状态以模块里的开关为准，按钮只负责显示。
Public Sub Ribbon_OnToggle(control As IRibbonControl, ByVal pressed As Boolean)
    modAction.RunAction control.Tag
    RefreshControl control.Id
End Sub

Public Sub Ribbon_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = modAction.IsActionPressed(control.Tag)
End Sub

'==============================================================================
' 刷新
'==============================================================================

' 刷新整个功能区。业务代码改变了任何影响按钮状态的东西之后调用它。
Public Sub RefreshRibbon()
    Dim rb As IRibbonUI
    Set rb = GetRibbon()
    If rb Is Nothing Then Exit Sub
    On Error Resume Next
    rb.Invalidate
    On Error GoTo 0
End Sub

Public Sub RefreshControl(ByVal controlId As String)
    Dim rb As IRibbonUI
    Set rb = GetRibbon()
    If rb Is Nothing Then Exit Sub
    On Error Resume Next
    rb.InvalidateControl controlId
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Ribbon 是否已成功加载。
'
' customUI14.xml 里只要有一处错误（重复 id、无效属性、坏掉的 imageMso），
' Excel 就会【静默】丢掉整个选项卡，VBA 编译却照样通过。所以需要一个能自动
' 检测的信号：OnLoad 只有在 customUI 被正确解析后才会触发。
'------------------------------------------------------------------------------
Public Function IsRibbonLoaded() As Boolean
    IsRibbonLoaded = Not (GetRibbon() Is Nothing)
End Function

'==============================================================================
' 指针持久化
'==============================================================================

Private Sub StoreRibbonPointer(ribbon As IRibbonUI)
    On Error Resume Next
    ThisWorkbook.Names(RIBBON_PTR_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=RIBBON_PTR_NAME, _
                           RefersTo:="=" & CStr(ObjPtr(ribbon)), _
                           Visible:=False
End Sub

Private Function GetRibbon() As IRibbonUI
    If Not mRibbon Is Nothing Then
        Set GetRibbon = mRibbon
        Exit Function
    End If

    ' 工程被重置过，从 Name 里把指针捞回来
    Dim s As String
#If VBA7 Then
    Dim p As LongPtr
#Else
    Dim p As Long
#End If

    On Error GoTo Fail
    s = ThisWorkbook.Names(RIBBON_PTR_NAME).RefersTo   ' 形如 "=123456789"
    s = Replace(s, "=", "")
    If Len(s) = 0 Or Not IsNumeric(s) Then GoTo Fail

#If VBA7 Then
    p = CLngPtr(CDec(s))
    CopyMemory VarPtr(mRibbon), VarPtr(p), LenB(p)
#Else
    p = CLng(s)
    CopyMemory VarPtr(mRibbon), VarPtr(p), 4&
#End If

    Set GetRibbon = mRibbon
    Exit Function
Fail:
    Set GetRibbon = Nothing
End Function
