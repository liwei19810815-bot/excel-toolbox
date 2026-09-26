Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - Word 工具箱功能区回调
'
' 和 excel/code/Core/modRibbon.bas 同一个约定：回调只负责解析控件参数，
' 转派给 modAction.RunAction，不写业务逻辑。
'
' IRibbonUI 指针持久化：VBA 工程被重置后模块级对象变量会清空，用
' ThisDocument.Variables 存 ObjPtr，重置后再从这里捞回来——和 Excel 版
' 用 ThisWorkbook.Names、PPT 版用 Presentation.Tags 是同一个技巧，
' 换成 Word 这边真实可用的持久化点（真实 COM 调用验证过
' Document.Variables.Add/.Item(name).Value 可读写）。
'
' 没有开关型按钮（没有 misc.spotlight 那种），所以没有 OnToggle/
' GetPressed；没有搜索框（首批三个命令用不上，PPT 那边同样没做）。
'==============================================================================
Option Explicit
' 注意：本模块【不能】加 Option Private Module。
' Ribbon 回调由 Word 通过和 Application.Run 相同的机制按名字查找，
' 私有模块里的过程找不到，表现为按钮点了没反应。

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const RIBBON_PTR_NAME As String = "WordToolbox_RibbonPtr"

Private mRibbon As IRibbonUI

'==============================================================================
' 回调
'==============================================================================

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    StoreRibbonPointer ribbon
End Sub

Public Sub Ribbon_OnAction(control As IRibbonControl)
    Select Case control.Tag
        Case "core.about"
            MsgBox modApp.AboutText(), vbInformation, modApp.APP_NAME
        Case Else
            modAction.RunAction control.Tag
    End Select
End Sub

'==============================================================================
' 【所有 get* 回调都必须自己兜住异常】，理由和 Excel 版一致：出错时 Word
' 会静默画坏控件而不给提示，宁可按钮多亮着，也不能让整个选项卡渲染异常。
'==============================================================================

Public Sub Ribbon_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo Fallback
    If control.Tag = "core.about" Then
        returnedVal = True
    Else
        returnedVal = modAction.IsActionEnabled(control.Tag)
    End If
    Exit Sub
Fallback:
    returnedVal = True
End Sub

Public Sub Ribbon_GetLabel(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo Fallback
    If control.Tag = "core.about" Then
        returnedVal = "关于"
    Else
        returnedVal = modAction.ActionLabel(control.Tag)
    End If
    Exit Sub
Fallback:
    returnedVal = control.Tag
End Sub

Public Sub Ribbon_GetScreentip(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo Fallback
    If control.Tag = "core.about" Then
        returnedVal = vbNullString
    Else
        returnedVal = modAction.ActionScreentip(control.Tag)
    End If
    Exit Sub
Fallback:
    returnedVal = vbNullString
End Sub

Public Sub Ribbon_GetSupertip(control As IRibbonControl, ByRef returnedVal)
    Ribbon_GetScreentip control, returnedVal
End Sub

'==============================================================================
' 刷新
'==============================================================================

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

Public Function IsRibbonLoaded() As Boolean
    IsRibbonLoaded = Not (GetRibbon() Is Nothing)
End Function

'==============================================================================
' 指针持久化
'==============================================================================

Private Sub StoreRibbonPointer(ribbon As IRibbonUI)
    On Error Resume Next
    ThisDocument.Variables(RIBBON_PTR_NAME).Value = CStr(ObjPtr(ribbon))
    If Err.Number <> 0 Then
        Err.Clear
        ThisDocument.Variables.Add RIBBON_PTR_NAME, CStr(ObjPtr(ribbon))
    End If
    On Error GoTo 0
End Sub

Private Function GetRibbon() As IRibbonUI
    If Not mRibbon Is Nothing Then
        Set GetRibbon = mRibbon
        Exit Function
    End If

    Dim s As String
#If VBA7 Then
    Dim p As LongPtr
#Else
    Dim p As Long
#End If

    On Error GoTo Fail
    s = ThisDocument.Variables(RIBBON_PTR_NAME).Value
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
