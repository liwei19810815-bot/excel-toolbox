Attribute VB_Name = "modPublic"
'==============================================================================
' modPublic - 对外暴露的入口
'
' 其余模块一律带 Option Private Module，对外不可见；只有这里的过程能被
' Application.Run 调用。构建脚本、测试脚本、以及用户自己的宏都从这里进。
'
' 本模块【不能】加 Option Private Module。
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 自检。构建后由 build\verify.ps1 调用。
'
' 关键作用是【触发 VBA 编译】：Import 只是把源码塞进工程，语法和引用错误
' 要等到第一次执行时才暴露。只要这个函数能返回，就说明整个工程编译通过了。
'------------------------------------------------------------------------------
Public Function Toolbox_SelfCheck() As String
    Dim parts As String

    parts = "OK|" & APP_ID & "|" & APP_VERSION
    parts = parts & "|host=" & Application.Name
    parts = parts & "|bitness=" & modApp.HostBitness()
    parts = parts & "|actions=" & modAction.ActionCount()
    parts = parts & "|canUndo=" & CStr(modUndo.CanUndo())
    parts = parts & "|undoDepth=" & modUndo.Depth()
    parts = parts & "|ribbon=" & CStr(modRibbon.IsRibbonLoaded())

    Toolbox_SelfCheck = parts
End Function

' 所有已注册的 actionId，换行分隔。供测试脚本与 customUI14.xml 里的 tag 做一致性比对。
Public Function Toolbox_ListActions() As String
    Toolbox_ListActions = modAction.AllActionIds()
End Function

'------------------------------------------------------------------------------
' 供测试脚本调用：执行一个已注册的命令，返回结果文字。
'------------------------------------------------------------------------------
Public Function Toolbox_Run(ByVal actionId As String) As String
    modAction.RunAction actionId
    Toolbox_Run = modAction.LastMessage()
End Function

' 静默模式：不弹任何对话框。测试脚本必须先打开它，否则第一个 MsgBox 就会卡死。
Public Sub Toolbox_SetSilent(ByVal value As Boolean)
    modAction.SetSilent value
End Sub

' 静默模式下给工具预设参数（替代输入框）。key 见各工具里 AskXxx 的第一个参数。
Public Sub Toolbox_SetParam(ByVal key As String, ByVal value As String)
    modPrompt.SetParam key, value
End Sub

Public Sub Toolbox_ClearParams()
    modPrompt.ClearParams
End Sub

' 最近一次执行的结果文字。静默模式下出错时这里会是 "ERROR: <原因>"。
Public Function Toolbox_LastMessage() As String
    Toolbox_LastMessage = modAction.LastMessage()
End Function

'------------------------------------------------------------------------------
' 撤销上一步。成功返回空串，失败返回 "ERROR: <原因>"。
'
' 【必须自己处理错误】：这是个对外入口，会被 Application.Run 直接调用。
' 未捕获的 VBA 错误在这里不会变成干净的 COM 异常，而是弹出"运行时错误 1004"
' 的调试对话框——无界面运行时那个框谁也看不见，调用方就永久挂住了。
'
' Ribbon 上的撤销按钮走的是 RunAction，那条路本来就有统一错误处理；
' 漏的正是这个绕过管线的入口。
'------------------------------------------------------------------------------
Public Function Toolbox_Undo() As String
    On Error GoTo Failed
    modUndo.UndoLast
    Exit Function
Failed:
    Toolbox_Undo = "ERROR: " & Err.Description & " [" & Err.Number & " @ " & Err.Source & "]"
End Function

' 宏被强行中断后恢复 Excel 环境。也挂在 Ribbon 上，这里额外暴露给用户手工调用。
Public Sub Toolbox_ResetEnvironment()
    modPerf.FastModeReset
End Sub

'------------------------------------------------------------------------------
' 分段诊断。
'
' 无头环境下一旦某段代码挂住（典型是弹出了一个不可见的模态框，或者对失效的
' IRibbonUI 指针做了 Invalidate），调用方只会看到永久无响应，拿不到任何信息。
' 这里把 RunAction 管线拆成可以单独调用的小段，逐段试就能定位到具体是哪一环。
'------------------------------------------------------------------------------
Public Function Toolbox_DiagStage(ByVal stage As Long) As String
    Select Case stage
        Case 1
            modPerf.FastModeOn
            modPerf.FastModeOff
            Toolbox_DiagStage = "1 ok: FastMode 开关"

        Case 2
            modPerf.SetStatus "diag"
            modPerf.ClearStatus
            Toolbox_DiagStage = "2 ok: 状态栏"

        Case 3
            modUndo.BeginTx "diag"
            modUndo.Commit
            Toolbox_DiagStage = "3 ok: 空事务开启与提交"

        Case 4
            modRibbon.RefreshControl "btnUndoLast"
            Toolbox_DiagStage = "4 ok: Ribbon 刷新"

        Case 5
            Dim rng As Range
            Set rng = modRange.NormalizeSelection(Selection)
            If rng Is Nothing Then
                Toolbox_DiagStage = "5 ok: 选区归一化 -> Nothing"
            Else
                Toolbox_DiagStage = "5 ok: 选区归一化 -> " & rng.Address(False, False)
            End If

        Case 6
            modUndo.BeginTx "diag-capture"
            modUndo.CaptureSheet ActiveSheet
            modUndo.Commit
            Toolbox_DiagStage = "6 ok: 整表快照，canUndo=" & CStr(modUndo.CanUndo())

        Case Else
            Toolbox_DiagStage = "未知阶段：" & stage
    End Select
End Function
