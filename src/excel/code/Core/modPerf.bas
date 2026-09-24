Attribute VB_Name = "modPerf"
'==============================================================================
' modPerf - 高速模式：Application 环境开关的保存与可靠还原
'
' 批量操作时关掉屏幕刷新/自动重算/事件，能把耗时降一到两个数量级。
' 但这里真正的难点不是"关"，而是"无论如何都要还原"——一旦出错路径漏了还原，
' 用户的 Excel 会一直停在手动重算、无屏幕刷新的状态，且毫无提示。
'
' 因此：
'   1. 只在 RunAction 管线里成对调用，业务代码不许自己碰这些属性；
'   2. 支持嵌套（内层调用只加减计数，不重复保存/还原）；
'   3. 保存的是"进入时的真实值"，还原的也是它，而不是无脑设 True/xlCalculationAutomatic
'      ——用户原本就处于手动重算时，不该被我们改成自动。
'==============================================================================
Option Explicit
Option Private Module

Private mDepth As Long

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mCalculation As XlCalculation
Private mStatusBar As Variant
Private mCursor As XlMousePointer
Private mCalcSaved As Boolean

'------------------------------------------------------------------------------
' 进入/退出高速模式。必须成对调用，退出务必放在错误处理路径里。
'------------------------------------------------------------------------------
Public Sub FastModeOn()
    mDepth = mDepth + 1
    If mDepth > 1 Then Exit Sub          ' 嵌套：已经在高速模式里了

    With Application
        mScreenUpdating = .ScreenUpdating
        mEnableEvents = .EnableEvents
        mDisplayAlerts = .DisplayAlerts
        mStatusBar = .StatusBar
        mCursor = .Cursor

        ' 无工作簿打开时读写 Calculation 会报错，单独兜一下
        mCalcSaved = False
        On Error Resume Next
        mCalculation = .Calculation
        mCalcSaved = (Err.Number = 0)
        On Error GoTo 0

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        If mCalcSaved Then .Calculation = xlCalculationManual
    End With
End Sub

Public Sub FastModeOff()
    If mDepth = 0 Then Exit Sub
    mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub

    ' 还原逐条独立容错：任何一条失败都不能连累后面几条
    With Application
        On Error Resume Next
        If mCalcSaved Then .Calculation = mCalculation
        .DisplayAlerts = mDisplayAlerts
        .EnableEvents = mEnableEvents
        .Cursor = mCursor
        .StatusBar = mStatusBar        ' False 表示交还给 Excel 自己管
        .ScreenUpdating = mScreenUpdating
        On Error GoTo 0
    End With
End Sub

'------------------------------------------------------------------------------
' 强制复位。用于兜底：VBA 工程被重置后 mDepth 会清零，但 Application 可能
' 还停在高速模式。提供一个"急救"入口，也挂到 Ribbon 的设置菜单里。
'------------------------------------------------------------------------------
Public Sub FastModeReset()
    mDepth = 0
    On Error Resume Next
    With Application
        .Calculation = xlCalculationAutomatic
        .DisplayAlerts = True
        .EnableEvents = True
        .Cursor = xlDefault
        .StatusBar = False
        .ScreenUpdating = True
    End With
    On Error GoTo 0
End Sub

Public Function InFastMode() As Boolean
    InFastMode = (mDepth > 0)
End Function

'------------------------------------------------------------------------------
' 状态栏提示。高速模式下 StatusBar 仍然会刷新，是唯一低成本的进度反馈。
'------------------------------------------------------------------------------
Public Sub SetStatus(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = APP_NAME & "：" & text
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 带计数的进度提示：「批量插图 37/200：产品A.jpg」。
'
' 长任务只写一句"执行中…"，用户看不出是在干活还是卡死了，几分钟后就会去点
' 任务管理器结束进程——那才是真正丢数据的时刻。有了分母，至少知道还要等多久。
'
' 【必须节流】。写 StatusBar 要过一次 COM，一万行的循环里每次都写，
' 光刷状态栏就比干活还慢。每 25 次写一次，人眼看不出区别。
'------------------------------------------------------------------------------
Public Sub SetProgress(ByVal label As String, ByVal current As Long, _
                       ByVal total As Long, Optional ByVal detail As String = "")
    If total <= 0 Then Exit Sub
    ' 首尾必须报，中间节流：否则小批量（总数 < 25）可能一次都不显示
    If current > 1 And current < total Then
        If current Mod 25 <> 0 Then Exit Sub
    End If

    Dim text As String
    text = label & " " & current & "/" & total
    If Len(detail) > 0 Then text = text & "：" & detail
    SetStatus text
End Sub

Public Sub ClearStatus()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub
