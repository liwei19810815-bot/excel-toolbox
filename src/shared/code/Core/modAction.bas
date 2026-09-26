Attribute VB_Name = "modAction"
'==============================================================================
' modAction - 统一执行管线（宿主无关部分）
'
' 所有工具的唯一入口。RunAction 负责：
'     前置校验 → 确认 → 开高速模式 → 开撤销事务 → 执行 → 提交/回滚
'     → 恢复环境 → 统一报错
'
' 【这是从 excel/code/Core/modAction.bas 拆出来的】：逐行核对过，RunAction
' 本身、注册表存取、ActionLabel/ActionScreentip 这几部分完全不碰 Excel 对象，
' 只调 modHost.*/modRibbon.*/modTelemetry.*/clsActionDef，以及三个"本工程内
' 必须存在同名实现"的自由函数：
'     RegisterAll()                  —— 注册这个宿主支持哪些命令
'     Dispatch(actionId) As String   —— 真正执行，返回结果文字
'     ActionExtraEnabled(actionId)   —— core.undoLast 之外的按钮禁用规则
'                                        （没有就一律返回 True）
'     ActionExtraPressed(actionId)   —— 开关型按钮的按下状态
'                                        （没有就一律返回 False）
' 这四个函数在各宿主自己的 modActionRegistry.bas 里实现，和 modHost 是
' 同一个"同名模块/同名自由函数，各工程各一份"的技巧：调用处不写模块前缀，
' VBA 按"当前工程内查找"解析，RunAction 完全不用知道自己调的是哪个宿主的。
'
' 新增一个工具的完整步骤（以 Excel 为例，其它宿主同理）：
'     1. 在 customUI14.xml 里加按钮，tag = actionId
'     2. 在 excel/code/Core/modActionRegistry.bas 的 RegisterAll 里
'        RegisterAction 一行
'     3. 在同文件 Dispatch 的 Select Case 里加一行，指向业务过程
' 业务过程本身不碰 ScreenUpdating、不碰撤销、不写 On Error 弹窗。
'==============================================================================
Option Explicit
Option Private Module

Private mActions As Object          ' Scripting.Dictionary: actionId -> clsActionDef
Private mSilent As Boolean          ' 静默模式：不弹任何对话框，供自动化测试使用
Private mLastMessage As String      ' 最近一次执行的结果文字，静默模式下供测试断言

'------------------------------------------------------------------------------
' 静默模式。
'
' 业务过程一律【不许自己弹框】：结果以字符串返回，由 RunAction 统一呈现。
' 这不只是为了测试——弹框散落在业务代码里，就没法把工具组合起来批量执行。
'------------------------------------------------------------------------------
Public Sub SetSilent(ByVal value As Boolean)
    mSilent = value
    modPrompt.SetSilent value
End Sub

Public Function LastMessage() As String
    LastMessage = mLastMessage
End Function

' 静默模式下不许有任何"会打扰人"的副作用。modHelp 用它来决定
' 要不要真的弹浏览器——测试里弹一个浏览器窗口出来，
' 和弹一个 MsgBox 一样属于打扰，只是不会挂死而已。
Public Function IsSilent() As Boolean
    IsSilent = mSilent
End Function

Private Sub Notify(ByVal msg As String, ByVal style As VbMsgBoxStyle)
    mLastMessage = msg
    If mSilent Then Exit Sub
    MsgBox msg, style, APP_NAME
End Sub

'==============================================================================
' 注册表
'==============================================================================

Private Sub EnsureRegistry()
    If Not mActions Is Nothing Then Exit Sub
    Set mActions = CreateObject("Scripting.Dictionary")
    mActions.CompareMode = vbTextCompare
    RegisterAll   ' 未加模块前缀：解析到本工程内该宿主自己的 modActionRegistry.bas
End Sub

Public Sub RegisterAction(ByVal id As String, _
                          ByVal Label As String, _
                          Optional ByVal Screentip As String = "", _
                          Optional ByVal Undoable As Boolean = True, _
                          Optional ByVal ConfirmBeforeRun As Boolean = False, _
                          Optional ByVal RequiresSelection As Boolean = False, _
                          Optional ByVal RequiresWorkbook As Boolean = True, _
                          Optional ByVal SupportedInWps As Boolean = True, _
                          Optional ByVal PromptsForInput As Boolean = False)
    Dim d As clsActionDef
    Set d = New clsActionDef
    d.Id = id
    d.Label = Label
    d.Screentip = Screentip
    d.Undoable = Undoable
    d.ConfirmBeforeRun = ConfirmBeforeRun
    d.RequiresSelection = RequiresSelection
    d.RequiresWorkbook = RequiresWorkbook
    d.SupportedInWps = SupportedInWps
    d.PromptsForInput = PromptsForInput
    Set mActions(id) = d
End Sub

Public Function GetAction(ByVal actionId As String) As clsActionDef
    EnsureRegistry
    If mActions.Exists(actionId) Then Set GetAction = mActions(actionId)
End Function

Public Function ActionCount() As Long
    EnsureRegistry
    ActionCount = mActions.Count
End Function

Public Function AllActionIds() As String
    EnsureRegistry
    Dim k As Variant, buf As String
    For Each k In mActions.Keys
        If Len(buf) > 0 Then buf = buf & vbLf
        buf = buf & k
    Next k
    AllActionIds = buf
End Function

'==============================================================================
' Ribbon 查询
'==============================================================================

Public Function ActionLabel(ByVal actionId As String) As String
    Dim d As clsActionDef
    Set d = GetAction(actionId)
    If d Is Nothing Then
        ActionLabel = actionId
        Exit Function
    End If

    If actionId = "core.undoLast" Then
        ' 撤销按钮上直接显示待撤销的操作名，用户不用猜会撤销掉什么
        Dim lbl As String
        lbl = modHost.Host_PeekLabel()
        If Len(lbl) > 0 Then ActionLabel = "撤销 " & lbl Else ActionLabel = d.Label
        Exit Function
    End If

    ActionLabel = d.Label
    ' "…" 由 PromptsForInput 统一生成，不在 XML 里手写，两边不会再对不上
    If d.PromptsForInput Then ActionLabel = ActionLabel & "…"
End Function

'------------------------------------------------------------------------------
' 悬停提示。
'
' 在注册时写的说明后面【自动追加一行可撤销性】。
' 这件事必须自动做：手写的话一定会漏——改之前 59 条里只有零星几条
' 在 supertip 末尾写了"可撤销。"，用户站在按钮前没有统一线索判断有没有退路。
'------------------------------------------------------------------------------
Public Function ActionScreentip(ByVal actionId As String) As String
    Dim d As clsActionDef
    Set d = GetAction(actionId)
    If d Is Nothing Then Exit Function

    Dim buf As String
    buf = d.Screentip

    Select Case actionId
        Case "core.undoLast", "core.selfTest", "core.about", "core.resetEnv"
            ' 这几个本身就是元操作，标注可撤销性只会让人困惑
        Case Else
            If Len(buf) > 0 Then buf = buf & vbCrLf
            If d.Undoable Then
                buf = buf & "【可撤销】执行后可用工具箱的撤销按钮还原。"
            Else
                buf = buf & "【不可撤销】工具箱的撤销按钮还原不了，建议先保存。"
            End If
    End Select

    ActionScreentip = buf
End Function

' 开关型按钮的按下状态。具体哪些 actionId 是开关型、按下条件是什么，
' 是宿主业务模块的事，这里只转派给 ActionExtraPressed（各宿主自己实现）。
Public Function IsActionPressed(ByVal actionId As String) As Boolean
    IsActionPressed = ActionExtraPressed(actionId)
End Function

'------------------------------------------------------------------------------
' 按钮是否可点。由 Ribbon 的 getEnabled 回调调用——每次功能区刷新都会跑一遍
' 全部按钮，所以这里既要快，也【不许抛错】：回调里的未处理异常会让宿主
' 静默画坏控件，而不会给任何提示。
'
' 出错时一律返回 True（可点）：点下去还有 RunAction 的前置校验会拦住并说明原因，
' 比一个说不出理由的灰按钮强。
'------------------------------------------------------------------------------
Public Function IsActionEnabled(ByVal actionId As String) As Boolean
    On Error GoTo Fallback

    Dim d As clsActionDef
    Set d = GetAction(actionId)
    If d Is Nothing Then Exit Function

    If modApp.IsWps And Not d.SupportedInWps Then Exit Function
    If d.RequiresWorkbook And Not modHost.Host_HasDocument() Then Exit Function

    ' 宿主能力探测：不支持的命令直接灰显，而不是让用户点了才看到报错。
    ' core.undoLast 是唯一一个所有宿主共用的判据（modHost.Host_CanUndo，
    ' 没有能力的宿主老实返回 False）；其余的能力判据是各宿主业务模块的事，
    ' 转派给 ActionExtraEnabled（各宿主自己实现，没有特殊判据就一律 True）。
    If actionId = "core.undoLast" Then
        IsActionEnabled = modHost.Host_CanUndo()
    Else
        IsActionEnabled = ActionExtraEnabled(actionId)
    End If
    Exit Function

Fallback:
    IsActionEnabled = True
End Function

'==============================================================================
' 执行管线
'==============================================================================

Public Sub RunAction(ByVal actionId As String)
    mLastMessage = vbNullString

    ' 遥测计时。用 Timer 而不是 Now：Now 的分辨率是秒，大部分命令测出来都是 0。
    ' 【Timer 在午夜会回绕】，所以下面取耗时时要处理负数。
    Dim startedAt As Single
    startedAt = Timer

    Dim d As clsActionDef
    Set d = GetAction(actionId)
    If d Is Nothing Then
        Notify "命令未注册：" & actionId, vbExclamation
        modTelemetry.TrackAction actionId, "blocked", 0
        Exit Sub
    End If

    ' --- 前置校验 ---
    '
    ' 这几条被拦下的路径【也要记遥测】，而且很有价值：
    ' 「请先选中区域」如果某个命令上出现得特别频繁，说明它的适用条件
    ' 没跟用户讲清楚，那是产品问题不是用户问题。只记成功次数就看不到这些。
    If modApp.IsWps And Not d.SupportedInWps Then
        Notify "「" & d.Label & "」在 WPS 下不可用。", vbInformation
        modTelemetry.TrackAction actionId, "blocked_wps", ElapsedMs(startedAt)
        Exit Sub
    End If
    If d.RequiresWorkbook And Not modHost.Host_HasDocument() Then
        Notify "请先打开一个工作簿。", vbInformation
        modTelemetry.TrackAction actionId, "blocked_nodoc", ElapsedMs(startedAt)
        Exit Sub
    End If
    If d.RequiresSelection Then
        If Not modHost.Host_HasValidSelection() Then
            Notify "请先选中要处理的单元格区域。", vbInformation
            modTelemetry.TrackAction actionId, "blocked_nosel", ElapsedMs(startedAt)
            Exit Sub
        End If
    End If

    ' --- 不可撤销操作的强制确认 ---
    ' 静默模式下视为已确认——测试要的就是执行本身，不是确认逻辑。
    If d.ConfirmBeforeRun And Not mSilent Then
        If MsgBox("「" & d.Label & "」执行后无法撤销。" & vbCrLf & vbCrLf & _
                  "建议先保存或备份当前文件。是否继续？", _
                  vbExclamation + vbYesNo + vbDefaultButton2, APP_NAME) <> vbYes Then
            modTelemetry.TrackAction actionId, "declined", ElapsedMs(startedAt)
            Exit Sub
        End If
    End If

    ' --- 执行 ---
    ' 这一段的顺序和配对不能改：
    '   FastModeOn 必须在最外层，保证任何出口都会 FastModeOff；
    '   BeginTx 必须在 Dispatch 之前，业务代码才有事务可以往里写快照。
    Dim txOpened As Boolean

    On Error GoTo Failed
    modHost.Host_FastModeOn
    modHost.Host_SetStatus d.Label & " 执行中…"

    If d.Undoable Then
        modHost.Host_BeginUndo d.Label
        txOpened = True
    End If

    Dim result As String
    result = Dispatch(actionId)   ' 未加模块前缀：解析到本工程内该宿主自己的实现

    ' Commit 可能因为规模超限丢弃撤销记录。它返回的警告必须原样转达，
    ' 否则用户会以为这一步能撤销。
    Dim undoWarning As String
    If txOpened Then undoWarning = modHost.Host_CommitUndo()

    modHost.Host_ClearStatus
    modHost.Host_FastModeOff
    modRibbon.RefreshControl "btnUndoLast"

    ' 结果统一在这里呈现，业务过程只负责返回文字
    If Len(undoWarning) > 0 Then
        If Len(result) > 0 Then result = result & vbCrLf & vbCrLf & undoWarning Else result = undoWarning
    End If
    If Len(result) > 0 Then Notify result, vbInformation

    modTelemetry.TrackAction actionId, "ok", ElapsedMs(startedAt)
    Exit Sub

Failed:
    ' 别把变量叫 eNum：VBA 标识符大小写不敏感，它和保留字 Enum 是同一个词，
    ' 会被解析成 Enum 语句，报"语法错误"。同类要避开的还有 IsEmpty / Name / Type 等。
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source                  ' 没有来源的错误信息在无头调试时几乎无法定位

    ' 用户在参数输入框上点了取消：正常情况下安静退出，不弹框也不留撤销步骤。
    '
    ' 但【回滚失败必须说】：取消发生在业务代码已经改了一部分之后，
    ' 回滚又没成功的话，数据就停在中间状态。这时候还安静退出，
    ' 用户会以为"我取消了所以什么都没发生"——那是最危险的误解。
    If errNum = modPrompt.ERR_CANCELLED Then
        Dim cancelRollbackOk As Boolean
        cancelRollbackOk = True
        If txOpened Then
            On Error Resume Next
            cancelRollbackOk = modHost.Host_RollbackUndo()
            If Err.Number <> 0 Then cancelRollbackOk = False
            Err.Clear
            On Error GoTo 0
        End If

        modHost.Host_ClearStatus
        modHost.Host_FastModeOff
        modRibbon.RefreshControl "btnUndoLast"

        If cancelRollbackOk Then
            mLastMessage = "CANCELLED"
            modTelemetry.TrackAction actionId, "cancel", ElapsedMs(startedAt)
        Else
            mLastMessage = "CANCELLED_ROLLBACK_FAILED"
            Notify "「" & d.Label & "」已取消，但【回滚失败】。" & vbCrLf & vbCrLf & _
                   "数据可能停在中间状态，请立即检查，必要时关闭文件不保存。", vbCritical
            ' 回滚失败是最需要 IT 立刻知道的一类事件，单列一个 outcome 便于告警
            modTelemetry.TrackAction actionId, "cancel_rollback_failed", _
                                      ElapsedMs(startedAt), errNum, errDesc
        End If
        Exit Sub
    End If

    ' 先把改了一半的数据还原回去，再恢复环境，最后才弹框。
    ' 顺序颠倒的话，用户会在一个半改坏的表上看到错误提示。
    Dim rollbackOk As Boolean
    rollbackOk = True
    If txOpened Then
        On Error Resume Next
        rollbackOk = modHost.Host_RollbackUndo()
        If Err.Number <> 0 Then rollbackOk = False
        Err.Clear
        On Error GoTo 0
    End If

    modHost.Host_ClearStatus
    modHost.Host_FastModeOff
    modRibbon.RefreshControl "btnUndoLast"

    ' 绝不声称没验证过的事：
    '   - 不可撤销的操作本来就没有回滚可言；
    '   - 回滚本身也会失败（目标表被关掉、被保护、快照丢失），
    '     这时候说"已回滚"会让用户放弃检查数据，比不说更糟。
    Dim tail As String
    If Not txOpened Then
        tail = "该操作不支持回滚，请检查数据是否已被部分修改。"
    ElseIf rollbackOk Then
        tail = "改动已回滚。"
    Else
        tail = "【回滚也失败了】，数据可能停在中间状态，请立即检查，必要时关闭文件不保存。"
    End If

    ' 静默模式下把失败原因留给测试断言，不弹框
    mLastMessage = "ERROR: " & errDesc & " [" & errNum & " @ " & errSrc & "]"
    If Not mSilent Then
        ' 出错是用户最需要帮助的时刻，顺手给一条路，而不是只留一个错误号。
        ' 用 vbYesNo 而不是再加一个按钮：Excel 的 MsgBox 没法自定义按钮文字，
        ' 把问句写清楚比按钮上写什么更重要。
        If MsgBox("「" & d.Label & "」执行失败。" & tail & vbCrLf & vbCrLf & _
                  "错误 " & errNum & "：" & errDesc & vbCrLf & vbCrLf & _
                  "要查看这个功能的使用帮助吗？", _
                  vbCritical + vbYesNo + vbDefaultButton2, APP_NAME) = vbYes Then
            On Error Resume Next
            modHelp.ShowFor actionId
            Err.Clear
            On Error GoTo 0
        End If
    End If

    ' 回滚成不成功要分开记：同一个错误，回滚失败的那些才是真正会伤到数据的，
    ' 混在一起统计就分不出轻重缓急了。
    modTelemetry.TrackAction actionId, _
                             IIf(txOpened And Not rollbackOk, "fail_rollback_failed", "fail"), _
                             ElapsedMs(startedAt), errNum, errDesc & " @ " & errSrc
End Sub

'------------------------------------------------------------------------------
' Timer 起点到现在的毫秒数。
'
' 【Timer 在午夜会归零】，跨零点执行的命令会算出负数。
' 那种情况下返回 0 而不是一个荒谬的负值——遥测里出现负耗时，
' 会让后面做统计的人白白花时间去查一个不存在的 bug。
'------------------------------------------------------------------------------
Private Function ElapsedMs(ByVal startedAt As Single) As Long
    Dim secs As Single
    secs = Timer - startedAt
    If secs < 0 Then Exit Function
    ElapsedMs = CLng(secs * 1000)
End Function
