Attribute VB_Name = "modAction"
'==============================================================================
' modAction - 统一执行管线
'
' 所有工具的唯一入口。RunAction 负责：
'     前置校验 → 确认 → 开高速模式 → 开撤销事务 → 执行 → 提交/回滚
'     → 恢复环境 → 统一报错
'
' 新增一个工具的完整步骤：
'     1. 在 customUI14.xml 里加按钮，tag = actionId
'     2. 在 RegisterAll 里 RegisterAction 一行
'     3. 在 Dispatch 的 Select Case 里加一行，指向业务过程
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
    RegisterAll
End Sub

' 各功能模块的注册入口都挂在这里，按模块分段，便于增删。
Private Sub RegisterAll()
    ' --- Core ---
    RegisterAction "core.undoLast", "撤销", "撤销上一步工具箱操作", _
                   Undoable:=False, RequiresWorkbook:=False
    RegisterAction "core.selfTest", "自检", "验证加载宏已正确装载", _
                   Undoable:=False, RequiresWorkbook:=False
    RegisterAction "core.about", "关于", "", _
                   Undoable:=False, RequiresWorkbook:=False
    RegisterAction "core.help", "帮助", _
                   "打开使用帮助：每个功能什么时候用、怎么用、有什么坑", _
                   Undoable:=False, RequiresWorkbook:=False
    RegisterAction "core.resetEnv", "环境复位", _
                   "恢复屏幕刷新、自动重算、事件响应和状态栏。" & _
                   "宏被强行中断后如果 Excel 变得没有反应或不自动计算，点这里", _
                   Undoable:=False, RequiresWorkbook:=False

    ' --- M1 文本与单元格 ---
    RegisterAction "text.cleanSpaces", "清除空格", _
                   "去掉首尾空格、压缩中间连续空格，并清除不间断空格、全角空格、零宽字符等不可见字符", _
                   RequiresSelection:=True
    RegisterAction "text.toNumber", "文本转数值", _
                   "把文本型数字转成真正的数值，自动处理全角数字、千分位逗号和不间断空格。" & _
                   "这是求和结果为 0 的头号原因", _
                   RequiresSelection:=True
    RegisterAction "text.toUpper", "转大写", "", RequiresSelection:=True
    RegisterAction "text.toLower", "转小写", "", RequiresSelection:=True
    RegisterAction "text.toProper", "首字母大写", "", RequiresSelection:=True
    RegisterAction "text.toHalfWidth", "全角转半角", "", RequiresSelection:=True
    RegisterAction "text.toFullWidth", "半角转全角", "", RequiresSelection:=True
    RegisterAction "text.removeLineBreaks", "删除换行符", _
                   "去掉单元格内的换行符", RequiresSelection:=True
    RegisterAction "text.extractDigits", "提取数字", _
                   "只保留数字字符", RequiresSelection:=True
    RegisterAction "text.extractChinese", "提取中文", "", RequiresSelection:=True
    RegisterAction "text.extractEnglish", "提取字母", "", RequiresSelection:=True
    RegisterAction "text.addAffix", "添加前后缀", _
                   "批量给选区加前缀或后缀", RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "text.regexReplace", "正则替换", _
                   "用正则表达式批量查找替换", RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "text.splitColumn", "按分隔符拆列", _
                   "把一列按分隔符拆成多列，自动插入所需列数，不覆盖右侧数据", _
                   RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "cells.unmergeFill", "拆分并填充", _
                   "取消合并单元格，并把原值填满整个区域", RequiresSelection:=True
    RegisterAction "cells.mergeSame", "合并相同项", _
                   "把同列中相邻且内容相同的单元格合并", RequiresSelection:=True

    ' --- M2 数据处理 ---
    RegisterAction "data.deleteEmptyRows", "删除空行", _
                   "删除选区内完全为空的整行。只按选中的列判断是否为空", RequiresSelection:=True
    RegisterAction "data.deleteEmptyCols", "删除空列", _
                   "删除选区内完全为空的整列", RequiresSelection:=True
    RegisterAction "data.markDuplicates", "标记重复值", _
                   "把重复行标成浅红色，支持多列组合判重", _
                   RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "data.deleteDuplicates", "删除重复值", _
                   "删除重复行并保留首次出现，支持多列组合判重", _
                   RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "data.extractUnique", "提取唯一值", _
                   "把选区内的唯一值提取到新工作表", RequiresSelection:=True
    RegisterAction "data.compare", "两表对比", _
                   "按键列对比两个区域，输出差异报告：仅 A 有 / 仅 B 有 / 内容不同，" & _
                   "每行都能点击跳回源数据", _
                   Undoable:=False, RequiresWorkbook:=True, PromptsForInput:=True
    RegisterAction "data.unpivot", "二维转一维", _
                   "逆透视：把交叉表展开成明细表。透视表、Power Query 和数据库都需要一维明细表", _
                   RequiresSelection:=True, Undoable:=False, PromptsForInput:=True
    RegisterAction "data.transpose", "行列转置", _
                   "把选区转置后输出到新工作表", RequiresSelection:=True, Undoable:=False

    ' --- M3 工作表管理 ---
    ' 拆表/排序/重命名都会改动工作簿结构且撤不回来，一律强制确认。
    ' 「按列拆分」一次可能生成上百张表，手工删回去比出错本身还痛苦。
    RegisterAction "sheet.splitByColumn", "按列拆分工作表", _
                   "按某列的值把数据拆分成多个工作表", _
                   RequiresSelection:=True, Undoable:=False, _
                   ConfirmBeforeRun:=True, PromptsForInput:=True
    RegisterAction "sheet.mergeAll", "合并所有工作表", _
                   "把当前工作簿所有工作表合并到一张汇总表。" & _
                   "按标题名对齐，而不是按列位置——某张表少一列也不会整体错位", Undoable:=False
    RegisterAction "sheet.createIndex", "生成目录", _
                   "生成带超链接的工作表目录", Undoable:=False
    RegisterAction "sheet.sort", "工作表排序", "按名称排列工作表", _
                   Undoable:=False, ConfirmBeforeRun:=True, PromptsForInput:=True
    ' 深度隐藏的表往往是作者有意藏起来的（参数表、中间计算表），
    ' 一次全部放出来之后没有记录能还原回去——所以也要确认
    RegisterAction "sheet.showAll", "显示所有表", _
                   "显示全部隐藏工作表（含深度隐藏）。" & _
                   "哪些表原来是隐藏的不会被记录下来，之后无法一键还原", _
                   Undoable:=False, ConfirmBeforeRun:=True
    RegisterAction "sheet.batchRename", "批量重命名表", _
                   "按选中的名称列表批量重命名工作表", _
                   RequiresSelection:=True, Undoable:=False, ConfirmBeforeRun:=True

    ' --- M4 多文件合并 ---
    RegisterAction "merge.folder", "合并文件夹", _
                   "合并一个文件夹内所有 Excel 文件的数据，结果带来源文件和来源工作表列。" & _
                   "单个文件失败不会中断整批，最后出失败清单", _
                   Undoable:=False, RequiresWorkbook:=False, PromptsForInput:=True

    ' --- M5 文件批处理（全部不可撤销，强制确认）---
    RegisterAction "file.list", "文件清单", _
                   "把文件夹内的文件清单导入工作表，并生成可填写的新文件名列", _
                   Undoable:=False, RequiresWorkbook:=True, PromptsForInput:=True
    RegisterAction "file.batchRename", "批量重命名", _
                   "按当前「文件清单」表的 F 列批量重命名文件。先全表校验，全部通过才动手", _
                   Undoable:=False, ConfirmBeforeRun:=True
    RegisterAction "file.exportSheets", "导出工作表", _
                   "把每张工作表导出为独立的 xlsx / CSV / PDF 文件", _
                   Undoable:=False, ConfirmBeforeRun:=True, PromptsForInput:=True
    ' 插进去的图片是浮动对象，撤不回来，手工一张张删很痛苦——必须确认
    RegisterAction "file.insertImages", "批量插图", _
                   "按单元格内容在指定文件夹里找同名图片并插入到右侧单元格", _
                   Undoable:=False, RequiresSelection:=True, _
                   ConfirmBeforeRun:=True, PromptsForInput:=True

    ' --- M6 公式与引用 ---
    RegisterAction "formula.toValues", "公式转值", _
                   "把选区内的公式替换为计算结果", RequiresSelection:=True
    RegisterAction "formula.findErrors", "定位错误值", _
                   "找出并标黄选区内的所有错误值", RequiresSelection:=True
    RegisterAction "formula.wrapIfError", "套用 IFERROR", _
                   "给选区内的公式批量加上 IFERROR 容错", _
                   RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "formula.breakLinks", "断开外部链接", _
                   "把所有外部链接公式转为当前值", _
                   Undoable:=False, ConfirmBeforeRun:=True
    RegisterAction "formula.cleanNames", "清理失效名称", _
                   "删除指向 #REF! 的已定义名称", Undoable:=False, ConfirmBeforeRun:=True
    RegisterAction "formula.toggleView", "显示公式", _
                   "在显示公式和显示结果之间切换", Undoable:=False
    RegisterAction "formula.cleanRules", "清理格式规则", _
                   "清除选区内堆叠的条件格式和数据验证", RequiresSelection:=True

    ' --- M7 数据体检 ---
    RegisterAction "audit.scan", "数据体检", _
                   "扫描当前工作表，列出空行、文本型数字、文本型日期、错误值、合并单元格等问题，可点击跳转", _
                   Undoable:=False
    RegisterAction "audit.quickClean", "一键清洗", _
                   "修掉最常见且无歧义的几类问题：清理空白字符、文本型数字转数值、删除空行。" & _
                   "合并单元格和错误值需要人工判断，不会自动改", RequiresSelection:=True

    ' --- M8 数据可视化 ---
    RegisterAction "viz.dataBars", "数据条", "给数值单元格添加数据条", RequiresSelection:=True
    RegisterAction "viz.colorScale", "色阶热力图", "三色色阶", RequiresSelection:=True
    RegisterAction "viz.iconSet", "图标集", "三色交通灯图标集", RequiresSelection:=True
    RegisterAction "viz.clearCF", "清除条件格式", "清除选区内所有条件格式", RequiresSelection:=True
    RegisterAction "viz.sparklines", "批量迷你图", _
                   "每行生成一个迷你图，放在数据右侧一列", _
                   RequiresSelection:=True, PromptsForInput:=True
    RegisterAction "viz.quickChart", "快速图表", _
                   "按选区生成图表并套用统一格式", _
                   RequiresSelection:=True, Undoable:=False, PromptsForInput:=True
    ' 会覆盖用户手工调好的图表格式，且撤不回来
    RegisterAction "viz.unifyCharts", "统一图表格式", _
                   "把当前工作表所有图表的格式统一，会覆盖你手工调过的格式", _
                   Undoable:=False, ConfirmBeforeRun:=True

    ' --- M9 辅助增强 ---
    RegisterAction "misc.spotlight", "聚光灯", _
                   "高亮光标所在的整行整列。用条件格式实现，不会破坏原有底色", _
                   Undoable:=False
    RegisterAction "misc.amountToChinese", "金额大写", _
                   "把选中的一列金额转成人民币中文大写，写到右侧一列", RequiresSelection:=True
    RegisterAction "misc.parseId", "身份证解析", _
                   "解析出生日期、性别、年龄，并校验校验位", RequiresSelection:=True
    RegisterAction "misc.normalizeDates", "日期规范化", _
                   "把 20240115 / 2024.1.15 / 2024年1月15日 等文本统一转成日期值", _
                   RequiresSelection:=True
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
        lbl = modUndo.PeekLabel()
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

' 开关型按钮的按下状态
Public Function IsActionPressed(ByVal actionId As String) As Boolean
    Select Case actionId
        Case "misc.spotlight": IsActionPressed = modSpotlight.IsEnabled()
        Case Else:             IsActionPressed = False
    End Select
End Function

'------------------------------------------------------------------------------
' 按钮是否可点。由 Ribbon 的 getEnabled 回调调用——59 个按钮，每次功能区刷新
' 就跑 59 次，所以这里既要快，也【不许抛错】：回调里的未处理异常会让 Excel
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
    If d.RequiresWorkbook And ActiveWorkbook Is Nothing Then Exit Function

    ' 宿主能力探测：不支持的命令直接灰显，而不是让用户点了才看到报错。
    ' 这比在 clsActionDef 上人工维护一张"哪个宿主支持哪个 API"的表可靠——
    ' 那张表在开发机上根本没法验证，事实上也一直是空的。
    Select Case actionId
        Case "core.undoLast"
            IsActionEnabled = modUndo.CanUndo()

        Case "viz.sparklines"
            IsActionEnabled = modCaps.SupportsSparklines()

        Case "merge.folder", "file.list", "file.exportSheets", "file.insertImages"
            IsActionEnabled = modCaps.SupportsFileDialog()

        Case Else
            IsActionEnabled = True
    End Select
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
    If d.RequiresWorkbook And ActiveWorkbook Is Nothing Then
        Notify "请先打开一个工作簿。", vbInformation
        modTelemetry.TrackAction actionId, "blocked_nodoc", ElapsedMs(startedAt)
        Exit Sub
    End If
    If d.RequiresSelection Then
        If TypeName(Selection) <> "Range" Then
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
    modPerf.FastModeOn
    modPerf.SetStatus d.Label & " 执行中…"

    If d.Undoable Then
        modUndo.BeginTx d.Label
        txOpened = True
    End If

    Dim result As String
    result = Dispatch(actionId)

    ' Commit 可能因为规模超限丢弃撤销记录。它返回的警告必须原样转达，
    ' 否则用户会以为这一步能撤销。
    Dim undoWarning As String
    If txOpened Then undoWarning = modUndo.Commit()

    modPerf.ClearStatus
    modPerf.FastModeOff
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
            cancelRollbackOk = modUndo.Rollback()
            If Err.Number <> 0 Then cancelRollbackOk = False
            Err.Clear
            On Error GoTo 0
        End If

        modPerf.ClearStatus
        modPerf.FastModeOff
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
        rollbackOk = modUndo.Rollback()
        If Err.Number <> 0 Then rollbackOk = False
        Err.Clear
        On Error GoTo 0
    End If

    modPerf.ClearStatus
    modPerf.FastModeOff
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

'------------------------------------------------------------------------------
' actionId -> 业务过程。这是全加载宏唯一的转派点。
'
' 返回值是给用户看的结果文字（"已删除 3 个空行"），由 RunAction 统一呈现。
' 业务过程不许自己弹框，否则批量执行和自动化测试都会被卡住。
'------------------------------------------------------------------------------
Private Function Dispatch(ByVal actionId As String) As String
    Select Case actionId
        ' --- Core ---
        Case "core.undoLast":  modUndo.UndoLast
        Case "core.resetEnv":  modPerf.FastModeReset
        Case "core.selfTest":  Dispatch = modApp.AboutText() & vbCrLf & vbCrLf & "加载宏工作正常。"
        Case "core.about":     Dispatch = modApp.AboutText()
        Case "core.help":      Dispatch = modHelp.ShowPane()

        ' --- M1 文本与单元格 ---
        Case "text.cleanSpaces":      Dispatch = modText.CleanSpaces(Selection)
        Case "text.toNumber":         Dispatch = modText.TextToNumber(Selection)
        Case "text.toUpper":          Dispatch = modText.ToUpper(Selection)
        Case "text.toLower":          Dispatch = modText.ToLower(Selection)
        Case "text.toProper":         Dispatch = modText.ToProper(Selection)
        Case "text.toHalfWidth":      Dispatch = modText.ToHalfWidth(Selection)
        Case "text.toFullWidth":      Dispatch = modText.ToFullWidth(Selection)
        Case "text.removeLineBreaks": Dispatch = modText.RemoveLineBreaks(Selection)
        Case "text.extractDigits":    Dispatch = modText.ExtractDigits(Selection)
        Case "text.extractChinese":   Dispatch = modText.ExtractChinese(Selection)
        Case "text.extractEnglish":   Dispatch = modText.ExtractEnglish(Selection)
        Case "text.addAffix":         Dispatch = modText.AddAffix(Selection)
        Case "text.regexReplace":     Dispatch = modText.RegexReplaceCells(Selection)
        Case "text.splitColumn":      Dispatch = modText.SplitColumn(Selection)
        Case "cells.unmergeFill":     Dispatch = modCells.UnmergeAndFill(Selection)
        Case "cells.mergeSame":       Dispatch = modCells.MergeSameValues(Selection)

        ' --- M2 数据处理 ---
        Case "data.deleteEmptyRows":  Dispatch = modRows.DeleteEmptyRows(Selection)
        Case "data.deleteEmptyCols":  Dispatch = modRows.DeleteEmptyColumns(Selection)
        Case "data.markDuplicates":   Dispatch = modDedupe.MarkDuplicates(Selection)
        Case "data.deleteDuplicates": Dispatch = modDedupe.DeleteDuplicates(Selection)
        Case "data.extractUnique":    Dispatch = modDedupe.ExtractUnique(Selection)
        Case "data.compare":          Dispatch = modCompare.CompareRanges()
        Case "data.unpivot":          Dispatch = modReshape.Unpivot(Selection)
        Case "data.transpose":        Dispatch = modReshape.TransposeRange(Selection)

        ' --- M3 工作表管理 ---
        Case "sheet.splitByColumn":   Dispatch = modSheets.SplitByColumn(Selection)
        Case "sheet.mergeAll":        Dispatch = modSheets.MergeAllSheets(ActiveWorkbook)
        Case "sheet.createIndex":     Dispatch = modSheets.CreateIndex(ActiveWorkbook)
        Case "sheet.sort":            Dispatch = modSheets.SortSheets(ActiveWorkbook)
        Case "sheet.showAll":         Dispatch = modSheets.ShowAllSheets(ActiveWorkbook)
        Case "sheet.batchRename":     Dispatch = modSheets.BatchRename(Selection)

        ' --- M4 多文件合并 ---
        Case "merge.folder":          Dispatch = modMergeFiles.MergeFolder()

        ' --- M5 文件批处理 ---
        Case "file.list":             Dispatch = modFileBatch.ListFiles()
        Case "file.batchRename":      Dispatch = modFileBatch.BatchRenameFiles(ActiveSheet)
        Case "file.exportSheets":     Dispatch = modFileBatch.ExportSheets(ActiveWorkbook)
        Case "file.insertImages":     Dispatch = modFileBatch.InsertImages(Selection)

        ' --- M6 公式与引用 ---
        Case "formula.toValues":      Dispatch = modFormula.FormulasToValues(Selection)
        Case "formula.findErrors":    Dispatch = modFormula.FindErrors(Selection)
        Case "formula.wrapIfError":   Dispatch = modFormula.WrapWithIfError(Selection)
        Case "formula.breakLinks":    Dispatch = modFormula.BreakExternalLinks(ActiveWorkbook)
        Case "formula.cleanNames":    Dispatch = modFormula.CleanBrokenNames(ActiveWorkbook)
        Case "formula.toggleView":    Dispatch = modFormula.ToggleFormulaView(ActiveSheet)
        Case "formula.cleanRules":    Dispatch = modFormula.CleanFormatRules(Selection)

        ' --- M7 数据体检 ---
        Case "audit.scan":            Dispatch = modAudit.ScanSheet(ActiveSheet)
        Case "audit.quickClean":      Dispatch = modAudit.QuickClean(Selection)

        ' --- M8 数据可视化 ---
        Case "viz.dataBars":          Dispatch = modViz.AddDataBars(Selection)
        Case "viz.colorScale":        Dispatch = modViz.AddColorScale(Selection)
        Case "viz.iconSet":           Dispatch = modViz.AddIconSet(Selection)
        Case "viz.clearCF":           Dispatch = modViz.ClearConditionalFormats(Selection)
        Case "viz.sparklines":        Dispatch = modViz.AddSparklines(Selection)
        Case "viz.quickChart":        Dispatch = modViz.QuickChart(Selection)
        Case "viz.unifyCharts":       Dispatch = modViz.UnifyCharts(ActiveSheet)

        ' --- M9 辅助增强 ---
        Case "misc.spotlight":        Dispatch = modSpotlight.Toggle()
        Case "misc.amountToChinese":  Dispatch = modUtils.AmountColumnToChinese(Selection)
        Case "misc.parseId":          Dispatch = modUtils.ParseIdCards(Selection)
        Case "misc.normalizeDates":   Dispatch = modUtils.NormalizeDates(Selection)

        Case Else
            Err.Raise vbObjectError + 1, "modAction.Dispatch", _
                      "actionId 已注册但未实现转派：" & actionId
    End Select
End Function
