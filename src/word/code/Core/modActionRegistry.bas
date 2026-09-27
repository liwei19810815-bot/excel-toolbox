Attribute VB_Name = "modActionRegistry"
'==============================================================================
' modActionRegistry - Word 命令注册表与派发（modAction 的宿主专属部分）
'
' 和 excel/code/Core/modActionRegistry.bas 是同一个技巧：shared/code/Core/
' modAction.bas 的 RunAction 不写模块前缀调用 RegisterAll()/Dispatch()/
' ActionExtraEnabled()/ActionExtraPressed()，VBA 在同一工程内按名字解析，
' 这四个函数在这里实现。
'
' 首批三个命令（见 docs/规划-Word与PPT.md 第八节设计），刻意覆盖
' RunAction 的三条分支：
'   word.audit         只读，Undoable=False，不强制确认
'   word.cleanSpaces    Undoable=True，走 Host_BeginUndo/CommitUndo
'                        （包一层 Application.UndoRecord）
'   word.updateFields   Undoable=False，ConfirmBeforeRun=True
'
' 第二批三个命令（同样每条用到的 Office API 都用真实 COM 调用单独
' 验证过，VBOM 信任这台机器仍然没开，没法走真正的 VBE 编译）：
'   word.removeEmptyParagraphs  Undoable=True，走 UndoRecord
'   word.replaceText            Undoable=False + ConfirmBeforeRun=True，
'                                第一次用到 PromptsForInput（modPrompt.
'                                AskText，Word 这边此前没有调用方，这次
'                                补上）——本来想标 Undoable=True，但
'                                Word 的 Host_RollbackUndo 一律返回
'                                False（UndoRecord 没有真回滚能力），
'                                带 PromptsForInput 的命令在第一个输入框
'                                就取消是最常见路径，这时候事务已经被
'                                RunAction 打开但根本没碰过文档，会被
'                                误判成"回滚失败"弹假警报——改成不可
'                                撤销+强制确认，避免这个坑
'   word.acceptAllRevisions     Undoable=True，AcceptAllRevisions() 包在
'                                UndoRecord 里真机验证过不会崩，且没有
'                                PromptsForInput，不会撞上面那个坑
'
' 不注册 core.undoLast：Word 的 UndoRecord 没有查询能力，Excel 那套
' "撤销上一步"按钮做不出来，见 modHost.bas 头部说明。
'==============================================================================
Option Explicit
Option Private Module

Public Sub RegisterAll()
    modAction.RegisterAction "word.audit", "文档体检", _
                   "扫描当前文档，统计空段落、手动换行符、超长段落、连续空格，只读不修改", _
                   Undoable:=False, RequiresWorkbook:=True
    modAction.RegisterAction "word.cleanSpaces", "清理多余空格", _
                   "清理全文的连续空格、全角空格、不间断空格和零宽字符，" & _
                   "合并成一条原生撤销记录，可用 Ctrl+Z 撤销", _
                   Undoable:=True, RequiresWorkbook:=True
    modAction.RegisterAction "word.updateFields", "更新域", _
                   "更新全文所有域（含目录）。域更新后的撤销语义复杂，不承诺能撤销", _
                   Undoable:=False, ConfirmBeforeRun:=True, RequiresWorkbook:=True
    modAction.RegisterAction "word.removeEmptyParagraphs", "清理多余空行", _
                   "把连续 2 个及以上的空行压缩成 1 个，孤立的单个空行保留。" & _
                   "合并成一条原生撤销记录，可用 Ctrl+Z 撤销", _
                   Undoable:=True, RequiresWorkbook:=True
    ' 【不能标 Undoable:=True】：Word 的 Host_RollbackUndo 只要事务被
    ' 打开过就固定返回 False（UndoRecord 没有真正的回滚能力，见
    ' modHost.bas 头部说明）。这个命令带 PromptsForInput，用户在第一个
    ' 输入框上点"取消"是最常见的取消路径——此时事务已经被 RunAction
    ' 打开（Undoable:=True 会让 Host_BeginUndo 在 Dispatch 之前就执行），
    ' 但根本还没碰过文档，RunAction 的取消分支却会因为 Host_RollbackUndo
    ' 返回 False 而弹"已取消，但回滚失败，数据可能停在中间状态"——一句
    ' 误导性的假警报。改成 Undoable:=False + ConfirmBeforeRun:=True，
    ' 和 word.updateFields 同一个理由：宁可少一点"合并成一条原生撤销
    ' 记录"的体验，也不要在最常见的取消路径上吓用户。
    modAction.RegisterAction "word.replaceText", "批量替换", _
                   "全文查找替换（精确匹配，最多 255 字符，不支持通配符）", _
                   Undoable:=False, ConfirmBeforeRun:=True, RequiresWorkbook:=True, PromptsForInput:=True
    modAction.RegisterAction "word.acceptAllRevisions", "接受所有修订", _
                   "接受文档里全部的修订标记。合并成一条原生撤销记录，可用 Ctrl+Z 撤销", _
                   Undoable:=True, RequiresWorkbook:=True
End Sub

Public Function Dispatch(ByVal actionId As String) As String
    Select Case actionId
        Case "word.audit":         Dispatch = modAudit.ScanDocument()
        Case "word.cleanSpaces":   Dispatch = modClean.CleanSpaces()
        Case "word.updateFields":  Dispatch = modFields.UpdateAllFields()
        Case "word.removeEmptyParagraphs": Dispatch = modClean.RemoveEmptyParagraphs()
        Case "word.replaceText":   Dispatch = modReplace.ReplaceText()
        Case "word.acceptAllRevisions": Dispatch = modRevisions.AcceptAllRevisions()

        Case Else
            Err.Raise vbObjectError + 1, "modActionRegistry.Dispatch", _
                      "actionId 已注册但未实现转派：" & actionId
    End Select
End Function

' 首批三个命令没有需要额外能力探测的情况（不像 Excel 的
' viz.sparklines/文件对话框那样依赖运行时环境），一律可点。
Public Function ActionExtraEnabled(ByVal actionId As String) As Boolean
    ActionExtraEnabled = True
End Function

' 首批三个命令都不是开关型按钮。
Public Function ActionExtraPressed(ByVal actionId As String) As Boolean
    ActionExtraPressed = False
End Function
