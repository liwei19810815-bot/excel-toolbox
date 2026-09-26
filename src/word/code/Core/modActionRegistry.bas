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
End Sub

Public Function Dispatch(ByVal actionId As String) As String
    Select Case actionId
        Case "word.audit":         Dispatch = modAudit.ScanDocument()
        Case "word.cleanSpaces":   Dispatch = modClean.CleanSpaces()
        Case "word.updateFields":  Dispatch = modFields.UpdateAllFields()

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
