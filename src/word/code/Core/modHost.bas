Attribute VB_Name = "modHost"
'==============================================================================
' modHost - 宿主契约的 Word 实现
'
' 和 excel/code/Core/modHost.bas 同一个理由：modAction.RunAction 只调
' modHost.* 这组函数名，不直接碰 Word 对象，新增宿主只要在自己的工程里
' 放一份实现了同名函数的 modHost.bas。
'
' 【Word 的撤销和 Excel 完全不是一回事】：Excel 没有可用的原生撤销
' （宏一执行原生 Ctrl+Z 就失效），所以自建 modUndo 快照栈；Word 有
' Application.UndoRecord，真实 COM 调用验证过 StartCustomRecord/
' EndCustomRecord 能把中间任意多次编辑合并成一条原生撤销记录，
' Host_BeginUndo/Host_CommitUndo 就是包一层这个。
'
' 但 UndoRecord 只能"开始记/结束记"，没有暴露"查询是否可撤销"或
' "看一眼上一条记录叫什么"的 API——这意味着 Excel 那套由
' Host_CanUndo/Host_PeekLabel 驱动的"撤销上一步"按钮（core.undoLast）
' 在 Word 上做不出来，不是没设计好，是 Office 没给这个能力。
' 所以 word/code/Core/modActionRegistry.bas 的 RegisterAll 不注册
' core.undoLast，customUI 里也没有这个按钮——GetAction("core.undoLast")
' 会返回 Nothing，modAction.IsActionEnabled/ActionLabel 都会在
' "d Is Nothing" 那一步提前退出，根本不会走到调用
' Host_CanUndo/Host_PeekLabel 的分支。这三个函数仍然要实现（modAction
' 里是显式 modHost.Host_CanUndo() 这样的早绑定调用，编译期就要求这个
' 符号存在），但老实返回"不支持"，不能弄虚作假。
'
' Host_RollbackUndo 同理老实：UndoRecord 没有"取消当前记录、把已经做的
' 编辑撤销掉"这个能力，出错时能做的只是把记录关掉（避免留下一个悬空
' 的 UndoRecord 影响后续操作），但不能真的把数据改回去，所以固定返回
' False——调用方（RunAction 的 Failed 分支）看到 False 会提示"数据可能
' 停在中间状态，请立即检查"，这是诚实的默认，不是最优体验，但没有
' 验证过的能力不能声称有。
'==============================================================================
Option Explicit
Option Private Module

Private mUndoOpen As Boolean

'------------------------------------------------------------------------------
' 前置校验：有没有文档、选区合不合法。
'------------------------------------------------------------------------------
Public Function Host_HasDocument() As Boolean
    Host_HasDocument = (Application.Documents.Count > 0)
End Function

Public Function Host_HasValidSelection() As Boolean
    On Error GoTo NoSelection
    Host_HasValidSelection = Not (Selection Is Nothing)
    Exit Function
NoSelection:
    Host_HasValidSelection = False
End Function

'------------------------------------------------------------------------------
' 高速模式：Word 没有 Application.EnableEvents 这个属性（和 PowerPoint
' 一样是 Excel 专有的），这里只管 ScreenUpdating。
'------------------------------------------------------------------------------
Public Sub Host_FastModeOn()
    Application.ScreenUpdating = False
End Sub

Public Sub Host_FastModeOff()
    Application.ScreenUpdating = True
End Sub

Public Sub Host_FastModeReset()
    Application.ScreenUpdating = True
End Sub

Public Sub Host_SetStatus(ByVal text As String)
    On Error Resume Next
    Application.StatusBar = text
    On Error GoTo 0
End Sub

Public Sub Host_ClearStatus()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 撤销事务：包一层 Application.UndoRecord。
'------------------------------------------------------------------------------
Public Sub Host_BeginUndo(ByVal Label As String)
    On Error GoTo Failed
    Application.UndoRecord.StartCustomRecord Label
    mUndoOpen = True
    Exit Sub
Failed:
    ' 开启失败也不能让整条管线崩掉——按"没有事务"处理，业务代码照常执行，
    ' 只是这次的改动不会被合并成一条原生撤销记录。
    mUndoOpen = False
End Sub

Public Function Host_CommitUndo() As String
    If mUndoOpen Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
        Err.Clear
        On Error GoTo 0
        mUndoOpen = False
    End If
    ' 没有"这次规模太大放弃快照"这类情况（UndoRecord 不像 Excel 的
    ' modUndo 那样有容量上限），老实返回空字符串，不编造警告。
    Host_CommitUndo = vbNullString
End Function

Public Function Host_RollbackUndo() As Boolean
    If mUndoOpen Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
        Err.Clear
        On Error GoTo 0
        mUndoOpen = False
    End If
    ' 见文件头部说明：UndoRecord 没有"取消并回滚"的能力，这里只是把
    ' 记录关掉，不代表数据真的被还原了，所以固定返回 False。
    Host_RollbackUndo = False
End Function

' core.undoLast 在 Word 上不注册（见文件头部说明），这三个函数不会被
' modAction 实际调用到，但符号必须存在才能编译通过——老实返回"不支持"。
Public Function Host_CanUndo() As Boolean
    Host_CanUndo = False
End Function

Public Function Host_PeekLabel() As String
    Host_PeekLabel = vbNullString
End Function

Public Sub Host_UndoLast()
    ' 不支持，什么都不做——正常情况下这条路径不会被触发。
End Sub
