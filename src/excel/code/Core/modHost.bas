Attribute VB_Name = "modHost"
'==============================================================================
' modHost - 宿主契约的 Excel 实现
'
' 【为什么需要这一层】：modAction.RunAction 是整套执行管线，未来 PPT/Word
' 也要复用它（各自一份 modAction.bas，Dispatch 里的 60 个 Select Case
' 是宿主专属的，不抽象；但"前置校验 → 高速模式 → 撤销事务 → 恢复环境"
' 这套控制流是通用的）。这里把 RunAction 用到的、真正碰 Excel 对象的部分
' 收进一组 Host_* 函数——RunAction 只调这些函数名，以后新增一个宿主，
' 只要在那个宿主的工程里放一份实现了同名函数的 modHost.bas，
' RunAction 的代码不用再改一次。
'
' 【这是包一层，不是重写】：FastMode/Undo 的实际实现仍然是 modPerf/modUndo，
' 一行逻辑都没有搬动。出问题时回滚只需要把 modAction 里的调用改回直接调用
' modPerf/modUndo，不涉及这两个模块内部。
'
' 【只实现 modAction 实际用到的函数】：规划文档里列的契约还有
' Host_Kind/Host_Version/Host_Bitness，本轮没有调用方（modApp/modCaps
' 已经各自处理了这部分），先不加空壳——没有调用方的抽象就是没人验证过的
' 抽象，等 PPT 那边真的需要按宿主分支时再补。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 前置校验：有没有文档、选区合不合法。
'------------------------------------------------------------------------------
Public Function Host_HasDocument() As Boolean
    Host_HasDocument = Not (ActiveWorkbook Is Nothing)
End Function

Public Function Host_HasValidSelection() As Boolean
    Host_HasValidSelection = (TypeName(Selection) = "Range")
End Function

'------------------------------------------------------------------------------
' 高速模式：包一层 modPerf，实现和调用配对关系原样不动。
'------------------------------------------------------------------------------
Public Sub Host_FastModeOn()
    modPerf.FastModeOn
End Sub

Public Sub Host_FastModeOff()
    modPerf.FastModeOff
End Sub

Public Sub Host_FastModeReset()
    modPerf.FastModeReset
End Sub

Public Sub Host_SetStatus(ByVal text As String)
    modPerf.SetStatus text
End Sub

Public Sub Host_ClearStatus()
    modPerf.ClearStatus
End Sub

'------------------------------------------------------------------------------
' 撤销事务：包一层 modUndo。
'------------------------------------------------------------------------------
Public Sub Host_BeginUndo(ByVal Label As String)
    modUndo.BeginTx Label
End Sub

Public Function Host_CommitUndo() As String
    Host_CommitUndo = modUndo.Commit()
End Function

Public Function Host_RollbackUndo() As Boolean
    Host_RollbackUndo = modUndo.Rollback()
End Function

Public Function Host_CanUndo() As Boolean
    Host_CanUndo = modUndo.CanUndo()
End Function

Public Function Host_PeekLabel() As String
    Host_PeekLabel = modUndo.PeekLabel()
End Function

Public Sub Host_UndoLast()
    modUndo.UndoLast
End Sub
