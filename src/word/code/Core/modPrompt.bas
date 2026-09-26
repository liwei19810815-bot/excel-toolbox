Attribute VB_Name = "modPrompt"
'==============================================================================
' modPrompt - 参数采集层（Word 最小实现）
'
' shared/code/Core/modAction.bas 的 SetSilent 无条件调用
' modPrompt.SetSilent，RunAction 的 Failed 分支无条件比较
' Err.Number = modPrompt.ERR_CANCELLED——这两个符号必须存在才能编译。
'
' 【只实现这两个】：Word 首批三个命令（word.audit/cleanSpaces/
' updateFields）都没有 PromptsForInput:=True，用不到 Excel 那边
' AskText/AskNumber/AskRange 等一整套采集函数（那几个还耦合了
' Range/Worksheet 类型，原样搬过来也编译不过）。等 Word 真的有命令
' 需要弹参数输入框时，再照 Excel 的 modPrompt.bas 补对应的 Ask*
' 函数——没有调用方的函数不要提前搭。
'==============================================================================
Option Explicit
Option Private Module

Public Const ERR_CANCELLED As Long = vbObjectError + 999

Private mSilent As Boolean

Public Sub SetSilent(ByVal value As Boolean)
    mSilent = value
End Sub

Public Function IsSilent() As Boolean
    IsSilent = mSilent
End Function

Public Sub Cancel()
    Err.Raise ERR_CANCELLED, "modPrompt", "用户取消"
End Sub
