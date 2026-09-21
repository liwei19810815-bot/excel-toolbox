Attribute VB_Name = "modSettings"
'==============================================================================
' modSettings - 配置读写
'
' 走 VBA 内置的 SaveSetting/GetSetting（落在 HKCU\Software\VB and VBA Program Settings），
' 不需要额外文件，也不碰 Office 自己的注册表分支。
'==============================================================================
Option Explicit
Option Private Module

Private Const SECTION_GENERAL As String = "General"

' 默认值集中在这里，避免各处散落魔数
Public Const DEF_MAX_UNDO_STEPS As Long = 5
Public Const DEF_MAX_UNDO_CELLS As Double = 1000000#

'------------------------------------------------------------------------------
' 注册表访问一律容错。
'
' 企业环境里 HKCU 被组策略限制、注册表虚拟化异常都是真实存在的。
' 配置读不出来最多是用默认值，绝不能因此让整个命令执行失败——
' 用户想做的是删个空行，不该被一个读配置的错误挡住。
'------------------------------------------------------------------------------
Private Function ReadRaw(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    ReadRaw = GetSetting(APP_ID, SECTION_GENERAL, key, defaultValue)
    If Err.Number <> 0 Then
        Err.Clear
        ReadRaw = defaultValue
    End If
    On Error GoTo 0
End Function

Public Function GetSettingString(ByVal key As String, ByVal defaultValue As String) As String
    GetSettingString = ReadRaw(key, defaultValue)
End Function

Public Function GetSettingLong(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim s As String
    s = ReadRaw(key, CStr(defaultValue))
    If IsNumeric(s) Then GetSettingLong = CLng(s) Else GetSettingLong = defaultValue
End Function

Public Function GetSettingDouble(ByVal key As String, ByVal defaultValue As Double) As Double
    Dim s As String
    s = ReadRaw(key, CStr(defaultValue))
    If IsNumeric(s) Then GetSettingDouble = CDbl(s) Else GetSettingDouble = defaultValue
End Function

Public Function GetSettingBool(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim s As String
    s = ReadRaw(key, CStr(defaultValue))
    GetSettingBool = (LCase$(s) = "true" Or s = "-1" Or s = "1")
End Function

' 写失败就算了：配置存不下来只影响下次的默认值，不值得打断用户
Public Sub PutSetting(ByVal key As String, ByVal value As Variant)
    On Error Resume Next
    SaveSetting APP_ID, SECTION_GENERAL, key, CStr(value)
    Err.Clear
    On Error GoTo 0
End Sub

'--- 具名快捷方式 -------------------------------------------------------------

Public Function MaxUndoSteps() As Long
    MaxUndoSteps = GetSettingLong("MaxUndoSteps", DEF_MAX_UNDO_STEPS)
    If MaxUndoSteps < 1 Then MaxUndoSteps = 1
End Function

' 单次操作的快照规模上限。超过则不入撤销栈，改为执行前确认。
Public Function MaxUndoCells() As Double
    MaxUndoCells = GetSettingDouble("MaxUndoCells", DEF_MAX_UNDO_CELLS)
    If MaxUndoCells < 1000# Then MaxUndoCells = 1000#
End Function
