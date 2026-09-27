Attribute VB_Name = "modPrompt"
'==============================================================================
' modPrompt - 参数采集层（Word 最小实现）
'
' shared/code/Core/modAction.bas 的 SetSilent 无条件调用
' modPrompt.SetSilent，RunAction 的 Failed 分支无条件比较
' Err.Number = modPrompt.ERR_CANCELLED——这两个符号必须存在才能编译。
'
' 【AskText 是第二批命令（word.replaceText）第一次用到 PromptsForInput
' 才补的】：第一批三个命令都没有参数，没有调用方之前没有加。AskText
' 本身不碰 Range/Worksheet 这类 Excel 专属类型（InputBox 是 VBA 内置
' 函数，任何宿主都能用），可以照抄 Excel 版这部分逻辑；Excel 版另外
' 几个 Ask*（AskRange/AskFolder 等）耦合了 Range/Worksheet/modIO，
' Word 这边还没有调用方，不提前搭。
'==============================================================================
Option Explicit
Option Private Module

Public Const ERR_CANCELLED As Long = vbObjectError + 999

Private mSilent As Boolean
Private mParams As Object            ' Scripting.Dictionary，静默模式下的预设参数

Public Sub SetSilent(ByVal value As Boolean)
    mSilent = value
End Sub

Public Function IsSilent() As Boolean
    IsSilent = mSilent
End Function

Public Sub SetParam(ByVal key As String, ByVal value As String)
    EnsureParams
    mParams(key) = value
End Sub

Public Sub ClearParams()
    Set mParams = Nothing
End Sub

Private Sub EnsureParams()
    If mParams Is Nothing Then
        Set mParams = CreateObject("Scripting.Dictionary")
        mParams.CompareMode = vbTextCompare
    End If
End Sub

Private Function Preset(ByVal key As String, ByRef found As Boolean) As String
    EnsureParams
    found = mParams.Exists(key)
    If found Then Preset = CStr(mParams(key))
End Function

Public Sub Cancel()
    Err.Raise ERR_CANCELLED, "modPrompt", "用户取消"
End Sub

'------------------------------------------------------------------------------
' 文本参数。和 Excel 版 AskText 同一个逻辑：静默模式下不弹框，从预设值
' 拿；InputBox 取消和"输入空串后确定"都返回""，无法区分，对不允许空值
' 的参数一律按取消处理。
'------------------------------------------------------------------------------
Public Function AskText(ByVal key As String, _
                        ByVal message As String, _
                        Optional ByVal defaultValue As String = "", _
                        Optional ByVal allowBlank As Boolean = False) As String
    Dim hit As Boolean, preVal As String
    preVal = Preset(key, hit)
    If mSilent Then
        If Not hit Then Cancel
        AskText = preVal
        Exit Function
    End If

    Dim answer As String
    answer = InputBox(message, modApp.APP_NAME, defaultValue)

    If Len(answer) = 0 And Not allowBlank Then Cancel

    AskText = answer
End Function
