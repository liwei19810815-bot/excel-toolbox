Attribute VB_Name = "modPrompt"
'==============================================================================
' modPrompt - 参数采集层
'
' 为什么需要这一层：
'   1. 很多工具要参数（分隔符、要比对的列、前后缀……）。用 UserForm 的话，
'      .frm 会带一个 .frx 二进制资源文件，放进"源码入 git、构建时导入"的流程里
'      既不能 diff 也容易损坏，收益不抵成本。所以 v1 一律用输入框。
'   2. 输入框在无头运行时会弹在一个不可见的 Excel 里，谁也点不到，测试永久挂死。
'      所以静默模式下改为从预设值里取——测试先 Toolbox_SetParam 塞好参数，
'      业务代码走的还是同一条路径。
'
' 取消的处理：
'   用户点了取消就抛 ERR_CANCELLED，由 RunAction 统一识别为"静默放弃"，
'   不报错、不留撤销步骤。业务代码因此不用在每个参数后面写一遍 If 判断。
'==============================================================================
Option Explicit
Option Private Module

Public Const ERR_CANCELLED As Long = vbObjectError + 999

Private mSilent As Boolean
Private mParams As Object            ' Scripting.Dictionary，静默模式下的预设参数

Public Sub SetSilent(ByVal value As Boolean)
    mSilent = value
End Sub

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
' 文本参数
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
    answer = InputBox(message, APP_NAME, defaultValue)

    ' InputBox 取消和"输入空串后确定"都返回 ""，无法区分。
    ' 对不允许空值的参数，一律按取消处理——这比默默用空串去改数据安全。
    If Len(answer) = 0 And Not allowBlank Then Cancel

    AskText = answer
End Function

'------------------------------------------------------------------------------
' 整数参数
'------------------------------------------------------------------------------
Public Function AskNumber(ByVal key As String, _
                          ByVal message As String, _
                          Optional ByVal defaultValue As Long = 1) As Long
    Dim raw As String
    raw = AskText(key, message, CStr(defaultValue))
    If Not IsNumeric(raw) Then
        Err.Raise vbObjectError + 300, "modPrompt.AskNumber", "「" & raw & "」不是有效的数字。"
    End If
    AskNumber = CLng(raw)
End Function

'------------------------------------------------------------------------------
' 是/否
'------------------------------------------------------------------------------
Public Function AskYesNo(ByVal key As String, ByVal message As String) As Boolean
    Dim hit As Boolean, preVal As String
    preVal = Preset(key, hit)
    If mSilent Then
        If Not hit Then Cancel
        AskYesNo = (LCase$(preVal) = "true" Or preVal = "1" Or preVal = "-1")
        Exit Function
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox(message, vbQuestion + vbYesNoCancel, APP_NAME)
    If answer = vbCancel Then Cancel
    AskYesNo = (answer = vbYes)
End Function

'------------------------------------------------------------------------------
' 区域参数。静默模式下预设值写地址字符串，如 "Sheet1!A1:D10" 或 "A1:D10"。
'------------------------------------------------------------------------------
Public Function AskRange(ByVal key As String, ByVal message As String) As Range
    Dim hit As Boolean, preVal As String
    preVal = Preset(key, hit)
    If mSilent Then
        If Not hit Then Cancel
        On Error GoTo BadAddress
        If InStr(preVal, "!") > 0 Then
            Set AskRange = ActiveWorkbook.Worksheets(Split(preVal, "!")(0)).Range(Split(preVal, "!")(1))
        Else
            Set AskRange = ActiveSheet.Range(preVal)
        End If
        Exit Function
BadAddress:
        Err.Raise vbObjectError + 301, "modPrompt.AskRange", "无效的区域地址：" & preVal
    End If

    Dim picked As Range
    On Error Resume Next
    Set picked = Application.InputBox(message, APP_NAME, Type:=8)
    On Error GoTo 0
    If picked Is Nothing Then Cancel

    Set AskRange = picked
End Function

'------------------------------------------------------------------------------
' 文件夹参数。
'
' 必须和其它参数一样走这一层：文件夹选择框同样是模态的，
' 无头运行时会弹在一个看不见的 Excel 里，把整个批处理永久挂死。
'------------------------------------------------------------------------------
Public Function AskFolder(ByVal key As String, ByVal message As String) As String
    Dim hit As Boolean, preVal As String
    preVal = Preset(key, hit)
    If mSilent Then
        If Not hit Then Cancel
        If Not modIO.FolderExists(preVal) Then
            Err.Raise vbObjectError + 303, "modPrompt.AskFolder", "文件夹不存在：" & preVal
        End If
        AskFolder = preVal
        Exit Function
    End If

    AskFolder = modIO.PickFolder(message)
End Function

'------------------------------------------------------------------------------
' 从若干选项里挑一个，返回 1 基序号。
' 用编号输入而不是下拉框——同样是为了不引入 UserForm。
'------------------------------------------------------------------------------
Public Function AskChoice(ByVal key As String, _
                          ByVal message As String, _
                          ByRef choices() As String) As Long
    Dim listText As String, i As Long
    For i = LBound(choices) To UBound(choices)
        listText = listText & (i - LBound(choices) + 1) & ". " & choices(i) & vbCrLf
    Next i

    Dim picked As Long
    picked = AskNumber(key, message & vbCrLf & vbCrLf & listText & vbCrLf & "请输入序号：", 1)

    If picked < 1 Or picked > (UBound(choices) - LBound(choices) + 1) Then
        Err.Raise vbObjectError + 302, "modPrompt.AskChoice", "序号超出范围：" & picked
    End If
    AskChoice = picked
End Function
