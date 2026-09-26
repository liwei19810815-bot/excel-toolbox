Attribute VB_Name = "modApp"
'==============================================================================
' modApp - Word 工具箱的应用级常量与自述信息
'
' 三期第二步（"Word 样板命令"）：命令注册表已接入
' modAction/modActionRegistry/clsActionDef（仿 Excel 那套，RunAction 管线
' 抽到了 shared/code），首批三个命令见 modActionRegistry.bas。
'==============================================================================
Option Explicit

Public Const APP_NAME As String = "Word 工具箱"
Public Const APP_VERSION As String = "0.2.0"

Public Function AboutText() As String
    ' 【不用 ActiveDocument.Name】：那是用户当前打开的文件，不是加载项
    ' 自身的身份。和 PPT 那边同一个理由，这里先不报文件名。
    AboutText = APP_NAME & "  v" & APP_VERSION & vbCrLf & vbCrLf & _
                "宿主：" & Application.Name & " " & Application.Version & vbCrLf & _
                "发布：CBG合同管理部"
End Function
