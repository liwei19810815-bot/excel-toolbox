Attribute VB_Name = "modPublic"
'==============================================================================
' modPublic - 对外暴露的入口
'
' 只有这里的过程允许被 Application.Run 调用，构建/测试脚本从这里进。
' 本模块【不能】加 Option Private Module。
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 自检。构建后由构建/测试脚本调用，触发 VBA 编译——Import 只是把源码
' 塞进工程，语法错误要等第一次执行时才暴露。
'------------------------------------------------------------------------------
Public Function PptToolbox_SelfCheck() As String
    PptToolbox_SelfCheck = "OK|" & modApp.APP_NAME & "|" & modApp.APP_VERSION
End Function
