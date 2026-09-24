Attribute VB_Name = "modApp"
'==============================================================================
' modApp - PPT 工具箱的应用级常量与自述信息
'
' 这是骨架阶段（二期第二步：构建管线）的最小实现，只为了证明
' 构建 -> 装载 -> 功能区 -> 回调这条链路走得通。真正的命令注册表
' （modAction/clsActionDef，仿 Excel 那套）留到下一步"PPT 样板命令"
' 再引入——这里没有调用方，提前搭起来就是没人验证过的抽象。
'==============================================================================
Option Explicit

Public Const APP_NAME As String = "PPT 工具箱"
Public Const APP_VERSION As String = "0.1.0（骨架）"

Public Function AboutText() As String
    ' 【不用 ActivePresentation.Name】：那是用户当前打开的文件，不是加载项
    ' 自身的身份。PowerPoint 的 VBA 加载项没有 Excel ThisWorkbook 那种
    ' "指向承载代码的那个文件"的隐式全局对象，这里先不报文件名，
    ' 免得报错或者报出一个和加载项无关的名字。
    AboutText = APP_NAME & "  v" & APP_VERSION & vbCrLf & vbCrLf & _
                "宿主：" & Application.Name & " " & Application.Version & vbCrLf & _
                "发布：CBG合同管理部"
End Function
