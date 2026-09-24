Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - PPT 工具箱功能区回调（骨架阶段）
'
' 骨架阶段只证明"customUI 被正确加载、按钮点得动"这条链路，回调直接弹
' AboutText，不经过 modAction/RunAction 管线——那套（撤销事务、高速模式、
' 统一确认框）要等真正的命令集（下一步"PPT 样板命令"）引入时再接上，
' 提前搭一个没有第二个调用方的管线是没人验证过的抽象。
'
' 【本模块不能加 Option Private Module】：Ribbon 回调由 PowerPoint 按名字
' 查找调用，和 Excel modRibbon 的约束一致。
'==============================================================================
Option Explicit

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    ' 骨架阶段还没有需要 InvalidateControl 的动态按钮，先不缓存指针——
    ' 等真的有开关型按钮或者要点亮/变灰的控件时再照抄 Excel modRibbon
    ' 那套 ObjPtr + Presentation.Tags 持久化方案（已用 COM 实测验证过
    ' Presentation.Tags 可以承载这个指针，是 Excel Workbook.Names 那招
    ' 在 PPT 上的对应写法）。
End Sub

Public Sub Ribbon_OnAction(control As IRibbonControl)
    Select Case control.Tag
        Case "core.about"
            MsgBox modApp.AboutText(), vbInformation, modApp.APP_NAME
        Case Else
            MsgBox "未识别的命令：" & control.Tag, vbExclamation, modApp.APP_NAME
    End Select
End Sub
