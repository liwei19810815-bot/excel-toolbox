Attribute VB_Name = "modRibbon"
'==============================================================================
' modRibbon - Word 工具箱功能区回调（骨架阶段）
'
' 骨架阶段只证明"customUI 被正确加载、按钮点得动"这条链路，回调直接弹
' AboutText，不经过 modAction/RunAction 管线——那套要等真正的命令集
' （下一步"Word 样板命令"）引入时再接上，提前搭一个没有第二个调用方的
' 管线是没人验证过的抽象。
'
' 【本模块不能加 Option Private Module】：Ribbon 回调由 Word 按名字
' 查找调用，和 Excel/PPT modRibbon 的约束一致。
'==============================================================================
Option Explicit

Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    ' 骨架阶段还没有需要 InvalidateControl 的动态按钮，先不缓存指针。
End Sub

Public Sub Ribbon_OnAction(control As IRibbonControl)
    Select Case control.Tag
        Case "core.about"
            MsgBox modApp.AboutText(), vbInformation, modApp.APP_NAME
        Case Else
            MsgBox "未识别的命令：" & control.Tag, vbExclamation, modApp.APP_NAME
    End Select
End Sub
