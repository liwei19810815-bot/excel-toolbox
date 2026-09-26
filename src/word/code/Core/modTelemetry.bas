Attribute VB_Name = "modTelemetry"
'==============================================================================
' modTelemetry - 遥测（Word 最小实现：空操作）
'
' shared/code/Core/modAction.bas 的 RunAction 在每条执行路径上都无条件
' 调用 modTelemetry.TrackAction，这个符号必须存在才能编译。
'
' 【暂时空操作，不落盘不上报】：Excel 那份 modTelemetry.bas
' （718 行）不是可以直接照抄的骨架——之前 Codex 复审就挑出过它耦合了
' Application.Version，真要给 Word 建一条同样完整的遥测管线（本地缓冲、
' 定时上报、失败重试）是一次独立的工程投入，不该为了让三个样板命令
' 编译通过就顺手搭一半。等 Word 命令集到了需要真实使用数据的规模，
' 再照 Excel 那份的思路单独做，不是现在。
'==============================================================================
Option Explicit
Option Private Module

Public Sub TrackAction(ByVal actionId As String, _
                       ByVal outcome As String, _
                       ByVal elapsedMs As Long, _
                       Optional ByVal errNum As Long = 0, _
                       Optional ByVal errDesc As String = "")
    ' 有意空操作，见文件头部说明。
End Sub
