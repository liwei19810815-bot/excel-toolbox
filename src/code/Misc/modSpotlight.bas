Attribute VB_Name = "modSpotlight"
'==============================================================================
' modSpotlight - 聚光灯（十字高亮）（M9）
'
' 实现方式的取舍：
'   方案 A（很多插件在用）：直接改单元格 Interior.Color，移动光标时再改回来。
'     致命问题是会【永久破坏用户原有的底色】——恢复时只能恢复成"无填充"，
'     原来有底色的单元格就此丢失。
'   方案 B（本实现）：用条件格式，公式引用 CELL("row")/CELL("col")。
'     条件格式是叠加在原格式之上的，取消时精确删除自己那一条，
'     用户原有的底色和条件格式都不受影响。
'
' 代价是每次选区变化要触发一次重算（CELL 是易失函数）。因此只在聚光灯
' 开启期间挂事件，关闭即摘钩。
'==============================================================================
Option Explicit
Option Private Module

' 用一个别人不会用到的颜色做标记，取消时据此认出"哪条条件格式是我加的"
Private Const MARKER_COLOR As Long = 14083324        ' RGB(252, 228, 214) 的 Long 值

Private mEnabled As Boolean
Private mEvents As clsAppEvents
Private mLastSheet As String

Public Function IsEnabled() As Boolean
    IsEnabled = mEnabled
End Function

'------------------------------------------------------------------------------
' 开关
'------------------------------------------------------------------------------
Public Function Toggle() As String
    If mEnabled Then
        TurnOff
        Toggle = "聚光灯已关闭。"
    Else
        TurnOn
        Toggle = "聚光灯已开启。移动光标时会高亮所在的整行和整列。" & vbCrLf & vbCrLf & _
                 "再点一次关闭。"
    End If
End Function

Private Sub TurnOn()
    ' 先把条件格式加上，成功了才算"已开启"。
    ' 反过来写的话，Apply 失败（比如表被保护）时状态仍是已开启、事件钩子也挂着，
    ' 按钮显示"开"但屏幕上什么都没有，用户只能困惑。
    If Not ActiveSheet Is Nothing Then
        Apply ActiveSheet
        mLastSheet = ActiveSheet.Name
    End If

    If mEvents Is Nothing Then Set mEvents = New clsAppEvents
    mEvents.Hook
    mEnabled = True
End Sub

Private Sub TurnOff()
    mEnabled = False
    If Not mEvents Is Nothing Then mEvents.Unhook
    Set mEvents = Nothing

    ' 所有打开的工作簿都要清，用户可能在开启期间切换过文件
    Dim wb As Workbook, ws As Worksheet
    On Error Resume Next
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            RemoveFrom ws
        Next ws
    Next wb
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 选区变化时刷新。
'
' 只在换了工作表时才重建条件格式；同一张表内移动光标，靠 CELL 函数的
' 重算自动更新高亮位置，不需要碰条件格式。这一点决定了流畅度。
'------------------------------------------------------------------------------
Public Sub OnSelectionChanged(ByVal target As Range)
    If Not mEnabled Then Exit Sub
    If target Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = target.Worksheet

    If ws.Name <> mLastSheet Then
        Apply ws
        mLastSheet = ws.Name
    End If

    ' CELL 是易失函数，但只有在重算时才更新。轻量地戳一下当前表即可。
    Application.ScreenUpdating = True
    ws.Calculate
End Sub

' 加条件格式的错误【不吞】：失败了就让它抛出去，
' 由 RunAction 统一报给用户，而不是让聚光灯假装开着。
Private Sub Apply(ByVal ws As Worksheet)
    RemoveFrom ws

    Dim used As Range
    Set used = modRange.RealUsedRange(ws)
    If used Is Nothing Then Set used = ws.Range("A1:Z100")

    ' 行高亮和列高亮各一条规则
    Dim fcRow As FormatCondition, fcCol As FormatCondition

    Set fcRow = used.FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=CELL(""row"")=ROW()")
    fcRow.Interior.Color = MARKER_COLOR
    fcRow.StopIfTrue = False

    Set fcCol = used.FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=CELL(""col"")=COLUMN()")
    fcCol.Interior.Color = MARKER_COLOR
    fcCol.StopIfTrue = False
End Sub

'------------------------------------------------------------------------------
' 只删掉聚光灯自己加的那两条规则，不碰用户的条件格式。
'
' 【不能只靠颜色认】：用户完全可能有一条底色恰好相同的条件格式，
' 关闭聚光灯时就把人家的规则删了，而且神不知鬼不觉。
' 所以判据是"公式是我们那两条 + 颜色也对"，两者都满足才删。
'------------------------------------------------------------------------------
Private Sub RemoveFrom(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        If IsOurRule(ws.Cells.FormatConditions(i)) Then
            On Error Resume Next
            ws.Cells.FormatConditions(i).Delete
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function IsOurRule(ByVal fc As Object) As Boolean
    Dim ruleFormula As String, ruleColor As Long

    On Error Resume Next
    ' 非 xlExpression 类型的规则没有 Formula1，取不到就不是我们的
    ruleFormula = fc.Formula1
    If Err.Number <> 0 Then Err.Clear: Exit Function
    ruleColor = fc.Interior.Color
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0

    If ruleColor <> MARKER_COLOR Then Exit Function

    IsOurRule = (InStr(1, ruleFormula, "CELL(""row"")", vbTextCompare) > 0) Or _
                (InStr(1, ruleFormula, "CELL(""col"")", vbTextCompare) > 0)
End Function

' 加载宏卸载时调用，避免把条件格式残留在用户文件里
Public Sub Cleanup()
    If mEnabled Then TurnOff
End Sub
