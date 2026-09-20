Attribute VB_Name = "modViz"
'==============================================================================
' modViz - 数据可视化（M8）
'
' 这些工具都只作用于【数值单元格】。把数据条套到文本列上不会报错，
' 但会画出一片毫无意义的色块，所以统一先筛出数值区域再动手。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 取出选区内的数值单元格
'------------------------------------------------------------------------------
Private Function NumericCells(ByVal srcRng As Range) As Range
    Dim result As Range
    On Error Resume Next
    Set result = srcRng.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo 0

    Dim formulaNums As Range
    On Error Resume Next
    Set formulaNums = srcRng.SpecialCells(xlCellTypeFormulas, xlNumbers)
    On Error GoTo 0

    If result Is Nothing Then
        Set NumericCells = formulaNums
    ElseIf formulaNums Is Nothing Then
        Set NumericCells = result
    Else
        Set NumericCells = Application.Union(result, formulaNums)
    End If
End Function

'------------------------------------------------------------------------------
' 数据条
'------------------------------------------------------------------------------
Public Function AddDataBars(ByVal target As Range) As String
    Dim numRng As Range
    Set numRng = PrepareNumeric(target)
    If numRng Is Nothing Then
        AddDataBars = "选区内没有数值。"
        Exit Function
    End If

    modUndo.Capture numRng

    Dim db As Databar
    Set db = numRng.FormatConditions.AddDatabar
    db.BarColor.Color = RGB(99, 142, 198)
    db.BarFillType = xlDataBarFillGradient

    ' 负值单独配色，否则正负两个方向都是蓝色，看不出差别
    db.NegativeBarFormat.ColorType = xlDataBarColor
    db.NegativeBarFormat.Color.Color = RGB(214, 99, 99)

    AddDataBars = "已为 " & CLng(modRange.CellCount(numRng)) & " 个数值单元格添加数据条。"
End Function

'------------------------------------------------------------------------------
' 三色色阶（热力图）
'------------------------------------------------------------------------------
Public Function AddColorScale(ByVal target As Range) As String
    Dim numRng As Range
    Set numRng = PrepareNumeric(target)
    If numRng Is Nothing Then
        AddColorScale = "选区内没有数值。"
        Exit Function
    End If

    modUndo.Capture numRng

    Dim cs As ColorScale
    Set cs = numRng.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)

    AddColorScale = "已添加三色色阶（低=绿 中=黄 高=红）。"
End Function

'------------------------------------------------------------------------------
' 图标集
'------------------------------------------------------------------------------
Public Function AddIconSet(ByVal target As Range) As String
    Dim numRng As Range
    Set numRng = PrepareNumeric(target)
    If numRng Is Nothing Then
        AddIconSet = "选区内没有数值。"
        Exit Function
    End If

    modUndo.Capture numRng

    Dim ic As IconSetCondition
    Set ic = numRng.FormatConditions.AddIconSetCondition
    ic.IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)

    AddIconSet = "已添加三色交通灯图标集。"
End Function

'------------------------------------------------------------------------------
' 清除选区内所有条件格式（上面三个的撤销出口之外，再给一个直接的清除入口）
'------------------------------------------------------------------------------
Public Function ClearConditionalFormats(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        ClearConditionalFormats = "选区内没有数据。"
        Exit Function
    End If

    Dim n As Long
    n = srcRng.FormatConditions.Count
    If n = 0 Then
        ClearConditionalFormats = "选区内没有条件格式。"
        Exit Function
    End If

    modUndo.Capture srcRng
    srcRng.FormatConditions.Delete

    ClearConditionalFormats = "已清除 " & n & " 条条件格式规则。"
End Function

Private Function PrepareNumeric(ByVal target As Range) As Range
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then Exit Function
    Set PrepareNumeric = NumericCells(srcRng)
End Function

'------------------------------------------------------------------------------
' 批量生成迷你图（每行一个，画在数据右侧一格）
'------------------------------------------------------------------------------
Public Function AddSparklines(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        AddSparklines = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Columns.Count < 2 Then
        AddSparklines = "迷你图至少需要 2 列数据。"
        Exit Function
    End If

    Dim typeChoice As Long
    Dim types(1 To 3) As String
    types(1) = "折线图"
    types(2) = "柱形图"
    types(3) = "盈亏图"
    typeChoice = modPrompt.AskChoice("sparkType", "生成哪种迷你图？", types)

    Dim sparkType As Long
    Select Case typeChoice
        Case 1: sparkType = xlSparkLine
        Case 2: sparkType = xlSparkColumn
        Case 3: sparkType = xlSparkColumnStacked100
    End Select

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim targetCol As Long
    targetCol = srcRng.Column + srcRng.Columns.Count

    ' 迷你图画在数据右侧一列，先把那一列快照下来
    Dim outRng As Range
    Set outRng = ws.Cells(srcRng.Row, targetCol).Resize(srcRng.Rows.Count, 1)
    modUndo.Capture outRng

    Dim phase As String
    On Error GoTo Failed

    ' 迷你图的创建依赖界面，ScreenUpdating 关着时 Add 会直接抛 1004。
    ' 这里临时打开，做完还原——注意不能走 modPerf.FastModeOff，
    ' 那个带嵌套计数，会把整条管线的高速模式一起关掉。
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True

    Dim rowIdx As Long, dataRng As Range, cellRng As Range
    For rowIdx = 1 To srcRng.Rows.Count
        Set dataRng = srcRng.Rows(rowIdx)
        Set cellRng = ws.Cells(srcRng.Row + rowIdx - 1, targetCol)

        phase = "清空 " & cellRng.Address(False, False)
        ' 集合为空时 Clear 会抛 1004，必须先判断
        If cellRng.SparklineGroups.Count > 0 Then cellRng.SparklineGroups.Clear

        phase = "添加 " & cellRng.Address(False, False) & " <- " & dataRng.Address(False, False) & " type=" & sparkType
        ' 用位置参数而不是 Type:=／SourceData:=，避开 Type 这个 VBA 关键字做具名参数
        cellRng.SparklineGroups.Add sparkType, dataRng.Address
    Next rowIdx

    Application.ScreenUpdating = prevScreenUpdating

    AddSparklines = "已生成 " & srcRng.Rows.Count & " 个迷你图，位于第 " & targetCol & " 列。"
    Exit Function

Failed:
    Application.ScreenUpdating = prevScreenUpdating
    Err.Raise Err.Number, "modViz.AddSparklines[" & phase & "]", Err.Description
End Function

'------------------------------------------------------------------------------
' 快速生成图表。
'
' 统一套一份克制的格式：去掉图表区边框和网格线杂色、字号统一、
' 单系列时不显示图例（只有一个系列还放图例纯属占地方）。
'------------------------------------------------------------------------------
Public Function QuickChart(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        QuickChart = "选区内没有数据。"
        Exit Function
    End If

    Dim typeChoice As Long
    Dim types(1 To 4) As String
    types(1) = "柱形图"
    types(2) = "折线图"
    types(3) = "饼图"
    types(4) = "条形图"
    typeChoice = modPrompt.AskChoice("chartType", "生成哪种图表？", types)

    Dim chartKind As Long
    Select Case typeChoice
        Case 1: chartKind = xlColumnClustered
        Case 2: chartKind = xlLine
        Case 3: chartKind = xlPie
        Case 4: chartKind = xlBarClustered
    End Select

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=srcRng.Left + srcRng.Width + 20, _
                                  Top:=srcRng.Top, Width:=420, Height:=260)
    cht.Chart.SetSourceData Source:=srcRng
    cht.Chart.ChartType = chartKind

    ApplyChartStyle cht.Chart

    QuickChart = "已生成" & types(typeChoice) & "。"
End Function

Private Sub ApplyChartStyle(ByVal cht As Chart)
    On Error Resume Next

    cht.ChartArea.Format.Line.Visible = msoFalse
    cht.ChartArea.Font.Size = 10
    cht.PlotArea.Format.Fill.Visible = msoFalse

    ' 只有一个系列时图例没有信息量
    If cht.SeriesCollection.Count <= 1 Then
        cht.HasLegend = False
    Else
        cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionBottom
    End If

    ' 网格线压淡，不要和数据抢注意力
    cht.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 把选定图表的格式套用到同一工作表的其他图表上（图表格式刷）
'------------------------------------------------------------------------------
Public Function UnifyCharts(ByVal ws As Worksheet) As String
    If ws.ChartObjects.Count = 0 Then
        UnifyCharts = "当前工作表没有图表。"
        Exit Function
    End If

    Dim i As Long
    For i = 1 To ws.ChartObjects.Count
        ApplyChartStyle ws.ChartObjects(i).Chart
    Next i

    UnifyCharts = "已统一 " & ws.ChartObjects.Count & " 个图表的格式。"
End Function
