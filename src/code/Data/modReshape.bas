Attribute VB_Name = "modReshape"
'==============================================================================
' modReshape - 表结构变形（M2）
'
' 二维表转一维表（逆透视）是做数据分析前最常见的一步：
' 交叉表人看着方便，但透视表、Power Query、数据库全都要一维明细表。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 二维表 -> 一维表（逆透视）
'
' 输入形如：
'     产品   1月   2月   3月
'     A      10    20    30
' 输出：
'     产品   项目   值
'     A      1月    10
'     A      2月    20
'
' 左侧保留几列由用户指定（通常是维度列），其余列全部展开为"项目 / 值"两列。
' 空值默认跳过——展开后空值行只会稀释数据，没有分析价值。
'------------------------------------------------------------------------------
Public Function Unpivot(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        Unpivot = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Rows.Count < 2 Or srcRng.Columns.Count < 2 Then
        Unpivot = "逆透视至少需要 2 行 2 列（含标题行）。"
        Exit Function
    End If

    Dim keepCols As Long
    keepCols = modPrompt.AskNumber("keepColumns", _
        "左侧保留几列作为维度列？（其余列将展开为「项目」「值」两列）", 1)

    If keepCols < 1 Or keepCols >= srcRng.Columns.Count Then
        Err.Raise vbObjectError + 370, "modReshape", _
                  "保留列数必须在 1 到 " & (srcRng.Columns.Count - 1) & " 之间。"
    End If

    Dim skipBlank As Boolean
    skipBlank = modPrompt.AskYesNo("skipBlank", "跳过空值吗？（建议是）")

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim totalRows As Long, totalCols As Long
    totalRows = UBound(srcArr, 1)
    totalCols = UBound(srcArr, 2)

    ' 先算出输出行数，一次性分配数组，避免 ReDim Preserve 反复拷贝
    Dim outCount As Long, rowIdx As Long, colIdx As Long
    For rowIdx = 2 To totalRows
        For colIdx = keepCols + 1 To totalCols
            If Not (skipBlank And IsBlankCell(srcArr(rowIdx, colIdx))) Then
                outCount = outCount + 1
            End If
        Next colIdx
    Next rowIdx

    If outCount = 0 Then
        Unpivot = "没有可展开的数据。"
        Exit Function
    End If

    Dim outCols As Long
    outCols = keepCols + 2

    Dim outArr() As Variant
    ReDim outArr(1 To outCount + 1, 1 To outCols)

    ' 表头
    For colIdx = 1 To keepCols
        outArr(1, colIdx) = srcArr(1, colIdx)
    Next colIdx
    outArr(1, keepCols + 1) = "项目"
    outArr(1, keepCols + 2) = "值"

    Dim outIdx As Long, keepIdx As Long
    outIdx = 2
    For rowIdx = 2 To totalRows
        For colIdx = keepCols + 1 To totalCols
            If Not (skipBlank And IsBlankCell(srcArr(rowIdx, colIdx))) Then
                For keepIdx = 1 To keepCols
                    outArr(outIdx, keepIdx) = srcArr(rowIdx, keepIdx)
                Next keepIdx
                outArr(outIdx, keepCols + 1) = srcArr(1, colIdx)     ' 原列标题
                outArr(outIdx, keepCols + 2) = srcArr(rowIdx, colIdx)
                outIdx = outIdx + 1
            End If
        Next colIdx
    Next rowIdx

    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(srcRng.Worksheet.Parent, "一维表")
    outWs.Range("A1").Resize(outCount + 1, outCols).Value = outArr
    modSheetUtil.FormatHeader outWs, outCols

    Unpivot = "已转换为一维表，共 " & outCount & " 行，结果在工作表「" & outWs.Name & "」。"
End Function

Private Function IsBlankCell(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankCell = False
    ElseIf IsNull(v) Then
        IsBlankCell = True
    Else
        IsBlankCell = (Len(CStr(v)) = 0)
    End If
End Function

'------------------------------------------------------------------------------
' 行列转置到新表。
'
' 不用选择性粘贴的转置：那个在源区域和目标区域重叠时会直接报错，
' 而且会把公式里的相对引用一起转过去，结果往往是错的。这里只转值。
'------------------------------------------------------------------------------
Public Function TransposeRange(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        TransposeRange = "选区内没有数据。"
        Exit Function
    End If

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim rowCount As Long, colCount As Long
    rowCount = UBound(srcArr, 1)
    colCount = UBound(srcArr, 2)

    Dim outArr() As Variant
    ReDim outArr(1 To colCount, 1 To rowCount)

    Dim rowIdx As Long, colIdx As Long
    For rowIdx = 1 To rowCount
        For colIdx = 1 To colCount
            outArr(colIdx, rowIdx) = srcArr(rowIdx, colIdx)
        Next colIdx
    Next rowIdx

    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(srcRng.Worksheet.Parent, "转置")
    outWs.Range("A1").Resize(colCount, rowCount).Value = outArr
    outWs.Columns.AutoFit

    TransposeRange = "已转置 " & rowCount & " 行 × " & colCount & " 列，结果在工作表「" & outWs.Name & "」。"
End Function
