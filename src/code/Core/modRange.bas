Attribute VB_Name = "modRange"
'==============================================================================
' modRange - 区域工具
'
' 集中处理 Excel 区域操作里几个反复咬人的坑：
'   1. UsedRange 会被"删过内容但留了格式"的单元格撑大，必须用 Find 重新定位边界；
'   2. 在循环里 Union 是 O(n^2)，几千个区域就会卡死，必须分块累积；
'   3. Range("A1,A5,...") 的地址字符串有 255 字符上限，超了直接报错；
'   4. 逐单元格读写比一次性读入数组慢一到两个数量级。
'==============================================================================
Option Explicit
Option Private Module

'==============================================================================
' 边界定位
'==============================================================================

'------------------------------------------------------------------------------
' 真实已用区域。ws.UsedRange 会把"清过内容但残留格式/批注"的单元格算进去，
' 导致动辄几万空行。这里用 Find 从两个方向找真正有内容的最后一行/列。
' 整表为空时返回 Nothing。
'------------------------------------------------------------------------------
Public Function RealUsedRange(ByVal ws As Worksheet) As Range
    Dim fallback As Range
    On Error Resume Next
    Set fallback = ws.UsedRange
    On Error GoTo 0
    If fallback Is Nothing Then Exit Function

    Dim lastRow As Long, lastCol As Long
    lastRow = LastDataRow(ws)
    lastCol = LastDataColumn(ws)

    ' Find 并不总是可靠——起点落在合并单元格上时它会【静默返回 Nothing】。
    ' 这时候必须退回 UsedRange：宁可多算上几行残留格式，也绝不能把一张
    ' 有数据的表判定成空表，那会让所有工具都直接罢工。
    If lastRow = 0 Or lastCol = 0 Then
        If Application.WorksheetFunction.CountA(fallback) = 0 Then Exit Function
        Set RealUsedRange = fallback
        Exit Function
    End If

    Dim firstRow As Long, firstCol As Long
    firstRow = FirstDataRow(ws)
    firstCol = FirstDataColumn(ws)

    Set RealUsedRange = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
End Function

' After 用右下角那个单元格而不是 A1：Find 的起点一旦落在合并单元格里，
' 它会静默返回 Nothing，导致整张表被判定为空。A1 恰好被合并是很常见的
' （表头合并标题），这个坑不绕过去，合并表就全都用不了。
Public Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not c Is Nothing Then LastDataRow = c.Row
End Function

Public Function LastDataColumn(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                          LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If Not c Is Nothing Then LastDataColumn = c.Column
End Function

Public Function FirstDataRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlNext)
    On Error GoTo 0
    If c Is Nothing Then FirstDataRow = 1 Else FirstDataRow = c.Row
End Function

Public Function FirstDataColumn(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(ws.Rows.Count, ws.Columns.Count), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    On Error GoTo 0
    If c Is Nothing Then FirstDataColumn = 1 Else FirstDataColumn = c.Column
End Function

'------------------------------------------------------------------------------
' 用户选区归一化：选中整行/整列/整表时，收敛到真实已用区域，
' 避免对一百万行做无意义的遍历。
'------------------------------------------------------------------------------
Public Function NormalizeSelection(ByVal rng As Range) As Range
    If rng Is Nothing Then Exit Function

    Dim used As Range
    Set used = RealUsedRange(rng.Worksheet)
    If used Is Nothing Then Exit Function

    On Error Resume Next
    Set NormalizeSelection = Application.Intersect(rng, used)
    On Error GoTo 0
End Function

'==============================================================================
' 多区域累积见 clsAreaAccumulator（循环 Union 是 O(n^2)，必须分块）
'==============================================================================

'==============================================================================
' 数组读写
'==============================================================================

'------------------------------------------------------------------------------
' 区域 -> 二维数组（下标恒为 1 基的 (1 To rows, 1 To cols)）。
' 单个单元格时 .Value 返回的是标量而非数组，必须单独处理——这是经典崩点。
'------------------------------------------------------------------------------
Public Function ToArray(ByVal rng As Range) As Variant
    If rng.Cells.Count = 1 Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = rng.Value
        ToArray = one
    Else
        ToArray = rng.Value
    End If
End Function

'------------------------------------------------------------------------------
' 把处理结果写回区域，但【不碰公式单元格】。
'
' 这是文本类工具最容易造成不可逆损失的地方：即使计算时跳过了公式单元格，
' 只要最后用 FromArray 整块写回，公式就会被它自己的计算结果覆盖掉，
' 而且看不出异常——直到某天源数据变了，结果不再更新。
'
' 区域内没有公式时走整块写回的快路径；有公式时逐格写，只写真正变了的单元格。
' 返回实际写入的单元格数。
'------------------------------------------------------------------------------
Public Function WriteBack(ByVal areaRng As Range, _
                          ByRef srcArr As Variant, _
                          ByRef outArr As Variant) As Long
    Dim hasFormula As Variant
    hasFormula = areaRng.HasFormula

    ' HasFormula 返回 True/False/Null（Null = 混合）
    If VarType(hasFormula) = vbBoolean Then
        If Not CBool(hasFormula) Then
            FromArray areaRng, outArr
            WriteBack = CLng(CellCount(areaRng))
            Exit Function
        End If
    End If

    Dim rowIdx As Long, colIdx As Long, written As Long
    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
            If Not ValuesEqual(srcArr(rowIdx, colIdx), outArr(rowIdx, colIdx)) Then
                If Not areaRng.Cells(rowIdx, colIdx).HasFormula Then
                    areaRng.Cells(rowIdx, colIdx).Value = outArr(rowIdx, colIdx)
                    written = written + 1
                End If
            End If
        Next colIdx
    Next rowIdx

    WriteBack = written
End Function

' 【类型不同就必须算"不相等"】。原先只比 CStr(a) = CStr(b)，
' 于是文本 "4" 和数值 4 被判成相等——WriteBack 跳过不写，
' 「文本转数值」对这种单元格【完全不起作用】。
'
' 而它只在选区里【同时含有公式】时才发作：没有公式时 WriteBack 走
' FromArray 整块写回的快路径，根本不做逐格比较，一切正常。
' 一旦选区里有一列公式（很常见），就落到逐格路径上，
' 所有"纯数字文本"全部被静默跳过。
'
' 最恶劣的地方是它不报错：TextToNumber 报的是【打算转换的个数】，
' 不是实际写回的个数，所以用户看到"已处理 94 个单元格"，
' 数据却一格没变。没有任何测试会因此变红。
Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then
        ValuesEqual = (IsError(a) And IsError(b))
    ElseIf IsNull(a) Or IsNull(b) Then
        ValuesEqual = (IsNull(a) And IsNull(b))
    ElseIf VarType(a) <> VarType(b) Then
        ValuesEqual = False
    Else
        ValuesEqual = (CStr(a) = CStr(b))
    End If
End Function

'------------------------------------------------------------------------------
' 二维数组 -> 区域。以数组尺寸为准重新框定目标左上角，避免尺寸不匹配时静默截断。
'
' 会覆盖公式。需要保留公式时用 WriteBack。
'------------------------------------------------------------------------------
Public Sub FromArray(ByVal topLeft As Range, ByRef arr As Variant)
    Dim nR As Long, nC As Long
    nR = UBound(arr, 1) - LBound(arr, 1) + 1
    nC = UBound(arr, 2) - LBound(arr, 2) + 1
    topLeft.Cells(1, 1).Resize(nR, nC).Value = arr
End Sub

'==============================================================================
' 其它
'==============================================================================

' 可见单元格；全部隐藏时返回 Nothing 而不是报错
Public Function VisibleCells(ByVal rng As Range) As Range
    On Error Resume Next
    Set VisibleCells = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
End Function

' 区域内是否存在合并单元格。很多数据类工具遇到合并单元格必须先拦下来。
Public Function HasMergedCells(ByVal rng As Range) As Boolean
    Dim v As Variant
    v = rng.MergeCells
    ' 混合状态时返回 Null
    HasMergedCells = IsNull(v) Or (v = True)
End Function

' 单元格数量（用 CDbl 避免超过 Long 上限的大区域溢出）
Public Function CellCount(ByVal rng As Range) As Double
    Dim a As Range, n As Double
    For Each a In rng.Areas
        n = n + CDbl(a.Rows.Count) * CDbl(a.Columns.Count)
    Next a
    CellCount = n
End Function
