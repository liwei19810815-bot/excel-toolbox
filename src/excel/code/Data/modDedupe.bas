Attribute VB_Name = "modDedupe"
'==============================================================================
' modDedupe - 重复值处理（M2）
'
' 判重一律走 Dictionary（哈希），不做两两比较：10 万行两两比较是 50 亿次，
' 走哈希是 10 万次。
'
' 组合键用 vbNullChar 做分隔符——它不会出现在正常单元格数据里。
' 如果用逗号之类的可见字符，("a,b", "c") 和 ("a", "b,c") 会被判成同一行。
'==============================================================================
Option Explicit
Option Private Module

Private Const KEY_SEP As String = vbNullChar

'------------------------------------------------------------------------------
' 把一行的若干列拼成组合键
'------------------------------------------------------------------------------
Private Function RowKey(ByRef srcArr As Variant, ByVal rowIdx As Long, ByRef cols() As Long) As String
    Dim i As Long, buf As String
    For i = LBound(cols) To UBound(cols)
        If i > LBound(cols) Then buf = buf & KEY_SEP
        buf = buf & CStr(NzText(srcArr(rowIdx, cols(i))))
    Next i
    RowKey = buf
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsError(v) Then
        NzText = "#ERR"
    ElseIf IsNull(v) Then
        NzText = ""
    Else
        NzText = CStr(v)
    End If
End Function

'------------------------------------------------------------------------------
' 询问按哪些列判重。空输入 = 全部列参与。
'------------------------------------------------------------------------------
Private Function AskKeyColumns(ByVal colCount As Long) As Long()
    Dim answer As String
    answer = modPrompt.AskText("keyColumns", _
        "按哪些列判断重复？填选区内的列序号，逗号分隔（例如 1,3）。" & vbCrLf & _
        "留空表示所有列都要相同才算重复。", "", True)

    Dim cols() As Long
    If Len(Trim$(answer)) = 0 Then
        ReDim cols(1 To colCount)
        Dim i As Long
        For i = 1 To colCount
            cols(i) = i
        Next i
    Else
        Dim parts As Variant
        parts = Split(answer, ",")
        ReDim cols(1 To UBound(parts) + 1)
        For i = 0 To UBound(parts)
            If Not IsNumeric(Trim$(parts(i))) Then
                Err.Raise vbObjectError + 340, "modDedupe", "「" & parts(i) & "」不是有效的列序号。"
            End If
            cols(i + 1) = CLng(Trim$(parts(i)))
            If cols(i + 1) < 1 Or cols(i + 1) > colCount Then
                Err.Raise vbObjectError + 341, "modDedupe", _
                          "列序号 " & cols(i + 1) & " 超出选区范围（选区共 " & colCount & " 列）。"
            End If
        Next i
    End If

    AskKeyColumns = cols
End Function

'------------------------------------------------------------------------------
' 标记重复行（整行底色）。
'
' 默认不标记首次出现的那一行——用户要看的是"多出来的那些"。
'------------------------------------------------------------------------------
Public Function MarkDuplicates(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        MarkDuplicates = "选区内没有数据。"
        Exit Function
    End If

    Dim hasHeader As Boolean
    hasHeader = modPrompt.AskYesNo("hasHeader", "选区的第一行是标题行吗？")

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim cols() As Long
    cols = AskKeyColumns(UBound(srcArr, 2))

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim acc As clsAreaAccumulator
    Set acc = New clsAreaAccumulator
    acc.Init ws

    Dim rowIdx As Long, startRow As Long, keyText As String
    startRow = IIf(hasHeader, 2, 1)

    For rowIdx = startRow To UBound(srcArr, 1)
        keyText = RowKey(srcArr, rowIdx, cols)
        If seen.Exists(keyText) Then
            acc.AddRange srcRng.Rows(rowIdx)
        Else
            seen(keyText) = 1
        End If
    Next rowIdx

    If acc.HasNoAreas Then
        MarkDuplicates = "没有找到重复行。"
        Exit Function
    End If

    Dim hitRng As Range
    Set hitRng = acc.Result

    modUndo.Capture hitRng
    hitRng.Interior.Color = RGB(255, 199, 206)      ' Excel 内置"浅红填充"的配色

    MarkDuplicates = "已标记 " & acc.Count & " 个重复行（不含首次出现）。"
End Function

'------------------------------------------------------------------------------
' 删除重复行，保留首次出现。
'------------------------------------------------------------------------------
Public Function DeleteDuplicates(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        DeleteDuplicates = "选区内没有数据。"
        Exit Function
    End If

    If modRange.HasMergedCells(srcRng) Then
        DeleteDuplicates = "选区内存在合并单元格，删除整行的结果不可预期。请先取消合并。"
        Exit Function
    End If

    Dim hasHeader As Boolean
    hasHeader = modPrompt.AskYesNo("hasHeader", "选区的第一行是标题行吗？")

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim cols() As Long
    cols = AskKeyColumns(UBound(srcArr, 2))

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim acc As clsAreaAccumulator
    Set acc = New clsAreaAccumulator
    acc.Init ws

    Dim rowIdx As Long, startRow As Long, keyText As String
    startRow = IIf(hasHeader, 2, 1)

    For rowIdx = startRow To UBound(srcArr, 1)
        keyText = RowKey(srcArr, rowIdx, cols)
        If seen.Exists(keyText) Then
            acc.AddRow srcRng.Row + rowIdx - 1
        Else
            seen(keyText) = 1
        End If
    Next rowIdx

    If acc.HasNoAreas Then
        DeleteDuplicates = "没有找到重复行。"
        Exit Function
    End If

    Dim removed As Long
    removed = acc.Count

    ' 删行是结构性操作
    modUndo.CaptureSheet ws
    acc.Result.EntireRow.Delete

    DeleteDuplicates = "已删除 " & removed & " 个重复行，保留首次出现。"
End Function

'------------------------------------------------------------------------------
' 提取唯一值到新工作表。
'------------------------------------------------------------------------------
Public Function ExtractUnique(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        ExtractUnique = "选区内没有数据。"
        Exit Function
    End If

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    Dim rowIdx As Long, colIdx As Long, txt As String
    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
            txt = NzText(srcArr(rowIdx, colIdx))
            If Len(txt) > 0 Then
                If Not seen.Exists(txt) Then seen(txt) = seen.Count + 1
            End If
        Next colIdx
    Next rowIdx

    If seen.Count = 0 Then
        ExtractUnique = "选区内没有非空值。"
        Exit Function
    End If

    ' 输出到新表：新增工作表本身不可撤销，但不改动任何原始数据，所以是安全的
    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(srcRng.Worksheet.Parent, "唯一值")

    Dim outArr() As Variant
    ReDim outArr(1 To seen.Count, 1 To 1)
    Dim k As Variant, i As Long
    i = 1
    For Each k In seen.Keys
        outArr(i, 1) = k
        i = i + 1
    Next k

    outWs.Range("A1").Value = "唯一值"
    outWs.Range("A1").Font.Bold = True
    outWs.Range("A2").Resize(seen.Count, 1).Value = outArr
    outWs.Columns(1).AutoFit

    ExtractUnique = "已提取 " & seen.Count & " 个唯一值到工作表「" & outWs.Name & "」。"
End Function
