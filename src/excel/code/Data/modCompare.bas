Attribute VB_Name = "modCompare"
'==============================================================================
' modCompare - 两表数据对比（M2）
'
' 输出一张对比报告表：仅 A 有 / 仅 B 有 / 两边都有但内容不同 / 完全一致的统计。
' 报告里每一行都带可点击的超链接，能直接跳回源数据——对比结果如果不能跳转定位，
' 实际用起来还是要靠人肉找行，价值大打折扣。
'==============================================================================
Option Explicit
Option Private Module

Private Const KEY_SEP As String = vbNullChar

Public Function CompareRanges() As String
    Dim rngA As Range, rngB As Range
    Set rngA = modPrompt.AskRange("rangeA", "选择第一个区域（表 A，含标题行）：")
    Set rngB = modPrompt.AskRange("rangeB", "选择第二个区域（表 B，含标题行）：")

    Dim keyCol As Long
    keyCol = modPrompt.AskNumber("keyColumn", _
        "用第几列作为匹配键？（填区域内的列序号，1 表示第一列）", 1)

    If keyCol < 1 Or keyCol > rngA.Columns.Count Or keyCol > rngB.Columns.Count Then
        Err.Raise vbObjectError + 350, "modCompare", _
                  "键列序号 " & keyCol & " 超出区域范围。"
    End If

    Dim arrA As Variant, arrB As Variant
    arrA = modRange.ToArray(rngA)
    arrB = modRange.ToArray(rngB)

    ' 键 -> 行号。重复键只记首次，并单独统计——静默覆盖会让对比结果失真。
    Dim mapA As Object, mapB As Object
    Set mapA = BuildKeyMap(arrA, keyCol)
    Set mapB = BuildKeyMap(arrB, keyCol)

    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(rngA.Worksheet.Parent, "对比结果")

    outWs.Range("A1:E1").Value = Array("类别", "键值", "表 A 行", "表 B 行", "差异说明")
    modSheetUtil.FormatHeader outWs, 5

    Dim outRow As Long
    outRow = 2

    Dim onlyA As Long, onlyB As Long, diffCount As Long, sameCount As Long
    Dim k As Variant, rowA As Long, rowB As Long, diffText As String

    ' 以 A 为基准扫一遍
    For Each k In mapA.Keys
        rowA = mapA(k)
        If mapB.Exists(k) Then
            rowB = mapB(k)
            diffText = DiffRow(arrA, rowA, arrB, rowB, keyCol)
            If Len(diffText) = 0 Then
                sameCount = sameCount + 1
            Else
                diffCount = diffCount + 1
                WriteRow outWs, outRow, "内容不同", CStr(k), rngA, rowA, rngB, rowB, diffText
                outRow = outRow + 1
            End If
        Else
            onlyA = onlyA + 1
            WriteRow outWs, outRow, "仅表 A 有", CStr(k), rngA, rowA, Nothing, 0, ""
            outRow = outRow + 1
        End If
    Next k

    ' 再补 B 独有的
    For Each k In mapB.Keys
        If Not mapA.Exists(k) Then
            onlyB = onlyB + 1
            WriteRow outWs, outRow, "仅表 B 有", CStr(k), Nothing, 0, rngB, mapB(k), ""
            outRow = outRow + 1
        End If
    Next k

    outWs.Columns.AutoFit
    If outRow > 2 Then outWs.Range("A2").Select

    CompareRanges = "对比完成，结果在工作表「" & outWs.Name & "」。" & vbCrLf & vbCrLf & _
                    "完全一致：" & sameCount & " 行" & vbCrLf & _
                    "内容不同：" & diffCount & " 行" & vbCrLf & _
                    "仅表 A 有：" & onlyA & " 行" & vbCrLf & _
                    "仅表 B 有：" & onlyB & " 行"
End Function

'------------------------------------------------------------------------------
' 建立 键 -> 行号 的映射。跳过标题行和空键。
'------------------------------------------------------------------------------
Private Function BuildKeyMap(ByRef srcArr As Variant, ByVal keyCol As Long) As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare

    Dim rowIdx As Long, keyText As String
    For rowIdx = 2 To UBound(srcArr, 1)          ' 第 1 行是标题
        keyText = CellText(srcArr(rowIdx, keyCol))
        If Len(keyText) > 0 Then
            If Not m.Exists(keyText) Then m(keyText) = rowIdx
        End If
    Next rowIdx

    Set BuildKeyMap = m
End Function

'------------------------------------------------------------------------------
' 逐列比较两行，返回差异说明；完全一致则返回空串。
' 列数不同时只比较公共列数，并在说明里点出来。
'------------------------------------------------------------------------------
Private Function DiffRow(ByRef arrA As Variant, ByVal rowA As Long, _
                         ByRef arrB As Variant, ByVal rowB As Long, _
                         ByVal keyCol As Long) As String
    Dim colsA As Long, colsB As Long, common As Long
    colsA = UBound(arrA, 2)
    colsB = UBound(arrB, 2)
    common = IIf(colsA < colsB, colsA, colsB)

    Dim buf As String, colIdx As Long
    Dim valA As String, valB As String

    For colIdx = 1 To common
        If colIdx <> keyCol Then
            valA = CellText(arrA(rowA, colIdx))
            valB = CellText(arrB(rowB, colIdx))
            If valA <> valB Then
                If Len(buf) > 0 Then buf = buf & "；"
                buf = buf & "第" & colIdx & "列：[" & valA & "] -> [" & valB & "]"
            End If
        End If
    Next colIdx

    If colsA <> colsB Then
        If Len(buf) > 0 Then buf = buf & "；"
        buf = buf & "列数不同（A=" & colsA & "，B=" & colsB & "），只比较了前 " & common & " 列"
    End If

    DiffRow = buf
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = "#ERR"
    ElseIf IsNull(v) Then
        CellText = ""
    Else
        CellText = Trim$(CStr(v))
    End If
End Function

'------------------------------------------------------------------------------
' 写一行报告，并给行号加上可跳转的超链接
'------------------------------------------------------------------------------
Private Sub WriteRow(ByVal outWs As Worksheet, ByVal outRow As Long, _
                     ByVal category As String, ByVal keyText As String, _
                     ByVal rngA As Range, ByVal rowA As Long, _
                     ByVal rngB As Range, ByVal rowB As Long, _
                     ByVal diffText As String)
    outWs.Cells(outRow, 1).Value = category
    outWs.Cells(outRow, 2).Value = keyText
    outWs.Cells(outRow, 5).Value = diffText

    If Not rngA Is Nothing Then AddJumpLink outWs.Cells(outRow, 3), rngA, rowA
    If Not rngB Is Nothing Then AddJumpLink outWs.Cells(outRow, 4), rngB, rowB
End Sub

Private Sub AddJumpLink(ByVal cellRng As Range, ByVal srcRng As Range, ByVal rowOffset As Long)
    Dim absRow As Long
    absRow = srcRng.Row + rowOffset - 1

    Dim addr As String
    addr = "'" & srcRng.Worksheet.Name & "'!" & srcRng.Worksheet.Cells(absRow, srcRng.Column).Address

    cellRng.Worksheet.Hyperlinks.Add Anchor:=cellRng, Address:="", _
                                     SubAddress:=addr, TextToDisplay:=CStr(absRow)
End Sub
