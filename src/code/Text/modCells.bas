Attribute VB_Name = "modCells"
'==============================================================================
' modCells - 合并单元格处理（M1）
'
' 合并单元格是数据处理的头号障碍：排序、筛选、透视、公式引用全都会出问题。
' 这两个工具是一对互逆操作——先拆开填充做完数据处理，再按需合并回去做展示。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 拆分合并单元格并向下填充。
'
' Excel 自带的"取消合并"只把值留在左上角，其余全变空，正是数据处理时最不想要的结果。
'------------------------------------------------------------------------------
Public Function UnmergeAndFill(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        UnmergeAndFill = "选区内没有数据。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    ' 先把每个合并区域的范围和值收集起来，拆完再回填。
    ' 不能边拆边填：拆开的瞬间 MergeArea 就失效了。
    Dim areaAddrs As Object, areaValues As Object
    Set areaAddrs = CreateObject("Scripting.Dictionary")
    Set areaValues = CreateObject("Scripting.Dictionary")

    Dim cellRng As Range, mergedRng As Range, keyText As String
    For Each cellRng In srcRng.Cells
        If cellRng.MergeCells Then
            Set mergedRng = cellRng.MergeArea
            keyText = mergedRng.Address
            If Not areaAddrs.Exists(keyText) Then
                areaAddrs(keyText) = keyText
                areaValues(keyText) = mergedRng.Cells(1, 1).Value
            End If
        End If
    Next cellRng

    If areaAddrs.Count = 0 Then
        UnmergeAndFill = "选区内没有合并单元格。"
        Exit Function
    End If

    ' 拆分改变单元格结构，整表快照
    modUndo.CaptureSheet ws

    Dim k As Variant
    For Each k In areaAddrs.Keys
        ws.Range(CStr(k)).UnMerge
        ws.Range(CStr(k)).Value = areaValues(k)
    Next k

    UnmergeAndFill = "已拆分 " & areaAddrs.Count & " 处合并单元格并填充。"
End Function

'------------------------------------------------------------------------------
' 把同一列中相邻且内容相同的单元格合并。
'
' 只合并【相邻】的，不跨越不同值——否则会把本来分属不同分组的行并到一起，
' 数据看着整齐了，实际含义已经错了。
'------------------------------------------------------------------------------
Public Function MergeSameValues(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        MergeSameValues = "选区内没有数据。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    modUndo.CaptureSheet ws

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False       ' 合并时"仅保留左上角值"的提示

    Dim colIdx As Long, rowIdx As Long, runStart As Long, mergedCount As Long
    Dim curVal As String, prevVal As String

    For colIdx = 1 To UBound(srcArr, 2)
        runStart = 1
        prevVal = CStr(srcArr(1, colIdx))

        For rowIdx = 2 To UBound(srcArr, 1) + 1
            If rowIdx > UBound(srcArr, 1) Then
                curVal = ChrW$(&H1) & "END"     ' 哨兵，保证最后一段也会被收尾
            Else
                curVal = CStr(srcArr(rowIdx, colIdx))
            End If

            If curVal <> prevVal Then
                If rowIdx - runStart > 1 And Len(prevVal) > 0 Then
                    srcRng.Cells(runStart, colIdx).Resize(rowIdx - runStart, 1).Merge
                    mergedCount = mergedCount + 1
                End If
                runStart = rowIdx
                prevVal = curVal
            End If
        Next rowIdx
    Next colIdx

    Application.DisplayAlerts = prevAlerts

    If mergedCount = 0 Then
        MergeSameValues = "没有找到可合并的连续相同值。"
    Else
        MergeSameValues = "已合并 " & mergedCount & " 处相同内容。"
    End If
End Function
