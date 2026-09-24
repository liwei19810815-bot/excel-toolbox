Attribute VB_Name = "modRows"
'==============================================================================
' modRows - 行级批量操作（M2 数据处理）
'
' 本模块也是撤销框架的首个验证用例：删除整行是典型的【结构性操作】，
' 执行后下方所有行整体上移，只快照"将被删的那几行"是没用的——
' 还原时地址早就对不上了。所以必须用 CaptureSheet 快照整张表。
'
' 约定（全工具箱通用）：
'   - 不碰 ScreenUpdating / Calculation，那是 modPerf 的事；
'   - 不写 On Error 弹窗，异常直接抛给 RunAction 统一处理并回滚；
'   - 不弹任何对话框，结果以字符串返回，由 RunAction 统一呈现；
'   - 改动之前先 Capture。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 删除选区内完全为空的整行。
'
' "空"的判定只看用户选中的那几列：选了 A:C 就只看 A:C，
' 哪怕 D 列有值该行也算空行——这符合"我圈定的范围内是空的"这一直觉。
' 选整行/整列/整表时会先收敛到真实已用区域，避免遍历一百万行。
'------------------------------------------------------------------------------
Public Function DeleteEmptyRows(ByVal target As Range) As String
    Dim rng As Range
    Set rng = modRange.NormalizeSelection(target)
    If rng Is Nothing Then
        DeleteEmptyRows = "选区内没有数据。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = rng.Worksheet

    ' 合并单元格会让"整行删除"产生难以预期的结果，先拦下来
    If modRange.HasMergedCells(rng) Then
        DeleteEmptyRows = "选区内存在合并单元格，删除整行的结果不可预期。" & vbCrLf & _
                          "请先取消合并，或缩小选区后重试。"
        Exit Function
    End If

    Dim box As Range
    Set box = ws.Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, rng.Columns.Count))

    ' 一次性读入内存判空，比逐单元格 CountA 快一到两个数量级
    Dim arr As Variant
    arr = modRange.ToArray(box)

    Dim firstRow As Long
    firstRow = box.Row

    Dim acc As clsAreaAccumulator
    Set acc = New clsAreaAccumulator
    acc.Init ws

    Dim r As Long, c As Long, isBlank As Boolean
    For r = LBound(arr, 1) To UBound(arr, 1)
        isBlank = True
        For c = LBound(arr, 2) To UBound(arr, 2)
            If Len(CStr(arr(r, c))) > 0 Then
                isBlank = False
                Exit For
            End If
        Next c
        If isBlank Then acc.AddRow firstRow + r - 1
    Next r

    If acc.HasNoAreas Then
        DeleteEmptyRows = "选区内没有空行。"
        Exit Function
    End If

    Dim n As Long
    n = acc.Count

    ' 结构性操作：必须整表快照，且必须在删除之前
    modUndo.CaptureSheet ws

    ' 一次性删除多区域，比循环 Delete 快得多，也不用倒着遍历处理行号位移
    acc.Result.EntireRow.Delete

    DeleteEmptyRows = "已删除 " & n & " 个空行。"
End Function

'------------------------------------------------------------------------------
' 删除选区内完全为空的整列。与删除空行同理，判空只看用户选中的那些行。
'------------------------------------------------------------------------------
Public Function DeleteEmptyColumns(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        DeleteEmptyColumns = "选区内没有数据。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    If modRange.HasMergedCells(srcRng) Then
        DeleteEmptyColumns = "选区内存在合并单元格，删除整列的结果不可预期。请先取消合并。"
        Exit Function
    End If

    Dim box As Range
    Set box = ws.Range(srcRng.Cells(1, 1), srcRng.Cells(srcRng.Rows.Count, srcRng.Columns.Count))

    Dim srcArr As Variant
    srcArr = modRange.ToArray(box)

    Dim firstCol As Long
    firstCol = box.Column

    Dim acc As clsAreaAccumulator
    Set acc = New clsAreaAccumulator
    acc.Init ws

    Dim rowIdx As Long, colIdx As Long, isBlank As Boolean
    For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
        isBlank = True
        For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
            If Len(CStr(srcArr(rowIdx, colIdx))) > 0 Then
                isBlank = False
                Exit For
            End If
        Next rowIdx
        If isBlank Then acc.AddColumn firstCol + colIdx - 1
    Next colIdx

    If acc.HasNoAreas Then
        DeleteEmptyColumns = "选区内没有空列。"
        Exit Function
    End If

    Dim n As Long
    n = acc.Count

    modUndo.CaptureSheet ws
    acc.Result.EntireColumn.Delete

    DeleteEmptyColumns = "已删除 " & n & " 个空列。"
End Function
