Attribute VB_Name = "modSheets"
'==============================================================================
' modSheets - 工作表与工作簿管理（M3）
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 按某列的值把数据拆分成多个工作表。
'
' 表名直接用列值是最常见的翻车点：列值可能超过 31 字符、含 : \ / ? * [ ] 、
' 或者重名。统一交给 modSheetUtil.UniqueSheetName 处理。
'------------------------------------------------------------------------------
Public Function SplitByColumn(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        SplitByColumn = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Rows.Count < 2 Then
        SplitByColumn = "至少需要标题行加一行数据。"
        Exit Function
    End If

    Dim keyCol As Long
    keyCol = modPrompt.AskNumber("keyColumn", "按第几列的值拆分？（选区内的列序号）", 1)
    If keyCol < 1 Or keyCol > srcRng.Columns.Count Then
        Err.Raise vbObjectError + 390, "modSheets", "列序号超出选区范围。"
    End If

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim totalCols As Long
    totalCols = UBound(srcArr, 2)

    ' 先按键分组收集行号，再一次性成表——边扫边建表会反复触发重算
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    Dim rowIdx As Long, keyText As String
    For rowIdx = 2 To UBound(srcArr, 1)
        keyText = Trim$(CStr(NzVal(srcArr(rowIdx, keyCol))))
        If Len(keyText) = 0 Then keyText = "(空白)"
        If Not groups.Exists(keyText) Then
            Set groups(keyText) = New Collection
        End If
        groups(keyText).Add rowIdx
    Next rowIdx

    If groups.Count = 0 Then
        SplitByColumn = "没有可拆分的数据行。"
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = srcRng.Worksheet.Parent

    Dim k As Variant, rowList As Collection, outWs As Worksheet
    Dim outArr() As Variant, outIdx As Long, colIdx As Long, i As Long
    Dim created As Long

    For Each k In groups.Keys
        Set rowList = groups(k)

        ReDim outArr(1 To rowList.Count + 1, 1 To totalCols)
        For colIdx = 1 To totalCols
            outArr(1, colIdx) = srcArr(1, colIdx)        ' 标题行
        Next colIdx

        outIdx = 2
        For i = 1 To rowList.Count
            For colIdx = 1 To totalCols
                outArr(outIdx, colIdx) = srcArr(rowList(i), colIdx)
            Next colIdx
            outIdx = outIdx + 1
        Next i

        Set outWs = modSheetUtil.AddSheet(wb, CStr(k))
        outWs.Range("A1").Resize(rowList.Count + 1, totalCols).Value = outArr
        modSheetUtil.FormatHeader outWs, totalCols
        created = created + 1
    Next k

    SplitByColumn = "已按第 " & keyCol & " 列拆分为 " & created & " 个工作表。"
End Function

Private Function NzVal(ByVal v As Variant) As Variant
    If IsError(v) Then
        NzVal = "#ERR"
    ElseIf IsNull(v) Then
        NzVal = ""
    Else
        NzVal = v
    End If
End Function

'------------------------------------------------------------------------------
' 合并当前工作簿内所有工作表到一张汇总表。
'
' 以第一张表的标题行为准，其余表按标题名对齐——直接按列位置拼接的话，
' 只要有一张表少了一列，后面所有数据就整体错位，而且看不出来。
'------------------------------------------------------------------------------
Public Function MergeAllSheets(ByVal wb As Workbook) As String
    If wb.Worksheets.Count < 2 Then
        MergeAllSheets = "当前工作簿只有一张工作表，无需合并。"
        Exit Function
    End If

    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(wb, "汇总")

    ' 标题名 -> 汇总表列号
    Dim headerMap As Object
    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    headerMap("来源工作表") = 1

    Dim outRow As Long
    outRow = 2

    Dim ws As Worksheet, used As Range, srcArr As Variant
    Dim rowIdx As Long, colIdx As Long, headerText As String, targetCol As Long
    Dim merged As Long, skipped As Long

    For Each ws In wb.Worksheets
        If ws.Name <> outWs.Name Then
            Set used = modRange.RealUsedRange(ws)
            If used Is Nothing Then
                skipped = skipped + 1
            ElseIf used.Rows.Count < 2 Then
                skipped = skipped + 1
            Else
                srcArr = modRange.ToArray(used)

                For rowIdx = 2 To UBound(srcArr, 1)
                    outWs.Cells(outRow, 1).Value = ws.Name
                    For colIdx = 1 To UBound(srcArr, 2)
                        headerText = Trim$(CStr(NzVal(srcArr(1, colIdx))))
                        If Len(headerText) = 0 Then headerText = "列" & colIdx

                        If headerMap.Exists(headerText) Then
                            targetCol = headerMap(headerText)
                        Else
                            targetCol = headerMap.Count + 1
                            headerMap(headerText) = targetCol
                        End If

                        outWs.Cells(outRow, targetCol).Value = srcArr(rowIdx, colIdx)
                    Next colIdx
                    outRow = outRow + 1
                Next rowIdx

                merged = merged + 1
            End If
        End If
    Next ws

    ' 最后再写标题行——列的集合要全部扫完才知道
    Dim k As Variant
    For Each k In headerMap.Keys
        outWs.Cells(1, headerMap(k)).Value = k
    Next k
    modSheetUtil.FormatHeader outWs, headerMap.Count

    MergeAllSheets = "已合并 " & merged & " 张工作表，共 " & (outRow - 2) & " 行数据。" & _
                     IIf(skipped > 0, vbCrLf & "跳过 " & skipped & " 张空表。", "") & vbCrLf & _
                     "结果在工作表「" & outWs.Name & "」。"
End Function

'------------------------------------------------------------------------------
' 生成带超链接的目录导航表
'------------------------------------------------------------------------------
Public Function CreateIndex(ByVal wb As Workbook) As String
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = wb.Worksheets("目录")
    On Error GoTo 0

    Dim outWs As Worksheet
    If existing Is Nothing Then
        Set outWs = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        outWs.Name = modSheetUtil.UniqueSheetName(wb, "目录")
    Else
        Set outWs = existing
        modUndo.CaptureSheet outWs
        outWs.Cells.Clear
    End If

    outWs.Range("A1:C1").Value = Array("序号", "工作表", "已用区域")
    modSheetUtil.FormatHeader outWs, 3

    Dim ws As Worksheet, outRow As Long, idx As Long, used As Range
    outRow = 2
    For Each ws In wb.Worksheets
        If ws.Name <> outWs.Name Then
            idx = idx + 1
            outWs.Cells(outRow, 1).Value = idx
            outWs.Hyperlinks.Add Anchor:=outWs.Cells(outRow, 2), Address:="", _
                                 SubAddress:="'" & ws.Name & "'!A1", _
                                 TextToDisplay:=ws.Name

            Set used = modRange.RealUsedRange(ws)
            If used Is Nothing Then
                outWs.Cells(outRow, 3).Value = "(空表)"
            Else
                outWs.Cells(outRow, 3).Value = used.Address(False, False)
            End If
            outRow = outRow + 1
        End If
    Next ws

    outWs.Columns.AutoFit
    CreateIndex = "已生成目录，共 " & idx & " 张工作表。"
End Function

'------------------------------------------------------------------------------
' 按名称排序工作表。
'
' 用 Move 做插入排序，不用先取名字再排数组——Move 之后索引会变，
' 按旧索引操作必然错位。
'------------------------------------------------------------------------------
Public Function SortSheets(ByVal wb As Workbook) As String
    Dim ascending As Boolean
    ascending = modPrompt.AskYesNo("ascending", "按升序排列吗？（否 = 降序）")

    Dim i As Long, j As Long, shouldSwap As Boolean
    For i = 1 To wb.Worksheets.Count - 1
        For j = i + 1 To wb.Worksheets.Count
            If ascending Then
                shouldSwap = (StrComp(wb.Worksheets(j).Name, wb.Worksheets(i).Name, vbTextCompare) < 0)
            Else
                shouldSwap = (StrComp(wb.Worksheets(j).Name, wb.Worksheets(i).Name, vbTextCompare) > 0)
            End If
            If shouldSwap Then wb.Worksheets(j).Move Before:=wb.Worksheets(i)
        Next j
    Next i

    SortSheets = "已按名称" & IIf(ascending, "升序", "降序") & "排列 " & wb.Worksheets.Count & " 张工作表。"
End Function

'------------------------------------------------------------------------------
' 显示全部隐藏的工作表（含"深度隐藏"的）
'------------------------------------------------------------------------------
Public Function ShowAllSheets(ByVal wb As Workbook) As String
    Dim ws As Worksheet, shown As Long
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            shown = shown + 1
        End If
    Next ws

    If shown = 0 Then
        ShowAllSheets = "没有被隐藏的工作表。"
    Else
        ShowAllSheets = "已显示 " & shown & " 张隐藏的工作表。"
    End If
End Function

'------------------------------------------------------------------------------
' 按选区里的名称列表批量重命名工作表。
' 选区第 N 个单元格 -> 第 N 张工作表。
'------------------------------------------------------------------------------
Public Function BatchRename(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        BatchRename = "请先选中一列新表名。"
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = srcRng.Worksheet.Parent

    Dim names As Collection
    Set names = New Collection

    Dim cellRng As Range, nameText As String
    For Each cellRng In srcRng.Cells
        nameText = Trim$(CStr(cellRng.Value))
        If Len(nameText) > 0 Then names.Add nameText
    Next cellRng

    If names.Count = 0 Then
        BatchRename = "选区内没有有效的名称。"
        Exit Function
    End If
    If names.Count > wb.Worksheets.Count Then
        BatchRename = "名称有 " & names.Count & " 个，但工作簿只有 " & _
                      wb.Worksheets.Count & " 张工作表。"
        Exit Function
    End If

    ' 先全部改成临时名，再改成目标名。
    ' 否则一旦新名和某张还没改到的表当前重名，就会中途报错，留下改了一半的烂摊子。
    Dim i As Long
    For i = 1 To names.Count
        wb.Worksheets(i).Name = "~tmp_" & i & "_" & Int(Rnd() * 100000)
    Next i

    Dim renamed As Long
    For i = 1 To names.Count
        wb.Worksheets(i).Name = modSheetUtil.UniqueSheetName(wb, names(i))
        renamed = renamed + 1
    Next i

    BatchRename = "已重命名 " & renamed & " 张工作表。"
End Function
