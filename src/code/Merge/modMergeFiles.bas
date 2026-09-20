Attribute VB_Name = "modMergeFiles"
'==============================================================================
' modMergeFiles - 多文件合并（M4）
'
' 设计上的三个决定：
'   1. 源文件一律【只读】打开，合并任务没有任何理由去写别人的文件；
'   2. 单个文件失败不中断整批，最后统一出一份失败清单——批量任务里
'      一个坏文件毁掉其余 99 个是不可接受的；
'   3. 结果带"来源文件 / 来源工作表"两列。合并完之后如果查不出某行来自哪，
'      数据对不上时就完全没法追溯。
'==============================================================================
Option Explicit
Option Private Module

Public Function MergeFolder() As String
    Dim folderPath As String
    folderPath = modPrompt.AskFolder("folder", "选择要合并的文件夹")

    Dim recursive As Boolean
    recursive = modPrompt.AskYesNo("recursive", "包含子文件夹吗？")

    Dim sheetMode As Long
    Dim modes(1 To 2) As String
    modes(1) = "每个文件的第一张工作表"
    modes(2) = "每个文件的所有工作表"
    sheetMode = modPrompt.AskChoice("sheetMode", "要合并哪些工作表？", modes)

    Dim headerRows As Long
    headerRows = modPrompt.AskNumber("headerRows", _
        "每个文件的标题行占几行？（0 表示没有标题行）", 1)
    If headerRows < 0 Then headerRows = 0

    Dim files As Collection
    Set files = modIO.ListExcelFiles(folderPath, recursive)

    If files.Count = 0 Then
        MergeFolder = "文件夹内没有找到 Excel 文件。"
        Exit Function
    End If

    ' 结果写到一个【新工作簿】，不污染用户当前文件
    Dim outWb As Workbook
    Set outWb = Application.Workbooks.Add
    Dim outWs As Worksheet
    Set outWs = outWb.Worksheets(1)
    outWs.Name = "合并结果"

    Dim outRow As Long, headerWritten As Boolean, dataCols As Long
    outRow = 1

    Dim failures As Collection
    Set failures = New Collection

    Dim okFiles As Long, okSheets As Long
    Dim i As Long, filePath As String
    Dim srcWb As Workbook, ws As Worksheet

    For i = 1 To files.Count
        filePath = files(i)
        modPerf.SetStatus "合并中 " & i & "/" & files.Count & "：" & modIO.FileNameOf(filePath)

        Set srcWb = modIO.OpenQuiet(filePath)
        If srcWb Is Nothing Then
            failures.Add modIO.FileNameOf(filePath) & "：无法打开"
        Else
            Dim fileHadData As Boolean
            fileHadData = False

            On Error GoTo FileFailed
            For Each ws In srcWb.Worksheets
                If sheetMode = 2 Or ws.Index = 1 Then
                    If AppendSheet(ws, outWs, outRow, headerWritten, dataCols, _
                                   headerRows, filePath) Then
                        okSheets = okSheets + 1
                        fileHadData = True
                    End If
                End If
                If sheetMode = 1 Then Exit For
            Next ws
            On Error GoTo 0

            If fileHadData Then okFiles = okFiles + 1
            modIO.CloseQuiet srcWb
            Set srcWb = Nothing
        End If
    Next i

    GoTo Summarize

FileFailed:
    failures.Add modIO.FileNameOf(filePath) & "：" & Err.Description
    modIO.CloseQuiet srcWb
    Set srcWb = Nothing
    Resume Next

Summarize:
    modPerf.ClearStatus

    If outRow = 1 Then
        outWb.Close SaveChanges:=False
        MergeFolder = "没有合并到任何数据。"
        Exit Function
    End If

    modSheetUtil.FormatHeader outWs, dataCols + 2

    Dim summary As String
    summary = "已合并 " & okFiles & " 个文件、" & okSheets & " 张工作表，共 " & _
              (outRow - 2) & " 行数据。"

    If failures.Count > 0 Then
        Dim failWs As Worksheet
        Set failWs = modSheetUtil.AddSheet(outWb, "失败清单")
        failWs.Range("A1").Value = "失败的文件"
        failWs.Range("A1").Font.Bold = True
        For i = 1 To failures.Count
            failWs.Cells(i + 1, 1).Value = failures(i)
        Next i
        failWs.Columns.AutoFit
        summary = summary & vbCrLf & vbCrLf & _
                  failures.Count & " 个文件失败，详见「" & failWs.Name & "」工作表。"
    End If

    MergeFolder = summary
End Function

'------------------------------------------------------------------------------
' 把一张源表追加到结果表。返回是否真的写入了数据。
'
' 前两列固定是来源文件和来源工作表，业务列从第 3 列开始。
'------------------------------------------------------------------------------
Private Function AppendSheet(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, _
                             ByRef outRow As Long, ByRef headerWritten As Boolean, _
                             ByRef dataCols As Long, ByVal headerRows As Long, _
                             ByVal filePath As String) As Boolean
    Dim used As Range
    Set used = modRange.RealUsedRange(srcWs)
    If used Is Nothing Then Exit Function
    If used.Rows.Count <= headerRows Then Exit Function

    Dim srcArr As Variant
    srcArr = modRange.ToArray(used)

    Dim colCount As Long
    colCount = UBound(srcArr, 2)

    ' 标题行取第一个有数据的文件的，后续文件的标题行直接跳过
    If Not headerWritten Then
        outWs.Cells(1, 1).Value = "来源文件"
        outWs.Cells(1, 2).Value = "来源工作表"
        If headerRows > 0 Then
            Dim colIdx As Long
            For colIdx = 1 To colCount
                outWs.Cells(1, colIdx + 2).Value = srcArr(1, colIdx)
            Next colIdx
        End If
        dataCols = colCount
        headerWritten = True
        outRow = 2
    End If

    ' 后续文件列数更多时，多出来的列照样写入，不静默丢数据
    If colCount > dataCols Then dataCols = colCount

    Dim dataRows As Long
    dataRows = UBound(srcArr, 1) - headerRows
    If dataRows <= 0 Then Exit Function

    Dim outArr() As Variant
    ReDim outArr(1 To dataRows, 1 To colCount + 2)

    Dim rowIdx As Long, c As Long
    Dim fileName As String
    fileName = modIO.FileNameOf(filePath)

    For rowIdx = 1 To dataRows
        outArr(rowIdx, 1) = fileName
        outArr(rowIdx, 2) = srcWs.Name
        For c = 1 To colCount
            outArr(rowIdx, c + 2) = srcArr(rowIdx + headerRows, c)
        Next c
    Next rowIdx

    outWs.Cells(outRow, 1).Resize(dataRows, colCount + 2).Value = outArr
    outRow = outRow + dataRows

    AppendSheet = True
End Function
