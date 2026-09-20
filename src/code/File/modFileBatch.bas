Attribute VB_Name = "modFileBatch"
'==============================================================================
' modFileBatch - 文件批处理（M5）
'
' 这个模块里的操作【全部不可撤销】：它们动的是磁盘上的文件，不是单元格。
' 所以每个命令都注册成 ConfirmBeforeRun，执行前强制确认。
'
' 批量重命名尤其危险，因此拆成两步：先生成预览表，人工核对无误后再执行。
' 一次改错几百个文件名是找不回来的。
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 把文件夹内的文件清单导入工作表，同时生成"新文件名"列供填写。
' 这张表既是清单，也是批量重命名的输入。
'------------------------------------------------------------------------------
Public Function ListFiles() As String
    Dim folderPath As String
    folderPath = modPrompt.AskFolder("folder", "选择要列出文件的文件夹")

    Dim recursive As Boolean
    recursive = modPrompt.AskYesNo("recursive", "包含子文件夹吗？")

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim paths As Collection
    Set paths = New Collection
    CollectAll fso, fso.GetFolder(folderPath), recursive, paths

    If paths.Count = 0 Then
        ListFiles = "文件夹内没有文件。"
        Exit Function
    End If

    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(ActiveWorkbook, "文件清单")
    outWs.Range("A1:F1").Value = Array("完整路径", "文件名", "扩展名", "大小(KB)", "修改时间", "新文件名")
    modSheetUtil.FormatHeader outWs, 6

    Dim outArr() As Variant
    ReDim outArr(1 To paths.Count, 1 To 5)

    Dim i As Long, fileObj As Object
    For i = 1 To paths.Count
        Set fileObj = fso.GetFile(paths(i))
        outArr(i, 1) = fileObj.Path
        outArr(i, 2) = fileObj.Name
        outArr(i, 3) = fso.GetExtensionName(fileObj.Path)
        outArr(i, 4) = Round(fileObj.Size / 1024, 1)
        outArr(i, 5) = fileObj.DateLastModified
    Next i

    outWs.Range("A2").Resize(paths.Count, 5).Value = outArr
    outWs.Columns.AutoFit

    ListFiles = "已列出 " & paths.Count & " 个文件到工作表「" & outWs.Name & "」。" & vbCrLf & vbCrLf & _
                "如需批量重命名：在 F 列填写新文件名（含扩展名），然后执行「批量重命名」。"
End Function

Private Sub CollectAll(ByVal fso As Object, ByVal folderObj As Object, _
                       ByVal recursive As Boolean, ByVal result As Collection)
    Dim fileObj As Object
    For Each fileObj In folderObj.Files
        If Left$(fso.GetFileName(fileObj.Path), 2) <> "~$" Then result.Add fileObj.Path
    Next fileObj

    If recursive Then
        Dim subFolder As Object
        For Each subFolder In folderObj.SubFolders
            CollectAll fso, subFolder, True, result
        Next subFolder
    End If
End Sub

'------------------------------------------------------------------------------
' 按「文件清单」表执行批量重命名。
'
' A 列 = 原完整路径，F 列 = 新文件名。空白的 F 列表示不改。
'
' 安全措施：
'   - 先全表校验（源文件在不在、目标名合不合法、会不会撞已有文件、批内有没有重名），
'     全部通过才动手。改到一半失败是最糟的结果。
'   - 执行后把实际结果写回 G 列，留一份可追溯的记录。
'------------------------------------------------------------------------------
Public Function BatchRenameFiles(ByVal ws As Worksheet) As String
    Dim used As Range
    Set used = modRange.RealUsedRange(ws)
    If used Is Nothing Or used.Rows.Count < 2 Then
        BatchRenameFiles = "当前工作表没有文件清单。请先执行「文件清单」。"
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim lastRow As Long
    lastRow = used.Row + used.Rows.Count - 1

    ' --- 第一遍：只校验，不改任何文件 ---
    Dim plannedTargets As Object
    Set plannedTargets = CreateObject("Scripting.Dictionary")
    plannedTargets.CompareMode = vbTextCompare

    Dim rowIdx As Long, srcPath As String, newName As String, targetPath As String
    Dim problems As Collection
    Set problems = New Collection
    Dim planCount As Long

    For rowIdx = 2 To lastRow
        srcPath = Trim$(CStr(ws.Cells(rowIdx, 1).Value))
        newName = Trim$(CStr(ws.Cells(rowIdx, 6).Value))

        If Len(srcPath) > 0 And Len(newName) > 0 Then
            If Not fso.FileExists(srcPath) Then
                problems.Add "第 " & rowIdx & " 行：源文件不存在 " & srcPath
            ElseIf HasInvalidNameChars(newName) Then
                problems.Add "第 " & rowIdx & " 行：文件名含非法字符 " & newName
            Else
                targetPath = modIO.JoinPath(fso.GetParentFolderName(srcPath), newName)

                If StrComp(targetPath, srcPath, vbTextCompare) <> 0 Then
                    If fso.FileExists(targetPath) Then
                        problems.Add "第 " & rowIdx & " 行：目标文件已存在 " & newName
                    ElseIf plannedTargets.Exists(targetPath) Then
                        problems.Add "第 " & rowIdx & " 行：批内重名 " & newName
                    Else
                        plannedTargets(targetPath) = rowIdx
                        planCount = planCount + 1
                    End If
                End If
            End If
        End If
    Next rowIdx

    If planCount = 0 Then
        BatchRenameFiles = "没有需要重命名的文件（F 列为空或与原名相同）。"
        Exit Function
    End If

    If problems.Count > 0 Then
        Dim detail As String, i As Long
        For i = 1 To problems.Count
            If i > 10 Then
                detail = detail & vbCrLf & "…… 共 " & problems.Count & " 处问题"
                Exit For
            End If
            detail = detail & vbCrLf & problems(i)
        Next i
        BatchRenameFiles = "校验未通过，未改动任何文件：" & detail
        Exit Function
    End If

    ' --- 第二遍：校验全过，开始实际改名 ---
    ws.Cells(1, 7).Value = "重命名结果"
    ws.Cells(1, 7).Font.Bold = True

    Dim renamed As Long, failed As Long
    Dim k As Variant
    For Each k In plannedTargets.Keys
        rowIdx = plannedTargets(k)
        srcPath = Trim$(CStr(ws.Cells(rowIdx, 1).Value))

        ' 校验和执行之间有时间差，期间别人可能已经建了同名文件。
        ' 这里紧挨着 MoveFile 再查一次，把这个窗口缩到最小——
        ' 彻底消除做不到（文件系统没给我们原子的"不存在才改名"），
        ' 但至少不能眼睁睁覆盖掉别人的文件。
        On Error Resume Next
        Err.Clear
        If fso.FileExists(CStr(k)) Then
            failed = failed + 1
            ws.Cells(rowIdx, 7).Value = "失败：目标文件在校验之后被创建，已跳过"
        Else
            fso.MoveFile srcPath, CStr(k)
            If Err.Number = 0 Then
                renamed = renamed + 1
                ws.Cells(rowIdx, 7).Value = "成功"
                ws.Cells(rowIdx, 1).Value = CStr(k)      ' 路径已变，同步更新清单
            Else
                failed = failed + 1
                ws.Cells(rowIdx, 7).Value = "失败：" & Err.Description
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next k

    ' 部分失败时必须说清楚现状：已改的不会自动改回去。
    ' 含糊其辞会让用户以为"失败了就是没动"，然后按旧文件名去找文件。
    Dim tail As String
    If failed > 0 Then
        tail = "，失败 " & failed & " 个（详见 G 列）。" & vbCrLf & vbCrLf & _
               "【已成功的 " & renamed & " 个不会自动改回去】，G 列就是改名记录，" & _
               "需要还原请照着它手工处理。"
    Else
        tail = "。"
    End If

    BatchRenameFiles = "重命名完成：成功 " & renamed & " 个" & tail & vbCrLf & vbCrLf & _
                       "注意：文件重命名无法撤销。"
End Function

Private Function HasInvalidNameChars(ByVal fileName As String) As Boolean
    Dim bad As String, i As Long
    bad = "\/:*?""<>|"
    For i = 1 To Len(bad)
        If InStr(fileName, Mid$(bad, i, 1)) > 0 Then
            HasInvalidNameChars = True
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' 把当前工作簿的每张工作表导出为独立文件
'------------------------------------------------------------------------------
Public Function ExportSheets(ByVal wb As Workbook) As String
    Dim folderPath As String
    folderPath = modPrompt.AskFolder("folder", "选择导出到哪个文件夹")

    Dim formatChoice As Long
    Dim formats(1 To 3) As String
    formats(1) = "Excel 工作簿 (.xlsx)"
    formats(2) = "CSV (.csv)"
    formats(3) = "PDF (.pdf)"
    formatChoice = modPrompt.AskChoice("exportFormat", "导出成什么格式？", formats)

    Dim ws As Worksheet, exported As Long, skipped As Long
    Dim targetPath As String, baseName As String
    Dim tempWb As Workbook

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            baseName = modSheetUtil.SafeSheetName(ws.Name)
            modPerf.SetStatus "导出 " & baseName

            Select Case formatChoice
                Case 1
                    targetPath = modIO.JoinPath(folderPath, baseName & ".xlsx")
                    ws.Copy                              ' 复制成新工作簿
                    Set tempWb = ActiveWorkbook
                    tempWb.SaveAs targetPath, xlOpenXMLWorkbook
                    tempWb.Close SaveChanges:=False

                Case 2
                    targetPath = modIO.JoinPath(folderPath, baseName & ".csv")
                    ws.Copy
                    Set tempWb = ActiveWorkbook
                    tempWb.SaveAs targetPath, xlCSVUTF8      ' UTF-8，避免中文乱码
                    tempWb.Close SaveChanges:=False

                Case 3
                    targetPath = modIO.JoinPath(folderPath, baseName & ".pdf")
                    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=targetPath
            End Select

            exported = exported + 1
        Else
            skipped = skipped + 1
        End If
    Next ws

    modPerf.ClearStatus
    ExportSheets = "已导出 " & exported & " 张工作表到：" & vbCrLf & folderPath & _
                   IIf(skipped > 0, vbCrLf & vbCrLf & "跳过 " & skipped & " 张隐藏工作表。", "")
End Function

'------------------------------------------------------------------------------
' 按单元格内容批量插入图片。
'
' 选区里每个单元格的值当作图片文件名（不含扩展名），在指定文件夹里找同名图片，
' 插入到该单元格右侧一格，并按行高自动缩放。
'------------------------------------------------------------------------------
Public Function InsertImages(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        InsertImages = "请先选中包含图片名称的单元格。"
        Exit Function
    End If

    Dim folderPath As String
    folderPath = modPrompt.AskFolder("folder", "选择图片所在文件夹")

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim exts As Variant
    exts = Array("jpg", "jpeg", "png", "gif", "bmp")

    Dim cellRng As Range, nameText As String, found As String
    Dim inserted As Long, missing As Long
    Dim targetCell As Range, pic As Object
    Dim i As Long, candidate As String

    For Each cellRng In srcRng.Cells
        nameText = Trim$(CStr(cellRng.Value))
        If Len(nameText) > 0 Then
            found = ""
            For i = LBound(exts) To UBound(exts)
                candidate = modIO.JoinPath(folderPath, nameText & "." & exts(i))
                If fso.FileExists(candidate) Then
                    found = candidate
                    Exit For
                End If
            Next i

            If Len(found) = 0 Then
                missing = missing + 1
            Else
                Set targetCell = ws.Cells(cellRng.Row, cellRng.Column + 1)
                Set pic = ws.Shapes.AddPicture(Filename:=found, _
                            LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                            Left:=targetCell.Left + 1, Top:=targetCell.Top + 1, _
                            Width:=-1, Height:=-1)

                ' 等比缩放到刚好放进单元格
                pic.LockAspectRatio = msoTrue
                If pic.Height > targetCell.Height - 2 Then
                    pic.Height = targetCell.Height - 2
                End If
                If pic.Width > targetCell.Width - 2 Then
                    pic.Width = targetCell.Width - 2
                End If

                inserted = inserted + 1
            End If
        End If
    Next cellRng

    InsertImages = "已插入 " & inserted & " 张图片" & _
                   IIf(missing > 0, "，" & missing & " 个名称没有找到对应图片", "") & "。"
End Function
