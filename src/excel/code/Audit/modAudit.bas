Attribute VB_Name = "modAudit"
'==============================================================================
' modAudit - 数据体检（M7）
'
' 一次扫描出一整张问题清单，而不是让人一个个工具试过去。
' 报告里每条问题都带可点击的定位链接——只报个数字"有 37 个文本型数字"
' 而不能跳过去看，等于没说。
'
' 扫描是【只读】的，不改任何数据，所以不需要撤销。
'==============================================================================
Option Explicit
Option Private Module

' 单类问题最多列出多少条明细。全列出来的话，一张脏表能生成几十万行报告，
' 报告本身反而成了新的性能问题。
Private Const MAX_DETAIL_PER_TYPE As Long = 200

Public Function ScanSheet(ByVal ws As Worksheet) As String
    ' phase 只是个局部字符串，出错时把它塞进 Err.Source 一起抛出。
    ' 这样无头运行也能知道是哪一步炸的，而不是只看到一句
    ' "无效的过程调用或参数"。比模块级的全局打点安全，也不留状态。
    Dim phase As String
    On Error GoTo Failed

    phase = "定位已用区域"
    Dim used As Range
    Set used = modRange.RealUsedRange(ws)
    If used Is Nothing Then
        ScanSheet = "当前工作表是空的。"
        Exit Function
    End If

    phase = "读入数组"
    Dim srcArr As Variant
    srcArr = modRange.ToArray(used)

    phase = "新建报告表"
    Dim outWs As Worksheet
    Set outWs = modSheetUtil.AddSheet(ws.Parent, "体检报告")

    phase = "写报告表头"
    outWs.Range("A1:D1").Value = Array("问题类型", "位置", "内容", "说明")
    modSheetUtil.FormatHeader outWs, 4

    Dim outRow As Long
    outRow = 2

    ' 各类问题的计数
    Dim cntBlankRow As Long, cntTextNum As Long, cntSpace As Long
    Dim cntError As Long, cntMerged As Long, cntBadDate As Long, cntLong As Long

    Dim detailCount As Object
    Set detailCount = CreateObject("Scripting.Dictionary")

    Dim rowIdx As Long, colIdx As Long
    Dim cellVal As Variant, txt As String
    Dim addr As String

    phase = "逐格扫描"
    For rowIdx = 1 To UBound(srcArr, 1)
        For colIdx = 1 To UBound(srcArr, 2)
            phase = "取单元格地址 r" & rowIdx & "c" & colIdx
            cellVal = srcArr(rowIdx, colIdx)
            addr = used.Cells(rowIdx, colIdx).Address(False, False)

            If IsError(cellVal) Then
                cntError = cntError + 1
                AddDetail outWs, outRow, detailCount, "错误值", ws, addr, "#ERR", "公式返回了错误值"

            ElseIf VarType(cellVal) = vbString Then
                txt = CStr(cellVal)

                If Len(txt) > 0 Then
                    ' 文本型数字：SUM 结果为 0 却看不出原因的头号元凶
                    phase = "判定文本型数字 " & addr & " [" & txt & "]"
                    If IsTextNumber(txt) Then
                        cntTextNum = cntTextNum + 1
                        AddDetail outWs, outRow, detailCount, "文本型数字", ws, addr, txt, _
                                  "看着是数字，实际是文本，不参与求和"
                    End If

                    ' 首尾空格与不可见字符：导致"看着一样却匹配不上"
                    phase = "判定空白字符 " & addr
                    If HasStrayWhitespace(txt) Then
                        cntSpace = cntSpace + 1
                        AddDetail outWs, outRow, detailCount, "首尾空格/不可见字符", ws, addr, _
                                  "[" & txt & "]", "含首尾空格、不间断空格或零宽字符"
                    End If

                    ' 疑似日期却存成了文本
                    phase = "判定文本型日期 " & addr
                    If LooksLikeDate(txt) Then
                        cntBadDate = cntBadDate + 1
                        AddDetail outWs, outRow, detailCount, "文本型日期", ws, addr, txt, _
                                  "看着是日期，实际是文本，无法参与日期计算"
                    End If

                    If Len(txt) > 255 Then
                        cntLong = cntLong + 1
                        AddDetail outWs, outRow, detailCount, "超长文本", ws, addr, _
                                  Left$(txt, 50) & "…", "长度 " & Len(txt) & " 字符"
                    End If
                End If
            End If
        Next colIdx
    Next rowIdx

    Dim isBlank As Boolean
    For rowIdx = 1 To UBound(srcArr, 1)
        isBlank = True
        For colIdx = 1 To UBound(srcArr, 2)
            If Len(CStr(NzAudit(srcArr(rowIdx, colIdx)))) > 0 Then
                isBlank = False
                Exit For
            End If
        Next colIdx
        If isBlank Then
            cntBlankRow = cntBlankRow + 1
            AddDetail outWs, outRow, detailCount, "空行", ws, _
                      used.Rows(rowIdx).Address(False, False), "", "整行为空"
        End If
    Next rowIdx

    Dim mergedAreas As Object
    Set mergedAreas = CreateObject("Scripting.Dictionary")
    Dim cellRng As Range
    For Each cellRng In used.Cells
        If cellRng.MergeCells Then
            If Not mergedAreas.Exists(cellRng.MergeArea.Address) Then
                mergedAreas(cellRng.MergeArea.Address) = 1
                cntMerged = cntMerged + 1
                AddDetail outWs, outRow, detailCount, "合并单元格", ws, _
                          cellRng.MergeArea.Address(False, False), "", _
                          "合并单元格会破坏排序、筛选和透视"
            End If
        End If
    Next cellRng

    phase = "汇总"
    Dim summaryLines As String
    summaryLines = BuildSummary(cntBlankRow, cntTextNum, cntSpace, cntError, _
                                cntMerged, cntBadDate, cntLong)

    phase = "排版"
    outWs.Columns.AutoFit
    If outWs.Columns(3).ColumnWidth > 60 Then outWs.Columns(3).ColumnWidth = 60

    ScanSheet = "体检完成，明细见工作表「" & outWs.Name & "」。" & vbCrLf & vbCrLf & summaryLines
    Exit Function

Failed:
    Err.Raise Err.Number, "modAudit.ScanSheet[" & phase & "]", Err.Description
End Function

Private Function BuildSummary(ByVal blankRow As Long, ByVal textNum As Long, _
                              ByVal spaceIssue As Long, ByVal errVal As Long, _
                              ByVal merged As Long, ByVal badDate As Long, _
                              ByVal longText As Long) As String
    Dim buf As String
    buf = AddLine(buf, "空行", blankRow)
    buf = AddLine(buf, "文本型数字", textNum)
    buf = AddLine(buf, "首尾空格/不可见字符", spaceIssue)
    buf = AddLine(buf, "文本型日期", badDate)
    buf = AddLine(buf, "错误值", errVal)
    buf = AddLine(buf, "合并单元格", merged)
    buf = AddLine(buf, "超长文本", longText)

    If Len(buf) = 0 Then
        BuildSummary = "没有发现问题，数据很干净。"
    Else
        BuildSummary = buf
    End If
End Function

Private Function AddLine(ByVal buf As String, ByVal label As String, ByVal n As Long) As String
    If n = 0 Then
        AddLine = buf
    ElseIf Len(buf) = 0 Then
        AddLine = label & "：" & n
    Else
        AddLine = buf & vbCrLf & label & "：" & n
    End If
End Function

Private Sub AddDetail(ByVal outWs As Worksheet, ByRef outRow As Long, _
                      ByVal detailCount As Object, ByVal issueType As String, _
                      ByVal srcWs As Worksheet, ByVal addr As String, _
                      ByVal content As String, ByVal note As String)
    If Not detailCount.Exists(issueType) Then detailCount(issueType) = 0
    detailCount(issueType) = detailCount(issueType) + 1

    If detailCount(issueType) > MAX_DETAIL_PER_TYPE Then
        If detailCount(issueType) = MAX_DETAIL_PER_TYPE + 1 Then
            outWs.Cells(outRow, 1).Value = issueType
            outWs.Cells(outRow, 4).Value = "（同类问题过多，明细只列前 " & MAX_DETAIL_PER_TYPE & " 条）"
            outRow = outRow + 1
        End If
        Exit Sub
    End If

    outWs.Cells(outRow, 1).Value = issueType
    outWs.Hyperlinks.Add Anchor:=outWs.Cells(outRow, 2), Address:="", _
                         SubAddress:="'" & srcWs.Name & "'!" & addr, _
                         TextToDisplay:=addr
    outWs.Cells(outRow, 3).Value = content
    outWs.Cells(outRow, 4).Value = note
    outRow = outRow + 1
End Sub

Private Function NzAudit(ByVal v As Variant) As Variant
    If IsError(v) Then
        NzAudit = "#ERR"
    ElseIf IsNull(v) Then
        NzAudit = ""
    Else
        NzAudit = v
    End If
End Function

'------------------------------------------------------------------------------
' 判断"看着是数字的文本"。
' 要先把全角数字和千分位逗号归一化，否则最常见的那几种根本检不出来。
'------------------------------------------------------------------------------
Private Function IsTextNumber(ByVal txt As String) As Boolean
    Dim cleaned As String
    cleaned = modStr.NormalizeNumericText(txt)

    If Len(cleaned) = 0 Then Exit Function
    ' 日期也会被 IsNumeric 认作数字，这里排掉，交给文本型日期那条规则
    If InStr(cleaned, "/") > 0 Or InStr(cleaned, "-") > 1 Then Exit Function

    IsTextNumber = IsNumeric(cleaned)
End Function

Private Function HasStrayWhitespace(ByVal txt As String) As Boolean
    If txt <> Trim$(txt) Then
        HasStrayWhitespace = True
    ElseIf InStr(txt, ChrW$(&HA0)) > 0 Then
        HasStrayWhitespace = True
    ElseIf InStr(txt, ChrW$(&H200B)) > 0 Then
        HasStrayWhitespace = True
    ElseIf InStr(txt, ChrW$(&HFEFF)) > 0 Then
        HasStrayWhitespace = True
    ElseIf InStr(txt, ChrW$(&H3000)) > 0 Then
        HasStrayWhitespace = True
    End If
End Function

Private Function LooksLikeDate(ByVal txt As String) As Boolean
    Dim cleaned As String
    cleaned = Trim$(modStr.ToHalfWidth(txt))
    If Len(cleaned) < 6 Or Len(cleaned) > 25 Then Exit Function

    ' 必须同时含分隔符和数字，避免把普通短语误判成日期
    If InStr(cleaned, "/") = 0 And InStr(cleaned, "-") = 0 And InStr(cleaned, "年") = 0 Then
        Exit Function
    End If

    LooksLikeDate = IsDate(cleaned)
End Function

'------------------------------------------------------------------------------
' 一键清洗：把体检里最常见、且修复方式无歧义的几类问题一次性修掉。
'
' 只做三件确定安全的事：清理空白字符、文本型数字转数值、删除空行。
' 合并单元格和错误值不碰——它们的正确处理方式依赖业务语义，
' 自动改动很可能改错，那比不改更糟。
'------------------------------------------------------------------------------
Public Function QuickClean(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        QuickClean = "选区内没有数据。"
        Exit Function
    End If

    Dim parts As String
    parts = modText.CleanSpaces(srcRng) & vbCrLf & _
            modText.TextToNumber(srcRng) & vbCrLf & _
            modRows.DeleteEmptyRows(srcRng)

    QuickClean = "一键清洗完成：" & vbCrLf & parts & vbCrLf & vbCrLf & _
                 "合并单元格和错误值需要人工判断，未自动处理。"
End Function
