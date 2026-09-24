Attribute VB_Name = "modFormula"
'==============================================================================
' modFormula - 公式与引用工具（M6）
'==============================================================================
Option Explicit
Option Private Module

'------------------------------------------------------------------------------
' 公式批量转值。
'
' 不用"复制 -> 选择性粘贴值"：那会清掉剪贴板，而且在多区域选区上直接报错。
' 直接把 .Formula 换成 .Value 更干净，也不影响用户的剪贴板。
'------------------------------------------------------------------------------
Public Function FormulasToValues(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        FormulasToValues = "选区内没有数据。"
        Exit Function
    End If

    Dim formulaRng As Range
    Set formulaRng = FormulaCells(srcRng)
    If formulaRng Is Nothing Then
        FormulasToValues = "选区内没有公式。"
        Exit Function
    End If

    Dim converted As Long
    converted = CLng(modRange.CellCount(formulaRng))

    modUndo.Capture srcRng

    Dim areaRng As Range
    For Each areaRng In formulaRng.Areas
        areaRng.Value = areaRng.Value
    Next areaRng

    FormulasToValues = "已把 " & converted & " 个公式转换为数值。"
End Function

Private Function FormulaCells(ByVal srcRng As Range) As Range
    On Error Resume Next
    Set FormulaCells = srcRng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 定位并统计错误值
'------------------------------------------------------------------------------
Public Function FindErrors(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        FindErrors = "选区内没有数据。"
        Exit Function
    End If

    Dim errRng As Range
    On Error Resume Next
    Set errRng = srcRng.SpecialCells(xlCellTypeFormulas, xlErrors)
    On Error GoTo 0

    Dim constErrRng As Range
    On Error Resume Next
    Set constErrRng = srcRng.SpecialCells(xlCellTypeConstants, xlErrors)
    On Error GoTo 0

    Dim allErr As Range
    If errRng Is Nothing Then
        Set allErr = constErrRng
    ElseIf constErrRng Is Nothing Then
        Set allErr = errRng
    Else
        Set allErr = Application.Union(errRng, constErrRng)
    End If

    If allErr Is Nothing Then
        FindErrors = "选区内没有错误值。"
        Exit Function
    End If

    modUndo.Capture srcRng
    allErr.Interior.Color = RGB(255, 235, 156)
    allErr.Select

    FindErrors = "找到 " & CLng(modRange.CellCount(allErr)) & " 个错误值，已标黄并选中。"
End Function

'------------------------------------------------------------------------------
' 给含公式的单元格批量套 IFERROR。
'
' 已经套过的会跳过，重复执行不会叠成 IFERROR(IFERROR(...))。
'------------------------------------------------------------------------------
Public Function WrapWithIfError(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        WrapWithIfError = "选区内没有数据。"
        Exit Function
    End If

    Dim formulaRng As Range
    Set formulaRng = FormulaCells(srcRng)
    If formulaRng Is Nothing Then
        WrapWithIfError = "选区内没有公式。"
        Exit Function
    End If

    Dim fallback As String
    fallback = modPrompt.AskText("fallback", _
        "公式出错时显示什么？（直接确定则显示空白）", "", True)

    Dim fallbackExpr As String
    If Len(fallback) = 0 Then
        fallbackExpr = """"""
    ElseIf IsNumeric(fallback) Then
        fallbackExpr = fallback
    Else
        fallbackExpr = """" & Replace(fallback, """", """""") & """"
    End If

    modUndo.Capture srcRng

    Dim cellRng As Range, oldFormula As String, wrapped As Long, skipped As Long
    For Each cellRng In formulaRng.Cells
        oldFormula = cellRng.Formula
        If Left$(oldFormula, 1) = "=" Then
            If UCase$(Left$(oldFormula, 9)) = "=IFERROR(" Then
                skipped = skipped + 1
            Else
                cellRng.Formula = "=IFERROR(" & Mid$(oldFormula, 2) & "," & fallbackExpr & ")"
                wrapped = wrapped + 1
            End If
        End If
    Next cellRng

    WrapWithIfError = "已为 " & wrapped & " 个公式添加 IFERROR" & _
                      IIf(skipped > 0, "，跳过 " & skipped & " 个已有 IFERROR 的公式", "") & "。"
End Function

'------------------------------------------------------------------------------
' 断开所有外部链接（转为当前值）。
'
' 这是典型的不可撤销操作——链接断了就找不回来源了，所以要强制确认。
'------------------------------------------------------------------------------
Public Function BreakExternalLinks(ByVal wb As Workbook) As String
    Dim links As Variant
    links = wb.LinkSources(xlExcelLinks)

    If IsEmpty(links) Then
        BreakExternalLinks = "当前工作簿没有外部链接。"
        Exit Function
    End If

    Dim i As Long, broken As Long
    For i = LBound(links) To UBound(links)
        On Error Resume Next
        wb.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
        If Err.Number = 0 Then broken = broken + 1
        On Error GoTo 0
    Next i

    BreakExternalLinks = "已断开 " & broken & " 个外部链接（公式已转为当前值）。"
End Function

'------------------------------------------------------------------------------
' 清理无效的已定义名称。
'
' 指向 #REF! 的名称是工作簿体积膨胀和"打开时莫名弹框"的常见原因，
' 而 Excel 的名称管理器要一个一个删。
'------------------------------------------------------------------------------
Public Function CleanBrokenNames(ByVal wb As Workbook) As String
    Dim nm As Object, removed As Long, kept As Long
    Dim refersTo As String

    Dim i As Long
    For i = wb.Names.Count To 1 Step -1          ' 倒着遍历，删除不会影响后续索引
        Set nm = wb.Names(i)
        refersTo = ""
        On Error Resume Next
        refersTo = nm.RefersTo
        On Error GoTo 0

        If InStr(1, refersTo, "#REF!", vbTextCompare) > 0 Then
            On Error Resume Next
            nm.Delete
            If Err.Number = 0 Then removed = removed + 1
            On Error GoTo 0
        Else
            kept = kept + 1
        End If
    Next i

    If removed = 0 Then
        CleanBrokenNames = "没有找到失效的已定义名称（共 " & kept & " 个名称）。"
    Else
        CleanBrokenNames = "已删除 " & removed & " 个失效名称，保留 " & kept & " 个。"
    End If
End Function

'------------------------------------------------------------------------------
' 显示/隐藏全部公式（等价于 Ctrl+`，但做成可发现的按钮）
'------------------------------------------------------------------------------
Public Function ToggleFormulaView(ByVal ws As Worksheet) As String
    Dim win As Window
    Set win = ws.Parent.Windows(1)
    win.DisplayFormulas = Not win.DisplayFormulas

    ToggleFormulaView = IIf(win.DisplayFormulas, "已切换为显示公式。", "已切换为显示计算结果。")
End Function

'------------------------------------------------------------------------------
' 清理选区内多余的条件格式与数据验证。
'
' 反复复制粘贴会让同一片区域堆叠出成百上千条重复的条件格式规则，
' 表现为文件越来越大、滚动越来越卡。
'------------------------------------------------------------------------------
Public Function CleanFormatRules(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        CleanFormatRules = "选区内没有数据。"
        Exit Function
    End If

    Dim cfCount As Long, dvCount As Long
    cfCount = srcRng.FormatConditions.Count

    On Error Resume Next
    dvCount = 0
    Dim probe As Long
    probe = srcRng.Validation.Type          ' 没有验证时会抛错
    If Err.Number = 0 Then dvCount = 1
    Err.Clear
    On Error GoTo 0

    If cfCount = 0 And dvCount = 0 Then
        CleanFormatRules = "选区内没有条件格式或数据验证。"
        Exit Function
    End If

    modUndo.Capture srcRng

    srcRng.FormatConditions.Delete
    On Error Resume Next
    srcRng.Validation.Delete
    On Error GoTo 0

    CleanFormatRules = "已清除 " & cfCount & " 条条件格式规则" & _
                       IIf(dvCount > 0, "和选区内的数据验证", "") & "。"
End Function
