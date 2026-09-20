Attribute VB_Name = "modText"
'==============================================================================
' modText - 文本处理（M1）
'
' 全部走"一次性读入数组 -> 内存处理 -> 一次性写回"，不逐格读写。
' 只在确实有单元格发生变化时才写回，避免把一整片没动过的区域标记为已修改。
'==============================================================================
Option Explicit
Option Private Module

'==============================================================================
' 通用骨架
'==============================================================================

'------------------------------------------------------------------------------
' 对选区每个单元格的文本做一次变换。
'
' 只处理常量文本，跳过公式单元格——把公式的计算结果写回去会直接毁掉公式，
' 这是文本类工具最容易造成的不可逆损失。
'------------------------------------------------------------------------------
Private Function TransformCells(ByVal target As Range, _
                                ByVal transformId As String, _
                                ByVal arg1 As String, _
                                ByVal arg2 As String) As Long
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then Exit Function

    Dim areaRng As Range, changed As Long
    For Each areaRng In srcRng.Areas
        changed = changed + TransformArea(areaRng, transformId, arg1, arg2)
    Next areaRng

    TransformCells = changed
End Function

Private Function TransformArea(ByVal areaRng As Range, _
                               ByVal transformId As String, _
                               ByVal arg1 As String, _
                               ByVal arg2 As String) As Long
    Dim srcArr As Variant, outArr As Variant
    srcArr = modRange.ToArray(areaRng)
    outArr = srcArr

    ' 公式位置单独记下来，写回时原样跳过
    Dim hasFormula As Variant
    hasFormula = areaRng.HasFormula

    Dim rowIdx As Long, colIdx As Long, changed As Long
    Dim oldVal As String, newVal As String
    Dim cellIsFormula As Boolean

    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
            If VarType(hasFormula) = vbBoolean Then
                cellIsFormula = CBool(hasFormula)
            Else
                cellIsFormula = areaRng.Cells(rowIdx, colIdx).HasFormula
            End If
            If Not cellIsFormula Then
                If Not IsError(srcArr(rowIdx, colIdx)) Then
                    oldVal = CStr(Nz(srcArr(rowIdx, colIdx)))
                    If Len(oldVal) > 0 Then
                        newVal = ApplyTransform(transformId, oldVal, arg1, arg2)
                        If newVal <> oldVal Then
                            outArr(rowIdx, colIdx) = newVal
                            changed = changed + 1
                        End If
                    End If
                End If
            End If
        Next colIdx
    Next rowIdx

    If changed > 0 Then
        modUndo.Capture areaRng
        ' 必须走 WriteBack 而不是 FromArray：整块写回会把公式覆盖成它的计算结果
        modRange.WriteBack areaRng, srcArr, outArr
    End If

    TransformArea = changed
End Function

Private Function Nz(ByVal v As Variant) As Variant
    If IsNull(v) Then Nz = "" Else Nz = v
End Function

Private Function ApplyTransform(ByVal transformId As String, _
                                ByVal srcText As String, _
                                ByVal arg1 As String, _
                                ByVal arg2 As String) As String
    Select Case transformId
        Case "trim":        ApplyTransform = CleanText(srcText)
        Case "upper":       ApplyTransform = UCase$(srcText)
        Case "lower":       ApplyTransform = LCase$(srcText)
        Case "proper":      ApplyTransform = Application.WorksheetFunction.Proper(srcText)
        Case "halfWidth":   ApplyTransform = modStr.ToHalfWidth(srcText)
        Case "fullWidth":   ApplyTransform = modStr.ToFullWidth(srcText)
        Case "noBreaks":    ApplyTransform = Replace(Replace(srcText, vbCrLf, ""), vbLf, "")
        Case "digits":      ApplyTransform = KeepPattern(srcText, "[0-9]")
        Case "chinese":     ApplyTransform = KeepPattern(srcText, "[一-龥]")
        Case "english":     ApplyTransform = KeepPattern(srcText, "[A-Za-z]")
        Case "affix":       ApplyTransform = arg1 & srcText & arg2
        Case "regex":       ApplyTransform = RegexReplace(srcText, arg1, arg2)
        Case Else
            Err.Raise vbObjectError + 320, "modText", "未知的文本变换：" & transformId
    End Select
End Function

'------------------------------------------------------------------------------
' 清理空白与不可见字符。
'
' 光用 Trim 是不够的：从网页和系统导出的数据里最常见的是不间断空格（U+00A0）
' 和零宽字符，它们看起来就是空格，但 Trim 不认，导致"看着一样却匹配不上"。
'------------------------------------------------------------------------------
Private Function CleanText(ByVal srcText As String) As String
    Dim outText As String
    outText = srcText

    outText = Replace(outText, ChrW$(&HA0), " ")      ' 不间断空格
    outText = Replace(outText, ChrW$(&H3000), " ")    ' 全角空格
    outText = Replace(outText, ChrW$(&H200B), "")     ' 零宽空格
    outText = Replace(outText, ChrW$(&HFEFF), "")     ' 零宽不换行空格 / BOM
    outText = Replace(outText, vbTab, " ")

    ' 控制字符（含换行）一律去掉
    Dim i As Long, ch As String, buf As String
    For i = 1 To Len(outText)
        ch = Mid$(outText, i, 1)
        If AscW(ch) >= 32 Or ch = " " Then buf = buf & ch
    Next i

    ' 首尾去空格，中间连续空格压成一个
    buf = Trim$(buf)
    Do While InStr(buf, "  ") > 0
        buf = Replace(buf, "  ", " ")
    Loop

    CleanText = buf
End Function

Private Function KeepPattern(ByVal srcText As String, ByVal charClass As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "[^" & Mid$(charClass, 2, Len(charClass) - 2) & "]"
    KeepPattern = re.Replace(srcText, "")
End Function

Private Function RegexReplace(ByVal srcText As String, _
                              ByVal pattern As String, _
                              ByVal replacement As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = pattern
    RegexReplace = re.Replace(srcText, replacement)
End Function

Private Function DoneMsg(ByVal changed As Long) As String
    If changed = 0 Then
        DoneMsg = "没有需要修改的单元格。"
    Else
        DoneMsg = "已处理 " & changed & " 个单元格。"
    End If
End Function

'==============================================================================
' 对外命令
'==============================================================================

Public Function CleanSpaces(ByVal target As Range) As String
    CleanSpaces = DoneMsg(TransformCells(target, "trim", "", ""))
End Function

Public Function ToUpper(ByVal target As Range) As String
    ToUpper = DoneMsg(TransformCells(target, "upper", "", ""))
End Function

Public Function ToLower(ByVal target As Range) As String
    ToLower = DoneMsg(TransformCells(target, "lower", "", ""))
End Function

Public Function ToProper(ByVal target As Range) As String
    ToProper = DoneMsg(TransformCells(target, "proper", "", ""))
End Function

Public Function ToHalfWidth(ByVal target As Range) As String
    ToHalfWidth = DoneMsg(TransformCells(target, "halfWidth", "", ""))
End Function

Public Function ToFullWidth(ByVal target As Range) As String
    ToFullWidth = DoneMsg(TransformCells(target, "fullWidth", "", ""))
End Function

Public Function RemoveLineBreaks(ByVal target As Range) As String
    RemoveLineBreaks = DoneMsg(TransformCells(target, "noBreaks", "", ""))
End Function

Public Function ExtractDigits(ByVal target As Range) As String
    ExtractDigits = DoneMsg(TransformCells(target, "digits", "", ""))
End Function

Public Function ExtractChinese(ByVal target As Range) As String
    ExtractChinese = DoneMsg(TransformCells(target, "chinese", "", ""))
End Function

Public Function ExtractEnglish(ByVal target As Range) As String
    ExtractEnglish = DoneMsg(TransformCells(target, "english", "", ""))
End Function

Public Function AddAffix(ByVal target As Range) As String
    Dim prefixText As String, suffixText As String
    prefixText = modPrompt.AskText("prefix", "要添加的前缀（不需要就直接确定）：", "", True)
    suffixText = modPrompt.AskText("suffix", "要添加的后缀（不需要就直接确定）：", "", True)

    If Len(prefixText) = 0 And Len(suffixText) = 0 Then
        AddAffix = "前缀和后缀都为空，未做修改。"
        Exit Function
    End If

    AddAffix = DoneMsg(TransformCells(target, "affix", prefixText, suffixText))
End Function

Public Function RegexReplaceCells(ByVal target As Range) As String
    Dim pattern As String, replacement As String
    pattern = modPrompt.AskText("pattern", "正则表达式：")
    replacement = modPrompt.AskText("replacement", "替换为（可为空）：", "", True)

    ' 先验一下正则是否合法，否则错误会在循环里反复抛，信息还不直观
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    On Error GoTo BadPattern
    re.Pattern = pattern
    Dim probe As String
    probe = re.Replace("test", replacement)
    On Error GoTo 0

    RegexReplaceCells = DoneMsg(TransformCells(target, "regex", pattern, replacement))
    Exit Function

BadPattern:
    Err.Raise vbObjectError + 321, "modText.RegexReplaceCells", _
              "正则表达式无效：" & pattern
End Function

'------------------------------------------------------------------------------
' 文本型数字 -> 数值。
'
' 这是中文办公环境里最高频的痛点之一：从系统导出的数字常带不间断空格、全角
' 数字或千分位逗号，SUM 结果是 0 却看不出哪里不对。
'------------------------------------------------------------------------------
Public Function TextToNumber(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        TextToNumber = "选区内没有数据。"
        Exit Function
    End If

    Dim areaRng As Range, changed As Long
    For Each areaRng In srcRng.Areas
        changed = changed + TextToNumberArea(areaRng)
    Next areaRng

    TextToNumber = DoneMsg(changed)
End Function

Private Function TextToNumberArea(ByVal areaRng As Range) As Long
    Dim srcArr As Variant, outArr As Variant
    srcArr = modRange.ToArray(areaRng)
    outArr = srcArr

    Dim rowIdx As Long, colIdx As Long, changed As Long
    Dim rawText As String, cleaned As String

    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
            If VarType(srcArr(rowIdx, colIdx)) = vbString Then
                rawText = CStr(srcArr(rowIdx, colIdx))
                cleaned = modStr.NormalizeNumericText(rawText)
                If Len(cleaned) > 0 And IsNumeric(cleaned) Then
                    outArr(rowIdx, colIdx) = CDbl(cleaned)
                    changed = changed + 1
                End If
            End If
        Next colIdx
    Next rowIdx

    If changed > 0 Then
        modUndo.Capture areaRng
        ' 文本型数字所在单元格往往被设成了"文本"格式，不改回常规的话写进去还是文本
        areaRng.NumberFormat = "General"
        modRange.WriteBack areaRng, srcArr, outArr
    End If

    TextToNumberArea = changed
End Function

'------------------------------------------------------------------------------
' 按分隔符拆分为多列。
'
' 不用 TextToColumns：它受系统区域设置影响，且会静默覆盖右侧已有数据。
' 这里先算出需要几列、检查右侧是否会被覆盖，再插入足够的空列。
'------------------------------------------------------------------------------
Public Function SplitColumn(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        SplitColumn = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Columns.Count <> 1 Then
        SplitColumn = "请只选中一列。"
        Exit Function
    End If

    Dim sepText As String
    sepText = modPrompt.AskText("separator", "分隔符（例如 , 或 - 或 |）：", ",")

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    ' 先扫一遍决定要几列
    Dim rowIdx As Long, maxParts As Long, parts As Variant
    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        parts = Split(CStr(Nz(srcArr(rowIdx, 1))), sepText)
        If UBound(parts) + 1 > maxParts Then maxParts = UBound(parts) + 1
    Next rowIdx

    If maxParts <= 1 Then
        SplitColumn = "没有找到分隔符「" & sepText & "」，未做修改。"
        Exit Function
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To UBound(srcArr, 1), 1 To maxParts)
    Dim colIdx As Long
    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        parts = Split(CStr(Nz(srcArr(rowIdx, 1))), sepText)
        For colIdx = 0 To UBound(parts)
            outArr(rowIdx, colIdx + 1) = parts(colIdx)
        Next colIdx
    Next rowIdx

    ' 拆分会改变列结构，属于结构性操作，必须整表快照
    modUndo.CaptureSheet ws

    ' 插入 maxParts-1 列，避免覆盖右侧已有数据
    Dim insertAt As Long
    insertAt = srcRng.Column + 1
    ws.Range(ws.Columns(insertAt), ws.Columns(insertAt + maxParts - 2)).Insert Shift:=xlToRight

    ws.Cells(srcRng.Row, srcRng.Column).Resize(UBound(srcArr, 1), maxParts).Value = outArr

    SplitColumn = "已拆分为 " & maxParts & " 列。"
End Function
