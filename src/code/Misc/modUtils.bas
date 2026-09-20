Attribute VB_Name = "modUtils"
'==============================================================================
' modUtils - 辅助增强（M9）
'
' 金额大写、身份证解析、日期规范化——都是中文办公场景里高频、
' 但 Excel 内置函数覆盖不到的东西。
'==============================================================================
Option Explicit
Option Private Module

'==============================================================================
' 金额小写转大写
'==============================================================================

'------------------------------------------------------------------------------
' 人民币金额大写。
'
' 规则比想象中繁琐，坑主要在"零"的处理上：
'   - 连续多个零只写一个"零"
'   - 整数部分每个节（万、亿）末尾的零不写，但节与节之间要补零
'   - 角位为零而分位不为零时，要写"零"（如 100.05 -> 壹佰元零伍分）
'   - 无小数时补"整"
'------------------------------------------------------------------------------
Public Function AmountToChinese(ByVal amount As Double) As String
    Const DIGITS As String = "零壹贰叁肆伍陆柒捌玖"

    If amount = 0 Then
        AmountToChinese = "零元整"
        Exit Function
    End If

    Dim signText As String
    If amount < 0 Then
        signText = "负"
        amount = -amount
    End If

    If amount >= 1E+15 Then
        AmountToChinese = "金额超出可表示范围"
        Exit Function
    End If

    ' 四舍五入到分，避免浮点误差让 0.1+0.2 变成 0.30000000000000004
    Dim totalCents As Currency
    totalCents = CCur(Round(amount * 100, 0))

    Dim yuanPart As Currency, centPart As Long
    yuanPart = Int(totalCents / 100)
    centPart = CLng(totalCents - yuanPart * 100)

    Dim result As String
    result = IntegerToChinese(CDbl(yuanPart), DIGITS)
    If Len(result) > 0 Then result = result & "元"

    Dim jiao As Long, fen As Long
    jiao = centPart \ 10
    fen = centPart Mod 10

    If jiao = 0 And fen = 0 Then
        result = result & "整"
    Else
        If jiao > 0 Then
            result = result & Mid$(DIGITS, jiao + 1, 1) & "角"
        ElseIf yuanPart > 0 Then
            result = result & "零"           ' 100.05 -> 壹佰元零伍分
        End If
        If fen > 0 Then
            result = result & Mid$(DIGITS, fen + 1, 1) & "分"
        End If
    End If

    AmountToChinese = signText & result
End Function

'------------------------------------------------------------------------------
' 整数部分转中文。按"节"（每 4 位）处理，这是中文数字的自然分组方式。
'------------------------------------------------------------------------------
Private Function IntegerToChinese(ByVal n As Double, ByVal digits As String) As String
    If n = 0 Then
        IntegerToChinese = "零"
        Exit Function
    End If

    Dim sectionNames As Variant
    sectionNames = Array("", "万", "亿", "万亿")

    Dim result As String, sectionIdx As Long
    Dim sectionVal As Long, sectionText As String
    Dim needZero As Boolean

    ' 从低位节向高位节处理。needZero 表示"已处理的低位部分缺了千位"，
    ' 也就是更高的节接上来时中间要补一个零：
    '     10001  -> 壹万【零】壹      低节 1 不足四位
    '     100000001 -> 壹亿【零】壹   中间整个万节为零
    '     101000 -> 壹拾万壹仟        低节 1000 占满千位，不补零
    ' 判据统一为"低位部分非空且其值小于 1000"。
    Do While n > 0
        sectionVal = CLng(n - Int(n / 10000) * 10000)
        sectionText = SectionToChinese(sectionVal, digits)

        If Len(sectionText) > 0 Then
            result = sectionText & sectionNames(sectionIdx) & IIf(needZero, "零", "") & result
        End If

        ' 整节为零时不写节名，但"缺千位"的标记要保留给更高的节
        If Len(result) > 0 Then needZero = (sectionVal < 1000)

        n = Int(n / 10000)
        sectionIdx = sectionIdx + 1
        If sectionIdx > 3 Then Exit Do
    Loop

    IntegerToChinese = result
End Function

'------------------------------------------------------------------------------
' 四位以内的数字转中文
'------------------------------------------------------------------------------
Private Function SectionToChinese(ByVal n As Long, ByVal digits As String) As String
    If n = 0 Then Exit Function

    Dim units As Variant
    units = Array("", "拾", "佰", "仟")

    Dim result As String, pos As Long, d As Long
    Dim zeroPending As Boolean

    pos = 0
    Do While n > 0
        d = n Mod 10
        If d = 0 Then
            ' 连续的零只保留一个，且末尾的零不写
            If Len(result) > 0 Then zeroPending = True
        Else
            If zeroPending Then
                result = "零" & result
                zeroPending = False
            End If
            result = Mid$(digits, d + 1, 1) & units(pos) & result
        End If
        n = n \ 10
        pos = pos + 1
    Loop

    SectionToChinese = result
End Function

'------------------------------------------------------------------------------
' 批量把选区金额转成大写，写到右侧一列
'------------------------------------------------------------------------------
Public Function AmountColumnToChinese(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        AmountColumnToChinese = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Columns.Count <> 1 Then
        AmountColumnToChinese = "请只选中一列金额。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim outRng As Range
    Set outRng = ws.Cells(srcRng.Row, srcRng.Column + 1).Resize(srcRng.Rows.Count, 1)
    modUndo.Capture outRng

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim outArr() As Variant
    ReDim outArr(1 To UBound(srcArr, 1), 1 To 1)

    Dim rowIdx As Long, converted As Long
    For rowIdx = 1 To UBound(srcArr, 1)
        If IsNumeric(srcArr(rowIdx, 1)) And Not IsEmpty(srcArr(rowIdx, 1)) Then
            outArr(rowIdx, 1) = AmountToChinese(CDbl(srcArr(rowIdx, 1)))
            converted = converted + 1
        Else
            outArr(rowIdx, 1) = ""
        End If
    Next rowIdx

    outRng.Value = outArr
    ws.Columns(srcRng.Column + 1).AutoFit

    AmountColumnToChinese = "已转换 " & converted & " 个金额为中文大写，写入右侧一列。"
End Function

'==============================================================================
' 身份证解析
'==============================================================================

'------------------------------------------------------------------------------
' 解析中国大陆 18 位身份证：出生日期、性别、校验位是否正确。
'
' 校验位用的是国标 GB 11643 的 ISO 7064:1983 MOD 11-2 算法。
' 不验校验位的话，随便编一串 18 位数字也能"解析"出生日，反而更危险。
'------------------------------------------------------------------------------
Public Function ParseIdCards(ByVal target As Range) As String
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        ParseIdCards = "选区内没有数据。"
        Exit Function
    End If
    If srcRng.Columns.Count <> 1 Then
        ParseIdCards = "请只选中一列身份证号。"
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = srcRng.Worksheet

    Dim outRng As Range
    Set outRng = ws.Cells(srcRng.Row, srcRng.Column + 1).Resize(srcRng.Rows.Count, 4)
    modUndo.Capture ws.Cells(srcRng.Row - 1, srcRng.Column + 1).Resize(srcRng.Rows.Count + 1, 4)

    ' 表头写在选区上面一行
    If srcRng.Row > 1 Then
        ws.Cells(srcRng.Row - 1, srcRng.Column + 1).Resize(1, 4).Value = _
            Array("出生日期", "性别", "年龄", "校验")
    End If

    Dim srcArr As Variant
    srcArr = modRange.ToArray(srcRng)

    Dim outArr() As Variant
    ReDim outArr(1 To UBound(srcArr, 1), 1 To 4)

    Dim rowIdx As Long, idText As String, okCount As Long, badCount As Long
    Dim birthDate As Date, genderText As String

    For rowIdx = 1 To UBound(srcArr, 1)
        idText = Trim$(UCase$(CStr(srcArr(rowIdx, 1))))

        If Len(idText) = 18 And IsValidIdFormat(idText) Then
            On Error Resume Next
            birthDate = DateSerial(CInt(Mid$(idText, 7, 4)), _
                                   CInt(Mid$(idText, 11, 2)), _
                                   CInt(Mid$(idText, 13, 2)))
            If Err.Number <> 0 Then
                Err.Clear
                outArr(rowIdx, 1) = ""
                outArr(rowIdx, 4) = "出生日期无效"
                badCount = badCount + 1
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0

            genderText = IIf(CInt(Mid$(idText, 17, 1)) Mod 2 = 1, "男", "女")

            outArr(rowIdx, 1) = birthDate
            outArr(rowIdx, 2) = genderText
            outArr(rowIdx, 3) = AgeFrom(birthDate)

            If CheckDigit(idText) = Right$(idText, 1) Then
                outArr(rowIdx, 4) = "正确"
                okCount = okCount + 1
            Else
                outArr(rowIdx, 4) = "校验位错误"
                badCount = badCount + 1
            End If
        ElseIf Len(idText) > 0 Then
            outArr(rowIdx, 4) = "格式不正确"
            badCount = badCount + 1
        End If
NextRow:
    Next rowIdx

    outRng.Value = outArr
    outRng.Columns(1).NumberFormat = "yyyy-mm-dd"
    ws.Columns(srcRng.Column + 1).Resize(1, 4).AutoFit

    ParseIdCards = "解析完成：有效 " & okCount & " 条" & _
                   IIf(badCount > 0, "，有问题 " & badCount & " 条", "") & "。"
End Function

Private Function IsValidIdFormat(ByVal idText As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To 17
        ch = Mid$(idText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ch = Right$(idText, 1)
    IsValidIdFormat = ((ch >= "0" And ch <= "9") Or ch = "X")
End Function

Private Function CheckDigit(ByVal idText As String) As String
    Dim weights As Variant, codes As Variant
    weights = Array(7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)
    codes = Array("1", "0", "X", "9", "8", "7", "6", "5", "4", "3", "2")

    Dim total As Long, i As Long
    For i = 0 To 16
        total = total + CLng(Mid$(idText, i + 1, 1)) * weights(i)
    Next i

    CheckDigit = codes(total Mod 11)
End Function

Private Function AgeFrom(ByVal birthDate As Date) As Long
    Dim age As Long
    age = Year(Date) - Year(birthDate)
    If Month(Date) < Month(birthDate) Then
        age = age - 1
    ElseIf Month(Date) = Month(birthDate) And Day(Date) < Day(birthDate) Then
        age = age - 1
    End If
    AgeFrom = age
End Function

'==============================================================================
' 日期规范化
'==============================================================================

'------------------------------------------------------------------------------
' 把各种写法的文本日期统一成真正的日期值。
'
' 从各种系统导出的日期五花八门：20240115、2024.01.15、2024年1月15日、
' 2024/1/15。它们看着是日期，但全都是文本，不能排序也不能做日期计算。
'------------------------------------------------------------------------------
Public Function NormalizeDates(ByVal target As Range) As String
    Dim phase As String
    On Error GoTo Failed

    phase = "归一化选区"
    Dim srcRng As Range
    Set srcRng = modRange.NormalizeSelection(target)
    If srcRng Is Nothing Then
        NormalizeDates = "选区内没有数据。"
        Exit Function
    End If

    phase = "扫描并转换"
    Dim areaRng As Range, changed As Long
    For Each areaRng In srcRng.Areas
        changed = changed + NormalizeArea(areaRng)
    Next areaRng

    If changed = 0 Then
        NormalizeDates = "没有找到可转换的文本日期。"
    Else
        NormalizeDates = "已把 " & changed & " 个文本日期转换为日期值。"
    End If
    Exit Function

Failed:
    Err.Raise Err.Number, "modUtils.NormalizeDates[" & phase & "]", Err.Description
End Function

Private Function NormalizeArea(ByVal areaRng As Range) As Long
    Dim phase As String
    On Error GoTo Failed

    phase = "读入数组"
    Dim srcArr As Variant, outArr As Variant
    srcArr = modRange.ToArray(areaRng)
    outArr = srcArr

    Dim rowIdx As Long, colIdx As Long, changed As Long
    Dim parsed As Date

    For rowIdx = LBound(srcArr, 1) To UBound(srcArr, 1)
        For colIdx = LBound(srcArr, 2) To UBound(srcArr, 2)
            If VarType(srcArr(rowIdx, colIdx)) = vbString Then
                phase = "解析 r" & rowIdx & "c" & colIdx & " [" & CStr(srcArr(rowIdx, colIdx)) & "]"
                If TryParseDate(CStr(srcArr(rowIdx, colIdx)), parsed) Then
                    outArr(rowIdx, colIdx) = parsed
                    changed = changed + 1
                End If
            End If
        Next colIdx
    Next rowIdx

    If changed > 0 Then
        phase = "快照"
        modUndo.Capture areaRng

        ' 【必须先改格式再写值】：源单元格往往是文本格式（本来就是文本日期），
        ' 往文本格式的单元格里写日期，Excel 会原样存成文本字符串，
        ' 事后再设日期格式也救不回来——值已经不是日期了。
        phase = "设置日期格式"
        areaRng.NumberFormat = "yyyy-mm-dd"

        phase = "写回"
        modRange.WriteBack areaRng, srcArr, outArr
    End If

    NormalizeArea = changed
    Exit Function

Failed:
    Err.Raise Err.Number, "modUtils.NormalizeArea[" & phase & "]", Err.Description
End Function

Private Function TryParseDate(ByVal rawText As String, ByRef result As Date) As Boolean
    Dim txt As String
    txt = Trim$(modStr.ToHalfWidth(rawText))
    If Len(txt) = 0 Then Exit Function

    ' 中文年月日
    txt = Replace(txt, "年", "-")
    txt = Replace(txt, "月", "-")
    txt = Replace(txt, "日", "")
    txt = Replace(txt, ".", "-")

    ' 纯 8 位数字：20240115
    If Len(txt) = 8 And IsAllDigits(txt) Then
        On Error Resume Next
        result = DateSerial(CInt(Left$(txt, 4)), CInt(Mid$(txt, 5, 2)), CInt(Right$(txt, 2)))
        If Err.Number = 0 Then TryParseDate = True
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    Do While Right$(txt, 1) = "-"
        txt = Left$(txt, Len(txt) - 1)
    Loop

    If IsDate(txt) Then
        result = CDate(txt)
        TryParseDate = True
    End If
End Function

Private Function IsAllDigits(ByVal txt As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function
