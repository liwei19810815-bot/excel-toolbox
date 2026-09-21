Attribute VB_Name = "modStr"
'==============================================================================
' modStr - 字符串工具
'
' 存在的唯一理由：VBA 的 StrConv(s, vbNarrow) / StrConv(s, vbWide) 不能用。
'
' 这两个转换依赖操作系统的 DBCS 支持，在相当多的环境里会直接抛"运行时错误 5：
' 无效的过程调用或参数"——而且是【对任意输入都抛】，连纯 ASCII 字符串也不例外，
' 本机（Excel 2016 64 位 / 中文环境）实测就是如此。更糟的是它不会在编译期暴露，
' 只在运行到那一行时才炸。
'
' 所以全角半角转换一律走这里的显式码位映射：确定、可移植、和区域设置无关。
'
' 映射规则（Unicode 标准的全角形式）：
'   U+FF01..U+FF5E  全角 ！..～   <-> U+0021..U+007E  半角 !..~   （偏移 0xFEE0）
'   U+3000          表意空格      <-> U+0020          普通空格
'
' 注意 VBA 的两个坑：
'   1. 十六进制字面量超过 &H7FFF 会被当成负的 Integer，必须加 & 后缀写成 Long，
'      例如 &HFF5E& 而不是 &HFF5E；
'   2. AscW 返回的是带符号 Integer，码位大于 32767 时返回负数，要补回 65536。
'==============================================================================
Option Explicit
Option Private Module

Private Const FULL_WIDTH_FIRST As Long = &HFF01&
Private Const FULL_WIDTH_LAST As Long = &HFF5E&
Private Const FULL_HALF_OFFSET As Long = &HFEE0&
Private Const IDEOGRAPHIC_SPACE As Long = &H3000&

'------------------------------------------------------------------------------
' 全角 -> 半角
'------------------------------------------------------------------------------
Public Function ToHalfWidth(ByVal srcText As String) As String
    If Len(srcText) = 0 Then Exit Function

    Dim i As Long, code As Long, ch As String
    Dim buf As String

    For i = 1 To Len(srcText)
        ch = Mid$(srcText, i, 1)
        code = CodePointOf(ch)

        If code = IDEOGRAPHIC_SPACE Then
            buf = buf & " "
        ElseIf code >= FULL_WIDTH_FIRST And code <= FULL_WIDTH_LAST Then
            buf = buf & ChrW$(code - FULL_HALF_OFFSET)
        Else
            buf = buf & ch
        End If
    Next i

    ToHalfWidth = buf
End Function

'------------------------------------------------------------------------------
' 半角 -> 全角
'------------------------------------------------------------------------------
Public Function ToFullWidth(ByVal srcText As String) As String
    If Len(srcText) = 0 Then Exit Function

    Dim i As Long, code As Long, ch As String
    Dim buf As String

    For i = 1 To Len(srcText)
        ch = Mid$(srcText, i, 1)
        code = CodePointOf(ch)

        If code = 32 Then
            buf = buf & ChrW$(IDEOGRAPHIC_SPACE)
        ElseIf code >= &H21& And code <= &H7E& Then
            buf = buf & ChrW$(code + FULL_HALF_OFFSET)
        Else
            buf = buf & ch
        End If
    Next i

    ToFullWidth = buf
End Function

'------------------------------------------------------------------------------
' 取单个字符的 Unicode 码位。
' AscW 返回带符号 Integer，U+8000 以上会变成负数，这里补回来。
'
' 【要用码位就调这个，不要自己写 AscW】。直接写 AscW 再和数字比大小，
' 对 U+8000 以上的字符（辰、说、财、货、购、路、车、运、通、部、采、里、
' 金、银、销、长、问、间、题、风、高……全都在这个区间）结果是负数，
' 判断会整片出错。modText.CleanText 就这么丢过字符：
' 「北辰科技」被清洗成「北科技」，而且不报任何错。
'------------------------------------------------------------------------------
Public Function CodePointOf(ByVal ch As String) As Long
    Dim code As Long
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    CodePointOf = code
End Function

'------------------------------------------------------------------------------
' 把"看起来像数字的文本"归一化成可以交给 IsNumeric / CDbl 的形式：
' 全角转半角、去掉千分位逗号、去掉不间断空格和首尾空白。
' 文本型数字、文本型日期的判定都从这里开始。
'------------------------------------------------------------------------------
Public Function NormalizeNumericText(ByVal srcText As String) As String
    Dim buf As String
    buf = ToHalfWidth(srcText)
    buf = Replace(buf, ChrW$(&HA0&), "")      ' 不间断空格
    buf = Replace(buf, ",", "")
    NormalizeNumericText = Trim$(buf)
End Function
