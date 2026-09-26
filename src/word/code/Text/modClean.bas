Attribute VB_Name = "modClean"
'==============================================================================
' modClean - 空格清理（word.cleanSpaces）
'
' 范围（有意收窄，先解决最常见的痛点）：
'   1. 零宽字符（零宽空格 U+200B、零宽不换行空格/BOM U+FEFF）整段删除
'   2. 全角空格（U+3000）、不间断空格（U+00A0）统一换成普通空格
'   3. 连续 2 个及以上普通空格压缩成 1 个
'
' 【不做】段落首尾空格清理：Word 的段落首尾空格视觉上不像 Excel 单元格
' 那样直接影响观感，且要精确处理"首尾"需要逐段落操作而不是全文
' Find/Replace，复杂度和收益不成正比，等有真实需求再加。
'
' 全部用 Find/Replace 实现（真实 COM 调用验证过这个写法有效，
' 见 docs/规划-Word与PPT.md 第八节）：
'   - 精确字符替换（零宽字符、全角空格、NBSP）不开 MatchWildcards
'   - 压缩连续空格用通配符 " {2,}" → " "，必须开 MatchWildcards
'
' 【MatchByte 必须显式设为 True，这是真实 COM 调用踩出来的坑】：这台机器
' 装了中文语言包，Word 的 Find 默认会把全角字符和对应的半角字符当成
' 等价——搜索全角空格（U+3000）时，默认情况下连文档里的普通半角空格
' 也会一起命中（实测：3 字符的文档里搜 1 个全角空格，命中数算出来是
' 3——把两个不相关的半角空格也数了进去）。这不是 Find.Execute 的参数，
' 是 Find 对象单独的 MatchByte 属性，必须在每次用 Find 之前显式设成
' True（"按字节严格匹配，不做全角/半角等价折叠"），不设的话
' 这个命令会把文档里所有普通空格全部替换掉，是一个会破坏用户数据的
' 真实 bug，不是理论风险。
'==============================================================================
Option Explicit
Option Private Module

Private Const CODE_ZERO_WIDTH_SPACE As Long = &H200B
Private Const CODE_ZERO_WIDTH_NBSP As Long = &HFEFF
Private Const CODE_FULLWIDTH_SPACE As Long = &H3000
Private Const CODE_NBSP As Long = &HA0

Public Function CleanSpaces() As String
    Dim doc As Document
    Set doc = ActiveDocument

    Dim removedZeroWidth As Long
    removedZeroWidth = ReplaceExact(doc, ChrW(CODE_ZERO_WIDTH_SPACE), "")
    removedZeroWidth = removedZeroWidth + ReplaceExact(doc, ChrW(CODE_ZERO_WIDTH_NBSP), "")

    Dim normalizedFullwidth As Long
    normalizedFullwidth = ReplaceExact(doc, ChrW(CODE_FULLWIDTH_SPACE), " ")

    Dim normalizedNbsp As Long
    normalizedNbsp = ReplaceExact(doc, ChrW(CODE_NBSP), " ")

    Dim collapsedRuns As Long
    collapsedRuns = ReplaceWildcard(doc, " {2,}", " ")

    Dim total As Long
    total = removedZeroWidth + normalizedFullwidth + normalizedNbsp + collapsedRuns

    If total = 0 Then
        CleanSpaces = "没有发现需要清理的空格。"
    Else
        CleanSpaces = "已清理：零宽字符 " & removedZeroWidth & " 处，" & _
                      "全角空格 " & normalizedFullwidth & " 处，" & _
                      "不间断空格 " & normalizedNbsp & " 处，" & _
                      "连续空格压缩 " & collapsedRuns & " 处。"
    End If
End Function

'------------------------------------------------------------------------------
' 精确字符替换，返回替换次数。Word 的 Find.Execute 不直接返回替换次数
' （只返回是否至少替换过一次的布尔值），所以先数出现次数，再统一替换——
' 这样"清理了几处"这句话里的数字是真的数出来的，不是拍脑袋。
'------------------------------------------------------------------------------
Private Function ReplaceExact(ByVal doc As Document, ByVal findText As String, ByVal replaceText As String) As Long
    If Len(findText) = 0 Then Exit Function

    Dim count As Long
    count = CountOccurrences(doc, findText, False)
    If count = 0 Then Exit Function

    Dim f As Find
    Set f = doc.Content.Find
    f.ClearFormatting
    f.Replacement.ClearFormatting
    f.MatchByte = True   ' 见文件头部说明：不设这个全角空格会连带命中半角空格
    f.Text = findText
    f.Replacement.Text = replaceText
    f.Forward = True
    f.Wrap = wdFindContinue
    f.Format = False
    f.MatchWildcards = False
    f.Execute Replace:=wdReplaceAll

    ReplaceExact = count
End Function

Private Function ReplaceWildcard(ByVal doc As Document, ByVal pattern As String, ByVal replaceText As String) As Long
    Dim count As Long
    count = CountOccurrences(doc, pattern, True)
    If count = 0 Then Exit Function

    Dim f As Find
    Set f = doc.Content.Find
    f.ClearFormatting
    f.Replacement.ClearFormatting
    f.MatchByte = True
    f.Text = pattern
    f.Replacement.Text = replaceText
    f.Forward = True
    f.Wrap = wdFindContinue
    f.Format = False
    f.MatchWildcards = True
    f.Execute Replace:=wdReplaceAll

    ReplaceWildcard = count
End Function

'------------------------------------------------------------------------------
' 数出现次数：在一份 Content 的副本 Range 上反复 Find（不替换），
' 每命中一次把搜索起点挪到命中区域之后，避免死循环。
'------------------------------------------------------------------------------
Private Function CountOccurrences(ByVal doc As Document, ByVal pattern As String, ByVal useWildcards As Boolean) As Long
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    Dim f As Find
    Set f = rng.Find
    f.MatchByte = True

    Dim n As Long
    Do While f.Execute(pattern, False, False, useWildcards, False, False, True, wdFindStop, False)
        n = n + 1
        rng.Collapse wdCollapseEnd
        Set f = rng.Find
        f.MatchByte = True
        ' 极端保护：单份文档里同一个 pattern 命中次数不会有意义地超过
        ' 这个量级，防止 Find 对象行为异常时死循环拖死宏。
        If n > 100000 Then Exit Do
    Loop

    CountOccurrences = n
End Function
