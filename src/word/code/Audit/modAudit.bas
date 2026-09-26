Attribute VB_Name = "modAudit"
'==============================================================================
' modAudit - 文档体检（word.audit）
'
' 只读扫描，不修改文档。统计：
'   - 空段落数
'   - 含手动换行符（Shift+Enter，Chr(11)）的段落数
'   - 超长段落数（阈值见 LONG_PARAGRAPH_THRESHOLD）
'   - 连续空格（2 个及以上）出现次数
'
' 逐段落用 Range.Text 判断前三项（真实 COM 调用验证过：段落文字末尾
' 带一个段落标记字符，取 Text 后要去掉末尾的 Chr(13)/Chr(7) 才能拿到
' 真正的正文长度，否则每段都会多算 1-2 个字符）。连续空格统计复用
' modClean 同款 Find 计数写法，同样要设 MatchByte（这里搜的是纯 ASCII
' 空格，理论上不受全角/半角折叠影响，但统一设置更保险，不差这一行）。
'==============================================================================
Option Explicit
Option Private Module

Private Const LONG_PARAGRAPH_THRESHOLD As Long = 500
Private Const CHAR_MANUAL_BREAK As Long = 11   ' vbVerticalTab / Shift+Enter

Public Function ScanDocument() As String
    Dim doc As Document
    Set doc = ActiveDocument

    Dim emptyCount As Long, manualBreakCount As Long, longCount As Long

    Dim p As Paragraph
    For Each p In doc.Paragraphs
        ' 【手动换行符检测必须在去掉段落标记之前做】：如果 Shift+Enter
        ' 换行恰好是段落里最后一个字符（紧挨着段落标记），先去标记再找
        ' 会把它一起砍掉，误判成"没有手动换行符"——这是真实存在的边界
        ' case，不是假设。
        Dim raw As String
        raw = p.Range.Text
        If InStr(raw, ChrW(CHAR_MANUAL_BREAK)) > 0 Then
            manualBreakCount = manualBreakCount + 1
        End If

        Dim t As String
        t = TrimParagraphMark(raw)

        If Len(t) = 0 Then
            emptyCount = emptyCount + 1
        ElseIf Len(t) > LONG_PARAGRAPH_THRESHOLD Then
            longCount = longCount + 1
        End If
    Next p

    Dim spaceRunCount As Long
    spaceRunCount = CountSpaceRuns(doc)

    If emptyCount = 0 And manualBreakCount = 0 And longCount = 0 And spaceRunCount = 0 Then
        ScanDocument = "体检完成：没有发现问题。共 " & doc.Paragraphs.Count & " 段。"
        Exit Function
    End If

    ScanDocument = "体检完成，共 " & doc.Paragraphs.Count & " 段：" & vbCrLf & _
                   "- 空段落 " & emptyCount & " 个" & vbCrLf & _
                   "- 含手动换行符（Shift+Enter）的段落 " & manualBreakCount & " 个" & vbCrLf & _
                   "- 超长段落（超过 " & LONG_PARAGRAPH_THRESHOLD & " 字）" & longCount & " 个" & vbCrLf & _
                   "- 连续空格 " & spaceRunCount & " 处"
End Function

' 段落 Range.Text 末尾带段落标记（普通段落是 Chr(13)，节末尾/文档末尾
' 可能是 Chr(7) 之类的单元格标记）——真实测试过 Chr(13) 是常规情况，
' 这里做成"去掉末尾控制字符"而不是只砍固定长度，避免遗漏其它标记类型。
'
' 【必须用 modStr.CodePointOf，不能直接 AscW】：AscW 返回带符号 Integer，
' U+8000 以上（一大片常用汉字，比如"龘"U+9F98 之外，连很多常见字都在
' 这个范围）会被解析成负数，负数天然满足"< 32"，会被这个函数当成控制
' 字符一起砍掉——这是 docs/规划-Word与PPT.md 明确点名过的坑
' （Excel 侧已经踩过两次，代价是静默删汉字），第一版写这个函数时
' 直接用了 AscW，属于重蹈覆辙，改用 CodePointOf 修正。
Private Function TrimParagraphMark(ByVal text As String) As String
    Dim s As String
    s = text
    Do While Len(s) > 0 And modStr.CodePointOf(Right$(s, 1)) < 32
        s = Left$(s, Len(s) - 1)
    Loop
    TrimParagraphMark = s
End Function

Private Function CountSpaceRuns(ByVal doc As Document) As Long
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    Dim f As Find
    Set f = rng.Find
    f.MatchByte = True

    Dim n As Long
    Do While f.Execute(" {2,}", False, False, True, False, False, True, wdFindStop, False)
        n = n + 1
        rng.Collapse wdCollapseEnd
        Set f = rng.Find
        f.MatchByte = True
        If n > 100000 Then Exit Do
    Loop

    CountSpaceRuns = n
End Function
