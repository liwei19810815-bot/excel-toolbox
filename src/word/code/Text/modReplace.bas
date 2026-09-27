Attribute VB_Name = "modReplace"
'==============================================================================
' modReplace - 批量替换（word.replaceText）
'
' 全文查找替换，用户输入查找/替换文字（PromptsForInput，走
' modPrompt.AskText）。精确匹配，不支持通配符/正则——Word 的通配符
' 语法和用户熟悉的正则不是一回事，直接暴露给用户会更困惑，不如先做
' 精确匹配这个最常见的用例；真要支持正则再单独设计，不能顺手就宣称
' "正则替换"却只是 Word 通配符的转发。
'
' 和 modClean.CleanSpaces 同一套 Find 写法（MatchByte 必须设 True——
' 见该文件头部注释，全角/半角折叠不是只有搜空格才会踩，任何精确匹配
' 都可能被牵连，统一设置更安全）。
'==============================================================================
Option Explicit
Option Private Module

' Word 的 Find.Text 硬性限制最长 255 字符，超过直接抛"字符串超出规定
' 长度"运行时错误（真实 COM 调用验证过）。提前挡在 AskText 之后、
' 真正调用 Find 之前，给一句能看懂的提示，而不是让用户看着一个语焉
' 不详的运行时错误。
Private Const MAX_FIND_LENGTH As Long = 255

Public Function ReplaceText() As String
    Dim findText As String
    findText = modPrompt.AskText("word.replaceText.find", "要查找的文字（最多 255 字符）：", allowBlank:=False)

    If Len(findText) > MAX_FIND_LENGTH Then
        ReplaceText = "要查找的文字超过 " & MAX_FIND_LENGTH & " 字符（Word 的限制），请缩短后重试。"
        Exit Function
    End If

    Dim replaceWith As String
    replaceWith = modPrompt.AskText("word.replaceText.replace", "替换成（留空表示删除）：", allowBlank:=True)

    Dim doc As Document
    Set doc = ActiveDocument

    Dim count As Long
    count = CountOccurrences(doc, findText)
    If count = 0 Then
        ReplaceText = "全文没有找到「" & findText & "」，未做任何改动。"
        Exit Function
    End If

    Dim f As Find
    Set f = doc.Content.Find
    f.ClearFormatting
    f.Replacement.ClearFormatting
    f.MatchByte = True
    f.Text = findText
    f.Replacement.Text = replaceWith
    f.Forward = True
    f.Wrap = wdFindContinue
    f.Format = False
    f.MatchWildcards = False
    f.Execute Replace:=wdReplaceAll

    ReplaceText = "已把全文 " & count & " 处「" & findText & "」替换为「" & replaceWith & "」。"
End Function

Private Function CountOccurrences(ByVal doc As Document, ByVal pattern As String) As Long
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    Dim f As Find
    Set f = rng.Find
    f.MatchByte = True

    Dim n As Long
    Do While f.Execute(pattern, False, False, False, False, False, True, wdFindStop, False)
        n = n + 1
        rng.Collapse wdCollapseEnd
        Set f = rng.Find
        f.MatchByte = True
        If n > 100000 Then Exit Do
    Loop

    CountOccurrences = n
End Function
