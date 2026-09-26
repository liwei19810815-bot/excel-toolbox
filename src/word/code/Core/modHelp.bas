Attribute VB_Name = "modHelp"
'==============================================================================
' modHelp - 使用帮助（Word 最小实现）
'
' shared/code/Core/modAction.bas 的 RunAction 出错时，如果用户选择查看
' 帮助会调用 modHelp.ShowFor(actionId)，这个符号必须存在才能编译。
'
' 【暂不搬 Excel 那套 1183 行的帮助渲染管线】：Word 首批只有三个命令，
' 帮助内容直接在这里写一份精简的文字说明就够用，不需要 Excel 那套
' "Markdown 转 HTML、生成可搜索帮助页" 的完整机制——那是给 60+ 命令
' 规模准备的，套用在 3 个命令上是过度设计。命令集变大到需要搜索/
' 分类浏览时，再照 Excel 的 modHelp.bas 思路搭一份 Word 版。
'==============================================================================
Option Explicit
Option Private Module

Public Function ShowFor(ByVal actionId As String) As String
    Dim msg As String
    Select Case actionId
        Case "word.audit"
            msg = "「文档体检」：只读扫描，统计空段落、连续空格、" & _
                  "手动换行符、超长段落，不会修改文档。"
        Case "word.cleanSpaces"
            msg = "「清理多余空格」：清理全文的连续空格、全角空格、" & _
                  "不间断空格，合并到一条原生撤销记录，可用 Ctrl+Z 撤销。"
        Case "word.updateFields"
            msg = "「更新域」：更新全文所有域（含目录），域更新后的" & _
                  "撤销语义复杂，不承诺能撤销，执行前会要求确认。"
        Case Else
            msg = "这条命令还没有写帮助正文。"
    End Select

    If Not modAction.IsSilent() Then
        MsgBox msg, vbInformation, modApp.APP_NAME
    End If
    ShowFor = msg
End Function
