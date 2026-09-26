Attribute VB_Name = "modFields"
'==============================================================================
' modFields - 更新域（word.updateFields）
'
' 更新全文所有域（Document.Fields.Update，含页码、交叉引用等常规域），
' 再单独更新每个目录（TableOfContents.Update）。
'
' 【这两步不能只做一步，真实 COM 调用验证过】：目录本身确实算一个域，
' 会被 Document.Fields 计入（真机测过：只有一个目录时 Fields.Count
' 就是 1），但调用 Document.Fields.Update() 并不会让目录重新扫描文档、
' 收录新增的标题——插入目录之后在文档末尾新加一个"标题 1"样式的段落，
' 只调 Fields.Update() 之后目录文字完全没变；改调
' TablesOfContents(i).Update() 之后新标题才会出现在目录里。
' 两个 API 名字看着像做同一件事，实际行为不同，不能只调一个就假设
' 目录也更新了。
'
' 不做撤销承诺：域更新后原文字会被替换成新计算结果，如果之前那份
' 撤销语义复杂（比如目录展开的条目数变了），这里老实标 Undoable:=False，
' 强制确认。
'==============================================================================
Option Explicit
Option Private Module

Public Function UpdateAllFields() As String
    Dim doc As Document
    Set doc = ActiveDocument

    Dim fieldCount As Long
    fieldCount = doc.Fields.Count
    If fieldCount > 0 Then doc.Fields.Update

    Dim tocCount As Long
    tocCount = doc.TablesOfContents.Count
    Dim i As Long
    For i = 1 To tocCount
        doc.TablesOfContents(i).Update
    Next i

    If fieldCount = 0 And tocCount = 0 Then
        UpdateAllFields = "文档里没有域或目录，无需更新。"
    Else
        UpdateAllFields = "已更新 " & fieldCount & " 个域" & _
                          IIf(tocCount > 0, "，" & tocCount & " 个目录", "") & "。"
    End If
End Function
