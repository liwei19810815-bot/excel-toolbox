Attribute VB_Name = "modRevisions"
'==============================================================================
' modRevisions - 接受所有修订（word.acceptAllRevisions）
'
' Document.AcceptAllRevisions() 真实 COM 调用验证过：
'   - 零条修订时安全空操作，不报错
'   - 包在 Application.UndoRecord.StartCustomRecord/EndCustomRecord 里
'     不会崩，所以标 Undoable:=True，和 word.cleanSpaces 一样走原生
'     撤销记录，不需要 ConfirmBeforeRun。
'==============================================================================
Option Explicit
Option Private Module

Public Function AcceptAllRevisions() As String
    Dim doc As Document
    Set doc = ActiveDocument

    Dim count As Long
    count = doc.Revisions.Count

    If count = 0 Then
        AcceptAllRevisions = "文档里没有待处理的修订，无需操作。"
        Exit Function
    End If

    doc.AcceptAllRevisions

    AcceptAllRevisions = "已接受 " & count & " 条修订。"
End Function
