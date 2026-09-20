Attribute VB_Name = "modSheetUtil"
'==============================================================================
' modSheetUtil - 工作表基础操作（被多个模块复用）
'==============================================================================
Option Explicit
Option Private Module

' Excel 工作表名的硬性限制
Private Const MAX_SHEET_NAME As Long = 31
Private Const INVALID_CHARS As String = ":\/?*[]"

'------------------------------------------------------------------------------
' 把任意字符串变成合法的工作表名。
'
' Excel 对表名的限制很容易踩到：31 字符上限、7 个非法字符、不能为空、
' 不能叫 History（保留名）。按列拆分工作表时，列值直接拿来当表名必然翻车。
'------------------------------------------------------------------------------
Public Function SafeSheetName(ByVal rawName As String) As String
    Dim buf As String
    buf = Trim$(rawName)

    Dim i As Long
    For i = 1 To Len(INVALID_CHARS)
        buf = Replace(buf, Mid$(INVALID_CHARS, i, 1), "_")
    Next i

    ' 单引号不能出现在首尾
    Do While Left$(buf, 1) = "'"
        buf = Mid$(buf, 2)
    Loop
    Do While Right$(buf, 1) = "'"
        buf = Left$(buf, Len(buf) - 1)
    Loop

    If Len(buf) > MAX_SHEET_NAME Then buf = Left$(buf, MAX_SHEET_NAME)
    If Len(buf) = 0 Then buf = "Sheet"
    If StrComp(buf, "History", vbTextCompare) = 0 Then buf = "History_"

    SafeSheetName = buf
End Function

'------------------------------------------------------------------------------
' 在工作簿内生成一个不重名的表名（追加 _2 / _3 …，并保证总长不超 31）
'------------------------------------------------------------------------------
Public Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String
    candidate = SafeSheetName(baseName)

    If Not SheetExists(wb, candidate) Then
        UniqueSheetName = candidate
        Exit Function
    End If

    Dim n As Long, suffix As String, stem As String
    For n = 2 To 1000
        suffix = "_" & n
        stem = SafeSheetName(baseName)
        If Len(stem) + Len(suffix) > MAX_SHEET_NAME Then
            stem = Left$(stem, MAX_SHEET_NAME - Len(suffix))
        End If
        candidate = stem & suffix
        If Not SheetExists(wb, candidate) Then
            UniqueSheetName = candidate
            Exit Function
        End If
    Next n

    Err.Raise vbObjectError + 360, "modSheetUtil", "无法为「" & baseName & "」生成不重复的表名。"
End Function

Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

'------------------------------------------------------------------------------
' 新建一张工作表，放在最后，名字自动去重
'------------------------------------------------------------------------------
Public Function AddSheet(ByVal wb As Workbook, ByVal baseName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = UniqueSheetName(wb, baseName)
    Set AddSheet = ws
End Function

'------------------------------------------------------------------------------
' 给报表表头统一加粗 + 冻结首行，几个模块都要用
'------------------------------------------------------------------------------
Public Sub FormatHeader(ByVal ws As Worksheet, ByVal colCount As Long)
    With ws.Range("A1").Resize(1, colCount)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With
    ws.Columns.AutoFit
End Sub
