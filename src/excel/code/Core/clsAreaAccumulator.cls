VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAreaAccumulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsAreaAccumulator - 分块累积多区域
'
' 解决两个坑：
'   1. 在循环里反复 Application.Union 是 O(n^2)，几千个区域就会卡到不可用；
'   2. Range("A1,A5,...") 的地址字符串有 255 字符上限，超了直接报错。
'
' 做法：先把地址攒成字符串，逼近长度上限时 flush 成一个 Range 再 Union。
' Union 的调用次数因此从 N 降到 N/几十，十万行量级也能秒出。
'
' 用法：
'       Dim acc As New clsAreaAccumulator
'       acc.Init ws
'       For i = 1 To n
'           If 需要 Then acc.AddRow i          ' 或 acc.Add "A" & i
'       Next
'       Dim rng As Range: Set rng = acc.Result  ' 无任何区域时返回 Nothing
'==============================================================================
Option Explicit

' Range() 地址字符串上限 255 字符，留余量
Private Const ADDR_CHUNK_LIMIT As Long = 200

Private mWs As Worksheet
Private mBuffer As String
Private mResult As Range
Private mCount As Long

Public Sub Init(ByVal ws As Worksheet)
    Set mWs = ws
    mBuffer = vbNullString
    Set mResult = Nothing
    mCount = 0
End Sub

' 追加一个地址片段，如 "A5" / "B2:D9" / "7:7"
Public Sub Add(ByVal addr As String)
    If mWs Is Nothing Then Err.Raise vbObjectError + 100, "clsAreaAccumulator", "使用前必须先调用 Init。"

    If Len(mBuffer) = 0 Then
        mBuffer = addr
    ElseIf Len(mBuffer) + Len(addr) + 1 > ADDR_CHUNK_LIMIT Then
        Flush
        mBuffer = addr
    Else
        mBuffer = mBuffer & "," & addr
    End If

    mCount = mCount + 1
End Sub

' 追加整行（供批量删行使用）
Public Sub AddRow(ByVal rowIndex As Long)
    Add rowIndex & ":" & rowIndex
End Sub

' 追加整列
Public Sub AddColumn(ByVal colIndex As Long)
    Dim letter As String
    letter = Split(mWs.Cells(1, colIndex).Address(True, False), "$")(0)
    Add letter & ":" & letter
End Sub

Public Sub AddRange(ByVal rng As Range)
    Add rng.Address(False, False)
End Sub

Private Sub Flush()
    If Len(mBuffer) = 0 Then Exit Sub

    Dim chunk As Range
    Set chunk = mWs.Range(mBuffer)

    If mResult Is Nothing Then
        Set mResult = chunk
    Else
        Set mResult = Application.Union(mResult, chunk)
    End If

    mBuffer = vbNullString
End Sub

Public Property Get Result() As Range
    Flush
    Set Result = mResult
End Property

' 已累积的片段个数（不是单元格数）
Public Property Get Count() As Long
    Count = mCount
End Property

' 不叫 IsEmpty：那是 VBA 内置函数名，用作成员名容易在类内部引起歧义
Public Property Get HasNoAreas() As Boolean
    HasNoAreas = (mCount = 0)
End Property
