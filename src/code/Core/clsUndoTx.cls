VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUndoTx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsUndoTx - 一次可撤销的操作（事务）
'
' 一个工具的一次执行 = 一个事务，事务内可以有多个快照
' （例如"按列拆分"会同时改动多张表）。还原时按【逆序】回放，
' 这样后写的快照先还原，顺序敏感的场景才不会互相覆盖。
'==============================================================================
Option Explicit

Public Label As String               ' 显示名，撤销按钮上显示"撤销 <Label>"
Public Seq As Long                   ' 递增序号，用于生成快照表名
Public CreatedAt As Date

' 还原后要恢复的现场
Public ActiveSheetName As String
Public SelectionAddress As String

' 规模超限：快照【没有拍】，这个事务不可撤销。
' 由 Commit 转成一句明确的警告告诉用户，绝不能默默咽下去。
Public OverCapacity As Boolean
Public SkippedCells As Double

Private mSnapshots As Collection

Private Sub Class_Initialize()
    Set mSnapshots = New Collection
    CreatedAt = Now
End Sub

Public Sub AddSnapshot(ByVal snap As clsUndoSnapshot)
    mSnapshots.Add snap
End Sub

Public Property Get Snapshots() As Collection
    Set Snapshots = mSnapshots
End Property

Public Property Get Count() As Long
    Count = mSnapshots.Count
End Property

Public Property Get TotalCells() As Double
    Dim s As clsUndoSnapshot, n As Double
    For Each s In mSnapshots
        n = n + s.CellCount
    Next s
    TotalCells = n
End Property
