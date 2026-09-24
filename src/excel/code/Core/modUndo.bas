Attribute VB_Name = "modUndo"
'==============================================================================
' modUndo - 统一撤销框架
'
' 为什么必须自建：VBA 宏一旦执行，Excel 原生的 Ctrl+Z 撤销栈就被清空且
' 无法恢复。这不是 bug 而是设计——宏对对象模型的改动 Excel 无从追踪。
' 所以工具箱里每一个写操作都必须自己留后路。
'
' 机制：
'   执行前把将被改动的区域整块拷到一个隐藏的备份工作簿里（一个事务一批快照），
'   撤销时再整块拷回去。用 Range.Copy 而不是手工存 Variant 数组，是因为前者
'   一次就能带走值、公式、数字格式、字体、边框、底纹、批注、合并状态，
'   自己存数组的话这些全都要单独处理，且极易漏项。
'
' 能力边界（必须如实对用户讲，不能假装可撤销）：
'   - 只覆盖【当前 Excel 会话】，关掉 Excel 即失效；
'   - 只覆盖【单元格内容与格式】，不覆盖跨文件写入、文件重命名、导出、打印；
'   - 超过规模阈值（modSettings.MaxUndoCells）的操作【不拍快照、不入栈】，
'     但 Commit 会返回一句警告，由 RunAction 原样转达给用户。
'     执行前无法预知规模，所以只能事后如实相告——但绝不能不说。
'
' 备份工作簿用 IsAddin = True 而不是普通隐藏工作簿：
'   普通隐藏工作簿会在用户关掉自己所有文件后仍把 Excel 进程吊住不退出，
'   而 IsAddin 工作簿不计入"还有文件开着"的判断，行为和加载宏本身一致。
'==============================================================================
Option Explicit
Option Private Module

Private mBackupWb As Workbook
Private mStack As Collection         ' LIFO，末尾是最近一次操作
Private mCurrent As clsUndoTx        ' 当前打开的事务
Private mSeq As Long

'==============================================================================
' 事务
'==============================================================================

'------------------------------------------------------------------------------
' 开启事务。由 modAction.RunAction 调用，业务代码不要直接调。
'------------------------------------------------------------------------------
Public Sub BeginTx(ByVal Label As String)
    If Not mCurrent Is Nothing Then
        ' 上一个事务没收尾就开新的，说明管线有 bug；丢弃旧的并继续，避免越滚越乱
        Set mCurrent = Nothing
    End If

    mSeq = mSeq + 1
    Set mCurrent = New clsUndoTx
    mCurrent.Label = Label
    mCurrent.Seq = mSeq

    ' 记录现场，撤销后把用户送回原位
    On Error Resume Next
    If Not ActiveSheet Is Nothing Then mCurrent.ActiveSheetName = ActiveSheet.Name
    If TypeName(Selection) = "Range" Then mCurrent.SelectionAddress = Selection.Address
    On Error GoTo 0
End Sub

Public Function InTx() As Boolean
    InTx = Not (mCurrent Is Nothing)
End Function

'------------------------------------------------------------------------------
' 快照一个区域。业务代码在【改动之前】调用。
'
' 多区域会被收敛成外接矩形——还原时按矩形整块拷回，比逐区域还原更安全，
' 代价是可能多存一些没改过的单元格。
'------------------------------------------------------------------------------
Public Sub Capture(ByVal rng As Range)
    If mCurrent Is Nothing Then Exit Sub        ' 不可撤销的操作，静默跳过
    If rng Is Nothing Then Exit Sub

    Dim box As Range
    Set box = BoundingBox(rng)
    If box Is Nothing Then Exit Sub

    CaptureBox box, False
End Sub

'------------------------------------------------------------------------------
' 快照整张表。
'
' 【结构性操作必须用这个】：删除/插入行列、排序、拆分合并单元格、逆透视……
' 这类操作执行后行列会整体位移，只快照"将被删的那几行"是没用的——
' 还原时地址早已对不上，会把数据写到错误的位置。
'------------------------------------------------------------------------------
Public Sub CaptureSheet(ByVal ws As Worksheet)
    If mCurrent Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub

    Dim used As Range
    Set used = modRange.RealUsedRange(ws)
    If used Is Nothing Then
        ' 空表也要留一条记录，否则"往空表里写东西"这个操作无法撤销
        Set used = ws.Range("A1")
    End If

    CaptureBox used, True
End Sub

'------------------------------------------------------------------------------
' 实际写快照
'------------------------------------------------------------------------------
Private Sub CaptureBox(ByVal box As Range, ByVal fullSheet As Boolean)
    Dim cellCount As Double
    cellCount = CDbl(box.Rows.Count) * CDbl(box.Columns.Count)

    ' 【先算账，再决定拷不拷】。
    ' 超限的事务反正要丢弃，先老老实实拷完 100 万格再扔掉纯属白费——
    ' 用户会莫名其妙等上十几秒，然后得到一句"无法撤销"。
    If mCurrent.TotalCells + cellCount > modSettings.MaxUndoCells Then
        mCurrent.OverCapacity = True
        mCurrent.SkippedCells = mCurrent.SkippedCells + cellCount
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = box.Worksheet

    Dim snap As clsUndoSnapshot
    Set snap = New clsUndoSnapshot
    snap.TargetWorkbook = ws.Parent.Name
    snap.TargetSheet = ws.Name
    snap.TargetAddress = box.Address(False, False)
    snap.IsFullSheet = fullSheet
    snap.CellCount = CDbl(box.Rows.Count) * CDbl(box.Columns.Count)

    Dim dest As Worksheet
    Set dest = NewSnapshotSheet(mCurrent.Seq, mCurrent.Count + 1)
    snap.SnapshotSheet = dest.Name

    ' 拷到快照表里【同样的地址】，还原时直接对位，不用做偏移换算
    box.Copy Destination:=dest.Range(snap.TargetAddress)

    ' 行高列宽 Range.Copy 带不走，单独记
    If fullSheet Then
        snap.ColumnWidths = CaptureColumnWidths(box)
        snap.RowHeights = CaptureRowHeights(box)
    End If

    ' 快照用的是 Range.Copy，它会让 Excel 停在复制模式（选区跑马灯）。
    ' 不清掉的话，后续某些 API（例如 SparklineGroups.Add）会直接抛 1004，
    ' 而且报错内容和复制模式毫无关系，极难联想到这里。
    Application.CutCopyMode = False

    mCurrent.AddSnapshot snap
    MarkBackupSaved
End Sub

Private Function CaptureColumnWidths(ByVal box As Range) As Variant
    Dim n As Long: n = box.Columns.Count
    Dim arr() As Double
    ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = box.Columns(i).ColumnWidth
    Next i
    CaptureColumnWidths = arr
End Function

Private Function CaptureRowHeights(ByVal box As Range) As Variant
    Dim n As Long: n = box.Rows.Count
    Dim arr() As Double
    ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = box.Rows(i).RowHeight
    Next i
    CaptureRowHeights = arr
End Function

'------------------------------------------------------------------------------
' 提交事务：推入栈顶。
'
' 返回一句需要转达给用户的警告；没有警告就返回空串。
'
' 规模超限时事务会被丢弃（留着既占内存、还原也慢），但【必须告诉用户】：
' 默默丢掉会让用户以为这一步能撤销，等真去点撤销才发现撤不了，
' 那时候早就没法补救了。这正是"不做假装可撤销"的底线。
'------------------------------------------------------------------------------
Public Function Commit() As String
    If mCurrent Is Nothing Then Exit Function

    ' 这个判断必须排在 Count = 0 之前：超限时压根没拍快照，Count 就是 0，
    ' 顺序反了就会被当成"没改动"而静默放过，警告也就发不出去了。
    If mCurrent.OverCapacity Then
        Dim cells As Double
        cells = mCurrent.TotalCells + mCurrent.SkippedCells
        DiscardCurrent
        Commit = "注意：本次改动涉及约 " & Format$(cells, "#,##0") & " 个单元格，" & _
                 "超过撤销上限（" & Format$(modSettings.MaxUndoCells, "#,##0") & "），" & _
                 "已【不保留】撤销记录，无法撤销。"
        Exit Function
    End If

    If mCurrent.Count = 0 Then
        ' 事务开了但没留任何快照：业务代码没改东西，或本来就没什么可撤销的
        DiscardCurrent
        Exit Function
    End If

    EnsureStack
    mStack.Add mCurrent
    Set mCurrent = Nothing

    TrimStack
    modRibbon.RefreshControl "btnUndoLast"
End Function

'------------------------------------------------------------------------------
' 回滚：业务代码执行到一半出错时立即还原，不入栈。
'
' 返回是否真的还原成功。调用方【必须】看这个返回值：
' 回滚本身也可能失败（目标表被关掉、被保护、快照丢了），
' 这时候还对用户说"改动已回滚"就是撒谎，而且会让他放弃检查数据。
'------------------------------------------------------------------------------
Public Function Rollback() As Boolean
    If mCurrent Is Nothing Then
        Rollback = True                  ' 没有待回滚的事务，视为成功
        Exit Function
    End If

    Dim tx As clsUndoTx
    Set tx = mCurrent
    Set mCurrent = Nothing

    On Error Resume Next
    Err.Clear
    RestoreTx tx
    Rollback = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    DropSnapshotSheets tx
End Function

Private Sub DiscardCurrent()
    Dim tx As clsUndoTx
    Set tx = mCurrent
    Set mCurrent = Nothing
    If Not tx Is Nothing Then DropSnapshotSheets tx
End Sub

'==============================================================================
' 撤销
'==============================================================================

Public Function CanUndo() As Boolean
    EnsureStack
    CanUndo = (mStack.Count > 0)
End Function

' 撤销栈里还有几步。自检串里带上它，测试才能断言
' "这一步到底有没有进栈"，而不是只能看有没有记录。
Public Function Depth() As Long
    EnsureStack
    Depth = mStack.Count
End Function

' 待撤销操作的名称，供 Ribbon 按钮显示
Public Function PeekLabel() As String
    EnsureStack
    If mStack.Count = 0 Then Exit Function
    PeekLabel = mStack(mStack.Count).Label
End Function

Public Sub UndoLast()
    EnsureStack
    If mStack.Count = 0 Then Exit Sub

    Dim tx As clsUndoTx
    Set tx = mStack(mStack.Count)

    ' 【先还原，成功了才出栈】。
    ' 反过来写的话，一旦还原失败（目标表被重命名、删除、关闭或保护），
    ' 这一步的撤销记录就永久丢了，用户连重试的机会都没有——
    ' 而"表被保护"这类问题恰恰是解除保护后重试就能成功的。
    RestoreTx tx

    mStack.Remove mStack.Count
    DropSnapshotSheets tx

    modRibbon.RefreshControl "btnUndoLast"
End Sub

'------------------------------------------------------------------------------
' 回放一个事务。快照按【逆序】还原：同一事务里后写的快照先还原，
' 顺序敏感的多表操作才不会互相覆盖。
'------------------------------------------------------------------------------
Private Sub RestoreTx(ByVal tx As clsUndoTx)
    Dim i As Long
    For i = tx.Snapshots.Count To 1 Step -1
        RestoreSnapshot tx.Snapshots(i)
    Next i

    ' 送回原来的位置
    On Error Resume Next
    If Len(tx.ActiveSheetName) > 0 Then
        ActiveWorkbook.Worksheets(tx.ActiveSheetName).Activate
        If Len(tx.SelectionAddress) > 0 Then ActiveSheet.Range(tx.SelectionAddress).Select
    End If
    On Error GoTo 0
End Sub

Private Sub RestoreSnapshot(ByVal snap As clsUndoSnapshot)
    Dim ws As Worksheet
    Set ws = ResolveSheet(snap)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 210, "modUndo.RestoreSnapshot", _
                  "撤销失败：找不到工作表 [" & snap.TargetWorkbook & "]" & snap.TargetSheet & _
                  "，它可能已被重命名、删除或关闭。"
    End If

    Dim src As Worksheet
    Set src = SnapshotSheet(snap.SnapshotSheet)
    If src Is Nothing Then
        Err.Raise vbObjectError + 211, "modUndo.RestoreSnapshot", "撤销失败：快照数据已丢失。"
    End If

    If snap.IsFullSheet Then
        ' 整表还原必须先彻底清干净：操作可能往快照范围之外写了东西，
        ' 也可能新建了合并单元格，不清就会有残留。
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If

    src.Range(snap.TargetAddress).Copy Destination:=ws.Range(snap.TargetAddress)

    If snap.IsFullSheet Then
        RestoreColumnWidths ws, snap
        RestoreRowHeights ws, snap
    End If

    Application.CutCopyMode = False
End Sub

Private Sub RestoreColumnWidths(ByVal ws As Worksheet, ByVal snap As clsUndoSnapshot)
    If IsEmpty(snap.ColumnWidths) Then Exit Sub
    Dim box As Range: Set box = ws.Range(snap.TargetAddress)
    Dim arr As Variant: arr = snap.ColumnWidths
    Dim i As Long
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        box.Columns(i).ColumnWidth = arr(i)
    Next i
    On Error GoTo 0
End Sub

Private Sub RestoreRowHeights(ByVal ws As Worksheet, ByVal snap As clsUndoSnapshot)
    If IsEmpty(snap.RowHeights) Then Exit Sub
    Dim box As Range: Set box = ws.Range(snap.TargetAddress)
    Dim arr As Variant: arr = snap.RowHeights
    Dim i As Long
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        box.Rows(i).RowHeight = arr(i)
    Next i
    On Error GoTo 0
End Sub

Private Function ResolveSheet(ByVal snap As clsUndoSnapshot) As Worksheet
    On Error Resume Next
    Set ResolveSheet = Application.Workbooks(snap.TargetWorkbook).Worksheets(snap.TargetSheet)
    On Error GoTo 0
End Function

'==============================================================================
' 备份工作簿
'==============================================================================

Private Function BackupWorkbook() As Workbook
    If Not mBackupWb Is Nothing Then
        ' 用户可能手工关掉了它；访问任意属性来探活
        On Error Resume Next
        Dim probe As String
        probe = mBackupWb.Name
        If Err.Number <> 0 Then Set mBackupWb = Nothing
        On Error GoTo 0
    End If

    If mBackupWb Is Nothing Then
        Dim prevSheets As Long
        prevSheets = Application.SheetsInNewWorkbook
        Application.SheetsInNewWorkbook = 1
        Set mBackupWb = Application.Workbooks.Add
        Application.SheetsInNewWorkbook = prevSheets

        ' 关键：IsAddin 的工作簿不会把 Excel 进程吊住，也不会出现在窗口列表里
        mBackupWb.IsAddin = True
        mBackupWb.Saved = True
    End If

    Set BackupWorkbook = mBackupWb
End Function

Private Function NewSnapshotSheet(ByVal seq As Long, ByVal idx As Long) As Worksheet
    Dim wb As Workbook
    Set wb = BackupWorkbook()

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "s" & seq & "_" & idx
    Set NewSnapshotSheet = ws
End Function

' 参数不叫 name：Name 是 VBA 的语句关键字（Name x As y），当标识符用会报语法错误
Private Function SnapshotSheet(ByVal sheetName As String) As Worksheet
    If mBackupWb Is Nothing Then Exit Function
    On Error Resume Next
    Set SnapshotSheet = mBackupWb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub DropSnapshotSheets(ByVal tx As clsUndoTx)
    If mBackupWb Is Nothing Then Exit Sub

    Dim snap As clsUndoSnapshot
    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    For Each snap In tx.Snapshots
        ' 备份簿至少要留一张表，否则删到最后一张会报错
        If mBackupWb.Worksheets.Count > 1 Then
            mBackupWb.Worksheets(snap.SnapshotSheet).Delete
        Else
            mBackupWb.Worksheets(snap.SnapshotSheet).Cells.Clear
        End If
    Next snap
    On Error GoTo 0

    Application.DisplayAlerts = prevAlerts
    MarkBackupSaved
End Sub

Private Sub MarkBackupSaved()
    On Error Resume Next
    If Not mBackupWb Is Nothing Then mBackupWb.Saved = True
    On Error GoTo 0
End Sub

'==============================================================================
' 栈维护与清理
'==============================================================================

Private Sub EnsureStack()
    If mStack Is Nothing Then Set mStack = New Collection
End Sub

Private Sub TrimStack()
    Dim maxN As Long
    maxN = modSettings.MaxUndoSteps

    Do While mStack.Count > maxN
        DropSnapshotSheets mStack(1)       ' 丢最旧的
        mStack.Remove 1
    Loop
End Sub

' 清空撤销栈并关闭备份工作簿。由 App_Shutdown 调用。
Public Sub Cleanup()
    Set mCurrent = Nothing

    If Not mStack Is Nothing Then
        Do While mStack.Count > 0
            mStack.Remove 1
        Loop
    End If

    On Error Resume Next
    If Not mBackupWb Is Nothing Then
        mBackupWb.Saved = True
        mBackupWb.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Set mBackupWb = Nothing
End Sub

'==============================================================================
' 工具
'==============================================================================

'------------------------------------------------------------------------------
' 多区域 -> 外接矩形。快照按矩形存，还原时整块拷回，
' 比逐区域还原少一大类边界 bug。
'------------------------------------------------------------------------------
Private Function BoundingBox(ByVal rng As Range) As Range
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count = 1 Then
        Set BoundingBox = rng
        Exit Function
    End If

    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    Dim a As Range, first As Boolean
    first = True

    For Each a In rng.Areas
        If first Then
            r1 = a.Row: c1 = a.Column
            r2 = a.Row + a.Rows.Count - 1
            c2 = a.Column + a.Columns.Count - 1
            first = False
        Else
            If a.Row < r1 Then r1 = a.Row
            If a.Column < c1 Then c1 = a.Column
            If a.Row + a.Rows.Count - 1 > r2 Then r2 = a.Row + a.Rows.Count - 1
            If a.Column + a.Columns.Count - 1 > c2 Then c2 = a.Column + a.Columns.Count - 1
        End If
    Next a

    Set BoundingBox = rng.Worksheet.Range(rng.Worksheet.Cells(r1, c1), rng.Worksheet.Cells(r2, c2))
End Function
