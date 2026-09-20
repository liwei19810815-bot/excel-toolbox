VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUndoSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsUndoSnapshot - 单个区域的快照
'
' 只记录"从哪来、存到哪、怎么还原"，不含策略判断。
'
' 目标一律按【名称】而非对象引用记录：工作簿可能被关掉又打开，
' 对象引用会变成失效指针，而名称至少能给出可读的失败原因。
'==============================================================================
Option Explicit

Public TargetWorkbook As String      ' 工作簿名（非全路径，跨会话本来就不支持）
Public TargetSheet As String
Public TargetAddress As String       ' 快照区域地址，快照表内用同一地址存放，还原时直接对位

Public SnapshotSheet As String       ' 备份工作簿里的表名
Public IsFullSheet As Boolean        ' True 表示整表快照，还原时要先清空目标表

Public ColumnWidths As Variant       ' 整表快照时的列宽数组（1 基），Range.Copy 带不回来
Public RowHeights As Variant         ' 整表快照时的行高数组（1 基）

Public CellCount As Double           ' 用于规模统计
