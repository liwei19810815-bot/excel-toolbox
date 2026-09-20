VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsActionDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsActionDef - 一个工具（Ribbon 命令）的元数据
'
' 只描述"这个命令是什么、有什么前置条件、能不能撤销"，不含任何执行逻辑。
'==============================================================================
Option Explicit

Public Id As String                  ' 唯一标识，与 Ribbon 控件的 tag 一致，如 "data.dedupe"
Public Label As String               ' 显示名，也用于撤销按钮上的"撤销 XXX"
Public Screentip As String           ' 悬停提示

Public Undoable As Boolean           ' 是否走撤销框架。False 表示不可撤销（跨文件写入、导出等）
Public ConfirmBeforeRun As Boolean   ' 执行前是否强制弹确认（不可撤销的破坏性操作必须为 True）

Public RequiresSelection As Boolean  ' 需要有效的单元格选区
Public RequiresWorkbook As Boolean   ' 需要有活动工作簿
Public SupportedInWps As Boolean     ' WPS 下是否可用；False 则在 WPS 中灰显
