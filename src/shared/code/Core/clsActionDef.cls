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
Public Label As String               ' 显示名。这是唯一真相——功能区按钮也走 getLabel 取这里，
                                     ' 不在 XML 里另写一份，否则按钮叫"删除重复"、
                                     ' 撤销按钮却显示"撤销 删除重复值"。
Public Screentip As String           ' 悬停提示。可撤销性由 ActionScreentip 自动追加，这里不要手写。

Public Undoable As Boolean           ' 是否走撤销框架。False 表示不可撤销（跨文件写入、导出等）
Public ConfirmBeforeRun As Boolean   ' 执行前是否强制弹确认（不可撤销的破坏性操作必须为 True）

' 点下去之后还会追问参数。用来在标签末尾自动加"…"——
' Office 的通行约定就是"带省略号 = 还要再问你"，用户据此判断点下去会不会立刻动手。
' 手工往 XML 的 label 里敲"…"会漏：实测漏的恰好是问得最多的几个大按钮
' （合并文件夹连问 4 遍，却没有省略号）。所以改成由这个标志统一生成。
Public PromptsForInput As Boolean

Public RequiresSelection As Boolean  ' 需要有效的单元格选区
Public RequiresWorkbook As Boolean   ' 需要有活动工作簿
Public SupportedInWps As Boolean     ' WPS 下是否可用；False 则在 WPS 中灰显
