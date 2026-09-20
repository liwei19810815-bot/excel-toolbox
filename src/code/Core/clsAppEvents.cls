VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' clsAppEvents - Application 级事件挂钩
'
' 加载宏里想响应"用户换了个单元格"这类事件，只能走 WithEvents 挂 Application，
' 因为加载宏本身的 ThisWorkbook 事件只覆盖它自己那个（空的）工作簿。
'
' 事件处理里必须极度克制：SheetSelectionChange 是用户每动一下光标就触发一次的，
' 任何耗时操作都会让整个 Excel 变卡。
'==============================================================================
Option Explicit

Private WithEvents mApp As Application
Attribute mApp.VB_VarHelpID = -1

Public Sub Hook()
    Set mApp = Application
End Sub

Public Sub Unhook()
    Set mApp = Nothing
End Sub

Private Sub mApp_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    ' 事件处理里一旦抛错，Excel 会静默关掉事件响应，后面所有事件都失效且没有提示。
    ' 所以这里必须全程 On Error Resume Next。
    On Error Resume Next
    modSpotlight.OnSelectionChanged target
End Sub
