Attribute VB_Name = "modApp"
'==============================================================================
' modApp - 加载宏全局入口与环境信息
'
' 职责：版本号、宿主探测、启动/卸载时的全局初始化与清理。
' 其他模块一律通过这里获取环境信息，不要各自去读 Application 属性。
'==============================================================================
Option Explicit
Option Private Module

Public Const APP_NAME As String = "Excel 通用工具箱"
Public Const APP_ID As String = "ExcelToolbox"
Public Const APP_VERSION As String = "1.0.0"

' 宿主类型
Public Enum HostKind
    HostExcel = 0
    HostWps = 1
    HostUnknown = 2
End Enum

Private mInitialized As Boolean

'------------------------------------------------------------------------------
' 加载宏装载时调用（由 ThisWorkbook.Workbook_Open 触发）
'------------------------------------------------------------------------------
Public Sub App_Startup()
    If mInitialized Then Exit Sub
    mInitialized = True
End Sub

'------------------------------------------------------------------------------
' 加载宏卸载时调用（由 ThisWorkbook.Workbook_BeforeClose 触发）
' 这里负责释放事件钩子、清理临时文件，后续模块会往里追加。
'------------------------------------------------------------------------------
Public Sub App_Shutdown()
    If Not mInitialized Then Exit Sub
    mInitialized = False

    ' 关掉备份工作簿并清空撤销栈。撤销本来就只在当前会话内有效。
    On Error Resume Next
    modSpotlight.Cleanup      ' 先摘事件钩子并清掉条件格式，别把它残留在用户文件里
    modUndo.Cleanup
    modPerf.FastModeReset
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 宿主探测。WPS 的 Application.Name 返回 "WPS表格" 之类的本地化名称，
' 因此用「不是 Microsoft Excel 就当作非 Excel」的保守判断。
'------------------------------------------------------------------------------
Public Function Host() As HostKind
    Dim n As String
    On Error Resume Next
    n = Application.Name
    On Error GoTo 0

    If InStr(1, n, "Microsoft Excel", vbTextCompare) > 0 Then
        Host = HostExcel
    ElseIf InStr(1, n, "WPS", vbTextCompare) > 0 Or InStr(1, n, "表格") > 0 Then
        Host = HostWps
    Else
        Host = HostUnknown
    End If
End Function

Public Function IsWps() As Boolean
    IsWps = (Host() = HostWps)
End Function

'------------------------------------------------------------------------------
' 位数。64 位 Office 下 Win64 编译常量为 True。
'------------------------------------------------------------------------------
Public Function HostBitness() As String
#If Win64 Then
    HostBitness = "64 位"
#Else
    HostBitness = "32 位"
#End If
End Function

Public Function AboutText() As String
    AboutText = APP_NAME & "  v" & APP_VERSION & vbCrLf & vbCrLf & _
                "宿主：" & Application.Name & " " & Application.Version & "（" & HostBitness() & "）" & vbCrLf & _
                "加载宏：" & ThisWorkbook.Name
End Function
