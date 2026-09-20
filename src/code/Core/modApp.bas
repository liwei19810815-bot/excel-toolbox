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
' 宿主探测。
'
' 【不能靠 Application.Name 判断】——这是实测踩出来的：
' WPS 表格的 Application.Name 返回的就是字符串 "Microsoft Excel"，
' 它是刻意伪装成 Excel 的（连 COM ProgID 都会接管）。
' 原先"名字里有 Microsoft Excel 就是 Excel"的写法，在 WPS 上永远判成 Excel，
' 整套 SupportedInWps 机制因此从来没生效过。
'
' 可靠的判据是安装路径：WPS 的 Application.Path 指向它自己的安装目录
' （含 WPSOFF… / KINGSO…，注意可能是 8.3 短名）。路径拿不到时再退回看版本号——
' 真 Excel 从 2016 起都是 16.0，而 WPS 报 12.0。
'------------------------------------------------------------------------------
Public Function Host() As HostKind
    Dim appPath As String, appName As String, appVer As String

    On Error Resume Next
    appPath = Application.Path
    appName = Application.Name
    appVer = Application.Version
    On Error GoTo 0

    ' 【必须用短名也能命中的前缀】。实测 WPS 的 Application.Path 返回的是
    ' 8.3 短路径，形如 D:\PROGRA~3\WPSOFF~1\...\office6
    ' 拿 "WPSOFFICE" 去匹配 "WPSOFF~1" 是匹配不上的——
    ' 我第一版就栽在这里，改完自以为修好了，一跑才发现还是认不出来。
    If InStr(1, appPath, "WPSOFF", vbTextCompare) > 0 _
       Or InStr(1, appPath, "KINGSO", vbTextCompare) > 0 Then
        Host = HostWps
        Exit Function
    End If

    ' 名字本身仍有参考价值：某些 WPS 版本不伪装
    If InStr(1, appName, "WPS", vbTextCompare) > 0 Then
        Host = HostWps
        Exit Function
    End If

    If InStr(1, appName, "Microsoft Excel", vbTextCompare) > 0 Then
        ' 伪装成 Excel 但版本号对不上：12.0 是 Excel 2007，而本工具箱
        ' 最低只支持 2010(14.0)，所以真 Excel 不可能报 12.0 或更低
        If Val(appVer) > 0 And Val(appVer) < 14 Then
            Host = HostWps
        Else
            Host = HostExcel
        End If
        Exit Function
    End If

    Host = HostUnknown
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
