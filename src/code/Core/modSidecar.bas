Attribute VB_Name = "modSidecar"
'==============================================================================
' modSidecar - 拉起 sidecar 伴生进程
'
' sidecar 是本机上的一个小进程，专门干 Office.js 任务窗格【做不到】的事
' （刷新 Power Query、数据模型、调用工作簿里的宏）。
'
' 为什么由 .xlam 来拉：这样它跟着 Excel 生死，
' 不用开机自启、不用常驻、不用托盘图标，用户完全不知道它存在。
' 前提是 Office 能创建子进程——这一条已在实测中确认
' （tools\check-office-child-process.ps1）。
'
' 【这里的每一条错误都必须吞掉】。sidecar 起不来只意味着少几个 AI 能力，
' 而抛错会把整个加载宏的启动打断——那是 60 个命令一个都用不了。
' 这个取舍和 modTelemetry.FlushOnStartup 是一样的。
'==============================================================================
Option Explicit
Option Private Module

' 装在哪：安装器把 exe 和 config.json 一起放到这个目录
Private Const SIDECAR_DIR_NAME As String = "ExcelToolbox\sidecar"
Private Const SIDECAR_EXE_NAME As String = "ExcelToolboxSidecar.exe"
Private Const SIDECAR_CFG_NAME As String = "config.json"

' 单实例的名字。开几个 Excel 只该有一个 sidecar——
' 否则每个都占一个端口，任务窗格探到哪个全看运气。
Private Const SIDECAR_INSTANCE_KEY As String = "ExcelToolboxSidecar"

Private mTried As Boolean

'------------------------------------------------------------------------------
' 确保 sidecar 在跑。由 App_Startup 调用。
'
' 【一次会话只试一次】。试不成就算了，不要每次用到都重试——
' 没装 sidecar 组件的用户占多数，反复 Shell 一个不存在的 exe 毫无意义。
'------------------------------------------------------------------------------
Public Sub Sidecar_EnsureRunning()
    If mTried Then Exit Sub
    mTried = True

    On Error Resume Next

    Dim dirPath As String, exePath As String, cfgPath As String
    dirPath = SidecarDir()
    If Len(dirPath) = 0 Then Exit Sub

    exePath = dirPath & "\" & SIDECAR_EXE_NAME
    cfgPath = dirPath & "\" & SIDECAR_CFG_NAME

    ' 【exe 和配置缺一不可】。没装 sidecar 组件的用户走到这里就是两个都没有，
    ' 直接退出，不报错也不提示——这是正常情况，不是故障。
    If Not FileExists(exePath) Then Exit Sub
    If Not FileExists(cfgPath) Then Exit Sub

    Dim cmd As String
    ' 路径一律加引号：%LOCALAPPDATA% 在中文域账号下可能带空格，
    ' 不加引号会被拆成两个参数，sidecar 收到的配置路径是残的。
    cmd = """" & exePath & """" & _
          " --config """ & cfgPath & """" & _
          " --single-instance " & SIDECAR_INSTANCE_KEY & _
          " --watch-process EXCEL"

    ' 0 = 隐藏窗口。用户不该看见一个黑框闪过。
    Dim pid As Double
    pid = Shell(cmd, 0)

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' sidecar 的安装目录。取不到 %LOCALAPPDATA% 就返回空串。
'------------------------------------------------------------------------------
Public Function SidecarDir() As String
    On Error Resume Next
    Dim base As String
    base = Environ$("LOCALAPPDATA")
    If Len(base) = 0 Then
        SidecarDir = ""
    Else
        SidecarDir = base & "\" & SIDECAR_DIR_NAME
    End If
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 这台机器上装了 sidecar 组件没有。
' 【只看文件在不在，不去连它】——连接要走网络栈，会拖慢启动。
'------------------------------------------------------------------------------
Public Function Sidecar_IsInstalled() As Boolean
    On Error Resume Next
    Dim d As String
    d = SidecarDir()
    If Len(d) = 0 Then
        Sidecar_IsInstalled = False
    Else
        Sidecar_IsInstalled = FileExists(d & "\" & SIDECAR_EXE_NAME) And _
                              FileExists(d & "\" & SIDECAR_CFG_NAME)
    End If
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 【不要用 WScript.Shell / FileSystemObject 去判存在】。
' 本仓库实测过：CreateObject("WScript.Shell") 会被安全策略【静默拦截】，
' 拿不到对象也不抛错（README 第 13 条）。Dir 是内建函数，不受这个影响。
'------------------------------------------------------------------------------
Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(path)) > 0)
    If Err.Number <> 0 Then FileExists = False
    Err.Clear
    On Error GoTo 0
End Function
