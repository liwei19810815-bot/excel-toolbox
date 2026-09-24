Attribute VB_Name = "modCaps"
'==============================================================================
' modCaps - 宿主能力探测
'
' 为什么要有这个模块：
'
' 工具箱声称支持 Excel 2010+ 和 WPS，但"声称"和"验证过"是两回事。
' 原来的做法是在 clsActionDef 上挂一个 SupportedInWps 标志，靠人工填——
' 结果 59 个命令一个都没填，那个机制等于不存在。
'
' 人工维护一张"哪个宿主支持哪个 API"的表是不现实的：宿主版本在变，
' WPS 每个版本的 API 覆盖度也在变，而开发机上根本没有那些环境可验证。
'
' 所以改成【运行时探测】：直接问宿主"你有没有这个东西"，问不到就禁用相应命令。
' 这样不需要预先知道跑在什么上面，新宿主也不必改代码。
'
' 探测必须是【只读】的。为了知道"能不能加数据条"就真去加一条再删掉，
' 会弄脏用户的工作表、破坏撤销栈，代价远大于收益。
' 所以只探那些能只读判断的能力；其余的交给 RunAction 的统一错误处理，
' 真失败时至少有一条清楚的报错，而不是工具箱整个崩掉。
'==============================================================================
Option Explicit
Option Private Module

' 缓存状态：0 = 还没探出结论（下次再探），1 = 支持，-1 = 确定不支持。
'
' 【0 和 -1 的区别是这个模块最容易出错的地方】：
' "当前上下文不具备探测条件"（没有活动工作簿、活动的是图表工作表）
' 【不等于】"宿主不支持"。把前者缓存成 -1，后果是用户只要在加载宏启动的那一瞬间
' 恰好停在一张图表工作表上，迷你图按钮就会永久灰掉——之后切回普通工作表也不恢复，
' 除非重启 Excel。而且没有任何提示，用户只会觉得"这个功能坏了"。
'
' 所以：只有【确实调到了 API 并且失败】才记 -1；条件不具备一律留 0 等下次。
Private mSparkline As Long
Private mFileDialog As Long
Private mAutomationSecurity As Long

'------------------------------------------------------------------------------
' 清空缓存，下次访问重新探测。
'
' 在 Ribbon_OnLoad 里调用：功能区重建意味着加载宏刚装载或 VBA 工程被重置过，
' 此时上一轮缓存的结论未必还成立。
'------------------------------------------------------------------------------
Public Sub Reset()
    mSparkline = 0
    mFileDialog = 0
    mAutomationSecurity = 0
End Sub

'------------------------------------------------------------------------------
' 迷你图。WPS 至今没有 SparklineGroups 这套 API。
'------------------------------------------------------------------------------
Public Function SupportsSparklines() As Boolean
    If mSparkline = 0 Then mSparkline = ProbeSparklines()
    SupportsSparklines = (mSparkline = 1)
End Function

Private Function ProbeSparklines() As Long
    ' 必须拿一张【真正的工作表】来探。
    ' 不能用 ActiveSheet：活动的可能是图表工作表（Chart），它根本没有 .Range，
    ' 访问就抛错——那是上下文问题，不是宿主不支持迷你图。
    Dim ws As Object
    Set ws = AnyWorksheet()
    If ws Is Nothing Then
        ProbeSparklines = 0          ' 没有可用的工作表，这次不下结论
        Exit Function
    End If

    On Error GoTo NotSupported
    ' 只读访问：取一下集合的 Count，不创建任何东西
    Dim n As Long
    n = ws.Range("A1").SparklineGroups.Count
    ProbeSparklines = 1
    Exit Function

NotSupported:
    ProbeSparklines = -1
End Function

'------------------------------------------------------------------------------
' 找一张可用于只读探测的工作表。找不到返回 Nothing（= 这次别下结论）。
'------------------------------------------------------------------------------
Private Function AnyWorksheet() As Object
    On Error Resume Next

    If Not ActiveSheet Is Nothing Then
        If TypeName(ActiveSheet) = "Worksheet" Then
            Set AnyWorksheet = ActiveSheet
            Exit Function
        End If
    End If

    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook.Worksheets.Count > 0 Then
            Set AnyWorksheet = ActiveWorkbook.Worksheets(1)
        End If
    End If

    Err.Clear
End Function

'------------------------------------------------------------------------------
' 文件夹选择对话框。合并文件夹、文件批处理都靠它。
'------------------------------------------------------------------------------
Public Function SupportsFileDialog() As Boolean
    If mFileDialog = 0 Then mFileDialog = ProbeFileDialog()
    SupportsFileDialog = (mFileDialog = 1)
End Function

Private Function ProbeFileDialog() As Long
    On Error GoTo NotSupported

    ' 只取对象，不 Show——Show 会真的弹窗。
    ' 常量走 modIO 的本地副本，不能写 msoFileDialogFolderPicker（MSO 库编译期依赖）
    Dim dlg As Object
    Set dlg = Application.FileDialog(modIO.MSO_FILEDIALOG_FOLDERPICKER)
    If dlg Is Nothing Then GoTo NotSupported

    ProbeFileDialog = 1
    Exit Function

NotSupported:
    ProbeFileDialog = -1
End Function

'------------------------------------------------------------------------------
' AutomationSecurity。用来在打开别人的工作簿时禁用其中的宏。
'
' 【这条不支持是安全问题，不只是功能问题】：拿不到这个开关，
' 合并文件夹时扫到的带宏文件，它的 Workbook_Open 会在我们的进程里执行。
'------------------------------------------------------------------------------
Public Function SupportsAutomationSecurity() As Boolean
    If mAutomationSecurity = 0 Then mAutomationSecurity = ProbeAutomationSecurity()
    SupportsAutomationSecurity = (mAutomationSecurity = 1)
End Function

Private Function ProbeAutomationSecurity() As Long
    On Error GoTo NotSupported

    Dim cur As Long
    cur = Application.AutomationSecurity      ' 只读一下
    ProbeAutomationSecurity = 1
    Exit Function

NotSupported:
    ProbeAutomationSecurity = -1
End Function

'------------------------------------------------------------------------------
' 完整的能力报告。
'
' 拿到 WPS / Office 2024 / 32 位 Excel 上跑一次，就知道那台机器上
' 到底什么能用、什么不能用——比任何静态推算都可靠。
' 用法见 docs\兼容性验证.md。
'------------------------------------------------------------------------------
Public Function Report() As String
    Dim buf As String

    buf = "host=" & Application.Name & _
          "|version=" & Application.Version & _
          "|bitness=" & modApp.HostBitness() & _
          "|isWps=" & CStr(modApp.IsWps())

    buf = buf & "|sparklines=" & CStr(SupportsSparklines())
    buf = buf & "|fileDialog=" & CStr(SupportsFileDialog())
    buf = buf & "|automationSecurity=" & CStr(SupportsAutomationSecurity())

    ' 这几项没有安全的只读探法，只能报告"按版本推算"的结论，
    ' 真实结果要靠在目标机器上执行对应命令来确认
    buf = buf & "|csvFormat=" & IIf(Val(Application.Version) >= 16, "UTF8", "本地编码")

    Report = buf
End Function
