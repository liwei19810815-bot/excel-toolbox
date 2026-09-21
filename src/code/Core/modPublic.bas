Attribute VB_Name = "modPublic"
'==============================================================================
' modPublic - 对外暴露的入口
'
' 其余模块一律带 Option Private Module，对外不可见；只有这里的过程能被
' Application.Run 调用。构建脚本、测试脚本、以及用户自己的宏都从这里进。
'
' 本模块【不能】加 Option Private Module。
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 自检。构建后由 build\verify.ps1 调用。
'
' 关键作用是【触发 VBA 编译】：Import 只是把源码塞进工程，语法和引用错误
' 要等到第一次执行时才暴露。只要这个函数能返回，就说明整个工程编译通过了。
'------------------------------------------------------------------------------
Public Function Toolbox_SelfCheck() As String
    Dim parts As String

    parts = "OK|" & APP_ID & "|" & APP_VERSION
    parts = parts & "|host=" & Application.Name
    parts = parts & "|bitness=" & modApp.HostBitness()
    parts = parts & "|actions=" & modAction.ActionCount()
    parts = parts & "|canUndo=" & CStr(modUndo.CanUndo())
    parts = parts & "|undoDepth=" & modUndo.Depth()
    parts = parts & "|ribbon=" & CStr(modRibbon.IsRibbonLoaded())

    Toolbox_SelfCheck = parts
End Function

' 宿主能力报告。拿到 WPS / Office 2024 / 32 位 Excel 上跑一次，
' 就知道那台机器上到底什么能用——比任何静态推算都可靠。
Public Function Toolbox_ProbeHost() As String
    Toolbox_ProbeHost = modCaps.Report()
End Function

' 清空宿主能力缓存。供回归测试构造"先在图表工作表上探测、再切回普通工作表"
' 这种场景——那是 modCaps 最容易误判的路径（图表工作表没有 .Range）。
Public Sub Toolbox_ResetCaps()
    modCaps.Reset
End Sub

'------------------------------------------------------------------------------
' 遥测状态。不含任何用户数据，可以直接截图发给 IT 排障。
'------------------------------------------------------------------------------
Public Function Toolbox_TelemetryStatus() As String
    Toolbox_TelemetryStatus = modTelemetry.Status()
End Function

' 供回归测试配置遥测（端点、开关），免得测试去动真实配置
Public Sub Toolbox_SetTelemetry(ByVal enabled As Boolean, ByVal endpoint As String)
    modSettings.PutSetting "TelemetryEnabled", enabled
    modSettings.PutSetting "TelemetryEndpoint", endpoint
End Sub

' 立刻把缓冲冲出去。测试用来验证"上报成功才删本地文件"。
Public Sub Toolbox_FlushTelemetry()
    modTelemetry.ForceFlush
End Sub

' 路径清洗的直接入口。供测试逐条验证"错误描述里的路径确实被抹掉了"——
' 这是使用说明里对员工的承诺，必须能单独断言，不能只靠端到端碰运气。
Public Function Toolbox_ScrubPaths(ByVal src As String) As String
    Toolbox_ScrubPaths = modTelemetry.ScrubPaths(src)
End Function

' 走一遍关闭时的遥测路径。供测试验证"端点已知挂掉时关闭不再白等一个超时"——
' 那是用户唯一会感知到遥测存在的时刻，必须有断言守着。
Public Sub Toolbox_Shutdown()
    modTelemetry.FlushOnShutdown
End Sub

'------------------------------------------------------------------------------
' 帮助
'------------------------------------------------------------------------------

' 某个命令有没有帮助正文。供 tests\check-help.ps1 断言——
' 帮助漏写不会让任何功能测试变红，必须专门检查。
Public Function Toolbox_HasHelp(ByVal actionId As String) As Boolean
    Toolbox_HasHelp = modHelp.HasEntry(actionId)
End Function

' 「我要做什么」搜索的【只解析不执行】版本，返回命中的 actionId（换行分隔）。
' 测试用它断言匹配逻辑，不会真的动用户数据。
Public Function Toolbox_ResolveHelp(ByVal query As String) As String
    Toolbox_ResolveHelp = modHelp.Resolve(query)
End Function

'------------------------------------------------------------------------------
' 检查一批 imageMso 在【当前这台机器的 Excel】上是否存在。
' 入参用 | 分隔，返回【不存在的那些】，同样用 | 分隔；全部存在则返回空串。
'
' 为什么要放在 VBA 里而不是让 PowerShell 直接调：
'   GetImageMso 返回的是 IPictureDisp。跨进程 COM 把这个对象 marshal 回
'   PowerShell 时会【直接挂死，而且不报错】——加了可见窗口和工作簿也一样。
'   在进程内调用就没这个问题：VBA 只判断有没有抛错，不把图片对象传出去。
'
' 为什么需要这个检查：
'   imageMso 无效时 Excel 只是【静默不画图标】，功能区照常加载、按钮照常能点，
'   check-ribbon.ps1 照样报 ribbon=True。用户看到的是一排没有图标的按钮。
'   而且同为 16.0，Microsoft 365 和 Excel 2021 的图标集并不一样——
'   在 365 开发机上验过不等于在 2021 上没问题。
'------------------------------------------------------------------------------
Public Function Toolbox_CheckImageMso(ByVal idList As String) As String
    Dim ids() As String, i As Long, one As String, buf As String

    ids = Split(idList, "|")
    For i = LBound(ids) To UBound(ids)
        one = Trim$(ids(i))
        If Len(one) > 0 Then
            If Not ImageMsoExists(one) Then
                If Len(buf) > 0 Then buf = buf & "|"
                buf = buf & one
            End If
        End If
    Next i

    Toolbox_CheckImageMso = buf
End Function

' 【必须把常用尺寸都试一遍】。GetImageMso 是按尺寸取位图的，
' 不少 ID（尤其是库/菜单型控件，如 ConditionalFormattingDataBars、
' FunctionsInsertGallery）在 16×16 下取不到，换 32×32 就有。
' 只试一个尺寸会把大量【确实存在】的 Excel 原生图标误判成不存在——
' 第一版就是这么得出"13 个无效"的，差点照着那份假名单去改图标。
'------------------------------------------------------------------------------
' 把一批 imageMso 导出成 BMP 文件，供人肉眼比对。
'
' 为什么需要：挑图标时只能看见 ID 名字，看不见图。"TableEraser 到底长什么样"
' 靠名字猜，结论就是各人猜各人的，争不出结果。导出来看一眼就定了。
'
' 入参 idList 用 | 分隔；返回实际导出成功的个数。
' 文件名就是 ID 本身，放在 folderPath 下。
'------------------------------------------------------------------------------
Public Function Toolbox_DumpImageMso(ByVal idList As String, _
                                     ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    Dim ids() As String, i As Long, one As String
    Dim okCount As Long, failed As String

    ids = Split(idList, "|")
    For i = LBound(ids) To UBound(ids)
        one = Trim$(ids(i))
        If Len(one) > 0 Then
            If DumpOne(one, folderPath & "\" & one & ".bmp") Then
                okCount = okCount + 1
            Else
                If Len(failed) > 0 Then failed = failed & ","
                failed = failed & one
            End If
        End If
    Next i

    Toolbox_DumpImageMso = "ok=" & okCount & "|failed=" & failed
End Function

Private Function DumpOne(ByVal idMso As String, ByVal outPath As String) As Boolean
    Dim pic As Object
    On Error GoTo Failed

    ' 32×32 是功能区大按钮的尺寸，比对时看这个最贴近实际观感
    Set pic = Application.CommandBars.GetImageMso(idMso, 32, 32)
    If pic Is Nothing Then Exit Function

    SavePicture pic, outPath
    DumpOne = True
    Exit Function

Failed:
    DumpOne = False
End Function

Private Function ImageMsoExists(ByVal idMso As String) As Boolean
    If TryGetImage(idMso, 16) Then ImageMsoExists = True: Exit Function
    If TryGetImage(idMso, 32) Then ImageMsoExists = True: Exit Function
    If TryGetImage(idMso, 64) Then ImageMsoExists = True: Exit Function
End Function

Private Function TryGetImage(ByVal idMso As String, ByVal px As Long) As Boolean
    Dim pic As Object
    On Error GoTo NotFound
    Set pic = Application.CommandBars.GetImageMso(idMso, px, px)
    TryGetImage = Not (pic Is Nothing)
    Exit Function
NotFound:
    TryGetImage = False
End Function

' 所有已注册的 actionId，换行分隔。供测试脚本与 customUI14.xml 里的 tag 做一致性比对。
Public Function Toolbox_ListActions() As String
    Toolbox_ListActions = modAction.AllActionIds()
End Function

'------------------------------------------------------------------------------
' 供测试脚本调用：执行一个已注册的命令，返回结果文字。
'------------------------------------------------------------------------------
Public Function Toolbox_Run(ByVal actionId As String) As String
    modAction.RunAction actionId
    Toolbox_Run = modAction.LastMessage()
End Function

' 静默模式：不弹任何对话框。测试脚本必须先打开它，否则第一个 MsgBox 就会卡死。
Public Sub Toolbox_SetSilent(ByVal value As Boolean)
    modAction.SetSilent value
End Sub

' 静默模式下给工具预设参数（替代输入框）。key 见各工具里 AskXxx 的第一个参数。
Public Sub Toolbox_SetParam(ByVal key As String, ByVal value As String)
    modPrompt.SetParam key, value
End Sub

Public Sub Toolbox_ClearParams()
    modPrompt.ClearParams
End Sub

' 最近一次执行的结果文字。静默模式下出错时这里会是 "ERROR: <原因>"。
Public Function Toolbox_LastMessage() As String
    Toolbox_LastMessage = modAction.LastMessage()
End Function

'------------------------------------------------------------------------------
' 撤销上一步。成功返回空串，失败返回 "ERROR: <原因>"。
'
' 【必须自己处理错误】：这是个对外入口，会被 Application.Run 直接调用。
' 未捕获的 VBA 错误在这里不会变成干净的 COM 异常，而是弹出"运行时错误 1004"
' 的调试对话框——无界面运行时那个框谁也看不见，调用方就永久挂住了。
'
' Ribbon 上的撤销按钮走的是 RunAction，那条路本来就有统一错误处理；
' 漏的正是这个绕过管线的入口。
'------------------------------------------------------------------------------
Public Function Toolbox_Undo() As String
    On Error GoTo Failed
    modUndo.UndoLast
    Exit Function
Failed:
    Toolbox_Undo = "ERROR: " & Err.Description & " [" & Err.Number & " @ " & Err.Source & "]"
End Function

' 宏被强行中断后恢复 Excel 环境。也挂在 Ribbon 上，这里额外暴露给用户手工调用。
Public Sub Toolbox_ResetEnvironment()
    modPerf.FastModeReset
End Sub

'------------------------------------------------------------------------------
' 分段诊断。
'
' 无头环境下一旦某段代码挂住（典型是弹出了一个不可见的模态框，或者对失效的
' IRibbonUI 指针做了 Invalidate），调用方只会看到永久无响应，拿不到任何信息。
' 这里把 RunAction 管线拆成可以单独调用的小段，逐段试就能定位到具体是哪一环。
'------------------------------------------------------------------------------
Public Function Toolbox_DiagStage(ByVal stage As Long) As String
    Select Case stage
        Case 1
            modPerf.FastModeOn
            modPerf.FastModeOff
            Toolbox_DiagStage = "1 ok: FastMode 开关"

        Case 2
            modPerf.SetStatus "diag"
            modPerf.ClearStatus
            Toolbox_DiagStage = "2 ok: 状态栏"

        Case 3
            modUndo.BeginTx "diag"
            modUndo.Commit
            Toolbox_DiagStage = "3 ok: 空事务开启与提交"

        Case 4
            modRibbon.RefreshControl "btnUndoLast"
            Toolbox_DiagStage = "4 ok: Ribbon 刷新"

        Case 5
            Dim rng As Range
            Set rng = modRange.NormalizeSelection(Selection)
            If rng Is Nothing Then
                Toolbox_DiagStage = "5 ok: 选区归一化 -> Nothing"
            Else
                Toolbox_DiagStage = "5 ok: 选区归一化 -> " & rng.Address(False, False)
            End If

        Case 6
            modUndo.BeginTx "diag-capture"
            modUndo.CaptureSheet ActiveSheet
            modUndo.Commit
            Toolbox_DiagStage = "6 ok: 整表快照，canUndo=" & CStr(modUndo.CanUndo())

        Case Else
            Toolbox_DiagStage = "未知阶段：" & stage
    End Select
End Function
