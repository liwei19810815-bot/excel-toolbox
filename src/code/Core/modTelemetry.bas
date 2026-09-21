Attribute VB_Name = "modTelemetry"
'==============================================================================
' modTelemetry - 使用情况与错误回传
'
' 目的（面向 IT 运维，用户运行时无感）：
'   1. 问题与 BUG 的记录和聚合，用于定位修复
'   2. 日 / 月 / 年的功能点击次数、使用人数、日活
'
' 挂在 modAction.RunAction 的统一出口上——这是当初做统一执行管线最大的红利：
' 59 个命令一个都不用改。
'
'==============================================================================
' 【本模块的铁律】
'
'   遥测失败绝不能影响功能。
'
'   用户想做的是删一行空行。收集端点挂了、网络断了、代理拦了、磁盘满了，
'   都不允许让那一行删不掉，更不允许弹任何框。
'   所以本模块【所有】对外行为都包在 On Error Resume Next 里，
'   并且没有任何一条路径会 Raise、会 MsgBox、会阻塞。
'
' 【采集什么】
'   actionId、宿主与版本、加载宏版本、耗时毫秒、成功/失败、
'   错误号与错误描述（截断）、用户名、机器名、时间戳。
'
' 【绝不采集什么】
'   文件内容、文件名、路径、单元格数据、区域地址、任何用户数据。
'   这条写进了 install\使用说明.txt，对员工是公开的。
'
' 【为什么不用 WScript.Shell】
'   它是宏病毒的典型载体，企业 AV 和 Windows ASR 规则普遍阻止 Office 创建它，
'   而且失败是静默的——本项目实测过（见 docs\部署与自动更新.md）。
'   这里用 MSXML2.ServerXMLHTTP，它不在那类拦截规则里。
'==============================================================================
Option Explicit
Option Private Module

' 单条记录的字段分隔符。选这几个字符是因为它们不会出现在 actionId、
' 版本号、用户名里，也不需要转义。
Private Const FIELD_SEP As String = "|"

' 本地缓冲上限。超过就丢最旧的——遥测数据的价值随时间衰减很快，
' 而把用户的磁盘塞满是实打实的伤害。
Private Const MAX_BUFFER_BYTES As Long = 5242880      ' 5 MB
Private Const MAX_BUFFER_DAYS As Long = 30

' 网络超时。宁可丢这条记录，也不能让用户对着一个卡住的 Excel 干等。
Private Const HTTP_TIMEOUT_MS As Long = 3000

Private mFlushedThisSession As Boolean

'==============================================================================
' 对外接口
'==============================================================================

'------------------------------------------------------------------------------
' 记录一次命令执行。由 modAction.RunAction 在每个出口调用。
'
' outcome: "ok" / "fail" / "cancel" / "blocked"
' errNum / errDesc: 仅 fail 时有值
'------------------------------------------------------------------------------
Public Sub TrackAction(ByVal actionId As String, _
                       ByVal outcome As String, _
                       ByVal elapsedMs As Long, _
                       Optional ByVal errNum As Long = 0, _
                       Optional ByVal errDesc As String = "")
    On Error Resume Next
    If Not IsEnabled() Then Exit Sub

    Dim line As String
    line = BuildRecord(actionId, outcome, elapsedMs, errNum, errDesc)
    AppendToBuffer line

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 会话开始时调用一次：把上次遗留的缓冲发出去。
'
' 【为什么在启动时发而不是每条都发】：
'   每执行一个命令就发一次 HTTP，会在用户连点几个命令时串起明显的卡顿，
'   而且收集端点要承受 N 倍的请求。攒着批量发，对两边都好。
'   代价是最后一次会话的数据要等下次启动才上报——对"使用情况统计"完全够用。
'------------------------------------------------------------------------------
Public Sub FlushOnStartup()
    On Error Resume Next
    If mFlushedThisSession Then Exit Sub
    mFlushedThisSession = True
    If Not IsEnabled() Then Exit Sub

    PruneBuffer
    FlushBuffer

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 会话结束时再冲一次。失败也无所谓，缓冲还在，下次启动会再试。
'------------------------------------------------------------------------------
Public Sub FlushOnShutdown()
    On Error Resume Next
    If Not IsEnabled() Then Exit Sub
    FlushBuffer
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 供自检与排障使用：当前缓冲里有多少条、端点是什么、开关状态。
' 不含任何用户数据，可以直接给 IT 看。
'------------------------------------------------------------------------------
Public Function Status() As String
    On Error Resume Next
    Status = "enabled=" & CStr(IsEnabled()) & _
             "|endpoint=" & Endpoint() & _
             "|pending=" & PendingCount() & _
             "|bufferDir=" & BufferDir()
    Err.Clear
    On Error GoTo 0
End Function

'==============================================================================
' 配置
'==============================================================================

' 默认关闭：没有配置收集端点的环境（比如个人从 GitHub 下载来用），
' 不应该悄悄往任何地方发东西。IT 部署时通过 modSettings 打开并填端点。
Public Function IsEnabled() As Boolean
    On Error Resume Next
    IsEnabled = modSettings.GetSettingBool("TelemetryEnabled", False) _
                And Len(Endpoint()) > 0
    Err.Clear
    On Error GoTo 0
End Function

Public Function Endpoint() As String
    On Error Resume Next
    Endpoint = modSettings.GetSettingString("TelemetryEndpoint", vbNullString)
    Err.Clear
    On Error GoTo 0
End Function

'==============================================================================
' 记录构造
'==============================================================================

Private Function BuildRecord(ByVal actionId As String, _
                             ByVal outcome As String, _
                             ByVal elapsedMs As Long, _
                             ByVal errNum As Long, _
                             ByVal errDesc As String) As String
    Dim parts(0 To 9) As String

    parts(0) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    parts(1) = Clean(Environ$("USERNAME"))
    parts(2) = Clean(Environ$("COMPUTERNAME"))
    parts(3) = "excel"
    parts(4) = Clean(SafeAppVersion())
    parts(5) = Clean(APP_VERSION)
    parts(6) = Clean(actionId)
    parts(7) = Clean(outcome)
    parts(8) = CStr(elapsedMs)
    ' 错误描述可能很长、可能带换行，截断并清洗——它只是给 IT 聚类用的线索
    parts(9) = CStr(errNum) & " " & Clean(Left$(errDesc, 200))

    BuildRecord = Join(parts, FIELD_SEP)
End Function

Private Function SafeAppVersion() As String
    On Error Resume Next
    SafeAppVersion = Application.Version & "." & Application.Build
    Err.Clear
    On Error GoTo 0
End Function

' 把分隔符和换行从字段值里清掉，避免一条记录被拆成两条。
' 不做转义而是直接替换：这些字段本来就不该含这些字符，
' 保留原样没有价值，而转义会让服务端解析复杂化。
Private Function Clean(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, FIELD_SEP, "/")
    Clean = Trim$(t)
End Function

'==============================================================================
' 本地缓冲
'
' 每天一个文件，方便按天清理与按天补传。
'==============================================================================

Private Function BufferDir() As String
    BufferDir = Environ$("LOCALAPPDATA") & "\ExcelToolbox\telemetry"
End Function

Private Function TodayFile() As String
    TodayFile = BufferDir() & "\" & Format$(Now, "yyyymmdd") & ".log"
End Function

Private Sub AppendToBuffer(ByVal line As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub

    If Not fso.FolderExists(BufferDir()) Then
        EnsureFolder fso, BufferDir()
        If Not fso.FolderExists(BufferDir()) Then Exit Sub
    End If

    Dim ts As Object
    Set ts = fso.OpenTextFile(TodayFile(), 8, True)    ' 8 = ForAppending
    If ts Is Nothing Then Exit Sub
    ts.WriteLine line
    ts.Close

    Err.Clear
    On Error GoTo 0
End Sub

' FileSystemObject 的 CreateFolder 不会自动建多级目录
Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    On Error Resume Next
    If fso.FolderExists(path) Then Exit Sub
    EnsureFolder fso, fso.GetParentFolderName(path)
    fso.CreateFolder path
    Err.Clear
    On Error GoTo 0
End Sub

Private Function PendingCount() As Long
    On Error Resume Next
    Dim fso As Object, f As Object, n As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(BufferDir()) Then Exit Function
    For Each f In fso.GetFolder(BufferDir()).Files
        n = n + 1
    Next f
    PendingCount = n
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 清理过期与超量的缓冲。
'
' 两条判据：超过 MAX_BUFFER_DAYS 天的直接删；总量超 MAX_BUFFER_BYTES 时
' 从最旧的开始删到达标为止。
'------------------------------------------------------------------------------
Private Sub PruneBuffer()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(BufferDir()) Then Exit Sub

    Dim f As Object, total As Double

    ' 第一遍：删过期的，顺便统计剩下的总量
    For Each f In fso.GetFolder(BufferDir()).Files
        If DateDiff("d", f.DateLastModified, Now) > MAX_BUFFER_DAYS Then
            Err.Clear
            fso.DeleteFile f.Path, True
            Err.Clear
        Else
            total = total + f.Size
        End If
    Next f

    If total <= MAX_BUFFER_BYTES Then Exit Sub

    ' 第二遍：仍然超量，从最旧的开始删。
    ' 文件名是 yyyymmdd.log，按名字排序就是按时间排序。
    Dim names() As String, count As Long
    ReDim names(0 To 512)
    For Each f In fso.GetFolder(BufferDir()).Files
        If count > UBound(names) Then Exit For
        names(count) = f.Name
        count = count + 1
    Next f
    If count = 0 Then Exit Sub

    SortStrings names, count

    Dim i As Long, p As String
    For i = 0 To count - 1
        If total <= MAX_BUFFER_BYTES Then Exit For
        p = BufferDir() & "\" & names(i)
        If fso.FileExists(p) Then
            total = total - fso.GetFile(p).Size
            Err.Clear
            fso.DeleteFile p, True
            Err.Clear
        End If
    Next i

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SortStrings(ByRef arr() As String, ByVal count As Long)
    Dim i As Long, j As Long, t As String
    For i = 0 To count - 2
        For j = i + 1 To count - 1
            If arr(j) < arr(i) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
End Sub

'==============================================================================
' 上报
'==============================================================================

'------------------------------------------------------------------------------
' 把缓冲目录里的文件逐个发出去。
'
' 【上报成功才删】——这是关键。发失败就留着，下次启动再试。
' 反过来"先删再发"会在网络抖动时静默丢数据，而丢了没人知道。
'------------------------------------------------------------------------------
Private Sub FlushBuffer()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(BufferDir()) Then Exit Sub

    Dim f As Object, body As String, sent As Long
    For Each f In fso.GetFolder(BufferDir()).Files
        ' 一次会话最多发 10 个文件，避免积压很多天时启动变慢
        If sent >= 10 Then Exit For

        body = ReadAll(fso, f.Path)
        If Len(body) > 0 Then
            If PostBody(body) Then
                Err.Clear
                fso.DeleteFile f.Path, True
                Err.Clear
                sent = sent + 1
            Else
                ' 发不出去就整体放弃这一轮：多半是端点不可达，
                ' 继续试后面的文件只是白等超时
                Exit For
            End If
        Else
            ' 空文件没有价值，直接删掉，免得每次都来一遍
            Err.Clear
            fso.DeleteFile f.Path, True
            Err.Clear
        End If
    Next f

    Err.Clear
    On Error GoTo 0
End Sub

Private Function ReadAll(ByVal fso As Object, ByVal path As String) As String
    On Error Resume Next
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 1)      ' 1 = ForReading
    If ts Is Nothing Then Exit Function
    If Not ts.AtEndOfStream Then ReadAll = ts.ReadAll
    ts.Close
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 发一次 POST。成功返回 True。
'
' 任何异常都吞掉并返回 False——这里是最需要克制的地方：
' 收集端点的可用性和用户能不能用工具箱，必须完全无关。
'------------------------------------------------------------------------------
Private Function PostBody(ByVal body As String) As Boolean
    On Error GoTo Failed

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Exit Function

    ' ServerXMLHTTP 才有 setTimeouts；这是选它而不是 XMLHTTP 的主要原因——
    ' 没有超时控制的 HTTP 调用可能把 Excel 挂死好几十秒。
    http.setTimeouts HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS, HTTP_TIMEOUT_MS

    http.Open "POST", Endpoint(), False
    http.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
    http.setRequestHeader "X-Toolbox-Version", APP_VERSION
    http.send body

    PostBody = (http.Status >= 200 And http.Status < 300)
    Exit Function

Failed:
    Err.Clear
    PostBody = False
End Function
