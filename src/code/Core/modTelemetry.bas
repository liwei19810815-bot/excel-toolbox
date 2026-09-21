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
'
' 【用户最多会等多久：把账算清楚】
'   HTTP 是同步调用，所以必须把最坏情况写明白，不能含糊过去。
'
'     单条命令执行        0 秒。单条只写本地文件，完全不碰网络。
'     启动（端点正常）    N × 往返；内网往返几十毫秒，N ≤ 5，可忽略。
'     启动（端点挂掉）    1.5 秒 —— 第一个文件发失败就 Exit For，不逐个重试。
'     关闭（端点挂掉）    0 秒 —— 启动时已知端点死了，关闭时直接跳过。
'
'   最后一条是专门为"关 Excel"做的：那是用户最不耐烦、
'   也最容易把卡顿归咎于插件的时刻。
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
'
' 【1.5 秒是按"用户能不能察觉"定的，不是按"网络够不够快"定的】。
' 内网端点正常时往返是几十毫秒，1.5 秒绰绰有余；端点挂了的时候，
' 这 1.5 秒会直接加在用户关闭 Excel 的等待上——那是他最不耐烦的时刻。
Private Const HTTP_TIMEOUT_MS As Long = 1500

' 一次会话最多发几个文件。积压多天时不能让启动变慢。
Private Const MAX_FILES_PER_FLUSH As Long = 5

Private mFlushedThisSession As Boolean

' 【本次会话已经确认端点不可达】。
'
' 启动时试过一次发不出去，就不必在关闭时再试一次——网络不会因为用户
' 点了关闭按钮就恢复。这一条专门用来消除"关 Excel 时卡一下"：
' 端点挂掉的环境里，那是用户唯一会感知到遥测存在的时刻。
Private mEndpointDeadThisSession As Boolean

' 每追加多少条检查一次缓冲大小。每条都查太浪费（要遍历目录），
' 完全不查又会让"上限 5MB"这个承诺在长会话里失效。
Private Const PRUNE_EVERY_N_APPENDS As Long = 200
Private mAppendsSincePrune As Long

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

    ' 【启动时就发不出去的话，关闭时不要再试】。
    ' 网络不会因为用户点了关闭按钮就恢复，再试一次只是让 Excel 多卡 1.5 秒——
    ' 而那正是用户最不耐烦、也最容易把锅扣到插件头上的时刻。
    ' 数据不会丢：缓冲还在，下次启动再发。
    If mEndpointDeadThisSession Then Exit Sub

    FlushBuffer
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 无视"本会话端点已死"的标记，强制冲一次。
' 只给测试用——测试需要在同一个会话里模拟"端点恢复"。
'------------------------------------------------------------------------------
Public Sub ForceFlush()
    On Error Resume Next
    If Not IsEnabled() Then Exit Sub
    mEndpointDeadThisSession = False
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

    ' 错误描述可能很长、可能带换行，截断并清洗——它只是给 IT 聚类用的线索。
    '
    ' 【必须先抹掉路径和文件名再截断】。Excel 的错误描述里经常带着完整路径，
    ' 比如"'D:\财务\2026年薪资.xlsx' 无法访问"——而我们在使用说明里
    ' 白纸黑字承诺了"不采集文件名、文件路径"。
    ' 光靠"我们没主动去读路径"是不够的：错误描述是 Excel 给的，里面有什么
    ' 不由我们决定。承诺了就要在代码里堵死。
    parts(9) = CStr(errNum) & " " & Clean(Left$(ScrubPaths(errDesc), 200))

    BuildRecord = Join(parts, FIELD_SEP)
End Function

'------------------------------------------------------------------------------
' 把错误描述里的路径与文件名替换成占位符。
'
' 判据（宁可多抹，不可漏抹——错杀一个技术词不影响聚类，漏掉一个路径就是违诺）：
'   - 含反斜杠或正斜杠的词        → 路径
'   - 形如 xxx.ext 的词           → 文件名
'   - 单双引号、书名号括起来的段  → Excel 报错时路径几乎总是被引号括着
'
' 保留的是错误号和句子结构，足够 IT 按错误类型聚类。
'------------------------------------------------------------------------------
Public Function ScrubPaths(ByVal src As String) As String
    On Error GoTo Failed
    If Len(src) = 0 Then Exit Function

    ' 先把引号/书名号里的内容整段换掉——Excel 报路径时基本都带引号
    Dim s As String
    s = ReplaceQuoted(src, "'", "'")
    s = ReplaceQuoted(s, """", """")
    s = ReplaceQuoted(s, ChrW$(&H300C), ChrW$(&H300D))   ' 「」

    ' 再逐词扫，抓漏网的裸路径和裸文件名
    Dim words() As String, i As Long
    words = Split(s, " ")
    For i = LBound(words) To UBound(words)
        If LooksLikePath(words(i)) Then words(i) = "<path>"
    Next i

    ScrubPaths = Join(words, " ")
    Exit Function

Failed:
    ' 清洗本身出错时【不要把原文放出去】——那正是要防的东西
    Err.Clear
    ScrubPaths = "<scrub-failed>"
End Function

Private Function ReplaceQuoted(ByVal src As String, _
                               ByVal openCh As String, _
                               ByVal closeCh As String) As String
    Dim result As String, rest As String
    Dim p1 As Long, p2 As Long

    rest = src
    Do
        p1 = InStr(rest, openCh)
        If p1 = 0 Then Exit Do
        p2 = InStr(p1 + Len(openCh), rest, closeCh)
        If p2 = 0 Then Exit Do

        result = result & Left$(rest, p1 - 1) & "<path>"
        rest = Mid$(rest, p2 + Len(closeCh))
    Loop

    ReplaceQuoted = result & rest
End Function

Private Function LooksLikePath(ByVal w As String) As Boolean
    If Len(w) = 0 Then Exit Function

    If InStr(w, "\") > 0 Then LooksLikePath = True: Exit Function
    If InStr(w, "/") > 0 Then LooksLikePath = True: Exit Function

    ' 形如 name.ext：点号后面跟 1-5 个字母，且点号不在首尾
    Dim dotPos As Long, extPart As String
    dotPos = InStrRev(w, ".")
    If dotPos > 1 And dotPos < Len(w) Then
        extPart = Mid$(w, dotPos + 1)
        If Len(extPart) >= 1 And Len(extPart) <= 5 Then
            If IsAllLetters(extPart) Then LooksLikePath = True
        End If
    End If
End Function

Private Function IsAllLetters(ByVal txt As String) As Boolean
    Dim i As Long, ch As String
    If Len(txt) = 0 Then Exit Function
    For i = 1 To Len(txt)
        ch = LCase$(Mid$(txt, i, 1))
        If ch < "a" Or ch > "z" Then Exit Function
    Next i
    IsAllLetters = True
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

    ' 【追加时也要清理，不能只在启动时清】。
    ' 只在启动清的话，一次会话里疯狂点命令就能把缓冲顶到远超上限，
    ' 而"上限 5MB"是写给用户看的承诺。
    ' 但每条都去遍历目录太浪费，所以按次数节流。
    mAppendsSincePrune = mAppendsSincePrune + 1
    If mAppendsSincePrune >= PRUNE_EVERY_N_APPENDS Then
        mAppendsSincePrune = 0
        PruneBuffer
    End If

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

    ' 第二遍：仍然超量，从最旧的开始删到达标为止。
    Dim names As Collection
    Set names = ListBufferFiles(fso)
    If names Is Nothing Then Exit Sub

    Dim i As Long, p As String
    For i = 1 To names.Count
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

'------------------------------------------------------------------------------
' 缓冲目录里的日志文件名，按【从旧到新】排好序。
'
' 【不用定长数组】：第一版写的是 ReDim names(0 To 512) + 超出就 Exit For，
' 超过 513 个文件的部分会被静默跳过——清理清不干净、补传也补不到，
' 而且这种"静默截断"恰恰是最难发现的一类 bug。Collection 没有上限。
'
' 只收自己的 yyyymmdd.log，目录里万一有别的东西不关我们的事。
' 文件名即日期，按名字排序就是按时间排序，不需要读 DateLastModified
' （那个会被文件复制、备份软件改掉，反而不如文件名可靠）。
'------------------------------------------------------------------------------
Private Function ListBufferFiles(ByVal fso As Object) As Collection
    On Error GoTo Failed

    Dim result As New Collection
    Dim f As Object, nm As String

    For Each f In fso.GetFolder(BufferDir()).Files
        nm = f.Name
        If Len(nm) = 12 Then
            If LCase$(Right$(nm, 4)) = ".log" And IsAllDigits(Left$(nm, 8)) Then
                InsertSorted result, nm
            End If
        End If
    Next f

    Set ListBufferFiles = result
    Exit Function

Failed:
    Err.Clear
    Set ListBufferFiles = New Collection
End Function

' 插入排序。文件数是天数量级（配合 30 天上限最多几十个），
' 用不着更复杂的算法。
Private Sub InsertSorted(ByVal c As Collection, ByVal value As String)
    Dim i As Long
    For i = 1 To c.Count
        If value < c(i) Then
            c.Add value, , i
            Exit Sub
        End If
    Next i
    c.Add value
End Sub

Private Function IsAllDigits(ByVal txt As String) As Boolean
    Dim i As Long, ch As String
    If Len(txt) = 0 Then Exit Function
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

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

    ' 【按文件名排序后再发，保证先发最旧的】。
    ' 直接遍历 Files 集合的顺序是文件系统给的，不保证按时间——
    ' 那会导致积压时每次都优先发某几个文件，最旧的永远轮不到、
    ' 最后被 PruneBuffer 当过期数据删掉。文件名是 yyyymmdd.log，
    ' 按名字排序就是按时间排序。
    Dim names As Collection
    Set names = ListBufferFiles(fso)
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub

    Dim i As Long, path As String, body As String, sent As Long
    For i = 1 To names.Count
        If sent >= MAX_FILES_PER_FLUSH Then Exit For

        path = BufferDir() & "\" & names(i)
        If Not fso.FileExists(path) Then GoTo NextFile

        body = ReadAll(fso, path)
        If Len(body) = 0 Then
            ' 空文件没有价值，直接删掉，免得每次都来一遍
            Err.Clear
            fso.DeleteFile path, True
            Err.Clear
            GoTo NextFile
        End If

        If PostBody(body) Then
            Err.Clear
            fso.DeleteFile path, True
            Err.Clear
            sent = sent + 1
        Else
            ' 发不出去就整体放弃这一轮：多半是端点不可达，
            ' 继续试后面的文件只是白等超时。
            ' 记下来，关闭时就不必再白等一次了。
            mEndpointDeadThisSession = True
            Exit For
        End If
NextFile:
    Next i

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
