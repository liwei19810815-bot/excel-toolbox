Attribute VB_Name = "modLoader"
'==============================================================================
' modLoader - 瘦加载器（自动更新）
'
' 这个加载宏本身【永不改动】，装一次就不用再管。它只干三件事：
'   1. 读共享盘上的版本清单，看有没有新版；
'   2. 有新版就把载荷复制到本地缓存；
'   3. 打开本地缓存里的载荷。真正的 59 个命令全在载荷里。
'
' 为什么要拆成"加载器 + 载荷"，而不是直接让大家链到共享盘上那一个文件：
'   Excel 加载加载宏时会【占住文件锁】。只要公司里有一个人开着 Excel，
'   你就覆盖不了共享盘上那个文件。几十个人的环境里，这等于永远没有窗口期。
'   载荷用版本号做文件名，发版就变成了"新增一个文件"，锁的问题不存在了。
'
' 为什么要复制到本地再打开，而不是直接打开共享盘上的载荷：
'   直接打开同样会锁住共享盘上的载荷文件，你就删不掉旧版本了；
'   而且共享盘一断，所有人的工具箱当场失效。复制到本地之后，
'   断网、出差、VPN 掉线都还能用上一次缓存的版本。
'
' 安全边界（交付文档里必须写明）：
'   能往共享目录写文件的人 = 能让全公司的 Excel 执行任意 VBA。
'   那个目录的写权限必须收紧到只有维护者。这不是代码能解决的问题。
'==============================================================================
Option Explicit

' 共享盘上的发布目录。部署时改这一行，或者在注册表里覆盖（见 SharePath）。
Private Const DEFAULT_SHARE As String = "\\server\share\ExcelToolbox"

Private Const MANIFEST_NAME As String = "manifest.txt"
Private Const PAYLOAD_PREFIX As String = "ExcelToolbox_"
Private Const PAYLOAD_EXT As String = ".xlam"

' 载荷自检串里用来确认"这确实是本工具箱"的标识。
' 必须和载荷侧 modApp.APP_ID 保持一致——加载器和载荷是两个独立工程，
' 这里引用不到那边的常量，只能各存一份。
Private Const PAYLOAD_APP_ID As String = "ExcelToolbox"

Private mPayloadWb As Workbook

' 最近一次更新的结果。内网支持同事时，"没更新成功"必须能说清是哪一种：
' 断网、权限不足、被杀毒锁住、还是载荷本身坏了。全都静默的话只能靠猜。
Private mLastUpdateNote As String

#If VBA7 Then
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

'==============================================================================
' 入口
'==============================================================================

'------------------------------------------------------------------------------
' 由 ThisWorkbook.Workbook_Open 调用。
'
' 整个过程【绝不打断用户】：更新检查失败、共享盘连不上、清单格式不对——
' 任何一步出问题都退回"用本地已有的版本"，而不是弹框。
' 开 Excel 时被一个更新失败的对话框拦住，比晚一天拿到新版糟糕得多。
'------------------------------------------------------------------------------
Public Sub Startup()
    On Error Resume Next

    Dim localPath As String
    localPath = EnsureLatestPayload()

    If Len(localPath) = 0 Then
        ' 一个可用版本都没有：这是唯一值得打扰用户的情况
        MsgBox "Excel 工具箱无法启动：本地没有缓存版本，也连不上发布目录。" & vbCrLf & vbCrLf & _
               "发布目录：" & SharePath() & vbCrLf & vbCrLf & _
               "请检查网络或联系维护者。", vbExclamation, "Excel 工具箱"
        Exit Sub
    End If

    OpenPayload localPath

    ' 【顺序不能反】：必须确认新版本真的跑起来了，才能删旧版本。
    ' 在那之前，旧版本是唯一的退路。
    If VerifyPayloadRuns() Then PurgeOldPayloads localPath

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 确认刚打开的载荷【真的能运行】，而不只是"文件打开成功"。
'
' 光看 Workbooks.Open 有没有报错是不够的：一个 VBA 工程编译不过的 .xlam
' 照样能被打开，只是里面一行代码都跑不了。拿这种状态去删旧版本，
' 等于把唯一能用的版本删掉，换来一个打得开但没功能的空壳。
'
' 所以这里真的去调一次载荷里的自检函数——能返回就说明工程编译通过、
' 宏可执行、入口点齐全。
'------------------------------------------------------------------------------
Private Function VerifyPayloadRuns() As Boolean
    If mPayloadWb Is Nothing Then Exit Function

    Dim reply As String
    On Error GoTo NotRunning
    reply = CStr(Application.Run("'" & mPayloadWb.Name & "'!Toolbox_SelfCheck"))

    ' 【只看开头是不是 "OK" 太松了】。自检串的格式是
    '     OK|ExcelToolbox|1.0.0|host=...|actions=59|...
    ' 要删掉用户唯一的退路，判据就得严一点：除了 OK 之外，
    ' 还要确认这确实是本工具箱的载荷、且命令注册表非空——
    ' 一个编译通过但注册表是空的载荷，装上了也没有任何按钮可用。
    If Left$(reply, 3) <> "OK|" Then GoTo NotRunning
    If InStr(1, reply, "|" & PAYLOAD_APP_ID & "|", vbTextCompare) = 0 Then GoTo NotRunning

    Dim actionsPos As Long
    actionsPos = InStr(1, reply, "actions=", vbTextCompare)
    If actionsPos = 0 Then GoTo NotRunning
    If Val(Mid$(reply, actionsPos + Len("actions="))) <= 0 Then GoTo NotRunning

    VerifyPayloadRuns = True
    Exit Function

NotRunning:
    mLastUpdateNote = "新版本自检未通过，保留旧版本缓存：" & Left$(reply, 120)
    VerifyPayloadRuns = False
End Function

'------------------------------------------------------------------------------
' 删除缓存目录里除【当前正在用的这一个】之外的所有历史载荷。
'
' 为什么可以只留一个：当前这个版本已经验证过能跑，它就是下一次升级失败时的退路。
' 再往前的版本没有任何价值，只是在用户的 %LOCALAPPDATA% 里越堆越多——
' 每个约 250KB，几十个版本之后在漫游配置文件和瘦客户端上是会有感的。
'
' 两条硬约束：
'   1.【删不掉就跳过，绝不能让启动失败】。别的 Excel 实例可能正开着旧版本，
'      文件被占用。那不是错误，下次启动再删就是了。
'   2.【只认自己的命名规则】。缓存目录里万一有别的东西，不关我们的事。
'------------------------------------------------------------------------------
Private Sub PurgeOldPayloads(ByVal activePath As String)
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub

    Dim dirPath As String
    dirPath = CacheDir()
    If Not fso.FolderExists(dirPath) Then Exit Sub

    Dim activeName As String
    activeName = LCase$(fso.GetFileName(activePath))

    Dim f As Object, nameOnly As String
    Dim deleted As Long, skipped As String
    For Each f In fso.GetFolder(dirPath).Files
        nameOnly = fso.GetFileName(f.Path)

        If LCase$(nameOnly) <> activeName _
           And Left$(nameOnly, Len(PAYLOAD_PREFIX)) = PAYLOAD_PREFIX _
           And LCase$(Right$(nameOnly, Len(PAYLOAD_EXT))) = LCase$(PAYLOAD_EXT) Then

            Err.Clear
            fso.DeleteFile f.Path, True

            ' 【删不掉恰恰说明别人正用着它，这是安全行为不是错误】。
            ' Windows 会锁住被 Excel 打开的文件，所以"另一个实例正开着旧版本"
            ' 这种情况下 DeleteFile 必然失败——不需要额外的占用检测，
            ' 文件锁本身就是那道保险。跳过即可，下次启动再删。
            '
            ' 但【要记下来】：静默吞掉的话，缓存一直清不干净时没人知道为什么。
            If Err.Number <> 0 Then
                If Len(skipped) > 0 Then skipped = skipped & ", "
                skipped = skipped & nameOnly
                Err.Clear
            Else
                deleted = deleted + 1
            End If
        End If
    Next f

    If deleted > 0 Or Len(skipped) > 0 Then
        mLastUpdateNote = mLastUpdateNote & _
            IIf(Len(mLastUpdateNote) > 0, "；", "") & _
            "清理旧版本：删除 " & deleted & " 个" & _
            IIf(Len(skipped) > 0, "，跳过（被占用）：" & skipped, "")
    End If

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 由 ThisWorkbook.Workbook_BeforeClose 调用
'------------------------------------------------------------------------------
Public Sub Shutdown()
    On Error Resume Next
    If Not mPayloadWb Is Nothing Then
        mPayloadWb.Close SaveChanges:=False
        Set mPayloadWb = Nothing
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' 更新逻辑
'==============================================================================

'------------------------------------------------------------------------------
' 确保本地有最新可用的载荷，返回它的完整路径。
' 一个可用版本都没有时返回空串。
'------------------------------------------------------------------------------
Private Function EnsureLatestPayload() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    mLastUpdateNote = vbNullString

    Dim cacheRoot As String
    cacheRoot = CacheDir()
    If Not fso.FolderExists(cacheRoot) Then
        On Error Resume Next
        fso.CreateFolder cacheRoot
        On Error GoTo 0
    End If

    ' 清单说了算。读不到（断网、出差、VPN 掉线）就用缓存里最新的那个版本。
    Dim wanted As String
    wanted = ReadManifest(fso)

    If Len(wanted) = 0 Then
        If Not fso.FolderExists(SharePath()) Then
            mLastUpdateNote = "连不上发布目录，使用本地缓存"
        Else
            mLastUpdateNote = "清单缺失或内容非法，使用本地缓存"
        End If
    End If

    If Len(wanted) > 0 Then
        Dim wantedPath As String
        wantedPath = fso.BuildPath(cacheRoot, PAYLOAD_PREFIX & wanted & PAYLOAD_EXT)

        ' 缓存里已有【且完整】才直接用。回滚场景下目标版本往往已经在缓存里，不必重下。
        '
        ' 【只判断文件存在是不够的】：并发或上次中断都可能留下一个尺寸不对的文件，
        ' 那时候直接打开，用户看到的是"文件已损坏"，还完全不知道为什么。
        If Not IsPayloadComplete(fso, wantedPath) Then
            CopyPayload fso, wanted, cacheRoot
        End If

        If IsPayloadComplete(fso, wantedPath) Then
            EnsureLatestPayload = wantedPath
            mLastUpdateNote = "ok:" & wanted
            Exit Function
        End If
    End If

    ' 走到这里说明"清单要的那个版本"没拿到（读不到清单、或下载/校验失败），
    ' 只能退回缓存里最新的。诊断信息必须把【想要哪个】和【实际用哪个】都写清楚，
    ' 否则维护者看到 active 和 remote 不一致时根本不知道卡在哪一步。
    Dim fallback As String
    fallback = NewestCachedVersion(fso, True)      ' 只挑内容完好的

    If Len(fallback) = 0 Then
        ' 【不要覆盖掉原因】。CopyPayload 已经写明了是"源文件不存在"还是"复制失败"，
        ' 这里直接改写成一句笼统的话，维护者就分不出是断网、没发版还是权限不足了。
        If Len(wanted) > 0 Then
            mLastUpdateNote = "需要 " & wanted & " 但未能获取（" & _
                              IIf(Len(mLastUpdateNote) > 0, mLastUpdateNote, "原因未知") & _
                              "），且缓存里没有可用版本"
        Else
            mLastUpdateNote = mLastUpdateNote & "；缓存里也没有可用版本"
        End If
        Exit Function
    End If

    ' 走到这里 fallback 已经通过完整性检查，直接用
    Dim fallbackPath As String
    fallbackPath = fso.BuildPath(cacheRoot, PAYLOAD_PREFIX & fallback & PAYLOAD_EXT)

    EnsureLatestPayload = fallbackPath
    If Len(wanted) > 0 Then
        mLastUpdateNote = "需要 " & wanted & " 但未能获取（" & mLastUpdateNote & _
                          "），已退回缓存版本 " & fallback
    Else
        mLastUpdateNote = mLastUpdateNote & "，使用缓存版本 " & fallback
    End If
End Function

'------------------------------------------------------------------------------
' 缓存里的载荷是否完整。
'
' 判据三层，按可靠性从高到低：
'   1. 大小下限（xlam 是 zip 包，再小也不止几 KB）；
'   2. 发布目录可达时，与源文件大小严格比对——最权威；
'   3. 断网时退而看内容：zip 魔数 "PK"，挡住"大小还在但内容已坏"。
'
' 【核心原则：要正面证据才定罪】。
'
' 早先这里用 OpenTextFile 探"有没有被别人占用"，那是错的——
' FileSystemObject 根本没有独占打开模式，打得开只说明能按文本读；
' 更糟的是反方向：同事开着两个 Excel 窗口是家常便饭，第一个已经打开了载荷，
' 第二个来探测就可能失败，于是把【完好的缓存】判成不可用，弹出"无法启动"。
'
' 所以现在读不到内容时一律放行——"说不清"不等于"有问题"。
' 至于"半截文件"，本设计里复制一律走唯一临时名 + 改名，
' 最终文件按构造要么不存在、要么完整，这个担心本就不成立。
'------------------------------------------------------------------------------
Private Function IsPayloadComplete(ByVal fso As Object, ByVal filePath As String, _
                                   Optional ByVal strictRead As Boolean = False, _
                                   Optional ByVal shareUp As Long = -1) As Boolean
    On Error GoTo NotComplete

    If Not fso.FileExists(filePath) Then Exit Function

    Dim localSize As Double
    localSize = fso.GetFile(filePath).Size

    ' xlam 是 zip 包，再小也不止几 KB
    If localSize < 8192 Then Exit Function

    ' 发布目录可达时用源文件大小做权威比对。
    ' shareUp 由调用方传进来，避免在候选文件多的时候对着一个不可达的 UNC
    ' 反复探测——每次都要等网络超时，启动会被拖得很慢。
    If shareUp = -1 Then shareUp = IIf(fso.FolderExists(SharePath()), 1, 0)

    If shareUp = 1 Then
        Dim srcPath As String
        srcPath = fso.BuildPath(SharePath(), fso.GetFileName(filePath))
        If fso.FileExists(srcPath) Then
            ' 大小对不上就当不完整，随后重新复制——多下一次而已，是自愈的
            If fso.GetFile(srcPath).Size <> localSize Then Exit Function
        End If
    End If

    ' 内容自检：xlam 是 zip 包，头两个字节必然是 "PK"。
    ' 这一步挡住磁盘损坏、杀毒软件截断这类"大小还在但内容已坏"的情况。
    If Not LooksLikeZip(fso, filePath, strictRead) Then Exit Function

    IsPayloadComplete = True
    Exit Function

NotComplete:
    IsPayloadComplete = False
End Function

'------------------------------------------------------------------------------
' 文件是否以 zip 魔数开头。
'
' 【读不出来时怎么办，取决于这个文件是哪来的】——这两种情况性质完全不同：
'
'   strictRead = True：刚复制完的文件。我们自己刚写的东西居然读不了，
'     那就是可疑（盘坏了、杀毒插手了），拒绝，重新来过代价很小。
'
'   strictRead = False：断网降级时用的既有缓存。读不到的常见原因是
'     "另一个 Excel 正开着它"，这完全正常。这时候拒绝就是误杀——
'     两轮之前正是因为把"读不出来"当成"文件有问题"，
'     导致同事开两个 Excel 窗口时完好的缓存被判成不可用。
'
' 换句话说：能重来的场合从严，重来就没得用的场合从宽。
'------------------------------------------------------------------------------
Private Function LooksLikeZip(ByVal fso As Object, ByVal filePath As String, _
                              ByVal strictRead As Boolean) As Boolean
    On Error GoTo CannotTell

    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)

    Dim head As String
    If Not ts.AtEndOfStream Then head = ts.Read(2)
    ts.Close

    If Len(head) = 2 Then
        LooksLikeZip = (head = "PK")
    Else
        ' 读到的字节数不对：文件是空的或者截断了
        LooksLikeZip = Not strictRead
    End If
    Exit Function

CannotTell:
    LooksLikeZip = Not strictRead
End Function

'------------------------------------------------------------------------------
' 读共享盘上的清单。内容就是一行版本号，例如 1.2.0。
'------------------------------------------------------------------------------
Private Function ReadManifest(ByVal fso As Object) As String
    Dim manifestPath As String
    manifestPath = fso.BuildPath(SharePath(), MANIFEST_NAME)

    On Error GoTo NoManifest
    If Not fso.FileExists(manifestPath) Then Exit Function

    Dim ts As Object
    Set ts = fso.OpenTextFile(manifestPath, 1)     ' ForReading
    Dim raw As String
    If Not ts.AtEndOfStream Then raw = ts.ReadLine
    ts.Close

    ReadManifest = SanitizeVersion(raw)
    Exit Function

NoManifest:
    ' 共享盘连不上是常态（出差、VPN 断开），静默退回本地缓存
    ReadManifest = vbNullString
End Function

'------------------------------------------------------------------------------
' 版本号只允许数字和点。
'
' 这一步是必须的：版本号会被直接拼进文件路径，不过滤的话，清单里写一个
' "..\..\Windows\System32\evil" 就能让加载器去别处取文件。
'------------------------------------------------------------------------------
' 只接受 N.N / N.N.N / N.N.N.N，每段 1-4 位数字。
'
' 光过滤字符是不够的：".", "1..2", "00000000000000001" 都只含数字和点，
' 但会变成怪异的文件名，超长数字段还会让后面 CompareVersions 的 CLng 溢出。
' 清单被写坏时宁可整条作废、继续用旧版本，也不能拿它去拼路径。
Private Function SanitizeVersion(ByVal raw As String) As String
    Dim buf As String
    buf = Trim$(raw)
    If Len(buf) = 0 Then Exit Function

    Dim parts As Variant
    parts = Split(buf, ".")

    If UBound(parts) < 1 Or UBound(parts) > 3 Then Exit Function

    Dim i As Long, seg As String, j As Long, ch As String
    For i = 0 To UBound(parts)
        seg = parts(i)
        If Len(seg) < 1 Or Len(seg) > 4 Then Exit Function
        For j = 1 To Len(seg)
            ch = Mid$(seg, j, 1)
            If ch < "0" Or ch > "9" Then Exit Function
        Next j
    Next i

    SanitizeVersion = buf
End Function

'------------------------------------------------------------------------------
' 把指定版本的载荷从共享盘复制到本地缓存。
'
' 先复制到临时名再改名：复制到一半断网的话，留下的是个半截文件；
' 直接用目标名复制会让下次启动加载到这个残缺文件，而且版本号记录显示是好的。
'------------------------------------------------------------------------------
Private Function CopyPayload(ByVal fso As Object, ByVal version As String, _
                             ByVal cacheRoot As String) As Boolean
    Dim srcPath As String, finalPath As String, tempPath As String
    srcPath = fso.BuildPath(SharePath(), PAYLOAD_PREFIX & version & PAYLOAD_EXT)
    finalPath = fso.BuildPath(cacheRoot, PAYLOAD_PREFIX & version & PAYLOAD_EXT)

    ' 【临时文件名必须每个进程唯一】。
    ' 早上九点几十号人同时开 Excel 是常态，固定用 "<版本>.part" 的话：
    ' 进程 A 正在写，进程 B 一上来就把它删了重写，两边再交错改名，
    ' 最后谁也说不清缓存里那个文件是完整的还是拼出来的。
    tempPath = finalPath & "." & CStr(GetCurrentProcessId()) & "." & _
               Format$(Now, "hhnnss") & CStr(Int(Rnd() * 100000)) & ".part"

    On Error GoTo CopyFailed
    If Not fso.FileExists(srcPath) Then
        mLastUpdateNote = "载荷在发布目录里不存在：" & PAYLOAD_PREFIX & version & PAYLOAD_EXT
        Exit Function
    End If

    fso.CopyFile srcPath, tempPath, True

    ' 复制完先自检一次，半截文件不许进缓存
    ' 自己刚写完的文件，有权要求它读得出来：用严格模式
    If Not IsPayloadComplete(fso, tempPath, True) Then
        mLastUpdateNote = "复制结果不完整，已丢弃"
        GoTo CopyFailed
    End If

    ' 改名到最终名。这里【不先删目标】：
    ' 并发时另一个进程可能刚好已经放好了同一个版本，把它删掉反而制造空窗，
    ' 让第三个进程读到"文件不存在"。改名失败就看目标是不是已经好了——
    ' 是的话本来就该用它，等于别人替我们干完了。
    On Error Resume Next
    Err.Clear
    fso.MoveFile tempPath, finalPath
    Dim moveErr As Long
    moveErr = Err.Number
    Err.Clear
    On Error GoTo CopyFailed

    If moveErr <> 0 Then
        If IsPayloadComplete(fso, finalPath) Then
            ' 别的进程已经装好了，清掉自己的临时文件即可
            On Error Resume Next
            fso.DeleteFile tempPath, True
            On Error GoTo 0
            mLastUpdateNote = "ok:" & version & "（由另一个 Excel 进程完成）"
            CopyPayload = True
            Exit Function
        End If
        mLastUpdateNote = "写入缓存失败（文件被占用？）"
        GoTo CopyFailed
    End If

    mLastUpdateNote = "ok:" & version
    CopyPayload = True
    Exit Function

CopyFailed:
    On Error Resume Next
    If Len(mLastUpdateNote) = 0 Then mLastUpdateNote = "复制失败：" & Err.Description
    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    On Error GoTo 0
    CopyPayload = False
End Function

'------------------------------------------------------------------------------
' 打开载荷。
'
' 用 Workbooks.Open 打开 .xlam，它的 customUI 功能区会正常加载——
' 这一点已由 tests\check-ribbon.ps1 验证过，整个方案就建立在这个行为上。
'------------------------------------------------------------------------------
Private Sub OpenPayload(ByVal payloadPath As String)
    On Error GoTo OpenFailed

    Set mPayloadWb = Application.Workbooks.Open( _
                        Filename:=payloadPath, _
                        UpdateLinks:=0, _
                        ReadOnly:=False, _
                        AddToMru:=False, _
                        Notify:=False)
    Exit Sub

OpenFailed:
    ' 打开失败时把引用清干净：留一个半吊子的 Workbook 对象，
    ' 后面所有"mPayloadWb Is Nothing"的判断都会得出错误结论
    Set mPayloadWb = Nothing
    mLastUpdateNote = "载荷打开失败：" & Err.Description

    MsgBox "Excel 工具箱载荷打开失败：" & vbCrLf & payloadPath & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "常见原因：该目录不在 Excel 的受信任位置里，宏被禁用了。", _
           vbExclamation, "Excel 工具箱"
End Sub

'==============================================================================
' 路径与版本记录
'==============================================================================

' 发布目录。构建时由 build-loader.ps1 -SharePath 写死进来。
'
' 【刻意不做运行时配置】。原本这里读注册表，但实测 Excel 宏里
' CreateObject("WScript.Shell") 会被安全策略拦下（它是宏病毒的典型载体，
' 企业 AV 和 Windows ASR 规则普遍会阻止 Office 创建它），读取静默失败、
' 悄悄退回默认值——在同事的机器上这种失败根本看不出来。
'
' 发布目录本来也几乎不会变，为它引入一条随时可能被安全软件掐断的依赖不划算。
' 要改路径就重新构建一个加载器，反而清楚。
Private Function SharePath() As String
    SharePath = DEFAULT_SHARE
End Function

' 本地缓存目录。放 LOCALAPPDATA，不会被漫游配置文件同步来同步去。
Private Function CacheDir() As String
    CacheDir = Environ$("LOCALAPPDATA") & "\ExcelToolbox"
End Function

'------------------------------------------------------------------------------
' 缓存里现有的最高版本。
'
' 【不记录"已安装版本"】：状态存哪都可能和实际文件对不上——
' 记录说装了 1.2.0 但文件被杀毒删了，加载器就会一直打不开还找不到原因。
' 直接看目录里有什么，是什么就是什么，自带自愈能力。
'------------------------------------------------------------------------------
' 参数 requireComplete = True 时，只挑【内容完整】的版本。
'
' 降级时必须这样挑：只看"版本号最大"的话，一旦最新那个缓存坏了就直接放弃，
' 而旁边可能正躺着一个完好的上一版。用户明明有得用，却被告知"无可用载荷"。
Private Function NewestCachedVersion(ByVal fso As Object, _
                                     Optional ByVal requireComplete As Boolean = False) As String
    Dim dirPath As String
    dirPath = CacheDir()
    If Not fso.FolderExists(dirPath) Then Exit Function

    ' 【可达性只探一次】。否则候选版本一多，就会对着一个连不上的 UNC
    ' 反复等网络超时，降级路径反而比正常路径还慢。
    Dim shareUp As Long
    shareUp = IIf(fso.FolderExists(SharePath()), 1, 0)

    Dim f As Object, nameOnly As String, ver As String
    Dim best As String

    For Each f In fso.GetFolder(dirPath).Files
        nameOnly = fso.GetFileName(f.Path)
        If LCase$(Right$(nameOnly, Len(PAYLOAD_EXT))) = LCase$(PAYLOAD_EXT) Then
            If Left$(nameOnly, Len(PAYLOAD_PREFIX)) = PAYLOAD_PREFIX Then
                ver = Mid$(nameOnly, Len(PAYLOAD_PREFIX) + 1, _
                           Len(nameOnly) - Len(PAYLOAD_PREFIX) - Len(PAYLOAD_EXT))
                ver = SanitizeVersion(ver)
                If Len(ver) > 0 Then
                    If (Not requireComplete) Or _
                       IsPayloadComplete(fso, f.Path, False, shareUp) Then
                        If Len(best) = 0 Then
                            best = ver
                        ElseIf CompareVersions(ver, best) > 0 Then
                            best = ver
                        End If
                    End If
                End If
            End If
        End If
    Next f

    NewestCachedVersion = best
End Function

' 按数字逐段比较，不能直接比字符串：字符串比较会认为 "1.9" > "1.10"
Private Function CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim pa As Variant, pb As Variant
    pa = Split(a, ".")
    pb = Split(b, ".")

    Dim n As Long, i As Long, va As Long, vb As Long
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)

    For i = 0 To n
        va = 0: vb = 0
        If i <= UBound(pa) Then If IsNumeric(pa(i)) Then va = CLng(pa(i))
        If i <= UBound(pb) Then If IsNumeric(pb(i)) Then vb = CLng(pb(i))
        If va > vb Then CompareVersions = 1: Exit Function
        If va < vb Then CompareVersions = -1: Exit Function
    Next i
End Function

'==============================================================================
' 对外诊断入口（供 install/测试脚本调用）
'==============================================================================

' 当前真正打开的是哪个版本（从已打开的载荷文件名反推，不依赖任何记录）
Private Function ActiveVersion() As String
    If mPayloadWb Is Nothing Then Exit Function
    On Error Resume Next
    Dim n As String
    n = mPayloadWb.Name
    If Left$(n, Len(PAYLOAD_PREFIX)) = PAYLOAD_PREFIX Then
        ActiveVersion = Mid$(n, Len(PAYLOAD_PREFIX) + 1, _
                             Len(n) - Len(PAYLOAD_PREFIX) - Len(PAYLOAD_EXT))
    End If
    On Error GoTo 0
End Function

Public Function Loader_Status() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Loader_Status = "share=" & SharePath() & _
                    "|shareReachable=" & CStr(fso.FolderExists(SharePath())) & _
                    "|remote=" & ReadManifest(fso) & _
                    "|cached=" & NewestCachedVersion(fso) & _
                    "|active=" & ActiveVersion() & _
                    "|cacheDir=" & CacheDir() & _
                    "|payloadOpen=" & CStr(Not (mPayloadWb Is Nothing)) & _
                    "|lastUpdate=" & mLastUpdateNote
End Function

' 只做"检查 + 下载"，不打开载荷。
' 拆出来是为了能单独验证更新逻辑——打开载荷要真的加载一个加载宏，
' 无界面环境下那一步的行为和交互式差别很大，混在一起测不出问题在哪。
Public Function Loader_UpdateOnly() As String
    On Error GoTo Failed
    Dim p As String
    p = EnsureLatestPayload()
    If Len(p) = 0 Then Loader_UpdateOnly = "NONE" Else Loader_UpdateOnly = p
    Exit Function
Failed:
    Loader_UpdateOnly = "ERROR: " & Err.Description & " [" & Err.Number & "]"
End Function

' 强制重新检查更新（不用重启 Excel），供维护者排障
Public Function Loader_CheckNow() As String
    ' 【先确定新载荷在哪，再关旧的】。
    ' 反过来写的话，一旦解析或下载失败，用户就从"有一个能用的工具箱"
    ' 变成"什么都没有了"——为了检查更新反而把好好的东西弄丢，不可接受。
    Dim p As String
    p = EnsureLatestPayload()

    If Len(p) = 0 Then
        Loader_CheckNow = "没有可用版本，保持当前已加载的载荷不变。" & vbCrLf & Loader_Status()
        Exit Function
    End If

    ' 已经就是它了，不折腾
    If Not mPayloadWb Is Nothing Then
        On Error Resume Next
        Dim curPath As String
        curPath = mPayloadWb.FullName
        On Error GoTo 0
        If StrComp(curPath, p, vbTextCompare) = 0 Then
            Loader_CheckNow = "已是最新：" & p
            Exit Function
        End If
    End If

    Dim previousPath As String
    On Error Resume Next
    If Not mPayloadWb Is Nothing Then previousPath = mPayloadWb.FullName
    On Error GoTo 0

    Shutdown
    OpenPayload p

    If mPayloadWb Is Nothing Then
        ' 新载荷打不开：尽量把旧的接回来，别让用户两手空空
        If Len(previousPath) > 0 Then
            OpenPayload previousPath
            If Not mPayloadWb Is Nothing Then
                Loader_CheckNow = "新载荷打开失败，已退回原版本：" & previousPath
                Exit Function
            End If
        End If
        Loader_CheckNow = "载荷打开失败，当前没有可用的工具箱。" & vbCrLf & Loader_Status()
        Exit Function
    End If

    Loader_CheckNow = "已加载：" & p
End Function
