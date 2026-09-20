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

Private mPayloadWb As Workbook

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

    If Len(wanted) > 0 Then
        Dim wantedPath As String
        wantedPath = fso.BuildPath(cacheRoot, PAYLOAD_PREFIX & wanted & PAYLOAD_EXT)

        ' 缓存里没有就去拉。回滚场景下目标版本往往已经在缓存里，直接用，不必重下。
        If Not fso.FileExists(wantedPath) Then
            CopyPayload fso, wanted, cacheRoot
        End If

        If fso.FileExists(wantedPath) Then
            EnsureLatestPayload = wantedPath
            Exit Function
        End If
    End If

    Dim fallback As String
    fallback = NewestCachedVersion(fso)
    If Len(fallback) = 0 Then Exit Function

    Dim fallbackPath As String
    fallbackPath = fso.BuildPath(cacheRoot, PAYLOAD_PREFIX & fallback & PAYLOAD_EXT)
    If fso.FileExists(fallbackPath) Then EnsureLatestPayload = fallbackPath
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
Private Function SanitizeVersion(ByVal raw As String) As String
    Dim buf As String, i As Long, ch As String
    buf = Trim$(raw)

    For i = 1 To Len(buf)
        ch = Mid$(buf, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            SanitizeVersion = SanitizeVersion & ch
        Else
            ' 出现任何其它字符就整体作废，不做"尽量修复"——
            ' 清单被写坏时，宁可不更新也不能去加载一个来路不明的路径
            SanitizeVersion = vbNullString
            Exit Function
        End If
    Next i
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
    tempPath = finalPath & ".part"

    On Error GoTo CopyFailed
    If Not fso.FileExists(srcPath) Then Exit Function

    If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    fso.CopyFile srcPath, tempPath, True

    ' 目标文件可能是上次留下的同名旧文件（回滚场景），先删再改名
    If fso.FileExists(finalPath) Then fso.DeleteFile finalPath, True
    fso.MoveFile tempPath, finalPath

    CopyPayload = True
    Exit Function

CopyFailed:
    On Error Resume Next
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
Private Function NewestCachedVersion(ByVal fso As Object) As String
    Dim dirPath As String
    dirPath = CacheDir()
    If Not fso.FolderExists(dirPath) Then Exit Function

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
                    If Len(best) = 0 Then
                        best = ver
                    ElseIf CompareVersions(ver, best) > 0 Then
                        best = ver
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
                    "|payloadOpen=" & CStr(Not (mPayloadWb Is Nothing))
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
    Shutdown

    Dim p As String
    p = EnsureLatestPayload()
    If Len(p) = 0 Then
        Loader_CheckNow = "没有可用版本。" & vbCrLf & Loader_Status()
        Exit Function
    End If

    OpenPayload p
    Loader_CheckNow = "已加载：" & p
End Function
