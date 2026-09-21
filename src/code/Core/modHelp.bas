Attribute VB_Name = "modHelp"
'==============================================================================
' modHelp - 帮助系统
'
' 参考 Office 自己的帮助形态：功能区有入口、命令有上下文帮助、有搜索框。
' 但【内容不是手写死的】——标题、可否撤销、会不会追问参数全部从
' modAction 的命令注册表实时取，只有"什么时候用/怎么用/注意"来自
' src\help\help.md（构建时注入隐藏工作表 _Help）。
'
' 【为什么要这么拆】：功能区的标签当初就是因为 XML 和注册表各写一份，
' 出现过"按钮叫删除重复、撤销按钮却显示撤销删除重复值"。帮助文档比标签长得多，
' 两边各写一份必然漂移，而且漂移了没人会发现——帮助没人天天看。
' 所以凡是注册表里已有的信息，这里一律不重复存储。
'
' 产物是一个自包含的 HTML，写到 %TEMP% 再用默认浏览器打开：
' 不需要服务器、不需要额外分发文件，整个工具箱仍然只有一个 .xlam。
'==============================================================================
Option Explicit
Option Private Module

Private Const HELP_SHEET As String = "_Help"

'==============================================================================
' 对外入口
'==============================================================================

'------------------------------------------------------------------------------
' 打开帮助总览。功能区「帮助」按钮调用。
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' 打开右侧帮助侧边栏。功能区「帮助」按钮走这里。
'
' 【它不是 Office 原生的任务窗格】：真正的任务窗格只有 COM 加载项能创建，
' 纯 VBA 的 .xlam 拿不到那个接口。这里是一个贴在右缘的无模式窗体，
' 能常驻、能边看边操作，但不会把表格区域挤窄。
'
' 完整 HTML 那条路保留着（侧边栏底部有入口），要打印或全文检索时更合适。
'------------------------------------------------------------------------------
Public Function ShowPane(Optional ByVal entryId As String = vbNullString) As String
    ' 静默模式（回归测试）里绝不能弹窗体：它会一直等在那儿，把测试挂死
    If modAction.IsSilent() Then Exit Function

    On Error GoTo Failed

    frmHelpPane.DockRight
    ' 带了条目就直接定位过去——功能区的「帮助」是总览，
    ' 而命令执行失败时的「查看帮助」要落到出问题的那一条上。
    If Len(entryId) > 0 Then frmHelpPane.ShowEntry entryId
    frmHelpPane.Show vbModeless
    Exit Function

Failed:
    ' 侧边栏开不出来时退回浏览器版本，至少让用户看得到帮助
    ShowPane = ShowAll()
End Function

Public Function ShowAll() As String
    Dim path As String
    path = BuildHtml(vbNullString)
    If Len(path) = 0 Then
        ShowAll = "帮助内容不可用。加载宏可能损坏，请重新安装。"
        Exit Function
    End If

    OpenInBrowser path
    ShowAll = vbNullString          ' 成功时不弹框，浏览器已经开了
End Function

'------------------------------------------------------------------------------
' 打开某个命令的帮助。从报错对话框的「查看帮助」进来。
'------------------------------------------------------------------------------
Public Function ShowFor(ByVal actionId As String) As String
    Dim path As String
    path = BuildHtml(actionId)
    If Len(path) = 0 Then
        ShowFor = "帮助内容不可用。"
        Exit Function
    End If

    OpenInBrowser path
    ShowFor = vbNullString
End Function

'------------------------------------------------------------------------------
' 「我要做什么」搜索。功能区搜索框调用。
'
' 把命中的命令用帮助页展示出来，让用户看清楚再去点对应的按钮。
'
' 【为什么不"命中一条就直接执行"】：
' 第一版就是那么做的，理由是"用户打『求和是0』是想解决问题，
' 少一次点击就少一次放弃的机会"。这个理由本身没错，但代价没算清楚——
' 注册表里有 data.deleteDuplicates、data.deleteEmptyRows、file.batchRename、
' formula.breakLinks 这类会改数据甚至改磁盘文件的命令。
' 用户在搜索框里打几个字、按回车，然后数据就被改了，
' 这不是"省一次点击"，是【从一个输入框里静默触发了破坏性操作】。
'
' 搜索框的职责是"帮你找到功能"，不是"替你决定执行"。
' 省下的那一次点击，远不值得冒一次意外删除的风险。
'------------------------------------------------------------------------------
Public Function Search(ByVal query As String) As String
    Dim q As String
    q = Trim$(query)
    If Len(q) = 0 Then
        Search = "请输入你想做的事，例如「求和是0」「合并文件」「去重」。"
        Exit Function
    End If

    Dim hits As Collection
    Set hits = FindMatches(q)

    If hits.Count = 0 Then
        Search = "没找到和「" & q & "」相关的功能。" & vbCrLf & vbCrLf & _
                 "换个说法试试，或者点「帮助」浏览全部功能。"
        Exit Function
    End If

    Dim path As String
    path = BuildHtmlForSet(hits, q)
    If Len(path) = 0 Then
        ' 帮助内容不可用时至少把功能名告诉用户，别让他一无所获
        Dim names As String, i As Long
        For i = 1 To hits.Count
            If Len(names) > 0 Then names = names & "、"
            names = names & modAction.ActionLabel(CStr(hits(i)))
        Next i
        Search = "找到 " & hits.Count & " 个相关功能：" & vbCrLf & vbCrLf & names & _
                 vbCrLf & vbCrLf & "（帮助内容不可用，请在功能区里找上述按钮）"
        Exit Function
    End If

    OpenInBrowser path
    Search = vbNullString
End Function

'------------------------------------------------------------------------------
' 只解析不执行，返回命中的 actionId（换行分隔）。
'
' 【给测试用的】：Search 在只命中一条时会直接执行那条命令，
' 测试里没法断言"搜『求和是0』应该命中文本转数值"而不真的动数据。
' 拆一个只读入口出来，断言匹配逻辑本身。
'------------------------------------------------------------------------------
Public Function Resolve(ByVal query As String) As String
    Dim hits As Collection
    Set hits = FindMatches(Trim$(query))
    If hits Is Nothing Then Exit Function

    Dim i As Long, s As String
    For i = 1 To hits.Count
        If Len(s) > 0 Then s = s & vbLf
        s = s & CStr(hits(i))
    Next i
    Resolve = s
End Function

'==============================================================================
' 查询
'==============================================================================

Private Function HelpSheet() As Object
    On Error Resume Next
    Set HelpSheet = ThisWorkbook.Worksheets(HELP_SHEET)
    Err.Clear
    On Error GoTo 0
End Function

' 某个 actionId 有没有帮助条目。供 check-help.ps1 断言用。
Public Function HasEntry(ByVal actionId As String) As Boolean
    HasEntry = (Len(BodyOf(actionId)) > 0)
End Function

Public Function BodyOf(ByVal actionId As String) As String
    On Error Resume Next
    Dim sh As Object
    Set sh = HelpSheet()
    If sh Is Nothing Then Exit Function

    Dim r As Long, lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(-4162).Row      ' xlUp
    For r = 2 To lastRow
        If StrComp(CStr(sh.Cells(r, 1).Value2), actionId, vbTextCompare) = 0 Then
            BodyOf = CStr(sh.Cells(r, 3).Value2)
            Exit Function
        End If
    Next r

    Err.Clear
    On Error GoTo 0
End Function

Private Function KeywordsOf(ByVal actionId As String) As String
    On Error Resume Next
    Dim sh As Object
    Set sh = HelpSheet()
    If sh Is Nothing Then Exit Function

    Dim r As Long, lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(-4162).Row
    For r = 2 To lastRow
        If StrComp(CStr(sh.Cells(r, 1).Value2), actionId, vbTextCompare) = 0 Then
            KeywordsOf = CStr(sh.Cells(r, 2).Value2)
            Exit Function
        End If
    Next r

    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 找出和查询词相关的 actionId。
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' 找出和查询词相关的 actionId，按相关度从高到低。
'
' 【不能分成"第一轮有结果就不跑第二轮"】。第一版是那么写的，
' 后果是标题命中会把正文命中整个屏蔽掉：
' 搜"重复"命中了标题里带"重复"的两条，于是正文里讲"重复"的
' 「提取唯一值」就再也出不来了——而那很可能正是用户要找的。
'
' 现在一轮扫完全部来源，按命中位置打分，高分在前、低分在后，
' 不丢任何一条。
'------------------------------------------------------------------------------
Private Function FindMatches(ByVal q As String) As Collection
    Dim strong As New Collection      ' actionId / 标题命中
    Dim weak As New Collection        ' 关键词 / 正文命中

    Dim ids() As String, i As Long, id As String
    ids = Split(modAction.AllActionIds(), vbLf)

    For i = LBound(ids) To UBound(ids)
        id = ids(i)
        If Len(id) > 0 Then
            If InStr(1, id, q, vbTextCompare) > 0 _
               Or InStr(1, modAction.ActionLabel(id), q, vbTextCompare) > 0 Then
                strong.Add id
            ElseIf InStr(1, KeywordsOf(id), q, vbTextCompare) > 0 _
                Or InStr(1, BodyOf(id), q, vbTextCompare) > 0 Then
                weak.Add id
            End If
        End If
    Next i

    Dim result As New Collection
    For i = 1 To strong.Count
        result.Add strong(i)
    Next i
    For i = 1 To weak.Count
        result.Add weak(i)
    Next i

    Set FindMatches = result
End Function

'==============================================================================
' HTML 生成
'==============================================================================

Private Function BuildHtml(ByVal focusId As String) As String
    Dim ids() As String
    ids = Split(modAction.AllActionIds(), vbLf)

    Dim all As New Collection
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If Len(ids(i)) > 0 Then all.Add ids(i)
    Next i

    BuildHtml = BuildHtmlForSet(all, focusId)
End Function

Private Function BuildHtmlForSet(ByVal ids As Collection, ByVal note As String) As String
    On Error GoTo Failed
    If ids Is Nothing Then Exit Function
    If ids.Count = 0 Then Exit Function

    Dim sb As String
    sb = HtmlHead(note)

    Dim i As Long, id As String
    For i = 1 To ids.Count
        id = CStr(ids(i))
        sb = sb & HtmlEntry(id)
    Next i

    sb = sb & HtmlTail()

    Dim path As String
    path = Environ$("TEMP") & "\工具箱帮助.html"
    WriteUtf8 path, sb
    BuildHtmlForSet = path
    Exit Function

Failed:
    Err.Clear
End Function

Private Function HtmlEntry(ByVal id As String) As String
    Dim label As String, tip As String, body As String
    label = modAction.ActionLabel(id)
    tip = modAction.ActionScreentip(id)
    body = BodyOf(id)

    Dim s As String
    s = "<section id=""" & Esc(id) & """>" & vbCrLf
    s = s & "<h2>" & Esc(label) & " <code>" & Esc(id) & "</code></h2>" & vbCrLf

    ' 可撤销标注直接取注册表生成的那句，和按钮提示完全一致
    If Len(tip) > 0 Then
        s = s & "<p class=""tip"">" & Esc(tip) & "</p>" & vbCrLf
    End If

    If Len(body) > 0 Then
        s = s & MarkdownLite(body)
    Else
        s = s & "<p class=""missing"">（这条命令还没有写帮助正文）</p>" & vbCrLf
    End If

    s = s & "</section>" & vbCrLf
    HtmlEntry = s
End Function

'------------------------------------------------------------------------------
' 极简 Markdown 转换：只认 ### 小标题、**粗体**、段落。
'
' 【不引入完整 Markdown 解析】：帮助正文的格式是我们自己定的，
' 只用到这三种。为三种语法写一个解析器，比为了通用性引入一堆代码划算得多。
'------------------------------------------------------------------------------
Private Function MarkdownLite(ByVal src As String) As String
    Dim lines() As String, i As Long, line As String, s As String

    lines = Split(Replace(src, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo NextLine

        If Left$(line, 4) = "### " Then
            s = s & "<h3>" & Esc(Mid$(line, 5)) & "</h3>" & vbCrLf
        Else
            s = s & "<p>" & Bold(Esc(line)) & "</p>" & vbCrLf
        End If
NextLine:
    Next i

    MarkdownLite = s
End Function

' **粗体** → <strong>。成对出现才转，落单的星号原样保留。
Private Function Bold(ByVal src As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(src, "**")

    For i = LBound(parts) To UBound(parts)
        If i Mod 2 = 1 And i < UBound(parts) Then
            s = s & "<strong>" & parts(i) & "</strong>"
        ElseIf i Mod 2 = 1 Then
            s = s & "**" & parts(i)          ' 落单，还原
        Else
            s = s & parts(i)
        End If
    Next i

    Bold = s
End Function

Private Function HtmlHead(ByVal note As String) As String
    Dim s As String
    s = "<!doctype html><html lang=""zh-CN""><head><meta charset=""utf-8"">" & vbCrLf
    s = s & "<title>" & Esc(APP_NAME) & " 帮助</title>" & vbCrLf
    s = s & "<style>" & vbCrLf
    s = s & "body{font-family:""Microsoft YaHei"",sans-serif;max-width:820px;margin:0 auto;padding:24px;line-height:1.75;color:#222}" & vbCrLf
    s = s & "h1{font-size:22px;border-bottom:2px solid #217346;padding-bottom:8px}" & vbCrLf
    s = s & "h2{font-size:17px;margin-top:32px;color:#217346}" & vbCrLf
    s = s & "h2 code{font-size:12px;color:#888;font-weight:normal}" & vbCrLf
    s = s & "h3{font-size:14px;margin:14px 0 4px;color:#555}" & vbCrLf
    s = s & "p{margin:4px 0}" & vbCrLf
    s = s & ".tip{background:#f3f7f4;border-left:3px solid #217346;padding:8px 12px;color:#444;font-size:13px}" & vbCrLf
    s = s & ".missing{color:#c00}" & vbCrLf
    s = s & ".note{background:#fff8e1;border:1px solid #ffe082;padding:10px 14px;border-radius:4px}" & vbCrLf
    s = s & "section{border-bottom:1px solid #eee;padding-bottom:12px}" & vbCrLf
    s = s & "#toc{columns:3;font-size:13px;margin:16px 0 28px}" & vbCrLf
    s = s & "#toc a{color:#217346;text-decoration:none;display:block;padding:2px 0}" & vbCrLf
    s = s & "</style></head><body>" & vbCrLf
    s = s & "<h1>" & Esc(APP_NAME) & " 使用帮助</h1>" & vbCrLf

    If Len(note) > 0 Then
        s = s & "<p class=""note"">与「" & Esc(note) & "」相关的功能：</p>" & vbCrLf
    End If

    HtmlHead = s
End Function

Private Function HtmlTail() As String
    HtmlTail = "<p style=""margin-top:40px;color:#888;font-size:12px"">" & _
               Esc(APP_NAME) & " v" & APP_VERSION & "　本页由加载宏即时生成。</p>" & _
               "</body></html>"
End Function

' 【& 必须第一个换】，否则后面换出来的 &lt; 会被再换成 &amp;lt;。
' 单引号也一起转：当前 HTML 属性都用双引号，单引号不会突破边界，
' 但帮助正文来自可编辑的 help.md，统一转掉成本为零，不留口子。
Private Function Esc(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&#39;")
    Esc = t
End Function

'==============================================================================
' 输出
'==============================================================================

'------------------------------------------------------------------------------
' 写 UTF-8 文件。
'
' 【必须带 BOM】：浏览器读没有 BOM 的本地 HTML 文件时，即使有
' <meta charset="utf-8">，部分环境仍会按系统代码页解码，中文全变乱码。
' 加载项是本地文件打开（file://），没有 HTTP 头可依赖，BOM 是唯一可靠信号。
'------------------------------------------------------------------------------
Private Sub WriteUtf8(ByVal path As String, ByVal content As String)
    On Error Resume Next

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    If stm Is Nothing Then Exit Sub

    stm.Type = 2                 ' adTypeText
    stm.Charset = "utf-8"        ' ADODB 写 utf-8 时自带 BOM
    stm.Open
    stm.WriteText content
    stm.SaveToFile path, 2       ' adSaveCreateOverWrite
    stm.Close

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 用默认浏览器打开。
'
' 【不能用 WScript.Shell】——它是宏病毒的典型载体，企业 AV 和 ASR 规则
' 普遍阻止 Office 创建它，而且失败是静默的（本项目实测过）。
' ThisWorkbook.FollowHyperlink 是 Excel 自己的 API，不在那类拦截规则里。
'------------------------------------------------------------------------------
Private Sub OpenInBrowser(ByVal path As String)
    ' 静默模式下只生成文件不打开：测试里弹一个浏览器窗口出来，
    ' 和弹 MsgBox 一样是打扰，只是不会把测试挂死而已。
    If modAction.IsSilent() Then Exit Sub

    On Error Resume Next
    ThisWorkbook.FollowHyperlink path
    Err.Clear
    On Error GoTo 0
End Sub

'==============================================================================
' 侧边栏用的目录与正文组装
'
' 【逻辑全部放在这里，窗体只负责显示】。窗体在无头测试里跑不起来，
' 把组装逻辑写进窗体事件 = 又多一块没有任何测试覆盖的代码，
' 而这正是本项目一直在防的东西。下面每个函数都能被无头通路直接断言。
'==============================================================================

'------------------------------------------------------------------------------
' 分组清单，返回 "groupId|分组标题" 每行一条。
'
' 【使用配置排在最前】：用户装完打不开、宏被禁用的时候，
' 他要找的不是"文本处理"，是"为什么这玩意儿没反应"。
'------------------------------------------------------------------------------
Public Function CatalogGroups() As String
    Dim s As String
    s = "guide|使用前必读（Excel 配置）"
    s = s & vbLf & "core|撤销与关于"
    s = s & vbLf & "text|M1 文本处理"
    s = s & vbLf & "data|M2 数据处理"
    s = s & vbLf & "sheet|M3 工作表管理"
    s = s & vbLf & "merge|M4 多文件合并"
    s = s & vbLf & "file|M5 文件批处理"
    s = s & vbLf & "formula|M6 公式与引用"
    s = s & vbLf & "audit|M7 数据体检"
    s = s & vbLf & "viz|M8 数据可视化"
    s = s & vbLf & "misc|M9 辅助增强"
    CatalogGroups = s
End Function

'------------------------------------------------------------------------------
' 某一组下的条目，返回 "id|显示名" 每行一条。
'
' guide 组来自帮助表（它们不是命令，没有注册项）；
' 其余各组来自命令注册表，显示名实时取自注册表——
' 和功能区标签同源，不会漂移。
'------------------------------------------------------------------------------
Public Function CatalogItems(ByVal groupId As String) As String
    If LCase$(groupId) = "guide" Then
        CatalogItems = GuideItems()
        Exit Function
    End If

    Dim ids As Variant, i As Long, s As String, oneId As String
    ids = Split(modAction.AllActionIds(), vbLf)
    For i = LBound(ids) To UBound(ids)
        oneId = Trim$(CStr(ids(i)))
        If Len(oneId) > 0 Then
            If GroupOf(oneId) = LCase$(groupId) Then
                If Len(s) > 0 Then s = s & vbLf
                s = s & oneId & "|" & modAction.ActionLabel(oneId)
            End If
        End If
    Next i
    CatalogItems = s
End Function

'------------------------------------------------------------------------------
' actionId -> 分组。
'
' 【cells.* 归到 M1】：拆分合并单元格、合并相同项在功能清单里就列在
' M1 文本处理下。跟着功能清单走，别让帮助目录和文档各说一套。
'------------------------------------------------------------------------------
Private Function GroupOf(ByVal actionId As String) As String
    Dim prefix As String
    Dim dotPos As Long

    dotPos = InStr(actionId, ".")
    If dotPos <= 1 Then Exit Function
    prefix = LCase$(Left$(actionId, dotPos - 1))

    If prefix = "cells" Then
        GroupOf = "text"
    Else
        GroupOf = prefix
    End If
End Function

' 给测试用：断言没有命令落在目录之外
Public Function GroupOfAction(ByVal actionId As String) As String
    GroupOfAction = GroupOf(actionId)
End Function

'------------------------------------------------------------------------------
' guide.* 条目：id 与标题都来自帮助表（标题在第 4 列）。
'------------------------------------------------------------------------------
Private Function GuideItems() As String
    On Error Resume Next
    Dim sh As Object
    Set sh = HelpSheet()
    If sh Is Nothing Then Exit Function

    Dim r As Long, lastRow As Long, s As String
    Dim oneId As String, title As String

    lastRow = sh.Cells(sh.Rows.Count, 1).End(-4162).Row
    For r = 2 To lastRow
        oneId = CStr(sh.Cells(r, 1).Value2)
        If LCase$(Left$(oneId, 6)) = "guide." Then
            title = CStr(sh.Cells(r, 4).Value2)
            If Len(title) = 0 Then title = oneId
            If Len(s) > 0 Then s = s & vbLf
            s = s & oneId & "|" & title
        End If
    Next r

    Err.Clear
    On Error GoTo 0
    GuideItems = s
End Function

' 帮助表第 4 列（只有 guide.* 用得上）
Private Function TitleOf(ByVal entryId As String) As String
    On Error Resume Next
    Dim sh As Object
    Set sh = HelpSheet()
    If sh Is Nothing Then Exit Function

    Dim r As Long, lastRow As Long
    lastRow = sh.Cells(sh.Rows.Count, 1).End(-4162).Row
    For r = 2 To lastRow
        If StrComp(CStr(sh.Cells(r, 1).Value2), entryId, vbTextCompare) = 0 Then
            TitleOf = CStr(sh.Cells(r, 4).Value2)
            Exit Function
        End If
    Next r

    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' 侧边栏正文（纯文本，不是 HTML）。
'
' 命令条目的标题与可撤销性【实时取自注册表】，不从 help.md 读——
' 两边各写一份必然漂移，而帮助没人天天看，漂移了几个月都不会有人发现。
'------------------------------------------------------------------------------
Public Function RenderEntry(ByVal entryId As String) As String
    Dim body As String
    Dim head As String

    body = BodyOf(entryId)

    If LCase$(Left$(entryId, 6)) = "guide." Then
        head = TitleOf(entryId)
        If Len(head) = 0 Then head = entryId
    Else
        Dim d As clsActionDef
        Set d = modAction.GetAction(entryId)
        If d Is Nothing Then
            RenderEntry = "没有找到这个条目：" & entryId
            Exit Function
        End If
        head = d.Label
        If d.Undoable Then
            head = head & "　【可撤销】"
        Else
            head = head & "　【不可撤销】"
        End If
    End If

    If Len(body) = 0 Then
        RenderEntry = head & vbCrLf & String$(28, "-") & vbCrLf & _
                      "（这一条还没有帮助正文）"
        Exit Function
    End If

    RenderEntry = head & vbCrLf & String$(28, "-") & vbCrLf & vbCrLf & _
                  StripMarkdown(body)
End Function

'------------------------------------------------------------------------------
' help.md 是给 HTML 用的，带 ### 和 **。侧边栏是纯文本控件，
' 直接把这些标记显示出来很难看，这里做一次轻量清洗。
'
' 【不做完整 Markdown 渲染】：TextBox 显示不了富文本，做了也没用。
'------------------------------------------------------------------------------
Private Function StripMarkdown(ByVal src As String) As String
    Dim lines As Variant, i As Long, s As String, ln As String

    lines = Split(Replace(src, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = CStr(lines(i))
        If Left$(ln, 4) = "### " Then
            ln = "【" & Trim$(Mid$(ln, 5)) & "】"
        End If
        ln = Replace(ln, "**", "")
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & ln
    Next i

    StripMarkdown = s
End Function

'==============================================================================
' 环境体检
'
' 【只报告，绝不代改】。信任中心、受信任位置属于【安全设置】，
' 插件替用户改等于替他降低防护等级——哪怕他点了同意也不该由插件来做。
' 这里只把能观察到的事实摆出来，怎么改由用户按指引自己操作。
'==============================================================================
Public Function EnvReport() As String
    Dim s As String

    s = "环境检查结果" & vbCrLf & String$(28, "=") & vbCrLf & vbCrLf

    s = s & "工具箱版本：" & modApp.APP_VERSION & vbCrLf
    s = s & "宿主程序　：" & SafeHostName() & vbCrLf
    s = s & "版本 / 位数：" & SafeHostVersion() & " / " & modApp.HostBitness() & vbCrLf
    s = s & "已注册命令：" & modAction.ActionCount() & " 个" & vbCrLf
    s = s & "功能区加载：" & IIf(modRibbon.IsRibbonLoaded(), "正常", "【未加载】") & vbCrLf

    ' 宏能不能用不需要去读注册表——你能看到这份报告本身就是证据
    s = s & "宏的状态　：已启用（否则这份报告根本出不来）" & vbCrLf

    ' "改了数不更新"的高频元凶，而用户几乎不会想到是这里
    s = s & "重算模式　：" & CalcModeText() & vbCrLf

    s = s & vbCrLf & String$(28, "-") & vbCrLf
    s = s & "【无法自动检测的项】" & vbCrLf
    s = s & "是否设了受信任位置：查不到。" & vbCrLf
    s = s & "  读注册表要用的 WScript.Shell 会被企业安全策略静默拦截，" & vbCrLf
    s = s & "  本工具箱因此不依赖它，这里也不假装检测。" & vbCrLf
    s = s & "  【按症状判断】每次打开 Excel 都弹宏安全警告，" & vbCrLf
    s = s & "  【可能】是没设受信任位置——但也可能是宏设置本身、文件带着网络" & vbCrLf
    s = s & "  来源标记、或者组策略统一管控。要确认还得去信任中心看一眼。" & vbCrLf
    s = s & "  这几种情况的处理办法见左侧「使用前必读」那一组。" & vbCrLf

    s = s & vbCrLf & String$(28, "-") & vbCrLf
    s = s & "以上只是检查，工具箱不会替你修改任何安全设置。" & vbCrLf
    s = s & "需要改信任中心的，请按左侧「使用前必读」里的指引自己操作。"

    EnvReport = s
End Function

Private Function SafeHostName() As String
    On Error Resume Next
    SafeHostName = Application.Name
    If Len(SafeHostName) = 0 Then SafeHostName = "(取不到)"
    Err.Clear
    On Error GoTo 0
End Function

Private Function SafeHostVersion() As String
    On Error Resume Next
    SafeHostVersion = Application.Version
    If Len(SafeHostVersion) = 0 Then SafeHostVersion = "(取不到)"
    Err.Clear
    On Error GoTo 0
End Function

Private Function CalcModeText() As String
    On Error Resume Next
    Select Case Application.Calculation
        Case -4105: CalcModeText = "自动（正常）"
        Case -4135: CalcModeText = "【手动】——公式不会自动重算，这是「改了数不更新」的常见原因"
        Case 2:     CalcModeText = "除模拟运算表外自动"
        Case Else:  CalcModeText = "未知"
    End Select
    If Len(CalcModeText) = 0 Then CalcModeText = "(取不到)"
    Err.Clear
    On Error GoTo 0
End Function
