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
