'==============================================================================
' 帮助侧边栏。
'
' 【这不是 Office 原生的任务窗格】。真正的任务窗格（CustomTaskPane）只有
' COM 加载项能创建——它通过 ICTPFactory 交给实现了 IDTExtensibility2 的
' COM 加载项，纯 VBA 的 .xlam 拿不到。所以这里是一个无模式窗体，
' 贴在 Excel 窗口右缘：能常驻、能边看边操作表格，但【不会】像原生窗格
' 那样把表格区域挤窄，它浮在上层。UI 文案按实际情况写，不自称"窗格"。
'
' 【控件全部在运行时创建】，布局就是下面这段代码，git 里看得见、评审得了。
' 代价是事件要靠 clsPaneCtl 包一层（运行时控件的事件接不到窗体代码模块上）。
'
' 【本窗体不含任何业务逻辑】：目录、正文、体检报告全部来自 modHelp，
' 那些函数在无头测试里被直接断言过。窗体只是它们的显示外壳——
' 否则又是一块测试覆盖不到的代码。
'==============================================================================
Option Explicit

Private Const PANE_WIDTH As Single = 330

Private mCtls As Collection            ' 必须留住引用，否则事件静默失效
Private mGroups As Variant             ' 每项 "groupId|标题"
Private mItems As Variant              ' 每项 "entryId|显示名"

Private mLstGroups As MSForms.ListBox
Private mLstItems As MSForms.ListBox
Private mTxtBody As MSForms.TextBox
Private mTxtSearch As MSForms.TextBox

Private Sub UserForm_Initialize()
    Set mCtls = New Collection
    Me.Caption = "工具箱帮助"
    BuildUi
    LoadGroups
    ShowText WelcomeText()
End Sub

'------------------------------------------------------------------------------
' 贴到 Excel 窗口右缘。
' 取不到宿主窗口尺寸时不要硬算，宁可让它按默认位置弹出来。
'------------------------------------------------------------------------------
Public Sub DockRight()
    On Error Resume Next
    Dim appW As Single, appH As Single, appL As Single, appT As Single

    appW = Application.Width
    appH = Application.Height
    appL = Application.Left
    appT = Application.Top

    If appW > 0 And appH > 0 Then
        Me.StartUpPosition = 0          ' 手动定位
        Me.Width = PANE_WIDTH
        Me.Height = appH - 60
        If Me.Height < 260 Then Me.Height = 260
        Me.Left = appL + appW - PANE_WIDTH - 16
        Me.Top = appT + 40
        Layout
    End If

    Err.Clear
    On Error GoTo 0
End Sub

'==============================================================================
' 界面
'==============================================================================
Private Sub BuildUi()
    Me.Width = PANE_WIDTH
    Me.Height = 560

    Set mTxtSearch = AddText("txtSearch", "search")
    Set mLstGroups = AddList("lstGroups", "groups")
    Set mLstItems = AddList("lstItems", "items")

    Set mTxtBody = Me.Controls.Add("Forms.TextBox.1", "txtBody", True)
    mTxtBody.MultiLine = True
    mTxtBody.WordWrap = True
    mTxtBody.ScrollBars = 2                 ' fmScrollBarsVertical
    mTxtBody.Locked = True                  ' 只读，但仍可选中复制
    mTxtBody.BackColor = &H80000005
    mTxtBody.SpecialEffect = 2

    AddButton "btnSearch", "search", "搜索"
    AddButton "btnEnv", "env", "检查我的环境"
    AddButton "btnHtml", "html", "在浏览器打开完整帮助"

    Layout
End Sub

Private Function AddList(ByVal ctlName As String, ByVal tagName As String) As MSForms.ListBox
    Dim c As MSForms.ListBox
    Set c = Me.Controls.Add("Forms.ListBox.1", ctlName, True)

    Dim w As clsPaneCtl
    Set w = New clsPaneCtl
    w.BindList c, Me, tagName
    mCtls.Add w

    Set AddList = c
End Function

Private Function AddText(ByVal ctlName As String, ByVal tagName As String) As MSForms.TextBox
    Dim c As MSForms.TextBox
    Set c = Me.Controls.Add("Forms.TextBox.1", ctlName, True)

    Dim w As clsPaneCtl
    Set w = New clsPaneCtl
    w.BindText c, Me, tagName
    mCtls.Add w

    Set AddText = c
End Function

Private Sub AddButton(ByVal ctlName As String, ByVal tagName As String, ByVal caption As String)
    Dim c As MSForms.CommandButton
    Set c = Me.Controls.Add("Forms.CommandButton.1", ctlName, True)
    c.caption = caption

    Dim w As clsPaneCtl
    Set w = New clsPaneCtl
    w.BindButton c, Me, tagName
    mCtls.Add w
End Sub

'------------------------------------------------------------------------------
' 按当前窗体尺寸排布。窗体被拉伸后重排，所以单独成一个过程。
'------------------------------------------------------------------------------
Private Sub Layout()
    On Error Resume Next

    Dim m As Single, w As Single, y As Single
    m = 8
    w = Me.InsideWidth - m * 2
    If w < 120 Then Exit Sub
    y = m

    Me.Controls("txtSearch").Move m, y, w - 60, 20
    Me.Controls("btnSearch").Move m + w - 56, y, 56, 20
    y = y + 26

    Me.Controls("lstGroups").Move m, y, w, 88
    y = y + 92

    Me.Controls("lstItems").Move m, y, w, 108
    y = y + 112

    Dim bodyH As Single
    bodyH = Me.InsideHeight - y - 60
    If bodyH < 80 Then bodyH = 80
    Me.Controls("txtBody").Move m, y, w, bodyH
    y = y + bodyH + 8

    Me.Controls("btnEnv").Move m, y, 120, 22
    Me.Controls("btnHtml").Move m + 128, y, w - 128, 22

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    Layout
End Sub

'==============================================================================
' 数据装填 —— 内容全部来自 modHelp
'==============================================================================
Private Sub LoadGroups()
    mGroups = Split(modHelp.CatalogGroups(), vbLf)

    mLstGroups.Clear
    Dim i As Long
    For i = LBound(mGroups) To UBound(mGroups)
        mLstGroups.AddItem PartAfterBar(CStr(mGroups(i)))
    Next i

    If mLstGroups.ListCount > 0 Then mLstGroups.ListIndex = 0
End Sub

Private Sub LoadItems(ByVal groupId As String)
    Dim raw As String
    raw = modHelp.CatalogItems(groupId)

    mLstItems.Clear
    If Len(raw) = 0 Then
        mItems = Split("", vbLf)
        Exit Sub
    End If

    mItems = Split(raw, vbLf)
    Dim i As Long
    For i = LBound(mItems) To UBound(mItems)
        mLstItems.AddItem PartAfterBar(CStr(mItems(i)))
    Next i
End Sub

'==============================================================================
' 事件入口（clsPaneCtl 回调到这里）
'==============================================================================
Public Sub OnPaneEvent(ByVal tagName As String)
    On Error GoTo Failed

    Select Case tagName
        Case "groups"
            If mLstGroups.ListIndex < 0 Then Exit Sub
            LoadItems PartBeforeBar(CStr(mGroups(mLstGroups.ListIndex)))
            ShowText "← 在上面选一个功能，这里会显示它的说明和示例。"

        Case "items"
            If mLstItems.ListIndex < 0 Then Exit Sub
            ShowText modHelp.RenderEntry(PartBeforeBar(CStr(mItems(mLstItems.ListIndex))))

        Case "search"
            DoSearch

        Case "env"
            ShowText modHelp.EnvReport()

        Case "html"
            ' 完整 HTML 保留下来：要打印、要全文检索的时候它更合适
            modHelp.ShowAll
    End Select
    Exit Sub

Failed:
    ' 【侧边栏出错绝不能弹框打断用户】——他可能正在编辑单元格。
    ShowText "这一步出错了：" & Err.Description & vbCrLf & vbCrLf & _
             "可以关掉侧边栏重新打开。如果一直这样，请把这句话报给 IT。"
End Sub

'------------------------------------------------------------------------------
' 搜索：只定位到帮助条目，【绝不执行命令】。
'
' 功能区那个搜索框当初就是因为"只命中一条就直接执行"被改掉的——
' 注册表里有删除重复值、批量改名这类会动数据的命令，
' 在输入框打几个字一回车数据就被改了。这里同样只带你去看说明。
'------------------------------------------------------------------------------
Private Sub DoSearch()
    Dim q As String
    q = Trim$(mTxtSearch.Text)
    If Len(q) = 0 Then Exit Sub

    Dim hits As String
    hits = modHelp.Resolve(q)

    If Len(hits) = 0 Then
        ShowText "没找到和「" & q & "」相关的功能。" & vbCrLf & vbCrLf & _
                 "试试换个说法，按你遇到的【问题】搜，比如：" & vbCrLf & _
                 "　　求和是0　／　匹配不上　／　合并文件　／　每次都弹警告"
        Exit Sub
    End If

    Dim ids As Variant
    ids = Split(hits, vbLf)

    If UBound(ids) = LBound(ids) Then
        ShowText modHelp.RenderEntry(CStr(ids(LBound(ids))))
        Exit Sub
    End If

    Dim s As String, i As Long
    s = "「" & q & "」找到 " & (UBound(ids) - LBound(ids) + 1) & " 个相关功能：" & vbCrLf & _
        String$(28, "-") & vbCrLf & vbCrLf
    For i = LBound(ids) To UBound(ids)
        s = s & "· " & modAction.ActionLabel(CStr(ids(i))) & vbCrLf
    Next i
    s = s & vbCrLf & "在上面的目录里点开对应的功能可以看详细说明。"
    ShowText s
End Sub

'==============================================================================
' 小工具
'==============================================================================
'==============================================================================
' 定位到某个条目（功能区「帮助」带 actionId 进来、报错框里的"查看帮助"都走这儿）
'==============================================================================
Public Sub ShowEntry(ByVal entryId As String)
    On Error Resume Next

    Dim gid As String
    If LCase$(Left$(entryId, 6)) = "guide." Then
        gid = "guide"
    Else
        gid = modHelp.GroupOfAction(entryId)
    End If

    If Len(gid) = 0 Then Exit Sub
    SelectGroupById gid
    SelectItemById entryId

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SelectGroupById(ByVal gid As String)
    Dim i As Long
    For i = LBound(mGroups) To UBound(mGroups)
        If PartBeforeBar(CStr(mGroups(i))) = gid Then
            mLstGroups.ListIndex = i
            ' 【显式调用 LoadItems】：不能指望"设了 ListIndex 就会触发 Click"，
            ' 那取决于控件实现，靠它等于把正确性寄托在一个没写进文档的行为上。
            LoadItems gid
            Exit Sub
        End If
    Next i
End Sub

Private Sub SelectItemById(ByVal entryId As String)
    If IsEmpty(mItems) Then Exit Sub

    Dim i As Long
    For i = LBound(mItems) To UBound(mItems)
        If PartBeforeBar(CStr(mItems(i))) = entryId Then
            mLstItems.ListIndex = i
            ShowText modHelp.RenderEntry(entryId)
            Exit Sub
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' 把整条交互链跑一遍，返回一行可断言的结果。
'
' 【为什么需要它】：只证明"控件建得出来"是不够的——真正会坏的是
' 事件接线（运行时控件的事件接不到窗体代码模块上，漏了 clsPaneCtl
' 那一层的表现是"点了没反应"，而且不报错）。
' 这里直接走 OnPaneEvent，把点分组、点功能、搜索、体检、定位全跑一遍。
'------------------------------------------------------------------------------
Public Function PaneSelfTest() As String
    On Error GoTo Failed

    Dim r As String

    ' 1) 点分组 -> 功能列表要被填上
    mLstGroups.ListIndex = 0
    OnPaneEvent "groups"
    r = "groups=" & mLstGroups.ListCount & "|items=" & mLstItems.ListCount

    ' 2) 点功能 -> 正文要出来
    If mLstItems.ListCount > 0 Then
        mLstItems.ListIndex = 0
        OnPaneEvent "items"
    End If
    r = r & "|bodyLen=" & Len(mTxtBody.Text)

    ' 3) 搜索 -> 按症状能找到，且【不执行任何命令】
    mTxtSearch.Text = "求和是0"
    OnPaneEvent "search"
    r = r & "|search=" & CStr(InStr(mTxtBody.Text, "文本转数值") > 0)

    ' 4) 体检
    OnPaneEvent "env"
    r = r & "|env=" & CStr(InStr(mTxtBody.Text, "环境检查结果") > 0)

    ' 5) 定位到指定条目（跨分组）
    ShowEntry "misc.parseId"
    r = r & "|entry=" & CStr(InStr(mTxtBody.Text, "身份证") > 0)

    ' 6) 定位到使用配置条目
    ShowEntry "guide.macroTrust"
    r = r & "|guide=" & CStr(InStr(mTxtBody.Text, "宏被禁用") > 0)

    ' 7) 【事件接线探针】
    ' 上面几步都是直接调 OnPaneEvent，测的是处理逻辑，
    ' 并【没有】测 clsPaneCtl 那层 WithEvents 到底接上没有。
    ' 这里只改 ListIndex、不手工调任何东西，看功能列表会不会自己变：
    ' 变了说明 Click 事件真的传到了窗体，没变就是接线断了
    ' （那正是"点了没反应且不报错"的故障形态）。
    mLstGroups.ListIndex = 0
    mLstItems.Clear
    mLstGroups.ListIndex = 2
    r = r & "|wired=" & CStr(mLstItems.ListCount > 0)

    PaneSelfTest = "OK|" & r
    Exit Function

Failed:
    PaneSelfTest = "ERR|" & Err.Number & "|" & Err.Description
End Function

Private Sub UserForm_Terminate()
    ' 事件包装对象的引用显式放掉，别指望回收时机
    Set mCtls = Nothing
End Sub

Private Sub ShowText(ByVal s As String)
    On Error Resume Next
    mTxtBody.Text = s
    mTxtBody.CurLine = 0            ' 滚回顶部，否则换条目后还停在上一条的位置
    Err.Clear
    On Error GoTo 0
End Sub

Private Function PartBeforeBar(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "|")
    If p > 0 Then PartBeforeBar = Left$(s, p - 1) Else PartBeforeBar = s
End Function

Private Function PartAfterBar(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "|")
    If p > 0 Then PartAfterBar = Mid$(s, p + 1) Else PartAfterBar = s
End Function

Private Function WelcomeText() As String
    WelcomeText = _
        "工具箱帮助" & vbCrLf & String$(28, "=") & vbCrLf & vbCrLf & _
        "左边选分组 → 选功能，这里会显示：" & vbCrLf & _
        "　　什么时候用　／　怎么用　／　示例　／　注意" & vbCrLf & vbCrLf & _
        "【第一次用，或者装完没反应】" & vbCrLf & _
        "先看最上面那组「使用前必读（Excel 配置）」，" & vbCrLf & _
        "宏被禁用、每次弹安全警告、看不到选项卡都在里面。" & vbCrLf & vbCrLf & _
        "【不知道该用哪个功能】" & vbCrLf & _
        "在上面的框里按你遇到的问题搜，比如「求和是0」。" & vbCrLf & _
        "搜索只会带你看说明，不会动你的数据。" & vbCrLf & vbCrLf & _
        "这个窗口可以一直开着，不影响你操作表格。"
End Function
