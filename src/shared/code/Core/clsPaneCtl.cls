VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPaneCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' 帮助侧边栏的控件事件包装。
'
' 【为什么需要这个类】：用 Controls.Add 在运行时创建的控件，
' 它们的事件【不会】自动接到窗体的代码模块上——窗体里写
' Private Sub lstGroups_Click() 根本不会被调用，而且不报错，
' 表现就是"点了没反应"。只有声明成 WithEvents 的对象变量才收得到事件。
'
' 【为什么不干脆在 VBE 里把窗体画出来】：那样控件布局会存进同名的
' .frx 二进制资源，git 看不出任何差别，评审也无从审起。
' 本仓库的源码一律要可 diff，所以宁可多写这个包装类，
' 把布局留在代码里。
'
' 每个控件配一个本类的实例，实例存在窗体的集合里——
' 【必须留引用】，否则对象一被回收，事件就静默失效了。
'==============================================================================
Option Explicit

Public WithEvents Lst As MSForms.ListBox
Attribute Lst.VB_VarHelpID = -1
Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public WithEvents Txt As MSForms.TextBox
Attribute Txt.VB_VarHelpID = -1

Private mPane As Object        ' frmHelpPane，用 Object 避免和窗体互相编译依赖
Private mTag As String

Public Sub BindList(ByVal ctl As Object, ByVal pane As Object, ByVal tagName As String)
    Set Lst = ctl
    Set mPane = pane
    mTag = tagName
End Sub

Public Sub BindButton(ByVal ctl As Object, ByVal pane As Object, ByVal tagName As String)
    Set Btn = ctl
    Set mPane = pane
    mTag = tagName
End Sub

Public Sub BindText(ByVal ctl As Object, ByVal pane As Object, ByVal tagName As String)
    Set Txt = ctl
    Set mPane = pane
    mTag = tagName
End Sub

'------------------------------------------------------------------------------
' 自述：绑的是哪个控件、挂的什么 tag。
'
' 【给测试用】。只数包装对象的个数是不够的——数量对了，仍然可能
' 绑错了控件、挂错了 tag，或者某一个控件根本没绑上。
' 那几种情况的表现都是"点某个按钮没反应"，而且不报错。
' 这里把绑定关系吐出来，测试逐条比对。
'------------------------------------------------------------------------------
Public Function Describe() As String
    On Error Resume Next
    If Not Lst Is Nothing Then
        Describe = mTag & ":" & Lst.Name & ":list"
    ElseIf Not Btn Is Nothing Then
        Describe = mTag & ":" & Btn.Name & ":button"
    ElseIf Not Txt Is Nothing Then
        Describe = mTag & ":" & Txt.Name & ":text"
    Else
        Describe = mTag & ":(未绑定):none"
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub Lst_Click()
    If mPane Is Nothing Then Exit Sub
    mPane.OnPaneEvent mTag
End Sub

Private Sub Btn_Click()
    If mPane Is Nothing Then Exit Sub
    mPane.OnPaneEvent mTag
End Sub

' 搜索框按回车即搜索。用 KeyDown 而不是 Change：
' 每敲一个字就搜一次既慢又会把正在看的内容冲掉。
Private Sub Txt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0                 ' 吃掉回车，免得窗体发出提示音
        If mPane Is Nothing Then Exit Sub
        mPane.OnPaneEvent mTag
    End If
End Sub
