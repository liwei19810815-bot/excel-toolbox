Attribute VB_Name = "modIO"
'==============================================================================
' modIO - 文件与工作簿 IO
'
' 批量处理外部文件时，"打开别人的工作簿"这一步藏着一堆坑，全部在这里统一处理：
'   - 对方文件带外部链接 -> 弹"是否更新链接"对话框，批处理直接卡死
'   - 对方文件带宏        -> 弹安全警告；而且我们只是读数据，没必要执行它的宏
'   - 对方文件已被占用    -> 弹"以只读方式打开？"
'   - 对方文件有打开密码  -> 弹密码框
' 一律用 Open 的完整参数把这些关掉，并以只读方式打开——批量汇总没有任何理由
' 去写别人的源文件。
'==============================================================================
Option Explicit
Option Private Module

'==============================================================================
' MSO（Office 对象库）常量的本地副本。
'
' 【绝对不要直接写 msoXxx 常量，也不要用 MsoXxx / FileDialog 这类 MSO 类型声明】。
'
' 它们不属于 Excel 类型库，而属于 Office 共享库（MSO）。只要目标宿主的 VBA 工程
' 没有引用 MSO 库，这些符号在编译期就解析不了——后果和 xlCSVUTF8 那次一模一样：
' 不是"文件夹对话框用不了"，而是【整个工程编译不过、59 个命令一个都加载不了】。
'
' 尤其注意：modCaps 的运行时探测【救不了这件事】。探测是运行期行为，
' 而符号解析发生在编译期——工程根本轮不到运行起来。
'
' 值来自 MSO 类型库的枚举定义：
'   MsoFileDialogType:     Open=1, SaveAs=2, FilePicker=3, FolderPicker=4
'   MsoAutomationSecurity: Low=1, ByUI=2, ForceDisable=3
'   MsoTriState:           msoTrue=-1, msoFalse=0, msoCTrue=1
'
' 【msoTrue / msoFalse 同样是 MSO 符号】，很容易漏——它们看起来像语言内置的
' True/False，实际不是，而且取值也不一样（msoTrue = -1，msoFalse = 0）。
' 用在 Shapes.AddPicture、Format.Line.Visible 这类地方。
'==============================================================================
Public Const MSO_FILEDIALOG_FOLDERPICKER As Long = 4
Public Const MSO_AUTOMATIONSECURITY_FORCEDISABLE As Long = 3
Public Const MSO_TRUE As Long = -1
Public Const MSO_FALSE As Long = 0

'------------------------------------------------------------------------------
' 静默打开一个工作簿。打不开就返回 Nothing，由调用方汇总失败清单，
' 而不是中断整个批处理——批量任务里一个坏文件不该毁掉其余 99 个。
'------------------------------------------------------------------------------
Public Function OpenQuiet(ByVal filePath As String) As Workbook
    ' 【安全】强制禁用被打开文件里的宏。
    '
    ' 只设 Open 的参数是不够的：一个带 Workbook_Open 或 Auto_Open 的 .xlsm
    ' 被批量合并扫到时，它的宏会在我们的进程里直接执行。我们只是来读数据的，
    ' 没有任何理由运行别人文件里的代码——内网共享目录里混进一个带宏的文件，
    ' 这就成了一条现成的执行通道。
    '
    ' 【读取和设置都必须包在错误保护里】。宿主不支持这个属性时读它就会抛错，
    ' 而这两行原本在 On Error Resume Next 之前，异常会一路冒到 RunAction，
    ' 表现成"合并文件夹整个失败"，而不是"这台机器上这个功能本就该停用"。
    Dim prevSecurity As Long
    Dim securityApplied As Boolean

    On Error Resume Next
    prevSecurity = Application.AutomationSecurity
    If Err.Number = 0 Then
        Application.AutomationSecurity = MSO_AUTOMATIONSECURITY_FORCEDISABLE
        ' 【设了不等于生效】——回读确认。拿"我设过了"当保证，
        ' 正是这类防护最常见的失效方式。
        If Err.Number = 0 Then
            securityApplied = (Application.AutomationSecurity = _
                               MSO_AUTOMATIONSECURITY_FORCEDISABLE)
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' 关不掉宏就【不打开】。宁可这个文件算作失败进汇总清单，
    ' 也不能在没有防护的情况下把可能带 Auto_Open 的文件加载进本进程。
    If Not securityApplied Then
        Set OpenQuiet = Nothing
        Exit Function
    End If

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.Open( _
                Filename:=filePath, _
                UpdateLinks:=0, _
                ReadOnly:=True, _
                AddToMru:=False, _
                Password:=vbNullString, _
                WriteResPassword:=vbNullString, _
                IgnoreReadOnlyRecommended:=True, _
                Notify:=False, _
                CorruptLoad:=xlNormalLoad)
    On Error GoTo 0

    ' 还原必须无条件执行：留在 ForceDisable 状态会让用户之后正常打开的
    ' 带宏文件也一并失效，而且毫无提示
    On Error Resume Next
    Application.AutomationSecurity = prevSecurity
    On Error GoTo 0

    Set OpenQuiet = wb
End Function

Public Sub CloseQuiet(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' 让用户选一个文件夹
'------------------------------------------------------------------------------
Public Function PickFolder(ByVal title As String) As String
    ' dlg 必须声明为 Object：FileDialog 是 MSO 库的类型，写成强类型就又成了
    ' 编译期依赖（见本模块顶部的说明）
    Dim dlg As Object
    Set dlg = Application.FileDialog(MSO_FILEDIALOG_FOLDERPICKER)
    dlg.title = title
    dlg.AllowMultiSelect = False
    If dlg.Show <> -1 Then modPrompt.Cancel
    PickFolder = dlg.SelectedItems(1)
End Function

'------------------------------------------------------------------------------
' 列出文件夹内的 Excel 文件。
'
' 用 Dir 递归会出问题：Dir 是全局状态，在递归里被内层调用重置后，
' 外层的遍历就串了。所以这里用 FileSystemObject。
'------------------------------------------------------------------------------
Public Function ListExcelFiles(ByVal folderPath As String, _
                               ByVal recursive As Boolean) As Collection
    Dim result As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Err.Raise vbObjectError + 380, "modIO", "文件夹不存在：" & folderPath
    End If

    CollectFiles fso, fso.GetFolder(folderPath), recursive, result
    Set ListExcelFiles = result
End Function

Private Sub CollectFiles(ByVal fso As Object, ByVal folderObj As Object, _
                         ByVal recursive As Boolean, ByVal result As Collection)
    Dim fileObj As Object, ext As String
    For Each fileObj In folderObj.Files
        ext = LCase$(fso.GetExtensionName(fileObj.Path))
        ' ~$ 开头的是 Excel 的锁文件，不是真实数据文件
        If IsExcelExt(ext) And Left$(fso.GetFileName(fileObj.Path), 2) <> "~$" Then
            result.Add fileObj.Path
        End If
    Next fileObj

    If recursive Then
        Dim subFolder As Object
        For Each subFolder In folderObj.SubFolders
            CollectFiles fso, subFolder, True, result
        Next subFolder
    End If
End Sub

Private Function IsExcelExt(ByVal ext As String) As Boolean
    Select Case ext
        Case "xlsx", "xlsm", "xls", "xlsb", "csv": IsExcelExt = True
        Case Else: IsExcelExt = False
    End Select
End Function

'------------------------------------------------------------------------------
' 路径工具
'------------------------------------------------------------------------------
Public Function FileNameOf(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileNameOf = fso.GetFileName(filePath)
End Function

Public Function BaseNameOf(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    BaseNameOf = fso.GetBaseName(filePath)
End Function

Public Function JoinPath(ByVal folderPath As String, ByVal leafName As String) As String
    If Right$(folderPath, 1) = "\" Then
        JoinPath = folderPath & leafName
    Else
        JoinPath = folderPath & "\" & leafName
    End If
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function
