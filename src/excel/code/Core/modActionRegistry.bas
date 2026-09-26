Attribute VB_Name = "modActionRegistry"
'==============================================================================
' modActionRegistry - Excel 命令注册表与派发（modAction 的宿主专属部分）
'
' 【这是从原来的单文件 modAction.bas 拆出来的】：shared/code/Core/modAction.bas
' 现在只放 RunAction 那套宿主无关的管线逻辑，真正碰 Excel 对象
' （Selection/ActiveWorkbook/ActiveSheet）的 RegisterAll/Dispatch，以及
' IsActionEnabled/IsActionPressed 里那几条 Excel 专属能力判据，都搬到这里。
' 函数名和签名【一字未改】，RunAction 通过不写模块前缀的方式调用它们
' （VBA 在同一工程内按名字解析），和 modHost 是同一个技巧。
'
' 新增一个工具的完整步骤：
'     1. 在 customUI14.xml 里加按钮，tag = actionId
'     2. 在 RegisterAll 里 RegisterAction 一行
'     3. 在 Dispatch 的 Select Case 里加一行，指向业务过程
' 业务过程本身不碰 ScreenUpdating、不碰撤销、不写 On Error 弹窗。
'==============================================================================
Option Explicit
Option Private Module

' 各功能模块的注册入口都挂在这里，按模块分段，便于增删。
Public Sub RegisterAll()
    ' --- Core ---
    modAction.RegisterAction "core.undoLast", "撤销", "撤销上一步工具箱操作", _
                   Undoable:=False, RequiresWorkbook:=False
    modAction.RegisterAction "core.selfTest", "自检", "验证加载宏已正确装载", _
                   Undoable:=False, RequiresWorkbook:=False
    modAction.RegisterAction "core.about", "关于", "", _
                   Undoable:=False, RequiresWorkbook:=False
    modAction.RegisterAction "core.help", "帮助", _
                   "打开使用帮助：每个功能什么时候用、怎么用、有什么坑", _
                   Undoable:=False, RequiresWorkbook:=False
    modAction.RegisterAction "core.resetEnv", "环境复位", _
                   "恢复屏幕刷新、自动重算、事件响应和状态栏。" & _
                   "宏被强行中断后如果 Excel 变得没有反应或不自动计算，点这里", _
                   Undoable:=False, RequiresWorkbook:=False

    ' --- M1 文本与单元格 ---
    modAction.RegisterAction "text.cleanSpaces", "清除空格", _
                   "去掉首尾空格、压缩中间连续空格，并清除不间断空格、全角空格、零宽字符等不可见字符", _
                   RequiresSelection:=True
    modAction.RegisterAction "text.toNumber", "文本转数值", _
                   "把文本型数字转成真正的数值，自动处理全角数字、千分位逗号和不间断空格。" & _
                   "这是求和结果为 0 的头号原因", _
                   RequiresSelection:=True
    modAction.RegisterAction "text.toUpper", "转大写", "", RequiresSelection:=True
    modAction.RegisterAction "text.toLower", "转小写", "", RequiresSelection:=True
    modAction.RegisterAction "text.toProper", "首字母大写", "", RequiresSelection:=True
    modAction.RegisterAction "text.toHalfWidth", "全角转半角", "", RequiresSelection:=True
    modAction.RegisterAction "text.toFullWidth", "半角转全角", "", RequiresSelection:=True
    modAction.RegisterAction "text.removeLineBreaks", "删除换行符", _
                   "去掉单元格内的换行符", RequiresSelection:=True
    modAction.RegisterAction "text.extractDigits", "提取数字", _
                   "只保留数字字符", RequiresSelection:=True
    modAction.RegisterAction "text.extractChinese", "提取中文", "", RequiresSelection:=True
    modAction.RegisterAction "text.extractEnglish", "提取字母", "", RequiresSelection:=True
    modAction.RegisterAction "text.addAffix", "添加前后缀", _
                   "批量给选区加前缀或后缀", RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "text.regexReplace", "正则替换", _
                   "用正则表达式批量查找替换", RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "text.splitColumn", "按分隔符拆列", _
                   "把一列按分隔符拆成多列，自动插入所需列数，不覆盖右侧数据", _
                   RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "cells.unmergeFill", "拆分并填充", _
                   "取消合并单元格，并把原值填满整个区域", RequiresSelection:=True
    modAction.RegisterAction "cells.mergeSame", "合并相同项", _
                   "把同列中相邻且内容相同的单元格合并", RequiresSelection:=True

    ' --- M2 数据处理 ---
    modAction.RegisterAction "data.deleteEmptyRows", "删除空行", _
                   "删除选区内完全为空的整行。只按选中的列判断是否为空", RequiresSelection:=True
    modAction.RegisterAction "data.deleteEmptyCols", "删除空列", _
                   "删除选区内完全为空的整列", RequiresSelection:=True
    modAction.RegisterAction "data.markDuplicates", "标记重复值", _
                   "把重复行标成浅红色，支持多列组合判重", _
                   RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "data.deleteDuplicates", "删除重复值", _
                   "删除重复行并保留首次出现，支持多列组合判重", _
                   RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "data.extractUnique", "提取唯一值", _
                   "把选区内的唯一值提取到新工作表", RequiresSelection:=True
    modAction.RegisterAction "data.compare", "两表对比", _
                   "按键列对比两个区域，输出差异报告：仅 A 有 / 仅 B 有 / 内容不同，" & _
                   "每行都能点击跳回源数据", _
                   Undoable:=False, RequiresWorkbook:=True, PromptsForInput:=True
    modAction.RegisterAction "data.unpivot", "二维转一维", _
                   "逆透视：把交叉表展开成明细表。透视表、Power Query 和数据库都需要一维明细表", _
                   RequiresSelection:=True, Undoable:=False, PromptsForInput:=True
    modAction.RegisterAction "data.transpose", "行列转置", _
                   "把选区转置后输出到新工作表", RequiresSelection:=True, Undoable:=False

    ' --- M3 工作表管理 ---
    ' 拆表/排序/重命名都会改动工作簿结构且撤不回来，一律强制确认。
    ' 「按列拆分」一次可能生成上百张表，手工删回去比出错本身还痛苦。
    modAction.RegisterAction "sheet.splitByColumn", "按列拆分工作表", _
                   "按某列的值把数据拆分成多个工作表", _
                   RequiresSelection:=True, Undoable:=False, _
                   ConfirmBeforeRun:=True, PromptsForInput:=True
    modAction.RegisterAction "sheet.mergeAll", "合并所有工作表", _
                   "把当前工作簿所有工作表合并到一张汇总表。" & _
                   "按标题名对齐，而不是按列位置——某张表少一列也不会整体错位", Undoable:=False
    modAction.RegisterAction "sheet.createIndex", "生成目录", _
                   "生成带超链接的工作表目录", Undoable:=False
    modAction.RegisterAction "sheet.sort", "工作表排序", "按名称排列工作表", _
                   Undoable:=False, ConfirmBeforeRun:=True, PromptsForInput:=True
    ' 深度隐藏的表往往是作者有意藏起来的（参数表、中间计算表），
    ' 一次全部放出来之后没有记录能还原回去——所以也要确认
    modAction.RegisterAction "sheet.showAll", "显示所有表", _
                   "显示全部隐藏工作表（含深度隐藏）。" & _
                   "哪些表原来是隐藏的不会被记录下来，之后无法一键还原", _
                   Undoable:=False, ConfirmBeforeRun:=True
    modAction.RegisterAction "sheet.batchRename", "批量重命名表", _
                   "按选中的名称列表批量重命名工作表", _
                   RequiresSelection:=True, Undoable:=False, ConfirmBeforeRun:=True

    ' --- M4 多文件合并 ---
    modAction.RegisterAction "merge.folder", "合并文件夹", _
                   "合并一个文件夹内所有 Excel 文件的数据，结果带来源文件和来源工作表列。" & _
                   "单个文件失败不会中断整批，最后出失败清单", _
                   Undoable:=False, RequiresWorkbook:=False, PromptsForInput:=True

    ' --- M5 文件批处理（全部不可撤销，强制确认）---
    modAction.RegisterAction "file.list", "文件清单", _
                   "把文件夹内的文件清单导入工作表，并生成可填写的新文件名列", _
                   Undoable:=False, RequiresWorkbook:=True, PromptsForInput:=True
    modAction.RegisterAction "file.batchRename", "批量重命名", _
                   "按当前「文件清单」表的 F 列批量重命名文件。先全表校验，全部通过才动手", _
                   Undoable:=False, ConfirmBeforeRun:=True
    modAction.RegisterAction "file.exportSheets", "导出工作表", _
                   "把每张工作表导出为独立的 xlsx / CSV / PDF 文件", _
                   Undoable:=False, ConfirmBeforeRun:=True, PromptsForInput:=True
    ' 插进去的图片是浮动对象，撤不回来，手工一张张删很痛苦——必须确认
    modAction.RegisterAction "file.insertImages", "批量插图", _
                   "按单元格内容在指定文件夹里找同名图片并插入到右侧单元格", _
                   Undoable:=False, RequiresSelection:=True, _
                   ConfirmBeforeRun:=True, PromptsForInput:=True

    ' --- M6 公式与引用 ---
    modAction.RegisterAction "formula.toValues", "公式转值", _
                   "把选区内的公式替换为计算结果", RequiresSelection:=True
    modAction.RegisterAction "formula.findErrors", "定位错误值", _
                   "找出并标黄选区内的所有错误值", RequiresSelection:=True
    modAction.RegisterAction "formula.wrapIfError", "套用 IFERROR", _
                   "给选区内的公式批量加上 IFERROR 容错", _
                   RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "formula.breakLinks", "断开外部链接", _
                   "把所有外部链接公式转为当前值", _
                   Undoable:=False, ConfirmBeforeRun:=True
    modAction.RegisterAction "formula.cleanNames", "清理失效名称", _
                   "删除指向 #REF! 的已定义名称", Undoable:=False, ConfirmBeforeRun:=True
    modAction.RegisterAction "formula.toggleView", "显示公式", _
                   "在显示公式和显示结果之间切换", Undoable:=False
    modAction.RegisterAction "formula.cleanRules", "清理格式规则", _
                   "清除选区内堆叠的条件格式和数据验证", RequiresSelection:=True

    ' --- M7 数据体检 ---
    modAction.RegisterAction "audit.scan", "数据体检", _
                   "扫描当前工作表，列出空行、文本型数字、文本型日期、错误值、合并单元格等问题，可点击跳转", _
                   Undoable:=False
    modAction.RegisterAction "audit.quickClean", "一键清洗", _
                   "修掉最常见且无歧义的几类问题：清理空白字符、文本型数字转数值、删除空行。" & _
                   "合并单元格和错误值需要人工判断，不会自动改", RequiresSelection:=True

    ' --- M8 数据可视化 ---
    modAction.RegisterAction "viz.dataBars", "数据条", "给数值单元格添加数据条", RequiresSelection:=True
    modAction.RegisterAction "viz.colorScale", "色阶热力图", "三色色阶", RequiresSelection:=True
    modAction.RegisterAction "viz.iconSet", "图标集", "三色交通灯图标集", RequiresSelection:=True
    modAction.RegisterAction "viz.clearCF", "清除条件格式", "清除选区内所有条件格式", RequiresSelection:=True
    modAction.RegisterAction "viz.sparklines", "批量迷你图", _
                   "每行生成一个迷你图，放在数据右侧一列", _
                   RequiresSelection:=True, PromptsForInput:=True
    modAction.RegisterAction "viz.quickChart", "快速图表", _
                   "按选区生成图表并套用统一格式", _
                   RequiresSelection:=True, Undoable:=False, PromptsForInput:=True
    ' 会覆盖用户手工调好的图表格式，且撤不回来
    modAction.RegisterAction "viz.unifyCharts", "统一图表格式", _
                   "把当前工作表所有图表的格式统一，会覆盖你手工调过的格式", _
                   Undoable:=False, ConfirmBeforeRun:=True

    ' --- M9 辅助增强 ---
    modAction.RegisterAction "misc.spotlight", "聚光灯", _
                   "高亮光标所在的整行整列。用条件格式实现，不会破坏原有底色", _
                   Undoable:=False
    modAction.RegisterAction "misc.amountToChinese", "金额大写", _
                   "把选中的一列金额转成人民币中文大写，写到右侧一列", RequiresSelection:=True
    modAction.RegisterAction "misc.parseId", "身份证解析", _
                   "解析出生日期、性别、年龄，并校验校验位", RequiresSelection:=True
    modAction.RegisterAction "misc.normalizeDates", "日期规范化", _
                   "把 20240115 / 2024.1.15 / 2024年1月15日 等文本统一转成日期值", _
                   RequiresSelection:=True
End Sub

'------------------------------------------------------------------------------
' actionId -> 业务过程。这是全加载宏唯一的转派点。
'
' 返回值是给用户看的结果文字（"已删除 3 个空行"），由 RunAction 统一呈现。
' 业务过程不许自己弹框，否则批量执行和自动化测试都会被卡住。
'------------------------------------------------------------------------------
Public Function Dispatch(ByVal actionId As String) As String
    Select Case actionId
        ' --- Core ---
        Case "core.undoLast":  modHost.Host_UndoLast
        Case "core.resetEnv":  modHost.Host_FastModeReset
        Case "core.selfTest":  Dispatch = modApp.AboutText() & vbCrLf & vbCrLf & "加载宏工作正常。"
        Case "core.about":     Dispatch = modApp.AboutText()
        Case "core.help":      Dispatch = modHelp.ShowPane()

        ' --- M1 文本与单元格 ---
        Case "text.cleanSpaces":      Dispatch = modText.CleanSpaces(Selection)
        Case "text.toNumber":         Dispatch = modText.TextToNumber(Selection)
        Case "text.toUpper":          Dispatch = modText.ToUpper(Selection)
        Case "text.toLower":          Dispatch = modText.ToLower(Selection)
        Case "text.toProper":         Dispatch = modText.ToProper(Selection)
        Case "text.toHalfWidth":      Dispatch = modText.ToHalfWidth(Selection)
        Case "text.toFullWidth":      Dispatch = modText.ToFullWidth(Selection)
        Case "text.removeLineBreaks": Dispatch = modText.RemoveLineBreaks(Selection)
        Case "text.extractDigits":    Dispatch = modText.ExtractDigits(Selection)
        Case "text.extractChinese":   Dispatch = modText.ExtractChinese(Selection)
        Case "text.extractEnglish":   Dispatch = modText.ExtractEnglish(Selection)
        Case "text.addAffix":         Dispatch = modText.AddAffix(Selection)
        Case "text.regexReplace":     Dispatch = modText.RegexReplaceCells(Selection)
        Case "text.splitColumn":      Dispatch = modText.SplitColumn(Selection)
        Case "cells.unmergeFill":     Dispatch = modCells.UnmergeAndFill(Selection)
        Case "cells.mergeSame":       Dispatch = modCells.MergeSameValues(Selection)

        ' --- M2 数据处理 ---
        Case "data.deleteEmptyRows":  Dispatch = modRows.DeleteEmptyRows(Selection)
        Case "data.deleteEmptyCols":  Dispatch = modRows.DeleteEmptyColumns(Selection)
        Case "data.markDuplicates":   Dispatch = modDedupe.MarkDuplicates(Selection)
        Case "data.deleteDuplicates": Dispatch = modDedupe.DeleteDuplicates(Selection)
        Case "data.extractUnique":    Dispatch = modDedupe.ExtractUnique(Selection)
        Case "data.compare":          Dispatch = modCompare.CompareRanges()
        Case "data.unpivot":          Dispatch = modReshape.Unpivot(Selection)
        Case "data.transpose":        Dispatch = modReshape.TransposeRange(Selection)

        ' --- M3 工作表管理 ---
        Case "sheet.splitByColumn":   Dispatch = modSheets.SplitByColumn(Selection)
        Case "sheet.mergeAll":        Dispatch = modSheets.MergeAllSheets(ActiveWorkbook)
        Case "sheet.createIndex":     Dispatch = modSheets.CreateIndex(ActiveWorkbook)
        Case "sheet.sort":            Dispatch = modSheets.SortSheets(ActiveWorkbook)
        Case "sheet.showAll":         Dispatch = modSheets.ShowAllSheets(ActiveWorkbook)
        Case "sheet.batchRename":     Dispatch = modSheets.BatchRename(Selection)

        ' --- M4 多文件合并 ---
        Case "merge.folder":          Dispatch = modMergeFiles.MergeFolder()

        ' --- M5 文件批处理 ---
        Case "file.list":             Dispatch = modFileBatch.ListFiles()
        Case "file.batchRename":      Dispatch = modFileBatch.BatchRenameFiles(ActiveSheet)
        Case "file.exportSheets":     Dispatch = modFileBatch.ExportSheets(ActiveWorkbook)
        Case "file.insertImages":     Dispatch = modFileBatch.InsertImages(Selection)

        ' --- M6 公式与引用 ---
        Case "formula.toValues":      Dispatch = modFormula.FormulasToValues(Selection)
        Case "formula.findErrors":    Dispatch = modFormula.FindErrors(Selection)
        Case "formula.wrapIfError":   Dispatch = modFormula.WrapWithIfError(Selection)
        Case "formula.breakLinks":    Dispatch = modFormula.BreakExternalLinks(ActiveWorkbook)
        Case "formula.cleanNames":    Dispatch = modFormula.CleanBrokenNames(ActiveWorkbook)
        Case "formula.toggleView":    Dispatch = modFormula.ToggleFormulaView(ActiveSheet)
        Case "formula.cleanRules":    Dispatch = modFormula.CleanFormatRules(Selection)

        ' --- M7 数据体检 ---
        Case "audit.scan":            Dispatch = modAudit.ScanSheet(ActiveSheet)
        Case "audit.quickClean":      Dispatch = modAudit.QuickClean(Selection)

        ' --- M8 数据可视化 ---
        Case "viz.dataBars":          Dispatch = modViz.AddDataBars(Selection)
        Case "viz.colorScale":        Dispatch = modViz.AddColorScale(Selection)
        Case "viz.iconSet":           Dispatch = modViz.AddIconSet(Selection)
        Case "viz.clearCF":           Dispatch = modViz.ClearConditionalFormats(Selection)
        Case "viz.sparklines":        Dispatch = modViz.AddSparklines(Selection)
        Case "viz.quickChart":        Dispatch = modViz.QuickChart(Selection)
        Case "viz.unifyCharts":       Dispatch = modViz.UnifyCharts(ActiveSheet)

        ' --- M9 辅助增强 ---
        Case "misc.spotlight":        Dispatch = modSpotlight.Toggle()
        Case "misc.amountToChinese":  Dispatch = modUtils.AmountColumnToChinese(Selection)
        Case "misc.parseId":          Dispatch = modUtils.ParseIdCards(Selection)
        Case "misc.normalizeDates":   Dispatch = modUtils.NormalizeDates(Selection)

        Case Else
            Err.Raise vbObjectError + 1, "modActionRegistry.Dispatch", _
                      "actionId 已注册但未实现转派：" & actionId
    End Select
End Function

'------------------------------------------------------------------------------
' modAction.IsActionEnabled 对 core.undoLast 之外的按钮转派到这里。
' 宿主能力探测：不支持的命令直接灰显，而不是让用户点了才看到报错。
' 这比在 clsActionDef 上人工维护一张"哪个宿主支持哪个 API"的表可靠——
' 那张表在开发机上根本没法验证，事实上也一直是空的。
'------------------------------------------------------------------------------
Public Function ActionExtraEnabled(ByVal actionId As String) As Boolean
    Select Case actionId
        Case "viz.sparklines"
            ActionExtraEnabled = modCaps.SupportsSparklines()

        Case "merge.folder", "file.list", "file.exportSheets", "file.insertImages"
            ActionExtraEnabled = modCaps.SupportsFileDialog()

        Case Else
            ActionExtraEnabled = True
    End Select
End Function

' modAction.IsActionPressed 转派到这里——开关型按钮的按下状态。
Public Function ActionExtraPressed(ByVal actionId As String) As Boolean
    Select Case actionId
        Case "misc.spotlight": ActionExtraPressed = modSpotlight.IsEnabled()
        Case Else:             ActionExtraPressed = False
    End Select
End Function
