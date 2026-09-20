<#
.SYNOPSIS
    自动化回归测试：加载 dist\ExcelToolbox.xlam，构造数据、执行工具、断言结果。

.DESCRIPTION
    覆盖三类检查：
      1. 结构一致性 —— Ribbon 是否真的加载、按钮 tag 与注册命令是否一一对应。
         VBA 编译不校验 customUI14.xml，一处错误 Excel 就静默丢掉整个选项卡。
      2. 功能正确性 —— 每个模块至少一个用例。
      3. 撤销往返 —— 最容易出错的部分。断言比的是整块区域的完整指纹
         （值/公式/数字格式/字体/底纹/行高/列宽），只比值测不出格式被吃掉。

    全程开启静默模式：业务代码里任何一个对话框都会让无人值守测试永久挂死。
    参数化的工具通过 Toolbox_SetParam 预设参数，走的还是和交互时同一条代码路径。

    调用方必须套 timeout —— 见 README。退出码 0 = 全过。

.EXAMPLE
    timeout 300 powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"
$RibbonXml = Join-Path $RepoRoot "src\package\customUI\customUI14.xml"

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$script:pass = 0
$script:fail = 0
$script:section = ""

function Section($name) {
    $script:section = $name
    Write-Host ""
    Write-Host "== $name ==" -ForegroundColor Cyan
}

function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") {
        $script:pass++
        Write-Host "  PASS  $what" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected]" -ForegroundColor Red
        Write-Host "        实际 [$actual]" -ForegroundColor Red
    }
}

function Assert-True($condition, [string]$what) {
    Assert-Equal $true ([bool]$condition) $what
}

function Assert-Match($actual, [string]$pattern, [string]$what) {
    if ("$actual" -like $pattern) {
        $script:pass++
        Write-Host "  PASS  $what" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望匹配 [$pattern]" -ForegroundColor Red
        Write-Host "        实际     [$actual]" -ForegroundColor Red
    }
}

# 选中一块区域。
# 必须先 Activate：Range.Select 只对【活动工作表】有效，否则报
# "不能取得类 Range 的 Select 属性"。前面的工具一旦新建了工作表，
# 活动表就变了，这个坑非常容易踩。
function Select-On($ws, [string]$addr) {
    $null = $ws.Activate()
    $null = $ws.Range($addr).Select()
}

# 一次性写入一块数据。
# 不逐格 .Value2 赋值有两个原因：一是快得多，二是 PowerShell 的 COM 适配器在
# 逐格混写字符串和数字时会出现类型绑定错误（Int32 无法转换为 String）。
function Set-Grid($ws, [int]$top, [int]$left, $rows) {
    $h = $rows.Count
    $w = $rows[0].Count
    $arr = [Array]::CreateInstance([object], $h, $w)
    for ($r = 0; $r -lt $h; $r++) {
        for ($c = 0; $c -lt $w; $c++) { $arr[$r, $c] = $rows[$r][$c] }
    }
    $ws.Cells($top, $left).Resize($h, $w).Value2 = $arr
}

# 抓取一块区域的完整状态指纹。撤销测试的价值全在这里——只比值测不出格式丢失。
function Get-Fingerprint($ws, [int]$rows, [int]$cols) {
    $sb = New-Object System.Text.StringBuilder
    for ($r = 1; $r -le $rows; $r++) {
        [void]$sb.AppendLine("rowH$r=" + [math]::Round($ws.Rows($r).RowHeight, 2))
        for ($c = 1; $c -le $cols; $c++) {
            $cell = $ws.Cells($r, $c)
            [void]$sb.AppendLine("$r,$c|v=$($cell.Value2)|f=$($cell.Formula)|nf=$($cell.NumberFormat)|b=$($cell.Font.Bold)|clr=$($cell.Interior.Color)")
        }
    }
    for ($c = 1; $c -le $cols; $c++) {
        [void]$sb.AppendLine("colW$c=" + [math]::Round($ws.Columns($c).ColumnWidth, 2))
    }
    return $sb.ToString()
}

# 取单元格的日期并格式化成 yyyy-MM-dd。
# .Value2 可能返回 DateTime（PowerShell 的 COM 适配器自动转过），也可能返回
# 原始的日期序列值，两种都要能处理。直接比显示串不行——那个随区域设置变化。
function Get-CellDate($ws, [string]$addr) {
    $v = $ws.Range($addr).Value2
    if ($v -is [datetime]) { return $v.ToString("yyyy-MM-dd") }
    if ($v -is [string]) { return "仍是文本:$v" }   # 说明转换没生效，直接让断言失败
    return [datetime]::FromOADate([double]$v).ToString("yyyy-MM-dd")
}

function Show-FirstDiff($before, $after) {
    $bl = $before -split "`r?`n"
    $al = $after  -split "`r?`n"
    for ($i = 0; $i -lt [math]::Max($bl.Count, $al.Count); $i++) {
        $b = if ($i -lt $bl.Count) { $bl[$i] } else { "<缺失>" }
        $a = if ($i -lt $al.Count) { $al[$i] } else { "<缺失>" }
        if ($b -ne $a) {
            Write-Host "        首处差异：" -ForegroundColor Yellow
            Write-Host "          撤销前 $b" -ForegroundColor Yellow
            Write-Host "          撤销后 $a" -ForegroundColor Yellow
            return
        }
    }
}

$xl = $null
try {
    Write-Host "==> 启动 Excel 并加载加载宏" -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    # 保持不可见：可见模式下 Excel 会真正创建功能区和窗口，COM 调用时序变得不稳定。
    # 功能测试不需要 UI，Ribbon 是否真的加载由 tests\check-ribbon.ps1 单独验证。
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    $Run   = { param($id) $xl.Run("'$OutputName'!Toolbox_Run", $id) }
    $Param = { param($k, $v) $xl.Run("'$OutputName'!Toolbox_SetParam", $k, $v) }
    $Undo  = { $null = $xl.Run("'$OutputName'!Toolbox_Undo") }   # 返回值这里不关心

    $xl.Run("'$OutputName'!Toolbox_SetSilent", $true)

    #==========================================================================
    Section "结构一致性"

    $sc = $xl.Run("'$OutputName'!Toolbox_SelfCheck")
    Write-Host "  $sc" -ForegroundColor DarkGray
    Assert-Match $sc "OK|*" "自检返回 OK"

    # tag 与注册命令双向一致
    [xml]$ribbon = Get-Content $RibbonXml -Raw -Encoding UTF8
    $tags = @($ribbon.SelectNodes("//*[@tag]") | ForEach-Object { $_.tag }) | Sort-Object -Unique
    $registered = @(($xl.Run("'$OutputName'!Toolbox_ListActions") -split "`n") | Where-Object { $_ }) | Sort-Object -Unique

    $orphanTags = @($tags | Where-Object { $registered -notcontains $_ })
    Assert-Equal 0 $orphanTags.Count "所有按钮 tag 都已注册（孤儿：$($orphanTags -join ', ')）"

    $unexposed = @($registered | Where-Object { $tags -notcontains $_ })
    Assert-Equal 0 $unexposed.Count "所有注册命令都有按钮（未暴露：$($unexposed -join ', ')）"

    # Ribbon 控件 id 必须唯一，重复会让整个选项卡加载失败
    $ids = @($ribbon.SelectNodes("//*[@id]") | ForEach-Object { $_.id })
    $dupIds = @($ids | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    Assert-Equal 0 $dupIds.Count "Ribbon 控件 id 无重复（重复：$($dupIds -join ', '))"

    #==========================================================================
    Section "M1 文本处理"

    $ws = $wb.Worksheets.Item(1)
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  hello  world  "
    $ws.Range("A2").Value2 = "abc" + [char]0x00A0 + "def"     # 不间断空格
    $ws.Range("A3").Value2 = "already clean"
    Select-On $ws "A1:A3"
    $r = & $Run "text.cleanSpaces"
    Assert-Equal "hello world" $ws.Range("A1").Value2 "清除空格：压缩首尾和中间空格"
    Assert-Equal "abc def" $ws.Range("A2").Value2 "清除空格：不间断空格转普通空格"

    $ws.Cells.Clear()
    # 列必须先设成文本格式，否则 Excel 在写入时就把 "1,234.5" 和全角数字
    # 自动转成了数值，这个用例会变成【假通过】——工具根本没被执行到。
    $ws.Columns(1).NumberFormat = "@"
    $ws.Range("A1").Value2 = "1,234.5"
    $ws.Range("A2").Value2 = [char]0xFF11 + [char]0xFF12      # 全角 12
    $ws.Range("A3").Value2 = "not a number"
    Assert-Equal "1,234.5" $ws.Range("A1").Value2 "前置条件：写入后仍是文本"
    Select-On $ws "A1:A3"
    $r = & $Run "text.toNumber"
    Assert-Equal 1234.5 $ws.Range("A1").Value2 "文本转数值：处理千分位逗号"
    Assert-Equal 12 $ws.Range("A2").Value2 "文本转数值：处理全角数字"
    Assert-Equal "not a number" $ws.Range("A3").Value2 "文本转数值：非数字保持不变"
    $ws.Columns(1).NumberFormat = "General"

    $ws.Cells.Clear()
    $ws.Columns(1).NumberFormat = "@"
    $ws.Range("A1").Value2 = [char]0xFF21 + [char]0xFF10 + [char]0x3000 + "x"   # 全角 A0 + 表意空格
    Select-On $ws "A1"
    $r = & $Run "text.toHalfWidth"
    Assert-Equal "A0 x" $ws.Range("A1").Value2 "全角转半角（含表意空格）"
    $ws.Columns(1).NumberFormat = "General"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "订单A123号"
    Select-On $ws "A1"
    $r = & $Run "text.extractDigits"
    Assert-Equal 123 $ws.Range("A1").Value2 "提取数字"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "订单A123号"
    Select-On $ws "A1"
    $r = & $Run "text.extractChinese"
    Assert-Equal "订单号" $ws.Range("A1").Value2 "提取中文"

    # 公式单元格必须跳过——把计算结果写回去会毁掉公式
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  x  "
    $ws.Range("A2").Formula = "=A1"
    Select-On $ws "A1:A2"
    $r = & $Run "text.cleanSpaces"
    Assert-Equal "=A1" $ws.Range("A2").Formula "文本处理跳过公式单元格"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "a-b-c"
    $ws.Range("A2").Value2 = "d-e"
    & $Param "separator" "-"
    Select-On $ws "A1:A2"
    $r = & $Run "text.splitColumn"
    Assert-Equal "a" $ws.Range("A1").Value2 "按分隔符拆列：第 1 列"
    Assert-Equal "c" $ws.Range("C1").Value2 "按分隔符拆列：第 3 列"
    Assert-Equal "" "$($ws.Range('C2').Value2)" "按分隔符拆列：不足的部分留空"

    #==========================================================================
    Section "M2 数据处理"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(
        @("名称", "值"),
        @("甲", 1), @("乙", 2), @("甲", 1), @("丙", 3), @("乙", 2)
    )
    & $Param "hasHeader" "true"
    & $Param "keyColumns" ""
    Select-On $ws "A1:B6"
    $r = & $Run "data.deleteDuplicates"
    Assert-Match $r "已删除 2 个重复行*" "删除重复值：按全部列判重"
    Assert-Equal "丙" $ws.Range("A4").Value2 "删除重复值：保留首次出现且顺序正确"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(
        @("产品", "1月", "2月"),
        @("甲", 10, 20),
        @("乙", 30, 40)
    )
    & $Param "keepColumns" "1"
    & $Param "skipBlank" "true"
    Select-On $ws "A1:C3"
    $r = & $Run "data.unpivot"
    Assert-Match $r "已转换为一维表，共 4 行*" "逆透视：行数正确"
    $uws = $wb.Worksheets.Item($wb.Worksheets.Count)
    Assert-Equal "项目" $uws.Range("B1").Value2 "逆透视：生成项目列"
    Assert-Equal "1月" $uws.Range("B2").Value2 "逆透视：项目取自原列标题"
    Assert-Equal 20 $uws.Range("C3").Value2 "逆透视：值正确"

    #==========================================================================
    Section "M6 公式工具"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = 10
    $ws.Range("A2").Formula = "=A1*2"
    $ws.Range("A3").Formula = "=1/0"
    Select-On $ws "A1:A3"
    $r = & $Run "formula.toValues"
    Assert-Equal 20 $ws.Range("A2").Value2 "公式转值：结果正确"
    Assert-Equal "20" $ws.Range("A2").Formula "公式转值：公式已消失"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = 0
    $ws.Range("A2").Formula = "=1/A1"
    & $Param "fallback" "无"
    Select-On $ws "A1:A2"
    $r = & $Run "formula.wrapIfError"
    Assert-Equal "无" $ws.Range("A2").Value2 "套用 IFERROR：错误被兜住"
    $r = & $Run "formula.wrapIfError"
    Assert-Match $r "*跳过 1 个已有 IFERROR 的公式*" "套用 IFERROR：不重复包裹"

    #==========================================================================
    Section "M7 数据体检"

    $ws.Cells.Clear()
    # 必须先设文本格式：否则 Excel 在写入时就把 "123" 转成数值、"2024/1/15"
    # 转成日期，体检自然一个问题也检不出来，用例会变成假通过。
    $ws.Columns(1).NumberFormat = "@"
    $ws.Range("A1").Value2 = "值"
    $ws.Range("A2").Value2 = "123"            # 文本型数字
    $ws.Range("A3").Value2 = " x "            # 首尾空格
    $ws.Range("A5").Value2 = "2024/1/15"      # 文本型日期
    Select-On $ws "A1"
    $r = & $Run "audit.scan"
    Write-Host "  返回：$r" -ForegroundColor DarkGray
    Assert-Match $r "*文本型数字：1*" "体检：检出文本型数字"
    Assert-Match $r "*首尾空格*" "体检：检出首尾空格"
    Assert-Match $r "*文本型日期：1*" "体检：检出文本型日期"
    Assert-Match $r "*空行：1*" "体检：检出空行"
    $ws.Columns(1).NumberFormat = "General"

    #==========================================================================
    Section "M9 辅助增强"

    # 金额大写的零处理是最容易错的地方，专门多测几个边界
    $ws.Cells.Clear()
    $amounts = @(0, 100.05, 1000, 10001, 100000000, 1234.56, -50)
    Set-Grid $ws 1 1 @(@(0), @(100.05), @(1000), @(10001), @(100000000), @(1234.56), @(-50))
    Select-On $ws ("A1:A" + $amounts.Count)
    $r = & $Run "misc.amountToChinese"
    Assert-Equal "零元整"                 $ws.Range("B1").Value2 "金额大写：零"
    Assert-Equal "壹佰元零伍分"           $ws.Range("B2").Value2 "金额大写：角位为零而分位不为零"
    Assert-Equal "壹仟元整"               $ws.Range("B3").Value2 "金额大写：整数补整"
    Assert-Equal "壹万零壹元整"           $ws.Range("B4").Value2 "金额大写：节间补零"
    Assert-Equal "壹亿元整"               $ws.Range("B5").Value2 "金额大写：亿"
    Assert-Equal "壹仟贰佰叁拾肆元伍角陆分" $ws.Range("B6").Value2 "金额大写：角分齐全"
    Assert-Equal "负伍拾元整"             $ws.Range("B7").Value2 "金额大写：负数"

    $ws.Cells.Clear()
    $ws.Columns(1).NumberFormat = "@"    # 不设文本格式的话 18 位会被当成数值，精度直接丢失
    $ws.Range("A1").Value2 = "标题"
    $ws.Range("A2").Value2 = "110101199003072519"   # 校验位正确（加权和 190，190 mod 11 = 3 -> "9"）
    $ws.Range("A3").Value2 = "110101199003072516"   # 校验位错误
    $ws.Range("A4").Value2 = "12345"                 # 格式不对
    Select-On $ws "A2:A4"
    $r = & $Run "misc.parseId"
    Assert-Equal "男" $ws.Range("C2").Value2 "身份证：性别"
    Assert-Equal "正确" $ws.Range("E2").Value2 "身份证：校验位正确"
    Assert-Equal "校验位错误" $ws.Range("E3").Value2 "身份证：检出错误校验位"
    Assert-Equal "格式不正确" $ws.Range("E4").Value2 "身份证：检出格式错误"

    $ws.Cells.Clear()
    $ws.Columns(1).NumberFormat = "@"    # 否则 20240115 会被 Excel 直接存成数值
    $ws.Range("A1").Value2 = "20240115"
    $ws.Range("A2").Value2 = "2024.1.15"
    $ws.Range("A3").Value2 = "2024年1月15日"
    Select-On $ws "A1:A3"
    $r = & $Run "misc.normalizeDates"
    Write-Host "  返回：$r" -ForegroundColor DarkGray
    # 单元格里存的是日期序列值。用 FromOADate 转成日期再比 yyyy-MM-dd，
    # 不要比 .Value2 的显示串——那个随区域设置变化。
    # 另外 .Value 在 PowerShell 的 COM 适配器里是带参属性，不能直接取值，只能用 .Value2。
    Assert-Equal "2024-01-15" (Get-CellDate $ws "A1") "日期规范化：8 位数字"
    Assert-Equal "2024-01-15" (Get-CellDate $ws "A2") "日期规范化：点分隔"
    Assert-Equal "2024-01-15" (Get-CellDate $ws "A3") "日期规范化：中文年月日"
    $ws.Columns(1).NumberFormat = "General"

    #==========================================================================
    Section "M8 可视化"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@(10), @(20), @(30), @(40), @(50))
    Select-On $ws "A1:A5"
    $r = & $Run "viz.dataBars"
    Assert-Equal 1 $ws.Range("A1:A5").FormatConditions.Count "数据条：已添加条件格式"
    $r = & $Run "viz.clearCF"
    Assert-Equal 0 $ws.Range("A1:A5").FormatConditions.Count "清除条件格式"

    #==========================================================================
    Section "M1 文本处理（其余变体）"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "hello world"
    Select-On $ws "A1"
    $null = & $Run "text.toProper"
    Assert-Equal "Hello World" $ws.Range("A1").Value2 "首字母大写"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "AB"
    Select-On $ws "A1"
    $null = & $Run "text.toFullWidth"
    Assert-Equal ([string][char]0xFF21 + [string][char]0xFF22) $ws.Range("A1").Value2 "半角转全角"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "第一行" + [char]10 + "第二行"
    Select-On $ws "A1"
    $null = & $Run "text.removeLineBreaks"
    Assert-Equal "第一行第二行" $ws.Range("A1").Value2 "删除换行符"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "订单Abc123号"
    Select-On $ws "A1"
    $null = & $Run "text.extractEnglish"
    Assert-Equal "Abc" $ws.Range("A1").Value2 "提取字母"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "电话010-1234"
    & $Param "pattern" "[0-9]+"
    & $Param "replacement" "#"
    Select-On $ws "A1"
    $null = & $Run "text.regexReplace"
    Assert-Equal "电话#-#" $ws.Range("A1").Value2 "正则替换"

    # 非法正则必须给出清楚的报错，而不是在循环里反复抛
    $ws.Range("A1").Value2 = "x"
    & $Param "pattern" "[unclosed"
    & $Param "replacement" ""
    Select-On $ws "A1"
    $r = & $Run "text.regexReplace"
    Assert-Match $r "*正则表达式无效*" "正则替换：非法表达式被拦下"

    #==========================================================================
    Section "M1 合并单元格"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("甲"), @(""), @("乙"))
    $null = $ws.Activate()
    $ws.Range("A1:A2").Merge()
    Select-On $ws "A1:A3"
    $r = & $Run "cells.unmergeFill"
    Assert-Match $r "已拆分 1 处合并单元格并填充*" "拆分并填充：处理了合并区"
    Assert-Equal $false $ws.Range("A1").MergeCells "拆分并填充：已取消合并"
    Assert-Equal "甲" $ws.Range("A2").Value2 "拆分并填充：原值填满整个区域（不是只留左上角）"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("甲"), @("甲"), @("乙"))
    Select-On $ws "A1:A3"
    $r = & $Run "cells.mergeSame"
    Assert-Match $r "已合并 1 处相同内容*" "合并相同项：只合并相邻相同的"
    Assert-Equal $true $ws.Range("A1").MergeCells "合并相同项：前两行已合并"
    Assert-Equal $false $ws.Range("A3").MergeCells "合并相同项：不同值没有被并进去"
    $ws.Cells.UnMerge()

    #==========================================================================
    Section "M2 数据处理（其余）"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("a", "", "c"), @("d", "", "f"))
    Select-On $ws "A1:C2"
    $r = & $Run "data.deleteEmptyCols"
    Assert-Equal "已删除 1 个空列。" $r "删除空列"
    Assert-Equal "c" $ws.Range("B1").Value2 "删除空列：右侧列已左移"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("名称"), @("甲"), @("乙"), @("甲"))
    & $Param "hasHeader" "true"
    & $Param "keyColumns" ""
    Select-On $ws "A1:A4"
    $r = & $Run "data.markDuplicates"
    Assert-Match $r "已标记 1 个重复行*" "标记重复值：只标非首次出现的"
    Assert-Equal 13551615 $ws.Range("A4").Interior.Color "标记重复值：重复行已着色"
    Assert-Equal 16777215 $ws.Range("A2").Interior.Color "标记重复值：首次出现的没被标"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("甲", "乙"), @("甲", "丙"))
    Select-On $ws "A1:B2"
    $r = & $Run "data.extractUnique"
    Assert-Match $r "已提取 3 个唯一值*" "提取唯一值：跨行跨列去重"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@(1, 2, 3), @(4, 5, 6))
    Select-On $ws "A1:C2"
    $r = & $Run "data.transpose"
    Assert-Match $r "已转置 2 行 × 3 列*" "行列转置"
    $tws = $wb.Worksheets.Item($wb.Worksheets.Count)
    Assert-Equal 4 $tws.Range("B1").Value2 "行列转置：值落位正确"

    # 两表对比：三类差异都要能认出来
    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("编号", "值"), @("k1", "a"), @("k2", "b"), @("k3", "c"))
    $wsB = $wb.Worksheets.Add()
    $wsB.Name = "对比B"
    Set-Grid $wsB 1 1 @(@("编号", "值"), @("k1", "a"), @("k2", "改过"), @("k4", "d"))
    & $Param "rangeA" ($ws.Name + "!A1:B4")
    & $Param "rangeB" "对比B!A1:B4"
    & $Param "keyColumn" "1"
    $r = & $Run "data.compare"
    Assert-Match $r "*完全一致：1 行*" "两表对比：完全一致的行数"
    Assert-Match $r "*内容不同：1 行*" "两表对比：内容不同的行数"
    Assert-Match $r "*仅表 A 有：1 行*" "两表对比：仅 A 有"
    Assert-Match $r "*仅表 B 有：1 行*" "两表对比：仅 B 有"

    #==========================================================================
    Section "M6 公式工具（其余）"

    $ws.Cells.Clear()
    $ws.Range("A1").Formula = "=1/0"
    $ws.Range("A2").Value2 = 1
    Select-On $ws "A1:A2"
    $r = & $Run "formula.findErrors"
    Assert-Match $r "找到 1 个错误值*" "定位错误值"

    $r = & $Run "formula.toggleView"
    Assert-Equal "已切换为显示公式。" $r "显示公式：切换开"
    $r = & $Run "formula.toggleView"
    Assert-Equal "已切换为显示计算结果。" $r "显示公式：切换回"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = 1
    $ws.Range("A1:A3").FormatConditions.Delete()
    $null = $ws.Range("A1:A3").FormatConditions.Add(1, 5, "0")   # xlCellValue / xlGreater
    Select-On $ws "A1:A3"
    $r = & $Run "formula.cleanRules"
    Assert-Match $r "已清除 1 条条件格式规则*" "清理格式规则"

    $r = & $Run "formula.cleanNames"
    Assert-Match $r "*失效的已定义名称*" "清理失效名称：没有失效名称时如实说明"

    $r = & $Run "formula.breakLinks"
    Assert-Equal "当前工作簿没有外部链接。" $r "断开外部链接：无链接时如实说明"

    #==========================================================================
    Section "M7 一键清洗"

    $ws.Cells.Clear()
    $ws.Columns(1).NumberFormat = "@"
    $ws.Range("A1").Value2 = " 甲 "
    $ws.Range("A2").Value2 = "123"
    $ws.Range("A4").Value2 = "乙"
    Select-On $ws "A1:A4"
    $r = & $Run "audit.quickClean"
    Assert-Match $r "*一键清洗完成*" "一键清洗：执行完成"
    Assert-Match $r "*合并单元格和错误值需要人工判断*" "一键清洗：如实说明未自动处理的部分"
    $ws.Columns(1).NumberFormat = "General"

    #==========================================================================
    Section "M8 可视化（其余）"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@(10), @(20), @(30))
    Select-On $ws "A1:A3"
    $null = & $Run "viz.colorScale"
    Assert-Equal 1 $ws.Range("A1:A3").FormatConditions.Count "色阶：已添加"
    $null = & $Run "viz.clearCF"

    $null = & $Run "viz.iconSet"
    Assert-Equal 1 $ws.Range("A1:A3").FormatConditions.Count "图标集：已添加"
    $null = & $Run "viz.clearCF"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@(1, 5, 3), @(4, 2, 6))
    & $Param "sparkType" "1"
    Select-On $ws "A1:C2"
    $r = & $Run "viz.sparklines"
    Assert-Match $r "已生成 2 个迷你图*" "批量迷你图"
    Assert-Equal 1 $ws.Range("D1").SparklineGroups.Count "批量迷你图：画在数据右侧一列"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("甲", 1), @("乙", 2))
    & $Param "chartType" "1"
    Select-On $ws "A1:B2"
    $r = & $Run "viz.quickChart"
    Assert-Equal "已生成柱形图。" $r "快速图表"
    Assert-Equal 1 $ws.ChartObjects().Count "快速图表：图表对象已创建"

    $r = & $Run "viz.unifyCharts"
    Assert-Equal "已统一 1 个图表的格式。" $r "统一图表格式"
    $ws.ChartObjects().Delete()

    #==========================================================================
    Section "M9 聚光灯（开关与清理）"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "x"
    Select-On $ws "A1"
    $cfBefore = $ws.Cells.FormatConditions.Count
    $r = & $Run "misc.spotlight"
    Assert-Match $r "聚光灯已开启*" "聚光灯：开启"
    Assert-Equal ($cfBefore + 2) $ws.Cells.FormatConditions.Count "聚光灯：加了行列两条条件格式"

    $r = & $Run "misc.spotlight"
    Assert-Equal "聚光灯已关闭。" $r "聚光灯：关闭"
    Assert-Equal $cfBefore $ws.Cells.FormatConditions.Count "聚光灯：关闭后精确移除，不残留用户格式"

    #==========================================================================
    Section "Core"

    $r = & $Run "core.selfTest"
    Assert-Match $r "*加载宏工作正常*" "自检命令"
    $r = & $Run "core.about"
    Assert-Match $r "*Excel 通用工具箱*" "关于命令"
    $null = & $Run "core.resetEnv"
    Assert-Equal $true $xl.ScreenUpdating "环境复位"

    #==========================================================================
    Section "撤销往返（完整指纹比对）"

    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(
        @("A1",  10, "C1"),
        @("A2",  20, "C2"),
        @("",    "", ""),
        @("A4",  40, "C4"),
        @("A5",  50, "C5"),
        @("",    "", ""),
        @("A7",  70, "C7"),
        @("A8",  80, "C8"),
        @("",    "", ""),
        @("A10", 100, "C10")
    )
    $ws.Cells(1, 4).Formula = "=B1*2"
    $ws.Cells(2, 4).Formula = "=SUM(B1:B2)"
    $ws.Cells(2, 2).NumberFormat = "0.00"
    $ws.Cells(1, 1).Font.Bold = $true
    $ws.Cells(4, 2).Interior.Color = 255
    $ws.Rows(1).RowHeight = 30
    $ws.Columns(1).ColumnWidth = 22

    $before = Get-Fingerprint $ws 10 4

    Select-On $ws "A1:D10"
    $r = & $Run "data.deleteEmptyRows"
    Assert-Equal "已删除 3 个空行。" $r "删除空行：条数正确"
    Assert-Equal "A4"  $ws.Cells(3, 1).Value2 "删除空行：下方行已上移"

    & $Undo
    $after = Get-Fingerprint $ws 10 4
    Assert-Equal $before $after "结构性操作撤销后完整还原（值/公式/数字格式/字体/底纹/行高/列宽）"
    if ($before -ne $after) { Show-FirstDiff $before $after }

    # 非结构性操作的撤销走的是另一条路径（Capture 而非 CaptureSheet），单独测
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  padded  "
    $ws.Range("A2").Value2 = "ALSO"
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A2").Interior.Color = 65535
    $before2 = Get-Fingerprint $ws 2 1

    Select-On $ws "A1:A2"
    $r = & $Run "text.cleanSpaces"
    Assert-Equal "padded" $ws.Range("A1").Value2 "非结构性操作已生效"
    & $Undo
    $after2 = Get-Fingerprint $ws 2 1
    Assert-Equal $before2 $after2 "非结构性操作撤销后完整还原"
    if ($before2 -ne $after2) { Show-FirstDiff $before2 $after2 }

    # 多步撤销
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "one"
    $ws.Range("A2").Value2 = "two"
    Select-On $ws "A1:A2"
    $null = & $Run "text.toUpper"
    $null = & $Run "text.toLower"
    Assert-Equal "one" $ws.Range("A1").Value2 "连续两次操作后的状态"
    & $Undo
    Assert-Equal "ONE" $ws.Range("A1").Value2 "撤销第 1 步：回到大写"
    & $Undo
    Assert-Equal "one" $ws.Range("A1").Value2 "撤销第 2 步：回到原始"

    #==========================================================================
    Section "撤销框架的失败路径"

    # 这一段测的全是"出错时会不会骗用户"。
    # 之前的用例只走成功路径，而恰恰是失败路径上最容易出现
    # "说了没验证过的话"——Codex 验收就是在这里挑出的问题。

    # --- 规模超限时必须【告诉用户】撤销记录没留下，不能默默丢掉 ---
    # 用默认的 100 万上限，不去改注册表：在 A1 和最后一行各放一个值，
    # 已用区域就是 1,048,576 格。因为超限判断在拷贝【之前】做，
    # 这里不会真的去拷 100 万格，用例照样很快。
    function Get-UndoDepth {
        $m = [regex]::Match($xl.Run("'$OutputName'!Toolbox_SelfCheck"), "undoDepth=(\d+)")
        return [int]$m.Groups[1].Value
    }

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "x"
    $ws.Range("A1048576").Value2 = "y"
    Select-On $ws "A1:A3"                    # 只扫 3 行，但整表快照会超限
    $depthBefore = Get-UndoDepth
    $r = & $Run "data.deleteEmptyRows"
    Assert-Match $r "*超过撤销上限*" "超限时如实告知撤销记录未保留"
    Assert-Match $r "*无法撤销*" "超限时明确说明无法撤销"
    # 比深度而不是比 canUndo：前面的用例已经在栈里留了记录，
    # 这里要验的是"这一步没有进栈"，提示与实际必须一致
    Assert-Equal $depthBefore (Get-UndoDepth) "超限时该步确实没进撤销栈（提示与实际一致）"
    $ws.Cells.Clear()

    # --- 还原失败时，撤销记录必须【保留】，让用户能排除障碍后重试 ---
    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(@("x"), @(""), @("y"))
    Select-On $ws "A1:A3"
    $null = & $Run "data.deleteEmptyRows"
    Assert-Match ($xl.Run("'$OutputName'!Toolbox_SelfCheck")) "*canUndo=True*" "操作后撤销栈里有记录"

    # 保护工作表，让还原必然失败。
    # Toolbox_Undo 自己捕获错误并以字符串返回，不会抛出未处理的 VBA 错误——
    # 否则会弹出"运行时错误 1004"调试框，把无人值守的测试永久挂死。
    $ws.Protect()
    $undoErr = $xl.Run("'$OutputName'!Toolbox_Undo")
    $ws.Unprotect()
    Assert-Match $undoErr "ERROR:*" "还原失败时如实返回错误，而不是假装成功"
    Assert-Match ($xl.Run("'$OutputName'!Toolbox_SelfCheck")) "*canUndo=True*" "还原失败后撤销记录没丢，可以重试"

    # 障碍排除后重试应当成功
    $depthBeforeRetry = Get-UndoDepth
    $retry = $xl.Run("'$OutputName'!Toolbox_Undo")
    Assert-Equal "" "$retry" "解除保护后重试撤销：没有报错"
    Assert-Equal "" "$($ws.Range('A2').Value2)" "解除保护后重试撤销：空行已还原"
    # 同样比深度：栈里还有前面用例留下的记录，不能指望它变空
    Assert-Equal ($depthBeforeRetry - 1) (Get-UndoDepth) "撤销成功后才出栈（深度减一）"

    #==========================================================================
    Section "聚光灯不误删用户的条件格式"

    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "x"
    Select-On $ws "A1"

    # 造一条和聚光灯【同底色】但公式不同的用户规则。
    # 只按颜色识别的话，关闭聚光灯就会把它一起删掉。
    $userRule = $ws.Range("A1:A5").FormatConditions.Add(2, 0, "=TRUE")   # xlExpression
    $userRule.Interior.Color = 14083324                                  # 与 MARKER_COLOR 相同
    $cfWithUserRule = $ws.Cells.FormatConditions.Count

    $null = & $Run "misc.spotlight"
    $null = & $Run "misc.spotlight"
    Assert-Equal $cfWithUserRule $ws.Cells.FormatConditions.Count "关闭聚光灯后用户的同色规则仍在"

    $ws.Cells.FormatConditions.Delete()

    #==========================================================================
    Section "错误处理与边界"

    # 参数取消必须安静退出，不能留下半成品，也不能弹框
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "x"
    $xl.Run("'$OutputName'!Toolbox_ClearParams")
    Select-On $ws "A1"
    $r = & $Run "text.addAffix"
    Assert-Equal "CANCELLED" $r "缺少参数时按取消处理"
    Assert-Equal "x" $ws.Range("A1").Value2 "取消后数据未被修改"

    # 合并单元格必须被拦下，而不是产生不可预期的结果
    $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "m"
    $null = $ws.Activate()
    $ws.Range("A1:B1").Merge()
    Select-On $ws "A1:B3"
    $r = & $Run "data.deleteEmptyRows"
    Assert-Match $r "*合并单元格*" "删除空行遇到合并单元格时拒绝执行"
    $ws.Range("A1:B1").UnMerge()

    #==========================================================================
    Section "环境还原"

    Assert-Equal $true $xl.ScreenUpdating "ScreenUpdating 已还原"
    Assert-Equal $true $xl.EnableEvents   "EnableEvents 已还原"
    Assert-Equal "xlCalculationAutomatic" $xl.Calculation "Calculation 已还原为自动"
}
catch {
    $script:fail++
    Write-Host ""
    Write-Host "测试过程异常（$script:section）：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($xl) {
        try { $xl.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 400
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }
}

Write-Host ""
Write-Host ("通过 {0} / 失败 {1}" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
