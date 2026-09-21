<#
.SYNOPSIS
    生成业务培训用的演示数据（一整套，含故意做脏的表和多文件夹夹具）。

.DESCRIPTION
    培训现场最怕两件事：数据被学员改乱了没法重来，以及讲稿里写的
    "第 7 行"和屏幕上对不上。所以这个脚本有两条硬规矩：

      1.【完全确定，不用任何随机数】。同一个版本跑一百次，生成的
         内容逐字节一致。讲稿里可以放心写行号、写具体的客户名，
         截图也不会过期。

      2.【可以反复重跑】。学员把数据改乱了，重跑一次就回到干净的
         初始状态。默认会拒绝覆盖已存在的目录，加 -Force 才覆盖——
         演示目录下可能有学员自己存的东西。

    生成的每一处"脏"都是真实业务里天天遇到的，不是为了演示编出来的：
    文本型数字导致求和是 0、不间断空格导致 VLOOKUP 匹配不上、
    四种写法的文本型日期、散落的空行、完全重复的行、错误值。

.PARAMETER OutDir
    输出目录。默认是桌面上的「工具箱培训演示」。

.PARAMETER Force
    目录已存在时覆盖。不给这个开关就拒绝运行。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File training\New-DemoData.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File training\New-DemoData.ps1 -Force
#>
[CmdletBinding()]
param(
    [string]$OutDir = "",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot = Split-Path -Parent $PSScriptRoot
# 复用构建脚本那套 Excel 生命周期管理：WPS 劫持 COM 的守卫、
# 进程登记与回收都在里面，没必要在这儿再写一遍（也不该写第二份）
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

if (-not $OutDir) {
    $desktop = [Environment]::GetFolderPath("Desktop")
    $OutDir  = Join-Path $desktop "工具箱培训演示"
}

function Write-Step($m) { Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-Ok  ($m) { Write-Host "    [完成] $m" -ForegroundColor Green }

#-----------------------------------------------------------------------------
# 写一个单元格的值。
#
# 【不能直接写 $range.Value2 = $number】，会抛
#     "无法将 System.Double 强制转换为 System.String"
#
# 原因：PowerShell 的 COM 适配器把 Value2 这个 setter 的参数类型
# 【按第一次用法缓存住了】。脚本总是先写表头字符串，适配器于是认定
# Value2 只收 String，之后每一次数字赋值都失败。
#
# 症状极具迷惑性：单独拎一行出来跑完全正常，放回脚本里就报类型转换错，
# 很容易被误判成"数据有问题"。把 Range 取到中间变量【也没用】——
# 缓存是挂在属性上的，不是挂在调用点上的（实测确认）。
#
# 唯一可靠的解法是后期绑定 InvokeMember，彻底绕开适配器。
#-----------------------------------------------------------------------------
function Set-Cell($Sheet, [int]$Row, [int]$Col, $Value) {
    $cell = $Sheet.Cells.Item($Row, $Col)
    # 走后期绑定，绕开 PowerShell 的 COM 适配器（原因见上面的注释）
    [void]$cell.GetType().InvokeMember('Value2', 'SetProperty', $null, $cell, @($Value))
}

#-----------------------------------------------------------------------------
# 目录准备
#
# 【不加 -Force 就不覆盖】。这个目录在学员桌面上，很可能混进了他们
# 自己的文件。一个"生成演示数据"的脚本把讲师的备课笔记删掉，
# 比它能省的那点事严重得多。
#-----------------------------------------------------------------------------
if (Test-Path -LiteralPath $OutDir) {
    if (-not $Force) {
        Write-Host ""
        Write-Host "目录已存在：$OutDir" -ForegroundColor Yellow
        Write-Host "重新生成会删掉里面的内容。确认没有你自己的文件后，加 -Force 重跑：" -ForegroundColor Yellow
        Write-Host "    powershell -ExecutionPolicy Bypass -File training\New-DemoData.ps1 -Force" -ForegroundColor Yellow
        exit 2
    }
    Remove-Item -LiteralPath $OutDir -Recurse -Force
}
$null = New-Item -ItemType Directory -Path $OutDir -Force

# 看不见却要命的三种空白字符
$NBSP = [char]0x00A0      # 不间断空格：从网页/系统导出的数据里到处都是
$FULL = [char]0x3000      # 全角空格
$ZWSP = [char]0x200B      # 零宽空格：连选中都看不出来

$xlUp = -4162

#-----------------------------------------------------------------------------
# 演示数据（写死，不用随机数——讲稿要能引用具体行号和客户名）
#-----------------------------------------------------------------------------
$customers = @("北辰科技", "海川贸易", "云岭实业", "长风物流", "青禾食品")
$products  = @("A100 控制器", "B200 传感器", "C300 连接件", "D400 电源模块")
$sellers   = @("张伟", "李娜", "王强", "刘洋")

# 日期故意四种写法，全部是文本——这是「排序乱套、筛选不出来」的根因
$dateForms = @("20240115", "2024.1.15", "2024年1月15日", "2024-01-15")

$rows = @()
for ($i = 0; $i -lt 40; $i++) {
    $day  = [int](1 + ($i % 28))
    $mon  = [int](1 + [math]::Floor($i / 28))
    $cust = $customers[$i % $customers.Count]

    # 每隔几行给客户名掺进不可见字符：看着和对照表一模一样，
    # VLOOKUP 就是匹配不上——培训里最有共鸣的一个场景
    switch ($i % 4) {
        0 { $custDirty = " $cust " }
        1 { $custDirty = "$cust$NBSP" }
        2 { $custDirty = "$FULL$cust" }
        3 { $custDirty = "$cust$ZWSP" }
    }

    switch ($i % 4) {
        0 { $dateText = "{0:D4}{1:D2}{2:D2}" -f 2024, $mon, $day }
        1 { $dateText = "2024.$mon.$day" }
        2 { $dateText = "2024年${mon}月${day}日" }
        3 { $dateText = "2024-{0:D2}-{1:D2}" -f $mon, $day }
    }

    $qty   = 1 + ($i % 9) * 3
    $price = 120 + ($i % 7) * 35

    # 数量和单价【全部是文本型数字】，其中几行还带千分位逗号和不间断空格
    $qtyText = if ($i % 5 -eq 0) { "$qty$NBSP" } else { "$qty" }
    $priceText = if ($price -ge 300) { "{0:N0}" -f $price } else { "$price" }

    $rows += ,@(
        ("SO2024{0:D4}" -f (1001 + $i)),
        $dateText,
        $custDirty,
        $products[$i % $products.Count],
        $qtyText,
        $priceText,
        $sellers[$i % $sellers.Count]
    )
}

# 完全重复的两行（演示「标记重复值」「删除重复值」）
$rows += ,$rows[3]
$rows += ,$rows[11]

Write-Step "启动 Excel"
$xl = $null
try {
    $xl = New-RealExcel
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false

    #=========================================================================
    Write-Step "生成 01_销售明细（脏数据）.xlsx"
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "销售明细"

    # 标题区用合并单元格——很多工具遇到合并会拒绝执行，正好演示"为什么拒绝"
    $ws.Range("A1:G1").Merge()
    $ws.Range("A1").Value2 = "2024 年上半年销售明细（系统导出，未清洗）"
    $ws.Range("A1").HorizontalAlignment = -4108
    $ws.Range("A1").Font.Bold = $true

    $headers = @("订单号", "日期", "客户名称", "产品", "数量", "单价", "销售员")
    for ($c = 0; $c -lt $headers.Count; $c++) {
        Set-Cell $ws 2 ($c + 1) $headers[$c]
        $ws.Cells.Item(2, $c + 1).Font.Bold = $true
    }

    # 【先把整列设成文本，再写值】。顺序反了的话 Excel 会把
    # "20240115" 认成数字、把 "2024-01-15" 认成日期，
    # 演示数据就自己"干净"了，该演的问题反而演不出来。
    $ws.Range("B:B").NumberFormat = "@"   # 日期（文本型）
    $ws.Range("C:C").NumberFormat = "@"   # 客户名
    $ws.Range("E:F").NumberFormat = "@"   # 数量、单价（文本型数字）

    # 每隔 9 行插一个空行，模拟系统导出的分页残留
    $r = 3
    $n = 0
    foreach ($row in $rows) {
        if ($n -gt 0 -and $n % 9 -eq 0) { $r++ }     # 留一个空行
        for ($c = 0; $c -lt $row.Count; $c++) {
            Set-Cell $ws $r ($c + 1) $row[$c]
        }
        $r++; $n++
    }
    $lastRow = $r - 1

    # 单件均摊列：F/E 两列都是文本型数字，相除直接得 #VALUE!。
    # 【这个错误值不是硬造的】——它就是"数据是文本"的直接后果，
    # 清洗完文本型数字之后它会自己变成正常数值，演示里这一步很有说服力。
    Set-Cell $ws 2 8 "单件均摊"
    $ws.Cells.Item(2, 8).Font.Bold = $true
    for ($i = 3; $i -le $lastRow; $i++) {
        if ($ws.Cells.Item($i, 1).Value2) {
            $ws.Cells.Item($i, 8).Formula = "=F$i/E$i"
        }
    }

    $ws.Columns.Item("A:H").AutoFit() | Out-Null

    # --- 第二张表：干净的客户对照表，用来演示"看着一样却匹配不上" ---
    $ws2 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws)
    $ws2.Name = "客户对照表"
    $ws2.Range("A1").Value2 = "客户名称"
    $ws2.Range("B1").Value2 = "信用等级"
    $ws2.Range("A1:B1").Font.Bold = $true
    $grades = @("AAA", "AA", "A", "BBB", "BB")
    for ($i = 0; $i -lt $customers.Count; $i++) {
        Set-Cell $ws2 ($i + 2) 1 $customers[$i]     # 干净的名字
        Set-Cell $ws2 ($i + 2) 2 $grades[$i]
    }
    $ws2.Columns.Item("A:B").AutoFit() | Out-Null

    # --- 第三张表：把"求和是 0"摆到台面上 ---
    $ws3 = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $ws2)
    $ws3.Name = "汇总"
    $ws3.Range("A1").Value2 = "本表的公式没有错，但结果都不对。为什么？"
    $ws3.Range("A1").Font.Bold = $true
    $ws3.Range("A3").Value2 = "数量合计"
    $ws3.Range("B3").Formula = "=SUM(销售明细!E:E)"
    $ws3.Range("A4").Value2 = "单价合计"
    $ws3.Range("B4").Formula = "=SUM(销售明细!F:F)"
    $ws3.Range("A6").Value2 = "查「北辰科技」的信用等级"
    $ws3.Range("B6").Formula = '=VLOOKUP(销售明细!C3,客户对照表!A:B,2,0)'
    $ws3.Range("A8").Value2 = "↑ 求和是 0、查询是 #N/A，都不是公式的问题。"
    $ws3.Range("A9").Value2 = "   用「工具箱 → 数据体检」一扫就知道。"
    $ws3.Columns.Item("A:B").AutoFit() | Out-Null

    $ws.Activate()
    $wb.SaveAs((Join-Path $OutDir "01_销售明细（脏数据）.xlsx"), 51)
    $wb.Close($false)
    Write-Ok "01_销售明细（脏数据）.xlsx —— $($rows.Count) 行数据，含 2 行完全重复"

    #=========================================================================
    Write-Step "生成 02_员工信息.xlsx"
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "员工名册"

    $empHeaders = @("工号", "姓名", "身份证号", "月度报销金额")
    for ($c = 0; $c -lt $empHeaders.Count; $c++) {
        Set-Cell $ws 1 ($c + 1) $empHeaders[$c]
        $ws.Cells.Item(1, $c + 1).Font.Bold = $true
    }

    # 【身份证号必须是文本】：数值型会变成科学计数法，且末位被抹成 0。
    #
    # 【校验位是算出来的，不是编的】。「身份证解析」会按 GB 11643 验校验位，
    # 手写一串数字有 10/11 的概率校验不过——那样演示出来全是"校验位不符"，
    # 学员看到的就成了"这功能有问题"。所以这里现算。
    $ws.Range("C:C").NumberFormat = "@"

    function Get-IdCheckDigit([string]$first17) {
        $w = @(7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2)
        $map = @('1','0','X','9','8','7','6','5','4','3','2')
        $sum = 0
        for ($k = 0; $k -lt 17; $k++) { $sum += [int]::Parse($first17[$k]) * $w[$k] }
        return $map[$sum % 11]
    }

    $idBodies = @(
        "11010519880413567", "31010419900307425", "44030519951128123",
        "33010619870615231", "51010719920908441"
    )
    $ids = @()
    foreach ($b in $idBodies) { $ids += ($b + (Get-IdCheckDigit $b)) }

    # 【第 4 个故意改坏】：把校验位换成另一个字符，用来演示
    # "解析会如实拒绝"。不验校验位的工具随便编一串都能"解析"出生日，
    # 那比报错危险得多。
    $badTail = if ($ids[3].Substring(17) -eq "1") { "2" } else { "1" }
    $ids[3] = $ids[3].Substring(0, 17) + $badTail

    $names = @("陈静", "赵鹏", "孙丽", "周涛", "吴敏")
    $amts  = @(1280.5, 3465, 890.25, 12000, 56.8)
    for ($i = 0; $i -lt 5; $i++) {
        Set-Cell $ws ($i + 2) 1 ("EMP{0:D3}" -f (101 + $i))
        Set-Cell $ws ($i + 2) 2 $names[$i]
        Set-Cell $ws ($i + 2) 3 $ids[$i]
        Set-Cell $ws ($i + 2) 4 $amts[$i]
    }
    $ws.Range("A8").Value2 = "说明：以上为构造的测试数据，不是真实个人信息。"
    $ws.Range("A9").Value2 = "第 5 行（周涛）的身份证号校验位是错的，用来演示「身份证解析」会如实报错。"
    $ws.Columns.Item("A:D").AutoFit() | Out-Null

    $wb.SaveAs((Join-Path $OutDir "02_员工信息.xlsx"), 51)
    $wb.Close($false)
    Write-Ok "02_员工信息.xlsx —— 5 人，含 1 个校验位错误的号码"

    #=========================================================================
    Write-Step "生成 03_分公司数据\（合并文件夹演示）"
    $branchDir = Join-Path $OutDir "03_分公司数据"
    $null = New-Item -ItemType Directory -Path $branchDir -Force

    # 【三个文件的列顺序故意不一致】。「合并所有表」是按标题名对齐的，
    # 不是按列位置——这正是它和"手工复制粘贴"的关键区别，必须演出来。
    $branches = @(
        @{ Name = "华东分公司"; Cols = @("订单号", "客户", "金额", "销售员") }
        @{ Name = "华南分公司"; Cols = @("订单号", "金额", "客户", "销售员") }
        @{ Name = "华北分公司"; Cols = @("订单号", "客户", "金额") }          # 少一列
    )
    $bi = 0
    foreach ($b in $branches) {
        $wb = $xl.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = $b.Name

        for ($c = 0; $c -lt $b.Cols.Count; $c++) {
            Set-Cell $ws 1 ($c + 1) $b.Cols[$c]
            $ws.Cells.Item(1, $c + 1).Font.Bold = $true
        }
        for ($i = 0; $i -lt 6; $i++) {
            $vals = @{
                "订单号"  = "{0}{1:D3}" -f $b.Name.Substring(0, 2), (1 + $i)
                "客户"    = $customers[($i + $bi) % $customers.Count]
                "金额"    = 5000 + ($i * 1300) + ($bi * 700)
                "销售员"  = $sellers[($i + $bi) % $sellers.Count]
            }
            for ($c = 0; $c -lt $b.Cols.Count; $c++) {
                Set-Cell $ws ($i + 2) ($c + 1) $vals[$b.Cols[$c]]
            }
        }
        $ws.Columns.Item("A:D").AutoFit() | Out-Null
        $wb.SaveAs((Join-Path $branchDir "$($b.Name).xlsx"), 51)
        $wb.Close($false)
        $bi++
    }
    Write-Ok "03_分公司数据\ —— 3 个文件，列顺序各不相同，其中一个少一列"

    #=========================================================================
    Write-Step "生成 04_待重命名文件\（文件批处理演示）"
    $renameDir = Join-Path $OutDir "04_待重命名文件"
    $null = New-Item -ItemType Directory -Path $renameDir -Force

    # 文件名故意乱：空格、中英混杂、日期格式不统一——正是需要批量改名的场景
    $files = @(
        "报价单 - 北辰科技(最终版).xlsx",
        "报价单-海川贸易 2024.1.5.xlsx",
        "云岭实业 报价 20240210.xlsx",
        "长风物流报价单 copy.xlsx"
    )
    foreach ($f in $files) {
        $wb = $xl.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $wb.Worksheets.Item(1).Range("A1").Value2 = "演示用空表：$f"
        $wb.SaveAs((Join-Path $renameDir $f), 51)
        $wb.Close($false)
    }
    Write-Ok "04_待重命名文件\ —— 4 个文件，命名风格各不相同"
}
finally {
    if ($xl) {
        try { $xl.ScreenUpdating = $true } catch {}
        Close-ExcelInstance $xl
    }
}

#-----------------------------------------------------------------------------
# 给讲师的一张小抄，和数据放在一起，免得临场翻文档
#-----------------------------------------------------------------------------
$cheat = @()
$cheat += "工具箱培训演示数据"
$cheat += "============================================"
$cheat += ""
$cheat += "本目录由 training\New-DemoData.ps1 生成，可随时重跑还原。"
$cheat += "学员把数据改乱了不要紧，重跑一次就回到初始状态。"
$cheat += ""
$cheat += "01_销售明细（脏数据）.xlsx"
$cheat += "    「汇总」表：求和是 0、VLOOKUP 是 #N/A —— 开场用这个"
$cheat += "    「销售明细」表故意埋了："
$cheat += "        · 数量/单价是文本型数字（求和为 0 的元凶）"
$cheat += "        · 客户名带不可见字符（VLOOKUP 匹配不上的元凶）"
$cheat += "        · 日期有四种写法，全是文本（排序筛选都不对）"
$cheat += "        · 每 9 行一个空行、2 行完全重复、H 列是 #VALUE!（文本相除）"
$cheat += "        · 标题行是合并单元格（演示「为什么这个功能拒绝执行」）"
$cheat += ""
$cheat += "02_员工信息.xlsx      金额大写、身份证解析"
$cheat += "                      第 4 行校验位是错的，会如实报错"
$cheat += "03_分公司数据\        合并文件夹。三个文件列顺序不同，一个少一列"
$cheat += "04_待重命名文件\      生成文件清单 → 批量重命名"
$cheat += ""
$cheat += "完整讲稿见仓库 docs\培训演示用例.md"
$cheat += ""
$cheat += "重要：文件批处理（改名、导出）动的是磁盘文件，【不可撤销】。"
$cheat += "      演示前确认用的是本目录里的副本，不是谁的真实文件。"

[IO.File]::WriteAllLines((Join-Path $OutDir "讲师小抄.txt"), $cheat, [Text.UTF8Encoding]::new($true))

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  演示数据已生成" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "  $OutDir" -ForegroundColor Green
Write-Host ""
Write-Host "讲稿见 docs\培训演示用例.md；目录里另有一份「讲师小抄.txt」。" -ForegroundColor DarkGray
Write-Host "学员改乱了随时重跑本脚本（加 -Force）即可还原。" -ForegroundColor DarkGray
exit 0
