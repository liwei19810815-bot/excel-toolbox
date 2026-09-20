<#
.SYNOPSIS
    校验 customUI14.xml 里每一个 imageMso 在【当前这台机器的 Excel】上真的存在。

.DESCRIPTION
    为什么需要这个：

    imageMso 是 Office 内置图标的 ID，用它就不必往加载宏里塞图片资源。
    但有两个陷阱：

      1. 【这些 ID 在 Office 各应用之间共享，而每个应用只自带它自己用到的图标】。
         从 Word / Outlook / Access 抄来的 ID（TableMerge、AcceptInvitation、
         FilterBySelection……）在 Excel 里可能根本不存在。

      2. 【同样是 16.0，Microsoft 365 和 Excel 2021 的图标集不一样】。
         365 是滚动更新、图标集最全；2021 是冻结版本，集合更小。
         在 365 开发机上好好的图标，到 2021 上就是空白。

    而这种失败【不会让任何测试变红】：
    Excel 遇到无效的 imageMso 只是静默不画图标，功能区照样加载、
    按钮照样能点、check-ribbon.ps1 照样报 ribbon=True。
    用户看到的是一排没有图标的按钮——难看，而且不专业。

    Application.CommandBars.GetImageMso 对无效 ID 会抛错，
    所以可以逐个探测，把静默失败变成一条会红的断言。

.NOTES
    这个脚本【只能证明当前这台机器】。在开发机（Microsoft 365）上报 30/30，
    不等于 Excel 2021 / 2019 / 2016 上也是 30/30——图标集确实存在版本差异。

    所以交付流程里必须有这一步：【在目标版本的 Excel 上跑一次并留存输出】。
    本项目已知的 13 个无效 ID 是在所有版本上都不显示的（属于 ID 本身写错），
    但"当前全部有效"这个结论不能跨版本外推。

.EXAMPLE
    timeout 300 powershell -ExecutionPolicy Bypass -File tests\check-imagemso.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam",

    # 用来取"占位色块长什么样"的参照 ID。
    # 它必须是一个【GetImageMso 不报错、但画不出真实图形】的 idMso。
    # MergeCenterMenu 是实测出来的这么一个：菜单型 idMso 没有独立位图。
    [string]$PlaceholderRefId = "MergeCenterMenu"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

$XmlPath = Join-Path $RepoRoot "src\package\customUI\customUI14.xml"
if (-not (Test-Path $XmlPath)) { throw "找不到 $XmlPath。" }

Add-Type -AssemblyName System.Drawing

# 取图标文件的内容哈希。
#
# 【为什么用哈希比对而不是"数颜色"】：
# 第一版用"颜色少于 N 种就算占位色块"，结果误报了一大片——
# Undo、ViewSideBySide、FontDialog 这些都是正常的单色线条图标，
# 颜色数天然就只有四五种。颜色数区分不了"简洁"和"空白"。
#
# 真正稳的判据是：Excel 对所有画不出图的 ID 返回的是【同一张占位图】。
# 所以拿一个已知的占位 ID 当参照，谁和它一模一样谁就是占位。
# 没有阈值，不需要调参，也不会误伤简洁图标。
function Get-IconHash {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }
    catch { return $null }
}

# 把所有图标拼成一张总览图。自动判据挡不住"语义不对"，那只能靠看。
function New-IconContactSheet {
    param([string]$Dir)
    try {
        $files = @(Get-ChildItem $Dir -Filter *.bmp -ErrorAction SilentlyContinue | Sort-Object Name)
        if ($files.Count -eq 0) { return $null }

        $cols = 5; $cell = 120
        $rows = [int][math]::Ceiling($files.Count / $cols)
        $bmp = New-Object System.Drawing.Bitmap(([int]($cols * $cell)), ([int]($rows * $cell)))
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $fnt = New-Object System.Drawing.Font('Segoe UI', 7.5)

        for ($i = 0; $i -lt $files.Count; $i++) {
            [int]$x = ($i % $cols) * $cell
            [int]$y = [math]::Floor($i / $cols) * $cell
            $img = [System.Drawing.Image]::FromFile($files[$i].FullName)
            $g.DrawImage($img, [int]($x + 44), [int]($y + 10), 32, 32)
            $img.Dispose()
            $rect = New-Object System.Drawing.RectangleF([single]($x + 3), [single]($y + 48), [single]($cell - 6), [single]66)
            $g.DrawString([string]$files[$i].BaseName, $fnt, [System.Drawing.Brushes]::Black, $rect)
            $g.DrawRectangle([System.Drawing.Pens]::LightGray, [int]$x, [int]$y, [int]$cell, [int]$cell)
        }
        $g.Dispose()
        $png = Join-Path $Dir "_总览.png"
        $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        return $png
    }
    catch { return $null }
}

# ---- 收集 XML 里用到的全部 imageMso（控件 id -> 图标 id）----
[xml]$rx = Get-Content -LiteralPath $XmlPath -Raw -Encoding UTF8
$ns = New-Object System.Xml.XmlNamespaceManager($rx.NameTable)
$ns.AddNamespace("ui", "http://schemas.microsoft.com/office/2009/07/customui")

# 【按属性选，不要按控件类型枚举】。原先写的是
#   //ui:button | //ui:toggleButton | //ui:menu | //ui:gallery | //ui:splitButton
# 这份清单今天恰好覆盖全，但以后加了 dynamicMenu、checkBox 之类的控件就会
# 【静默漏检】——而漏检的表现就是又出现空白图标、又没有任何东西报警。
# 直接选"凡是带 imageMso 属性的元素"，控件类型再怎么变都漏不掉。
$uses = @()
foreach ($node in $rx.SelectNodes("//*[@imageMso]")) {
    $uses += [pscustomobject]@{
        ControlId = $node.GetAttribute("id")
        ImageMso  = $node.GetAttribute("imageMso")
    }
}

if ($uses.Count -eq 0) { Write-Host "XML 里没有用到 imageMso。"; exit 0 }

$distinct = $uses | Select-Object -ExpandProperty ImageMso -Unique | Sort-Object
Write-Host "==> 共 $($uses.Count) 处引用，$($distinct.Count) 个不同的 imageMso" -ForegroundColor Cyan

# Excel 进程的登记与回收统一由 build\_ExcelHost.ps1 负责
# （New-RealExcel 会自动登记，Close-ExcelInstance 负责收尾和回收孤儿）。

$bad = @()
$report = @()
$xl = $null
try {
    $xl = New-RealExcel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)

    $hostDesc = "$($xl.Name) $($xl.Version) build $($xl.Build)"
    Write-Host "    宿主：$hostDesc" -ForegroundColor DarkGray

    # 【必须让 VBA 在进程内去问，不能从 PowerShell 直接调 GetImageMso】。
    # 它返回 IPictureDisp，跨进程 marshal 回来时会直接挂死而且【不报错】——
    # 加了可见窗口和工作簿也一样。所以改成把 ID 列表传进去、只收字符串回来。
    $Xlam = Join-Path $RepoRoot "dist\$OutputName"
    if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }
    $null = $xl.Workbooks.Open($Xlam)

    $reply = $xl.Run("'$OutputName'!Toolbox_CheckImageMso", ($distinct -join "|"))
    $missing = @()
    if (-not [string]::IsNullOrWhiteSpace($reply)) {
        $missing = $reply -split '\|' | Where-Object { $_ }
    }

    # ---------------------------------------------------------------------
    # 【"存在"不等于"画得出来"】。
    #
    # GetImageMso 对某些 ID（尤其是 ...Menu 这类菜单型 idMso）会返回一个
    # 对象，但那个对象根本没有图形内容——功能区上呈现为一个无意义的色块。
    # 只判断"有没有抛错"会把这种放过去，本项目就漏掉了两个
    # （MergeCenterMenu、PivotTableInsertMenu），直到把图导出来看才发现。
    #
    # 所以这里把图标真的存成 BMP，和一个【已知占位 ID】的图做哈希比对。
    # 同时拼一张总览图，让人能一眼扫完——自动判据再好也替代不了看一眼。
    # ---------------------------------------------------------------------
    $iconDir = Join-Path $env:TEMP "ExcelToolbox_imagemso"
    if (Test-Path $iconDir) { Remove-Item $iconDir -Recurse -Force -ErrorAction SilentlyContinue }

    # 一起导出参照 ID，用来拿到"占位图长什么样"
    $null = $xl.Run("'$OutputName'!Toolbox_DumpImageMso",
                    (($distinct + $PlaceholderRefId) -join "|"), $iconDir)

    $refPath = Join-Path $iconDir "$PlaceholderRefId.bmp"
    $placeholderHash = Get-IconHash $refPath
    if (-not $placeholderHash) {
        # 【参照取不到必须判失败，不能降级成"只查存不存在"】。
        # 静默降级是不安全的失败模式：占位色块会被放过去、脚本还报绿，
        # 而这正是这个检查存在的唯一理由。宁可红，不要假绿。
        throw @"
取不到占位图参照（$PlaceholderRefId），无法判定"渲染为占位色块"这一类问题。

本脚本的核心价值就是识别这类图标：GetImageMso 不报错、返回了对象，
但那个对象画不出任何图形，功能区上呈现为无意义色块。
没有参照就做不了这个判断，所以【不允许继续并报通过】。

请确认参照 ID 仍然是一个"存在但无位图"的 idMso，
或用 -PlaceholderRefId 指定另一个。
"@
    }
    Write-Host "    占位图参照：$PlaceholderRefId" -ForegroundColor DarkGray
    # 参照本身不在被检之列，别让它混进总览图
    Remove-Item -LiteralPath $refPath -Force -ErrorAction SilentlyContinue

    Write-Host ""
    foreach ($id in $distinct) {
        $owners = ($uses | Where-Object { $_.ImageMso -eq $id } |
                   Select-Object -ExpandProperty ControlId) -join ", "

        if ($missing -contains $id) {
            Write-Host ("  不存在    {0}   <- {1}" -f $id, $owners) -ForegroundColor Red
            $bad += [pscustomobject]@{ ImageMso = $id; 问题 = "不存在"; Controls = $owners }
            continue
        }

        $h = Get-IconHash (Join-Path $iconDir "$id.bmp")
        if (-not $h) {
            Write-Host ("  导不出    {0}   <- {1}" -f $id, $owners) -ForegroundColor Red
            $bad += [pscustomobject]@{ ImageMso = $id; 问题 = "导不出"; Controls = $owners }
        }
        elseif ($h -eq $placeholderHash) {
            Write-Host ("  占位色块  {0}   <- {1}" -f $id, $owners) -ForegroundColor Red
            $bad += [pscustomobject]@{ ImageMso = $id; 问题 = "渲染为占位色块"; Controls = $owners }
        }
        else {
            Write-Host ("  OK        {0}" -f $id) -ForegroundColor DarkGreen
            $report += [pscustomobject]@{ ImageMso = $id; Hash = $h.Substring(0,16); Controls = $owners }
        }
    }

    # 两个不同的 ID 画出一模一样的图，多半也是某种占位/兜底，值得看一眼。
    # 这里只提示，不判失败——确实存在合法的"同图不同 ID"。
    $dupes = Get-ChildItem $iconDir -Filter *.bmp -ErrorAction SilentlyContinue |
             ForEach-Object { [pscustomobject]@{ Name = $_.BaseName; Hash = (Get-IconHash $_.FullName) } } |
             Group-Object Hash | Where-Object { $_.Count -gt 1 }
    foreach ($d in $dupes) {
        Write-Host ("  提示：这几个图标完全相同 -> {0}" -f (($d.Group.Name) -join ", ")) -ForegroundColor DarkYellow
    }

    $sheet = New-IconContactSheet $iconDir
    if ($sheet) {
        Write-Host ""
        Write-Host "总览图：$sheet" -ForegroundColor Cyan
        Write-Host "（自动判据只能筛掉明显的占位色块，语义对不对还得自己看这张图）" -ForegroundColor DarkGray
    }

    # 【把结果落成仓库里可核对的文件】。
    # "我跑过了，30 个都通过"这种口头结论没法复核；把宿主版本、每个 ID
    # 和它的图像哈希写进文档，谁都能在自己机器上重跑一遍对比。
    if ($bad.Count -eq 0) {
        $lines = @()
        $lines += "# 图标验收记录"
        $lines += ""
        $lines += "本文件由 ``tests\check-imagemso.ps1`` 自动生成，请勿手工编辑。"
        $lines += ""
        $lines += "记录的是「每个 imageMso 在该宿主上导出的 32×32 位图的 SHA-256 前 16 位」。"
        $lines += "换一台机器重跑，哈希不一致就说明那台机器的图标集不同——"
        $lines += "这正是需要在目标版本 Excel 上实测的原因。"
        $lines += ""
        $lines += "- 宿主：``$hostDesc``"
        $lines += "- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $lines += "- 占位图参照：``$PlaceholderRefId``（哈希 ``$($placeholderHash.Substring(0,16))``）"
        $lines += "- 结论：$($report.Count) 个图标全部存在且非占位色块"
        $lines += ""
        $lines += "| imageMso | 图像哈希 | 使用它的控件 |"
        $lines += "|---|---|---|"
        foreach ($r in ($report | Sort-Object ImageMso)) {
            $lines += "| ``$($r.ImageMso)`` | ``$($r.Hash)`` | $($r.Controls) |"
        }

        $reportPath = Join-Path $RepoRoot "docs\图标验收记录.md"
        Set-Content -LiteralPath $reportPath -Value $lines -Encoding UTF8
        Write-Host "验收记录：$reportPath" -ForegroundColor Cyan
    }
}
finally {
    Close-ExcelInstance $xl
}

Write-Host ""
if ($bad.Count -eq 0) {
    Write-Host "全部 $($distinct.Count) 个 imageMso 在这台机器上都存在且能正常渲染。" -ForegroundColor Green
    exit 0
}

Write-Host "有 $($bad.Count) 个 imageMso 有问题，对应按钮会显示为空白或无意义色块：" -ForegroundColor Red
$bad | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Host "换成这台 Excel 上确实能渲染的图标 ID，再重跑本脚本。" -ForegroundColor Yellow
exit 1
