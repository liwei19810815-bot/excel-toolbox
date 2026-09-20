<#
.SYNOPSIS
    在 WPS 表格上实测工具箱的兼容性，输出一份能力矩阵。

.DESCRIPTION
    这个脚本【只在需要验证 WPS 兼容性时手工运行】，不进日常回归。

    为什么单独一个脚本：WPS 安装后会接管 Excel 的 COM 注册
    （HKCU/HKCR 两层的 CLSID{00024500-...}\LocalServer32 都被改成 et.exe，
    连 Excel.Application.16 也一样），并且【自称 "Microsoft Excel"】——
    只有 Application.Path 和 Version(12.0) 能认出它。

    所以 WPS 兼容性不能和 Excel 的回归测试混在一起跑：
    要么 WPS 占着注册、Excel 测试全废；要么反过来。
    正确做法是平时让 Excel 拿着注册，需要验 WPS 时临时切过去跑一次这个脚本。

    脚本不做任何断言，只如实汇报"这条在 WPS 上行不行"——
    结论要拿去更新 docs\兼容性.md，而不是让 CI 去判定成败。

.EXAMPLE
    timeout 600 powershell -ExecutionPolicy Bypass -File tests\probe-wps.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Continue"     # 探测脚本要尽量跑完，不能一错就停
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"

# 和 Excel 侧共用同一套 WPS 判据，避免两边各写一套、各错一半
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$preExisting = @(Get-Process EXCEL, et, wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$script:rows = @()
function Record([string]$item, [string]$result, [string]$detail = "") {
    $script:rows += [pscustomobject]@{ 项目 = $item; 结果 = $result; 详情 = $detail }
    $color = if ($result -eq "可用") { "Green" } elseif ($result -eq "不可用") { "Red" } else { "Yellow" }
    Write-Host ("  {0,-22} {1,-8} {2}" -f $item, $result, $detail) -ForegroundColor $color
}

# 探一个命令：跑通且没返回 ERROR 就算可用
function Probe([string]$label, [string]$actionId, [scriptblock]$setup) {
    try {
        if ($setup) { & $setup }
        $r = $app.Run("'$OutputName'!Toolbox_Run", $actionId)
        if ("$r" -like "ERROR:*") { Record $label "不可用" "$r" }
        else { Record $label "可用" "$r" }
    }
    catch { Record $label "不可用" $_.Exception.Message }
}

$app = $null
try {
    Write-Host "==> 连接 WPS 表格" -ForegroundColor Cyan
    # 【有意例外】：这里【不能】用 New-RealExcel——本脚本要的恰恰是 WPS。
    # 直接创建 COM，下面用 Test-IsWpsHost 做反向守卫（拿到真 Excel 反而报错）。
    $app = New-Object -ComObject Excel.Application

    $path = ""; $name = ""; $ver = ""
    try { $path = $app.Path } catch {}
    try { $name = $app.Name } catch {}
    try { $ver  = $app.Version } catch {}

    Write-Host "    Name=[$name] Version=$ver" -ForegroundColor DarkGray
    Write-Host "    Path=$path" -ForegroundColor DarkGray

    # 反向守卫：这个脚本要的是 WPS，拿到真 Excel 说明关联还没切过来
    if (-not (Test-IsWpsHost $app)) {
        throw @"
COM 拿到的不是 WPS，而是 Microsoft Excel（Path=$path）。

本脚本用于在 WPS 上实测兼容性，需要 WPS 临时接管 Excel 的 COM 注册：
WPS → 设置/配置工具 → 兼容设置 → 勾选接管 Office 文件关联，然后重跑。
验完记得切回 Excel，否则日常的 Excel 回归测试都跑不了。
"@
    }

    $app.Visible = $false
    $app.DisplayAlerts = $false

    Write-Host ""
    Write-Host "== 基础加载 ==" -ForegroundColor Cyan

    $wb = $null
    try { $wb = $app.Workbooks.Add() ; Record "新建工作簿" "可用" }
    catch { Record "新建工作簿" "不可用" $_.Exception.Message }

    $addin = $null
    try { $addin = $app.Workbooks.Open($Xlam); Record "打开 .xlam 加载宏" "可用" }
    catch { Record "打开 .xlam 加载宏" "不可用" $_.Exception.Message }

    if ($null -eq $addin) {
        Write-Host ""
        Write-Host "加载宏都打不开，后续无法探测。" -ForegroundColor Red
        return
    }

    # 自检能返回，说明整个 VBA 工程在 WPS 上编译通过了——这是最关键的一关
    try {
        $sc = $app.Run("'$OutputName'!Toolbox_SelfCheck")
        Record "VBA 工程编译" "可用" "$sc"
        if ("$sc" -like "*ribbon=True*") { Record "功能区 customUI" "可用" }
        else { Record "功能区 customUI" "未确认" "无界面模式下 onLoad 可能不触发，需人工看 WPS 里有没有『工具箱』选项卡" }
    }
    catch { Record "VBA 工程编译" "不可用" $_.Exception.Message; return }

    try { $r = $app.Run("'$OutputName'!Toolbox_ProbeHost"); Record "宿主能力探测" "可用" "$r" }
    catch { Record "宿主能力探测" "不可用" $_.Exception.Message }

    try { $app.Run("'$OutputName'!Toolbox_SetSilent", $true) } catch { }

    Write-Host ""
    Write-Host "== 各模块代表性命令 ==" -ForegroundColor Cyan

    $ws = $wb.Worksheets.Item(1)
    $Prep = {
        try {
            $null = $ws.Cells.Clear()
            $ws.Range("A1").Value2 = "  甲  "
            $ws.Range("A3").Value2 = "乙"
            $ws.Range("B1").Value2 = 10
            $ws.Range("B3").Value2 = 20
            $null = $ws.Activate()
            $null = $ws.Range("A1:B3").Select()
        } catch {}
    }

    Probe "M1 清除空格"      "text.cleanSpaces"     $Prep
    Probe "M1 全角转半角"    "text.toHalfWidth"     $Prep
    Probe "M2 删除空行"      "data.deleteEmptyRows" $Prep
    Probe "M2 提取唯一值"    "data.extractUnique"   $Prep
    Probe "M3 生成目录"      "sheet.createIndex"    $null
    Probe "M6 定位错误值"    "formula.findErrors"   $Prep
    Probe "M7 数据体检"      "audit.scan"           $Prep
    Probe "M8 数据条"        "viz.dataBars"         $Prep
    Probe "M8 色阶"          "viz.colorScale"       $Prep
    Probe "M8 迷你图"        "viz.sparklines"       $Prep
    Probe "M9 金额大写"      "misc.amountToChinese" $Prep
    Probe "M9 聚光灯"        "misc.spotlight"       $Prep

    Write-Host ""
    Write-Host "== 撤销框架 ==" -ForegroundColor Cyan
    try {
        & $Prep
        $null = $app.Run("'$OutputName'!Toolbox_Run", "text.cleanSpaces")
        $u = $app.Run("'$OutputName'!Toolbox_Undo")
        if ("$u" -like "ERROR:*") { Record "撤销" "不可用" "$u" }
        elseif ("$($ws.Range('A1').Value2)" -eq "  甲  ") { Record "撤销" "可用" "值已还原" }
        else { Record "撤销" "不可用" "执行了但没还原：[$($ws.Range('A1').Value2)]" }
    }
    catch { Record "撤销" "不可用" $_.Exception.Message }
}
catch {
    Write-Host ""
    Write-Host "探测中断：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($app) {
        try { $app.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($app.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $app.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 600
    Get-Process EXCEL, et, wps -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }
}

Write-Host ""
Write-Host "================ WPS 兼容性汇总 ================" -ForegroundColor Cyan
$script:rows | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Host "把结论更新到 docs\兼容性.md。注意：功能区是否真的出现，" -ForegroundColor DarkGray
Write-Host "必须在 WPS 里肉眼确认一次——无界面模式下 onLoad 不一定触发。" -ForegroundColor DarkGray
