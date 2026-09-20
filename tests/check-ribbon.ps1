<#
.SYNOPSIS
    验证 customUI14.xml 真的被 Excel 接受并加载了。

.DESCRIPTION
    为什么要单独一个脚本：这是整套验证里唯一需要【可见 Excel】的检查。
    customUI 的 onLoad 只有在 Excel 真正创建功能区时才会触发，无界面模式下
    永远得不到信号。而可见模式会真的开窗口、加载功能区，COM 调用时序明显更
    不稳定——把它和功能测试混在一起，会让整套测试变得时灵时不灵。

    所以功能测试跑无界面（tests\run-tests.ps1），这一条单独跑。

    这个检查的价值：customUI14.xml 里只要有一处错误（重复 id、无效属性、
    坏掉的 imageMso），Excel 就会【静默】丢掉整个选项卡——VBA 照样编译通过，
    测试照样全绿，但用户打开 Excel 什么按钮都看不到。

    调用方需要套 timeout。退出码 0 = 通过。

.EXAMPLE
    timeout 180 powershell -ExecutionPolicy Bypass -File tests\check-ribbon.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$ok = $false
$xl = $null
try {
    Write-Host "==> 启动可见的 Excel（功能区只有在这种模式下才会创建）" -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    Write-Host "==> 等待 onLoad 触发" -ForegroundColor Cyan
    $sc = ""
    for ($i = 1; $i -le 15; $i++) {
        Start-Sleep -Milliseconds 800
        $sc = $xl.Run("'$OutputName'!Toolbox_SelfCheck")
        if ($sc -like "*ribbon=True") { $ok = $true; break }
    }

    Write-Host "    $sc" -ForegroundColor DarkGray
    if ($ok) {
        Write-Host ""
        Write-Host "Ribbon 加载正常：customUI14.xml 已被 Excel 接受。" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "Ribbon 未加载。customUI14.xml 里多半有错误，Excel 静默丢弃了整个选项卡。" -ForegroundColor Red
        Write-Host "排查：Excel 选项 -> 高级 -> 常规 -> 勾选「显示加载项用户界面错误」后重新打开加载宏。" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "检查过程异常：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($xl) {
        try { $xl.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }
}

if ($ok) { exit 0 } else { exit 1 }
