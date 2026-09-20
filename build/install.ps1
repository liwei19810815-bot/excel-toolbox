<#
.SYNOPSIS
    把 dist\ExcelToolbox.xlam 部署到用户加载宏目录并在 Excel 中激活。

.DESCRIPTION
    通过 Excel 的 AddIns 集合注册（等价于「开发工具 → Excel 加载项 → 浏览 → 勾选」），
    比直接写 HKCU\...\Excel\Options 的 OPENn 键更稳妥。

.PARAMETER Uninstall
    取消勾选并从加载宏目录删除。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\install.ps1
    powershell -ExecutionPolicy Bypass -File build\install.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam",
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot  = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Excel 而不是 WPS。安装尤其不能装错宿主：
# 装到 WPS 的加载项目录里，用户打开 Excel 什么都不会看到，还很难排查。
. (Join-Path $PSScriptRoot "_ExcelHost.ps1")
$SrcXlam   = Join-Path $RepoRoot "dist\$OutputName"
$AddInsDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$DestXlam  = Join-Path $AddInsDir $OutputName

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    throw "检测到 Excel 正在运行。请先完全关闭 Excel 再执行安装/卸载。"
}

$xl = New-RealExcel
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # 找到已注册的同名加载宏
    $existing = $null
    foreach ($a in $xl.AddIns) {
        if ($a.Name -eq $OutputName) { $existing = $a; break }
    }

    if ($Uninstall) {
        if ($existing) {
            Write-Step "取消勾选加载宏"
            $existing.Installed = $false
        }
        if (Test-Path $DestXlam) {
            Write-Step "删除 $DestXlam"
            Remove-Item $DestXlam -Force
        }
        Write-Host "已卸载。" -ForegroundColor Green
    }
    else {
        if (-not (Test-Path $SrcXlam)) {
            throw "找不到 $SrcXlam。请先运行 build\build.ps1。"
        }
        if (-not (Test-Path $AddInsDir)) { New-Item -ItemType Directory -Path $AddInsDir | Out-Null }

        # 已勾选时先取消，否则文件被占用无法覆盖
        if ($existing -and $existing.Installed) { $existing.Installed = $false }

        Write-Step "复制到 $AddInsDir"
        Copy-Item $SrcXlam $DestXlam -Force

        Write-Step "在 Excel 中激活"
        if ($existing) {
            $existing.Installed = $true
        } else {
            $xl.AddIns.Add($DestXlam, $false).Installed = $true
        }

        Write-Host ""
        Write-Host "安装完成。打开 Excel，应看到「工具箱」选项卡。" -ForegroundColor Green
    }
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
