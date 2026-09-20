<#
.SYNOPSIS
    把当前构建的工具箱发布到共享目录，供瘦加载器自动分发。

.DESCRIPTION
    发布 = 往共享目录【新增】一个版本化的载荷文件，然后更新清单。

    全程不覆盖任何正在使用的文件，所以【不需要要求大家关掉 Excel】——
    这正是版本化文件名的意义：加载器占的是本地缓存的锁，共享盘上的
    旧版本文件即使还被某些客户端读着，也不影响我们新增一个文件。

    清单写的是"当前应该用哪个版本"。把它改回旧版本号就能回滚全公司，
    客户端下次启动自动退回去，不需要动任何一台机器。

.PARAMETER SharePath
    发布目录（UNC 或本地路径）。

.PARAMETER Version
    版本号。不传就从 src\code\Core\modApp.bas 的 APP_VERSION 里读。

.PARAMETER Rollback
    只把清单改成指定版本，不上传文件。用于回滚到已发布过的版本。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\publish.ps1 -SharePath "\\fs01\tools\ExcelToolbox"
    powershell -ExecutionPolicy Bypass -File build\publish.ps1 -SharePath "\\fs01\tools\ExcelToolbox" -Version 1.0.0 -Rollback
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SharePath,
    [string]$Version = "",
    [switch]$Rollback,
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

# 版本号从源码里读，避免和加载器记录的版本对不上
if (-not $Version) {
    $appSrc = Get-Content (Join-Path $RepoRoot "src\code\Core\modApp.bas") -Raw -Encoding UTF8
    $m = [regex]::Match($appSrc, 'APP_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
    if (-not $m.Success) { throw "无法从 modApp.bas 读出 APP_VERSION，请用 -Version 显式指定。" }
    $Version = $m.Groups[1].Value
}

# 版本号会被拼进文件路径，加载器那边也会再过滤一次，这里先挡住
if ($Version -notmatch '^[0-9.]+$') {
    throw "版本号只允许数字和点：$Version"
}

if (-not (Test-Path $SharePath)) { throw "发布目录不存在或不可达：$SharePath" }

$payloadName = "ExcelToolbox_$Version.xlam"
$payloadPath = Join-Path $SharePath $payloadName
$manifestPath = Join-Path $SharePath "manifest.txt"

if ($Rollback) {
    if (-not (Test-Path $payloadPath)) {
        throw "回滚目标不存在：$payloadName（只能回滚到已发布过的版本）"
    }
    Write-Step "回滚清单到 $Version"
} else {
    if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

    if (Test-Path $payloadPath) {
        # 覆盖已发布的版本是危险操作：已经拿到该版本的客户端【不会再下载】
        # （它们记录的版本号已经一致），于是同一个版本号在不同机器上内容不同，
        # 出了问题根本无法追溯。正确做法是发一个新版本号。
        throw "$payloadName 已存在。同一版本号不要重复发布——请先把 APP_VERSION 递增后重新构建。"
    }

    Write-Step "上传载荷 $payloadName"
    Copy-Item $Xlam $payloadPath
}

Write-Step "更新清单"
# 加载器按"版本号不相等"判断是否更新，所以清单就是唯一的真相来源
[System.IO.File]::WriteAllText($manifestPath, $Version, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ""
Write-Host "已发布：$Version" -ForegroundColor Green
Write-Host "客户端下次打开 Excel 时自动生效，无需任何操作。" -ForegroundColor DarkGray
Write-Host ""
Write-Host "当前发布目录内容：" -ForegroundColor DarkGray
Get-ChildItem $SharePath -Filter "ExcelToolbox_*.xlam" | Sort-Object Name |
    ForEach-Object { Write-Host ("  " + $_.Name) -ForegroundColor DarkGray }
Write-Host ("  manifest.txt -> " + $Version) -ForegroundColor DarkGray
