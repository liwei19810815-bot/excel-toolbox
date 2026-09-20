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

    # 【先把载荷放好，最后才动清单】。
    #
    # 顺序反了的话会出现一个致命窗口：清单已经指向新版本，载荷却还没复制完，
    # 这期间启动 Excel 的同事会去拉一个半成品。
    # 所以这里保证：清单指向的版本，一定已经完整可用。
    Write-Step "上传载荷 $payloadName"

    # 先传到唯一临时名，再原子改名。直接往最终名写的话，
    # 客户端可能在复制途中就把它读走了。
    $tempName = "$payloadName.$([guid]::NewGuid().ToString('N')).uploading"
    $tempPath = Join-Path $SharePath $tempName

    try {
        Copy-Item $Xlam $tempPath

        # 校验大小一致，确认没传残
        $srcSize = (Get-Item $Xlam).Length
        $dstSize = (Get-Item $tempPath).Length
        if ($srcSize -ne $dstSize) {
            throw "上传后大小不一致（源 $srcSize / 目标 $dstSize），已中止。"
        }

        # 临时文件和最终文件【在同一个目录】，所以这是一次重命名，不是跨卷复制，
        # 不存在"改名改到一半"的中间状态。
        # 注意这条保证止于 SMB 服务器端的文件系统（NTFS 上成立）；
        # 若把发布目录放在不保证原子重命名的存储上，需要另行评估。
        Move-Item $tempPath $payloadPath

        # 后置校验：改名之后再确认一次，别只相信没抛异常
        if (-not (Test-Path $payloadPath)) { throw "改名后目标文件不存在，发布失败。" }
        $finalSize = (Get-Item $payloadPath).Length
        if ($finalSize -ne $srcSize) {
            throw "改名后大小不一致（源 $srcSize / 目标 $finalSize），发布失败。"
        }
    }
    catch {
        if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
        throw
    }
}

Write-Step "更新清单"
# 清单本身也用"临时文件 + 改名"替换。
# 直接覆盖的话，客户端可能读到写了一半的空内容——虽然加载器会当成
# "清单非法"退回缓存，但那是把运维失误当成了正常降级，不该依赖它。
$manifestTemp   = Join-Path $SharePath ("manifest.$([guid]::NewGuid().ToString('N')).tmp")
$manifestBackup = Join-Path $SharePath ("manifest.$([guid]::NewGuid().ToString('N')).bak")
$hadOldManifest = $false

try {
    # 先把旧清单备份出来。校验是在替换【之后】做的，不留后路的话，
    # 一旦读回不对，发布目录就停在一个说不清的状态上——而清单是整套机制的
    # 唯一真相来源，它不确定就等于全公司的客户端都不确定。
    #
    # 备份动作本身也放在 try 里：备份失败要能顺手清掉残留的 .bak。
    if (Test-Path $manifestPath) {
        Copy-Item $manifestPath $manifestBackup -Force
        $hadOldManifest = $true
    }

    [System.IO.File]::WriteAllText($manifestTemp, $Version, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item $manifestTemp $manifestPath -Force

    # 后置校验：读回来确认写对了，不只相信"没抛异常"
    $readBack = (Get-Content $manifestPath -Raw -Encoding UTF8).Trim()
    if ($readBack -ne $Version) {
        throw "清单写入校验失败（期望 $Version，读回 $readBack）。"
    }

    if ($hadOldManifest) { Remove-Item $manifestBackup -Force -ErrorAction SilentlyContinue }
}
catch {
    if (Test-Path $manifestTemp) { Remove-Item $manifestTemp -Force -ErrorAction SilentlyContinue }

    # 还原前先确认【现在的清单确实还是我们刚写的那份】。
    # 两个人同时发布时，另一个人可能已经成功把清单更新到了更新的版本；
    # 这时候拿我们的旧备份盖回去，等于把别人成功的发布给回滚了。
    if ($hadOldManifest -and (Test-Path $manifestBackup)) {
        $current = ""
        try { $current = (Get-Content $manifestPath -Raw -Encoding UTF8).Trim() } catch { }

        if ($current -eq $Version -or $current -eq "") {
            try {
                Move-Item $manifestBackup $manifestPath -Force
                Write-Host "    清单已还原到发布前的内容" -ForegroundColor Yellow
            }
            catch {
                Write-Host "    清单还原失败，请手工检查 $manifestPath" -ForegroundColor Red
            }
        } else {
            Remove-Item $manifestBackup -Force -ErrorAction SilentlyContinue
            Write-Host "    清单已被其他发布进程更新为 $current，不做还原" -ForegroundColor Yellow
        }
    }
    throw
}

Write-Host ""
Write-Host "已发布：$Version" -ForegroundColor Green
Write-Host "客户端下次打开 Excel 时自动生效，无需任何操作。" -ForegroundColor DarkGray
Write-Host ""
Write-Host "当前发布目录内容：" -ForegroundColor DarkGray
Get-ChildItem $SharePath -Filter "ExcelToolbox_*.xlam" | Sort-Object Name |
    ForEach-Object { Write-Host ("  " + $_.Name) -ForegroundColor DarkGray }
Write-Host ("  manifest.txt -> " + $Version) -ForegroundColor DarkGray
