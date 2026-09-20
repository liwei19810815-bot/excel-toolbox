<#
.SYNOPSIS
    瘦加载器（自动更新）的回归测试。

.DESCRIPTION
    在 %TEMP% 下造一个假的"共享目录"，用 -SharePath 单独构建一个指向它的
    测试用加载器，然后走完整条路径：首次安装 → 升级 → 回滚 → 断网降级 →
    清单被写坏 → 真实的 Workbook_Open 启动。

    这几条路径必须自动化验证：它们只在真实环境才会发生，而真出问题时
    受影响的是全公司所有人的 Excel。

    加载器是【无状态】的——不记录"已安装版本"，直接看缓存目录里有什么，
    所以这里也没有任何注册表清理，跑完只还原缓存目录和删临时目录。

    调用方必须套 timeout。退出码 0 = 全过。

.EXAMPLE
    timeout 700 powershell -ExecutionPolicy Bypass -File tests\run-tests-loader.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$Payload    = Join-Path $RepoRoot "dist\ExcelToolbox.xlam"
$LoaderName = "ExcelToolboxLoader.test.xlam"
$Loader     = Join-Path $RepoRoot "dist\$LoaderName"

if (-not (Test-Path $Payload)) { throw "找不到 $Payload。请先运行 build\build.ps1。" }

$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$SandBox  = Join-Path $env:TEMP ("LoaderTest_" + [guid]::NewGuid().ToString("N"))
$Share    = Join-Path $SandBox "share"
$CacheDir = Join-Path $env:LOCALAPPDATA "ExcelToolbox"

# 测试会清空本地缓存目录。先把原有内容挪走、跑完还回去，
# 免得把开发机上已经装好的工具箱缓存弄没了。
$CacheBackup = Join-Path $SandBox "cache-backup"

$script:pass = 0
$script:fail = 0
$script:section = ""

function Section($name) {
    $script:section = $name
    Write-Host ""
    Write-Host "== $name ==" -ForegroundColor Cyan
}
function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected]" -ForegroundColor Red
        Write-Host "        实际 [$actual]" -ForegroundColor Red
    }
}

function Set-Manifest([string]$text) {
    [System.IO.File]::WriteAllText((Join-Path $Share "manifest.txt"), $text,
                                   (New-Object System.Text.UTF8Encoding($false)))
}
function Publish-Version([string]$version) {
    Copy-Item $Payload (Join-Path $Share "ExcelToolbox_$version.xlam") -Force
    Set-Manifest $version
}
function Get-Field([string]$status, [string]$name) {
    $m = [regex]::Match($status, "(^|\|)$name=([^|]*)")
    if ($m.Success) { return $m.Groups[2].Value }
    return ""
}

$xl = $null
try {
    New-Item -ItemType Directory -Path $Share -Force | Out-Null
    New-Item -ItemType Directory -Path $CacheBackup -Force | Out-Null
    Write-Host "假共享目录：$Share" -ForegroundColor DarkGray

    if (Test-Path $CacheDir) {
        Copy-Item (Join-Path $CacheDir "*") $CacheBackup -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item (Join-Path $CacheDir "*") -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 发布目录是构建时写死的，所以测试要单独构建一个指向沙箱的加载器。
    # 这正是生产里的部署方式——测的就是真实机制。
    Write-Host "==> 构建指向沙箱的测试加载器" -ForegroundColor Cyan
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "build\build-loader.ps1") `
        -SharePath $Share -OutputName $LoaderName | Out-Null
    if (-not (Test-Path $Loader)) { throw "测试加载器构建失败。" }

    Publish-Version "9.0.0"

    Write-Host "==> 启动 Excel 并加载【加载器】" -ForegroundColor Cyan
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)

    # 关掉事件再打开：Workbook_Open 里又去 Workbooks.Open 打开载荷，
    # 这种"打开中再打开"在无界面的 COM 场景下会挂住。
    # 更新逻辑用 Loader_CheckNow 显式驱动即可完整覆盖；
    # Workbook_Open 那条真实路径放最后用可见模式单独验。
    $xl.EnableEvents = $false
    $null = $xl.Workbooks.Open($Loader)
    $xl.EnableEvents = $true

    $Status   = { $xl.Run("'$LoaderName'!Loader_Status") }
    $CheckNow = { $xl.Run("'$LoaderName'!Loader_CheckNow") }

    #==========================================================================
    Section "首次安装：本地缓存是空的"

    $null = & $CheckNow
    $st = & $Status
    Write-Host "  $st" -ForegroundColor DarkGray
    Assert-Equal "True"  (Get-Field $st "shareReachable") "能连上发布目录"
    Assert-Equal "9.0.0" (Get-Field $st "remote")      "读到远端清单版本"
    Assert-Equal "9.0.0" (Get-Field $st "cached")      "载荷已下载到本地缓存"
    Assert-Equal "9.0.0" (Get-Field $st "active")      "打开的就是该版本"
    Assert-Equal "True"  (Get-Field $st "payloadOpen") "载荷已打开"

    # 载荷真的能用——不是把文件拷过来就算数
    $sc = $xl.Run("'ExcelToolbox_9.0.0.xlam'!Toolbox_SelfCheck")
    Assert-Equal $true ($sc -like "OK|*") "载荷里的命令可以正常调用"

    #==========================================================================
    Section "升级：清单指向新版本"

    Publish-Version "9.1.0"
    $null = & $CheckNow
    $st = & $Status
    Assert-Equal "9.1.0" (Get-Field $st "active") "自动升级到新版本"
    Assert-Equal $true (Test-Path (Join-Path $CacheDir "ExcelToolbox_9.1.0.xlam")) "新载荷已缓存"

    #==========================================================================
    Section "回滚：清单改回旧版本号"

    # 运维上最关键的能力：线上出问题时改一行清单就能让全公司退回去，
    # 不需要碰任何一台客户端。所以判断依据是"清单说了算"，而不是"版本号更大才更新"。
    Set-Manifest "9.0.0"
    $null = & $CheckNow
    $st = & $Status
    Assert-Equal "9.0.0" (Get-Field $st "active") "清单改回旧版本后确实回滚了"

    #==========================================================================
    Section "断网降级：连不上发布目录"

    $Offline = $Share + "_offline"
    Rename-Item $Share $Offline
    try {
        $null = & $CheckNow
        $st = & $Status
        Assert-Equal "False" (Get-Field $st "shareReachable") "发布目录确实不可达"
        Assert-Equal ""      (Get-Field $st "remote")         "读不到远端清单"
        Assert-Equal "9.1.0" (Get-Field $st "active")         "退回缓存里最新的版本，工具箱照常可用"
        Assert-Equal "True"  (Get-Field $st "payloadOpen")    "断网时不是整个失效"
    }
    finally { Rename-Item $Offline $Share }

    #==========================================================================
    Section "清单被写坏：版本号含非法字符"

    # 版本号会被直接拼进文件路径。不过滤的话，清单里写一个相对路径
    # 就能让加载器去别处取文件执行——这是条现成的注入通道。
    Set-Manifest "..\..\evil"
    $null = & $CheckNow
    $st = & $Status
    Assert-Equal ""      (Get-Field $st "remote") "非法版本号被整体作废"
    Assert-Equal "9.1.0" (Get-Field $st "active") "不会因为清单被写坏就去加载别处的文件"
    Assert-Equal "True"  (Get-Field $st "payloadOpen") "仍然可用"

    #==========================================================================
    Section "真实启动路径：Workbook_Open 自动拉取（可见模式）"

    # 这是"用户无感"的关键路径，必须真的验一次。
    # 需要可见 Excel：无界面时"打开中再打开"会挂住，
    # 和 tests\check-ribbon.ps1 需要可见 Excel 是同一类原因。
    Set-Manifest "9.0.0"

    try { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
    try { $xl.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 800

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Loader)      # 事件开着，Workbook_Open 自己跑

    $st = ""
    for ($i = 1; $i -le 15; $i++) {
        Start-Sleep -Milliseconds 700
        $st = $xl.Run("'$LoaderName'!Loader_Status")
        if ((Get-Field $st "active") -eq "9.0.0") { break }
    }
    Write-Host "  $st" -ForegroundColor DarkGray
    Assert-Equal "9.0.0" (Get-Field $st "active") "打开 Excel 即按清单自动加载，用户无需任何操作"
    Assert-Equal "True"  (Get-Field $st "payloadOpen") "载荷已自动加载"
}
catch {
    $script:fail++
    Write-Host ""
    Write-Host "测试过程异常（$script:section）：$($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  出错行：$($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
}
finally {
    if ($xl) {
        try { $xl.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 600
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }

    # 还原开发机的缓存目录。
    # 刚杀掉的 Excel 释放文件句柄需要一点时间，直接复制会撞上
    # "文件正由另一进程使用"。这里重试几次，并且整段包在 try 里——
    # 清理失败不该把一次全绿的测试判成失败。
    try {
        for ($attempt = 1; $attempt -le 10; $attempt++) {
            $busy = $false
            try {
                if (Test-Path $CacheDir) {
                    Remove-Item (Join-Path $CacheDir "*") -Recurse -Force -ErrorAction Stop
                }
                if (Test-Path $CacheBackup) {
                    Copy-Item (Join-Path $CacheBackup "*") $CacheDir -Recurse -Force -ErrorAction Stop
                }
            }
            catch { $busy = $true }
            if (-not $busy) { break }
            Start-Sleep -Milliseconds 500
        }
    }
    catch { Write-Host "  （缓存目录还原失败，可手工清理 $CacheDir）" -ForegroundColor Yellow }
    if (Test-Path $Loader) { Remove-Item $Loader -Force -ErrorAction SilentlyContinue }
    if (Test-Path $SandBox) { Remove-Item $SandBox -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
Write-Host ("通过 {0} / 失败 {1}" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
