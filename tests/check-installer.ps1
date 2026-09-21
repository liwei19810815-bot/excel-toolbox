<#
.SYNOPSIS
    安装器的端到端回归：安装、卸载、参数校验、失败回滚。

.DESCRIPTION
    为什么需要这个：

    内网安装包的核心风险恰恰在【注册表 / 证书 / manifest / 卸载副作用】
    这些安装器行为上，而这些东西没有一条会被 VBA 的测试套件覆盖。
    此前这部分一直是手工跑一遍、口头说"实测通过"——
    评审者拿到仓库时看不到任何证据，只能选择信或不信。
    这个脚本把那些手工步骤变成会红的断言，并落进 docs\验收记录.md。

    【本脚本会真的改本机状态】，但全部是安装器自己会改的那些，
    并且每条用例跑完都还原：
      %APPDATA%\Microsoft\AddIns\ExcelToolbox.xlam
      %LOCALAPPDATA%\ExcelToolbox\
      HKCU\...\WEF\Developer 下本项目的那一个值

    【不碰的东西】：
      受信任位置  —— 全程加 -NoTrustedLocation，不改 Office 安全设置
      CA 证书     —— 夹具里不放 ca.crt，不碰用户的受信任根存储
      别人的东西  —— 断言里专门检查 WEF 缓存目录没被删（那是本项目
                     踩过的一个真 bug：卸载时清空了所有 Office.js
                     加载项的缓存，包括别的公司装的）

.EXAMPLE
    timeout 900 powershell -ExecutionPolicy Bypass -File tests\check-installer.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# 【必须设】：被测脚本自己把输出编码设成了 UTF-8，父进程若按系统代码页
# （中文机器上是 GBK）去解码，收到的全是乱码，所有中文断言一律失败——
# 而失败原因看起来像"功能没实现"，极具误导性。
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot = Split-Path -Parent $PSScriptRoot
$InstallDir = Join-Path $RepoRoot "install"
$Xlam = Join-Path $RepoRoot "dist\ExcelToolbox.xlam"
if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$AddInsDir    = Join-Path $env:APPDATA "Microsoft\AddIns"
$CacheDir     = Join-Path $env:LOCALAPPDATA "ExcelToolbox"
$AITargetDir  = Join-Path $CacheDir "ai"
$WefDeveloper = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$WefCache     = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
$TelemetryDir = Join-Path $CacheDir "telemetry"

# 【缓存目录存在 ≠ 装过工具箱】。%LOCALAPPDATA%\ExcelToolbox 被三样东西共用：
# 自动更新的载荷缓存、AI 的 ai\、以及【加载宏运行时写的 telemetry\ 缓冲】。
# 遥测缓冲只要有人用过工具箱就会有，跟装没装没关系——
# 拿"目录存在"当判据，会把只跑过遥测套件的机器误判成"已安装"，
# 于是本脚本拒绝运行（实测就是这么被 run-all 卡住的）。
# 判据只认【telemetry 以外的内容】。
function Test-InstallCachePresent {
    if (-not (Test-Path -LiteralPath $CacheDir)) { return $false }
    return @(Get-ChildItem -LiteralPath $CacheDir -Force -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -ne "telemetry" }).Count -gt 0
}

$script:pass = 0
$script:fail = 0

function Section($n) { Write-Host ""; Write-Host "== $n ==" -ForegroundColor Cyan }
function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected] 实际 [$actual]" -ForegroundColor Red
    }
}
function Assert-True($c, [string]$w) { Assert-Equal $true ([bool]$c) $w }
function Assert-Match($actual, [string]$pattern, [string]$what) {
    if ("$actual" -like $pattern) { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望匹配 [$pattern]" -ForegroundColor Red
        Write-Host "        实际     [$($actual -replace "`r?`n", ' / ')]" -ForegroundColor Red
    }
}

function Get-WefEntryCount {
    try {
        $p = Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue
        if (-not $p) { return 0 }
        return @($p.PSObject.Properties.Name | Where-Object { $_ -like "*ExcelToolbox*" }).Count
    } catch { return 0 }
}

# 夹具：一个模拟的分发包
function New-Stage([switch]$WithAI, [string]$Gateway = "https://192.168.1.50:8443", [switch]$BogusCA) {
    $stage = Join-Path ([IO.Path]::GetTempPath()) ("tbinst_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stage | Out-Null
    Copy-Item (Join-Path $InstallDir "Install-Toolbox.ps1") $stage
    Copy-Item (Join-Path $InstallDir "Excel工具箱.bat") $stage
    Copy-Item $Xlam $stage
    if ($WithAI) {
        $ai = Join-Path $stage "ai"
        New-Item -ItemType Directory -Path $ai | Out-Null
        Copy-Item (Join-Path $InstallDir "ai\manifest.template.xml") $ai
        # 【不放 ca.crt】：那会改用户的受信任根存储，测试不该碰
        Set-Content -LiteralPath (Join-Path $ai "gateway.txt") -Value $Gateway -Encoding UTF8

        # 【故意放一份坏证书】：内容不是证书，certutil 必然失败并退出非零，
        # 【而且不会往受信任根存储里放进任何东西】——这正是我们要的：
        # 用一个绝对安全的方式触发"最后一步失败"，去验证回滚。
        if ($BogusCA) {
            Set-Content -LiteralPath (Join-Path $ai "ca.crt") `
                        -Value "this is not a certificate" -Encoding ASCII
        }
    }
    return $stage
}

# 【参数名不能叫 $Input】。它是 PowerShell 的自动变量（代表管道输入），
# 用作参数名时传进来的值会被管道语义覆盖掉——表现是所有按键都丢失、
# 被测脚本读到 EOF 后走默认分支，于是每个用例都变成"直接回车安装"，
# 而失败信息看起来像是"菜单没实现"，极具误导性。
function Invoke-Installer([string]$Stage, [string]$Keys, [string[]]$ExtraArgs) {
    $script = Join-Path $Stage "Install-Toolbox.ps1"
    $psArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $script, "-NoTrustedLocation") + $ExtraArgs
    return ($Keys | & powershell @psArgs 2>&1 | Out-String)
}

#-----------------------------------------------------------------------------
# 【前置检查必须在动任何东西之前，而且不干净就直接中止】。
#
# 本脚本的 finally 会按名字模式删加载宏、缓存目录和 WEF 注册项。
# 如果这台机器上用户本来就装着工具箱，那些删除【删的就是他的东西】——
# 一个测试脚本把生产环境的安装毁掉，比它能发现的任何 bug 都严重。
#
# 所以这里用 throw 而不是断言：断言只是记一笔失败然后继续往下跑，
# 照样会走到 finally 的清理。必须在创建任何状态之前就退出。
#-----------------------------------------------------------------------------
$preExistingAddins = @(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue)
$preExistingWef    = Get-WefEntryCount
$preExistingCache  = Test-InstallCachePresent

if ($preExistingAddins.Count -gt 0 -or $preExistingWef -gt 0 -or $preExistingCache) {
    Write-Host ""
    Write-Host "拒绝运行：这台机器上已经装了工具箱或 AI 助手。" -ForegroundColor Red
    Write-Host ""
    Write-Host "本脚本会反复安装/卸载并在结束时清理，那些清理会删掉你现有的安装。" -ForegroundColor Yellow
    Write-Host "检测到的现有状态：" -ForegroundColor Yellow
    if ($preExistingAddins.Count -gt 0) { Write-Host "    加载宏：$(($preExistingAddins.Name) -join ', ')" -ForegroundColor Yellow }
    if ($preExistingWef -gt 0)          { Write-Host "    WEF 注册项：$preExistingWef 个" -ForegroundColor Yellow }
    if ($preExistingCache)              { Write-Host "    缓存目录里有安装内容：$CacheDir" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "请先卸载（双击 install\Excel工具箱.bat 选「全部卸载」）再跑本脚本。" -ForegroundColor Yellow
    exit 2
}

$cleanupAllowed = $true      # 走到这里说明起点是干净的，finally 清理才安全

# 【先把遥测缓冲挪走】。本脚本会反复调用真正的卸载流程，而卸载是
# 整个删掉 %LOCALAPPDATA%\ExcelToolbox ——那里面有这台机器上尚未上报的
# 遥测数据。测试不该顺手毁掉它，所以先搬到临时目录，结束时再放回去。
$telemetryStash = $null
if (Test-Path -LiteralPath $TelemetryDir) {
    $telemetryStash = Join-Path ([IO.Path]::GetTempPath()) ("tbtelem_" + [guid]::NewGuid().ToString("N"))
    try {
        Move-Item -LiteralPath $TelemetryDir -Destination $telemetryStash -Force -ErrorAction Stop
    } catch {
        # 挪不走就别往下跑：继续跑等于明知会删掉它还照删
        Write-Host "拒绝运行：无法暂存遥测缓冲（$TelemetryDir）：$($_.Exception.Message)" -ForegroundColor Red
        exit 2
    }
}

$stages = @()
try {
    Section "前置状态"
    $wefCacheExistedBefore = Test-Path $WefCache
    Assert-Equal 0 $preExistingAddins.Count "开始前没有已安装的加载宏"
    Assert-Equal 0 $preExistingWef          "开始前没有本项目的 WEF 注册项"

    #=========================================================================
    Section "不带 AI 包：行为与加入 AI 组件之前一致"

    $s1 = New-Stage; $stages += $s1
    $out = Invoke-Installer $s1 "" @()
    Assert-Match $out "*安装工具箱*"   "菜单只提供安装工具箱一项"
    Assert-True  ($out -notlike "*AI 助手*") "不带 AI 包时菜单不提 AI"
    Assert-Match $out "*安装完成*"     "回车即完成安装"
    Assert-Equal 1 (@(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*.xlam" -ErrorAction SilentlyContinue)).Count "加载宏已就位"

    $out = Invoke-Installer $s1 "2" @()
    Assert-Match $out "*卸载完成*" "卸载完成"
    Assert-Equal 0 (@(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue)).Count "加载宏已清除"

    #=========================================================================
    Section "带 AI 包：菜单与按用户名生成 manifest"

    $s2 = New-Stage -WithAI; $stages += $s2
    $out = Invoke-Installer $s2 "0" @()
    Assert-Match $out "*全部安装*"   "菜单提供全部安装"
    Assert-Match $out "*只装工具箱*" "菜单提供只装工具箱"
    Assert-Match $out "*只装 AI*"    "菜单提供只装 AI"
    Assert-Match $out "*已取消*"     "选 0 不做任何改动"
    Assert-Equal 0 (Get-WefEntryCount) "取消后没有写入注册项"

    $out = Invoke-Installer $s2 "" @()
    Assert-Match $out "*manifest 已生成*" "manifest 已生成"
    Assert-Match $out "*已注册到 Excel*"  "已写入 WEF 注册项"
    Assert-Equal 1 (Get-WefEntryCount)    "WEF 注册项存在且唯一"

    $manifest = Join-Path $AITargetDir "manifest.xml"
    Assert-True (Test-Path $manifest) "manifest 文件已落盘"

    $mx = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8
    Assert-True ($mx -notmatch '\{\{')          "占位符全部被替换"
    Assert-Match $mx "*?u=$($env:USERNAME)*"    "URL 里带当前用户名"
    Assert-True ([bool]([xml]$mx))              "生成的 manifest 是合法 XML"

    # 示例 GUID 没换掉时必须告警——两个组织用同一个 Id 会互相覆盖，极难排查
    Assert-Match $out "*示例 GUID*" "示例 GUID 触发告警"

    #=========================================================================
    Section "全部卸载：清干净，且不碰别人的东西"

    $out = Invoke-Installer $s2 "2" @()
    Assert-Match $out "*卸载完成*" "卸载完成"
    Assert-Equal 0 (@(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue)).Count "加载宏已清除"
    Assert-Equal 0 (Get-WefEntryCount) "WEF 注册项已清除"
    Assert-Equal $false (Test-Path $AITargetDir) "AI 目录已删除"
    Assert-Equal $false (Test-Path $CacheDir)    "缓存目录已删除"

    # 这一条守的是一个真 bug：卸载曾经清空整个 Wef 目录，
    # 把别的公司、别的项目装的 Office.js 加载项缓存一并删了
    if ($wefCacheExistedBefore) {
        Assert-True (Test-Path $WefCache) "【未误删】Office WEF 缓存目录完好保留"
    } else {
        Write-Host "  跳过  本机原本就没有 WEF 缓存目录，无法验证误删" -ForegroundColor DarkYellow
    }

    #=========================================================================
    Section "网关地址校验"

    $s3 = New-Stage -WithAI -Gateway 'https://bad host&x'; $stages += $s3
    $out = Invoke-Installer $s3 "" @("-AIOnly")
    Assert-Match $out "*不合法*"       "非法网关地址被拦下"
    Assert-Equal 0 (Get-WefEntryCount) "被拦下时没有写入注册项"
    Assert-Equal $false (Test-Path (Join-Path $AITargetDir "manifest.xml")) "被拦下时没有生成 manifest"

    $s4 = New-Stage -WithAI -Gateway 'http://192.168.1.50:8080'; $stages += $s4
    $out = Invoke-Installer $s4 "" @("-AIOnly")
    Assert-Match $out "*http 而不是 https*" "非 https 网关触发告警"
    # 告警不等于拒绝：装还是要装上的
    Assert-Equal 1 (Get-WefEntryCount) "告警后仍完成注册"
    $null = Invoke-Installer $s4 "2" @()

    #=========================================================================
    # 这一节守的是一类"没人会手工复现"的状态：安装走到最后一步才失败。
    # 之前的回滚只做到"不删新写的东西"，重装场景下等于把用户原来
    # 能用的那份配置改坏了又不还原。这里用一份坏证书逼出那条路径。
    Section "最后一步失败时的回滚"

    # --- 情况一：机器上本来就没装过，失败后必须不留痕 ---
    $s6 = New-Stage -WithAI -BogusCA; $stages += $s6
    $out = Invoke-Installer $s6 "" @("-AIOnly")
    Assert-Match $out "*CA 证书安装失败*" "坏证书让最后一步失败"
    Assert-Match $out "*已回滚*"          "失败后声明已回滚"
    Assert-Equal 0 (Get-WefEntryCount)    "全新安装失败后不留注册项"
    Assert-Equal $false (Test-Path (Join-Path $AITargetDir "manifest.xml")) "全新安装失败后不留 manifest"

    # --- 情况二：机器上已有一份能用的配置，失败后必须还原成原样 ---
    # 先装一份好的（不带 ca.crt，所以不碰证书存储）
    $s7 = New-Stage -WithAI; $stages += $s7
    $null = Invoke-Installer $s7 "" @("-AIOnly")
    $manifest2 = Join-Path $AITargetDir "manifest.xml"
    Assert-True (Test-Path $manifest2) "回滚用例的前置安装已就位"

    # 把它改成可识别的"用户原有配置"。manifest 仍是合法 XML，
    # 免得将来有人加了 XML 校验之后这条用例变成假通过。
    $sentinelXml = "<OfficeApp><Id>SENTINEL-ORIGINAL</Id></OfficeApp>"
    [IO.File]::WriteAllText($manifest2, $sentinelXml, [Text.UTF8Encoding]::new($false))
    $sentinelReg = "$AITargetDir|SENTINEL"
    New-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Value $sentinelReg `
                     -PropertyType String -Force | Out-Null

    # 再用坏证书重装一次：会覆盖上面两样，然后在最后一步失败
    $s8 = New-Stage -WithAI -BogusCA; $stages += $s8
    $out = Invoke-Installer $s8 "" @("-AIOnly")
    Assert-Match $out "*CA 证书安装失败*" "重装同样在最后一步失败"

    $after = [IO.File]::ReadAllText($manifest2, [Text.UTF8Encoding]::new($false))
    Assert-Equal $sentinelXml $after "【已还原】原 manifest 内容被恢复，而不是留着覆盖后的版本"

    $regNow = (Get-ItemProperty -Path $WefDeveloper -Name $AITargetDir -ErrorAction SilentlyContinue).$AITargetDir
    Assert-Equal $sentinelReg $regNow "【已还原】原注册值被恢复，而不是被删掉或留成新值"

    # 收尾：把这一节造出来的状态清掉，别影响后面的用例
    $null = Invoke-Installer $s7 "2" @()
    Assert-Equal 0 (Get-WefEntryCount) "回滚用例收尾后无残留注册项"

    #=========================================================================
    Section "IT 非交互部署"

    $s5 = New-Stage; $stages += $s5
    $out = Invoke-Installer $s5 "" @("-Install")
    Assert-True ($out -notlike "*请输入数字*") "-Install 跳过交互菜单"
    Assert-Match $out "*安装完成*" "-Install 直接完成安装"
    $out = Invoke-Installer $s5 "" @("-Uninstall")
    Assert-Match $out "*卸载完成*" "-Uninstall 直接完成卸载"
    Assert-Equal 0 (Get-WefEntryCount) "非交互卸载后无残留注册项"
}
catch {
    $script:fail++
    Write-Host ""
    Write-Host "测试过程异常：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # 保险起见再清一次：断言失败时上面的卸载可能没跑到。
    #
    # 【只有起点确认干净时才允许清】。$cleanupAllowed 在前置检查通过后才置位；
    # 前置不干净的路径根本走不到这里（那里直接 exit 2）。
    # 这个标志是第二道保险：万一将来有人改动控制流，别让清理逻辑
    # 在一台有真实安装的机器上跑起来。
    if ($cleanupAllowed) {
        try {
            Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            if (Test-Path $CacheDir) { Remove-Item $CacheDir -Recurse -Force -ErrorAction SilentlyContinue }
            $p = Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue
            if ($p) {
                $p.PSObject.Properties.Name | Where-Object { $_ -like "*ExcelToolbox*" } | ForEach-Object {
                    Remove-ItemProperty -Path $WefDeveloper -Name $_ -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {}
    }
    foreach ($s in $stages) { Remove-Item $s -Recurse -Force -ErrorAction SilentlyContinue }

    # 把暂存的遥测缓冲放回去。放不回去要【喊出来】——
    # 静默失败等于悄悄吞掉这台机器上尚未上报的数据。
    if ($telemetryStash -and (Test-Path -LiteralPath $telemetryStash)) {
        try {
            $parent = Split-Path -Parent $TelemetryDir
            if (-not (Test-Path -LiteralPath $parent)) { $null = New-Item -ItemType Directory -Path $parent -Force }
            Move-Item -LiteralPath $telemetryStash -Destination $TelemetryDir -Force -ErrorAction Stop
        } catch {
            Write-Host "警告：遥测缓冲未能还原，它还在 $telemetryStash" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "通过 $script:pass / 失败 $script:fail" -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
