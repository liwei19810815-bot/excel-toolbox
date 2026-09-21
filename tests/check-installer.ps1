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
function New-Stage([switch]$WithAI, [string]$Gateway = "https://192.168.1.50:8443") {
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

$stages = @()
try {
    # 前置：本机必须是干净的，否则后面的断言全都不可信
    Section "前置状态"
    $wefCacheExistedBefore = Test-Path $WefCache
    Assert-Equal 0 (@(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue)).Count "开始前没有已安装的加载宏"
    Assert-Equal 0 (Get-WefEntryCount) "开始前没有本项目的 WEF 注册项"

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
    # 保险起见再清一次：断言失败时上面的卸载可能没跑到
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
    foreach ($s in $stages) { Remove-Item $s -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
Write-Host "通过 $script:pass / 失败 $script:fail" -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
