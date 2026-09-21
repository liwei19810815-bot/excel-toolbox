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
function Invoke-Installer([string]$Stage, [string]$Keys, [string[]]$ExtraArgs, [switch]$BreakCertutil) {
    $sc = Join-Path $Stage "Install-Toolbox.ps1"
    $psArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $sc, "-NoTrustedLocation") + $ExtraArgs

    # 【用全路径启动子进程】。下面可能要把 PATH 换掉，
    # 而 `powershell` 这个名字本身就是靠 PATH 解析的。
    $exe = Join-Path $PSHOME "powershell.exe"

    # 【模拟 certutil 根本起不来】。把子进程的 PATH 换成只有夹具目录的值，
    # System32 不在里面，于是 `& certutil` 连解析都失败、直接抛
    # CommandNotFoundException——这跟"certutil 跑起来了但返回非零"
    # 是【两条不同的代码路径】，只有前者能验证 try/catch 兜没兜住。
    # 现实里对应的是 certutil 被组策略禁用、被 AV 拦、或不在 PATH 上。
    $savedPath = $env:PATH
    if ($BreakCertutil) { $env:PATH = $Stage }
    try {
        $out = ($Keys | & $exe @psArgs 2>&1 | Out-String)
        $script:lastInstallerExit = $LASTEXITCODE
    }
    finally {
        $env:PATH = $savedPath
    }
    return $out
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
#
# 【上一次跑崩过就别再跑】。搬走之后如果进程被杀或机器断电，
# 缓冲会永久留在 %TEMP%\tbtelem_*，而原路径空着。这时再跑一次，
# 新的一轮会把"空的原路径"当成正常起点，旧 stash 就再也没人认领了。
# 发现遗留就直接拒绝，让人先把数据挪回去——宁可不跑，也不要悄悄丢。
$staleStash = @(Get-ChildItem -Path ([IO.Path]::GetTempPath()) -Filter "tbtelem_*" `
                              -Directory -ErrorAction SilentlyContinue)
if ($staleStash.Count -gt 0) {
    Write-Host ""
    Write-Host "拒绝运行：发现上一次运行遗留的遥测缓冲暂存目录。" -ForegroundColor Red
    foreach ($d in $staleStash) { Write-Host "    $($d.FullName)" -ForegroundColor Yellow }
    Write-Host "说明上一轮中途崩了。请先把里面的文件挪回 $TelemetryDir" -ForegroundColor Yellow
    Write-Host "（或确认不要了再删掉），然后重跑本脚本。" -ForegroundColor Yellow
    exit 2
}

$telemetryStash = $null
if (Test-Path -LiteralPath $TelemetryDir) {
    # 【重解析点要拒绝】。目录联接/符号链接下 Move-Item 的语义不可靠，
    # 可能搬的是链接本身而不是内容，也可能把内容搬到链接指向的别处。
    # 这种情况下"搬走再放回"的承诺兑现不了，直接不跑。
    $ti = Get-Item -LiteralPath $TelemetryDir -Force
    if ($ti.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        Write-Host "拒绝运行：$TelemetryDir 是符号链接或目录联接，无法安全暂存。" -ForegroundColor Red
        exit 2
    }

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
    # 【退出码要单独断言】。只匹配输出文字的话，安装失败却 exit 0
    # 这种错误在这里完全看不出来，而 IT 的批量部署脚本正是靠退出码
    # 判断该不该重试、该不该告警的。
    Assert-Equal 1 $script:lastInstallerExit "安装失败时退出码为 1"
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

    # --- 情况三：原值是【空的】，回滚必须还原成空，而不是删掉 ---
    # 这条守的是一个纯靠肉眼看不出来的差别：用 `$backup -ne ""` 判断
    # "有没有备份"时，本来就是空值的配置会被当成"没有备份"，
    # 于是回滚去删它。状态对不上，而且不会有任何测试因此变红。
    $null = Invoke-Installer $s7 "" @("-AIOnly")
    [IO.File]::WriteAllText($manifest2, "", [Text.UTF8Encoding]::new($false))
    New-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Value "" `
                     -PropertyType String -Force | Out-Null

    $out = Invoke-Installer $s8 "" @("-AIOnly")
    Assert-Match $out "*CA 证书安装失败*" "空原值场景同样在最后一步失败"
    Assert-Equal "" ([IO.File]::ReadAllText($manifest2, [Text.UTF8Encoding]::new($false))) `
                 "【空值也算原值】空 manifest 被还原成空，而不是留着覆盖后的内容"
    $emptyProps = (Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue)
    Assert-True ($emptyProps -and ($emptyProps.PSObject.Properties.Name -contains $AITargetDir)) `
                "【空值也算原值】空注册值被还原，而不是被删掉"
    Assert-Equal "" $emptyProps.$AITargetDir "空注册值还原后仍是空串"

    # --- 情况四：certutil 根本起不来（抛异常，不是返回非零） ---
    # 只看退出码的写法在这条路径上会直接把异常抛出函数、【绕过回滚】，
    # 留下"manifest 和注册项都写了、证书没装"的半成品。
    $null = Invoke-Installer $s7 "2" @()
    $s9 = New-Stage -WithAI -BogusCA; $stages += $s9
    $out = Invoke-Installer $s9 "" @("-AIOnly") -BreakCertutil
    Assert-Match $out "*无法运行 certutil*" "certutil 起不来时被 try/catch 兜住"
    Assert-Match $out "*已回滚*"            "certutil 起不来时同样触发回滚"
    Assert-Equal 0 (Get-WefEntryCount)      "certutil 起不来后不留注册项"
    Assert-Equal $false (Test-Path (Join-Path $AITargetDir "manifest.xml")) "certutil 起不来后不留 manifest"

    # 收尾：把这一节造出来的状态清掉，别影响后面的用例
    $null = Invoke-Installer $s7 "2" @()
    Assert-Equal 0 (Get-WefEntryCount) "回滚用例收尾后无残留注册项"

    #=========================================================================
    Section "IT 非交互部署"

    $s5 = New-Stage; $stages += $s5
    $out = Invoke-Installer $s5 "" @("-Install")
    Assert-True ($out -notlike "*请输入数字*") "-Install 跳过交互菜单"
    Assert-Match $out "*安装完成*" "-Install 直接完成安装"
    Assert-Equal 0 $script:lastInstallerExit "安装成功时退出码为 0"
    $out = Invoke-Installer $s5 "" @("-Uninstall")
    Assert-Match $out "*卸载完成*" "-Uninstall 直接完成卸载"
    Assert-Equal 0 $script:lastInstallerExit "卸载成功时退出码为 0"
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
    # 【只删本次跑出来的东西，不做按名字模式的通扫】。
    # 原先是"删掉所有 ExcelToolbox*、整个缓存目录、所有含 ExcelToolbox 的
    # WEF 项"。前置门能保证起点是空的，但保证不了运行期间没有别人
    # （另一个用户会话、一台共享机器上的同事）装了东西进来——
    # 通扫会把那些一并删掉。范围收窄到确知是自己造的那几样。
    if ($cleanupAllowed) {
        $cleanupProblems = @()

        # 加载宏：只删前置快照里【没有】的文件（起点为空，所以就是本次新增的）
        try {
            $before = @($preExistingAddins | ForEach-Object { $_.FullName })
            foreach ($f in @(Get-ChildItem $AddInsDir -Filter "ExcelToolbox*" -ErrorAction SilentlyContinue)) {
                if ($before -notcontains $f.FullName) {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                }
            }
        } catch { $cleanupProblems += "加载宏：$($_.Exception.Message)" }

        # 缓存目录：telemetry 已经被暂存走了，这里只清其余内容；
        # 清完若已空则连目录一起删，留着空目录会让下一轮的判据更难写
        try {
            if (Test-Path -LiteralPath $CacheDir) {
                foreach ($c in @(Get-ChildItem -LiteralPath $CacheDir -Force -ErrorAction SilentlyContinue |
                                 Where-Object { $_.Name -ne "telemetry" })) {
                    Remove-Item -LiteralPath $c.FullName -Recurse -Force -ErrorAction Stop
                }
                if (@(Get-ChildItem -LiteralPath $CacheDir -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                    Remove-Item -LiteralPath $CacheDir -Force -ErrorAction SilentlyContinue
                }
            }
        } catch { $cleanupProblems += "缓存目录：$($_.Exception.Message)" }

        # WEF：只删【本程序自己那一条】（键名就是 $AITargetDir），不按模式扫
        try {
            $p = Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties.Name -contains $AITargetDir) {
                Remove-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Force -ErrorAction Stop
            }
        } catch { $cleanupProblems += "WEF 注册项：$($_.Exception.Message)" }

        # 【清理失败不能静默吞掉】。吞掉的话，残留会被下一轮的前置门
        # 当成"用户已有安装"而拒绝运行，排查时完全看不出是上一轮没清干净。
        if ($cleanupProblems.Count -gt 0) {
            $script:fail++
            Write-Host "  FAIL  收尾清理有残留：$($cleanupProblems -join '；')" -ForegroundColor Red
        }
    }
    foreach ($s in $stages) { Remove-Item $s -Recurse -Force -ErrorAction SilentlyContinue }

    # 把暂存的遥测缓冲放回去。
    #
    # 【还不回去必须让测试变红】。只打一行警告的话，就是这个项目一直在防的
    # 那类失败：测试全绿，数据却没了。
    #
    # 【不能直接 Move-Item 到 $TelemetryDir】。如果这中间有真实的 Excel
    # 带着加载宏跑过，那个目录会被重新建出来；此时 Move-Item 的语义是
    # "移动到该目录【里面】"，结果是 telemetry\tbtelem_xxx\... ——
    # 看着成功了，实际上数据埋到了一层没人读的子目录里。
    # 所以目标已存在时逐个文件搬。
    if ($telemetryStash -and (Test-Path -LiteralPath $telemetryStash)) {
        try {
            if (Test-Path -LiteralPath $TelemetryDir) {
                foreach ($f in @(Get-ChildItem -LiteralPath $telemetryStash -Force)) {
                    $dest = Join-Path $TelemetryDir $f.Name
                    # 同名的话保留两份，别替用户决定哪份该留
                    if (Test-Path -LiteralPath $dest) {
                        $dest = Join-Path $TelemetryDir ($f.BaseName + ".restored" + $f.Extension)
                    }
                    Move-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
                }
                Remove-Item -LiteralPath $telemetryStash -Recurse -Force -ErrorAction SilentlyContinue
            }
            else {
                $parent = Split-Path -Parent $TelemetryDir
                if (-not (Test-Path -LiteralPath $parent)) { $null = New-Item -ItemType Directory -Path $parent -Force }
                Move-Item -LiteralPath $telemetryStash -Destination $TelemetryDir -Force -ErrorAction Stop
            }
        } catch {
            $script:fail++
            Write-Host "  FAIL  遥测缓冲未能还原，它还在 $telemetryStash（$($_.Exception.Message)）" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "通过 $script:pass / 失败 $script:fail" -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
