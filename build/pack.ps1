<#
.SYNOPSIS
    组装发给业务用户的【一键安装包】：dist\Excel工具箱_v<版本>.zip

.DESCRIPTION
    打包 = 【从源码重新构建】+ 组装 + 自检 + 出清单。

    为什么默认要重新构建，而不是拿 dist\ 里现成的 .xlam：
    "打了个包，结果里面是上周的构建" 这种事不会让任何测试变红，
    也不会有任何报错——直到用户报告"新功能没有啊"。
    从源码构建一遍的代价是几十秒，换的是包里装的一定是当前代码。
    确实只想复用现有产物时用 -SkipBuild，但那要你自己为它负责。

    两种分发形态，对应两种 -Mode：

      standalone  包里放 ExcelToolbox.xlam 本体。装完就是完整功能，
                  和服务器没有任何关系。适合：人少、没有共享盘、
                  或者给外部/临时用户。缺点是以后每次升级都要重新发包。

      loader      包里放 ExcelToolboxLoader.xlam 瘦加载器。它启动时
                  去共享目录取最新载荷。适合公司内网常规分发：
                  以后升级只需 build\publish.ps1 往共享目录发新载荷，
                  【不用再碰任何一台客户端】。

    【loader 模式的共享目录是构建时写死的】，运行时没有任何办法改
    （实测 Excel 宏里 WScript.Shell 会被安全策略静默拦下，
    表现是"更新莫名其妙不生效"且查不出原因，详见 modLoader.bas）。
    所以换共享目录 = 重新打包 + 每台机器重装一次加载器。
    正因为改不了，这里对 -SharePath 的校验才格外严。

.PARAMETER Mode
    standalone（默认）或 loader。见上。

.PARAMETER SharePath
    loader 模式【必填】：载荷的共享目录，UNC 或本地路径。
    会原样写进加载器。这里会拦掉源码里那个占位值。

.PARAMETER Gateway
    带上就在包里放 ai\ 目录（AI 助手组件），值是内网网关地址。
    不带就不放，安装程序检测不到 ai\ 就不会向用户提这一项。

.PARAMETER SkipBuild
    不重新构建，直接用 dist\ 里现有的产物。见上面的告诫。

.PARAMETER OutDir
    输出目录，默认 dist。

.EXAMPLE
    # 最常见：内网常规分发，以后升级只发载荷不动客户端
    powershell -ExecutionPolicy Bypass -File build\pack.ps1 -Mode loader -SharePath "\\fs01\tools\ExcelToolbox"

.EXAMPLE
    # 独立版 + AI 助手
    powershell -ExecutionPolicy Bypass -File build\pack.ps1 -Gateway "https://ai.corp.example.com"

.NOTES
    【本工具不再接受 CA 证书】。早先支持 -CaCert，把内网自签根证书打进包里
    由安装程序装进用户的受信任根存储——那是降低用户整台机器防护等级的操作，
    影响远不止这一个加载项。现在的前提是网关用【已被客户端信任】的证书
    （域内 PKI 统一下发或公网证书）。若 IT 只能提供自签证书，
    应由 IT 用组策略统一下发根证书，而不是让安装包替用户做这个决定。
#>
[CmdletBinding()]
param(
    [ValidateSet("standalone", "loader")]
    [string]$Mode = "standalone",

    [string]$SharePath = "",
    [string]$Gateway   = "",

    [string]$OutDir    = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# 子脚本会把自己的输出设成 UTF-8；父进程按 GBK 解码的话，
# 下面那些靠输出判断成败的地方会看到乱码
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot   = Split-Path -Parent $PSScriptRoot
$InstallDir = Join-Path $RepoRoot "install"
$DistDir    = Join-Path $RepoRoot "dist"
if (-not $OutDir) { $OutDir = $DistDir }

function Write-Step($m) { Write-Host "`n==> $m" -ForegroundColor Cyan }
function Write-Ok  ($m) { Write-Host "    [完成] $m" -ForegroundColor Green }
function Write-Warn($m) { Write-Host "    [注意] $m" -ForegroundColor Yellow }

$problems = @()
function Add-Problem($m) { $script:problems += $m; Write-Host "    [失败] $m" -ForegroundColor Red }

#-----------------------------------------------------------------------------
# 1. 参数校验 —— 【在构建之前】
#
# 花两分钟构建完才发现网关地址拼错，纯属浪费；更糟的是校验不严时
# 错误配置会被打进包发出去，而那些错误全都是【在用户机器上才显形】的：
# 非法 XML 的 manifest，Office 是【静默不加载】，用户只看到"按钮没出现"。
#-----------------------------------------------------------------------------
Write-Step "校验参数"

if ($Mode -eq "loader") {
    if (-not $SharePath) {
        throw "loader 模式必须给 -SharePath（载荷的共享目录）。这个地址写死在加载器里，运行时改不了。"
    }
    # 源码里的占位值。带着它发出去，全公司的加载器都会去找一台不存在的服务器，
    # 而加载器找不到共享目录时是【静默用本地缓存】——第一次装的人直接没有功能。
    if ($SharePath -match '^\\\\server\\share\\') {
        throw "-SharePath 还是源码里的占位值（$SharePath）。请填真实的共享目录。"
    }
    # 【单个反斜杠开头几乎一定是被外壳吃掉了一个】。实测在 Git Bash 等
    # 非 PowerShell 外壳里调用，`\\fs01\tools\X` 会变成 `\fs01\tools\X`。
    # 这个地址是【写死进加载器、运行时改不了】的，被吃掉一个反斜杠
    # 等于全公司的加载器都指向一个不存在的路径，而加载器找不到共享目录时
    # 是静默用本地缓存——没有任何报错。必须在这里拦死。
    if ($SharePath -match '^\\[^\\]' ) {
        throw @"
-SharePath 以【单个】反斜杠开头：$SharePath
UNC 路径应该是两个反斜杠开头（\\服务器\共享名\...）。
多半是被外壳吃掉了一个——请【在 PowerShell 窗口里】运行本脚本，
并把路径用双引号括起来，例如：
    -SharePath "\\fs01\tools\ExcelToolbox"
"@
    }
    if ($SharePath -notmatch '^(\\\\[^\\]+\\[^\\]+|[A-Za-z]:\\)') {
        throw "-SharePath 看起来不是 UNC 路径或本地路径：$SharePath"
    }
    if (-not (Test-Path -LiteralPath $SharePath)) {
        # 不直接拦死：打包机可能没权限访问发布盘，但地址本身是对的
        Write-Warn "共享目录现在访问不到：$SharePath"
        Write-Warn "如果地址没写错就继续；写错了的话，装到客户端上就再也改不了了。"
    } else {
        Write-Ok "共享目录可达：$SharePath"
    }
}
elseif ($SharePath) {
    Write-Warn "standalone 模式用不到 -SharePath，已忽略。"
}

$withAI = [bool]$Gateway
if ($withAI) {
    # 【和安装脚本用同一条正则】。两边不一致的话，打包时放行、
    # 安装时才拦下，问题就跑到用户机器上去了。
    if ($Gateway -notmatch '^https?://[^\s<>&"'']+$') {
        throw "-Gateway 地址不合法：$Gateway（不能含空格或 < > & 等字符，它会被拼进 XML）"
    }
    if ($Gateway -notmatch '^https://') {
        Write-Warn "网关用的是 http 而不是 https。Office 加载项通常要求 https，任务窗格可能加载不了。"
    }
    $Gateway = $Gateway.TrimEnd('/')
    Write-Ok "网关地址：$Gateway"

    # 【网关证书必须已被客户端信任】，这是走 Office.js 的硬前提：
    # 证书不受信任时任务窗格是【空白且不报错】的，用户只会以为工具坏了。
    # 本工具不再打包 CA、安装程序也不再装证书，所以这一条只能靠部署时保证。
    Write-Warn "请确认网关的证书【已被客户端信任】（域内 PKI 下发或公网证书）。"
    Write-Warn "用自签且未下发根证书的话，任务窗格会空白且不报错——这是最难排查的一种故障。"
}

#-----------------------------------------------------------------------------
# 2. 版本号 —— 从源码读，别手输
#-----------------------------------------------------------------------------
$appSrc = Get-Content (Join-Path $RepoRoot "src\code\Core\modApp.bas") -Raw -Encoding UTF8
$m = [regex]::Match($appSrc, 'APP_VERSION\s+As\s+String\s*=\s*"([^"]+)"')
if (-not $m.Success) { throw "无法从 modApp.bas 读出 APP_VERSION。" }
$Version = $m.Groups[1].Value
if ($Version -notmatch '^[0-9.]+$') { throw "版本号只允许数字和点：$Version" }
Write-Ok "版本号：$Version"

#-----------------------------------------------------------------------------
# 3. 构建
#-----------------------------------------------------------------------------
$artifactName = if ($Mode -eq "loader") { "ExcelToolboxLoader.xlam" } else { "ExcelToolbox.xlam" }
$artifact     = Join-Path $DistDir $artifactName

if ($SkipBuild) {
    Write-Step "跳过构建（-SkipBuild）"
    if (-not (Test-Path -LiteralPath $artifact)) { throw "dist\$artifactName 不存在，没法跳过构建。" }

    # 【产物比源码旧就是在打陈包】。这正是 -SkipBuild 最容易出的事故，
    # 而它没有任何其它征兆。
    #
    # 【只比对真正进了这个产物的源码】。拿整个 src\ 去比的话，
    # 打独立版时会被 src\loader 的改动触发——一条经常误报的检查
    # 等于没有检查，人会习惯性忽略它。
    # 【构建脚本本身也算输入】。只盯着 src\ 的话，改了 build.ps1
    # （比如换了注入 customUI 的方式）而没重新构建，产物照样"比源码新"，
    # 于是打出一个内容过期的包——同样没有任何征兆。
    if ($Mode -eq "loader") {
        $srcDirs   = @("src\loader")
        $srcFiles  = @("build\build-loader.ps1", "build\_ExcelHost.ps1")
    } else {
        $srcDirs   = @("src\code", "src\package", "src\help")
        $srcFiles  = @("build\build.ps1", "build\_ExcelHost.ps1")
    }

    $inputs = @()
    $inputs += $srcDirs  | ForEach-Object { Get-ChildItem (Join-Path $RepoRoot $_) -Recurse -File -ErrorAction SilentlyContinue }
    $inputs += $srcFiles | ForEach-Object { Get-Item (Join-Path $RepoRoot $_) -ErrorAction SilentlyContinue }
    $newestSrc = $inputs | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $art = Get-Item -LiteralPath $artifact
    if ($newestSrc -and $newestSrc.LastWriteTime -gt $art.LastWriteTime) {
        Add-Problem "dist\$artifactName 比源码旧（源码 $($newestSrc.Name) 改于 $($newestSrc.LastWriteTime)，产物生成于 $($art.LastWriteTime)）。去掉 -SkipBuild 重新构建。"
    } else {
        Write-Ok "现有产物不比源码旧：$artifactName"
    }
}
else {
    Write-Step "从源码构建 $artifactName"
    Write-Host "    构建要驱动 Excel，期间别去动它的窗口。" -ForegroundColor DarkGray

    if ($Mode -eq "loader") {
        & powershell -NoProfile -ExecutionPolicy Bypass `
            -File (Join-Path $PSScriptRoot "build-loader.ps1") -SharePath $SharePath
    } else {
        & powershell -NoProfile -ExecutionPolicy Bypass `
            -File (Join-Path $PSScriptRoot "build.ps1")
    }
    if ($LASTEXITCODE -ne 0) { throw "构建失败（退出码 $LASTEXITCODE），包没打。" }
    if (-not (Test-Path -LiteralPath $artifact)) { throw "构建说成功了，但找不到 dist\$artifactName。" }
    Write-Ok "构建完成：$artifactName"
}

#-----------------------------------------------------------------------------
# 4. 组装
#-----------------------------------------------------------------------------
Write-Step "组装安装包"

$pkgName  = "Excel工具箱_v$Version" + $(if ($Mode -eq "loader") { "_自动更新版" } else { "_独立版" })
$stageDir = Join-Path $OutDir $pkgName
if (Test-Path -LiteralPath $stageDir) { Remove-Item -LiteralPath $stageDir -Recurse -Force }
$null = New-Item -ItemType Directory -Path $stageDir -Force

# 【加载器在用户机器上也叫 ExcelToolboxLoader.xlam】，不要改名：
# 卸载逻辑按 ExcelToolbox*.xlam 找文件，改名会让卸载找不着。
Copy-Item -LiteralPath $artifact -Destination (Join-Path $stageDir $artifactName)
Copy-Item -LiteralPath (Join-Path $InstallDir "Excel工具箱.bat")      -Destination $stageDir
Copy-Item -LiteralPath (Join-Path $InstallDir "Install-Toolbox.ps1")  -Destination $stageDir
Copy-Item -LiteralPath (Join-Path $InstallDir "使用说明.txt")          -Destination $stageDir
Write-Ok "本体 4 个文件已就位"

if ($withAI) {
    $aiOut = Join-Path $stageDir "ai"
    $null = New-Item -ItemType Directory -Path $aiOut -Force

    $tplPath = Join-Path $InstallDir "ai\manifest.template.xml"
    Copy-Item -LiteralPath $tplPath -Destination $aiOut

    # 【README.txt 不进包】。它写的是"给 IT 打包时用的"，
    # 里面讲的是怎么填 gateway.txt、怎么换 GUID——业务用户看了只会困惑。
    #
    # gateway.txt 【不带 BOM】：安装脚本是 Get-Content -Raw 之后 Trim，
    # 带 BOM 的话首字符可能混进地址，拼出来的 URL 是坏的，
    # 而 Office 遇到非法 manifest 是静默不加载。
    [IO.File]::WriteAllText((Join-Path $aiOut "gateway.txt"), $Gateway, [Text.UTF8Encoding]::new($false))



    # 示例 GUID 没换会让两个组织的加载项互相覆盖，且极难排查
    $tplText = [IO.File]::ReadAllText($tplPath, [Text.UTF8Encoding]::new($false))
    if ($tplText -match '7b2e4c91-6a38-4d5f-9e10-3c8a5f2d6b47') {
        Write-Warn "manifest 模板里还是示例 GUID。正式分发前请换成你自己的（见 install\ai\README.txt）。"
    }
    Write-Ok "ai 组件已就位（manifest 模板 + gateway.txt）"
}

#-----------------------------------------------------------------------------
# 5. 自检 —— 包组装完了，但"文件都在"不等于"能用"
#
# 这里查的每一条，都对应一个【在用户机器上才显形、且报错完全不指向真因】
# 的故障。宁可在打包机上红一次。
#-----------------------------------------------------------------------------
Write-Step "自检"

# .ps1 必须 UTF-8 with BOM：中文 Windows 上没 BOM 会按 GBK 读，直接语法错误，
# 用户看到的是一屏英文报错
$ps1 = Join-Path $stageDir "Install-Toolbox.ps1"
$head = [IO.File]::ReadAllBytes($ps1)[0..2]
if ($head[0] -eq 0xEF -and $head[1] -eq 0xBB -and $head[2] -eq 0xBF) {
    Write-Ok "Install-Toolbox.ps1 是 UTF-8 with BOM"
} else {
    Add-Problem "Install-Toolbox.ps1 缺 BOM。中文机器会按 GBK 解码，脚本直接语法错误。"
}

# .bat 必须纯 ASCII：里面的中文取决于控制台代码页，换台机器就是乱码
$batBytes = [IO.File]::ReadAllBytes((Join-Path $stageDir "Excel工具箱.bat"))
if (@($batBytes | Where-Object { $_ -gt 127 }).Count -eq 0) {
    Write-Ok "Excel工具箱.bat 是纯 ASCII"
} else {
    Add-Problem "Excel工具箱.bat 含非 ASCII 字符，换台机器可能显示成乱码。"
}

# .xlam 得是个 OOXML 包，且里面真的有 VBA 工程——
# 构建环节出岔子时最典型的产物就是"文件在、宏没了"
$xlamPath = Join-Path $stageDir $artifactName
$xb = [IO.File]::ReadAllBytes($xlamPath)
if ($xb.Length -lt 1024 -or $xb[0] -ne 0x50 -or $xb[1] -ne 0x4B) {
    Add-Problem "$artifactName 不是有效的 xlsx/xlam 包（大小 $($xb.Length) 字节）。"
} else {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($xlamPath)
    try {
        $names = @($zip.Entries | ForEach-Object { $_.FullName })
        if ($names -contains "xl/vbaProject.bin") {
            Write-Ok "$artifactName 内含 VBA 工程"
        } else {
            Add-Problem "$artifactName 里没有 xl\vbaProject.bin —— 宏丢了，装上也没有任何功能。"
        }
        if ($names -contains "customUI/customUI14.xml") {
            Write-Ok "$artifactName 内含功能区定义"
        } elseif ($Mode -eq "standalone") {
            Add-Problem "$artifactName 里没有 customUI14.xml —— 装上不会出现「工具箱」选项卡。"
        } else {
            # 加载器本来就没有 customUI，功能区随载荷下发
            Write-Ok "加载器不含功能区定义（符合预期，功能区随载荷下发）"
        }
    } finally { $zip.Dispose() }
}

if ($withAI) {
    # 占位符替换后必须还是合法 XML，否则 Office 静默不加载。
    # 这里先用打包时的网关值试算一遍——真正的替换在用户机器上做，
    # 但网关是现在定下的，能在这儿就把它验了。
    $probe = $tplText.Replace("{{USER}}", "packtest").Replace("{{GATEWAY}}", $Gateway)
    try {
        [void]([xml]$probe)
        Write-Ok "用这个网关地址替换后，manifest 仍是合法 XML"
    } catch {
        Add-Problem "用这个网关地址替换后 manifest 不是合法 XML：$($_.Exception.Message)"
    }

    $gw = [IO.File]::ReadAllText((Join-Path $stageDir "ai\gateway.txt"), [Text.UTF8Encoding]::new($false))
    if ($gw -eq $Gateway) { Write-Ok "gateway.txt 内容无误且无 BOM" }
    else { Add-Problem "gateway.txt 内容和 -Gateway 不一致（可能混进了 BOM）：「$gw」" }
}

#-----------------------------------------------------------------------------
# 6. 打包清单 —— 让别人能核对，而不是只能相信
#-----------------------------------------------------------------------------
$listLines = @()
$listLines += "Excel 工具箱 安装包清单"
$listLines += "============================================"
$listLines += ""
$listLines += "版本      ：$Version"
$listLines += "形态      ：$(if ($Mode -eq 'loader') { 'loader（瘦加载器，载荷从共享目录自动更新）' } else { 'standalone（独立版，功能全在包里）' })"
if ($Mode -eq "loader") { $listLines += "共享目录  ：$SharePath" }
$listLines += "AI 组件   ：$(if ($withAI) { "有（网关 $Gateway；不含证书，网关证书须已被客户端信任）" } else { '无' })"
$listLines += "打包时间  ：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$listLines += "打包机器  ：$env:COMPUTERNAME"
$listLines += "打包账号  ：$env:USERNAME"
$listLines += ""
$listLines += "文件清单（SHA256）"
$listLines += "--------------------------------------------"
foreach ($f in (Get-ChildItem -LiteralPath $stageDir -Recurse -File | Sort-Object FullName)) {
    $rel = $f.FullName.Substring($stageDir.Length + 1)
    $listLines += ("{0}  {1}" -f (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash, $rel)
}
$listLines += ""
$listLines += "核对方法：在解包后的目录里跑"
$listLines += '    Get-ChildItem -Recurse -File | Get-FileHash -Algorithm SHA256'
$listLines += "对不上说明包在传输途中被改过或损坏了。"

$listPath = Join-Path $stageDir "打包清单.txt"
[IO.File]::WriteAllLines($listPath, $listLines, [Text.UTF8Encoding]::new($true))

#-----------------------------------------------------------------------------
# 7. 压缩
#
# 【自检没过就不要生成 zip】。生成了再报错，那个 zip 还是躺在 dist\ 里，
# 下一个人（或几天后的自己）很容易直接拿去发——错误提示早就滚出屏幕了。
# 不留下能发的东西，比留下一句警告可靠。
#-----------------------------------------------------------------------------
if ($problems.Count -gt 0) {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "  自检没过，已中止，未生成 zip" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    foreach ($p in $problems) { Write-Host "  - $p" -ForegroundColor Red }
    Write-Host ""
    Write-Host "半成品目录留在这里供你排查：$stageDir" -ForegroundColor DarkGray
    exit 1
}

Write-Step "压缩"
$zipPath = Join-Path $OutDir "$pkgName.zip"
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $zipPath -Force

# 【压完要把 zip 再打开核一遍】。前面所有自检查的都是 stage 目录里的文件，
# 而发出去的是这个 zip。压缩这一步本身也会出事（磁盘满、被杀毒软件改写、
# 路径太长导致漏文件），结果就是"脚本说成功了，用户解开发现少东西"。
# 核不过就把这个 zip 删掉——不留下能发的坏包。
if (-not (Test-Path -LiteralPath $zipPath)) {
    throw "压缩命令没报错，但 $zipPath 不存在。"
}

$expected = @(Get-ChildItem -LiteralPath $stageDir -Recurse -File |
              ForEach-Object { $_.FullName.Substring($stageDir.Length + 1).Replace('\', '/') })
$zipBad = @()
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zf = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        # 【两边都要归一化分隔符】。zip 规范用 /，但 PowerShell 5.1 的
        # Compress-Archive 写进去的是 \。直接比字符串会把每一个子目录
        # 里的文件都判成"少了"——一条永远为真的告警等于没有告警。
        $inZip = @($zf.Entries | ForEach-Object { $_.FullName.Replace('\', '/') })
        # 【双向比对】。只查"该有的都在"是不够的：多出来的条目同样是问题
        # （上一次的残留被并进来、通配符匹配到了不该进包的东西），
        # 而多出来的文件一样会跟着发到用户手里。
        foreach ($e in $expected) {
            if ($inZip -notcontains $e) { $zipBad += "少了 $e" }
        }
        foreach ($e in $inZip) {
            if ($expected -notcontains $e) { $zipBad += "多了不该有的 $e" }
        }
        # 条目大小为 0 而源文件不是 0，说明内容没写进去
        foreach ($entry in $zf.Entries) {
            $src = Join-Path $stageDir ($entry.FullName.Replace('/', '\'))
            if ((Test-Path -LiteralPath $src) -and $entry.Length -eq 0 -and (Get-Item -LiteralPath $src).Length -gt 0) {
                $zipBad += "$($entry.FullName) 在包里是空的"
            }
        }
    } finally { $zf.Dispose() }
}
catch { $zipBad += "zip 打不开：$($_.Exception.Message)" }

if ($zipBad.Count -gt 0) {
    Write-Host ""
    Write-Host "生成的 zip 核对不过，不要发：" -ForegroundColor Red
    foreach ($b in $zipBad) { Write-Host "  - $b" -ForegroundColor Red }

    # 【删不掉一定要喊出来】。静默删除失败的话，一个坏包会原封不动
    # 躺在 dist\ 里，而报错早就滚出屏幕了——下一个人看到的就是
    # "有个 zip，看起来能发"。
    Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $zipPath) {
        Write-Host ""
        Write-Host "！！这个坏包没能删掉，请手工删除，【绝对不要分发】：" -ForegroundColor Red
        Write-Host "    $zipPath" -ForegroundColor Red
    } else {
        Write-Host "已删除该 zip。" -ForegroundColor DarkGray
    }
    exit 1
}
Write-Ok "zip 核对通过（$($expected.Count) 个文件都在，内容非空）"
Write-Ok "$zipPath"

#-----------------------------------------------------------------------------
# 收尾
#-----------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  打包完成：$pkgName.zip" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "发之前【务必】在一台干净机器上实装一次：" -ForegroundColor Yellow
Write-Host "  解压 → 双击 Excel工具箱.bat → 回车 → 打开 Excel 看「工具箱」选项卡" -ForegroundColor Yellow
Write-Host "自检只能证明文件对，证明不了它在别人的机器上装得上。" -ForegroundColor DarkGray
if ($Mode -eq "loader") {
    Write-Host ""
    Write-Host "loader 模式还要记得发载荷，否则客户端装完没有任何功能：" -ForegroundColor Yellow
    Write-Host "  powershell -File build\publish.ps1 -SharePath `"$SharePath`"" -ForegroundColor Yellow
}
exit 0
