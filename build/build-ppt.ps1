<#
.SYNOPSIS
    src\shared\code + src\ppt\code + src\ppt\package  ->  dist\PptToolbox.ppam

.DESCRIPTION
    和 build.ps1（Excel）是同一个思路，但 PowerPoint 的 COM 自动化在好几个
    地方和 Excel 不一样，已经用真实 COM 调用逐条验证过，写法必须照这些来
    （详细排查过程见 docs\规划-Word与PPT.md 的"PowerPoint COM 自动化的
    实测差异"一节）：

      · Application.Visible 不能设 False，会直接抛异常——用
        WindowState = 2（ppWindowMinimized）代替。
      · 没有 Application.EnableEvents 这个属性，这里直接不设。
      · Application.DisplayAlerts 是枚举 PpAlertLevel，不是布尔：
        ppAlertsNone = 1。
      · 新建的 Presentation.VBProject 默认零组件，没有 Excel ThisWorkbook
        那种自带的文档模块，不需要"文档模块不能 Import，只能塞
        CodeModule"那段特殊处理。
      · 【最关键】：customUI 的注入不能像 Excel 那样用
        [System.IO.Compression.ZipFile]::Open(path, "Update") 原地改——
        同样的改法 PowerPoint 会拒绝加载这个 .ppam（真机验证过：自动化
        下表现为 AddIns.Add(...).Loaded=True 无限期卡死，手动通过
        「文件→选项→加载项」操作则是明确报错"由于某种原因，PowerPoint
        无法加载加载项"）。必须完全展开到临时目录、加文件、改
        _rels/.rels，再从目录重新打包成新 zip。

    前置条件：PowerPoint 的「信任对 VBA 工程对象模型的访问」必须开启
    （这是逐宿主独立的开关，Excel 开了不代表 PowerPoint 也开了）。
    位置：PowerPoint 选项 → 信任中心 → 信任中心设置 → 宏设置 →
    勾选"信任对 VBA 工程对象模型的访问"。脚本会检测该设置，
    未开启时直接中止（不自动改注册表）。

    调用前必须确保桌面上没有其它 PowerPoint 实例在运行——PowerPoint
    的 COM 自动化实测在"已有实例运行"时容易互相阻塞（New-Object 会
    尝试和已有实例交互而不是干净地开一个新的），而且本脚本用进程
    快照比对识别自己创建的那个实例（PowerPoint 没有 Excel 那种
    Application.Hwnd 可以精确反查 PID，见 _PptHost.ps1 里的注释），
    台面上有别的实例会让这个识别不准。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build-ppt.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "PptToolbox.ppam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 PowerPoint 而不是 WPS
. (Join-Path $PSScriptRoot "_PptHost.ps1")

$SharedCodeDir = Join-Path $RepoRoot "src\shared\code"
$PptCodeDir    = Join-Path $RepoRoot "src\ppt\code"
$CodeDirs      = @($SharedCodeDir, $PptCodeDir)
$PackageDir    = Join-Path $RepoRoot "src\ppt\package"
$DistDir       = Join-Path $RepoRoot "dist"
$OutPath       = Join-Path $DistDir $OutputName

# PowerPoint 常量
$ppSaveAsOpenXMLAddin = 30
$ppAlertsNone = 1

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor DarkGray }

#------------------------------------------------------------------------------
# 源码 -> VBE 能正确导入的临时副本。和 build.ps1 的 ConvertTo-VbaSource
# 是同一份逻辑（CRLF + 系统 ANSI 代码页），原样复用。
#------------------------------------------------------------------------------
function ConvertTo-VbaSource([string]$srcPath, [string]$destDir) {
    $text = [System.IO.File]::ReadAllText($srcPath, (New-Object System.Text.UTF8Encoding($false)))
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`n", "`r`n"
    $dest = Join-Path $destDir (Split-Path $srcPath -Leaf)
    [System.IO.File]::WriteAllText($dest, $text, [System.Text.Encoding]::Default)
    return $dest
}

#------------------------------------------------------------------------------
# 检查「信任对 VBA 工程对象模型的访问」（PowerPoint 独立的注册表分支）
#------------------------------------------------------------------------------
function Test-PptVbomTrust([string]$pptVersion) {
    $key = "HKCU:\Software\Microsoft\Office\$pptVersion\PowerPoint\Security"
    if (-not (Test-Path $key)) { return $false }
    $v = Get-ItemProperty -Path $key -Name "AccessVBOM" -ErrorAction SilentlyContinue
    return ($null -ne $v -and $v.AccessVBOM -eq 1)
}

#------------------------------------------------------------------------------
# 把 customUI 注入 .ppam（zip）。
#
# 【不能用 ZipFile.Open(path, "Update") 原地改】——PowerPoint 会拒绝加载
# 这样产出的包。必须完全展开、加文件、改关系、重新打包整个 zip。
# 这是本脚本和 build.ps1（Excel）的 Add-CustomUI 最大的不同点，
# 别把两边搞混或者互相抄。
#------------------------------------------------------------------------------
function Add-PptCustomUI([string]$ppamPath, [string]$customUiPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

    $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

    $tempRoot = [System.IO.Path]::GetTempPath()
    $stageName = "PptToolboxBuild_" + [guid]::NewGuid().ToString("N")
    $stage = Join-Path $tempRoot $stageName
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ppamPath, $stage)

        $customUiDir = Join-Path $stage "customUI"
        New-Item -ItemType Directory -Path $customUiDir -Force | Out-Null
        Copy-Item $customUiPath (Join-Path $customUiDir "customUI14.xml") -Force

        $relsPath = Join-Path $stage "_rels\.rels"
        $relsXml = [System.IO.File]::ReadAllText($relsPath)
        if ($relsXml -notmatch [regex]::Escape($relType)) {
            $newRel = "<Relationship Id=`"rIdPptToolboxUI`" Type=`"$relType`" Target=`"customUI/customUI14.xml`"/>"
            $relsXml = $relsXml -replace "</Relationships>", "$newRel</Relationships>"
            [System.IO.File]::WriteAllText($relsPath, $relsXml, (New-Object System.Text.UTF8Encoding($false)))
        }

        Remove-Item $ppamPath -Force
        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $stage, $ppamPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
    }
    finally {
        Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#------------------------------------------------------------------------------
# 主流程
#------------------------------------------------------------------------------
foreach ($d in $CodeDirs) {
    if (-not (Test-Path $d)) { throw "找不到源码目录：$d" }
}
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }

$customUi = Join-Path $PackageDir "customUI\customUI14.xml"
if (-not (Test-Path $customUi)) { throw "找不到 Ribbon 定义：$customUi" }

# 台面上不能有别的 PowerPoint——见文件头部注释，识别自己创建的实例
# 靠进程快照比对，别的实例在跑会让这个识别不准。
if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
    throw "检测到已有 PowerPoint 实例在运行。请先关闭所有 PowerPoint 窗口再重新构建。"
}

if (Test-Path $OutPath) {
    try {
        $fs = [System.IO.File]::Open($OutPath, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        throw "$OutputName 正被占用。请先在 PowerPoint 中取消勾选该加载项并关闭所有 PowerPoint 实例，再重新构建。"
    }
}

Write-Step "启动 PowerPoint"
$ppt = New-RealPpt
$pres = $null
try {
    # 【不能设 Visible = False】：PowerPoint 会直接抛异常
    # "Hiding the application window is not allowed"。退而求其次最小化。
    $ppt.WindowState = 2   # ppWindowMinimized
    $ppt.DisplayAlerts = $ppAlertsNone
    Write-Ok "PowerPoint $($ppt.Version)"

    if (-not (Test-PptVbomTrust $ppt.Version)) {
        throw @"
未开启「信任对 VBA 工程对象模型的访问」，无法导入 VBA 源码。

请在 PowerPoint 中开启：
  文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
  → 勾选「信任对 VBA 工程对象模型的访问」
然后重新运行本脚本。
"@
    }

    Write-Step "新建演示文稿"
    $pres = $ppt.Presentations.Add($true)

    Write-Step "导入 VBA 源码"
    # 【modAction.bas 暂不给 PPT 用】：它调用 modPrompt/modTelemetry/modHelp/
    # modActionRegistry.RegisterAll/modActionRegistry.Dispatch，这几个
    # PPT 工程里都还没有（PPT 的 Ribbon_OnAction 目前直接弹 MsgBox，
    # 完全没经过 RunAction）。如果不排除，会编译出"Sub 或 Function 未定义"
    # 的加载项——Office 的 SaveAs 不会在这一步报错，只会静默产出损坏的
    # 文件，等真正调用时才炸，这是 Codex 复审挑出的真实问题。等 PPT 也
    # 有了自己的 modHost.bas + modActionRegistry.bas（"PPT 样板命令"那步）
    # 再把这一条排除去掉。
    $files = Get-ChildItem -Path $CodeDirs -Recurse -Include *.bas, *.cls, *.frm |
             Where-Object { $_.Name -ne "modAction.bas" } |
             Sort-Object FullName
    if ($files.Count -eq 0) { throw "src\shared\code / src\ppt\code 下没有可导入的 .bas/.cls/.frm。" }

    $stage = Join-Path ([System.IO.Path]::GetTempPath()) ("PptToolboxBuild_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        foreach ($f in $files) {
            $frx = [System.IO.Path]::ChangeExtension($f.FullName, ".frx")
            if (Test-Path $frx) { Copy-Item $frx (Join-Path $stage (Split-Path $frx -Leaf)) -Force }

            $staged = ConvertTo-VbaSource $f.FullName $stage
            $comp = $pres.VBProject.VBComponents.Import($staged)

            $expected = switch ($f.Extension.ToLower()) { ".bas" { 1 } ".cls" { 2 } ".frm" { 3 } }
            if ([int]$comp.Type -ne $expected) {
                throw "$($f.Name) 导入后类型为 $($comp.Type)，预期 $expected。源文件头部未被 VBE 识别。"
            }

            $base = $CodeDirs | Where-Object { $f.FullName.StartsWith($_ + "\") } | Select-Object -First 1
            Write-Ok $f.FullName.Substring($base.Length + 1)
        }
    }
    finally {
        Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
    }

    try { $pres.VBProject.Name = "PptToolbox" } catch { Write-Ok "VBProject 改名失败（不影响功能）" }

    Write-Step "另存为 .ppam"
    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
    $pres.SaveAs($OutPath, $ppSaveAsOpenXMLAddin)
    $pres.Close()
    $pres = $null
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    Close-PptInstance $ppt
}

Write-Step "注入 Ribbon 定义"
Add-PptCustomUI -ppamPath $OutPath -customUiPath $customUi
Write-Ok "customUI/customUI14.xml"

Write-Host ""
Write-Host "构建完成：$OutPath" -ForegroundColor Green
Write-Host "手动安装：把文件放到 %APPDATA%\Microsoft\Addins\，" -ForegroundColor DarkGray
Write-Host "         然后 PowerPoint 里「文件→选项→加载项→转到」勾选启用。" -ForegroundColor DarkGray
