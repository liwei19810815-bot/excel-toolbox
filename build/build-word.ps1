<#
.SYNOPSIS
    src\shared\code + src\word\code + src\word\package  ->  dist\WordToolbox.dotm

.DESCRIPTION
    和 build.ps1（Excel）/ build-ppt.ps1（PPT）是同一个思路。Word 的 COM
    自动化实测下来比 PowerPoint 更接近 Excel（细节见 _WordHost.ps1），
    但有自己的差异点，已用真实 COM 调用验证过：

      · Application.Visible 可以直接设 False（这点像 Excel，不像
        PowerPoint 需要 WindowState 变通）。
      · 没有 Application.EnableEvents，直接不设；ScreenUpdating 正常。
      · Application.DisplayAlerts 是数值：wdAlertsNone = 0（不是
        PowerPoint 的 1，也不是 Excel 的布尔 False）。
      · 新建 Document 的 VBProject 默认组件数没有实测确认过是否为零，
        构建脚本按"可能已有组件"处理，只管 Import，不假设起始状态。
      · customUI 注入沿用 PPT 那套"完全展开重新打包"的安全做法（不用
        Excel 的 ZipFile.Open(path,"Update") 原地改）——PPT 那边已经
        真机验证过原地改会导致加载项拒绝加载，本脚本没有反过来验证
        "Word 是否也有这个问题"，直接用已证明安全的做法，不重复冒险。

    前置条件：Word 的「信任对 VBA 工程对象模型的访问」必须开启（这是
    逐宿主独立的开关，Excel/PowerPoint 开了不代表 Word 也开了）。
    位置：Word 选项 → 信任中心 → 信任中心设置 → 宏设置 →
    勾选"信任对 VBA 工程对象模型的访问"。脚本会检测该设置，
    未开启时直接中止（不自动改注册表）。

    调用前必须确保桌面上没有其它 Word 实例在运行——本脚本用进程快照
    比对识别自己创建的那个实例（Word 没有可用的 Application.Hwnd，
    见 _WordHost.ps1 里的注释），台面上有别的实例会让这个识别不准。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build-word.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "WordToolbox.dotm"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Word 而不是 WPS
. (Join-Path $PSScriptRoot "_WordHost.ps1")

$SharedCodeDir = Join-Path $RepoRoot "src\shared\code"
$WordCodeDir   = Join-Path $RepoRoot "src\word\code"
$CodeDirs      = @($SharedCodeDir, $WordCodeDir)
$PackageDir    = Join-Path $RepoRoot "src\word\package"
$DistDir       = Join-Path $RepoRoot "dist"
$OutPath       = Join-Path $DistDir $OutputName

# Word 常量
$wdFormatXMLTemplateMacroEnabled = 15
$wdAlertsNone = 0

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor DarkGray }

#------------------------------------------------------------------------------
# 源码 -> VBE 能正确导入的临时副本。和 build.ps1 / build-ppt.ps1 的
# ConvertTo-VbaSource 是同一份逻辑（CRLF + 系统 ANSI 代码页），原样复用。
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
# 检查「信任对 VBA 工程对象模型的访问」（Word 独立的注册表分支）
#------------------------------------------------------------------------------
function Test-WordVbomTrust([string]$wordVersion) {
    $key = "HKCU:\Software\Microsoft\Office\$wordVersion\Word\Security"
    if (-not (Test-Path $key)) { return $false }
    $v = Get-ItemProperty -Path $key -Name "AccessVBOM" -ErrorAction SilentlyContinue
    return ($null -ne $v -and $v.AccessVBOM -eq 1)
}

#------------------------------------------------------------------------------
# 把 customUI 注入 .dotm（zip）。做法和 build-ppt.ps1 的 Add-PptCustomUI
# 完全一致（完全展开、加文件、改关系、重新打包），只是换了变量名——
# 见文件头部注释，这里不重复解释为什么不用 ZipFile.Open(path,"Update")。
#------------------------------------------------------------------------------
function Add-WordCustomUI([string]$dotmPath, [string]$customUiPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

    $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

    $tempRoot = [System.IO.Path]::GetTempPath()
    $stageName = "WordToolboxBuild_" + [guid]::NewGuid().ToString("N")
    $stage = Join-Path $tempRoot $stageName
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($dotmPath, $stage)

        $customUiDir = Join-Path $stage "customUI"
        New-Item -ItemType Directory -Path $customUiDir -Force | Out-Null
        Copy-Item $customUiPath (Join-Path $customUiDir "customUI14.xml") -Force

        $relsPath = Join-Path $stage "_rels\.rels"
        $relsXml = [System.IO.File]::ReadAllText($relsPath)
        if ($relsXml -notmatch [regex]::Escape($relType)) {
            $newRel = "<Relationship Id=`"rIdWordToolboxUI`" Type=`"$relType`" Target=`"customUI/customUI14.xml`"/>"
            $relsXml = $relsXml -replace "</Relationships>", "$newRel</Relationships>"
            [System.IO.File]::WriteAllText($relsPath, $relsXml, (New-Object System.Text.UTF8Encoding($false)))
        }

        Remove-Item $dotmPath -Force
        [System.IO.Compression.ZipFile]::CreateFromDirectory(
            $stage, $dotmPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
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

# 台面上不能有别的 Word——见文件头部注释，识别自己创建的实例靠进程
# 快照比对，别的实例在跑会让这个识别不准。
if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
    throw "检测到已有 Word 实例在运行。请先关闭所有 Word 窗口再重新构建。"
}

if (Test-Path $OutPath) {
    try {
        $fs = [System.IO.File]::Open($OutPath, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        throw "$OutputName 正被占用。请先在 Word 中取消勾选该加载项并关闭所有 Word 实例，再重新构建。"
    }
}

Write-Step "启动 Word"
$word = New-RealWord
$doc = $null
try {
    # Word 可以直接设 Visible = False，不需要 PPT 那种 WindowState 变通。
    $word.Visible = $false
    $word.DisplayAlerts = $wdAlertsNone
    Write-Ok "Word $($word.Version)"

    if (-not (Test-WordVbomTrust $word.Version)) {
        throw @"
未开启「信任对 VBA 工程对象模型的访问」，无法导入 VBA 源码。

请在 Word 中开启：
  文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
  → 勾选「信任对 VBA 工程对象模型的访问」
然后重新运行本脚本。
"@
    }

    Write-Step "新建文档"
    $doc = $word.Documents.Add()

    Write-Step "导入 VBA 源码"
    $files = Get-ChildItem -Path $CodeDirs -Recurse -Include *.bas, *.cls, *.frm |
             Sort-Object FullName
    if ($files.Count -eq 0) { throw "src\shared\code / src\word\code 下没有可导入的 .bas/.cls/.frm。" }

    $stage = Join-Path ([System.IO.Path]::GetTempPath()) ("WordToolboxBuild_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        foreach ($f in $files) {
            $frx = [System.IO.Path]::ChangeExtension($f.FullName, ".frx")
            if (Test-Path $frx) { Copy-Item $frx (Join-Path $stage (Split-Path $frx -Leaf)) -Force }

            $staged = ConvertTo-VbaSource $f.FullName $stage
            $comp = $doc.VBProject.VBComponents.Import($staged)

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

    try { $doc.VBProject.Name = "WordToolbox" } catch { Write-Ok "VBProject 改名失败（不影响功能）" }

    Write-Step "另存为 .dotm"
    if (Test-Path $OutPath) { Remove-Item $OutPath -Force }
    $doc.SaveAs($OutPath, $wdFormatXMLTemplateMacroEnabled)
    $doc.Close($false)
    $doc = $null
}
finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    Close-WordInstance $word
}

Write-Step "注入 Ribbon 定义"
Add-WordCustomUI -dotmPath $OutPath -customUiPath $customUi
Write-Ok "customUI/customUI14.xml"

Write-Host ""
Write-Host "构建完成：$OutPath" -ForegroundColor Green
Write-Host "手动安装：Word 里「文件→选项→加载项→模板→管理:模板→转到」，" -ForegroundColor DarkGray
Write-Host "         「模板和加载项」对话框 →「添加」选中该文件 → 勾选启用。" -ForegroundColor DarkGray
