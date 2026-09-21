<#
.SYNOPSIS
    src\code + src\package  ->  dist\ExcelToolbox.xlam

.DESCRIPTION
    .xlam 是 zip 二进制，无法进 git 做 diff，所以仓库里只存源码。
    本脚本用 COM 驱动 Excel 新建一个空工作簿、导入全部 VBA 源码、另存为 .xlam，
    然后在 Excel 退出后把 customUI14.xml 注入到 zip 包里（VBA 对象模型无法写 customUI）。

    前置条件：Excel 的「信任对 VBA 工程对象模型的访问」必须开启。
    位置：Excel 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选"信任对 VBA 工程对象模型的访问"。
    脚本会检测该设置，未开启时直接中止（不自动改注册表）。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot   = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Excel 而不是 WPS（WPS 会劫持 Excel 的 COM 注册并自称 Microsoft Excel）
. (Join-Path $PSScriptRoot "_ExcelHost.ps1")
$CodeDir    = Join-Path $RepoRoot "src\code"
$PackageDir = Join-Path $RepoRoot "src\package"
$DistDir    = Join-Path $RepoRoot "dist"
$OutPath    = Join-Path $DistDir $OutputName

# Excel 常量
$xlWBATWorksheet = -4167
$xlOpenXMLAddIn  = 55

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor DarkGray }

#------------------------------------------------------------------------------
# 源码 -> VBE 能正确导入的临时副本
#
# VBE 的 Import 有两个不讲道理的硬性要求，踩中任何一个都不会报错，只会静默产出
# 一个编译不过的工程：
#
#   1. 换行必须是 CRLF。LF 换行时解析器认不出 .cls 的 "VERSION 1.0 CLASS" 头部
#      和 .bas 的 "Attribute VB_Name" 头部，会把整个文件当成代码塞进一个标准模块
#      ——类模块就这么变成了标准模块，头部行变成语法错误。
#   2. 编码必须是系统 ANSI 代码页（本机 936/GBK）。UTF-8 的中文注释会变成乱码。
#
# git 里统一存 UTF-8 + LF，只在构建时转换，不污染仓库。
#------------------------------------------------------------------------------
function ConvertTo-VbaSource([string]$srcPath, [string]$destDir) {
    $text = [System.IO.File]::ReadAllText($srcPath, (New-Object System.Text.UTF8Encoding($false)))

    # 先统一成 LF 再换成 CRLF，避免原文件本来就混着 CRLF 时变成 CRCRLF
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`n", "`r`n"

    $dest = Join-Path $destDir (Split-Path $srcPath -Leaf)
    [System.IO.File]::WriteAllText($dest, $text, [System.Text.Encoding]::Default)
    return $dest
}

#------------------------------------------------------------------------------
# 检查「信任对 VBA 工程对象模型的访问」
#------------------------------------------------------------------------------
function Test-VbomTrust([string]$excelVersion) {
    $key = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security"
    if (-not (Test-Path $key)) { return $false }
    $v = Get-ItemProperty -Path $key -Name "AccessVBOM" -ErrorAction SilentlyContinue
    return ($null -ne $v -and $v.AccessVBOM -eq 1)
}

#------------------------------------------------------------------------------
# 把 customUI 注入 xlam（zip）
#------------------------------------------------------------------------------
function Add-CustomUI([string]$xlamPath, [string]$customUiPath) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

    $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $entryName = "customUI/customUI14.xml"
    $relId = "rIdExcelToolboxUI"

    $zip = [System.IO.Compression.ZipFile]::Open($xlamPath, "Update")
    try {
        # 1) 写入/覆盖 customUI14.xml
        $existing = $zip.Entries | Where-Object { $_.FullName -eq $entryName }
        if ($existing) { $existing.Delete() }
        $entry = $zip.CreateEntry($entryName)
        $sw = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write((Get-Content -Path $customUiPath -Raw -Encoding UTF8))
        $sw.Flush(); $sw.Dispose()

        # 2) 在 _rels/.rels 里登记关系
        $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq "_rels/.rels" }
        if (-not $relsEntry) { throw "xlam 包内找不到 _rels/.rels，文件可能已损坏。" }

        $sr = New-Object System.IO.StreamReader($relsEntry.Open())
        $relsXml = $sr.ReadToEnd(); $sr.Dispose()

        if ($relsXml -notmatch [regex]::Escape($relType)) {
            $newRel = "<Relationship Id=`"$relId`" Type=`"$relType`" Target=`"customUI/customUI14.xml`"/>"
            $relsXml = $relsXml -replace "</Relationships>", "$newRel</Relationships>"

            $stream = $relsEntry.Open()
            $stream.SetLength(0)
            $sw2 = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
            $sw2.Write($relsXml)
            $sw2.Flush(); $sw2.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

#------------------------------------------------------------------------------
# 解析 src\help\help.md
#
# 格式见该文件头部。这里只认三样东西：
#   ## <actionId>     —— 一条帮助的开始
#   关键词: ...        —— 搜索用词
#   其余行             —— 正文（原样保留，含 ### 小标题）
#
# 解析不出条目时【由调用方抛错】而不是静默产出一个空帮助——
# 帮助失效是不会让任何功能测试变红的那类问题。
#------------------------------------------------------------------------------
function ConvertFrom-HelpMarkdown {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path -Encoding UTF8
    $result = @()
    $cur = $null

    foreach ($line in $lines) {
        if ($line -match '^##\s+([A-Za-z]+\.[A-Za-z]+)\s*$') {
            if ($cur) { $cur.Body = ($cur.BodyLines -join "`r`n").Trim(); $result += $cur }
            $cur = [pscustomobject]@{
                Id        = $Matches[1]
                Keywords  = ""
                Body      = ""
                BodyLines = @()
            }
            continue
        }
        if ($null -eq $cur) { continue }          # 文件头部的说明注释，跳过

        if ($line -match '^关键词[:：]\s*(.+)$') {
            $cur.Keywords = $Matches[1].Trim()
            continue
        }
        $cur.BodyLines += $line
    }
    if ($cur) { $cur.Body = ($cur.BodyLines -join "`r`n").Trim(); $result += $cur }

    return $result
}

#------------------------------------------------------------------------------
# 主流程
#------------------------------------------------------------------------------
if (-not (Test-Path $CodeDir)) { throw "找不到源码目录：$CodeDir" }
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }

$customUi = Join-Path $PackageDir "customUI\customUI14.xml"
if (-not (Test-Path $customUi)) { throw "找不到 Ribbon 定义：$customUi" }

# 输出文件被占用（加载宏正在 Excel 里加载）时提前报错，避免走完一圈才失败
if (Test-Path $OutPath) {
    try {
        $fs = [System.IO.File]::Open($OutPath, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    } catch {
        throw "$OutputName 正被占用。请先在 Excel 中取消勾选该加载宏并关闭所有 Excel 实例，再重新构建。"
    }
}

Write-Step "启动 Excel"
$xl = New-RealExcel
$wb = $null
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    Write-Ok "Excel $($xl.Version)"

    if (-not (Test-VbomTrust $xl.Version)) {
        throw @"
未开启「信任对 VBA 工程对象模型的访问」，无法导入 VBA 源码。

请在 Excel 中开启：
  文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
  → 勾选「信任对 VBA 工程对象模型的访问」
然后重新运行本脚本。
"@
    }

    Write-Step "新建工作簿"
    $wb = $xl.Workbooks.Add($xlWBATWorksheet)

    Write-Step "导入 VBA 源码"
    $files = Get-ChildItem -Path $CodeDir -Recurse -Include *.bas, *.cls, *.frm |
             Sort-Object FullName
    if ($files.Count -eq 0) { throw "src\code 下没有可导入的 .bas/.cls/.frm。" }

    $stage = Join-Path ([System.IO.Path]::GetTempPath()) ("ExcelToolboxBuild_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        foreach ($f in $files) {
            # .frm 依赖同名 .frx 二进制资源，必须一起放进暂存目录
            $frx = [System.IO.Path]::ChangeExtension($f.FullName, ".frx")
            if (Test-Path $frx) { Copy-Item $frx (Join-Path $stage (Split-Path $frx -Leaf)) -Force }

            $staged = ConvertTo-VbaSource $f.FullName $stage
            $comp = $wb.VBProject.VBComponents.Import($staged)

            # 导入类型不对 = 头部没被识别，多半又是换行或编码退化了，不能让它静默过去
            $expected = switch ($f.Extension.ToLower()) { ".bas" { 1 } ".cls" { 2 } ".frm" { 3 } }
            if ([int]$comp.Type -ne $expected) {
                throw "$($f.Name) 导入后类型为 $($comp.Type)，预期 $expected。源文件头部未被 VBE 识别。"
            }

            Write-Ok $f.FullName.Substring($CodeDir.Length + 1)
        }
    }
    finally {
        Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
    }

    # 文档模块不能 Import，只能往 CodeModule 里塞源码
    $docModule = Join-Path $CodeDir "Core\ThisWorkbook.doccls"
    if (Test-Path $docModule) {
        Write-Step "注入 ThisWorkbook 文档模块"
        # 走 COM 传字符串（BSTR/UTF-16），没有编码问题，但换行仍需 CRLF
        $code = [System.IO.File]::ReadAllText($docModule, (New-Object System.Text.UTF8Encoding($false)))
        $code = ($code -replace "`r`n", "`n") -replace "`n", "`r`n"
        $cm = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
        Write-Ok "Core\ThisWorkbook.doccls"
    }

    try { $wb.VBProject.Name = "ExcelToolbox" } catch { Write-Ok "VBProject 改名失败（不影响功能）" }

    #--------------------------------------------------------------------------
    # 帮助内容注入成隐藏工作表
    #
    # 【为什么用工作表而不是生成一个 .bas】：中文正文嵌进 VBA 字符串字面量
    # 要处理单行长度上限、续行数上限和引号转义，很容易在某条帮助里踩雷，
    # 而且踩了是编译错误。.xlam 本身就是工作簿，用单元格存文本没有这些限制，
    # 也保住了"整个工具箱就一个文件"这个分发前提。
    #--------------------------------------------------------------------------
    $helpFile = Join-Path $RepoRoot "src\help\help.md"
    if (Test-Path $helpFile) {
        Write-Step "注入帮助内容"
        $entries = ConvertFrom-HelpMarkdown $helpFile
        if ($entries.Count -eq 0) { throw "src\help\help.md 解析不出任何条目，格式可能坏了。" }

        $sh = $wb.Worksheets.Add()
        $sh.Name = "_Help"
        $sh.Cells(1, 1).Value2 = "actionId"
        $sh.Cells(1, 2).Value2 = "keywords"
        $sh.Cells(1, 3).Value2 = "body"

        $r = 2
        foreach ($e in $entries) {
            $sh.Cells($r, 1).Value2 = $e.Id
            $sh.Cells($r, 2).Value2 = $e.Keywords
            $sh.Cells($r, 3).Value2 = $e.Body
            $r++
        }
        # xlSheetVeryHidden = 2：用户从右键菜单取消隐藏也看不到它，
        # 免得有人误删了导致帮助功能整个失效
        $sh.Visible = 2
        Write-Ok "$($entries.Count) 条帮助"
    }

    Write-Step "另存为加载宏"
    $wb.IsAddin = $true
    $wb.SaveAs($OutPath, $xlOpenXMLAddIn)
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.EnableEvents = $true } catch {} }
    Close-ExcelInstance $xl
}

Write-Step "注入 Ribbon 定义"
Add-CustomUI -xlamPath $OutPath -customUiPath $customUi
Write-Ok "customUI/customUI14.xml"

Write-Host ""
Write-Host "构建完成：$OutPath" -ForegroundColor Green
Write-Host "安装：powershell -ExecutionPolicy Bypass -File build\install.ps1" -ForegroundColor DarkGray
