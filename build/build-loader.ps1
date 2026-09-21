<#
.SYNOPSIS
    src\loader\code  ->  dist\ExcelToolboxLoader.xlam

.DESCRIPTION
    构建瘦加载器。它只负责"检查更新 + 打开载荷"，不含任何业务功能。

    加载器【装一次就不再动】，所以它自己没有自动更新机制——
    给更新器做自动更新会陷入鸡生蛋问题，而且一旦更新器自己更坏了，
    全公司都得手工重装。宁可让它保持简单到不需要改。

    加载器没有 customUI：功能区来自载荷。这样功能区改动也能随载荷下发，
    不需要碰客户端。

.PARAMETER SharePath
    写进加载器的默认发布目录。不传就用源码里的 DEFAULT_SHARE。

    【这是唯一的配置时机】。加载器运行时没有任何办法改这个地址——
    原本打算用注册表做运行时配置，但实测 Excel 宏里
    CreateObject("WScript.Shell") 会被安全策略静默拦下，
    表现为"更新莫名其妙不生效"且查不出原因。详见 modLoader.bas 的注释。

    换发布目录 = 重新构建加载器 + 每台机器重装一次加载器。
    载荷（功能本身）的更新不受影响，那个才是天天在变的东西。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build-loader.ps1 -SharePath "\\fs01\tools\ExcelToolbox"
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolboxLoader.xlam",
    [string]$SharePath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Excel 而不是 WPS（WPS 会劫持 Excel 的 COM 注册并自称 Microsoft Excel）
. (Join-Path $PSScriptRoot "_ExcelHost.ps1")
$CodeDir  = Join-Path $RepoRoot "src\loader\code"
$DistDir  = Join-Path $RepoRoot "dist"
$OutPath  = Join-Path $DistDir $OutputName

$xlWBATWorksheet = -4167
$xlOpenXMLAddIn  = 55

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor DarkGray }

# 和 build.ps1 同一套转换：VBE 的 Import 只认 CRLF + 系统 ANSI，
# 两者中任何一个不对都不会报错，只会静默产出编译不过的工程。
function ConvertTo-VbaSource([string]$srcPath, [string]$destDir) {
    $text = [System.IO.File]::ReadAllText($srcPath, (New-Object System.Text.UTF8Encoding($false)))
    $text = $text -replace "`r`n", "`n"
    $text = $text -replace "`n", "`r`n"
    $dest = Join-Path $destDir (Split-Path $srcPath -Leaf)
    [System.IO.File]::WriteAllText($dest, $text, [System.Text.Encoding]::Default)
    return $dest
}

if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir | Out-Null }

if (Test-Path $OutPath) {
    try { $fs = [System.IO.File]::Open($OutPath, 'Open', 'ReadWrite', 'None'); $fs.Close() }
    catch { throw "$OutputName 正被占用。请先关闭所有 Excel 实例。" }
}

Write-Step "启动 Excel"
$xl = New-RealExcel
$wb = $null
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false

    $key = "HKCU:\Software\Microsoft\Office\$($xl.Version)\Excel\Security"
    $vbom = Get-ItemProperty -Path $key -Name "AccessVBOM" -ErrorAction SilentlyContinue
    if ($null -eq $vbom -or $vbom.AccessVBOM -ne 1) {
        throw "未开启「信任对 VBA 工程对象模型的访问」，无法导入 VBA 源码。见 README。"
    }

    Write-Step "新建工作簿"
    $wb = $xl.Workbooks.Add($xlWBATWorksheet)

    Write-Step "导入加载器源码"
    $stage = Join-Path ([System.IO.Path]::GetTempPath()) ("LoaderBuild_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $stage | Out-Null
    try {
        foreach ($f in (Get-ChildItem -Path $CodeDir -Filter *.bas | Sort-Object Name)) {
            $staged = ConvertTo-VbaSource $f.FullName $stage

            # 部署时把发布目录烧进源码，省得每台机器都去配注册表
            if ($SharePath -and $f.Name -eq "modLoader.bas") {
                $text = [System.IO.File]::ReadAllText($staged, [System.Text.Encoding]::Default)
                $escaped = $SharePath -replace '"', '""'
                $text = $text -replace 'Private Const DEFAULT_SHARE As String = "[^"]*"', `
                                       ('Private Const DEFAULT_SHARE As String = "' + $escaped + '"')
                [System.IO.File]::WriteAllText($staged, $text, [System.Text.Encoding]::Default)
                Write-Ok "发布目录写入为 $SharePath"
            }

            $comp = $wb.VBProject.VBComponents.Import($staged)
            if ([int]$comp.Type -ne 1) { throw "$($f.Name) 导入后类型为 $($comp.Type)，预期 1。" }
            Write-Ok $f.Name
        }
    }
    finally { Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue }

    $docModule = Join-Path $CodeDir "ThisWorkbook.doccls"
    if (Test-Path $docModule) {
        Write-Step "注入 ThisWorkbook 文档模块"
        $code = [System.IO.File]::ReadAllText($docModule, (New-Object System.Text.UTF8Encoding($false)))
        $code = ($code -replace "`r`n", "`n") -replace "`n", "`r`n"
        $cm = $wb.VBProject.VBComponents.Item("ThisWorkbook").CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
        Write-Ok "ThisWorkbook.doccls"
    }

    try { $wb.VBProject.Name = "ExcelToolboxLoader" } catch {}

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

Write-Host ""
Write-Host "加载器构建完成：$OutPath" -ForegroundColor Green
Write-Host "发布载荷：powershell -File build\publish.ps1 -SharePath <发布目录>" -ForegroundColor DarkGray
