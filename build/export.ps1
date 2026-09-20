<#
.SYNOPSIS
    dist\ExcelToolbox.xlam  ->  src\code（回写 VBA 源码）

.DESCRIPTION
    在 VBE 里直接改代码调试很方便，但改动只存在于二进制 xlam 里。
    本脚本把 xlam 中的全部 VBA 组件导出回 src\code，让 git 能看到 diff。

    目录归属：按组件名在 src\code 下递归查找同名文件，找到就原地覆盖，
    保持原有的 Core\ Text\ Data\ 分层；找不到的（新建的模块）落到 src\code\_new，
    由人工挪到正确的子目录。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\export.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Excel 而不是 WPS（WPS 会劫持 Excel 的 COM 注册并自称 Microsoft Excel）
. (Join-Path $PSScriptRoot "_ExcelHost.ps1")
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"
$CodeDir  = Join-Path $RepoRoot "src\code"
$NewDir   = Join-Path $CodeDir "_new"

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok($msg)   { Write-Host "    $msg" -ForegroundColor DarkGray }

#------------------------------------------------------------------------------
# VBE 的 Export 写出的是【系统 ANSI 代码页 + CRLF】。仓库里统一存 UTF-8 + LF，
# 所以导出后要转一道，否则每次导出都会把整个文件的换行和中文注释搅成一片假 diff。
# 与 build.ps1 的 ConvertTo-VbaSource 互为逆操作。
#------------------------------------------------------------------------------
function ConvertTo-RepoSource([string]$path) {
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::Default)
    $text = $text -replace "`r`n", "`n"
    [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。" }

# vbext_ComponentType
$typeExt = @{
    1 = ".bas"   # StdModule
    2 = ".cls"   # ClassModule
    3 = ".frm"   # MSForm
}

Write-Step "启动 Excel"
$xl = New-RealExcel
$wb = $null
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false

    $wb = $xl.Workbooks.Open($Xlam, $false, $true)   # UpdateLinks=False, ReadOnly=True

    Write-Step "导出组件"
    foreach ($comp in $wb.VBProject.VBComponents) {
        $name = $comp.Name

        # 文档模块（ThisWorkbook / Sheet1）：只导出代码正文
        if ($comp.Type -eq 100) {
            if ($name -ne "ThisWorkbook") { continue }
            $cm = $comp.CodeModule
            if ($cm.CountOfLines -eq 0) { continue }
            $dest = Join-Path $CodeDir "Core\ThisWorkbook.doccls"
            $text = $cm.Lines(1, $cm.CountOfLines) -replace "`r`n", "`n"
            [System.IO.File]::WriteAllText($dest, $text, (New-Object System.Text.UTF8Encoding($false)))
            Write-Ok "Core\ThisWorkbook.doccls"
            continue
        }

        if (-not $typeExt.ContainsKey([int]$comp.Type)) { continue }
        $ext = $typeExt[[int]$comp.Type]

        $match = Get-ChildItem -Path $CodeDir -Recurse -File -Filter "$name$ext" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($match) {
            $dest = $match.FullName
        } else {
            if (-not (Test-Path $NewDir)) { New-Item -ItemType Directory -Path $NewDir | Out-Null }
            $dest = Join-Path $NewDir "$name$ext"
        }

        $comp.Export($dest)
        ConvertTo-RepoSource $dest
        Write-Ok $dest.Substring($CodeDir.Length + 1)
    }

    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) {
        try { $xl.EnableEvents = $true } catch {}
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "导出完成。请用 git diff 检查改动。" -ForegroundColor Green
if (Test-Path $NewDir) {
    Write-Host "注意：src\code\_new 下有新组件，请手动移到对应子目录。" -ForegroundColor Yellow
}
