<#
.SYNOPSIS
    编译 sidecar 伴生进程。

.DESCRIPTION
    【故意不用 dotnet SDK】。用的是 Windows 自带的 .NET Framework 编译器
    （C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe）。

    理由和整个项目一贯的取向一致——不给打包的人、也不给用户增加安装步骤：
      · 开发机不用装 .NET SDK
      · 用户机不用装任何运行时（.NET Framework 4.x 是 Windows 自带的）

    代价是【只能写 C# 5 的语法】：字符串内插 $""、?. 、nameof、
    表达式体成员统统不能用，用了会直接编译失败。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\build-sidecar.ps1
#>
[CmdletBinding()]
param(
    [string]$OutPath
)

$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SrcFile = Join-Path $RepoRoot "sidecar\ExcelToolboxSidecar.cs"
if (-not $OutPath) { $OutPath = Join-Path $RepoRoot "dist\ExcelToolboxSidecar.exe" }

function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }

#------------------------------------------------------------------------------
# 找编译器
#------------------------------------------------------------------------------
$Csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $Csc)) {
    $Csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
if (-not (Test-Path $Csc)) {
    throw "找不到系统自带的 C# 编译器 csc.exe。这台机器的 .NET Framework 可能不完整。"
}

if (-not (Test-Path $SrcFile)) { throw "找不到源码：$SrcFile" }

$OutDir = Split-Path -Parent $OutPath
if (-not (Test-Path $OutDir)) { $null = New-Item -ItemType Directory -Path $OutDir -Force }

Write-Step "编译 sidecar"
Write-Host "    编译器：$Csc" -ForegroundColor DarkGray
Write-Host "    源码：  $SrcFile" -ForegroundColor DarkGray

# System.Web.Extensions 提供 JavaScriptSerializer，用来读配置里的 JSON。
# 它是 .NET Framework 自带的，不引入任何第三方依赖。
$refs = @(
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Web.Extensions.dll"
)

$args = @(
    "/nologo",
    "/target:exe",
    "/platform:anycpu",
    "/optimize+",
    "/warnaserror+",          # 【警告一律当错误】：这个项目反复吃过"不报错但坏了"的亏
    "/utf8output",
    "/out:$OutPath"
) + $refs + @($SrcFile)

$out = & $Csc @args 2>&1
$rc = $LASTEXITCODE

if ($rc -ne 0) {
    Write-Host ""
    Write-Host "编译失败：" -ForegroundColor Red
    $out | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    Write-Host ""
    Write-Host "提示：系统自带的 csc 只支持到 C# 5。" -ForegroundColor Yellow
    Write-Host "      字符串内插 `$\"\"、?.、nameof、表达式体成员都不能用。" -ForegroundColor Yellow
    exit 1
}

if ($out) { $out | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkGray } }

if (-not (Test-Path $OutPath)) { throw "编译器报成功，但产物不存在：$OutPath" }

$size = [math]::Round((Get-Item $OutPath).Length / 1KB, 1)
Write-Host ""
Write-Host "  OK  $OutPath （$size KB）" -ForegroundColor Green
exit 0
