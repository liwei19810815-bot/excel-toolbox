<#
.SYNOPSIS
    跑完全部验证，并把结果写成仓库里可复核的 docs\验收记录.md。

.DESCRIPTION
    为什么需要这个脚本：

    "我跑过了，全绿"这种口头结论没法复核。评审者拿到仓库时看不到任何证据，
    只能选择信或不信——而这恰恰是本项目一直在反对的东西
    （同样的理由催生了 docs\图标验收记录.md）。

    这个脚本把每一套的【运行时实际通过数】、退出码、宿主版本、时间戳
    落成一份文档。换台机器重跑，数字对不上就说明那台机器有问题，
    而不是靠谁的记忆。

    【断言数的口径】必须说清楚，否则数字永远对不上：
    记的是【运行时实际通过的断言数】，不是脚本源码里 Assert-* 的调用次数。
    两者不同，因为有些断言在循环里（跑几次算几条），
    有些在条件分支里（可能一次都不跑）。数源码行数会得出偏大的数字。

.EXAMPLE
    timeout 3000 powershell -ExecutionPolicy Bypass -File tests\run-all.ps1
#>
[CmdletBinding()]
param(
    [switch]$SkipLoader     # 加载器套件最慢，调试时可跳过
)

$ErrorActionPreference = "Continue"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

function Write-Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }

# 取宿主信息，写进记录——同一份代码在不同 Office 上结果可能不同
$hostDesc = "(取不到)"
$xl = $null
try {
    $xl = New-RealExcel
    $hostDesc = "$($xl.Name) $($xl.Version) build $($xl.Build)"
} catch {
    Write-Host "无法启动 Excel：$($_.Exception.Message)" -ForegroundColor Red
    exit 1
} finally { Close-ExcelInstance $xl }

Write-Host "宿主：$hostDesc" -ForegroundColor DarkGray
Write-Host ""

#-----------------------------------------------------------------------------
# 逐套件跑，从输出里抓【运行时实际通过数】
#-----------------------------------------------------------------------------
$suites = @(
    @{ Name = "整工程编译";       Script = "..\build\verify.ps1";        Kind = "exit" }
    @{ Name = "主套件";           Script = "run-tests.ps1";              Kind = "count" }
    @{ Name = "文件批处理套件";   Script = "run-tests-files.ps1";        Kind = "count" }
    @{ Name = "遥测套件";         Script = "run-tests-telemetry.ps1";    Kind = "count" }
    @{ Name = "加载器套件";       Script = "run-tests-loader.ps1";       Kind = "count"; Slow = $true }
    @{ Name = "功能区接线与加载"; Script = "check-ribbon.ps1";           Kind = "exit" }
    @{ Name = "功能区图标";       Script = "check-imagemso.ps1";         Kind = "exit" }
)

$results = @()
$totalAssertions = 0
$anyFailed = $false

foreach ($s in $suites) {
    if ($SkipLoader -and $s.ContainsKey('Slow')) {
        Write-Host "跳过 $($s.Name)" -ForegroundColor DarkYellow
        continue
    }

    Write-Step $s.Name
    $path = Join-Path $PSScriptRoot $s.Script
    $out = & powershell -NoProfile -ExecutionPolicy Bypass -File $path 2>&1 | Out-String
    $code = $LASTEXITCODE

    $passed = ""
    $failed = ""
    if ($s.Kind -eq "count") {
        # 各套件末尾统一是「通过 N / 失败 M」
        if ($out -match '通过\s+(\d+)\s*/\s*失败\s+(\d+)') {
            $passed = $Matches[1]
            $failed = $Matches[2]
            $totalAssertions += [int]$passed
        }
    }

    $ok = ($code -eq 0)
    if (-not $ok) { $anyFailed = $true }

    $detail = if ($s.Kind -eq "count") { "通过 $passed / 失败 $failed" } else { "退出码 $code" }
    Write-Host ("    {0}  {1}" -f $(if ($ok) { "OK  " } else { "失败" }), $detail) `
        -ForegroundColor $(if ($ok) { "Green" } else { "Red" })

    $results += [pscustomobject]@{
        Name     = $s.Name
        Ok       = $ok
        Passed   = $passed
        Failed   = $failed
        ExitCode = $code
    }
}

#-----------------------------------------------------------------------------
# 收尾检查：这一整轮跑完不该留下任何东西
#-----------------------------------------------------------------------------
Write-Step "残留检查"
$leftoverProc = @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count
$lockDir = Join-Path $env:TEMP "ExcelToolbox.pids"
$leftoverLock = @(Get-ChildItem $lockDir -ErrorAction SilentlyContinue).Count
Write-Host "    Excel 进程 $leftoverProc / 锁文件 $leftoverLock" `
    -ForegroundColor $(if ($leftoverProc -eq 0 -and $leftoverLock -eq 0) { "Green" } else { "Red" })
if ($leftoverProc -ne 0 -or $leftoverLock -ne 0) { $anyFailed = $true }

#-----------------------------------------------------------------------------
# 落成文档
#-----------------------------------------------------------------------------
$lines = @()
$lines += "# 验收记录"
$lines += ""
$lines += "本文件由 ``tests\run-all.ps1`` 自动生成，请勿手工编辑。"
$lines += ""
$lines += "**断言数的口径**：记的是【运行时实际通过的断言数】，"
$lines += "不是脚本源码里 ``Assert-*`` 的调用次数。两者不同——"
$lines += "有些断言在循环里（跑几次算几条），有些在条件分支里（可能一次都不跑）。"
$lines += "数源码行数会得出偏大的数字。"
$lines += ""
$lines += "- 宿主：``$hostDesc``"
$lines += "- 生成时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm')"
$lines += "- 总计通过断言：**$totalAssertions**"
$lines += "- 残留 Excel 进程：$leftoverProc　残留锁文件：$leftoverLock"
$lines += "- 结论：$(if ($anyFailed) { '**存在失败项**' } else { '**全部通过**' })"
$lines += ""
$lines += "| 套件 | 结果 | 明细 |"
$lines += "|---|---|---|"
foreach ($r in $results) {
    $mark = if ($r.Ok) { "✅" } else { "❌" }
    $detail = if ($r.Passed -ne "") { "通过 $($r.Passed) / 失败 $($r.Failed)" } else { "退出码 $($r.ExitCode)" }
    $lines += "| $($r.Name) | $mark | $detail |"
}
$lines += ""
$lines += "> 换一台机器重跑本脚本，数字对不上就说明那台机器的环境有差异，"
$lines += "> 而不是靠谁的记忆。图标的逐个哈希另见 ``图标验收记录.md``。"

$reportPath = Join-Path $RepoRoot "docs\验收记录.md"
Set-Content -LiteralPath $reportPath -Value $lines -Encoding UTF8

Write-Host ""
Write-Host "验收记录：$reportPath" -ForegroundColor Cyan
Write-Host "总计通过断言：$totalAssertions" -ForegroundColor $(if ($anyFailed) { "Red" } else { "Green" })

if ($anyFailed) { exit 1 } else { exit 0 }
