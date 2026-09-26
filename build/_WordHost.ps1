<#
.SYNOPSIS
    创建一个【确认是真 Microsoft Word】的 COM 实例。

.DESCRIPTION
    和 _ExcelHost.ps1 / _PptHost.ps1 是同一个问题、同一套解法（进程归属
    登记与回收、WPS 劫持校验），针对 Word 的 COM 自动化做的独立实现。
    没有直接合并成 _OfficeHost.ps1——三份脚本各自独立验证过，贸然合并
    有回归 Excel/PPT 那两套已经很稳的逻辑的风险。

    Word COM 自动化实测下来的结论（用真实 COM 调用验证过，写进这里以免
    以后重复验证）：
      · Application.Visible 可以直接设 False，不像 PowerPoint 那样会
        抛异常——Word 这一点和 Excel 一致，不需要 WindowState 变通。
      · 没有 Application.EnableEvents 这个属性（和 PowerPoint 一样，
        这是 Excel 专有的），对应逻辑整段跳过；但有 ScreenUpdating，
        可以正常设置。
      · Application.DisplayAlerts 是数值型：wdAlertsNone = 0，
        不是 PowerPoint 那种从 1 开始的 PpAlertLevel 枚举，也不是
        Excel 的布尔——三个宿主三种写法，不能互相抄数值。
      · 【没有 Application.Hwnd】：真机测试过，即使已经打开一个可见
        文档，$w.Hwnd 读出来也是空值，不能像 Excel 那样反查 PID。
        这一点和 PowerPoint 是同一个坑，Word 不能走 Excel 的路，
        必须走本文件这套"创建前后进程快照比对"的方案。

.EXAMPLE
    . "$PSScriptRoot\_WordHost.ps1"
    $w = New-RealWord
#>

# 判断一个 Word COM 实例到底是不是 WPS 文字。
# 和 Test-IsWpsHost（Excel 版）/ Test-IsWpsPptHost（PPT 版）同一个判据：
# 路径匹配用短名前缀、版本号低于 14（本工具箱最低支持 Word 2010）判定为 WPS。
function Test-IsWpsWordHost {
    param($App)

    $path = ""; $name = ""; $ver = ""
    try { $path = $App.Path } catch {}
    try { $name = $App.Name } catch {}
    try { $ver  = $App.Version } catch {}

    if ($path -match 'WPSOFF|KINGSO|WPS Office') { return $true }
    if ($name -match 'WPS') { return $true }

    $num = 0.0
    if ($ver) { [void][double]::TryParse(($ver -replace '[^0-9.].*$',''), [ref]$num) }
    if ($num -gt 0 -and $num -lt 14) { return $true }

    return $false
}

#==============================================================================
# 进程的归属登记与回收——和 _ExcelHost.ps1 / _PptHost.ps1 同一套机制，
# 只是进程名换成 WINWORD。详细注释见 _ExcelHost.ps1，这里不重复。
#==============================================================================

$script:WordLockDir = Join-Path $env:TEMP "ExcelToolbox.word.pids"
$script:WordLockFile = Join-Path $script:WordLockDir ("run_{0}.txt" -f $PID)
$script:WordOwned = @()

function Clear-StaleWord {
    if (-not (Test-Path $script:WordLockDir)) { return }

    foreach ($f in (Get-ChildItem -LiteralPath $script:WordLockDir -Filter "run_*.txt" -ErrorAction SilentlyContinue)) {
        $ownerPid = 0
        if ($f.BaseName -match '^run_(\d+)$') {
            $ownerPid = [int]$Matches[1]
        } else {
            Write-Host "    警告：无法识别的锁文件，已跳过：$($f.Name)" -ForegroundColor DarkYellow
            continue
        }
        if ($ownerPid -eq $PID) { continue }
        if (Test-WordProcessAlive $ownerPid) { continue }

        $allResolved = $true
        $lines = $null
        try { $lines = @(Get-Content -LiteralPath $f.FullName -ErrorAction Stop) }
        catch {
            Write-Host "    警告：锁文件读取失败，保留待下轮处理：$($f.Name)" -ForegroundColor DarkYellow
            continue
        }

        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $parts = $line -split '\|'
            $pid2 = 0
            $ticks = 0L
            if ($parts.Count -ne 2 -or
                -not [int]::TryParse($parts[0], [ref]$pid2) -or
                -not [long]::TryParse($parts[1], [ref]$ticks)) {
                Write-Host "    警告：锁文件里有无法解析的登记，保留待人工处理：$($f.Name)" -ForegroundColor DarkYellow
                $allResolved = $false
                continue
            }

            $owned = @{ Id = $pid2; Ticks = $ticks }
            switch (Get-OwnedWordProcessState $owned) {
                'Gone'    { }
                'NotOurs' { }
                'Ours'    {
                    Write-Host ("    回收上一轮遗留的 Word 进程 {0}" -f $pid2) -ForegroundColor DarkYellow
                    try { Stop-Process -Id $pid2 -Force } catch {}
                    $gone = $false
                    for ($k = 0; $k -lt 6; $k++) {
                        if (-not (Test-WordProcessAlive $pid2)) { $gone = $true; break }
                        Start-Sleep -Milliseconds 500
                    }
                    if (-not $gone) { $allResolved = $false }
                }
                default   {
                    Write-Host ("    警告：无法确认进程 {0} 的状态，保留锁文件待下轮处理。" -f $pid2) -ForegroundColor DarkYellow
                    $allResolved = $false
                }
            }
        }

        if ($allResolved) {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

#------------------------------------------------------------------------------
# 【Word 的 Application 对象也没有可用的 Hwnd】——真机测试过，即使已经
# 打开可见文档也读不到值，不能像 Excel 那样反查 PID。做法和 PPT 一样：
# 记录调用 New-Object 之前已存在的 WINWORD 进程 PID 集合，创建之后再查
# 一次，取差集里"新增"的那个。$Before 必须由调用方在 New-Object 之前
# 拍好传进来——这里不自己现拍，PPT 那边已经因为这个时序问题被 Codex
# 挑出过一次真实 bug（见 build/_PptHost.ps1 的同名注释），这里从一开始
# 就按正确顺序写，不重蹈覆辙。
#------------------------------------------------------------------------------
function Get-WordIdentity {
    param($App, $Before, [int]$Retries = 10)

    $before = @($Before)

    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
            $newOnes = @($after | Where-Object { $_.Id -notin $before })
            if ($newOnes.Count -eq 1) {
                return @{ Id = $newOnes[0].Id; Ticks = $newOnes[0].StartTime.Ticks }
            }
            if ($newOnes.Count -gt 1) {
                $latest = $newOnes | Sort-Object StartTime -Descending | Select-Object -First 1
                return @{ Id = $latest.Id; Ticks = $latest.StartTime.Ticks }
            }
        } catch {}
        Start-Sleep -Milliseconds 300
    }
    return $null
}

function Test-WordProcessAlive {
    param([int]$ProcId)
    try {
        return [bool](Get-Process -Id $ProcId -ErrorAction Stop)
    }
    catch [Microsoft.PowerShell.Commands.ProcessCommandException] {
        return $false
    }
    catch {
        return $true
    }
}

function Get-OwnedWordProcessState {
    param($Owned)

    $p = $null
    try { $p = Get-Process -Id $Owned.Id -ErrorAction Stop }
    catch [Microsoft.PowerShell.Commands.ProcessCommandException] { return 'Gone' }
    catch { return 'Unknown' }

    try {
        if ($p.ProcessName -ne 'WINWORD' -and $p.ProcessName -notin @('wps')) { return 'NotOurs' }
        if ($p.StartTime.Ticks -ne $Owned.Ticks) { return 'NotOurs' }
        return 'Ours'
    }
    catch { return 'Unknown' }
}

function Stop-UnregisteredWord {
    param($App, $Own)

    if ($App) {
        try { $App.DisplayAlerts = 0 } catch {}   # wdAlertsNone
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (-not $Own) { return }
    if ((Get-OwnedWordProcessState $Own) -eq 'Ours') {
        try { Stop-Process -Id $Own.Id -Force } catch {}
    }
}

function Register-WordInstance {
    param($App, $Before)

    $own = Get-WordIdentity $App $Before
    if (-not $own) {
        Stop-UnregisteredWord $App $null
        throw "取不到 Word 实例的进程身份（启动时间不可用）。已尝试关闭并中止。"
    }
    $procId = $own.Id

    $script:WordOwned += [pscustomobject]@{ Id = $own.Id; Ticks = $own.Ticks }

    try {
        if (-not (Test-Path $script:WordLockDir)) {
            $null = New-Item -ItemType Directory -Path $script:WordLockDir -Force
        }

        $lines = @($script:WordOwned | ForEach-Object { "$($_.Id)|$($_.Ticks)" })
        $tmp = "$script:WordLockFile.tmp"
        Set-Content -LiteralPath $tmp -Encoding ASCII -Value $lines
        Move-Item -LiteralPath $tmp -Destination $script:WordLockFile -Force

        $back = @(Get-Content -LiteralPath $script:WordLockFile -ErrorAction SilentlyContinue)
        if (($back -join "`n") -ne ($lines -join "`n")) {
            throw "PID 登记文件回读不一致：$script:WordLockFile"
        }
    }
    catch {
        $script:WordOwned = @($script:WordOwned | Where-Object { $_.Id -ne $procId })
        Stop-UnregisteredWord $App $own
        throw "登记 Word 进程失败，已关闭该实例：$($_.Exception.Message)"
    }
}

function Close-WordInstance {
    param($App)

    if ($App) {
        try { $App.DisplayAlerts = 0 } catch {}
        try { foreach ($d in @($App.Documents)) { try { $d.Close($false) } catch {} } } catch {}
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500

    foreach ($o in $script:WordOwned) {
        $state = Get-OwnedWordProcessState $o

        if ($state -eq 'Gone')    { continue }
        if ($state -eq 'NotOurs') { continue }

        if ($state -eq 'Unknown') {
            Write-Host "    警告：无法确认 Word 进程 $($o.Id) 的状态，锁文件保留以便下轮回收。" -ForegroundColor DarkYellow
            return
        }

        try { Stop-Process -Id $o.Id -Force } catch {}

        $exited = $false
        for ($i = 0; $i -lt 10; $i++) {
            if (-not (Test-WordProcessAlive $o.Id)) { $exited = $true; break }
            Start-Sleep -Milliseconds 500
        }

        if (-not $exited) {
            Write-Host "    警告：Word 进程 $($o.Id) 未能退出，锁文件保留以便下轮回收。" -ForegroundColor DarkYellow
            return
        }
    }

    $script:WordOwned = @()
    if (Test-Path $script:WordLockFile) {
        Remove-Item -LiteralPath $script:WordLockFile -Force -ErrorAction SilentlyContinue
        if (Test-Path $script:WordLockFile) {
            Write-Host "    警告：锁文件未能删除：$script:WordLockFile" -ForegroundColor DarkYellow
        }
    }
}

function New-RealWord {
    [CmdletBinding()]
    param()

    Clear-StaleWord

    $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $w = New-Object -ComObject Word.Application
    Register-WordInstance $w $before

    $path = ""; $name = ""; $ver = ""
    try { $path = $w.Path } catch {}
    try { $name = $w.Name } catch {}
    try { $ver  = $w.Version } catch {}

    if (Test-IsWpsWordHost $w) {
        Close-WordInstance $w
        throw @"
COM 自动化拿到的不是 Microsoft Word，而是 WPS：
    Name    = $name
    Version = $ver
    Path    = $path

恢复方法：WPS → 设置/配置工具 → 兼容设置 → 取消 Office 文件关联，
或点"恢复 Office 默认设置"，然后重新运行本脚本。
"@
    }

    return $w
}
