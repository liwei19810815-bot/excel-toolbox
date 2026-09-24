<#
.SYNOPSIS
    创建一个【确认是真 Microsoft PowerPoint】的 COM 实例。

.DESCRIPTION
    和 _ExcelHost.ps1 是同一个问题、同一套解法（进程归属登记与回收、
    WPS 劫持校验），针对 PowerPoint 的 COM 自动化做的独立实现——
    没有直接改造 _ExcelHost.ps1 成通用的 _OfficeHost.ps1，是因为那个
    脚本里的孤儿进程回收逻辑经过大量轮次验证，PPT 这边的调用方式
    和 Excel 有几处真实差异（下面会遇到），贸然合并有回归 Excel 那套
    已经很稳的逻辑的风险。等 Word 那边也有了实际需求、能验证"三宿主
    共用一份"这件事之后，再考虑合并成 _OfficeHost.ps1。

    PowerPoint COM 自动化和 Excel 不一样的地方（都已用真实 COM 调用
    实测过，写进这里以免以后重复踩）：
      · Application.Visible 不能设 False，会直接抛异常——用
        WindowState = 2（ppWindowMinimized）代替。
      · 没有 Application.EnableEvents 这个属性，对应逻辑整段跳过。
      · Application.DisplayAlerts 是枚举 PpAlertLevel，不是布尔：
        ppAlertsNone = 1、ppAlertsAll = 2。

.EXAMPLE
    . "$PSScriptRoot\_PptHost.ps1"
    $ppt = New-RealPpt
#>

# 判断一个 PowerPoint COM 实例到底是不是 WPS 演示。
# 和 Test-IsWpsHost（Excel 版）同一个判据：路径匹配用短名前缀、
# 版本号低于 14（本工具箱最低支持 PowerPoint 2010）判定为 WPS。
function Test-IsWpsPptHost {
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
# 进程的归属登记与回收——和 _ExcelHost.ps1 同一套机制，只是进程名换成
# POWERPNT。详细注释见 _ExcelHost.ps1，这里不重复。
#==============================================================================

if (-not ([System.Management.Automation.PSTypeName]'ExcelToolbox.Win32').Type) {
    Add-Type -Namespace ExcelToolbox -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@
}

$script:PptLockDir = Join-Path $env:TEMP "ExcelToolbox.ppt.pids"
$script:PptLockFile = Join-Path $script:PptLockDir ("run_{0}.txt" -f $PID)
$script:PptOwned = @()

function Clear-StalePpt {
    if (-not (Test-Path $script:PptLockDir)) { return }

    foreach ($f in (Get-ChildItem -LiteralPath $script:PptLockDir -Filter "run_*.txt" -ErrorAction SilentlyContinue)) {
        $ownerPid = 0
        if ($f.BaseName -match '^run_(\d+)$') {
            $ownerPid = [int]$Matches[1]
        } else {
            Write-Host "    警告：无法识别的锁文件，已跳过：$($f.Name)" -ForegroundColor DarkYellow
            continue
        }
        if ($ownerPid -eq $PID) { continue }
        if (Test-PptProcessAlive $ownerPid) { continue }

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
            switch (Get-OwnedPptProcessState $owned) {
                'Gone'    { }
                'NotOurs' { }
                'Ours'    {
                    Write-Host ("    回收上一轮遗留的 PowerPoint 进程 {0}" -f $pid2) -ForegroundColor DarkYellow
                    try { Stop-Process -Id $pid2 -Force } catch {}
                    $gone = $false
                    for ($k = 0; $k -lt 6; $k++) {
                        if (-not (Test-PptProcessAlive $pid2)) { $gone = $true; break }
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
# 【PowerPoint 的 Application 对象没有 Hwnd/HWND 属性】——这是和 Excel
# 真实存在的差异，不是疏漏。Excel 那边（_ExcelHost.ps1）靠
# Application.Hwnd 反查 PID，从根上避免"创建前后快照比对"的竞态窗口；
# PowerPoint 没有这条路可走（微软自己的文档和社区都确认了这一点，
# DocumentWindow 层面也没有等价属性）。
#
# 退而求其次：记录调用 New-Object 之前已存在的 POWERPNT 进程 PID 集合，
# 创建之后再查一次，取差集里"最新启动"的那个。这确实有竞态窗口——
# 如果用户在这几百毫秒内自己手动开了一个 PowerPoint，可能被误认成
# 本次创建的实例。调用方必须【自己保证调用前台面上没有别的 PowerPoint】，
# 这条约束写进了 New-RealPpt 的调用方需知（build-ppt.ps1 里会先检查）。
#------------------------------------------------------------------------------
function Get-PptIdentity {
    param($App, [int]$Retries = 10)

    $before = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            $after = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
            $newOnes = @($after | Where-Object { $_.Id -notin $before })
            if ($newOnes.Count -eq 1) {
                return @{ Id = $newOnes[0].Id; Ticks = $newOnes[0].StartTime.Ticks }
            }
            if ($newOnes.Count -gt 1) {
                # 出现不止一个新增进程，选不出来是哪个——按最新启动时间挑一个，
                # 总比完全不登记强，但这种情况本身就说明调用前提被违反了
                $latest = $newOnes | Sort-Object StartTime -Descending | Select-Object -First 1
                return @{ Id = $latest.Id; Ticks = $latest.StartTime.Ticks }
            }
        } catch {}
        Start-Sleep -Milliseconds 300
    }
    return $null
}

function Test-PptProcessAlive {
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

function Get-OwnedPptProcessState {
    param($Owned)

    $p = $null
    try { $p = Get-Process -Id $Owned.Id -ErrorAction Stop }
    catch [Microsoft.PowerShell.Commands.ProcessCommandException] { return 'Gone' }
    catch { return 'Unknown' }

    try {
        if ($p.ProcessName -ne 'POWERPNT' -and $p.ProcessName -notin @('wpp')) { return 'NotOurs' }
        if ($p.StartTime.Ticks -ne $Owned.Ticks) { return 'NotOurs' }
        return 'Ours'
    }
    catch { return 'Unknown' }
}

function Stop-UnregisteredPpt {
    param($App, $Own)

    if ($App) {
        try { $App.DisplayAlerts = 1 } catch {}   # ppAlertsNone
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (-not $Own) { return }
    if ((Get-OwnedPptProcessState $Own) -eq 'Ours') {
        try { Stop-Process -Id $Own.Id -Force } catch {}
    }
}

function Register-PptInstance {
    param($App)

    $own = Get-PptIdentity $App
    if (-not $own) {
        Stop-UnregisteredPpt $App $null
        throw "取不到 PowerPoint 实例的进程身份（HWND 或启动时间不可用）。已尝试关闭并中止。"
    }
    $procId = $own.Id

    $script:PptOwned += [pscustomobject]@{ Id = $own.Id; Ticks = $own.Ticks }

    try {
        if (-not (Test-Path $script:PptLockDir)) {
            $null = New-Item -ItemType Directory -Path $script:PptLockDir -Force
        }

        $lines = @($script:PptOwned | ForEach-Object { "$($_.Id)|$($_.Ticks)" })
        $tmp = "$script:PptLockFile.tmp"
        Set-Content -LiteralPath $tmp -Encoding ASCII -Value $lines
        Move-Item -LiteralPath $tmp -Destination $script:PptLockFile -Force

        $back = @(Get-Content -LiteralPath $script:PptLockFile -ErrorAction SilentlyContinue)
        if (($back -join "`n") -ne ($lines -join "`n")) {
            throw "PID 登记文件回读不一致：$script:PptLockFile"
        }
    }
    catch {
        $script:PptOwned = @($script:PptOwned | Where-Object { $_.Id -ne $procId })
        Stop-UnregisteredPpt $App $own
        throw "登记 PowerPoint 进程失败，已关闭该实例：$($_.Exception.Message)"
    }
}

function Close-PptInstance {
    param($App)

    if ($App) {
        try { $App.DisplayAlerts = 1 } catch {}
        try { foreach ($pr in @($App.Presentations)) { try { $pr.Close() } catch {} } } catch {}
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500

    foreach ($o in $script:PptOwned) {
        $state = Get-OwnedPptProcessState $o

        if ($state -eq 'Gone')    { continue }
        if ($state -eq 'NotOurs') { continue }

        if ($state -eq 'Unknown') {
            Write-Host "    警告：无法确认 PowerPoint 进程 $($o.Id) 的状态，锁文件保留以便下轮回收。" -ForegroundColor DarkYellow
            return
        }

        try { Stop-Process -Id $o.Id -Force } catch {}

        $exited = $false
        for ($i = 0; $i -lt 10; $i++) {
            if (-not (Test-PptProcessAlive $o.Id)) { $exited = $true; break }
            Start-Sleep -Milliseconds 500
        }

        if (-not $exited) {
            Write-Host "    警告：PowerPoint 进程 $($o.Id) 未能退出，锁文件保留以便下轮回收。" -ForegroundColor DarkYellow
            return
        }
    }

    $script:PptOwned = @()
    if (Test-Path $script:PptLockFile) {
        Remove-Item -LiteralPath $script:PptLockFile -Force -ErrorAction SilentlyContinue
        if (Test-Path $script:PptLockFile) {
            Write-Host "    警告：锁文件未能删除：$script:PptLockFile" -ForegroundColor DarkYellow
        }
    }
}

function New-RealPpt {
    [CmdletBinding()]
    param()

    Clear-StalePpt

    $ppt = New-Object -ComObject PowerPoint.Application
    Register-PptInstance $ppt

    $path = ""; $name = ""; $ver = ""
    try { $path = $ppt.Path } catch {}
    try { $name = $ppt.Name } catch {}
    try { $ver  = $ppt.Version } catch {}

    if (Test-IsWpsPptHost $ppt) {
        Close-PptInstance $ppt
        throw @"
COM 自动化拿到的不是 Microsoft PowerPoint，而是 WPS：
    Name    = $name
    Version = $ver
    Path    = $path

恢复方法：WPS → 设置/配置工具 → 兼容设置 → 取消 Office 文件关联，
或点"恢复 Office 默认设置"，然后重新运行本脚本。
"@
    }

    return $ppt
}
