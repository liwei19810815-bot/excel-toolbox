<#
.SYNOPSIS
    创建一个【确认是真 Microsoft Excel】的 COM 实例。

.DESCRIPTION
    为什么需要这个：装了 WPS 之后，它会把 Excel 的 COM 注册整个接管——
    HKCU 和 HKCR 两层的 CLSID{00024500-...}\LocalServer32 都被改成了 et.exe，
    连 Excel.Application.16 这个带版本号的 ProgID 也一样。

    更麻烦的是 WPS 会【自称 Microsoft Excel】：Application.Name 返回的就是
    "Microsoft Excel"，只有 Application.Path 和 Version(12.0) 能露馅。

    于是构建和测试脚本会在毫不知情的情况下驱动 WPS 跑，
    产出的 .xlam 和测试结论全都是错的靶子。这种"跑错宿主还全绿"的情况
    比直接报错危险得多，所以这里必须显式校验并【拿错就报错】。

.EXAMPLE
    . "$PSScriptRoot\_ExcelHost.ps1"
    $xl = New-RealExcel
#>

# 判断一个 Excel COM 实例到底是不是 WPS。
#
# 路径匹配必须用短名也能命中的前缀：实测 WPS 的 Application.Path 返回的是
# 8.3 短路径 D:\PROGRA~3\WPSOFF~1\...，"WPSOFFICE" 根本匹配不上。
# 版本号是第二道判据：WPS 报 12.0，而本工具箱最低支持的 Excel 2010 是 14.0，
# 所以真 Excel 不可能报 12.0。
function Test-IsWpsHost {
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
# 哪些脚本【不】走这套生命周期，以及为什么：
#
#   build\publish.ps1   —— 它根本不创建 Excel 实例，全程只做文件复制与改名
#                          （这也是它不要求大家关掉 Excel 的原因）。没有进程要管。
#   tests\probe-wps.ps1 —— 它要的是 WPS，不能用 New-RealExcel（那个拿到 WPS 会报错）。
#                          但它【仍然接入登记与回收】：直接创建 COM 之后
#                          调用 Register-ExcelInstance，收尾走 Close-ExcelInstance，
#                          所以也不再按进程名强杀。
#
# 除这两个之外，所有创建 Excel 的脚本都必须走 New-RealExcel + Close-ExcelInstance。
#==============================================================================

#==============================================================================
# Excel 进程的归属登记与回收
#
# 要解决的是这个反复出现的问题：脚本被外部 timeout 杀掉时，PowerShell 的
# finally 【不会执行】，它创建的 Excel 就变成孤儿留在后台。本项目实际发生过，
# 一次累积了 5 个，还把后续构建卡住了（.xlam 被占用）。
#
# 几个必须做对的点：
#
# 1.【PID 要从 Hwnd 反查，不要用前后快照比对】。
#    "创建前记一次进程列表、创建后再记一次、差集就是我的"——这中间有竞态窗口：
#    用户在这期间自己打开的 Excel 会被算成脚本创建的，然后被杀掉，
#    连同未保存的内容一起没了。Application.Hwnd 直接对应到本实例的窗口，
#    用 GetWindowThreadProcessId 反查出来的 PID 是精确的，没有窗口期。
#
# 2.【要连进程启动时间一起记】。只记 PID 不够：进程退出后 PID 会被系统复用，
#    下次回收时可能杀掉一个恰好拿到同一个 PID 的无关进程。
#    PID + 启动时间才能唯一确定一个进程。
#
# 3.【锁文件每次运行一个，不能共用】。两个脚本同时跑会互相覆盖。
#    文件名带上本 PowerShell 进程的 PID。
#
# 4.【登记要在拿到实例后立刻落盘】，否则中途被打断就没人知道该回收谁。
#==============================================================================

if (-not ([System.Management.Automation.PSTypeName]'ExcelToolbox.Win32').Type) {
    Add-Type -Namespace ExcelToolbox -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@
}

$script:ExcelLockDir = Join-Path $env:TEMP "ExcelToolbox.pids"
$script:ExcelLockFile = Join-Path $script:ExcelLockDir ("run_{0}.txt" -f $PID)
$script:ExcelOwned = @()

# 回收【上一轮被强杀后遗留】的 Excel。
# 只回收登记过的、且创建它的那个 PowerShell 已经不在了的条目。
function Clear-StaleExcel {
    if (-not (Test-Path $script:ExcelLockDir)) { return }

    foreach ($f in (Get-ChildItem -LiteralPath $script:ExcelLockDir -Filter "run_*.txt" -ErrorAction SilentlyContinue)) {
        # 文件名里的是创建者 PowerShell 的 PID。它还活着就说明那一轮还在跑，别碰。
        $ownerPid = 0
        if ($f.BaseName -match '^run_(\d+)$') { $ownerPid = [int]$Matches[1] }
        if ($ownerPid -eq $PID) { continue }
        if ($ownerPid -gt 0 -and (Get-Process -Id $ownerPid -ErrorAction SilentlyContinue)) { continue }

        foreach ($line in (Get-Content -LiteralPath $f.FullName -ErrorAction SilentlyContinue)) {
            # 每行格式：<pid>|<启动时间 ticks>
            $parts = $line -split '\|'
            if ($parts.Count -ne 2) { continue }

            $p = Get-Process -Id ([int]$parts[0]) -ErrorAction SilentlyContinue
            if (-not $p -or $p.ProcessName -ne 'EXCEL') { continue }

            # PID 可能被复用，必须连启动时间一起对上才敢动手
            $ticks = 0L
            if (-not [long]::TryParse($parts[1], [ref]$ticks)) { continue }
            if ($p.StartTime.Ticks -ne $ticks) { continue }

            Write-Host ("    回收上一轮遗留的 Excel 进程 {0}" -f $p.Id) -ForegroundColor DarkYellow
            try { Stop-Process -Id $p.Id -Force } catch {}
        }

        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
    }
}

# 从 COM 实例精确取出它自己的 PID（不做快照比对，没有竞态）。
#
# 【要重试】：Hwnd 在实例刚创建的一瞬间可能还是 0，窗口尚未建好。
# 第一版拿到 0 就放弃登记，那个实例此后就成了无人认领的孤儿。
function Get-ExcelProcessId {
    param($App, [int]$Retries = 10)

    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            $hwnd = [System.IntPtr]::new([int]$App.Hwnd)
            if ($hwnd -ne [System.IntPtr]::Zero) {
                $procId = 0
                [void][ExcelToolbox.Win32]::GetWindowThreadProcessId($hwnd, [ref]$procId)
                if ($procId -gt 0) { return $procId }
            }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    return 0
}

# 这个 PID 现在还活着吗？
#
# 【"查不到"和"查询失败"必须分开】。直接用 `-not (Get-Process -Id x)` 判断，
# 查询本身出错时也返回空，会被当成"进程已退出"——结论正好反了，
# 于是锁文件被删、孤儿留下。所以查询异常时一律按【还活着】处理：
# 宁可多保留一轮锁文件，也不能漏掉一个孤儿。
function Test-ProcessAlive {
    param([int]$ProcId)
    try {
        return [bool](Get-Process -Id $ProcId -ErrorAction Stop)
    }
    catch [Microsoft.PowerShell.Commands.ProcessCommandException] {
        return $false          # 明确的"没有这个进程"
    }
    catch {
        return $true           # 其它异常：判断不了，按还活着处理
    }
}

# 这个进程确实是我们登记的那一个吗？
# 【PID + 进程名 + 启动时间三者都要对上】——PID 会被系统回收复用，
# 只认 PID 就可能杀掉一个恰好拿到同一号码的无关进程。
function Test-OwnedProcess {
    param($Owned)
    try {
        $p = Get-Process -Id $Owned.Id -ErrorAction SilentlyContinue
        if (-not $p) { return $false }
        if ($p.ProcessName -ne 'EXCEL' -and $p.ProcessName -notin @('et','wps')) { return $false }
        if ($p.StartTime.Ticks -ne $Owned.Ticks) { return $false }
        return $true
    }
    catch { return $false }    # 属性读不到就不动它，宁可不杀也不误杀
}

# 登记失败时的兜底：把刚创建、还没来得及登记的实例就地关掉。
#
# 【不做这件事，"报错"本身就会制造孤儿】——登记失败抛异常，而那个 Excel
# 已经起来了、又没进 $script:ExcelOwned，Close-ExcelInstance 不认识它，
# 于是它永远留在后台。这正是我们要消灭的东西。
function Stop-UnregisteredExcel {
    param($App, [int]$ProcId)

    if ($App) {
        try { $App.DisplayAlerts = $false } catch {}
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if ($ProcId -gt 0) {
        # 【杀之前确认它确实是个 Excel/WPS 进程】。这个 PID 是几秒前从活着的
        # COM 对象上取到的，复用概率极低，但一次误杀的代价是用户的未保存文件，
        # 所以还是要认一下身份再动手。
        try {
            $p = Get-Process -Id $ProcId -ErrorAction SilentlyContinue
            if ($p -and ($p.ProcessName -eq 'EXCEL' -or $p.ProcessName -in @('et','wps'))) {
                Stop-Process -Id $ProcId -Force
            }
        } catch {}
    }
}

function Register-ExcelInstance {
    param($App)

    $procId = Get-ExcelProcessId $App
    if ($procId -le 0) {
        # 【不能默默放过】。登记不上意味着这个实例没人负责回收，
        # 脚本被 timeout 打断时它就永远留在后台，还会占住 .xlam 让后续构建失败。
        # 宁可当场失败，也不要留一个查不出来的隐患。
        Stop-UnregisteredExcel $App 0
        throw "取不到 Excel 实例的进程 ID（Hwnd 始终为 0）。已关闭该实例并中止。"
    }

    $p = $null
    try { $p = Get-Process -Id $procId -ErrorAction SilentlyContinue } catch {}
    if (-not $p) {
        Stop-UnregisteredExcel $App $procId
        throw "进程 $procId 已不存在，Excel 实例创建异常。"
    }

    $startTicks = 0L
    try { $startTicks = $p.StartTime.Ticks } catch {
        Stop-UnregisteredExcel $App $procId
        throw "取不到进程 $procId 的启动时间，无法安全登记。已关闭该实例并中止。"
    }

    $script:ExcelOwned += [pscustomobject]@{ Id = $procId; Ticks = $startTicks }

    # 从这里往下任何失败，都必须把实例收掉再抛
    try {
        if (-not (Test-Path $script:ExcelLockDir)) {
            $null = New-Item -ItemType Directory -Path $script:ExcelLockDir -Force
        }

        # 立刻落盘：中途被 timeout 打断也还有据可查。
        # 【先写临时文件再改名】——Set-Content 不是原子的，写到一半被打断
        # 会留下一个半截文件，下一轮解析不出来，等于没登记。
        $lines = @($script:ExcelOwned | ForEach-Object { "$($_.Id)|$($_.Ticks)" })
        $tmp = "$script:ExcelLockFile.tmp"
        Set-Content -LiteralPath $tmp -Encoding ASCII -Value $lines
        Move-Item -LiteralPath $tmp -Destination $script:ExcelLockFile -Force

        # 回读确认内容真的写进去了，而不只是文件存在
        $back = @(Get-Content -LiteralPath $script:ExcelLockFile -ErrorAction SilentlyContinue)
        if (($back -join "`n") -ne ($lines -join "`n")) {
            throw "PID 登记文件回读不一致：$script:ExcelLockFile"
        }
    }
    catch {
        $script:ExcelOwned = @($script:ExcelOwned | Where-Object { $_.Id -ne $procId })
        Stop-UnregisteredExcel $App $procId
        throw "登记 Excel 进程失败，已关闭该实例：$($_.Exception.Message)"
    }
}

# 收尾：关实例、释放 COM、确认本脚本创建的进程都已退出。
# 只碰登记过的 PID，不碰用户自己开的 Excel。
function Close-ExcelInstance {
    param($App)

    if ($App) {
        try { $App.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($App.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $App.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500

    foreach ($o in $script:ExcelOwned) {
        # 【进程属性访问全部要保护】。进程可能在 Get-Process 之后、读属性之前
        # 就退出了，那时访问 ProcessName / StartTime 会抛异常，
        # 整个清理循环就断在这里，后面的实例一个都收不掉。
        if (-not (Test-OwnedProcess $o)) { continue }

        try { Stop-Process -Id $o.Id -Force } catch {}

        # 【要等它真的退出再往下走】。Stop-Process 是异步的，立刻去删锁文件
        # 就会出现"锁文件没了、进程还在"的窗口——那正是孤儿逃掉的缝隙。
        #
        # 判断不了就按"没退出"处理——保留锁文件总比漏掉一个孤儿强。
        # 【不能用 `-not (Get-Process ...)` 当回退判据】：Get-Process 查询本身
        # 失败时也返回空，那会被当成"已退出"，恰好把结论倒过来。
        $exited = $false
        for ($i = 0; $i -lt 10; $i++) {
            if (-not (Test-ProcessAlive $o.Id)) { $exited = $true; break }
            Start-Sleep -Milliseconds 500
        }

        if (-not $exited) {
            Write-Host "    警告：Excel 进程 $($o.Id) 未能退出，锁文件保留以便下轮回收。" -ForegroundColor DarkYellow
            return      # 保留锁文件，下一轮 Clear-StaleExcel 会再收一次
        }
    }

    $script:ExcelOwned = @()
    if (Test-Path $script:ExcelLockFile) {
        Remove-Item -LiteralPath $script:ExcelLockFile -Force -ErrorAction SilentlyContinue
        # 删不掉要说出来：静默失败会让下一轮误以为还有孤儿要回收
        if (Test-Path $script:ExcelLockFile) {
            Write-Host "    警告：锁文件未能删除：$script:ExcelLockFile" -ForegroundColor DarkYellow
        }
    }
}

function New-RealExcel {
    [CmdletBinding()]
    param()

    # 每次创建前先把上一轮的孤儿收掉，避免越积越多把 .xlam 占住
    Clear-StaleExcel

    $xl = New-Object -ComObject Excel.Application
    Register-ExcelInstance $xl

    $path = ""
    $name = ""
    $ver  = ""
    try { $path = $xl.Path } catch {}
    try { $name = $xl.Name } catch {}
    try { $ver  = $xl.Version } catch {}

    if (Test-IsWpsHost $xl) {
        # 【登记之后的每一条异常出口都必须走 Close-ExcelInstance】。
        # 这里原先只调 Quit + Release：实例是关了，但 $script:ExcelOwned 里
        # 那条记录和磁盘上的锁文件都还在，下一轮 Clear-StaleExcel 会对着一个
        # 早就没了的 PID 空转，而真正的问题（WPS 劫持）反而被这些噪音盖住。
        Close-ExcelInstance $xl

        throw @"
COM 自动化拿到的不是 Microsoft Excel，而是 WPS：
    Name    = $name
    Version = $ver
    Path    = $path

WPS 安装后会接管 Excel 的 COM 注册（连 Excel.Application.16 也会被劫持），
并且刻意自称 "Microsoft Excel"，所以不校验就会在错误的宿主上构建和测试。

恢复方法：WPS → 设置/配置工具 → 兼容设置 → 取消 Office 文件关联，
或点"恢复 Office 默认设置"，然后重新运行本脚本。

（WPS 的兼容性请用 tests\probe-wps.ps1 单独验证，不要让它占着 COM 注册。）
"@
    }

    return $xl
}
