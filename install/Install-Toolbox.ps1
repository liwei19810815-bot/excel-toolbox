<#
.SYNOPSIS
    面向业务人员的一键安装。双击同目录下的「安装工具箱.bat」即可。

.DESCRIPTION
    手工安装会卡住绝大多数非技术用户，这个脚本就是来解决那三个坎的：

      1.【解除网络锁定】从邮件或共享盘拿到的 .xlam 带着 Mark of the Web，
        不解锁 Excel 会直接禁用宏——而且报错信息里【完全不提】这回事，
        用户只会看到"工具箱没反应"。

      2.【注册加载项】手工路径是"文件 → 选项 → 加载项 → 管理 Excel 加载项
        → 转到 → 浏览"，六层菜单。这里用 COM 的 AddIns 集合做，
        等价于手工勾选，但用户什么都不用点。

      3.【受信任位置】自动更新模式下载荷跑在 %LOCALAPPDATA%\ExcelToolbox，
        不把它设为受信任位置，每次开 Excel 都弹宏安全警告。

    不需要管理员权限：只写 %APPDATA%、%LOCALAPPDATA% 和 HKCU。

.PARAMETER Uninstall
    取消勾选并删除已安装的加载宏。

.PARAMETER NoTrustedLocation
    不添加受信任位置。
    【批量部署时应该用这个开关】——几十台机器各自改注册表不好管理，
    受信任位置交给 IT 用组策略统一下发更规范，将来信息安全审计也说得清。

.EXAMPLE
    双击「安装工具箱.bat」

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Install-Toolbox.ps1 -NoTrustedLocation
#>
[CmdletBinding()]
param(
    [switch]$Uninstall,
    [switch]$NoTrustedLocation,
    [string]$AddinName
)

$ErrorActionPreference = "Stop"

# 控制台按 UTF-8 输出，否则中文在部分机器上是乱码
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$Here       = Split-Path -Parent $MyInvocation.MyCommand.Path
$AddInsDir  = Join-Path $env:APPDATA "Microsoft\AddIns"
$CacheDir   = Join-Path $env:LOCALAPPDATA "ExcelToolbox"

#-----------------------------------------------------------------------------
# 可靠地关掉【本程序自己创建的】那个 Excel 实例。
#
# 为什么不能只调 Quit()：注册加载项之后 Excel 会立刻加载它，加载器的
# Workbook_Open 又会打开载荷工作簿。这种状态下 Quit() 不一定能让进程退出——
# 实测会留下一个标题为"Excel 开始屏幕"的进程赖在后台，
# 接着就把下一次安装/卸载卡死在"检测到 Excel 正在运行"。
#
# 安装包是单独发给业务用户的，不能依赖仓库里的 _ExcelHost.ps1，
# 所以这里自带一份精简实现。
#
# 【只动自己创建的那个 PID】：从 Application.Hwnd 反查，绝不按进程名杀，
# 否则会连用户自己开着的 Excel 一起干掉。
#-----------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'ToolboxSetup.Win32').Type) {
    Add-Type -Namespace ToolboxSetup -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@
}

# 返回 @{ Id = <pid>; Ticks = <启动时间> }。取不到返回 $null。
#
# 【必须连启动时间一起记】：进程退出后 PID 会被系统回收复用，
# 只凭 PID 去强制结束，可能杀掉一个恰好拿到同一号码的无关进程——
# 在用户机器上那可能是他正在编辑的另一个 Excel。
function Get-OwnExcelProcess {
    param($App)
    for ($i = 0; $i -lt 10; $i++) {
        try {
            $hwnd = [System.IntPtr]::new([int]$App.Hwnd)
            if ($hwnd -ne [System.IntPtr]::Zero) {
                $procId = 0
                [void][ToolboxSetup.Win32]::GetWindowThreadProcessId($hwnd, [ref]$procId)
                if ($procId -gt 0) {
                    $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
                    if ($p) {
                        return @{ Id = $procId; Ticks = $p.StartTime.Ticks }
                    }
                }
            }
        } catch {}
        Start-Sleep -Milliseconds 200
    }
    return $null
}

function Stop-ExcelSafely {
    param($App, $Own)

    if ($App) {
        try { $App.DisplayAlerts = $false } catch {}
        try { foreach ($w in @($App.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
        try { $App.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if (-not $Own) { return }

    # 先给它几秒自己退
    for ($i = 0; $i -lt 10; $i++) {
        if (-not (Test-OwnProcessAlive $Own)) { return }
        Start-Sleep -Milliseconds 500
    }

    # 还在就强制结束——但只结束【确认是我们自己那一个】的
    if (Test-OwnProcessAlive $Own) {
        try { Stop-Process -Id $Own.Id -Force } catch {}
    }
}

# 这个进程还在，并且确实是我们启动的那一个（PID + 进程名 + 启动时间都要对上）
function Test-OwnProcessAlive {
    param($Own)
    try {
        $p = Get-Process -Id $Own.Id -ErrorAction SilentlyContinue
        if (-not $p) { return $false }
        if ($p.ProcessName -ne 'EXCEL') { return $false }
        if ($p.StartTime.Ticks -ne $Own.Ticks) { return $false }   # PID 被复用了
        return $true
    }
    catch { return $false }
}

function Say      ($m) { Write-Host $m }
function Step     ($m) { Write-Host "`n>> $m" -ForegroundColor Cyan }
function Good     ($m) { Write-Host "   [完成] $m" -ForegroundColor Green }
function Warn     ($m) { Write-Host "   [注意] $m" -ForegroundColor Yellow }
function Bad      ($m) { Write-Host "   [失败] $m" -ForegroundColor Red }

Say "============================================"
Say "  Excel 通用工具箱 - $(if ($Uninstall) { '卸载' } else { '安装' })"
Say "============================================"

#-----------------------------------------------------------------------------
# 找到要安装的 .xlam：优先用参数指定的，否则取脚本同目录下的那一个
#-----------------------------------------------------------------------------
if (-not $Uninstall) {
    if ($AddinName) {
        $Source = Join-Path $Here $AddinName
    } else {
        $found = @(Get-ChildItem -LiteralPath $Here -Filter *.xlam -ErrorAction SilentlyContinue)
        if ($found.Count -eq 0) {
            Bad "这个文件夹里没有找到 .xlam 加载宏文件。"
            Say ""
            Say "请确认解压时把所有文件放在了同一个文件夹里，"
            Say "这个脚本旁边应该有一个 ExcelToolbox.xlam。"
            exit 1
        }
        if ($found.Count -gt 1) {
            # 【同时出现多个是打包出错】。正式分发包里应该只放一个：
            # 要么独立版 ExcelToolbox.xlam，要么带自动更新的 ExcelToolboxLoader.xlam。
            #
            # 真遇到了就【优先装独立版】，不要自作主张装加载器——
            # 加载器启动时要去共享目录拉载荷，共享目录没配好或连不上时
            # 它会弹对话框，而模态框会把整个自动安装流程卡死
            # （实测就是这么失败的：Installed 属性设不进去）。
            # 独立版没有这个依赖，装上就能用，是更安全的默认。
            $standalone = $found | Where-Object { $_.Name -notlike "*Loader*" } | Select-Object -First 1
            $pick = if ($standalone) { $standalone } else { $found[0] }
            Warn "这个文件夹里有多个加载宏（正式分发包应该只放一个）。"
            Say  "     将安装：$($pick.Name)"
            if ($pick.Name -like "*Loader*") {
                Say  "     注意：加载器版需要能访问发布共享目录，否则启动时会报错。"
            }
            $Source = $pick.FullName
        } else {
            $Source = $found[0].FullName
        }
    }

    if (-not (Test-Path -LiteralPath $Source)) {
        Bad "找不到文件：$Source"
        exit 1
    }
    $AddinName = Split-Path -Leaf $Source
}

#-----------------------------------------------------------------------------
# 卸载时要卸掉【实际装进去的那个】，不能写死文件名。
#
# 本程序可能装的是独立版 ExcelToolbox.xlam，也可能是带自动更新的
# ExcelToolboxLoader.xlam（取决于分发包里放的是哪一个）。
# 卸载写死 ExcelToolbox.xlam 的话，装了加载器的机器上会卸不干净：
# 文件还在、加载项还勾着，用户以为卸载成功了。
#
# 判断依据是加载项目录里实际存在哪些属于本工具箱的文件。
#-----------------------------------------------------------------------------
if ($Uninstall -and -not $AddinName) {
    $installed = @(Get-ChildItem -LiteralPath $AddInsDir -Filter "ExcelToolbox*.xlam" -ErrorAction SilentlyContinue)
    if ($installed.Count -eq 0) {
        Say ""
        Say "加载项目录里没有找到本工具箱的文件，可能已经卸载过了。"
        Say "仍会继续清理受信任位置和缓存目录。"
        $AddinName = "ExcelToolbox.xlam"      # 占位，后面的删除步骤会走"本来就不存在"分支
    }
    elseif ($installed.Count -eq 1) {
        $AddinName = $installed[0].Name
    }
    else {
        # 两个都装过（比如先试了独立版又换成加载器版），一并卸掉
        Warn "检测到多个已安装的加载宏，将全部卸载：$(($installed.Name) -join ', ')"
        $AddinName = $installed[0].Name
        $script:ExtraToRemove = @($installed | Select-Object -Skip 1 | ForEach-Object { $_.Name })
    }
}

if (-not $AddinName) { $AddinName = "ExcelToolbox.xlam" }
$Dest = Join-Path $AddInsDir $AddinName

#-----------------------------------------------------------------------------
# 通过 COM 注册加载项时 Excel 不能正在运行，否则注册进去了但当前实例看不到，
# 用户会以为没装上。
#
# 这里按进程名判断，所以【任何】Excel 窗口都会让它等待——这是故意的：
# 注册动作影响的是整个 Excel 配置，不区分是哪个窗口打开的。
#
# 【绝对不自动杀 Excel】——用户可能有没保存的文件。只提示，让他自己关。
Step "检查 Excel 是否已关闭"
$waited = 0
while (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    if ($waited -eq 0) {
        Warn "检测到 Excel 正在运行。"
        Say  "   请先保存并关闭所有 Excel 窗口，然后本脚本会自动继续。"
        Say  "   （不会替你关闭，以免丢失未保存的内容）"
    }
    Start-Sleep -Seconds 2
    $waited += 2
    if ($waited -ge 180) {
        Bad "等待超过 3 分钟，Excel 仍在运行。请关闭后重新运行本程序。"
        exit 1
    }
}
Good "Excel 未运行"

#=============================================================================
# 卸载
#=============================================================================
if ($Uninstall) {
    $toRemove = @($AddinName)
    if ($script:ExtraToRemove) { $toRemove += $script:ExtraToRemove }

    Step "取消勾选并删除加载宏"
    $xl = $null
    $xlOwn = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xlOwn = Get-OwnExcelProcess $xl
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        foreach ($a in @($xl.AddIns)) {
            try {
                if ($toRemove -contains $a.Name) { $a.Installed = $false }
            } catch {}
        }
    }
    catch { Warn "无法通过 Excel 取消勾选：$($_.Exception.Message)" }
    finally {
        Stop-ExcelSafely $xl $xlOwn
    }

    $anyRemoved = $false
    foreach ($n in $toRemove) {
        $path = Join-Path $AddInsDir $n
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $path) { Warn "文件删除失败：$path" }
            else { Good "已删除 $n"; $anyRemoved = $true }
        }
    }
    if (-not $anyRemoved) { Good "加载宏文件本来就不存在" }

    # 【卸载必须把安全设置也撤掉】。装的时候改了 Excel 的受信任位置，
    # 卸载却留着，等于在用户机器上留下一条他不知情、也没人再需要的安全豁免。
    # 只删【本程序自己加的那一条】（键名 ExcelToolbox 且路径匹配），
    # 绝不碰用户或 IT 加的其他受信任位置。
    Step "撤销受信任位置设置"
    $removed = 0
    try {
        foreach ($vk in @(Get-ChildItem "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
                          Where-Object { $_.PSChildName -match '^\d+\.\d+$' })) {
            $base = "HKCU:\Software\Microsoft\Office\$($vk.PSChildName)\Excel\Security\Trusted Locations"
            if (-not (Test-Path $base)) { continue }
            foreach ($loc in @(Get-ChildItem $base -ErrorAction SilentlyContinue)) {
                $p = (Get-ItemProperty $loc.PSPath -Name Path -ErrorAction SilentlyContinue).Path
                if ($p -eq $CacheDir) {
                    Remove-Item -LiteralPath $loc.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    if (-not (Test-Path $loc.PSPath)) { $removed++ }
                }
            }
        }
        if ($removed -gt 0) { Good "已移除 $removed 条" } else { Good "没有需要移除的（可能安装时用了 -NoTrustedLocation）" }
    }
    catch { Warn "移除受信任位置时出错：$($_.Exception.Message)" }

    Step "删除缓存目录"
    if (Test-Path -LiteralPath $CacheDir) {
        Remove-Item -LiteralPath $CacheDir -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $CacheDir) {
            Warn "缓存目录删除失败（可能有文件被占用）：$CacheDir"
        } else {
            Good "已删除 $CacheDir"
        }
    } else {
        Good "缓存目录本来就不存在"
    }

    Say ""
    Say "卸载完成，安装时改动的东西都已还原。"
    exit 0
}

#=============================================================================
# 安装
#=============================================================================

Step "解除文件的网络锁定"
# 从网上/邮件/共享盘来的文件带 Zone.Identifier 标记，不解除 Excel 会禁用宏。
# 这一步没做，后面全都白搭，而且 Excel 的报错完全不会提示是这个原因。
try {
    Unblock-File -LiteralPath $Source -ErrorAction Stop
    Good "已解除锁定"
} catch {
    Warn "解除锁定时出错（多数情况下不影响）：$($_.Exception.Message)"
}

Step "复制到加载项目录"
if (-not (Test-Path -LiteralPath $AddInsDir)) {
    $null = New-Item -ItemType Directory -Path $AddInsDir -Force
}
# 记下复制之前那里有没有文件：失败回滚时要知道该删掉还是该还原
$destExistedBefore = Test-Path -LiteralPath $Dest
try {
    Copy-Item -LiteralPath $Source -Destination $Dest -Force -ErrorAction Stop
    Unblock-File -LiteralPath $Dest -ErrorAction SilentlyContinue
    Good $Dest
} catch {
    Bad "复制失败：$($_.Exception.Message)"
    exit 1
}

Step "在 Excel 中启用加载项"
$xl = $null
$xlOwn = $null
$registered = $false
$excelVersion = ""
try {
    $xl = New-Object -ComObject Excel.Application
    $xlOwn = Get-OwnExcelProcess $xl
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    try { $excelVersion = [string]$xl.Version } catch {}

    # AddIns.Add 等价于"浏览"，Installed = True 等价于"勾选"。
    #
    # 【只比文件名是不够的，必须连路径一起比】。Excel 的加载项列表里可能
    # 已经有一条同名但指向别处的旧记录（比如用户以前从桌面或 U 盘装过）。
    # 认领了那一条，启用的就是那个旧文件，用户会以为装了新版、
    # 实际跑的还是老的——而且这种错法从界面上完全看不出来。
    $addin = $null
    foreach ($a in @($xl.AddIns)) {
        try {
            if ($a.Name -eq $AddinName) {
                $existingPath = ""
                try { $existingPath = [string]$a.FullName } catch {}
                if ($existingPath -and ($existingPath -ne $Dest)) {
                    Warn "加载项列表里有一条同名但指向别处的记录，将改用本次安装的文件："
                    Say  "     旧：$existingPath"
                    Say  "     新：$Dest"
                    try { $a.Installed = $false } catch {}
                    continue        # 不认领它，下面重新 Add 正确路径
                }
                $addin = $a
                break
            }
        } catch {}
    }
    if (-not $addin) { $addin = $xl.AddIns.Add($Dest, $true) }
    $addin.Installed = $true
    $registered = $true
    Good "已启用"
}
catch {
    Bad "自动启用失败：$($_.Exception.Message)"
}
finally {
    # 【注册完必须确保这个实例真的退出】。启用加载项会让 Excel 立刻加载它，
    # 加载器的 Workbook_Open 又会打开载荷工作簿，此时 Quit() 不一定能退出——
    # 实测留下过一个"Excel 开始屏幕"进程赖在后台，
    # 把下一次安装/卸载直接卡死在"检测到 Excel 正在运行"。
    Stop-ExcelSafely $xl $xlOwn
}

# 【注册失败就把刚复制进去的文件收回来】。
#
# 留着它是最坏的结果：文件在加载项目录里、但没被勾选，用户看不到任何效果；
# 而下次再装时又会走"同名记录"那条分支，把问题搅得更复杂。
# 如果那个位置本来就有文件（覆盖安装），则不动——那是用户原有的东西。
if (-not $registered -and -not $destExistedBefore) {
    Step "回滚已复制的文件"
    Remove-Item -LiteralPath $Dest -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $Dest) {
        Warn "回滚失败，文件仍在：$Dest"
    } else {
        Good "已移除 $Dest"
    }
}
elseif (-not $registered -and $destExistedBefore) {
    Warn "启用失败。加载项目录里原本就有同名文件，已被本次安装覆盖，未做回滚。"
    Say  "     位置：$Dest"
}

# 【注册失败就不要再改安全设置】。否则会留下最糟糕的中间状态：
# 加载项没装上，但 Excel 的受信任位置已经被改了——用户什么功能都没得到，
# 却平白多了一条安全豁免，而且卸载程序也找不到加载项、更不会去撤它。
if (-not $registered -and -not $NoTrustedLocation) {
    Warn "加载项未能启用，跳过受信任位置设置（避免只改安全设置却没装上工具）。"
}
elseif (-not $NoTrustedLocation) {
    Step "把工具箱目录设为受信任位置"
    Say  "   这一步会修改 Excel 的安全设置，作用是让工具箱的宏不再每次弹警告。"
    Say  "   涉及目录：$CacheDir"
    Say  "   （批量部署时建议改由 IT 用组策略统一下发，见 -NoTrustedLocation 开关）"
    try {
        if (-not (Test-Path -LiteralPath $CacheDir)) {
            $null = New-Item -ItemType Directory -Path $CacheDir -Force
        }

        # 【必须用实际运行的 Excel 版本，不能去注册表里猜】。
        # HKCU\Software\Microsoft\Office 下常年堆着 11.0 / 12.0 / 14.0 / 15.0
        # 这些历史残留键，挑错了就是写进一个没人读的地方。
        # 上一步创建 COM 实例时拿到的 Version 才是真的（16.0 = 2016/2019/2021/2024/365）。
        if ([string]::IsNullOrWhiteSpace($excelVersion)) {
            throw "拿不到 Excel 版本号（上一步启用加载项可能已失败）"
        }

        $base = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security\Trusted Locations"

        # 【这个键不存在是常态，必须创建而不是跳过】。
        # 它只有在用户从信任中心界面手工加过一次受信任位置之后才会出现，
        # 绝大多数机器上根本没有。第一版写的是"不存在就 continue"，
        # 结果这一步在干净机器上【静默什么都不做】，用户还以为设好了。
        if (-not (Test-Path $base)) { $null = New-Item -Path $base -Force }

        # 已经加过就不重复加，否则每装一次多一条
        $exists = Get-ChildItem $base -ErrorAction SilentlyContinue | Where-Object {
            (Get-ItemProperty $_.PSPath -Name Path -ErrorAction SilentlyContinue).Path -eq $CacheDir
        }

        if ($exists) {
            Good "已存在，无需重复添加（Excel $excelVersion）"
        } else {
            $key = Join-Path $base "ExcelToolbox"
            $null = New-Item -Path $key -Force
            New-ItemProperty -Path $key -Name "Path"            -Value $CacheDir -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $key -Name "AllowSubfolders" -Value 1         -PropertyType DWord  -Force | Out-Null
            New-ItemProperty -Path $key -Name "Description"     -Value "Excel 通用工具箱" -PropertyType String -Force | Out-Null

            # 【写完要读回来确认】。注册表写入可能被组策略拦掉而不报错，
            # 只报"已设置"会让用户以为搞定了，结果每次开 Excel 还是弹警告。
            $back = (Get-ItemProperty -Path $key -Name Path -ErrorAction SilentlyContinue).Path
            if ($back -eq $CacheDir) {
                Good "已设置（Excel $excelVersion）"
            } else {
                Warn "写入后回读不一致，受信任位置可能未生效（多半被组策略限制）"
            }
        }
    }
    catch {
        Warn "设置受信任位置失败：$($_.Exception.Message)"
        Warn "不影响使用，但每次打开 Excel 可能会看到宏安全提示。"
    }
}

Say ""
Say "============================================"
if ($registered) {
    Say "  安装完成"
    Say "============================================"
    Say ""
    Say "请打开 Excel，功能区上会多出一个「工具箱」选项卡。"
    Say ""
    Say "如果没看到，通常是这两个原因之一："
    Say "  1. Excel 的宏被公司安全策略完全禁用了 —— 请联系 IT"
    Say "  2. 装的是另一个版本的 Excel —— 请确认只装了一套 Office"
} else {
    Say "  安装未完成"
    Say "============================================"
    Say ""
    Say "文件已经复制到了："
    Say "  $Dest"
    Say ""
    Say "但自动启用失败了，需要手工勾选一次："
    Say "  Excel → 文件 → 选项 → 加载项"
    Say "  → 最下方「管理」选择「Excel 加载项」→ 转到"
    Say "  → 在列表里勾选「$([IO.Path]::GetFileNameWithoutExtension($AddinName))」→ 确定"
    exit 1
}
