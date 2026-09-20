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
            # 有主加载宏和瘦加载器两个时，优先装加载器（它带自动更新）
            $loader = $found | Where-Object { $_.Name -like "*Loader*" } | Select-Object -First 1
            $pick = if ($loader) { $loader } else { $found[0] }
            Warn "找到多个加载宏，将安装：$($pick.Name)"
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
    Step "取消勾选并删除加载宏"
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        foreach ($a in @($xl.AddIns)) {
            try {
                if ($a.Name -eq $AddinName) { $a.Installed = $false }
            } catch {}
        }
    }
    catch { Warn "无法通过 Excel 取消勾选：$($_.Exception.Message)" }
    finally {
        if ($xl) {
            try { $xl.Quit() } catch {}
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    if (Test-Path -LiteralPath $Dest) {
        Remove-Item -LiteralPath $Dest -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $Dest) { Warn "文件删除失败：$Dest" } else { Good "已删除 $AddinName" }
    } else {
        Good "加载宏文件本来就不存在"
    }

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
$registered = $false
$excelVersion = ""
try {
    $xl = New-Object -ComObject Excel.Application
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
    if ($xl) {
        try { $xl.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
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
