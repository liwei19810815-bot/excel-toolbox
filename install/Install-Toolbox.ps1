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

.PARAMETER Install
.PARAMETER Uninstall
    指定动作，跳过交互式选择。【供 IT 批量部署用】——
    登录脚本、组策略里不能有交互，必须能直接指定装还是卸。
    双击运行时两个都不给，脚本会先探测当前状态再询问。

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
    [switch]$Install,
    [switch]$Uninstall,
    [switch]$NoTrustedLocation,
    [string]$AddinName,

    # AI 助手（独立的 Office.js 加载项）。只有分发包里带了 ai\ 目录时才有意义。
    # 【和工具箱本体是两个东西】：工具箱是 VBA 加载宏，装在本机；
    # AI 是网页加载项，从内网地址实时加载。两者各装各的，互不依赖。
    [switch]$WithAI,
    [switch]$AIOnly,

    # 【测试缝，正常使用不要传】。默认值就是生产行为，不新增任何分支逻辑。
    #
    # 去掉装 CA 那一步之后，「注册表」成了安装的最后一步，
    # 原先靠"放一份坏 ca.crt 逼 certutil 失败"来触发回滚的办法就没了。
    # 回滚逻辑（覆盖前备份、失败还原）必须继续有测试守着，否则它会悄悄烂掉。
    #
    # 传一个可写的测试键 → 整套安装器测试不再碰用户真实的 WEF 注册项；
    # 传一个非法键路径 → 注册表写入失败，走到回滚，验证 manifest 被还原。
    [string]$WefRoot = ""
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

#-----------------------------------------------------------------------------
# 单一入口：不给参数时，先看当前装没装，再决定问什么。
#
# 为什么不做成"安装.bat + 卸载.bat"两个文件：对业务用户来说，
# 桌面上多一个文件就多一次"我该点哪个"的犹豫。一个入口 + 按状态提示，
# 该干什么是脚本自己判断出来的，用户只需要确认。
#
# 【-Install / -Uninstall 参数保留】：IT 批量部署走登录脚本或组策略时
# 不能有交互，必须能指定动作直接跑。
#-----------------------------------------------------------------------------
function Get-InstalledAddins {
    return @(Get-ChildItem -LiteralPath $AddInsDir -Filter "ExcelToolbox*.xlam" -ErrorAction SilentlyContinue)
}

#-----------------------------------------------------------------------------
# AI 助手（Office.js 加载项）
#
# 分发包里可选带一个 ai\ 目录：
#     ai\manifest.template.xml   带 {{USER}} 和 {{GATEWAY}} 占位符
#     ai\gateway.txt             一行，内网网关地址
#
# 【不再需要 ca.crt】：本安装程序不往用户的受信任根存储里装任何证书。
# 前提是网关用【已被客户端信任】的证书（域内 PKI 下发或公网证书）。
#
# 没有这个目录就整段跳过——个人从 GitHub 下载的包里不会有它。
#-----------------------------------------------------------------------------
$AIDir        = Join-Path $Here "ai"
$AITargetDir  = Join-Path $env:LOCALAPPDATA "ExcelToolbox\ai"
$WefDeveloper = if ($WefRoot) { $WefRoot } else { "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" }

function Test-AIPackagePresent {
    return (Test-Path (Join-Path $AIDir "manifest.template.xml"))
}

# 【判据不能只看目录】。工具箱卸载会删掉整个 %LOCALAPPDATA%\ExcelToolbox，
# 而 AI 的 manifest 就在它的子目录里。只看目录的话，删完缓存之后
# 这里就判成"没装过"，AI 的注册表项再也没人清——Excel 会留下一个
# 指向已不存在目录的加载项，而卸载程序永远够不着它。
# 所以目录和注册表【任一存在】都算装过。
function Test-AIInstalled {
    if (Test-Path (Join-Path $AITargetDir "manifest.xml")) { return $true }
    try {
        $props = Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue
        if ($props -and $props.PSObject.Properties.Name -contains $AITargetDir) { return $true }
    } catch {}
    return $false
}

#-----------------------------------------------------------------------------
# 安装 AI 助手。
#
# 三步，都不需要管理员权限：
#   1. 按【当前用户名】生成专属 manifest（身份靠 URL 参数传给任务窗格——
#      Office.js 沙箱读不到用户名，只能这么注入，详见 docs\AI接入与白名单.md）
#   2. 写 HKCU\...\WEF\Developer 告诉 Excel 去哪找 manifest
#
# 【不装证书】。网关必须用已被客户端信任的证书，理由见下面第 3 段。
#-----------------------------------------------------------------------------
function Install-AI {
    Step "安装 AI 助手"

    $tpl = Join-Path $AIDir "manifest.template.xml"
    if (-not (Test-Path $tpl)) { Warn "分发包里没有 ai\manifest.template.xml，跳过。"; return $false }

    $gatewayFile = Join-Path $AIDir "gateway.txt"
    if (-not (Test-Path $gatewayFile)) { Bad "缺少 ai\gateway.txt（内网网关地址）。"; return $false }
    $gateway = (Get-Content -LiteralPath $gatewayFile -Raw).Trim().TrimEnd('/')
    if ([string]::IsNullOrWhiteSpace($gateway)) { Bad "ai\gateway.txt 是空的。"; return $false }

    # 【网关地址必须校验】。它会被原样拼进 manifest 的 XML，
    # 地址里有 & 或 < 就会让整个 manifest 变成非法 XML，
    # 而 Office 遇到非法 manifest 是【静默不加载】——
    # 用户只看到"按钮没出现"，没有任何报错可查。宁可现在就拦下来。
    if ($gateway -notmatch '^https?://[^\s<>&"'']+$') {
        Bad "ai\gateway.txt 的地址不合法：$gateway"
        Say  "     应形如 https://192.168.1.50:8443，且不能含空格或 < > & 等字符。"
        return $false
    }
    if ($gateway -notmatch '^https://') {
        Warn "网关用的是 http 而不是 https。Office 加载项通常要求 https，任务窗格可能加载不了。"
    }

    # 模板里的示例 GUID 没换过的话提醒一下——加载项的 Id 是全局唯一标识，
    # 两个组织用同一个 Id 会互相覆盖，而这种冲突极难排查
    $tplText = [IO.File]::ReadAllText($tpl, [Text.UTF8Encoding]::new($false))
    if ($tplText -match '7b2e4c91-6a38-4d5f-9e10-3c8a5f2d6b47') {
        Warn "manifest 模板里还是示例 GUID。正式分发前请换成你自己的（见 ai\README.txt）。"
    }

    if (-not (Test-Path $AITargetDir)) { $null = New-Item -ItemType Directory -Path $AITargetDir -Force }

    # 失败时要知道回滚到什么状态。
    #
    # 【不能只记"之前有没有"，还要把原内容留住】。重装场景下 manifest 和
    # 注册值都已经存在，我们会覆盖它们；如果后面某一步失败，只是"不删除"
    # 并不能还原——用户原来能用的那份配置已经被我们改掉了。
    # 备份的代价是几 KB 内存，换的是"失败之后至少回到原样"。
    # 【备份拿不到就别开工】。原先读失败是吞掉继续装，那等于明知
    # "失败了还不回去"还照样覆盖用户原有的配置——回滚承诺当场作废，
    # 而用户只会在失败信息里看到一行"无备份可还原"。
    # 没有可恢复的快照时，正确做法是根本不要动他的东西。
    $manifest = Join-Path $AITargetDir "manifest.xml"
    $manifestExistedBefore = Test-Path -LiteralPath $manifest
    $manifestBackup = $null
    $hasManifestBackup = $false
    if ($manifestExistedBefore) {
        try {
            $manifestBackup = [IO.File]::ReadAllText($manifest, [Text.UTF8Encoding]::new($false))
            $hasManifestBackup = $true
        }
        catch {
            Bad "读不到现有的 manifest，无法保证失败时能还原：$($_.Exception.Message)"
            Say "     为避免覆盖掉你现在能用的配置，已中止安装。"
            Say "     可手工删除后重试：$manifest"
            return $false
        }
    }

    # 【"没有原值"和"原值是空串"必须分开记】。只看 -ne "" 的话，
    # 原本就是空值的注册项在回滚时会被【删掉】而不是还原成空值——
    # 状态对不上，而且没有任何测试会因此变红。manifest 同理。
    $regBackup = $null
    $hasRegBackup = $false
    try {
        if (Test-Path $WefDeveloper) {
            $p = Get-ItemProperty -Path $WefDeveloper -ErrorAction Stop
            if ($p -and $p.PSObject.Properties.Name -contains $AITargetDir) {
                $regBackup = [string]$p.$AITargetDir
                $hasRegBackup = $true
            }
        }
    }
    catch {
        Bad "读不到现有的注册项，无法保证失败时能还原：$($_.Exception.Message)"
        Say "     为避免覆盖掉你现在能用的配置，已中止安装。"
        return $false
    }

    # --- 1. 生成专属 manifest ---
    #
    # 用户名要做 URL 编码：域账号里可能有空格或中文，直接拼进 URL 会把
    # SourceLocation 变成非法值，而 Office 遇到非法 manifest 是【静默不加载】，
    # 用户只会看到"按钮没出现"，完全查不出原因。
    $user = $env:USERNAME
    $userEnc = [uri]::EscapeDataString($user)

    try {
        $xml = $tplText.Replace("{{USER}}", $userEnc).Replace("{{GATEWAY}}", $gateway)

        # 替换完必须还是合法 XML。这一步是最后一道闸：
        # 上面校验的是网关地址，用户名走了 URL 编码，理论上都安全，
        # 但模板本身也可能被改坏——与其让 Office 静默不加载，不如这里就失败。
        [void]([xml]$xml)

        [IO.File]::WriteAllText($manifest, $xml, [Text.UTF8Encoding]::new($false))
        Good "manifest 已生成（用户：$user）"
    }
    catch {
        Bad "manifest 生成失败：$($_.Exception.Message)"
        Undo-AIPartialInstall -RegWritten $false -ManifestPath $manifest `
                              -ManifestExistedBefore $manifestExistedBefore `
                              -ManifestBackup $manifestBackup -RegBackup $regBackup `
                              -HasManifestBackup $hasManifestBackup -HasRegBackup $hasRegBackup
        return $false
    }

    # --- 2. 让 Excel 认识它 ---
    #
    # 【顺序是刻意的：先注册，最后才装证书】。
    # 原先是"先装 CA 再注册"，那样注册失败时证书已经进了用户的
    # 受信任根存储，而我们又不该去删它（可能是 IT 统一下发的，
    # 别的内网系统也在用）——于是必然留下"证书已受信但加载项没装上"的
    # 半成品状态，排查成本很高。
    # 反过来先注册：注册失败时根本还没碰证书，回滚干净。
    $regWritten = $false
    try {
        if (-not (Test-Path $WefDeveloper)) { $null = New-Item -Path $WefDeveloper -Force }
        New-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Value $AITargetDir `
                         -PropertyType String -Force | Out-Null
        $regWritten = $true

        $back = (Get-ItemProperty -Path $WefDeveloper -Name $AITargetDir -ErrorAction SilentlyContinue).$AITargetDir
        if ($back -ne $AITargetDir) { throw "注册表写入后回读不一致" }
        Good "已注册到 Excel"
    }
    catch {
        Bad "注册失败：$($_.Exception.Message)"
        Undo-AIPartialInstall -RegWritten $regWritten -ManifestPath $manifest `
                              -ManifestExistedBefore $manifestExistedBefore `
                              -ManifestBackup $manifestBackup -RegBackup $regBackup `
                              -HasManifestBackup $hasManifestBackup -HasRegBackup $hasRegBackup
        return $false
    }

    # --- 3. 证书？不装。 ---
    #
    # 【本安装程序不碰证书存储】。
    #
    # 早先这里会把内网自签 CA 装进当前用户的受信任根存储，因为
    # Office.js 要求任务窗格必须走 https，自签证书不被信任时任务窗格
    # 【空白且不报错】。但"往用户的受信任根存储里塞证书"是降低他整台机器
    # 防护等级的操作——那一个根证书能为任意域签发被信任的证书，
    # 影响远不止这个加载项。装个 Excel 插件不该有这种副作用。
    #
    # 现在的前提是：**网关用已经被客户端信任的证书**
    # （域内 PKI 统一下发，或公网证书）。这样什么都不用装。
    #
    # 如果 IT 最终只能提供自签证书，正确做法是让 IT 用组策略统一下发根证书，
    # 而不是由这个安装包替用户做这个决定。
    #
    # 详见 docs\AI接入与白名单.md。

    return $true
}

#-----------------------------------------------------------------------------
# 回滚一次失败的 AI 安装。
#
# 【注册项和 manifest 必须一起撤】。原先只删 manifest 不删注册项，
# 留下的是最糟的状态：Excel 仍会按注册项去加载一个已经不存在的 manifest，
# 而卸载逻辑靠"目录或注册项任一存在"判定，会把这台机器认成"装过"——
# 用户既用不了，也说不清自己到底装没装。
#
# 【证书不在回滚范围内】：调用点已保证证书是最后一步，
# 走到需要回滚时要么还没装，要么就是装证书这步自己失败的。
# 而且 CA 可能是 IT 统一下发的，别的内网系统也在用，不能替用户删。
#-----------------------------------------------------------------------------
function Undo-AIPartialInstall {
    param(
        [bool]$RegWritten,
        [string]$ManifestPath,
        [bool]$ManifestExistedBefore,
        [string]$ManifestBackup,
        [string]$RegBackup,

        # 【"有没有原值"必须单独传一个布尔】，不能靠 $RegBackup -ne "" 推断：
        # 原值本来就是空串时，推断的结果是"没有原值"，于是回滚去【删】它，
        # 而正确行为是还原成空串。参数声明成 [string] 之后 $null 会被
        # 绑定成 ""，两种情况在函数里根本区分不开。
        [bool]$HasManifestBackup,
        [bool]$HasRegBackup
    )

    # --- 注册项 ---
    # 原先有值就还原成原值，原先没有才删掉。
    # 一律删的话，重装失败会把用户本来好好的那条注册项也抹掉。
    if ($RegWritten) {
        try {
            if ($HasRegBackup) {
                New-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Value $RegBackup `
                                 -PropertyType String -Force | Out-Null
                Say "     已还原注册项原值。"
            } else {
                Remove-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Force -ErrorAction SilentlyContinue
                $still = (Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue)
                if ($still -and $still.PSObject.Properties.Name -contains $AITargetDir) {
                    Warn "回滚时未能删除注册项：$AITargetDir"
                } else {
                    Say "     已回滚注册项。"
                }
            }
        } catch { Warn "回滚注册项时出错：$($_.Exception.Message)" }
    }

    # --- manifest ---
    if ($ManifestExistedBefore) {
        if ($HasManifestBackup) {
            try {
                [IO.File]::WriteAllText($ManifestPath, $ManifestBackup, [Text.UTF8Encoding]::new($false))
                Say "     已还原原 manifest。"
            } catch { Warn "还原原 manifest 失败：$($_.Exception.Message)" }
        } else {
            # 正常走不到这里：备份读失败时 Install-AI 已经中止，根本不会开始覆盖。
            # 留着是兜底——万一将来有人把那道中止去掉，至少要如实说出来，
            # 而不是假装还原过。
            Warn "原 manifest 已被覆盖且无备份可还原：$ManifestPath"
        }
    }
    elseif (Test-Path -LiteralPath $ManifestPath) {
        Remove-Item -LiteralPath $ManifestPath -Force -ErrorAction SilentlyContinue
        Say "     已回滚生成的 manifest。"
    }
}

function Uninstall-AI {
    Step "卸载 AI 助手"

    # 注册项
    try {
        if (Test-Path $WefDeveloper) {
            $props = Get-ItemProperty -Path $WefDeveloper -ErrorAction SilentlyContinue
            if ($props -and $props.PSObject.Properties.Name -contains $AITargetDir) {
                Remove-ItemProperty -Path $WefDeveloper -Name $AITargetDir -Force -ErrorAction SilentlyContinue
                Good "已从 Excel 注销"
            } else { Good "注册项本来就不存在" }
        } else { Good "注册项本来就不存在" }
    }
    catch { Warn "注销时出错：$($_.Exception.Message)" }

    # manifest 目录
    if (Test-Path $AITargetDir) {
        Remove-Item -LiteralPath $AITargetDir -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path $AITargetDir) { Warn "目录删除失败：$AITargetDir" } else { Good "已删除 $AITargetDir" }
    } else { Good "目录本来就不存在" }

    # 【证书存储从头到尾没碰过】：安装时就没装任何证书，卸载自然也不用删。
    # 网关证书是 IT 统一管理的，和这个加载项的生命周期无关。

    # 【Office 的 WEF 缓存目录也不动】。
    #
    # 这里曾经写的是 Remove-Item "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*"，
    # 那是【清空所有 Office.js 加载项的缓存】——包括别的公司、别的项目装的，
    # 和我们毫无关系。为了清自己的残留去删别人的东西，代价完全不成比例。
    #
    # 而且本来也不需要：加载项能不能出现取决于上面那条 WEF\Developer 注册项，
    # 它已经删了，Excel 就不会再加载我们的加载项。缓存里剩下的只是文件。
    #
    # 万一真遇到"按钮还在但打不开"的残留，让用户手工清一次即可——
    # 那是极少数情况，不值得让每一次卸载都冒误删别人缓存的风险。
}

Say "============================================"
Say "  Excel 通用工具箱"
Say "============================================"

if ($AIOnly) { $Install = $true }

if (-not $Uninstall -and -not $Install) {
    $current = Get-InstalledAddins
    $hasAI   = Test-AIPackagePresent
    $aiOn    = Test-AIInstalled

    Say ""
    if ($current.Count -eq 0 -and -not $aiOn) {
        Say "当前状态：尚未安装"
        Say ""
        if ($hasAI) {
            Say "  [1] 全部安装：工具箱 + AI 助手   （直接回车即可）"
            Say "  [2] 只装工具箱"
            Say "  [3] 只装 AI 助手"
        } else {
            Say "  [1] 安装工具箱   （直接回车即可）"
        }
        Say "  [0] 退出"
        Say ""
        $choice = Read-Host "请输入数字后回车"
        if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }

        # 没带 AI 包时「1」就只是装工具箱；带了就是两个都装
        switch ($choice.Trim()) {
            "1" { $Install = $true; if ($hasAI) { $WithAI = $true } }
            "2" { if ($hasAI) { $Install = $true } else { Say ""; Bad "无法识别的输入「$choice」。"; exit 1 } }
            "3" { if ($hasAI) { $Install = $true; $AIOnly = $true; $WithAI = $true } else { Say ""; Bad "无法识别的输入「$choice」。"; exit 1 } }
            "0" { Say ""; Say "已取消，未做任何改动。"; exit 0 }
            default { Say ""; Bad "无法识别的输入「$choice」，未做任何改动。"; exit 1 }
        }
        $choice = "handled"
    }
    else {
        $state = @()
        if ($current.Count -gt 0) { $state += "工具箱（$(($current.Name) -join ', ')）" }
        if ($aiOn)                { $state += "AI 助手" }
        Say "当前状态：已安装 $($state -join '、')"
        Say ""
        Say "  [1] 重新安装 / 升级到本目录里的版本   （直接回车即可）"
        Say "  [2] 全部卸载"
        Say "  [0] 退出"
        Say ""
        $choice = Read-Host "请输入数字后回车"
        # 和"尚未安装"那一支保持一致：回车即默认动作。
        # 让回车报「无法识别的输入」比重装一次糟糕得多——重装是幂等的。
        if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
    }

    if ($choice -ne "handled") {
        switch ($choice.Trim()) {
            "1"     { $Install = $true; if (Test-AIPackagePresent) { $WithAI = $true } }
            "2"     {
                if ($current.Count -eq 0 -and -not (Test-AIInstalled)) {
                    Say ""; Say "还没有安装，无需卸载。"; exit 0
                }
                $Uninstall = $true
            }
            "0"     { Say ""; Say "已取消，未做任何改动。"; exit 0 }
            default { Say ""; Bad "无法识别的输入「$choice」，未做任何改动。"; exit 1 }
        }
    }
}

Say ""
Say "—— 开始$(if ($Uninstall) { '卸载' } else { '安装' }) ——"

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

    # 【AI 必须在删缓存目录之前卸】。AI 的 manifest 就在缓存目录的子目录里，
    # 顺序反了的话这里的判据会失效，注册表项留成悬空的。
    if (Test-AIInstalled) { Uninstall-AI }

    # 【顺带会删掉未上报的遥测缓冲】（telemetry\ 就在这个目录下）。
    # 这是有意为之：用户都把工具箱卸了，再留着他机器上的待上报数据没有道理。
    # 代价是 IT 看不到这台机器最后那几条事件——可以接受。
    # 注意升级不走这条路（安装流程不调用卸载），所以升级不会丢遥测。
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

# 只装 AI 时跳过工具箱本体
if ($AIOnly) {
    $ok = Install-AI
    Say ""
    Say "============================================"
    if ($ok) {
        Say "  AI 助手安装完成"
        Say "============================================"
        Say ""
        Say "请完全关闭 Excel 再重新打开，「开始」选项卡上会出现 AI 按钮。"
    } else {
        Say "  AI 助手安装未完成"
        Say "============================================"
        exit 1
    }
    exit 0
}

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

# AI 助手放在工具箱之后装：它失败不该影响工具箱已经装好这件事
$aiOk = $null
if ($WithAI -and $registered) {
    if (Test-AIPackagePresent) { $aiOk = Install-AI }
    else { Warn "分发包里没有 ai\ 目录，跳过 AI 助手。" }
}

Say ""
Say "============================================"
if ($registered) {
    Say "  安装完成"
    Say "============================================"
    Say ""
    Say "请打开 Excel，功能区上会多出一个「工具箱」选项卡。"
    if ($aiOk -eq $true) {
        Say "「开始」选项卡上还会出现 AI 按钮。"
    } elseif ($aiOk -eq $false) {
        Say ""
        Warn "工具箱装好了，但 AI 助手没装上（原因见上面）。工具箱本身不受影响。"
    }
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
