<#
.SYNOPSIS
    实测这台机器上 Office 能不能创建子进程。

.DESCRIPTION
    sidecar 伴生进程的启动方式取决于这一条：

      能起子进程 → Excel 打开工具箱时由 .xlam 直接拉起 sidecar。
                   不用开机自启、不用常驻、不用托盘图标，用户完全无感。

      起不来     → 退到「登录时启动」（Startup 快捷方式，不需要管理员），
                   sidecar 变成常驻进程，必须加托盘图标让用户看得见、关得掉。

    企业常部署一条 Defender ASR 规则「阻止所有 Office 应用程序创建子进程」
    （GUID D4F940AB-401B-4EFC-AADC-AD5F3C50688A），开了它前一条路就断。

    【为什么不能只读策略】
    读注册表/策略只能说明"配没配"，说明不了"拦不拦"：
      · 规则可能配成 Audit（记录但放行）
      · 可能没配 ASR，却被别的安全软件挡住
      · 策略可能还没下发到这台机器
    所以本脚本【真的驱动 Excel 跑一次 VBA Shell()，并验证进程确实起来了】。
    以实际结果为准。

    【必须在真实的用户机器上跑】
    开发机通常没有部署企业 ASR 策略，在开发机上通过不代表用户那边能用。

.NOTES
    只读探测：不改任何设置，不装任何东西。
    起的子进程是隐藏的 cmd.exe，ping 本机几次后自行退出，不联网、不写盘。

    ── 两种执行方式 ──────────────────────────────────────────────
    用户机上几乎都没开「信任对 VBA 工程对象模型的访问」，
    所以【不能只靠往工作簿里注入模块这一条路】，否则脚本在真正该用它的
    机器上跑不起来。两条路都支持，按下面顺序自动选：

      1. 探针加载宏（推荐，用户机上用这个）
         先在开发机上生成一次：
             check-office-child-process.ps1 -MakeProbe
         把生成的 AsrProbe.xlam 和本脚本一起拷到用户机，直接跑即可。
         走这条路【不需要】开 VBOM 信任。

      2. 运行时注入模块（开发机上用这个）
         机器已开 VBOM 信任时自动走这条，不需要探针加载宏。

    两条路都没有 → 判 INCONCLUSIVE 并给出操作指引。
    【绝不会因为"没测成"就报成通过。】

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\check-office-child-process.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\check-office-child-process.ps1 -MakeProbe
#>
[CmdletBinding()]
param(
    # 生成探针加载宏（在开发机上跑一次），不做探测。
    [switch]$MakeProbe,

    # 探针加载宏位置，默认与本脚本同目录。
    [string]$ProbePath
)

$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

if (-not $ProbePath) { $ProbePath = Join-Path $PSScriptRoot "AsrProbe.xlam" }

# 「阻止所有 Office 应用程序创建子进程」
$ASR_CHILD_PROCESS = 'D4F940AB-401B-4EFC-AADC-AD5F3C50688A'
$ASR_ACTION = @{ 0 = '未启用'; 1 = '拦截（Block）'; 2 = '仅审核（Audit，记录但放行）'; 6 = '警告（Warn）' }

$XL_OPEN_XML_ADDIN  = 55   # xlOpenXMLAddIn
$MSO_AUTOMATION_LOW = 1    # msoAutomationSecurityLow：自动化打开时不弹宏警告

# 探针模块。两条执行路径共用同一份源码，不要写成两份。
$PROBE_BAS = @"
Attribute VB_Name = "modAsrProbe"
Option Explicit

Public Function ProbeShell(ByVal cmdLine As String) As String
    On Error GoTo Failed
    Dim childPid As Double
    childPid = Shell(cmdLine, 0)          ' 0 = 隐藏窗口
    ProbeShell = "RET|" & CStr(childPid)
    Exit Function
Failed:
    ProbeShell = "ERR|" & Err.Number & "|" & Err.Description
End Function
"@

#------------------------------------------------------------------------------
# 把探针模块塞进工作簿
#
# 【必须走 Import(.bas)，不能用 CodeModule.AddFromString】。
# 本机实测：AddFromString 进去的工程，存成 .xlam / .xlsm 之后
# 【再也打不开】（Workbooks.Open 抛 0x800A03EC），而 SaveAs 那一步不报错——
# 又是一个"不报错但坏了"的坑。Import 出来的一切正常。
# 这也是 build.ps1 一直用 Import 的原因，别改回去。
#
# .bas 的两条硬性要求（踩中不报错，只会静默产出坏东西）：
#   · CRLF 换行
#   · 系统 ANSI 编码
# 和 build.ps1 里那段注释说的是同一件事。
#------------------------------------------------------------------------------
function Add-ProbeModule($wb) {
    $bas = Join-Path ([IO.Path]::GetTempPath()) ("modAsrProbe_{0}.bas" -f [guid]::NewGuid().ToString("N"))
    $text = ($PROBE_BAS -replace "`r`n", "`n") -replace "`n", "`r`n"
    [IO.File]::WriteAllText($bas, $text, [Text.Encoding]::Default)
    try   { $null = $wb.VBProject.VBComponents.Import($bas) }
    finally { Remove-Item $bas -Force -ErrorAction SilentlyContinue }
}

#------------------------------------------------------------------------------
# 「信任对 VBA 工程对象模型的访问」——和 build.ps1 同一套判断
#------------------------------------------------------------------------------
function Test-VbomTrust([string]$excelVersion) {
    $key = "HKCU:\Software\Microsoft\Office\$excelVersion\Excel\Security"
    if (-not (Test-Path $key)) { return $false }
    $v = Get-ItemProperty -Path $key -Name "AccessVBOM" -ErrorAction SilentlyContinue
    return ($null -ne $v -and $v.AccessVBOM -eq 1)
}

#==============================================================================
# 模式一：生成探针加载宏（开发机）
#==============================================================================
if ($MakeProbe) {
    Write-Host ""
    Write-Host "生成探针加载宏" -ForegroundColor Cyan
    Write-Host "============================================"

    $xl = $null
    try {
        $xl = New-RealExcel
        $xl.DisplayAlerts = $false

        if (-not (Test-VbomTrust $xl.Version)) {
            Write-Host ""
            Write-Host "  ❌ 本机没有开启「信任对 VBA 工程对象模型的访问」，无法生成探针。" -ForegroundColor Red
            Write-Host "     Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 →" -ForegroundColor Yellow
            Write-Host "     勾选「信任对 VBA 工程对象模型的访问」" -ForegroundColor Yellow
            exit 3
        }

        $wb = $xl.Workbooks.Add()
        Add-ProbeModule $wb
        # 【IsAddin 必须在 SaveAs 之前置上】，和 build.ps1 一样
        $wb.IsAddin = $true
        if (Test-Path $ProbePath) { Remove-Item $ProbePath -Force }
        $wb.SaveAs($ProbePath, $XL_OPEN_XML_ADDIN)
        $wb.Close($false)

        # 【生成完必须当场验证它真能被打开、真能被调用】。
        # 存出一个打不开的探针而不自知，等于把问题推给用户机那边才发现。
        $xl.AutomationSecurity = $MSO_AUTOMATION_LOW
        $check = $xl.Workbooks.Open($ProbePath)
        $probe = $xl.Run("'" + $check.Name + "'!ProbeShell", "cmd.exe /c exit")
        $check.Close($false)
        if ($probe -notmatch '^(RET|ERR)\|') {
            throw "探针存出来了，但自检时调用 ProbeShell 返回了意料之外的内容：$probe"
        }

        Write-Host ""
        Write-Host "  ✅ 已生成并自检通过：$ProbePath" -ForegroundColor Green
        Write-Host ""
        Write-Host "  把它和本脚本一起拷到【真实的用户机器】上，在那边跑："
        Write-Host "      powershell -ExecutionPolicy Bypass -File check-office-child-process.ps1"
        Write-Host "  用户机上不需要开 VBOM 信任。"
        exit 0
    }
    catch {
        Write-Host ""
        Write-Host "  ❌ 生成失败：$($_.Exception.Message)" -ForegroundColor Red
        exit 3
    }
    finally {
        if ($xl) { Close-ExcelInstance $xl }
    }
}

#==============================================================================
# 模式二：探测
#==============================================================================
Write-Host ""
Write-Host "Office 子进程能力探测" -ForegroundColor Cyan
Write-Host "============================================"
Write-Host "机器：$env:COMPUTERNAME    用户：$env:USERNAME"
Write-Host ""

#------------------------------------------------------------------------------
# 1. 策略状态（说明性，不作判据）
#------------------------------------------------------------------------------
Write-Host "== ASR 策略状态（仅供解释，不作判据）==" -ForegroundColor Cyan
try {
    $pref = Get-MpPreference -ErrorAction Stop
    $ids = @($pref.AttackSurfaceReductionRules_Ids)
    $acts = @($pref.AttackSurfaceReductionRules_Actions)

    $idx = -1
    for ($i = 0; $i -lt $ids.Count; $i++) {
        if ($ids[$i] -and ($ids[$i].ToString().ToUpper() -eq $ASR_CHILD_PROCESS)) { $idx = $i; break }
    }

    if ($idx -lt 0) {
        Write-Host "    这条规则没有配置" -ForegroundColor DarkGray
    } else {
        $a = [int]$acts[$idx]
        $desc = if ($ASR_ACTION.ContainsKey($a)) { $ASR_ACTION[$a] } else { "未知($a)" }
        $color = if ($a -eq 1) { "Yellow" } else { "DarkGray" }
        Write-Host "    「阻止 Office 创建子进程」= $desc" -ForegroundColor $color
    }
    Write-Host "    已配置的 ASR 规则共 $($ids.Count) 条" -ForegroundColor DarkGray
}
catch {
    Write-Host "    读不到（可能没装 Defender 或没权限）：$($_.Exception.Message)" -ForegroundColor DarkGray
}

#------------------------------------------------------------------------------
# 2. 真的试一次（判据）
#
# 【不能只看 Shell() 有没有抛异常】。被 ASR 拦时的表现不一定是抛错，
# 也可能是"返回了个 PID 但进程根本没起来"。必须回头验证进程真的存在过。
#------------------------------------------------------------------------------
Write-Host ""
Write-Host "== 实测：让 Excel 跑一次 VBA Shell() ==" -ForegroundColor Cyan

$verdict = "UNKNOWN"
$detail = ""
$xl = $null
try {
    $xl = New-RealExcel
    $xl.DisplayAlerts = $false

    #-- 选执行路径 -------------------------------------------------------------
    $useProbe = Test-Path $ProbePath
    $hasVbom = Test-VbomTrust $xl.Version

    if ($useProbe) {
        Write-Host "    执行方式：探针加载宏（不需要 VBOM 信任）" -ForegroundColor DarkGray
    }
    elseif ($hasVbom) {
        Write-Host "    执行方式：运行时注入模块（本机已开 VBOM 信任）" -ForegroundColor DarkGray
    }
    else {
        # 【这是"没测成"，不是"通过"】
        Write-Host ""
        Write-Host "    ⚠ 两条执行路径都不可用，无法探测。" -ForegroundColor Yellow
        Write-Host "      · 没找到探针加载宏：$ProbePath" -ForegroundColor Yellow
        Write-Host "      · 本机也没开「信任对 VBA 工程对象模型的访问」" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "    怎么办：在开发机上跑一次" -ForegroundColor Cyan
        Write-Host "        check-office-child-process.ps1 -MakeProbe"
        Write-Host "    把生成的 AsrProbe.xlam 和本脚本一起拷过来，再跑一次。"
        Close-ExcelInstance $xl
        exit 2
    }

    # 用一个会存活几秒的子进程，好让我们回头查它在不在。
    # ping 本机 5 次 ≈ 4 秒，不联网、不写盘。
    # marker 直接留在命令行里——【窗口是隐藏的，按窗口标题找不到】，
    # 只能按命令行找（Win32_Process.CommandLine）。
    $marker = "ToolboxAsrProbe_" + [guid]::NewGuid().ToString("N").Substring(0, 8)
    $cmd = "cmd.exe /c ping -n 5 127.0.0.1 >nul & rem $marker"

    if ($useProbe) {
        # 自动化打开时压掉宏警告。这不改用户的宏安全设置，
        # 只作用于我们这一个 Excel 实例，实例关掉就没了。
        $xl.AutomationSecurity = $MSO_AUTOMATION_LOW
        $wb = $xl.Workbooks.Open((Resolve-Path $ProbePath).Path)
        $target = "'" + $wb.Name + "'!ProbeShell"
    }
    else {
        $wb = $xl.Workbooks.Add()
        Add-ProbeModule $wb
        $target = "ProbeShell"
    }

    $raw = $xl.Run($target, $cmd)
    Write-Host "    VBA 返回：$raw" -ForegroundColor DarkGray

    if ($raw -like 'ERR|*') {
        $verdict = "BLOCKED"
        $detail = "Shell() 直接抛错：$raw"
    }
    else {
        # 【回头验证进程真的存在过】。拿到 PID 不等于进程起来了。
        # 注意：变量不能叫 $pid，那是 PowerShell 的自动变量（当前进程的 PID），
        # 写进去就等于在查我们自己。
        $childPid = 0
        if ($raw -match 'RET\|([0-9.]+)') { $childPid = [int][double]$Matches[1] }

        $seen = $false
        $how = ""
        for ($i = 0; $i -lt 20; $i++) {
            if ($childPid -gt 0 -and (Get-Process -Id $childPid -ErrorAction SilentlyContinue)) {
                $seen = $true; $how = "按 PID 命中"; break
            }
            # 兜底：按命令行里的 marker 找。窗口是隐藏的，按窗口标题找不到。
            $byCmd = @(Get-CimInstance Win32_Process -Filter "Name='cmd.exe'" -ErrorAction SilentlyContinue |
                       Where-Object { $_.CommandLine -and $_.CommandLine.Contains($marker) })
            if ($byCmd.Count -gt 0) { $seen = $true; $how = "按命令行命中"; break }
            Start-Sleep -Milliseconds 200
        }

        if ($seen) {
            $verdict = "ALLOWED"
            $detail = "子进程 PID $childPid 确认存在（$how）"
        } else {
            # PID 拿到了但进程查不到：可能是它太快退了，也可能是被拦了。
            # 【不能直接判成通过】——这正是 ASR 的典型表现之一。
            $verdict = "INCONCLUSIVE"
            $detail = "Shell() 返回了 PID $childPid，但 4 秒内没观察到该进程。可能退得太快，也可能被安全策略拦下。"
        }
    }

    $wb.Close($false)
}
catch {
    $verdict = "ERROR"
    $detail = $_.Exception.Message
}
finally {
    if ($xl) { Close-ExcelInstance $xl }
}

#------------------------------------------------------------------------------
# 3. 结论
#------------------------------------------------------------------------------
Write-Host ""
Write-Host "== 结论 ==" -ForegroundColor Cyan

switch ($verdict) {
    "ALLOWED" {
        Write-Host "    ✅ Office 可以创建子进程（$detail）" -ForegroundColor Green
        Write-Host ""
        Write-Host "    sidecar 可以采用【首选方案】：" -ForegroundColor Green
        Write-Host "      Excel 打开工具箱时由 .xlam 直接拉起，关 Excel 就退出。"
        Write-Host "      不用开机自启、不用常驻、不用托盘图标，用户完全无感。"
        exit 0
    }
    "BLOCKED" {
        Write-Host "    ❌ Office 被禁止创建子进程" -ForegroundColor Red
        Write-Host "       $detail" -ForegroundColor Red
        Write-Host ""
        Write-Host "    sidecar 必须走【退路】：" -ForegroundColor Yellow
        Write-Host "      登录时启动（Startup 文件夹快捷方式，不需要管理员权限）。"
        Write-Host "      代价是 sidecar 变成常驻进程，必须加托盘图标让用户看得见、关得掉。"
        exit 1
    }
    "INCONCLUSIVE" {
        Write-Host "    ⚠ 没测准，不要据此下结论" -ForegroundColor Yellow
        Write-Host "       $detail" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "    建议：重跑一次；仍然如此就按【被拦】处理（保守假设）。" -ForegroundColor Yellow
        exit 2
    }
    default {
        Write-Host "    探测本身出错，结论无效：$detail" -ForegroundColor Red
        Write-Host "    （这不代表被拦，只代表没测成。请排查后重跑。）" -ForegroundColor Red
        exit 3
    }
}
