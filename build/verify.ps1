<#
.SYNOPSIS
    验证构建产物：加载 dist\ExcelToolbox.xlam 并调用自检入口。

.DESCRIPTION
    build.ps1 里的 Import 只是把源码塞进 VBA 工程，【不做编译】。
    语法错误、缺失引用、类型不匹配要到第一次执行时才暴露。
    本脚本通过 Application.Run 调一次 Toolbox_SelfCheck 强制触发编译：
    只要它能返回，整个工程就是编译通过的。

    退出码 0 = 通过，1 = 失败。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\verify.ps1
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

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$xl = $null
$addin = $null
$failed = $false

try {
    Write-Step "启动 Excel 并加载加载宏"
    $xl = New-RealExcel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    # 需要一个可见工作簿，否则 Application.Run 无处执行
    $wb = $xl.Workbooks.Add()

    $addin = $xl.Workbooks.Open($Xlam)

    # VBA 是【按需编译】的：调一个函数只会编译它用到的那条路径，别的过程里的语法错误
    # 一样能蒙混过关，直到用户点到那个按钮才炸。所以必须显式触发整工程编译。
    # 578 = VBE 的「编译 VBAProject」命令。
    Write-Step "整工程编译（VBE 命令 578）"
    # 「编译 VBAProject」编译的是 VBE 里【当前激活】的工程。刚 Add 出来的空工作簿
    # 也是一个工程，不把焦点挪到加载宏上，编译的就是那个空壳，等于什么都没验证。
    # 激活任一代码窗格即可把对应工程设为当前工程。
    $addin.VBProject.VBComponents.Item("modAction").CodeModule.CodePane.Show()

    $compile = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) {
        Write-Host "    找不到编译命令，跳过（仅做运行时验证）" -ForegroundColor Yellow
    } else {
        # 编译失败会弹 VBE 模态框把这里永久卡住，所以调用方必须加超时：
        #   timeout 90 powershell -File build\verify.ps1
        # 超时即视为编译失败。
        $compile.Execute()
        Write-Host "    编译通过" -ForegroundColor DarkGray
    }

    Write-Step "调用 Toolbox_SelfCheck"
    $result = $xl.Run("'$OutputName'!Toolbox_SelfCheck")

    Write-Host "    $result" -ForegroundColor DarkGray
    if ($result -notlike "OK|*") {
        throw "自检返回了非预期结果：$result"
    }

    Write-Host ""
    Write-Host "验证通过：VBA 工程编译无误。" -ForegroundColor Green
}
catch {
    $failed = $true
    Write-Host ""
    Write-Host "验证失败：" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "多半是 VBA 编译错误。用 VBE 定位：" -ForegroundColor Yellow
    Write-Host "  打开 Excel -> 加载 dist\$OutputName -> Alt+F11 -> 调试 -> 编译" -ForegroundColor Yellow
}
finally {
    if ($xl) {
        try { $xl.DisplayAlerts = $false } catch {}
        try {
            foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} }
        } catch {}
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($failed) { exit 1 } else { exit 0 }
