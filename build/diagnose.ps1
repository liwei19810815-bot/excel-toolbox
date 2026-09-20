<#
.SYNOPSIS
    定位是哪个 VBA 模块导致编译失败。

.DESCRIPTION
    VBA 编译错误会弹出 VBE 模态框，把 COM 调用彻底卡死——脚本拿不到任何错误信息，
    只会一直挂着。所以只能反过来做：按依赖顺序只导入前 N 个模块，然后执行一次
    最简单的函数强制编译。第一个卡住或报错的 N，就是出问题的模块。

    每次只跑一个 N，由外部循环驱动并施加超时，这样模态框卡死的是一个一次性进程。
    脚本只会清理【自己启动的】Excel 进程，绝不碰调用前就已存在的实例。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File build\diagnose.ps1 -Count 5
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][int]$Count
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
$CodeDir  = Join-Path $RepoRoot "src\code"

$files = Get-ChildItem -Path $CodeDir -Recurse -Include *.bas, *.cls, *.frm | Sort-Object FullName
if ($Count -gt $files.Count) { $Count = $files.Count }

# 记下调用前就存在的 Excel 进程，清理时避开它们（那可能是用户自己开着的）
$preExisting = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false

    $wb = $xl.Workbooks.Add(-4167)

    for ($i = 0; $i -lt $Count; $i++) {
        $wb.VBProject.VBComponents.Import($files[$i].FullName) | Out-Null
    }

    # 加一个最简函数作为编译触发器
    $pingCode = @'
Public Function PingTest() As String
    PingTest = "pong"
End Function
'@
    $ping = $wb.VBProject.VBComponents.Add(1)   # vbext_ct_StdModule
    $ping.Name = "modPing"
    $ping.CodeModule.AddFromString($pingCode)

    $r = $xl.Run("PingTest")

    $lastName = $files[$Count - 1].FullName.Substring($CodeDir.Length + 1)
    if ($r -eq "pong") {
        Write-Host "OK   [$Count] 含 $lastName" -ForegroundColor Green
        $script:code = 0
    } else {
        Write-Host "BAD  [$Count] 含 $lastName -> 返回 $r" -ForegroundColor Red
        $script:code = 1
    }
}
catch {
    $lastName = $files[$Count - 1].FullName.Substring($CodeDir.Length + 1)
    Write-Host "FAIL [$Count] 含 $lastName" -ForegroundColor Red
    Write-Host "     $($_.Exception.Message)" -ForegroundColor Red
    $script:code = 1
}
finally {
    if ($xl) {
        try { $xl.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    Start-Sleep -Milliseconds 300
    Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        ForEach-Object { try { Stop-Process -Id $_.Id -Force } catch {} }
}

exit $script:code
