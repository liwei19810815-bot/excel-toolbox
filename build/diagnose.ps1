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

# 校验 COM 拿到的是真 Excel 而不是 WPS（WPS 会劫持 Excel 的 COM 注册并自称 Microsoft Excel）
. (Join-Path $PSScriptRoot "_ExcelHost.ps1")
$SharedCodeDir = Join-Path $RepoRoot "src\shared\code"
$ExcelCodeDir  = Join-Path $RepoRoot "src\excel\code"
$CodeDirs      = @($SharedCodeDir, $ExcelCodeDir)

$files = Get-ChildItem -Path $CodeDirs -Recurse -Include *.bas, *.cls, *.frm | Sort-Object FullName
if ($Count -gt $files.Count) { $Count = $files.Count }


$xl = $null
try {
    $xl = New-RealExcel
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

    $lastFile = $files[$Count - 1].FullName
    $base = $CodeDirs | Where-Object { $lastFile.StartsWith($_ + "\") } | Select-Object -First 1
    $lastName = $lastFile.Substring($base.Length + 1)
    if ($r -eq "pong") {
        Write-Host "OK   [$Count] 含 $lastName" -ForegroundColor Green
        $script:code = 0
    } else {
        Write-Host "BAD  [$Count] 含 $lastName -> 返回 $r" -ForegroundColor Red
        $script:code = 1
    }
}
catch {
    $lastFile = $files[$Count - 1].FullName
    $base = $CodeDirs | Where-Object { $lastFile.StartsWith($_ + "\") } | Select-Object -First 1
    $lastName = $lastFile.Substring($base.Length + 1)
    Write-Host "FAIL [$Count] 含 $lastName" -ForegroundColor Red
    Write-Host "     $($_.Exception.Message)" -ForegroundColor Red
    $script:code = 1
}
finally {
    Close-ExcelInstance $xl
}

exit $script:code
