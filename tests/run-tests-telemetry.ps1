<#
.SYNOPSIS
    遥测的端到端回归：正常上报、端点挂掉、断网缓冲、恢复后重传。

.DESCRIPTION
    这套测试的重点不是"能不能上报成功"，而是【上报失败时会不会伤到用户】。

    遥测是纯粹的附加功能：用户想做的是删一行空行。收集端点挂了、网络断了、
    代理拦了、磁盘满了，都不允许让那一行删不掉，更不允许弹任何框。
    所以下面最关键的一条断言是「端点不可达时命令照常执行并照常返回结果」——
    这一条挂了，整个遥测特性就该从产品里拿掉。

    脚本自带一个最小 HTTP 收集端点（HttpListener），不需要外部服务。

.EXAMPLE
    timeout 600 powershell -ExecutionPolicy Bypass -File tests\run-tests-telemetry.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam",
    [int]$Port = 18731
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

$Xlam = Join-Path $RepoRoot "dist\$OutputName"
if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$BufferDir = Join-Path $env:LOCALAPPDATA "ExcelToolbox\telemetry"
$Endpoint  = "http://127.0.0.1:$Port/collect"

$script:pass = 0
$script:fail = 0

function Section($name) {
    Write-Host ""
    Write-Host "== $name ==" -ForegroundColor Cyan
}
function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") {
        $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected]" -ForegroundColor Red
        Write-Host "        实际 [$actual]" -ForegroundColor Red
    }
}
function Assert-True($cond, [string]$what) { Assert-Equal $true ([bool]$cond) $what }
function Assert-Match($actual, [string]$pattern, [string]$what) {
    if ("$actual" -like $pattern) {
        $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望匹配 [$pattern]" -ForegroundColor Red
        Write-Host "        实际     [$actual]" -ForegroundColor Red
    }
}

#-----------------------------------------------------------------------------
# 最小收集端点。
#
# 用 HttpListener 起在后台 Runspace 里：Start-Job 的开销太大（几秒），
# 而且跨进程取回收到的内容很麻烦。Runspace 共享同一个进程，
# 用一个同步的 ArrayList 就能把收到的 body 传回主线程。
#-----------------------------------------------------------------------------
$script:Received = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$script:Listener = $null
$script:Runspace = $null
$script:Handle   = $null

function Start-Collector {
    $script:Listener = [System.Net.HttpListener]::new()
    $script:Listener.Prefixes.Add("http://127.0.0.1:$Port/")
    $script:Listener.Start()

    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('listener', $script:Listener)
    $rs.SessionStateProxy.SetVariable('received', $script:Received)

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        while ($listener.IsListening) {
            try {
                $ctx = $listener.GetContext()
                $reader = [System.IO.StreamReader]::new($ctx.Request.InputStream, [Text.Encoding]::UTF8)
                [void]$received.Add($reader.ReadToEnd())
                $reader.Close()
                $ctx.Response.StatusCode = 200
                $ctx.Response.Close()
            } catch { break }
        }
    })
    $script:Runspace = $ps
    $script:Handle = $ps.BeginInvoke()
    Start-Sleep -Milliseconds 300
}

function Stop-Collector {
    try { if ($script:Listener -and $script:Listener.IsListening) { $script:Listener.Stop() } } catch {}
    try { if ($script:Listener) { $script:Listener.Close() } } catch {}
    try { if ($script:Runspace) { $script:Runspace.Dispose() } } catch {}
    $script:Listener = $null
    $script:Runspace = $null
}

function Clear-Buffer {
    if (Test-Path $BufferDir) { Remove-Item "$BufferDir\*" -Force -ErrorAction SilentlyContinue }
}
function Buffer-FileCount {
    if (-not (Test-Path $BufferDir)) { return 0 }
    return @(Get-ChildItem $BufferDir -File -ErrorAction SilentlyContinue).Count
}

$xl = $null
try {
    Write-Host "==> 启动 Excel 并加载加载宏" -ForegroundColor Cyan
    $xl = New-RealExcel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    $Run   = { param($id) $xl.Run("'$OutputName'!Toolbox_Run", $id) }
    $ws    = $wb.Worksheets.Item(1)
    $null = $xl.Run("'$OutputName'!Toolbox_SetSilent", $true)

    #=========================================================================
    Section "默认关闭：没配端点就什么都不发"

    $null = $xl.Run("'$OutputName'!Toolbox_SetTelemetry", $false, "")
    Clear-Buffer

    $null = $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  x  "
    $null = $ws.Activate(); $null = $ws.Range("A1").Select()
    $null = & $Run "text.cleanSpaces"

    Assert-Equal 0 (Buffer-FileCount) "遥测关闭时不写任何本地缓冲"
    Assert-Match ($xl.Run("'$OutputName'!Toolbox_TelemetryStatus")) "enabled=False*" "状态显示为关闭"

    #=========================================================================
    Section "开启后：命令执行会落盘"

    Start-Collector
    $null = $xl.Run("'$OutputName'!Toolbox_SetTelemetry", $true, $Endpoint)
    Clear-Buffer

    $null = $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  y  "
    $null = $ws.Activate(); $null = $ws.Range("A1").Select()
    $null = & $Run "text.cleanSpaces"

    Assert-Equal 1 (Buffer-FileCount) "执行后本地缓冲出现一个文件"

    #=========================================================================
    Section "上报：成功才删本地文件"

    $script:Received.Clear()
    $null = $xl.Run("'$OutputName'!Toolbox_FlushTelemetry")
    Start-Sleep -Milliseconds 800

    Assert-Equal 1 $script:Received.Count "收集端点收到一次上报"
    Assert-Equal 0 (Buffer-FileCount) "上报成功后本地缓冲已清空"

    $body = if ($script:Received.Count -gt 0) { [string]$script:Received[0] } else { "" }
    Assert-Match $body "*text.cleanSpaces*" "上报内容含 actionId"
    Assert-Match $body "*|ok|*"             "上报内容含执行结果"
    Assert-Match $body "*$env:USERNAME*"    "上报内容含用户名"

    # 这几条是承诺过「绝不采集」的，必须有断言守住
    Assert-True ($body -notmatch '\\')      "上报内容不含任何文件路径"
    Assert-True ($body -notlike "*  y  *")  "上报内容不含单元格数据"

    #=========================================================================
    Section "路径清洗：承诺「不采集文件名/路径」必须在代码层面成立"

    # Excel 的错误描述里经常自带完整路径，而我们在使用说明里向员工承诺了
    # 不采集文件名和路径。光说"我们没主动读路径"不够——错误描述是 Excel 给的，
    # 里面有什么不由我们决定。这几条就是守住那句承诺的。
    $Scrub = { param($s) $xl.Run("'$OutputName'!Toolbox_ScrubPaths", $s) }

    $cases = @(
        @{ In = "无法访问 'D:\财务\2026年薪资.xlsx'";        Bad = @('财务','薪资','xlsx','D:') }
        @{ In = '文件 "\\fs01\share\预算.xlsm" 被占用';        Bad = @('fs01','share','预算','xlsm') }
        @{ In = "找不到 C:\Users\zhangsan\Desktop\report.csv"; Bad = @('zhangsan','Desktop','report','csv') }
        @{ In = "打开 机密数据.xlsx 失败";                     Bad = @('机密数据','xlsx') }
        @{ In = "「客户名单.docx」已损坏";                     Bad = @('客户名单','docx') }
        # 以下几种是复验时发现会漏掉的形态，各补一条守着
        @{ In = "打开 报价单.xlsx，失败";                      Bad = @('报价单','xlsx') }   # 尾随中文标点
        @{ In = "(见 汇总表.xlsb) 第 3 行";                    Bad = @('汇总表','xlsb') }   # 括号包裹
        @{ In = "临时目录 %LOCALAPPDATA%Temp 不可写";          Bad = @('LOCALAPPDATA') }    # 环境变量形式
        @{ In = "路径 D:项目资料 无效";                        Bad = @('项目资料') }        # 裸盘符无斜杠
    )

    foreach ($c in $cases) {
        $out = & $Scrub $c.In
        $leaked = @($c.Bad | Where-Object { "$out" -like "*$_*" })
        Assert-Equal 0 $leaked.Count "清洗「$($c.In)」后不含敏感片段（残留：$($leaked -join ', ')）"
    }

    # 反向：正常的错误信息不该被抹成一片空白，否则 IT 没法按错误聚类
    $plain = & $Scrub "类型不匹配"
    Assert-Match $plain "*类型不匹配*" "不含路径的错误描述保持原样"

    #=========================================================================
    Section "端点挂掉：命令照常执行（最关键的一条）"

    Stop-Collector
    Clear-Buffer

    $null = $ws.Cells.Clear()
    $ws.Range("A1").Value2 = "  z  "
    $null = $ws.Activate(); $null = $ws.Range("A1").Select()

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = & $Run "text.cleanSpaces"
    $sw.Stop()

    Assert-Equal "z" $ws.Range("A1").Value2 "端点不可达时命令仍然正确执行"
    Assert-Match $r "*单元格*"               "端点不可达时结果照常返回"

    # 上报走的是启动/关闭时的批量路径，单条执行不该碰网络，
    # 所以这里不该有任何网络等待
    Assert-True ($sw.ElapsedMilliseconds -lt 3000) "单条命令执行不因遥测产生网络等待（实测 $($sw.ElapsedMilliseconds)ms）"

    #=========================================================================
    Section "断网：缓冲保留，不丢数据"

    Assert-Equal 1 (Buffer-FileCount) "端点不可达时记录仍落在本地"

    $null = $xl.Run("'$OutputName'!Toolbox_FlushTelemetry")
    Start-Sleep -Milliseconds 500
    Assert-Equal 1 (Buffer-FileCount) "上报失败时【不删】本地文件"

    #=========================================================================
    Section "关闭 Excel 时不再白等一次（端点已知挂掉）"

    # 启动时就发不出去的话，关闭时不该再试——网络不会因为用户点了关闭按钮
    # 就恢复，再试一次只是让 Excel 多卡一个超时。那正是用户最不耐烦、
    # 也最容易把锅扣到插件头上的时刻。
    #
    # 上一节的 FlushTelemetry 已经把「本会话端点已死」置位了，
    # 所以这次走 App_Shutdown 的那条路径应该立刻返回。
    $sw2 = [Diagnostics.Stopwatch]::StartNew()
    $null = $xl.Run("'$OutputName'!Toolbox_Shutdown")
    $sw2.Stop()
    Assert-True ($sw2.ElapsedMilliseconds -lt 500) "端点已知挂掉时，关闭路径不产生网络等待（实测 $($sw2.ElapsedMilliseconds)ms）"
    Assert-Equal 1 (Buffer-FileCount) "跳过重试不等于丢数据，缓冲仍在"

    #=========================================================================
    Section "恢复后重传"

    Start-Collector
    $script:Received.Clear()
    $null = $xl.Run("'$OutputName'!Toolbox_FlushTelemetry")
    Start-Sleep -Milliseconds 800

    Assert-Equal 1 $script:Received.Count "端点恢复后补传成功"
    Assert-Equal 0 (Buffer-FileCount)     "补传成功后缓冲清空"
    Assert-Match ([string]$script:Received[0]) "*text.cleanSpaces*" "补传的是之前攒下的那条"

    #=========================================================================
    Section "被前置校验拦下的命令也要记"

    Clear-Buffer
    $null = $ws.Activate()
    # 选中一个非 Range 对象让 RequiresSelection 判定失败：
    # 这里用图表工作表，它的 Selection 不是 Range
    $chart = $wb.Charts.Add()
    $null = $chart.Activate()
    $null = & $Run "text.cleanSpaces"
    $xl.DisplayAlerts = $false
    $chart.Delete()
    $null = $ws.Activate()

    $script:Received.Clear()
    $null = $xl.Run("'$OutputName'!Toolbox_FlushTelemetry")
    Start-Sleep -Milliseconds 800
    $blocked = if ($script:Received.Count -gt 0) { [string]$script:Received[0] } else { "" }
    Assert-Match $blocked "*blocked_nosel*" "「请先选中区域」这类拦截也被记录"

    #=========================================================================
    Section "环境还原"

    $null = $xl.Run("'$OutputName'!Toolbox_SetTelemetry", $false, "")
    Clear-Buffer
    Assert-Equal $true $xl.ScreenUpdating "ScreenUpdating 已还原"
    Assert-Equal $true $xl.EnableEvents   "EnableEvents 已还原"
}
catch {
    $script:fail++
    Write-Host ""
    Write-Host "测试过程异常：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Stop-Collector
    Close-ExcelInstance $xl
}

Write-Host ""
Write-Host "通过 $script:pass / 失败 $script:fail" -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
