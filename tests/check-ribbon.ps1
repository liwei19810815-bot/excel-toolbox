<#
.SYNOPSIS
    验证 customUI14.xml 真的被 Excel 接受并加载了。

.DESCRIPTION
    为什么要单独一个脚本：这是整套验证里唯一需要【可见 Excel】的检查。
    customUI 的 onLoad 只有在 Excel 真正创建功能区时才会触发，无界面模式下
    永远得不到信号。而可见模式会真的开窗口、加载功能区，COM 调用时序明显更
    不稳定——把它和功能测试混在一起，会让整套测试变得时灵时不灵。

    所以功能测试跑无界面（tests\run-tests.ps1），这一条单独跑。

    这个检查的价值：customUI14.xml 里只要有一处错误（重复 id、无效属性、
    坏掉的 imageMso），Excel 就会【静默】丢掉整个选项卡——VBA 照样编译通过，
    测试照样全绿，但用户打开 Excel 什么按钮都看不到。

    调用方需要套 timeout。退出码 0 = 通过。

.EXAMPLE
    timeout 180 powershell -ExecutionPolicy Bypass -File tests\check-ribbon.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot

# 校验 COM 拿到的是真 Excel 而不是 WPS（WPS 会劫持 Excel 的 COM 注册并自称 Microsoft Excel）
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")
$Xlam     = Join-Path $RepoRoot "dist\$OutputName"

if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

# ---------------------------------------------------------------------------
# 静态检查：每个按钮都必须把 label / supertip / enabled 交给回调去算。
#
# 为什么要查这个：modAction 里写了一整套 IsActionEnabled（问 modCaps 当前宿主
# 支不支持）和 ActionLabel（自动加"…"和"撤销 XXX"），但只要 XML 里漏挂
# getEnabled / getLabel，这些逻辑就【一次都不会被调用】——按钮照样亮着、
# 照样显示写死的旧名字，而 VBA 编译通过、功能测试全绿，没有任何东西会报警。
# 改版之前就是这个状态：59 个按钮里只有 1 个挂了 getEnabled。
# ---------------------------------------------------------------------------
$XmlPath = Join-Path $RepoRoot "src\package\customUI\customUI14.xml"
if (-not (Test-Path $XmlPath)) { throw "找不到 $XmlPath。" }

Write-Host "==> 静态检查：功能区回调接线" -ForegroundColor Cyan
[xml]$rx = Get-Content -LiteralPath $XmlPath -Raw -Encoding UTF8
$ns = New-Object System.Xml.XmlNamespaceManager($rx.NameTable)
$ns.AddNamespace("ui", "http://schemas.microsoft.com/office/2009/07/customui")

$wiringErrors = @()
$seenTags = @{}
foreach ($node in $rx.SelectNodes("//ui:button | //ui:toggleButton", $ns)) {
    $id = $node.GetAttribute("id")
    $tag = $node.GetAttribute("tag")

    if ([string]::IsNullOrWhiteSpace($tag)) {
        $wiringErrors += "$id : 缺少 tag（actionId）"
        continue
    }
    if ($seenTags.ContainsKey($tag)) {
        $wiringErrors += "$id : tag「$tag」与 $($seenTags[$tag]) 重复"
    } else {
        $seenTags[$tag] = $id
    }

    foreach ($cb in @("getLabel", "getEnabled", "getSupertip")) {
        if ([string]::IsNullOrWhiteSpace($node.GetAttribute($cb))) {
            $wiringErrors += "$id ($tag) : 缺少 $cb"
        }
    }
    # 写死的字面量会和注册表打架，两边说得不一样
    foreach ($lit in @("label", "supertip", "screentip")) {
        if (-not [string]::IsNullOrWhiteSpace($node.GetAttribute($lit))) {
            $wiringErrors += "$id ($tag) : 不该写死 $lit=，改由注册表生成"
        }
    }
}

if ($wiringErrors.Count -gt 0) {
    Write-Host "功能区接线检查未通过：" -ForegroundColor Red
    $wiringErrors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    exit 1
}
Write-Host "    $($seenTags.Count) 个按钮，回调接线完整" -ForegroundColor Green


$ok = $false
$xl = $null
try {
    Write-Host "==> 启动可见的 Excel（功能区只有在这种模式下才会创建）" -ForegroundColor Cyan
    $xl = New-RealExcel
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    Write-Host "==> 等待 onLoad 触发" -ForegroundColor Cyan

    # 【每次调用都要各自兜异常】。Excel 刚起来、还在忙着加载加载宏时，
    # COM 调用会被拒绝（RPC_E_CALL_REJECTED「应用程序正忙」）——
    # 这是【暂时】的，下一轮就好了。原先没有这层兜底，一次这样的瞬时
    # 错误会直接跳到外层 catch，整套判成失败并 exit 1，
    # 而实际上功能区完全正常。实测在 run-all 的连跑里偶发过一次：
    # 单独重跑立刻就绿，这种"随机变红"最能把人训练成忽略红色。
    #
    # 轮询上限也放宽到 30 次（约 24 秒）：连跑时 Excel 是冷启动，
    # 前面几套刚折腾完，12 秒不一定够。
    $sc = ""
    $lastErr = ""
    for ($i = 1; $i -le 30; $i++) {
        Start-Sleep -Milliseconds 800
        try {
            $sc = $xl.Run("'$OutputName'!Toolbox_SelfCheck")
            $lastErr = ""
            if ($sc -like "*ribbon=True") { $ok = $true; break }
        }
        catch {
            # 记下来但继续等——真的起不来的话，循环跑完照样会判失败
            $lastErr = $_.Exception.Message
        }
    }
    if (-not $ok -and $lastErr) {
        Write-Host "    最后一次调用仍在报错：$lastErr" -ForegroundColor DarkYellow
    }

    Write-Host "    $sc" -ForegroundColor DarkGray
    if ($ok) {
        Write-Host ""
        Write-Host "Ribbon 加载正常：customUI14.xml 已被 Excel 接受。" -ForegroundColor Green
    } else {
        Write-Host ""
        Write-Host "Ribbon 未加载。customUI14.xml 里多半有错误，Excel 静默丢弃了整个选项卡。" -ForegroundColor Red
        Write-Host "排查：Excel 选项 -> 高级 -> 常规 -> 勾选「显示加载项用户界面错误」后重新打开加载宏。" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "检查过程异常：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($xl) {
        Close-ExcelInstance $xl
    }
}

if ($ok) { exit 0 } else { exit 1 }
