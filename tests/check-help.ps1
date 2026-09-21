<#
.SYNOPSIS
    校验每个注册命令都有帮助条目，且帮助内容真的被注入进了加载宏。

.DESCRIPTION
    为什么需要这个：

    帮助漂移【不会让任何功能测试变红】。新加一个命令、忘了写帮助，
    59 个用例照样全绿，Ribbon 照样加载，用户点「帮助」也照样打开——
    只是那一条下面写着"（这条命令还没有写帮助正文）"。
    而帮助没人天天看，这种缺失可能几个月都没人发现。

    这与 check-ribbon.ps1 的静态接线断言、check-imagemso.ps1 的图标渲染检测
    是同一类检查：专门抓那些"不会让测试变红"的失败。

    检查三件事：
      1. 注册表里的每个 actionId 都有帮助条目（漏写）
      2. 帮助里没有已经不存在的 actionId（命令删了帮助没删）
      3. 帮助正文真的进了加载宏（构建步骤有没有生效）

.EXAMPLE
    timeout 300 powershell -ExecutionPolicy Bypass -File tests\check-help.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$RepoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $RepoRoot "build\_ExcelHost.ps1")

$Xlam = Join-Path $RepoRoot "dist\$OutputName"
if (-not (Test-Path $Xlam)) { throw "找不到 $Xlam。请先运行 build\build.ps1。" }

$HelpMd = Join-Path $RepoRoot "src\help\help.md"
if (-not (Test-Path $HelpMd)) { throw "找不到 $HelpMd。" }

#-----------------------------------------------------------------------------
# 先做静态检查：源码里的帮助条目 vs 源码里注册的命令
#
# 这一段不需要启动 Excel，失败时能给出最直接的提示。
#-----------------------------------------------------------------------------
Write-Host "==> 静态检查：帮助条目与命令注册表" -ForegroundColor Cyan

$helpIds = @(
    Select-String -LiteralPath $HelpMd -Pattern '^##\s+([A-Za-z]+\.[A-Za-z]+)\s*$' |
    ForEach-Object { $_.Matches[0].Groups[1].Value }
)

$modAction = Join-Path $RepoRoot "src\code\Core\modAction.bas"
$regIds = @(
    Select-String -LiteralPath $modAction -Pattern 'RegisterAction\s+"([A-Za-z]+\.[A-Za-z]+)"' |
    ForEach-Object { $_.Matches[0].Groups[1].Value }
)

Write-Host "    注册命令 $($regIds.Count) 个，帮助条目 $($helpIds.Count) 条" -ForegroundColor DarkGray

$missing = @($regIds | Where-Object { $helpIds -notcontains $_ })
$orphan  = @($helpIds | Where-Object { $regIds -notcontains $_ })
$dupe    = @($helpIds | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

# 注册表里重复注册同一个 actionId 也要抓：后注册的会静默覆盖先注册的，
# 两条定义只有一条生效，而 Ribbon 一致性断言照样通过
$regDupe = @($regIds | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

$errors = @()
foreach ($m in $missing) { $errors += "缺少帮助条目：$m（在 src\help\help.md 里加一节 ## $m）" }
foreach ($o in $orphan)  { $errors += "帮助条目对应的命令已不存在：$o" }
foreach ($d in $dupe)    { $errors += "帮助条目重复：$d" }
foreach ($d in $regDupe) { $errors += "命令重复注册：$d（后一条会静默覆盖前一条）" }

if ($errors.Count -gt 0) {
    Write-Host "静态检查未通过：" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    exit 1
}
Write-Host "    一一对应，无缺失、无孤儿、无重复" -ForegroundColor Green

#-----------------------------------------------------------------------------
# 再做运行时检查：帮助内容真的进了 .xlam
#
# 【静态检查过了不代表构建步骤生效】——help.md 写得再全，
# 构建脚本没把它注入进去的话，用户点帮助还是空的。
#-----------------------------------------------------------------------------
$bad = @()
$xl = $null
try {
    Write-Host "==> 运行时检查：帮助内容是否已注入加载宏" -ForegroundColor Cyan
    $xl = New-RealExcel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $null = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    $ids = @(($xl.Run("'$OutputName'!Toolbox_ListActions") -split "`n") | Where-Object { $_ })
    Write-Host "    加载宏报告 $($ids.Count) 个命令" -ForegroundColor DarkGray

    foreach ($id in $ids) {
        $has = $xl.Run("'$OutputName'!Toolbox_HasHelp", $id)
        if (-not $has) { $bad += $id }
    }

    if ($bad.Count -eq 0) {
        Write-Host "    全部 $($ids.Count) 个命令都能取到帮助正文" -ForegroundColor Green
    }

    #-------------------------------------------------------------------------
    # 搜索：用户会怎么描述问题
    #
    # 【搜的是症状，不是功能名】。用户打的是"求和是0"，不是"文本转数值"——
    # 他要是知道该用哪个功能，就直接点按钮了。
    # 这几条断言守住的就是"按症状能找到功能"这件事。
    #-------------------------------------------------------------------------
    Write-Host "==> 搜索：按症状能否找到功能" -ForegroundColor Cyan
    $searchCases = @(
        @{ Q = "求和是0";     Want = "text.toNumber" }
        @{ Q = "匹配不上";     Want = "text.cleanSpaces" }
        @{ Q = "合并文件夹";   Want = "merge.folder" }
        @{ Q = "逆透视";       Want = "data.unpivot" }
        @{ Q = "看串行";       Want = "misc.spotlight" }
        @{ Q = "身份证";       Want = "misc.parseId" }
    )

    foreach ($c in $searchCases) {
        $hit = @(($xl.Run("'$OutputName'!Toolbox_ResolveHelp", $c.Q) -split "`n") | Where-Object { $_ })
        if ($hit -contains $c.Want) {
            Write-Host ("    OK    「{0}」-> {1}" -f $c.Q, ($hit -join ", ")) -ForegroundColor DarkGreen
        } else {
            Write-Host ("    未命中 「{0}」期望 {1}，实际 {2}" -f $c.Q, $c.Want, ($hit -join ", ")) -ForegroundColor Red
            $bad += "搜索「$($c.Q)」没命中 $($c.Want)"
        }
    }

    # 搜不到时必须是明确的空结果，不能是"全部命令"——
    # 那样等于没搜，用户还得自己从 60 条里翻
    $none = @(($xl.Run("'$OutputName'!Toolbox_ResolveHelp", "紫色的大象") -split "`n") | Where-Object { $_ })
    if ($none.Count -eq 0) {
        Write-Host "    OK    无关词返回空结果" -ForegroundColor DarkGreen
    } else {
        Write-Host "    无关词返回了 $($none.Count) 条，应该是 0" -ForegroundColor Red
        $bad += "无关词没有返回空结果"
    }

    # 【标题命中不能屏蔽正文命中】。
    # 第一版的两轮匹配是"第一轮有结果就不跑第二轮"，
    # 搜"重复"时标题里带"重复"的两条会把正文里讲"重复"的
    # 「提取唯一值」整个挡掉——而那很可能正是用户要找的。
    $dupHits = @(($xl.Run("'$OutputName'!Toolbox_ResolveHelp", "重复") -split "`n") | Where-Object { $_ })
    if ($dupHits -contains "data.extractUnique") {
        Write-Host "    OK    「重复」同时命中标题与正文（$($dupHits.Count) 条）" -ForegroundColor DarkGreen
    } else {
        Write-Host "    「重复」漏掉了 data.extractUnique，实际：$($dupHits -join ', ')" -ForegroundColor Red
        $bad += "标题命中屏蔽了正文命中"
    }

    #-------------------------------------------------------------------------
    # 【搜索绝不能直接执行命令】
    #
    # 注册表里有 data.deleteDuplicates、file.batchRename、formula.breakLinks
    # 这类会改数据甚至改磁盘文件的命令。用户在搜索框里打几个字按回车，
    # 数据就被改了——那不是"省一次点击"，是从输入框静默触发破坏性操作。
    # 这条断言守住"搜索只负责找，不负责做"。
    #-------------------------------------------------------------------------
    Write-Host "==> 搜索不得执行命令" -ForegroundColor Cyan
    $probe = $xl.Workbooks.Add(-4167)
    $pw = $probe.Worksheets.Item(1)
    $null = $pw.Cells.Clear()
    $pw.Range("A1").Value2 = "x"
    $pw.Range("A2").Value2 = "x"
    $null = $pw.Activate()
    $null = $pw.Range("A1:A2").Select()

    # 这个词应该只命中「删除重复值」一条——正是最危险的单命中场景
    $null = $xl.Run("'$OutputName'!Toolbox_SearchHelp", "删除重复值")
    Start-Sleep -Milliseconds 300

    $a1 = "$($pw.Range('A1').Value2)"
    $a2 = "$($pw.Range('A2').Value2)"
    if ($a1 -eq "x" -and $a2 -eq "x") {
        Write-Host "    OK    搜索「删除重复值」没有动数据" -ForegroundColor DarkGreen
    } else {
        Write-Host "    搜索执行了命令！A1=[$a1] A2=[$a2]（原本都是 x）" -ForegroundColor Red
        $bad += "搜索直接执行了破坏性命令"
    }
    $xl.DisplayAlerts = $false
    $probe.Close($false)
}
catch {
    Write-Host "检查过程异常：$($_.Exception.Message)" -ForegroundColor Red
    $bad += "(异常)"
}
finally {
    Close-ExcelInstance $xl
}

Write-Host ""
if ($bad.Count -eq 0) {
    Write-Host "帮助系统检查通过。" -ForegroundColor Green
    exit 0
}

Write-Host "以下命令在加载宏里取不到帮助正文，构建的注入步骤可能没生效：" -ForegroundColor Red
$bad | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
exit 1
