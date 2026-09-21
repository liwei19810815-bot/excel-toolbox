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

# guide.* 是【使用配置条目】，不是命令：它们讲的是信任宏、受信任位置、
# 解除文件锁定这类 Excel 自身的设置，没有也不该有对应的 actionId。
# 拿它们去和注册表比对会全部被判成孤儿，所以这里分开处理。
$guideIds  = @($helpIds | Where-Object { $_ -like 'guide.*' })
$cmdHelpIds = @($helpIds | Where-Object { $_ -notlike 'guide.*' })

Write-Host "    注册命令 $($regIds.Count) 个，命令帮助 $($cmdHelpIds.Count) 条，使用配置 $($guideIds.Count) 条" -ForegroundColor DarkGray

$missing = @($regIds | Where-Object { $cmdHelpIds -notcontains $_ })
$orphan  = @($cmdHelpIds | Where-Object { $regIds -notcontains $_ })
$dupe    = @($helpIds | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

# 注册表里重复注册同一个 actionId 也要抓：后注册的会静默覆盖先注册的，
# 两条定义只有一条生效，而 Ribbon 一致性断言照样通过
$regDupe = @($regIds | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })

$errors = @()
foreach ($m in $missing) { $errors += "缺少帮助条目：$m（在 src\help\help.md 里加一节 ## $m）" }
foreach ($o in $orphan)  { $errors += "帮助条目对应的命令已不存在：$o" }
foreach ($d in $dupe)    { $errors += "帮助条目重复：$d" }
foreach ($d in $regDupe) { $errors += "命令重复注册：$d（后一条会静默覆盖前一条）" }

# 【每条命令都必须有示例】。"什么时候用/怎么用"是抽象描述，
# 业务人员看完照样不知道点下去会发生什么。示例写"处理前 → 处理后"，
# 是这次帮助改版的核心要求，漏写不会有任何其它征兆。
$helpText = Get-Content -LiteralPath $HelpMd -Raw -Encoding UTF8
$sections = [regex]::Split($helpText, '(?m)^##\s+') | Select-Object -Skip 1
foreach ($sec in $sections) {
    $id = ($sec -split '\r?\n', 2)[0].Trim()
    if ($id -like 'guide.*') {
        # 使用配置条目不要求示例，但必须有标题——侧边栏目录靠它显示
        if ($sec -notmatch '(?m)^标题[:：]\s*\S') {
            $errors += "使用配置条目缺少「标题:」一行：$id（侧边栏目录会显示成 id）"
        }
    }
    elseif ($sec -notmatch '(?m)^###\s+示例\s*$') {
        $errors += "缺少示例：$id（在 help.md 的该节里加一段 ### 示例，写处理前→处理后）"
    }
}

# 使用配置条目一条都不能少：它们是用户装完打不开时唯一的自助入口
$requiredGuides = @(
    'guide.macroTrust', 'guide.trustedLocation', 'guide.unblockFile',
    'guide.addinMissing', 'guide.multiOffice', 'guide.undoLimits', 'guide.telemetry'
)
foreach ($g in $requiredGuides) {
    if ($guideIds -notcontains $g) { $errors += "缺少使用配置条目：$g" }
}

if ($errors.Count -gt 0) {
    Write-Host "静态检查未通过：" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    exit 1
}
Write-Host "    一一对应，无缺失、无孤儿、无重复；$($cmdHelpIds.Count) 条命令都有示例" -ForegroundColor Green

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
    # 侧边栏目录
    #
    # 【每个命令都必须落在某个分组里】。分组是按 actionId 前缀派生的，
    # 将来加一个新前缀（比如 chart.*）而忘了在 CatalogGroups 里加一行，
    # 那些命令就会从目录里【静默消失】——功能还在、按钮还在，
    # 只是用户在帮助里永远找不到它们，没有任何测试会因此变红。
    #-------------------------------------------------------------------------
    Write-Host "==> 侧边栏目录：分组覆盖与内容" -ForegroundColor Cyan

    $groupIds = @(($xl.Run("'$OutputName'!Toolbox_HelpGroups") -split "`n") |
                  Where-Object { $_ } | ForEach-Object { ($_ -split '\|')[0] })
    Write-Host "    分组 $($groupIds.Count) 个" -ForegroundColor DarkGray

    $ungrouped = @()
    foreach ($id in $ids) {
        $g = $xl.Run("'$OutputName'!Toolbox_HelpGroupOf", $id)
        if (-not $g -or $groupIds -notcontains $g) { $ungrouped += "$id -> '$g'" }
    }
    if ($ungrouped.Count -gt 0) {
        $bad += "以下命令不属于任何目录分组：" + ($ungrouped -join '; ')
    } else {
        Write-Host "    $($ids.Count) 个命令全部归入分组，无遗漏" -ForegroundColor Green
    }

    # 目录里各组条目加起来要等于命令总数（既不重复也不漏）
    $catalogCount = 0
    foreach ($g in $groupIds) {
        if ($g -eq 'guide') { continue }
        $items = @(($xl.Run("'$OutputName'!Toolbox_HelpItems", $g) -split "`n") | Where-Object { $_ })
        $catalogCount += $items.Count
    }
    if ($catalogCount -ne $ids.Count) {
        $bad += "目录条目数 $catalogCount 与命令数 $($ids.Count) 对不上"
    } else {
        Write-Host "    目录条目数与命令数一致（$catalogCount）" -ForegroundColor Green
    }

    # 使用配置条目必须真的进了加载宏，并且目录里显示的是标题而不是 id
    $guideItems = @(($xl.Run("'$OutputName'!Toolbox_HelpItems", "guide") -split "`n") | Where-Object { $_ })
    if ($guideItems.Count -lt 7) {
        $bad += "使用配置条目只取到 $($guideItems.Count) 条，应有 7 条（构建可能没注入 title 列）"
    } else {
        $noTitle = @($guideItems | Where-Object { ($_ -split '\|')[1] -match '^guide\.' })
        if ($noTitle.Count -gt 0) {
            $bad += "使用配置条目在目录里显示成了 id 而不是标题：" + ($noTitle -join '; ')
        } else {
            Write-Host "    使用配置 $($guideItems.Count) 条，目录显示标题正常" -ForegroundColor Green
        }
    }

    # 正文渲染：命令条目要带可撤销标注，使用配置条目要带标题
    $render = $xl.Run("'$OutputName'!Toolbox_HelpRender", "text.cleanSpaces")
    if ($render -notmatch '可撤销') { $bad += "命令正文缺少可撤销标注：text.cleanSpaces" }
    if ($render -notmatch '示例')   { $bad += "命令正文里没有示例：text.cleanSpaces" }
    if ($render -match '###')       { $bad += "命令正文没清掉 Markdown 标记（侧边栏是纯文本控件）" }

    $renderGuide = $xl.Run("'$OutputName'!Toolbox_HelpRender", "guide.macroTrust")
    if ($renderGuide -notmatch '宏被禁用') { $bad += "使用配置正文渲染不出标题：guide.macroTrust" }

    #-------------------------------------------------------------------------
    # 环境体检：【只报告，不代改】
    #
    # 这条断言守的是一个承诺而不是一个功能：工具箱不替用户改安全设置。
    # 将来有人"顺手"加上自动修改注册表的代码，这里会红。
    #-------------------------------------------------------------------------
    Write-Host "==> 环境体检" -ForegroundColor Cyan
    $env = $xl.Run("'$OutputName'!Toolbox_EnvReport")
    foreach ($must in @('工具箱版本', '宿主程序', '功能区加载', '重算模式', '不会替你修改')) {
        if ($env -notmatch $must) { $bad += "环境体检报告缺少「$must」" }
    }
    # 查不到的事情要如实说查不到，不能假装检测
    if ($env -notmatch '无法自动检测') {
        $bad += "环境体检没有如实标注「受信任位置无法自动检测」"
    }
    if ($bad.Count -eq 0) { Write-Host "    报告内容完整，且明确声明不代改设置" -ForegroundColor Green }

    #-------------------------------------------------------------------------
    # 侧边栏窗体冒烟
    #
    # 【Controls.Add 失败是静默的】：窗体照样弹出来，只是一片空白。
    # 没有这条断言的话，运行时控件那套写法一旦被改坏，
    # 所有功能测试依然全绿，只有用户会看到一个空窗口。
    #-------------------------------------------------------------------------
    Write-Host "==> 帮助侧边栏（只实例化，不显示）" -ForegroundColor Cyan
    $smoke = $xl.Run("'$OutputName'!Toolbox_HelpPaneSmoke")
    if ($smoke -notlike 'OK|*') {
        $bad += "侧边栏窗体建不出来：$smoke"
    } else {
        $ctlCount = 0
        if ($smoke -match 'controls=(\d+)') { $ctlCount = [int]$Matches[1] }
        if ($ctlCount -lt 7) {
            $bad += "侧边栏控件只建出 $ctlCount 个，预期至少 7 个（搜索框/搜索钮/两个列表/正文/两个按钮）"
        }
        if ($smoke -notmatch 'body=True') {
            $bad += "侧边栏正文区是空的（欢迎文案没显示出来）"
        }
        if ($bad.Count -eq 0) {
            Write-Host "    窗体可实例化，控件 $ctlCount 个，正文已填充，可正常卸载" -ForegroundColor Green
        }
    }

    #-------------------------------------------------------------------------
    # 交互链：点分组 / 点功能 / 搜索 / 体检 / 定位条目
    #
    # 【"控件建得出来"证明不了"点下去有反应"】。运行时控件的事件接不到
    # 窗体代码模块上，全靠 clsPaneCtl 那层 WithEvents 包装；
    # 那层被改坏的表现是窗体照弹、控件都在、点谁都没反应，且不报错。
    #-------------------------------------------------------------------------
    Write-Host "==> 侧边栏交互链" -ForegroundColor Cyan
    $drive = $xl.Run("'$OutputName'!Toolbox_HelpPaneDrive")
    if ($drive -notlike 'OK|*') {
        $bad += "侧边栏交互链失败：$drive"
    } else {
        $checks = @(
            @{ Pat = 'items=([1-9]\d*)';  Msg = "点分组后功能列表是空的（分组 Click 事件没接上）" }
            @{ Pat = 'bodyLen=([1-9]\d*)'; Msg = "点功能后正文是空的（功能 Click 事件没接上）" }
            @{ Pat = 'search=True';        Msg = "搜索「求和是0」没命中文本转数值" }
            @{ Pat = 'env=True';           Msg = "体检按钮没出报告" }
            @{ Pat = 'entry=True';         Msg = "定位到 misc.parseId 失败（ShowEntry 跨分组定位不工作）" }
            @{ Pat = 'guide=True';         Msg = "定位到 guide.macroTrust 失败" }
            # 【这条守的是事件接线本身】。上面几条走的是直接调 OnPaneEvent，
            # 测的是处理逻辑，测不到 clsPaneCtl 那层 WithEvents。
            #
            # 真正会发生的故障是【忘了把包装对象存进集合】：对象一被回收，
            # 事件就静默失效，表现为"点了没反应且不报错"。
            # 这里断言活着的接收器个数，判据是确定的，
            # 不依赖"设 ListIndex 会不会触发 Click"这种随实现而变的行为。
            # 需要事件的控件有 6 个：搜索框、搜索钮、两个列表、体检钮、HTML 钮。
            # 正文框是只读显示区，不需要事件，所以不算在内。
            @{ Pat = 'sinks=([6-9]|\d{2,})'; Msg = "事件接收器少于 6 个（clsPaneCtl 包装对象没保活，点了不会有反应）" }
        )
        foreach ($c in $checks) {
            if ($drive -notmatch $c.Pat) { $bad += $c.Msg }
        }
        if ($bad.Count -eq 0) {
            Write-Host "    分组→功能→正文、搜索、体检、条目定位全部有反应" -ForegroundColor Green
        }

        # 事件是否真的送达，只作为【观察值】记录，不判失败——
        # 程序设 ListIndex 会不会触发 Click 取决于 MSForms 实现，
        # 拿它当判据会在别的 Office 版本上变成假红。
        if ($drive -match 'clickObserved=True') {
            Write-Host "    本机实测：程序设置 ListIndex 会触发 Click，事件确实送达" -ForegroundColor DarkGray
        } else {
            Write-Host "    本机观察：程序设置 ListIndex 不触发 Click（不影响功能，代码没有依赖这个行为）" -ForegroundColor DarkYellow
        }
    }

    # 真的显示一次再关掉：确认 Show 能成功、且不会留下窗体
    Write-Host "==> 侧边栏显示与关闭" -ForegroundColor Cyan
    $cycle = $xl.Run("'$OutputName'!Toolbox_HelpPaneShowCycle")
    if ($cycle -notlike 'OK|*') {
        $bad += "侧边栏 Show/Unload 失败：$cycle"
    } elseif ($cycle -notmatch 'visible=True') {
        $bad += "侧边栏 Show 之后 Visible 不是 True：$cycle"
    } else {
        Write-Host "    $cycle" -ForegroundColor Green
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
