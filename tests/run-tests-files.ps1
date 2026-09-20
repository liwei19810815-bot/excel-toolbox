<#
.SYNOPSIS
    文件与工作表类工具的回归测试（M3 / M4 / M5）。

.DESCRIPTION
    单独一个脚本，因为这批工具和前面的性质完全不同：
      - 它们会在【磁盘上】建文件、改文件名、导出文件，全部【不可撤销】；
      - 需要真实的临时目录和真实的 xlsx 夹具，准备和清理都比内存用例重得多。

    所有夹具都建在 %TEMP% 下一个随机命名的目录里，跑完无条件删除。
    脚本只清理本次启动的 Excel 进程，不碰用户已经开着的实例。

    调用方必须套 timeout。退出码 0 = 全过。

.EXAMPLE
    timeout 600 powershell -ExecutionPolicy Bypass -File tests\run-tests-files.ps1
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

$SandBox = Join-Path $env:TEMP ("ExcelToolboxTest_" + [guid]::NewGuid().ToString("N"))

$script:pass = 0
$script:fail = 0
$script:section = ""

function Section($name) {
    $script:section = $name
    Write-Host ""
    Write-Host "== $name ==" -ForegroundColor Cyan
}
function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected]" -ForegroundColor Red
        Write-Host "        实际 [$actual]" -ForegroundColor Red
    }
}
function Assert-Match($actual, [string]$pattern, [string]$what) {
    if ("$actual" -like $pattern) { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望匹配 [$pattern]" -ForegroundColor Red
        Write-Host "        实际     [$actual]" -ForegroundColor Red
    }
}
function Set-Grid($ws, [int]$top, [int]$left, $rows) {
    $h = $rows.Count
    $w = $rows[0].Count
    $arr = [Array]::CreateInstance([object], $h, $w)
    for ($r = 0; $r -lt $h; $r++) { for ($c = 0; $c -lt $w; $c++) { $arr[$r, $c] = $rows[$r][$c] } }
    $ws.Cells($top, $left).Resize($h, $w).Value2 = $arr
}
function Select-On($ws, [string]$addr) {
    $null = $ws.Activate()
    $null = $ws.Range($addr).Select()
}
function Sheet-Exists($wb, [string]$name) {
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { return $true } }
    return $false
}

$xl = $null
try {
    New-Item -ItemType Directory -Path $SandBox | Out-Null
    Write-Host "沙箱目录：$SandBox" -ForegroundColor DarkGray

    Write-Host "==> 启动 Excel 并加载加载宏" -ForegroundColor Cyan
    $xl = New-RealExcel
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add(-4167)
    $null = $xl.Workbooks.Open($Xlam)

    $Run   = { param($id) $xl.Run("'$OutputName'!Toolbox_Run", $id) }
    $Param = { param($k, $v) $xl.Run("'$OutputName'!Toolbox_SetParam", $k, $v) }

    $xl.Run("'$OutputName'!Toolbox_SetSilent", $true)

    #==========================================================================
    Section "M3 工作表管理"

    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "源数据"
    $ws.Cells.Clear()
    Set-Grid $ws 1 1 @(
        @("区域", "金额"),
        @("华东", 10), @("华北", 20), @("华东", 30), @("华南", 40)
    )
    & $Param "keyColumn" "1"
    Select-On $ws "A1:B5"
    $r = & $Run "sheet.splitByColumn"
    Assert-Match $r "已按第 1 列拆分为 3 个工作表*" "按列拆表：分组数正确"
    Assert-Equal $true (Sheet-Exists $wb "华东") "按列拆表：生成了以列值命名的表"
    $wsHD = $wb.Worksheets.Item("华东")
    Assert-Equal "区域" $wsHD.Range("A1").Value2 "按列拆表：带上了标题行"
    Assert-Equal 30 $wsHD.Range("B3").Value2 "按列拆表：同组两行都在"

    # 表名里的非法字符和超长必须被处理掉，否则 Worksheets.Add 直接报错
    $wsBad = $wb.Worksheets.Add()
    $wsBad.Name = "非法名测试"
    Set-Grid $wsBad 1 1 @(
        @("名称", "值"),
        @("a/b:c*d?e[f]g", 1),
        @("这是一个非常非常非常非常非常非常长的名称超过三十一个字符限制", 2)
    )
    & $Param "keyColumn" "1"
    Select-On $wsBad "A1:B3"
    $r = & $Run "sheet.splitByColumn"
    Assert-Match $r "*拆分为 2 个工作表*" "按列拆表：非法字符与超长名称不会导致失败"

    $r = & $Run "sheet.createIndex"
    Assert-Match $r "已生成目录*" "生成目录"
    Assert-Equal $true (Sheet-Exists $wb "目录") "生成目录：表已建立"
    $wsIdx = $wb.Worksheets.Item("目录")
    Assert-Equal 1 $wsIdx.Hyperlinks.Count.CompareTo(0) "生成目录：包含可跳转的超链接"

    $hiddenWs = $wb.Worksheets.Add()
    $hiddenWs.Name = "藏起来"
    $hiddenWs.Visible = 2          # xlSheetVeryHidden，界面上根本取消不了隐藏
    $r = & $Run "sheet.showAll"
    Assert-Match $r "已显示 1 张隐藏的工作表*" "显示所有隐藏表：含深度隐藏"
    Assert-Equal -1 $wb.Worksheets.Item("藏起来").Visible "显示所有隐藏表：确实可见了"

    & $Param "ascending" "true"
    $r = & $Run "sheet.sort"
    Assert-Match $r "已按名称升序排列*" "工作表排序"

    #==========================================================================
    Section "M3 合并所有工作表（按标题名对齐）"

    $mergeWb = $xl.Workbooks.Add(-4167)
    $m1 = $mergeWb.Worksheets.Item(1)
    $m1.Name = "一月"
    Set-Grid $m1 1 1 @(@("姓名", "销量"), @("张", 1), @("李", 2))

    $m2 = $mergeWb.Worksheets.Add()
    $m2.Name = "二月"
    # 故意把列顺序换一下，再多一列 —— 按列位置拼接的话这里就会整体错位
    Set-Grid $m2 1 1 @(@("销量", "姓名", "备注"), @(3, "王", "新增"))

    $null = $mergeWb.Activate()
    $r = & $Run "sheet.mergeAll"
    Assert-Match $r "已合并 2 张工作表，共 3 行数据*" "合并所有表：行数正确"

    $mOut = $mergeWb.Worksheets.Item("汇总")
    Assert-Equal "来源工作表" $mOut.Range("A1").Value2 "合并所有表：首列是来源工作表"

    # 关键断言：二月的"姓名"必须和一月落在同一列，而不是按列位置错位。
    # 不假设工作表顺序 —— Worksheets.Add() 不带参数是插在当前表【之前】的。
    $nameCol = 0; $qtyCol = 0
    for ($c = 1; $c -le 6; $c++) {
        if ($mOut.Cells(1, $c).Value2 -eq "姓名") { $nameCol = $c }
        if ($mOut.Cells(1, $c).Value2 -eq "销量") { $qtyCol = $c }
    }
    Assert-Equal $true ($nameCol -gt 0 -and $qtyCol -gt 0) "合并所有表：姓名和销量各占一列"

    $rowOf = @{}
    for ($r = 2; $r -le 4; $r++) { $rowOf[[string]$mOut.Cells($r, $nameCol).Value2] = $r }
    Assert-Equal "一月" $mOut.Cells($rowOf["张"], 1).Value2 "合并所有表：张来自一月"
    Assert-Equal "二月" $mOut.Cells($rowOf["王"], 1).Value2 "合并所有表：王来自二月"
    Assert-Equal 3 $mOut.Cells($rowOf["王"], $qtyCol).Value2 "合并所有表：列顺序不同的表也按标题名对齐"
    $mergeWb.Close($false)

    #==========================================================================
    Section "M4 合并文件夹"

    $dataDir = Join-Path $SandBox "数据源"
    New-Item -ItemType Directory -Path $dataDir | Out-Null
    $subDir = Join-Path $dataDir "子目录"
    New-Item -ItemType Directory -Path $subDir | Out-Null

    # 造三个夹具：两个在根目录、一个在子目录。
    # 注意 @(@(..),@(..)) 会被 PowerShell 拆平成一维，必须用一元逗号逐个追加。
    $specs = @()
    $specs += , @("甲.xlsx", $dataDir, "A")
    $specs += , @("乙.xlsx", $dataDir, "B")
    $specs += , @("丙.xlsx", $subDir,  "C")
    foreach ($spec in $specs) {
        $tmpWb = $xl.Workbooks.Add(-4167)
        $fixWs = $tmpWb.Worksheets.Item(1)
        $fixWs.Range("A1").Value2 = "编号"
        $fixWs.Range("B1").Value2 = "值"
        $fixWs.Range("A2").Value2 = $spec[2] + "1"
        $fixWs.Range("B2").Value2 = 10
        $fixWs.Range("A3").Value2 = $spec[2] + "2"
        $fixWs.Range("B3").Value2 = 20
        $tmpWb.SaveAs((Join-Path $spec[1] $spec[0]), 51)
        $tmpWb.Close($false)
    }

    & $Param "folder" $dataDir
    & $Param "recursive" "false"
    & $Param "sheetMode" "1"
    & $Param "headerRows" "1"
    $r = & $Run "merge.folder"
    Assert-Match $r "已合并 2 个文件*共 4 行数据*" "合并文件夹：不含子目录"

    $mergedWb = $xl.ActiveWorkbook
    $mOut2 = $mergedWb.Worksheets.Item("合并结果")
    Assert-Equal "来源文件" $mOut2.Range("A1").Value2 "合并文件夹：带来源文件列"
    Assert-Equal "来源工作表" $mOut2.Range("B1").Value2 "合并文件夹：带来源工作表列"
    Assert-Match $mOut2.Range("A2").Value2 "*.xlsx" "合并文件夹：来源文件名已填入"
    $mergedWb.Close($false)

    & $Param "recursive" "true"
    $r = & $Run "merge.folder"
    Assert-Match $r "已合并 3 个文件*共 6 行数据*" "合并文件夹：递归包含子目录"
    $xl.ActiveWorkbook.Close($false)

    # 坏文件不能中断整批
    $badFile = Join-Path $dataDir "损坏的.xlsx"
    [System.IO.File]::WriteAllText($badFile, "这不是一个 Excel 文件")
    & $Param "recursive" "false"
    $r = & $Run "merge.folder"
    Assert-Match $r "*个文件失败*" "合并文件夹：坏文件计入失败清单"
    Assert-Match $r "已合并 2 个文件*" "合并文件夹：坏文件不影响其余文件"
    $xl.ActiveWorkbook.Close($false)
    Remove-Item $badFile -Force

    #==========================================================================
    Section "M5 文件清单与批量重命名"

    $renameDir = Join-Path $SandBox "待改名"
    New-Item -ItemType Directory -Path $renameDir | Out-Null
    1..3 | ForEach-Object { [System.IO.File]::WriteAllText((Join-Path $renameDir "old$_.txt"), "x") }

    $listWb = $xl.Workbooks.Add(-4167)
    $null = $listWb.Activate()
    & $Param "folder" $renameDir
    & $Param "recursive" "false"
    $r = & $Run "file.list"
    Assert-Match $r "已列出 3 个文件*" "文件清单：数量正确"

    $wsList = $xl.ActiveSheet
    Assert-Equal "新文件名" $wsList.Range("F1").Value2 "文件清单：生成了新文件名列"

    # 只填两行，第三行留空 —— 留空的必须原样不动
    $wsList.Range("F2").Value2 = "新名A.txt"
    $wsList.Range("F3").Value2 = "新名B.txt"
    $r = & $Run "file.batchRename"
    Assert-Match $r "重命名完成：成功 2 个*" "批量重命名：只改填了新名的"
    Assert-Equal $true (Test-Path (Join-Path $renameDir "新名A.txt")) "批量重命名：新文件名存在"
    Assert-Equal 1 (@(Get-ChildItem $renameDir -Filter "old*.txt").Count) "批量重命名：留空的那个没被动"

    # 全表校验：只要有一处问题，一个文件都不能改
    $wsList2 = $xl.ActiveSheet
    $wsList2.Range("F2").Value2 = "含非法字符<>.txt"
    $wsList2.Range("F3").Value2 = "另一个.txt"
    $before = @(Get-ChildItem $renameDir | Select-Object -ExpandProperty Name) -join ","
    $r = & $Run "file.batchRename"
    Assert-Match $r "校验未通过，未改动任何文件*" "批量重命名：校验不过时整批拒绝"
    $after = @(Get-ChildItem $renameDir | Select-Object -ExpandProperty Name) -join ","
    Assert-Equal $before $after "批量重命名：校验失败后文件一个都没动"
    $listWb.Close($false)

    #==========================================================================
    Section "M5 导出工作表"

    $exportDir = Join-Path $SandBox "导出"
    New-Item -ItemType Directory -Path $exportDir | Out-Null

    $expWb = $xl.Workbooks.Add(-4167)
    $expWb.Worksheets.Item(1).Name = "报表甲"
    Set-Grid $expWb.Worksheets.Item(1) 1 1 @(@("列", "值"), @("x", 1))
    $e2 = $expWb.Worksheets.Add()
    $e2.Name = "报表乙"
    Set-Grid $e2 1 1 @(@("列", "值"), @("y", 2))
    $null = $expWb.Activate()

    & $Param "folder" $exportDir
    & $Param "exportFormat" "1"
    $r = & $Run "file.exportSheets"
    Assert-Match $r "已导出 2 张工作表*" "导出工作表：数量正确"
    Assert-Equal $true (Test-Path (Join-Path $exportDir "报表甲.xlsx")) "导出工作表：文件已生成"
    $expWb.Close($false)

    #==========================================================================
    Section "M3 按列表批量重命名工作表"

    $renWb = $xl.Workbooks.Add(-4167)
    $rs1 = $renWb.Worksheets.Item(1)
    $rs1.Name = "原名一"
    # 故意让新名和另一张表的当前名相同，验证"先改临时名"的两段式做法
    $rs1.Range("A1").Value2 = "原名二"
    $rs1.Range("A2").Value2 = "新名乙"
    # Worksheets.Add(Before, After, Count, Type) —— PowerShell 不支持 VBA 的
    # Name:=value 具名参数语法，只能按位置传，Before 用 Missing 占位
    $rs2 = $renWb.Worksheets.Add([System.Type]::Missing, $rs1)
    $rs2.Name = "原名二"

    $null = $renWb.Activate()
    Select-On $rs1 "A1:A2"
    $r = & $Run "sheet.batchRename"
    Assert-Match $r "已重命名 2 张工作表。" "批量重命名工作表：数量正确"
    Assert-Equal "原名二" $renWb.Worksheets.Item(1).Name "批量重命名工作表：新名与他表旧名相同也不会失败"
    Assert-Equal "新名乙" $renWb.Worksheets.Item(2).Name "批量重命名工作表：第二张也改了"
    $renWb.Close($false)

    #==========================================================================
    Section "M5 批量插入图片"

    $imgDir = Join-Path $SandBox "图片"
    New-Item -ItemType Directory -Path $imgDir | Out-Null

    # 用 System.Drawing 现造两张小图，避免依赖仓库里的二进制夹具
    Add-Type -AssemblyName System.Drawing
    foreach ($imgName in @("图A", "图B")) {
        $bmp = New-Object System.Drawing.Bitmap 20, 20
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::CornflowerBlue)
        $g.Dispose()
        $bmp.Save((Join-Path $imgDir ($imgName + ".png")), [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    }

    $imgWb = $xl.Workbooks.Add(-4167)
    $iws = $imgWb.Worksheets.Item(1)
    $iws.Range("A1").Value2 = "图A"
    $iws.Range("A2").Value2 = "图B"
    $iws.Range("A3").Value2 = "不存在的图"
    $null = $imgWb.Activate()

    & $Param "folder" $imgDir
    Select-On $iws "A1:A3"
    $r = & $Run "file.insertImages"
    Assert-Match $r "已插入 2 张图片*" "批量插图：按单元格内容匹配到图片"
    Assert-Match $r "*1 个名称没有找到对应图片*" "批量插图：找不到的如实报出来，不静默跳过"
    Assert-Equal 2 $iws.Shapes.Count "批量插图：图片对象已插入"
    $imgWb.Close($false)

    #==========================================================================
    Section "环境还原"
    Assert-Equal $true $xl.ScreenUpdating "ScreenUpdating 已还原"
    Assert-Equal $true $xl.EnableEvents   "EnableEvents 已还原"
}
catch {
    $script:fail++
    Write-Host ""
    Write-Host "测试过程异常（$script:section）：$($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  出错行：$($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
    Write-Host "  行号  ：$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
}
finally {
    if ($xl) {
        Close-ExcelInstance $xl
    }
    Start-Sleep -Milliseconds 500

    if (Test-Path $SandBox) { Remove-Item $SandBox -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
Write-Host ("通过 {0} / 失败 {1}" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -gt 0) { exit 1 } else { exit 0 }
