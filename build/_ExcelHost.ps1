<#
.SYNOPSIS
    创建一个【确认是真 Microsoft Excel】的 COM 实例。

.DESCRIPTION
    为什么需要这个：装了 WPS 之后，它会把 Excel 的 COM 注册整个接管——
    HKCU 和 HKCR 两层的 CLSID{00024500-...}\LocalServer32 都被改成了 et.exe，
    连 Excel.Application.16 这个带版本号的 ProgID 也一样。

    更麻烦的是 WPS 会【自称 Microsoft Excel】：Application.Name 返回的就是
    "Microsoft Excel"，只有 Application.Path 和 Version(12.0) 能露馅。

    于是构建和测试脚本会在毫不知情的情况下驱动 WPS 跑，
    产出的 .xlam 和测试结论全都是错的靶子。这种"跑错宿主还全绿"的情况
    比直接报错危险得多，所以这里必须显式校验并【拿错就报错】。

.EXAMPLE
    . "$PSScriptRoot\_ExcelHost.ps1"
    $xl = New-RealExcel
#>

# 判断一个 Excel COM 实例到底是不是 WPS。
#
# 路径匹配必须用短名也能命中的前缀：实测 WPS 的 Application.Path 返回的是
# 8.3 短路径 D:\PROGRA~3\WPSOFF~1\...，"WPSOFFICE" 根本匹配不上。
# 版本号是第二道判据：WPS 报 12.0，而本工具箱最低支持的 Excel 2010 是 14.0，
# 所以真 Excel 不可能报 12.0。
function Test-IsWpsHost {
    param($App)

    $path = ""; $name = ""; $ver = ""
    try { $path = $App.Path } catch {}
    try { $name = $App.Name } catch {}
    try { $ver  = $App.Version } catch {}

    if ($path -match 'WPSOFF|KINGSO|WPS Office') { return $true }
    if ($name -match 'WPS') { return $true }

    $num = 0.0
    if ($ver) { [void][double]::TryParse(($ver -replace '[^0-9.].*$',''), [ref]$num) }
    if ($num -gt 0 -and $num -lt 14) { return $true }

    return $false
}

function New-RealExcel {
    [CmdletBinding()]
    param()

    $xl = New-Object -ComObject Excel.Application

    $path = ""
    $name = ""
    $ver  = ""
    try { $path = $xl.Path } catch {}
    try { $name = $xl.Name } catch {}
    try { $ver  = $xl.Version } catch {}

    if (Test-IsWpsHost $xl) {
        try { $xl.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}

        throw @"
COM 自动化拿到的不是 Microsoft Excel，而是 WPS：
    Name    = $name
    Version = $ver
    Path    = $path

WPS 安装后会接管 Excel 的 COM 注册（连 Excel.Application.16 也会被劫持），
并且刻意自称 "Microsoft Excel"，所以不校验就会在错误的宿主上构建和测试。

恢复方法：WPS → 设置/配置工具 → 兼容设置 → 取消 Office 文件关联，
或点"恢复 Office 默认设置"，然后重新运行本脚本。

（WPS 的兼容性请用 tests\probe-wps.ps1 单独验证，不要让它占着 COM 注册。）
"@
    }

    return $xl
}
