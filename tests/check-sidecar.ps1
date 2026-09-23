<#
.SYNOPSIS
    sidecar 伴生进程骨架的回归：绑回环、令牌校验、Origin 校验、生命周期。

.DESCRIPTION
    为什么需要这个：

    sidecar 在 127.0.0.1 上开了一个端口，而【本机所有浏览器进程都能访问它】。
    用户打开的任意网页里的 JS 都能向它发请求。所以这里每一条断言守的
    都不是"功能对不对"，而是"会不会被一个远端网页利用"。

    最危险的失败模式是【放行本该拒绝的请求】——这种错不会让任何功能报错，
    表现完全正常，只有被人利用的时候才知道。

    【本脚本会真的启动 sidecar 进程并监听回环端口】，但：
      · 只绑 127.0.0.1，不对网络暴露
      · 用的是临时目录里的临时配置和临时令牌
      · 每条用例跑完都把进程收掉，末尾还有残留检查

.EXAMPLE
    timeout 600 powershell -ExecutionPolicy Bypass -File tests\check-sidecar.ps1
#>
[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SrcFile  = Join-Path $RepoRoot "sidecar\ExcelToolboxSidecar.cs"
$Exe      = Join-Path $RepoRoot "dist\ExcelToolboxSidecar.exe"
if (-not (Test-Path $Exe)) { throw "找不到 $Exe。请先运行 build\build-sidecar.ps1。" }

$script:pass = 0
$script:fail = 0

function Section($n) { Write-Host ""; Write-Host "== $n ==" -ForegroundColor Cyan }
function Assert-Equal($expected, $actual, [string]$what) {
    if ("$expected" -eq "$actual") { $script:pass++; Write-Host "  PASS  $what" -ForegroundColor Green }
    else {
        $script:fail++
        Write-Host "  FAIL  $what" -ForegroundColor Red
        Write-Host "        期望 [$expected] 实际 [$actual]" -ForegroundColor Red
    }
}
function Assert-True($c, [string]$w) { Assert-Equal $true ([bool]$c) $w }

#------------------------------------------------------------------------------
# 临时夹具
#------------------------------------------------------------------------------
$Stage = Join-Path ([IO.Path]::GetTempPath()) ("tbsidecar_" + [guid]::NewGuid().ToString("N"))
$null = New-Item -ItemType Directory -Path $Stage -Force

$Token = -join (1..64 | ForEach-Object { "0123456789abcdef"[(Get-Random -Maximum 16)] })
$GoodOrigin = "https://gw.test.local:8443"
$BadOrigin  = "https://evil.example.com"

function New-Config([string]$name, [string]$token, [int]$port, [string[]]$origins) {
    $p = Join-Path $Stage $name
    $obj = @{ port = $port }
    if ($null -ne $token) { $obj.token = $token }
    if ($origins) { $obj.allowedOrigins = $origins }
    ($obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $p -Encoding UTF8
    return $p
}

$script:Started = @()

function Start-Sidecar([string]$configPath, [int]$parentPid, [int]$portOverride, [string[]]$Extra) {
    $stdout = Join-Path $Stage ("out_" + [guid]::NewGuid().ToString("N") + ".txt")
    $stderr = $stdout -replace "^out_", "err_"
    $stderr = Join-Path $Stage ("err_" + (Split-Path $stdout -Leaf))

    $a = @("--config", $configPath)
    if ($parentPid -gt 0)   { $a += @("--parent-pid", "$parentPid") }
    if ($portOverride -gt 0) { $a += @("--port", "$portOverride") }
    if ($Extra) { $a += $Extra }

    $p = Start-Process -FilePath $Exe -ArgumentList $a -NoNewWindow -PassThru `
                       -RedirectStandardOutput $stdout -RedirectStandardError $stderr

    # 【必须在进程还活着的时候摸一下 .Handle】。这是 Start-Process -PassThru
    # 的老毛病：不提前把句柄缓存下来，进程一退 ExitCode 就永远读成空，
    # 断言会显示成"没有退出码"，看起来像是程序没设退出码——其实设了。
    try { $null = $p.Handle } catch { }

    $script:Started += $p

    # 等它打印「已就绪：127.0.0.1:<端口>」，从中拿到真正绑上的端口
    $port = 0
    for ($i = 0; $i -lt 60; $i++) {
        if ($p.HasExited -and -not (Test-Path $stdout)) { break }
        if (Test-Path $stdout) {
            $t = Get-Content -LiteralPath $stdout -Raw -ErrorAction SilentlyContinue
            if ($t -and $t -match '127\.0\.0\.1:(\d+)') { $port = [int]$Matches[1]; break }
        }
        if ($p.HasExited) { break }
        Start-Sleep -Milliseconds 200
    }

    return [pscustomobject]@{
        Process = $p; Port = $port; StdOut = $stdout; StdErr = $stderr
    }
}

# 【必须先 WaitForExit 再读 ExitCode】。Start-Process -PassThru 拿到的对象
# 在进程结束前 ExitCode 是空的，直接读会得到空字符串，断言看起来像"没有退出码"。
function Get-ExitCode($s) {
    try { $s.Process.WaitForExit(5000) | Out-Null } catch { }
    try { return $s.Process.ExitCode } catch { return $null }
}

function Stop-Sidecar($s) {
    if ($null -eq $s -or $null -eq $s.Process) { return }
    try { if (-not $s.Process.HasExited) { $s.Process.Kill() } } catch { }
    try { $s.Process.WaitForExit(5000) | Out-Null } catch { }
}

#------------------------------------------------------------------------------
# 发请求。【故意不用 Invoke-WebRequest】——它对 4xx 会抛异常，
# 而这个套件里"拿到 401/403"恰恰是正常预期，用 HttpWebRequest 更直接。
#------------------------------------------------------------------------------
function Invoke-Sidecar {
    param(
        [int]$Port, [string]$Path = "/health", [string]$Method = "GET",
        [string]$Token, [string]$Origin, [string]$Body
    )
    $url = "http://127.0.0.1:$Port$Path"
    try {
        $req = [Net.HttpWebRequest]::Create($url)
        $req.Method = $Method
        $req.Timeout = 8000
        $req.Proxy = $null            # 别让系统代理把回环请求绕出去
        if ($Token)  { $req.Headers.Add("X-Toolbox-Token", $Token) }
        if ($Origin) { $req.Headers.Add("Origin", $Origin) }

        if ($PSBoundParameters.ContainsKey('Body')) {
            $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
            $req.ContentType = "application/json"
            $req.ContentLength = $bytes.Length
            $rs = $req.GetRequestStream()
            $rs.Write($bytes, 0, $bytes.Length)
            $rs.Close()
        }

        $resp = $req.GetResponse()
        $code = [int]$resp.StatusCode
        $acao = $resp.Headers["Access-Control-Allow-Origin"]
        $body = ""
        $st = $resp.GetResponseStream()
        if ($st) { $body = (New-Object IO.StreamReader($st)).ReadToEnd() }
        $resp.Close()
        return [pscustomobject]@{ Status = $code; Body = $body; Acao = $acao; Failed = $false }
    }
    catch [Net.WebException] {
        $we = $_.Exception
        if ($we.Response) {
            $code = [int]$we.Response.StatusCode
            $acao = $we.Response.Headers["Access-Control-Allow-Origin"]
            $body = ""
            try { $body = (New-Object IO.StreamReader($we.Response.GetResponseStream())).ReadToEnd() } catch { }
            return [pscustomobject]@{ Status = $code; Body = $body; Acao = $acao; Failed = $false }
        }
        return [pscustomobject]@{ Status = 0; Body = $we.Message; Acao = $null; Failed = $true }
    }
}


#------------------------------------------------------------------------------
# 裸 socket 发一段原始 HTTP。
#
# 【不能用 HttpWebRequest 测这些】：它会自己把头部规范化、自己算
# Content-Length，畸形请求根本发不出去——那样测的是 .NET 的客户端，
# 不是我们的服务端。
#------------------------------------------------------------------------------
function Invoke-RawHttp([int]$Port, [string]$Raw) {
    try {
        $c = New-Object Net.Sockets.TcpClient
        $c.Connect("127.0.0.1", $Port)
        $c.ReceiveTimeout = 5000
        $st = $c.GetStream()
        $bytes = [Text.Encoding]::ASCII.GetBytes($Raw)
        $st.Write($bytes, 0, $bytes.Length)
        $st.Flush()

        $sr = New-Object IO.StreamReader($st)
        $text = $sr.ReadToEnd()
        $c.Close()
        return $text
    } catch {
        return "EXCEPTION: $($_.Exception.Message)"
    }
}

function Get-StatusCode([string]$rawResponse) {
    if ($rawResponse -match '^HTTP/1\.1 (\d{3})') { return [int]$Matches[1] }
    return 0
}

try {
    #==========================================================================
    Section "源码级约束（这些错编译不会报，跑起来也看不出来）"
    #==========================================================================
    # 【必须先把注释剥掉再断言】。源码里正是用大段注释解释"为什么不用
    # HttpListener、为什么不绑 IPAddress.Any"，直接全文匹配会把这些说明
    # 当成违例——断言红了，可代码其实是对的。要查的是代码，不是注释。
    $code = (Get-Content -LiteralPath $SrcFile) |
            Where-Object { $_.TrimStart() -notmatch '^//' }
    $code = $code -join "`n"

    # HttpListener 要求 URL 保留，非管理员拿不到。开发机上开发者通常是管理员，
    # 于是"在我这儿是好的"，铺到用户机才发现起不来。
    Assert-True ($code -notmatch 'HttpListener') "不用 HttpListener（非管理员用户起不来）"

    # 绑 IPAddress.Any 等于把端口暴露给整个内网，还会触发防火墙弹窗
    Assert-True ($code -notmatch 'IPAddress\.Any') "不绑 IPAddress.Any"
    Assert-True ($code -match 'IPAddress\.Loopback')  "确实绑的是回环地址"

    # 通用执行口子是这套设计里最不能开的东西
    Assert-True ($code -notmatch 'Process\.Start') "源码里没有任意进程启动的口子"

    # 剥注释这件事本身也得守住：剥完不能把代码也剥没了
    Assert-True ($code.Length -gt 2000) "剥掉注释后仍有实质代码（剥注释没把代码吃掉）"

    #==========================================================================
    Section "拒绝启动：配置不合格就不该把端口开起来"
    #==========================================================================
    $noTokenCfg = New-Config "no-token.json" $null 8971 @($GoodOrigin)
    $s = Start-Sidecar $noTokenCfg 0 0
    Start-Sleep -Milliseconds 500
    Assert-True ($s.Process.HasExited) "配置里没有 token：进程退出，不监听"
    Assert-Equal 2 (Get-ExitCode $s) "没有 token 时退出码为 2"
    Stop-Sidecar $s

    $shortCfg = New-Config "short-token.json" "abc123" 8972 @($GoodOrigin)
    $s = Start-Sidecar $shortCfg 0 0
    Start-Sleep -Milliseconds 500
    Assert-True ($s.Process.HasExited) "token 太短：进程退出，不监听"
    Assert-Equal 2 (Get-ExitCode $s) "token 太短时退出码为 2"
    Stop-Sidecar $s

    $missingCfg = Join-Path $Stage "not-there.json"
    $s = Start-Sidecar $missingCfg 0 0
    Start-Sleep -Milliseconds 500
    Assert-True ($s.Process.HasExited) "配置文件不存在：进程退出，不监听"
    Stop-Sidecar $s

    #==========================================================================
    Section "正常启动与健康检查"
    #==========================================================================
    $cfg = New-Config "good.json" $Token 8981 @($GoodOrigin)
    $sc = Start-Sidecar $cfg 0 0
    Assert-True ($sc.Port -gt 0) "sidecar 启动并报出了监听端口"

    $r = Invoke-Sidecar -Port $sc.Port -Token $Token
    Assert-Equal 200 $r.Status "带正确令牌：/health 返回 200"
    Assert-True ($r.Body -match '"ok"\s*:\s*true') "/health 的响应体里 ok=true"

    #==========================================================================
    Section "令牌校验（这里放行一条，整套防护就没了）"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port
    Assert-Equal 401 $r.Status "不带令牌：401"

    $r = Invoke-Sidecar -Port $sc.Port -Token "wrong-token-but-long-enough-0123456789abcdef"
    Assert-Equal 401 $r.Status "令牌错误：401"

    # 【守住"前缀比较"这类写法】：用 StartsWith / 截断比较都会让这条变绿
    $r = Invoke-Sidecar -Port $sc.Port -Token $Token.Substring(0, 32)
    Assert-Equal 401 $r.Status "令牌是正确令牌的前缀：401（不能只比前缀）"

    $r = Invoke-Sidecar -Port $sc.Port -Token ($Token + "x")
    Assert-Equal 401 $r.Status "令牌多一个字符：401"

    $r = Invoke-Sidecar -Port $sc.Port -Token $Token.ToUpper()
    Assert-Equal 401 $r.Status "令牌大小写不同：401（必须严格相等）"

    # 【查询串传令牌必须拒绝】。?token=xxx 属于简单请求、不触发预检，
    # 一旦接受，"自定义头强制预检"这道防线就等于自己拆了。
    $r = Invoke-Sidecar -Port $sc.Port -Path "/health?token=$Token"
    Assert-True ($r.Status -ne 200) "查询串里带令牌：不放行"
    Assert-Equal 400 $r.Status "查询串里带令牌：400 并提示令牌要放在头里"

    #==========================================================================
    Section "Origin 校验（第二道，不是唯一一道）"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port -Token $Token -Origin $BadOrigin
    Assert-Equal 403 $r.Status "非白名单 Origin：403（即使令牌是对的）"

    $r = Invoke-Sidecar -Port $sc.Port -Token $Token -Origin $GoodOrigin
    Assert-Equal 200 $r.Status "白名单 Origin：放行"
    Assert-Equal $GoodOrigin $r.Acao "回显的是那个具体 Origin"
    Assert-True ($r.Acao -ne "*") "绝不回 Access-Control-Allow-Origin: *"

    # 子串攻击：https://gw.test.local:8443.evil.com 不能被当成白名单成员
    $r = Invoke-Sidecar -Port $sc.Port -Token $Token -Origin ($GoodOrigin + ".evil.com")
    Assert-Equal 403 $r.Status "Origin 是白名单项的前缀延长：403"

    #==========================================================================
    Section "预检"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port -Method "OPTIONS" -Origin $GoodOrigin
    Assert-Equal 204 $r.Status "白名单 Origin 的预检：204"
    Assert-Equal $GoodOrigin $r.Acao "预检回显那个 Origin"

    $r = Invoke-Sidecar -Port $sc.Port -Method "OPTIONS" -Origin $BadOrigin
    Assert-Equal 403 $r.Status "非白名单 Origin 的预检：403"
    Assert-True ([string]::IsNullOrEmpty($r.Acao)) "被拒的预检不回 CORS 头"

    #==========================================================================
    Section "路由是白名单式的"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "GET" -Token $Token
    Assert-Equal 404 $r.Status "没实现的路径：404"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/eval" -Method "GET" -Token $Token
    Assert-Equal 404 $r.Status "没有通用执行口子"

    #==========================================================================
    Section "能力接口：白名单式，不接受代码"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port -Path "/queries" -Token $Token
    Assert-Equal 200 $r.Status "/queries 带令牌可用"
    Assert-True ($r.Body -match '"ok"') "/queries 返回结构化结果"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/queries"
    Assert-Equal 401 $r.Status "/queries 不带令牌：401"

    # 刷新必须指名道姓。不给名字就拒绝——这是"只暴露具体动作"的体现。
    $r = Invoke-Sidecar -Port $sc.Port -Path "/refresh-query" -Method "POST" -Token $Token -Body '{}'
    Assert-Equal 400 $r.Status "刷新不给查询名：400"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/refresh-query" -Method "POST" -Body '{"name":"x"}'
    Assert-Equal 401 $r.Status "刷新不带令牌：401"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/refresh-query" -Method "POST" -Token $Token -Body '{"name":"NoSuchQuery_xyz"}'
    Assert-Equal 200 $r.Status "刷新请求本身被受理"
    # Excel 没开就该如实说 Excel 没开；开着但没这个查询就该说找不到。
    # 【两种都不能报成成功】——报成功的话用户点了刷新什么也没发生，却以为好了。
    Assert-True ($r.Body -match 'excel_not_running|query_not_found|no_workbook') `
                "不存在的查询不会被报成刷新成功"

    #==========================================================================
    Section "调用宏：只认 AI_ 前缀，不接受代码"
    #==========================================================================
    $r = Invoke-Sidecar -Port $sc.Port -Path "/macros" -Token $Token
    Assert-Equal 200 $r.Status "/macros 带令牌可用"
    Assert-True ($r.Body -match '"ok"') "/macros 返回结构化结果"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/macros"
    Assert-Equal 401 $r.Status "/macros 不带令牌：401"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{}'
    Assert-Equal 400 $r.Status "调用宏不给名字：400"

    # 【这是这条接口最核心的一道防线】：没有 AI_ 前缀的宏名，
    # 请求要在真的去调用 Excel 之前就被拒绝——不能指望"Excel 没开"
    # 之类的下游失败顺便挡住它，那样只要哪天 Excel 恰好开着，
    # 这道门就形同虚设。
    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"DeleteAllSheets"}'
    Assert-Equal 400 $r.Status "宏名没有 AI_ 前缀：400（不去尝试调用）"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"AI_"}'
    Assert-Equal 400 $r.Status "AI_ 后面空着：400"

    # 大小写、前缀藏在中间——都不能被当成合法前缀
    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"ai_lowercase"}'
    Assert-Equal 400 $r.Status "前缀大小写不对：400"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"NotAI_Foo"}'
    Assert-Equal 400 $r.Status "AI_ 不在开头：400"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Body '{"name":"AI_Test"}'
    Assert-Equal 401 $r.Status "调用宏不带令牌：401"

    # 【参数只能是原子值】。传对象/数组进来一律拒绝——这条接口不接受
    # 任何"代码形状"的东西，只收字符串/数字/布尔。
    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"AI_Test","args":[{"x":1}]}'
    Assert-Equal 400 $r.Status "参数里混进对象：400"

    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"AI_Test","args":"not-an-array"}'
    Assert-Equal 400 $r.Status "参数不是数组：400"

    # 名字合法、参数合法，但没有 Excel——必须如实说没开，不能报成功
    $r = Invoke-Sidecar -Port $sc.Port -Path "/run-macro" -Method "POST" -Token $Token -Body '{"name":"AI_NoSuchMacro_xyz","args":["a",1,true]}'
    Assert-Equal 200 $r.Status "合法请求本身被受理"
    Assert-True ($r.Body -match 'excel_not_running|no_workbook|macro_failed') `
                "不存在/调不到的宏不会被报成调用成功"

    # 【这里不要停掉 $sc】。下面"只绑回环"那节要连它的端口，
    # 进程没了的话连接当然失败，那条断言就会【因为错误的原因变绿】——
    # 它本该证明的是"绑了回环所以外网连不上"，而不是"服务根本没在跑"。

    #==========================================================================
    Section "单实例：开几个 Excel 也只该有一个 sidecar"
    #==========================================================================
    $cfgSi = New-Config "single.json" $Token 8961 @($GoodOrigin)
    $key = "TbSidecarTest_" + [guid]::NewGuid().ToString("N").Substring(0, 8)
    $first = Start-Sidecar $cfgSi 0 0 @("--single-instance", $key)
    Assert-True ($first.Port -gt 0) "第一个实例正常启动"

    $second = Start-Sidecar $cfgSi 0 0 @("--single-instance", $key)
    Assert-True ($second.Process.HasExited) "第二个实例没有占住端口"
    Assert-Equal 0 (Get-ExitCode $second) "第二个实例安静退出（退出码 0，不是报错）"

    $r = Invoke-Sidecar -Port $first.Port -Token $Token
    Assert-Equal 200 $r.Status "第一个实例仍然正常服务"
    Stop-Sidecar $first
    Stop-Sidecar $second

    #==========================================================================
    Section "跟随宿主：Excel 全退了就自己退"
    #==========================================================================
    # 用一个不存在的进程名 + 1 秒宽限，等价于"宿主从来没出现过"。
    # 这条守的是"用户关掉所有 Excel 之后 sidecar 不该赖着不走"。
    $cfgW = New-Config "watch.json" $Token 8966 @($GoodOrigin)
    $w = Start-Sidecar $cfgW 0 0 @("--watch-process", "NoSuchHost_xyz", "--watch-grace-seconds", "1")
    Assert-True ($w.Port -gt 0) "带 --watch-process 时正常启动"

    $gone = $false
    for ($i = 0; $i -lt 40; $i++) {
        if ($w.Process.HasExited) { $gone = $true; break }
        Start-Sleep -Milliseconds 250
    }
    Assert-True $gone "宿主进程一个都不剩时，sidecar 自己退出"
    Stop-Sidecar $w

    #==========================================================================
    Section "HTTP 解析的边界（畸形请求不能被放行）"
    #==========================================================================
    # 这一节全部用裸 socket，因为要发的正是"正常客户端发不出来"的东西。

    # 【头部名大小写不敏感】。HTTP 规范要求如此。
    # 真实客户端（fetch / XHR / 代理）完全可能把头名小写化，
    # 区分大小写的话表现是【带着正确令牌却一直 401】。
    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nx-toolbox-token: $Token`r`nConnection: close`r`n`r`n"
    Assert-Equal 200 (Get-StatusCode $raw) "头部名小写也认（HTTP 头不区分大小写）"

    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nX-TOOLBOX-TOKEN: $Token`r`nConnection: close`r`n`r`n"
    Assert-Equal 200 (Get-StatusCode $raw) "头部名全大写也认"

    # 【重复令牌头只取第一个】。取最后一个的话，攻击者可以在一个
    # 合法请求后面追加自己的头来覆盖前面的值。
    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nX-Toolbox-Token: wrong-token-here-000000000000000000`r`nX-Toolbox-Token: $Token`r`nConnection: close`r`n`r`n"
    Assert-Equal 401 (Get-StatusCode $raw) "重复令牌头：以第一个为准，后面的覆盖不了"

    # 【Content-Length 不是数字时必须拒绝】。当成 0 的话请求体被丢掉，
    # 但请求仍按"没有正文"继续处理——我们和客户端对同一个请求的理解
    # 就不一致了，这正是 HTTP 走私类问题的温床。
    #
    # 【这几条必须打在 /health 上，不能打在 /refresh-query 上】。
    # 打在 /refresh-query 上是测不出来的：正文被丢掉之后它会因为
    # "没给查询名"而回 400，和"正确地拒绝了畸形请求"同样是 400——
    # 两种原因分不开，断言就会【因为错误的原因变绿】。
    # /health 本来就不需要正文，所以只有真的拒绝了才会是 400。
    # （变异测试正是这么发现这条断言不够格的。）
    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nX-Toolbox-Token: $Token`r`nContent-Length: abc`r`nConnection: close`r`n`r`n"
    Assert-Equal 400 (Get-StatusCode $raw) "Content-Length 非数字：400（不当成 0）"

    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nX-Toolbox-Token: $Token`r`nContent-Length: -5`r`nConnection: close`r`n`r`n"
    Assert-Equal 400 (Get-StatusCode $raw) "Content-Length 为负：400"

    # 声明一个超大正文：必须拒掉，不能真去分配那么多内存
    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`nX-Toolbox-Token: $Token`r`nContent-Length: 99999999`r`nConnection: close`r`n`r`n"
    Assert-Equal 400 (Get-StatusCode $raw) "声明超大正文：400（不分配那么多内存）"

    # 正常的 POST 仍要能走通——上面几条不能把合法请求也拦了
    $raw = Invoke-RawHttp $sc.Port "POST /refresh-query HTTP/1.1`r`nHost: 127.0.0.1`r`nX-Toolbox-Token: $Token`r`nContent-Length: 12`r`nConnection: close`r`n`r`n{`"name`":`"x`"}"
    Assert-Equal 200 (Get-StatusCode $raw) "Content-Length 正确的 POST 照常受理"

    # 超长头部不能把服务打挂
    $huge = "X-Junk: " + ("A" * 20000)
    $raw = Invoke-RawHttp $sc.Port "GET /health HTTP/1.1`r`nHost: 127.0.0.1`r`n$huge`r`nX-Toolbox-Token: $Token`r`nConnection: close`r`n`r`n"
    Assert-True ((Get-StatusCode $raw) -ne 200) "超长头部：不放行"

    # 【发完这些畸形请求，服务必须还活着】。
    # 一个坏请求把整个 sidecar 打挂的话，等于任何网页都能让它拒绝服务。
    $r = Invoke-Sidecar -Port $sc.Port -Token $Token
    Assert-Equal 200 $r.Status "一连串畸形请求之后，服务仍然正常"

    #==========================================================================
    Section "只绑回环：从本机的非回环地址连不上"
    #==========================================================================
    $lanIps = @(
        [Net.Dns]::GetHostAddresses([Net.Dns]::GetHostName()) |
        Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.ToString() -ne '127.0.0.1' }
    )
    if ($lanIps.Count -eq 0) {
        Write-Host "  跳过  这台机器没有非回环 IPv4 地址，这条测不了" -ForegroundColor DarkYellow
    }
    else {
        $ip = $lanIps[0].ToString()
        $reachable = $false
        try {
            $c = New-Object Net.Sockets.TcpClient
            $ar = $c.BeginConnect($ip, $sc.Port, $null, $null)
            if ($ar.AsyncWaitHandle.WaitOne(2000)) {
                try { $c.EndConnect($ar); $reachable = $c.Connected } catch { $reachable = $false }
            }
            $c.Close()
        } catch { $reachable = $false }
        Assert-Equal $false $reachable "从 $ip 连不上（没有绑到全部网卡）"
    }

    # 【先确认它在这一刻确实还活着】，否则上面那条"连不上"可能只是
    # 因为进程早退了，而不是因为绑的是回环。
    $r = Invoke-Sidecar -Port $sc.Port -Token $Token
    Assert-Equal 200 $r.Status "同一时刻从 127.0.0.1 连得上（证明上一条不是因为进程没了）"

    Stop-Sidecar $sc

    #==========================================================================
    Section "端口被占时顺延"
    #==========================================================================
    $busyPort = 8991
    $blocker = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, $busyPort)
    $blocker.Start()
    try {
        $cfg2 = New-Config "busy.json" $Token $busyPort @($GoodOrigin)
        $sc2 = Start-Sidecar $cfg2 0 0
        Assert-True ($sc2.Port -gt 0) "端口被占时仍能起来"
        Assert-True ($sc2.Port -ne $busyPort) "顺延到了别的端口（不是被占的那个）"
        $r = Invoke-Sidecar -Port $sc2.Port -Token $Token
        Assert-Equal 200 $r.Status "顺延后的端口上服务正常"
        Stop-Sidecar $sc2
    }
    finally { $blocker.Stop() }

    #==========================================================================
    Section "生命周期：父进程没了就自己退"
    #==========================================================================
    # 这是"关 Excel 就没了、用户完全无感"那个承诺的落点。
    # 少了它，sidecar 就成了用户看不见也关不掉的常驻进程。
    $parent = Start-Process -FilePath "powershell.exe" `
                -ArgumentList @("-NoProfile", "-Command", "Start-Sleep -Seconds 120") `
                -PassThru -WindowStyle Hidden
    try {
        $cfg3 = New-Config "child.json" $Token 8995 @($GoodOrigin)
        $sc3 = Start-Sidecar $cfg3 $parent.Id 0
        Assert-True ($sc3.Port -gt 0) "带 --parent-pid 时正常启动"

        $parent.Kill()
        $parent.WaitForExit(5000) | Out-Null

        $gone = $false
        for ($i = 0; $i -lt 50; $i++) {
            if ($sc3.Process.HasExited) { $gone = $true; break }
            Start-Sleep -Milliseconds 200
        }
        Assert-True $gone "父进程退出后，sidecar 自己也退出了"
        Stop-Sidecar $sc3
    }
    finally {
        try { if (-not $parent.HasExited) { $parent.Kill() } } catch { }
    }

    #==========================================================================
    Section "残留检查"
    #==========================================================================
    foreach ($p in $script:Started) {
        try { if (-not $p.HasExited) { $p.Kill(); $p.WaitForExit(3000) | Out-Null } } catch { }
    }
    $alive = @($script:Started | Where-Object { -not $_.HasExited })
    Assert-Equal 0 $alive.Count "没有残留的 sidecar 进程"
}
finally {
    foreach ($p in $script:Started) {
        try { if (-not $p.HasExited) { $p.Kill() } } catch { }
    }
    Start-Sleep -Milliseconds 300
    Remove-Item -LiteralPath $Stage -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "通过 $script:pass / 失败 $script:fail" -ForegroundColor $(if ($script:fail -eq 0) { "Green" } else { "Red" })
if ($script:fail -eq 0) { exit 0 } else { exit 1 }
