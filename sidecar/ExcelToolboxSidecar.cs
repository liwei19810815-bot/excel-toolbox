// Excel 工具箱 sidecar 伴生进程 —— 骨架
//
// 它是干什么的：Office.js 任务窗格跑在浏览器沙箱里，有几件事根本做不到
// （Power Query 刷新、数据模型、调用工作簿里的宏）。sidecar 是本机上的一个
// 小进程，由 Excel 打开工具箱时用 VBA Shell() 拉起，关 Excel 就跟着退出。
//
// ============================================================================
// 【威胁模型：真正的入口是浏览器，不是本机恶意软件】
// ============================================================================
// 本机恶意软件已经有用户的全部权限，多一个 sidecar 不改变什么。
// 真正的风险是：127.0.0.1 对本机所有浏览器进程都可达，用户打开的任意网页
// 里的 JS 都能向 sidecar 发请求。
//
// 常见误解是"有 CORS 挡着"。【CORS 挡的是读响应，不是发请求】——
// 一个 text/plain 的简单请求会真的发出去、真的被执行，浏览器只是不让
// 那个页面读返回值。在一个能操作 Excel 的端点上，"能触发但读不到结果"
// 依然是灾难。
//
// 所以防护的支点是这两条，缺一不可：
//
//   1. 【必须带自定义请求头 X-Toolbox-Token】
//      自定义头会强制浏览器先发 OPTIONS 预检。远端网页【根本没法】
//      把这个头发出来，除非预检先过——而预检只对白名单 Origin 放行。
//      这就是为什么【绝不接受从查询串里传令牌】：?token=xxx 属于简单请求，
//      不触发预检，等于把上面这道门自己拆了。下面有专门的代码拦这件事。
//
//   2. 【256 位随机令牌】，安装时生成，远端网页猜不到。
//
//   Origin 校验是第二道，不是唯一一道。
//
// ============================================================================
// 【为什么用 TcpListener 而不是 HttpListener】
// ============================================================================
// HttpListener 在 Windows 上要求 URL 保留（netsh http add urlacl），
// 非管理员用户拿不到，会直接抛 "拒绝访问"。
// 而用户机上的人【基本都不是管理员】。
//
// 要命的是：开发机上开发者通常是管理员，HttpListener 一跑就通，
// 等铺到用户机才发现起不来——又是一个"在我这儿是好的"。
// TcpListener 绑回环不需要任何权限，从根上绕开这件事。
//
// tests\check-sidecar.ps1 里有一条静态断言，专门守着本文件不许出现
// HttpListener，别改回去。
//
// ============================================================================
// 编译：build\build-sidecar.ps1（用 Windows 自带的 csc.exe，不装任何东西）
// 【只能用 C# 5 的语法】——系统自带的 csc 就到这个版本。
// 不要用字符串内插 $""、?. 、nameof、表达式体成员，编译会直接失败。
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace ExcelToolbox.Sidecar
{
    internal static class Program
    {
        public const string Version = "0.1.0";

        // 令牌走这个头。改名字要同步改任务窗格那边。
        public const string TokenHeader = "X-Toolbox-Token";

        // 令牌长度下限。安装器发的是 64 个十六进制字符（256 位）。
        // 【短令牌等于没有令牌】，所以这里宁可拒绝启动也不将就。
        private const int MinTokenLength = 32;

        // 端口冲突时往后顺延几个。范围要小——任务窗格那边是靠"带着令牌
        // 逐个探"来找我们的，范围大了探测就慢。
        private const int PortProbeSpan = 5;

        private static int Main(string[] args)
        {
            try
            {
                var opts = Options.Parse(args);
                if (opts == null) { PrintUsage(); return 2; }

                // 【单实例】。.xlam 每次 Excel 启动都会 Shell() 我们一下，
                // 不拦的话开几个 Excel 就有几个 sidecar，各占一个端口，
                // 任务窗格探到哪个全看运气。
                // 用命名互斥体：抢不到就说明已经有一个在跑，安静退出。
                if (!string.IsNullOrEmpty(opts.SingleInstanceKey))
                {
                    bool createdNew;
                    // Local\ 前缀 = 每个登录会话各一个。多用户共用一台机器
                    // （终端服务器）时，各人有各人的 sidecar，互不干扰。
                    _singleInstanceMutex = new Mutex(true, "Local\\" + opts.SingleInstanceKey, out createdNew);
                    if (!createdNew)
                    {
                        Console.WriteLine("已经有一个 sidecar 在跑，本次不重复启动。");
                        return 0;
                    }
                }

                SidecarConfig cfg;
                try
                {
                    cfg = SidecarConfig.Load(opts.ConfigPath);
                }
                catch (Exception ex)
                {
                    // 【失败就是不启动】。没有有效令牌还把端口开起来，
                    // 等于开了一个谁都能用的本机后门。
                    Console.Error.WriteLine("配置无效，拒绝启动：" + ex.Message);
                    return 2;
                }

                int port = opts.Port > 0 ? opts.Port : cfg.Port;
                var server = new SidecarServer(cfg);

                int bound = server.Start(port, PortProbeSpan);
                if (bound <= 0)
                {
                    Console.Error.WriteLine(string.Format(CultureInfo.InvariantCulture,
                        "端口 {0}-{1} 都被占用，起不来。", port, port + PortProbeSpan - 1));
                    return 3;
                }

                Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                    "sidecar {0} 已就绪：127.0.0.1:{1}  pid={2}", Version, bound, CurrentPid()));

                // 【跟随 Excel 生存】：父进程没了就自己退。
                // 这是"关 Excel 就没了、用户完全无感"这个承诺的落点——
                // 少了它，sidecar 就变成一个用户看不见也关不掉的常驻进程。
                if (opts.ParentPid > 0) { WatchParent(opts.ParentPid, server); }

                // 【按宿主进程存活来判定，比盯着某一个 PID 更准】。
                // 用户常同时开好几个 Excel 窗口。盯着"拉起我的那一个"的话，
                // 那个先关了 sidecar 就没了，而别的 Excel 还开着——
                // 表现是"AI 在这个工作簿能用、在那个不能用"，极难排查。
                if (!string.IsNullOrEmpty(opts.WatchProcess)) { WatchHost(opts.WatchProcess, opts.WatchGraceSeconds, server); }

                server.WaitForShutdown();
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("启动失败：" + ex);
                return 1;
            }
        }

        private static void PrintUsage()
        {
            Console.Error.WriteLine("用法：ExcelToolboxSidecar.exe --config <路径> [--parent-pid <PID>] [--port <端口>]");
        }

        internal static int CurrentPid()
        {
            using (var p = Process.GetCurrentProcess()) { return p.Id; }
        }

        // 互斥体要一直拿在手里，进程活多久它活多久。
        // 【不能是局部变量】——那样会被 GC 回收，互斥体提前释放，单实例就失效了。
        private static Mutex _singleInstanceMutex;

        // 连续查不到宿主状态多少次就认输退出。2 秒一次，30 次 = 1 分钟。
        private const int MaxProbeFailures = 30;

        private static void WatchHost(string processName, int graceSeconds, SidecarServer server)
        {
            var t = new Thread(delegate()
            {
                // 启动时宿主可能还没起来（.xlam 加载得比 Excel 窗口早），
                // 所以给一段宽限期，别刚起来就把自己关了。
                var deadline = DateTime.UtcNow.AddSeconds(graceSeconds);
                bool everSeen = false;
                int probeFailures = 0;

                while (true)
                {
                    Thread.Sleep(2000);
                    int n;
                    try
                    {
                        n = Process.GetProcessesByName(processName).Length;
                        probeFailures = 0;
                    }
                    catch
                    {
                        // 查不到先当它还在，宁可多活一会儿也别误杀。
                        // 【但不能永远这么认】：查询要是一直失败（权限、系统异常），
                        // 把它恒定解释成"宿主还在"，sidecar 就【永远不退出】了——
                        // 而那正是我们答应用户绝不会发生的事（关了 Excel 就该没了）。
                        probeFailures++;
                        if (probeFailures >= MaxProbeFailures)
                        {
                            Console.Error.WriteLine("连续 " + probeFailures +
                                " 次查不到宿主进程状态，保险起见退出。");
                            server.Shutdown();
                            return;
                        }
                        continue;
                    }

                    if (n > 0) { everSeen = true; continue; }
                    if (!everSeen && DateTime.UtcNow < deadline) { continue; }

                    Console.WriteLine(processName + " 已全部退出，sidecar 跟着退出。");
                    server.Shutdown();
                    return;
                }
            });
            t.IsBackground = true;
            t.Start();
        }

        private static void WatchParent(int parentPid, SidecarServer server)
        {
            Process parent;
            try { parent = Process.GetProcessById(parentPid); }
            catch (ArgumentException)
            {
                // 父进程已经没了（或者 PID 是错的）。这时候不该继续活着。
                Console.Error.WriteLine("父进程 " + parentPid + " 不存在，退出。");
                server.Shutdown();
                return;
            }

            var t = new Thread(delegate()
            {
                try { parent.WaitForExit(); }
                catch { /* 拿不到就当它没了 */ }
                Console.WriteLine("父进程已退出，sidecar 跟着退出。");
                server.Shutdown();
            });
            t.IsBackground = true;
            t.Start();
        }
    }

    //=========================================================================
    // 命令行
    //=========================================================================
    internal sealed class Options
    {
        public string ConfigPath;
        public int ParentPid;
        public int Port;
        public string WatchProcess;      // 这个名字的进程一个都不剩时退出

        // 宿主还没起来时的宽限秒数。
        // 【这是给测试留的缝】，默认值就是生产行为——和安装器的 -WefRoot 一个路子。
        // 不留这条缝的话，"宿主进程消失就退出"这条只能靠等满 30 秒来验证，
        // 慢到没人愿意把它放进回归里，于是就不会有人测它。
        public int WatchGraceSeconds = 30;
        public string SingleInstanceKey; // 同一个 key 只允许跑一个

        public static Options Parse(string[] args)
        {
            var o = new Options();
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (a == "--config" && i + 1 < args.Length) { o.ConfigPath = args[++i]; }
                else if (a == "--parent-pid" && i + 1 < args.Length) { o.ParentPid = ParseInt(args[++i]); }
                else if (a == "--port" && i + 1 < args.Length) { o.Port = ParseInt(args[++i]); }
                else if (a == "--watch-process" && i + 1 < args.Length) { o.WatchProcess = args[++i]; }
                else if (a == "--single-instance" && i + 1 < args.Length) { o.SingleInstanceKey = args[++i]; }
                else if (a == "--watch-grace-seconds" && i + 1 < args.Length) { o.WatchGraceSeconds = ParseInt(args[++i]); }
                else { return null; }
            }
            if (string.IsNullOrEmpty(o.ConfigPath)) { return null; }
            return o;
        }

        private static int ParseInt(string s)
        {
            int v;
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out v) ? v : 0;
        }
    }

    //=========================================================================
    // 配置
    //=========================================================================
    internal sealed class SidecarConfig
    {
        public string Token;
        public int Port = 8899;
        public readonly List<string> AllowedOrigins = new List<string>();

        public static SidecarConfig Load(string path)
        {
            if (!File.Exists(path)) { throw new FileNotFoundException("找不到配置文件：" + path); }

            var raw = File.ReadAllText(path, Encoding.UTF8);
            var ser = new JavaScriptSerializer();
            var map = ser.Deserialize<Dictionary<string, object>>(raw);
            if (map == null) { throw new InvalidDataException("配置不是一个 JSON 对象"); }

            var cfg = new SidecarConfig();

            object tokenObj;
            if (!map.TryGetValue("token", out tokenObj) || tokenObj == null)
            {
                throw new InvalidDataException("配置里没有 token");
            }
            cfg.Token = Convert.ToString(tokenObj, CultureInfo.InvariantCulture);
            if (cfg.Token.Length < 32)
            {
                // 这里的下限和 Program.MinTokenLength 是同一件事，写死在两边不好，
                // 但配置解析要在 Program 之外独立可测，就近校验更直接。
                throw new InvalidDataException(string.Format(CultureInfo.InvariantCulture,
                    "token 太短（{0} 个字符），至少要 32 个", cfg.Token.Length));
            }

            object portObj;
            if (map.TryGetValue("port", out portObj) && portObj != null)
            {
                int p = Convert.ToInt32(portObj, CultureInfo.InvariantCulture);
                if (p > 0 && p <= 65535) { cfg.Port = p; }
            }

            object originsObj;
            if (map.TryGetValue("allowedOrigins", out originsObj) && originsObj != null)
            {
                // 【不要写成 originsObj as object[]】。
                // JavaScriptSerializer 把 JSON 数组反序列化成 ArrayList，不是 object[]，
                // 那样转出来是 null —— 结果是【白名单静默变空】：
                // 配置解析不报错、sidecar 照常启动、/health 也正常，
                // 只是每一个正常来源都被判 403。这种坏法最难查。
                // 顺带也接受单个字符串，安装器只配一个来源时不至于踩坑。
                var one = originsObj as string;
                if (one != null)
                {
                    AddOrigin(cfg, one);
                }
                else
                {
                    var seq = originsObj as System.Collections.IEnumerable;
                    if (seq != null)
                    {
                        foreach (var item in seq) { AddOrigin(cfg, item); }
                    }
                }
            }

            // 白名单为空时【不要静默放行】。这里保持为空即可——
            // IsOriginAllowed 会一律返回 false，方向是拒绝，不是放行。
            return cfg;
        }

        private static void AddOrigin(SidecarConfig cfg, object item)
        {
            if (item == null) { return; }
            var s = Convert.ToString(item, CultureInfo.InvariantCulture);
            if (!string.IsNullOrEmpty(s)) { cfg.AllowedOrigins.Add(s.TrimEnd('/')); }
        }

        public bool IsOriginAllowed(string origin)
        {
            if (string.IsNullOrEmpty(origin)) { return false; }
            var probe = origin.TrimEnd('/');
            foreach (var o in AllowedOrigins)
            {
                // 【必须整串比，不能用 StartsWith】。
                // StartsWith 的话 https://gw.corp.example 会把
                // https://gw.corp.example.evil.com 也放进来。
                if (string.Equals(o, probe, StringComparison.OrdinalIgnoreCase)) { return true; }
            }
            return false;
        }
    }

    //=========================================================================
    // 服务器
    //=========================================================================
    internal sealed class SidecarServer
    {
        private readonly SidecarConfig _cfg;
        private readonly ManualResetEvent _stopped = new ManualResetEvent(false);
        private TcpListener _listener;
        private volatile bool _running;
        private int _port;

        // 请求大小上限。没有上限的话，一个连接慢慢灌数据就能把内存吃光。
        private const int MaxHeaderBytes = 8 * 1024;
        private const int MaxBodyBytes = 1024 * 1024;

        public SidecarServer(SidecarConfig cfg) { _cfg = cfg; }

        public int Port { get { return _port; } }

        /// <summary>
        /// 从 startPort 起试 span 个端口，返回真正绑上的那个；都失败返回 0。
        /// </summary>
        public int Start(int startPort, int span)
        {
            for (int p = startPort; p < startPort + span; p++)
            {
                if (p <= 0 || p > 65535) { continue; }
                try
                {
                    // 【只绑 IPAddress.Loopback，绝不 IPAddress.Any】。
                    // 绑 Any 就等于把这个端口暴露给整个内网，
                    // 那时候令牌是唯一一道门，而且 Windows 防火墙还会弹窗。
                    var l = new TcpListener(IPAddress.Loopback, p);

                    // 【不要开 ReuseAddress】。开了以后端口已被别人占用时
                    // 也可能"绑成功"，于是两个进程抢同一个端口，
                    // 请求随机落到谁那儿都有可能——这种问题极难查。
                    l.Start();

                    _listener = l;
                    _port = p;
                    _running = true;
                    var t = new Thread(AcceptLoop);
                    t.IsBackground = true;
                    t.Start();
                    return p;
                }
                catch (SocketException)
                {
                    // 被占了，试下一个
                }
            }
            return 0;
        }

        public void Shutdown()
        {
            if (!_running) { return; }
            _running = false;
            try { if (_listener != null) { _listener.Stop(); } }
            catch { }
            _stopped.Set();
        }

        public void WaitForShutdown() { _stopped.WaitOne(); }

        private void AcceptLoop()
        {
            while (_running)
            {
                TcpClient client;
                try { client = _listener.AcceptTcpClient(); }
                catch { break; }   // Stop() 会让这里抛，属于正常退出路径

                var c = client;
                ThreadPool.QueueUserWorkItem(delegate { HandleSafely(c); });
            }
        }

        private void HandleSafely(TcpClient client)
        {
            try { Handle(client); }
            catch { /* 单个连接出问题不能把整个 sidecar 带走 */ }
            finally { try { client.Close(); } catch { } }
        }

        private void Handle(TcpClient client)
        {
            client.ReceiveTimeout = 10000;
            client.SendTimeout = 10000;

            using (var stream = client.GetStream())
            {
                var req = HttpRequest.Read(stream, MaxHeaderBytes, MaxBodyBytes);
                if (req == null) { Write(stream, 400, null, "{\"error\":\"bad_request\"}"); return; }

                string origin = req.Header("Origin");
                bool originOk = string.IsNullOrEmpty(origin) || _cfg.IsOriginAllowed(origin);
                string corsOrigin = (!string.IsNullOrEmpty(origin) && originOk) ? origin : null;

                // 预检。浏览器要先问过这一关，才肯把自定义令牌头发出来。
                if (req.Method == "OPTIONS")
                {
                    if (!originOk) { Write(stream, 403, null, "{\"error\":\"origin_not_allowed\"}"); return; }
                    Write(stream, 204, corsOrigin, null);
                    return;
                }

                if (!originOk) { Write(stream, 403, corsOrigin, "{\"error\":\"origin_not_allowed\"}"); return; }

                // 【查询串里带令牌一律拒绝】。
                // ?token=xxx 是简单请求，不触发预检——接受它就等于把
                // "自定义头强制预检"这道防线自己拆了。
                // 顺带：查询串会进浏览器历史、代理日志、Referer，本来也不该放密钥。
                if (req.QueryContainsTokenLikeKey())
                {
                    Write(stream, 400, corsOrigin, "{\"error\":\"token_must_be_in_header\"}");
                    return;
                }

                if (!Authorized(req))
                {
                    Write(stream, 401, corsOrigin, "{\"error\":\"unauthorized\"}");
                    return;
                }

                // 白名单式路由：只认列出来的具体动作。
                // 【绝不做"传一段 VBA 进来执行"的通用口子】——那个口子会把风险
                // 从"几个动作"放大成"任意代码执行"。要什么能力就在这里加一条。
                if (req.Method == "GET" && req.Path == "/health")
                {
                    Write(stream, 200, corsOrigin, BuildHealthJson());
                    return;
                }

                // 列出工作簿里的查询。只读，不刷新任何东西。
                if (req.Method == "GET" && req.Path == "/queries")
                {
                    Write(stream, 200, corsOrigin, ExcelBridge.ListQueries());
                    return;
                }

                // 刷新查询。
                // 【只接受查询的名字，不接受任何代码或 M 表达式】。
                // 要什么能力就在这里加一条具体路由——通用执行口子一开，
                // 风险就从"几个动作"变成"任意代码执行"。
                if (req.Method == "POST" && req.Path == "/refresh-query")
                {
                    string name = ExtractJsonString(req.Body, "name");
                    if (string.IsNullOrEmpty(name))
                    {
                        Write(stream, 400, corsOrigin, "{\"error\":\"missing_name\"}");
                        return;
                    }
                    Write(stream, 200, corsOrigin, ExcelBridge.RefreshQuery(name));
                    return;
                }

                Write(stream, 404, corsOrigin, "{\"error\":\"not_found\"}");
            }
        }

        /// <summary>
        /// 从请求体里取一个字符串字段。
        /// 用 JavaScriptSerializer，不手搓解析——手搓的 JSON 解析器是 bug 温床。
        /// </summary>
        private static string ExtractJsonString(string body, string key)
        {
            if (string.IsNullOrEmpty(body)) { return null; }
            try
            {
                var ser = new JavaScriptSerializer();
                var map = ser.Deserialize<Dictionary<string, object>>(body);
                object v;
                if (map != null && map.TryGetValue(key, out v) && v != null)
                {
                    return Convert.ToString(v, CultureInfo.InvariantCulture);
                }
            }
            catch { }
            return null;
        }

        private bool Authorized(HttpRequest req)
        {
            var supplied = req.Header(Program.TokenHeader);
            if (string.IsNullOrEmpty(supplied)) { return false; }
            return FixedTimeEquals(supplied, _cfg.Token);
        }

        /// <summary>
        /// 定长比较。普通的 == 会在第一个不同的字符处就返回，
        /// 攻击者据此能一位一位地把令牌试出来。
        /// </summary>
        private static bool FixedTimeEquals(string a, string b)
        {
            if (a == null || b == null) { return false; }
            int diff = a.Length ^ b.Length;
            int n = Math.Min(a.Length, b.Length);
            for (int i = 0; i < n; i++) { diff |= a[i] ^ b[i]; }
            return diff == 0;
        }

        private string BuildHealthJson()
        {
            return string.Format(CultureInfo.InvariantCulture,
                "{{\"ok\":true,\"name\":\"excel-toolbox-sidecar\",\"version\":\"{0}\",\"pid\":{1},\"port\":{2}}}",
                Program.Version, Program.CurrentPid(), _port);
        }

        private static void Write(Stream stream, int status, string corsOrigin, string json)
        {
            var sb = new StringBuilder();
            sb.Append("HTTP/1.1 ").Append(status).Append(' ').Append(StatusText(status)).Append("\r\n");

            if (!string.IsNullOrEmpty(corsOrigin))
            {
                // 【只回显白名单里的那个 Origin，绝不回 *】。
                // 回 * 等于允许任意网页读响应内容。
                sb.Append("Access-Control-Allow-Origin: ").Append(corsOrigin).Append("\r\n");
                sb.Append("Vary: Origin\r\n");
                sb.Append("Access-Control-Allow-Headers: ").Append(Program.TokenHeader).Append(", Content-Type\r\n");
                sb.Append("Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n");
                sb.Append("Access-Control-Max-Age: 600\r\n");
            }

            byte[] body = json == null ? new byte[0] : Encoding.UTF8.GetBytes(json);
            if (json != null) { sb.Append("Content-Type: application/json; charset=utf-8\r\n"); }
            sb.Append("Content-Length: ").Append(body.Length).Append("\r\n");
            sb.Append("Cache-Control: no-store\r\n");
            sb.Append("X-Content-Type-Options: nosniff\r\n");
            sb.Append("Connection: close\r\n\r\n");

            var head = Encoding.ASCII.GetBytes(sb.ToString());
            stream.Write(head, 0, head.Length);
            if (body.Length > 0) { stream.Write(body, 0, body.Length); }
            stream.Flush();
        }

        private static string StatusText(int status)
        {
            switch (status)
            {
                case 200: return "OK";
                case 204: return "No Content";
                case 400: return "Bad Request";
                case 401: return "Unauthorized";
                case 403: return "Forbidden";
                case 404: return "Not Found";
                default: return "Error";
            }
        }
    }

    //=========================================================================
    // Excel 桥
    //
    // 【必须挂到用户正开着的那个 Excel 上，绝不能自己新开一个】。
    // 新开一个进程的话，刷新的是一个空白工作簿——用户点了"刷新"，
    // 界面上什么都没变，也没有任何报错。Marshal.GetActiveObject 拿的是
    // 运行中的实例；拿不到就如实说 Excel 没开。
    //
    // 【全部用后期绑定】。引用 Microsoft.Office.Interop.Excel 要求编译机
    // 装了对应版本的 PIA，而且【版本一变就对不上】。反射调 IDispatch
    // 没有这个问题，代价是写起来啰嗦、没有编译期检查。
    //=========================================================================
    internal static class ExcelBridge
    {
        public static string ListQueries()
        {
            return RunOnSta(delegate(object app)
            {
                object wb = Get(app, "ActiveWorkbook");
                if (wb == null) { return "{\"ok\":false,\"error\":\"no_workbook\"}"; }

                var names = new List<string>();
                object queries = Get(wb, "Queries");
                if (queries != null)
                {
                    int count = Convert.ToInt32(Get(queries, "Count"), CultureInfo.InvariantCulture);
                    for (int i = 1; i <= count; i++)
                    {
                        object q = Invoke(queries, "Item", i);
                        if (q == null) { continue; }
                        names.Add(Convert.ToString(Get(q, "Name"), CultureInfo.InvariantCulture));
                    }
                }

                var sb = new StringBuilder();
                sb.Append("{\"ok\":true,\"queries\":[");
                for (int i = 0; i < names.Count; i++)
                {
                    if (i > 0) { sb.Append(','); }
                    sb.Append('"').Append(JsonEscape(names[i])).Append('"');
                }
                sb.Append("]}");
                return sb.ToString();
            });
        }

        public static string RefreshQuery(string name)
        {
            return RunOnSta(delegate(object app)
            {
                object wb = Get(app, "ActiveWorkbook");
                if (wb == null) { return "{\"ok\":false,\"error\":\"no_workbook\"}"; }

                object connections = Get(wb, "Connections");
                if (connections == null) { return "{\"ok\":false,\"error\":\"no_connections\"}"; }

                // Power Query 建出来的连接叫「Query - <查询名>」。
                // 两种名字都认一下，省得因为版本差异找不到。
                object target = TryItem(connections, "Query - " + name);
                if (target == null) { target = TryItem(connections, name); }

                if (target == null)
                {
                    return "{\"ok\":false,\"error\":\"query_not_found\",\"name\":\"" + JsonEscape(name) + "\"}";
                }

                Invoke(target, "Refresh");
                return "{\"ok\":true,\"refreshed\":\"" + JsonEscape(name) + "\"}";
            });
        }

        //---------------------------------------------------------------------
        // 【COM 调用放到专用的 STA 线程上】。
        // 请求是在线程池线程（MTA）上处理的，从 MTA 直接打 Excel 的
        // IDispatch 要跨套间封送，时灵时不灵——而且失败方式很难看：
        // 偶发的 RPC_E_* 错误，重试一次又好了。专门起一个 STA 线程做这件事。
        //---------------------------------------------------------------------
        private delegate string ExcelWork(object app);

        private static string RunOnSta(ExcelWork work)
        {
            string result = null;
            Exception failure = null;

            var t = new Thread(delegate()
            {
                object app = null;
                _comScope = new List<object>();
                try
                {
                    try
                    {
                        app = Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        result = "{\"ok\":false,\"error\":\"excel_not_running\"}";
                        return;
                    }
                    result = work(app);
                }
                catch (Exception ex) { failure = ex; }
                finally
                {
                    // 先放中途拿到的那些，最后才放 app —— 逆序
                    ReleaseScope();
                    _comScope = null;
                    if (app != null) { try { Marshal.FinalReleaseComObject(app); } catch { } }
                }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();

            // 【必须有上限】。Excel 弹了个模态对话框的话，COM 调用会一直挂着，
            // 那条连接就永远不回包了。超时后如实回错，别让任务窗格干等。
            if (!t.Join(TimeSpan.FromSeconds(60)))
            {
                return "{\"ok\":false,\"error\":\"timeout\",\"hint\":\"Excel 可能正弹着对话框\"}";
            }

            if (failure != null)
            {
                return "{\"ok\":false,\"error\":\"com_failed\",\"message\":\"" +
                       JsonEscape(failure.Message) + "\"}";
            }
            return result ?? "{\"ok\":false,\"error\":\"no_result\"}";
        }

        // 【每一个拿到的 COM 对象都要还回去】。
        // 只释放 app 是不够的：ActiveWorkbook / Queries / 每个 Query /
        // Connections 都是独立的 RCW。漏掉它们，Excel 进程会被引用吊住——
        // 表现是【用户关掉 Excel 窗口，进程还在后台赖着】，而且越用越多。
        [ThreadStatic] private static List<object> _comScope;

        /// <summary>取属性，并把拿到的 COM 对象登记进待释放清单。</summary>
        private static object Get(object target, string name)
        {
            if (target == null) { return null; }
            var v = target.GetType().InvokeMember(
                name, BindingFlags.GetProperty, null, target, null, CultureInfo.InvariantCulture);
            Track(v);
            return v;
        }

        private static void Track(object v)
        {
            if (v == null) { return; }
            if (!Marshal.IsComObject(v)) { return; }
            if (_comScope != null) { _comScope.Add(v); }
        }

        /// <summary>按取得的逆序释放。逆序是因为子对象要先于父对象放掉。</summary>
        private static void ReleaseScope()
        {
            if (_comScope == null) { return; }
            for (int i = _comScope.Count - 1; i >= 0; i--)
            {
                try { Marshal.FinalReleaseComObject(_comScope[i]); } catch { }
            }
            _comScope.Clear();
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            if (target == null) { return null; }
            var v = target.GetType().InvokeMember(
                name, BindingFlags.InvokeMethod, null, target, args, CultureInfo.InvariantCulture);
            Track(v);
            return v;
        }

        private static object TryItem(object collection, string key)
        {
            // 取不到会抛，而"取不到"恰恰是正常分支——不能让它冒到上面去
            try { return Invoke(collection, "Item", key); }
            catch { return null; }
        }

        private static string JsonEscape(string s)
        {
            if (string.IsNullOrEmpty(s)) { return ""; }
            var sb = new StringBuilder(s.Length + 8);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 32) { sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture)); }
                        else { sb.Append(c); }
                        break;
                }
            }
            return sb.ToString();
        }
    }

    //=========================================================================
    // 极简 HTTP 请求解析
    //
    // 只解析我们自己用得到的那点东西。【这不是通用 HTTP 服务器】，
    // 不支持 keep-alive、分块传输、管线化——响应一律 Connection: close。
    //=========================================================================
    internal sealed class HttpRequest
    {
        public string Method;
        public string Path;
        public string Query;
        public string Body;
        private readonly Dictionary<string, string> _headers =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public string Header(string name)
        {
            string v;
            return _headers.TryGetValue(name, out v) ? v : null;
        }

        /// <summary>
        /// 查询串里有没有形如 token / auth / key 的参数。
        /// 有就拒绝——见 Handle() 里的说明。
        /// </summary>
        public bool QueryContainsTokenLikeKey()
        {
            if (string.IsNullOrEmpty(Query)) { return false; }
            var parts = Query.Split('&');
            foreach (var part in parts)
            {
                var eq = part.IndexOf('=');
                var key = (eq >= 0 ? part.Substring(0, eq) : part).Trim().ToLowerInvariant();
                if (key == "token" || key == "auth" || key == "apikey" || key == "api_key" || key == "key")
                {
                    return true;
                }
            }
            return false;
        }

        public static HttpRequest Read(Stream stream, int maxHeaderBytes, int maxBodyBytes)
        {
            var headerBytes = ReadUntilHeaderEnd(stream, maxHeaderBytes);
            if (headerBytes == null) { return null; }

            var text = Encoding.ASCII.GetString(headerBytes.Data, 0, headerBytes.HeaderLength);
            var lines = text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            if (lines.Length == 0) { return null; }

            var requestLine = lines[0].Split(' ');
            if (requestLine.Length < 2) { return null; }

            var req = new HttpRequest();
            req.Method = requestLine[0].ToUpperInvariant();

            var target = requestLine[1];
            var q = target.IndexOf('?');
            if (q >= 0) { req.Path = target.Substring(0, q); req.Query = target.Substring(q + 1); }
            else { req.Path = target; req.Query = ""; }

            for (int i = 1; i < lines.Length; i++)
            {
                var line = lines[i];
                if (line.Length == 0) { continue; }
                var colon = line.IndexOf(':');
                if (colon <= 0) { continue; }
                var name = line.Substring(0, colon).Trim();
                var value = line.Substring(colon + 1).Trim();
                // 同名头重复出现时保留第一个，避免"后一个覆盖前一个"被利用来绕过校验
                if (!req._headers.ContainsKey(name)) { req._headers[name] = value; }
            }

            int contentLength = 0;
            var cl = req.Header("Content-Length");
            if (!string.IsNullOrEmpty(cl))
            {
                // 【解析不了就拒绝，不能当成 0】。当成 0 的话请求体被丢掉，
                // 但请求仍按"没有正文"继续处理——我们和客户端对同一个请求的
                // 理解就不一致了，这类歧义正是 HTTP 走私类问题的温床。
                if (!int.TryParse(cl, NumberStyles.None, CultureInfo.InvariantCulture, out contentLength))
                {
                    return null;
                }
            }
            if (contentLength < 0 || contentLength > maxBodyBytes) { return null; }

            if (contentLength > 0)
            {
                var body = new byte[contentLength];
                int already = headerBytes.Length - headerBytes.HeaderLength;
                int copy = Math.Min(already, contentLength);
                if (copy > 0) { Array.Copy(headerBytes.Data, headerBytes.HeaderLength, body, 0, copy); }

                int got = copy;
                while (got < contentLength)
                {
                    int n = stream.Read(body, got, contentLength - got);
                    if (n <= 0) { break; }
                    got += n;
                }
                if (got < contentLength) { return null; }
                req.Body = Encoding.UTF8.GetString(body);
            }
            else { req.Body = ""; }

            return req;
        }

        private sealed class RawHead
        {
            public byte[] Data;
            public int Length;
            public int HeaderLength;   // 含结尾的 \r\n\r\n
        }

        private static RawHead ReadUntilHeaderEnd(Stream stream, int max)
        {
            var buf = new byte[max];
            int len = 0;
            while (len < max)
            {
                int n = stream.Read(buf, len, max - len);
                if (n <= 0) { break; }
                len += n;

                for (int i = 3; i < len; i++)
                {
                    if (buf[i] == '\n' && buf[i - 1] == '\r' && buf[i - 2] == '\n' && buf[i - 3] == '\r')
                    {
                        var h = new RawHead();
                        h.Data = buf; h.Length = len; h.HeaderLength = i + 1;
                        return h;
                    }
                }
            }
            return null;   // 头太大或者对方没发完就断了
        }
    }
}
