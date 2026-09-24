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
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
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

        //---------------------------------------------------------------------
        // 【调用宏的二次确认令牌】
        //
        // 光有 X-Toolbox-Token 不够——那个令牌只证明"这个页面在白名单里、
        // 拿到了令牌"，不证明"用户点了确认弹窗"。任务窗格的 mutate:structure
        // 确认只是 UI 策略，sidecar 完全不知道这件事，等于强制确认可以被
        // 绕过（持有令牌就能直接 POST /run-macro）。
        //
        // 所以把"确认"这件事下沉一步：任务窗格在用户点了确认之后，
        // 先调 /confirm-macro 换一张一次性、短时效、绑定"这个宏名 + 这些
        // 参数"的确认票；/run-macro 必须带着与本次调用完全匹配的票才放行，
        // 用一次就作废，换宏名或换参数原来的票就不认了。
        //
        // 这挡不住"可信来源自己被 XSS/供应链攻击"这种终极场景（那种场景下
        // 攻击脚本本来就能调用任务窗格能调用的一切），但能堵住"普通令牌
        // 单独泄露给白名单外/未触发过确认流程的调用方就能静默执行宏"。
        //---------------------------------------------------------------------
        private readonly Dictionary<string, ConfirmTicket> _confirmTickets =
            new Dictionary<string, ConfirmTicket>();
        private readonly object _confirmLock = new object();
        private static readonly TimeSpan ConfirmTicketTtl = TimeSpan.FromSeconds(120);

        private sealed class ConfirmTicket
        {
            public string ArgsKey;
            public DateTime ExpiresAtUtc;
        }

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

                // 列出「暴露给 AI 的宏」。只读，尽力而为——枚举 VBA 工程需要
                // 「信任对 VBA 工程对象模型的访问」，用户机上大概率没开，
                // 枚不出来不当错误处理，如实说「枚不出来」，不是「没有宏」。
                if (req.Method == "GET" && req.Path == "/macros")
                {
                    Write(stream, 200, corsOrigin, ExcelBridge.ListMacros());
                    return;
                }

                // 换一张"用户已确认"的一次性票。任务窗格在弹窗确认之后才调这个，
                // 不是在调用宏之前自动帮用户调——这一步本身不执行任何宏。
                if (req.Method == "POST" && req.Path == "/confirm-macro")
                {
                    string cname = ExtractJsonString(req.Body, "name");
                    if (string.IsNullOrEmpty(cname))
                    {
                        Write(stream, 400, corsOrigin, "{\"error\":\"missing_name\"}");
                        return;
                    }
                    if (!IsAllowedMacroName(cname))
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"macro_not_allowed\",\"hint\":\"宏名必须以 AI_ 开头，这是工作簿作者显式暴露给 AI 的宏\"}");
                        return;
                    }
                    object[] cargs;
                    if (!TryExtractJsonArgs(req.Body, "args", out cargs))
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"bad_args\",\"hint\":\"args 只能是字符串/数字/布尔组成的数组\"}");
                        return;
                    }

                    string ticketToken = GenerateConfirmToken();
                    string argsKey = CanonicalArgsKey(cname, cargs);
                    lock (_confirmLock)
                    {
                        PruneExpiredConfirmTickets();
                        _confirmTickets[ticketToken] = new ConfirmTicket
                        {
                            ArgsKey = argsKey,
                            ExpiresAtUtc = DateTime.UtcNow.Add(ConfirmTicketTtl)
                        };
                    }
                    Write(stream, 200, corsOrigin, string.Format(CultureInfo.InvariantCulture,
                        "{{\"ok\":true,\"confirmToken\":\"{0}\",\"expiresInSeconds\":{1}}}",
                        ticketToken, (int)ConfirmTicketTtl.TotalSeconds));
                    return;
                }

                // 调用工作簿里现成的宏。
                //
                // 【这是这套接口里风险最高的一条，因此约束也最严】：
                //   1. 只认 AI_ 前缀的宏名——这是工作簿作者的显式选择，
                //      不是"随便一个宏名都能调"。没有这个前缀，
                //      请求在业务逻辑跑起来之前就被拒绝。
                //   2. 只接受宏名和几个基本类型的参数（字符串/数字/布尔），
                //      不接受代码、不接受表达式——和 refresh-query 一样的原则。
                //   3. 必须带一张 /confirm-macro 发的、与本次宏名+参数完全匹配的
                //      confirmToken，否则拒绝执行——光有 X-Toolbox-Token 不够，
                //      那只证明"来源在白名单里"，不证明"用户点了确认"。
                //      详见上面 _confirmTickets 那段注释。
                if (req.Method == "POST" && req.Path == "/run-macro")
                {
                    string name = ExtractJsonString(req.Body, "name");
                    if (string.IsNullOrEmpty(name))
                    {
                        Write(stream, 400, corsOrigin, "{\"error\":\"missing_name\"}");
                        return;
                    }
                    if (!IsAllowedMacroName(name))
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"macro_not_allowed\",\"hint\":\"宏名必须以 AI_ 开头，这是工作簿作者显式暴露给 AI 的宏\"}");
                        return;
                    }
                    object[] macroArgs;
                    if (!TryExtractJsonArgs(req.Body, "args", out macroArgs))
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"bad_args\",\"hint\":\"args 只能是字符串/数字/布尔组成的数组\"}");
                        return;
                    }

                    string confirmToken = ExtractJsonString(req.Body, "confirmToken");
                    if (string.IsNullOrEmpty(confirmToken))
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"confirm_required\",\"hint\":\"先调用 /confirm-macro 换取确认令牌，用户确认后才能执行\"}");
                        return;
                    }

                    string expectedKey = CanonicalArgsKey(name, macroArgs);
                    bool ticketOk;
                    lock (_confirmLock)
                    {
                        PruneExpiredConfirmTickets();
                        ConfirmTicket ticket;
                        ticketOk = _confirmTickets.TryGetValue(confirmToken, out ticket)
                            && ticket.ExpiresAtUtc >= DateTime.UtcNow
                            && ticket.ArgsKey == expectedKey;
                        // 一次性：不管这次匹不匹配，用过（或试过）就废掉，不能被重放。
                        _confirmTickets.Remove(confirmToken);
                    }
                    if (!ticketOk)
                    {
                        Write(stream, 400, corsOrigin,
                            "{\"error\":\"confirm_invalid\",\"hint\":\"确认令牌无效、已过期，或与本次宏名/参数不匹配，请重新确认\"}");
                        return;
                    }

                    Write(stream, 200, corsOrigin, ExcelBridge.RunMacro(name, macroArgs));
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

        // 宏名必须以 AI_ 开头，其余只允许字母数字下划线——这是工作簿作者
        // 把某个 Sub 显式标记为"可以被 AI 调用"的方式，不是任意宏名都能调。
        private static readonly Regex AllowedMacroNamePattern =
            new Regex(@"^AI_[A-Za-z0-9_]+$", RegexOptions.Compiled);

        private static bool IsAllowedMacroName(string name)
        {
            return !string.IsNullOrEmpty(name) && AllowedMacroNamePattern.IsMatch(name);
        }

        /// <summary>
        /// 取 body 里 key 对应的数组，且【数组里只能是字符串/数字/布尔】。
        /// 没有这个字段时返回空数组（不算错误——很多宏不需要参数）；
        /// 字段存在但不是数组、或者里面混进了对象/数组，判成 false——
        /// 这里不接受任何"代码形状"的东西，只收原子值。
        /// </summary>
        private static bool TryExtractJsonArgs(string body, string key, out object[] args)
        {
            args = new object[0];
            if (string.IsNullOrEmpty(body)) { return true; }
            try
            {
                var ser = new JavaScriptSerializer();
                var map = ser.Deserialize<Dictionary<string, object>>(body);
                if (map == null) { return true; }

                object raw;
                if (!map.TryGetValue(key, out raw) || raw == null) { return true; }

                // 【不能写成 raw as object[]】。JavaScriptSerializer 把 JSON
                // 数组反序列化成 ArrayList，不是 object[]——那样转出来是
                // null，于是任何带参数的请求都会被判成"参数不是数组"而
                // 拒绝，即使调用方传的明明是一个合法数组。这正是本仓库
                // sidecar 配置解析踩过的同一个坑，写法要保持一致。
                var seq = raw as System.Collections.IEnumerable;
                if (seq == null || raw is string) { return false; }

                var list = new List<object>();
                foreach (var item in seq)
                {
                    if (item == null || item is string || item is bool ||
                        item is int || item is long || item is double || item is decimal)
                    {
                        list.Add(item);
                    }
                    else
                    {
                        return false;   // 字典或数组混进来了——不是原子值，拒绝
                    }
                }
                args = list.ToArray();
                return true;
            }
            catch { return false; }
        }

        /// <summary>256 位随机确认票。攻击者拿不到就没法伪造"用户已确认"。</summary>
        private static string GenerateConfirmToken()
        {
            var bytes = new byte[32];
            using (var rng = new RNGCryptoServiceProvider()) { rng.GetBytes(bytes); }
            var sb = new StringBuilder(bytes.Length * 2);
            for (int i = 0; i < bytes.Length; i++) { sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture)); }
            return sb.ToString();
        }

        /// <summary>
        /// 把宏名 + 参数序列化成一个规范字符串，票据据此和"本次调用"绑死——
        /// 换个宏名或换个参数，原来那张确认票就不再匹配。
        /// </summary>
        private static string CanonicalArgsKey(string name, object[] args)
        {
            var sb = new StringBuilder();
            sb.Append(name).Append('|').Append(args.Length);
            foreach (var a in args)
            {
                sb.Append('|');
                if (a == null) { sb.Append("null"); }
                else { sb.Append(a.GetType().Name).Append(':').Append(Convert.ToString(a, CultureInfo.InvariantCulture)); }
            }
            return sb.ToString();
        }

        /// <summary>清掉过期票，防止长时间运行后字典无限增长。调用前必须已持有 _confirmLock。</summary>
        private void PruneExpiredConfirmTickets()
        {
            if (_confirmTickets.Count == 0) { return; }
            var now = DateTime.UtcNow;
            var expired = new List<string>();
            foreach (var kv in _confirmTickets)
            {
                if (kv.Value.ExpiresAtUtc < now) { expired.Add(kv.Key); }
            }
            foreach (var k in expired) { _confirmTickets.Remove(k); }
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

        // 只认 Public（或不写修饰符，VBA 里默认就是 Public）的 AI_ 开头的 Sub。
        // 【Private 的不匹配】：这一行如果是 "Private Sub AI_x(" ，前导的
        // "Private" 不满足下面这个模式（模式只允许可选的 "Public "），
        // 整行匹配失败——Application.Run 本来也调不动 Private 的宏，
        // 列出来也没用，干脆不收进结果里。
        //
        // 【行首是 ' 的不匹配】：(?!\s*') 排除注释行，比如
        // "' Sub AI_Foo(" 这种写在注释里的伪声明不会被当成真宏。
        private static readonly Regex AiSubPattern = new Regex(
            @"(?im)^(?!\s*')\s*(?:Public\s+)?Sub\s+(AI_[A-Za-z0-9_]+)\s*\(",
            RegexOptions.Compiled);

        // VBA 续行符" _"后面跟换行——声明可能写成
        //   Public Sub AI_Foo _
        //       (arg1 As String)
        // 先把续行拼接成一行再跑 AiSubPattern，否则名字和括号被换行隔开，
        // 声明会被漏掉（漏掉≠禁止调用，run-macro 该有的正则校验独立存在，
        // 只是模型会以为这个宏不存在）。
        private static readonly Regex LineContinuationPattern = new Regex(
            @"[ \t]_[ \t]*\r?\n", RegexOptions.Compiled);

        public static string ListMacros()
        {
            return RunOnSta(delegate(object app)
            {
                object wb = Get(app, "ActiveWorkbook");
                if (wb == null) { return "{\"ok\":false,\"error\":\"no_workbook\"}"; }

                object vbProject;
                try { vbProject = Get(wb, "VBProject"); }
                catch (COMException)
                {
                    // 【枚不出来 ≠ 没有宏】。这几乎总是因为没开「信任对 VBA
                    // 工程对象模型的访问」——绝大多数用户机上就是没开，
                    // 而且不该为了"能列出宏名"这种锦上添花的功能去要求
                    // 用户开这个权限。如实说"枚不出来"，run-macro 本身
                    // 不需要这个权限，照样能用，只是用户得自己知道宏名。
                    //
                    // 【只吃 COMException】：VBProject 属性被拒绝访问时，
                    // 后期绑定的 InvokeMember 抛的就是这个类型。其他类型的
                    // 异常不代表"没开信任"，让它们照常往外传，走到 RunOnSta
                    // 统一的 com_failed 分支——那才是如实反映"枚举出错了"，
                    // 而不是被误判成"用户没开权限"。
                    return "{\"ok\":true,\"trusted\":false,\"macros\":[]}";
                }
                if (vbProject == null)
                {
                    return "{\"ok\":true,\"trusted\":false,\"macros\":[]}";
                }

                var names = new List<string>();
                object components = Get(vbProject, "VBComponents");
                int count = Convert.ToInt32(Get(components, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object comp = Invoke(components, "Item", i);
                    if (comp == null) { continue; }
                    object codeModule = Get(comp, "CodeModule");
                    if (codeModule == null) { continue; }

                    int lineCount = Convert.ToInt32(Get(codeModule, "CountOfLines"), CultureInfo.InvariantCulture);
                    if (lineCount <= 0) { continue; }

                    string src = Convert.ToString(
                        Invoke(codeModule, "Lines", 1, lineCount), CultureInfo.InvariantCulture);
                    if (string.IsNullOrEmpty(src)) { continue; }

                    string joined = LineContinuationPattern.Replace(src, " ");
                    foreach (Match m in AiSubPattern.Matches(joined))
                    {
                        string n = m.Groups[1].Value;
                        if (!names.Contains(n)) { names.Add(n); }
                    }
                }

                var sb = new StringBuilder();
                sb.Append("{\"ok\":true,\"trusted\":true,\"macros\":[");
                for (int i = 0; i < names.Count; i++)
                {
                    if (i > 0) { sb.Append(','); }
                    sb.Append('"').Append(JsonEscape(names[i])).Append('"');
                }
                sb.Append("]}");
                return sb.ToString();
            });
        }

        public static string RunMacro(string name, object[] args)
        {
            return RunOnSta(delegate(object app)
            {
                object wb = Get(app, "ActiveWorkbook");
                if (wb == null) { return "{\"ok\":false,\"error\":\"no_workbook\"}"; }

                var callArgs = new List<object>();
                callArgs.Add(name);
                if (args != null) { callArgs.AddRange(args); }

                // 【调用本身不需要 VBOM 信任】——ListMacros 枚举宏名才需要，
                // Application.Run 按名字调用是普通的 Automation 调用，
                // 用户机上不开那个信任设置一样能用。
                object result;
                try
                {
                    result = Invoke(app, "Run", callArgs.ToArray());
                }
                catch (Exception ex)
                {
                    var inner = ex.InnerException ?? ex;
                    // 【失败必须如实说，不能报成功】。宏名打错、宏内部抛错、
                    // 参数个数不对，都会走到这里——用户以为宏跑了，其实没跑，
                    // 比直接看到报错更糟。
                    return "{\"ok\":false,\"error\":\"macro_failed\",\"name\":\"" + JsonEscape(name) +
                           "\",\"message\":\"" + JsonEscape(inner.Message) + "\"}";
                }

                string resultText = result == null
                    ? ""
                    : Convert.ToString(result, CultureInfo.InvariantCulture);
                return "{\"ok\":true,\"name\":\"" + JsonEscape(name) +
                       "\",\"result\":\"" + JsonEscape(resultText) + "\"}";
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
