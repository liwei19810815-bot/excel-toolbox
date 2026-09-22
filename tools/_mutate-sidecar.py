"""把 sidecar 源码逐条改坏，确认 check-sidecar.ps1 真的会变红。

只在本地跑，不入库。用法：python tools/_mutate-sidecar.py
"""
import io, os, shutil, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "sidecar", "ExcelToolboxSidecar.cs")
BAK = SRC + ".bak"

MUTATIONS = [
    ("令牌改成前缀比较",
     "            int diff = a.Length ^ b.Length;",
     "            int diff = 0; if (b.StartsWith(a)) { return true; }"),

    ("Origin 改成 StartsWith",
     "                if (string.Equals(o, probe, StringComparison.OrdinalIgnoreCase)) { return true; }",
     "                if (probe.StartsWith(o, StringComparison.OrdinalIgnoreCase)) { return true; }"),

    ("CORS 回 *",
     '                sb.Append("Access-Control-Allow-Origin: ").Append(corsOrigin).Append("\\r\\n");',
     '                sb.Append("Access-Control-Allow-Origin: *\\r\\n");'),

    ("绑到全部网卡",
     "                    var l = new TcpListener(IPAddress.Loopback, p);",
     "                    var l = new TcpListener(IPAddress.Any, p);"),

    # 【变异体本身必须能编译】。写成 if (false) 会触发"无法访问的代码"警告，
    # 而 build-sidecar.ps1 开了 /warnaserror+，于是编译直接失败——
    # 这条变异就【根本没测到】，却容易被当成"测过了"。改成运行时恒假的条件。
    ("接受查询串里的令牌",
     "                if (req.QueryContainsTokenLikeKey())",
     '                if (req.Path == "/__never__")'),

    ("允许 allowedOrigins 解析失败时静默放行",
     "                    var seq = originsObj as System.Collections.IEnumerable;",
     "                    var seq = originsObj as object[];"),

    ("去掉父进程守护",
     "                if (opts.ParentPid > 0) { WatchParent(opts.ParentPid, server); }",
     "                if (opts.ParentPid < 0) { WatchParent(opts.ParentPid, server); }"),
]


def run(cmd):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def main():
    shutil.copy2(SRC, BAK)
    results = []
    try:
        for name, old, new in MUTATIONS:
            s = io.open(SRC, encoding="utf-8-sig").read()
            if old not in s:
                results.append((name, "跳过", "没找到锚点"))
                continue
            io.open(SRC, "w", encoding="utf-8-sig", newline="").write(
                s.replace(old, new, 1).replace("\r\n", "\n").replace("\n", "\r\n"))

            b = run('powershell -ExecutionPolicy Bypass -File "%s"' %
                    os.path.join(ROOT, "build", "build-sidecar.ps1"))
            if b.returncode != 0:
                results.append((name, "编译失败", "变异本身编译不过，这条没测到"))
                shutil.copy2(BAK, SRC)
                continue

            t = run('powershell -ExecutionPolicy Bypass -File "%s"' %
                    os.path.join(ROOT, "tests", "check-sidecar.ps1"))
            out = (t.stdout or "") + (t.stderr or "")
            red = [l.strip() for l in out.splitlines() if "FAIL" in l]
            verdict = "变红" if t.returncode != 0 else "【没变红！】"
            detail = red[0][:70] if red else "没有任何断言失败"
            results.append((name, verdict, detail))
            shutil.copy2(BAK, SRC)
    finally:
        shutil.copy2(BAK, SRC)
        os.remove(BAK)
        run('powershell -ExecutionPolicy Bypass -File "%s"' %
            os.path.join(ROOT, "build", "build-sidecar.ps1"))

    print("\n=== 变异测试结果 ===")
    bad = 0
    for name, verdict, detail in results:
        print("  %-34s %-12s %s" % (name, verdict, detail))
        if verdict != "变红":
            bad += 1
    print("\n没能变红/没测到的：%d" % bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
