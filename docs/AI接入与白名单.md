# AI 接入与白名单分流

工具箱的 AI 能力由一个**独立的 Office.js 加载项**提供（参考实现见 `Excel AI` 项目），
和本仓库的 VBA 工具箱是两个东西，各自安装、各自升级。

本文只定义一件事：**怎么把"这个人是谁"告诉任务窗格，从而决定他用公司配好的模型
还是自己配模型。**

---

## 先说清楚这不是安全机制

> **用户名分流是"分流"，不是"鉴权"。**
>
> 下面的方案里，用户名来自客户端，用户完全可以改掉它冒充别人。
> 这是**已知并接受的取舍**，前提是：IT 配置的那个模型本身不怕被多用几个人
> （没有敏感数据访问权、成本可控）。
>
> 一旦那个模型涉及成本分摊或能读到敏感数据，**必须换成网关侧的 Windows 集成认证**
> （Kerberos/Negotiate），由服务器从票据里解析域账号。那才是不可伪造的。
> 本文最后一节给了迁移路径。

---

## 为什么不能直接读用户名

最自然的想法是"任务窗格自己读一下当前用户名"。**做不到。**

Office.js 的任务窗格是一个沙箱网页，没有任何用户身份 API——
Excel / Word / PowerPoint 的 Office.js 都没有（只有 Outlook 有 `userProfile`）。
它读不到环境变量，也读不到本地文件。

所以身份只能**从外部注入**。

---

## 方案：安装时把身份写进 manifest 的 URL

安装脚本是以用户身份运行的，它知道 `%USERNAME%`。
于是让它生成一份**该用户专属的 manifest**，把身份放进任务窗格的 URL：

```
安装时（Install-Toolbox.ps1）
  读 %USERNAME%  →  把 manifest 模板里的 {{USER}} 替换掉
                 →  写到 %LOCALAPPDATA%\ExcelToolbox\ai\manifest.xml
                 →  注册 HKCU\...\Wef\Developer 指向该目录

manifest.xml 里
  <SourceLocation DefaultValue="https://网关/taskpane/index.html?u=zhangsan"/>

任务窗格启动时
  const u = new URLSearchParams(location.search).get('u')
  fetch(`https://网关/api/ai-config?u=${encodeURIComponent(u)}`)

网关返回
  命中白名单 → { "mode": "managed", "baseUrl": "...", "model": "...", "apiKey": "..." }
  未命中     → { "mode": "byok" }
```

- `managed`：任务窗格直接用返回的配置，**用户什么都不用填**
- `byok`：任务窗格显示设置页，用户自己填内网模型地址（复用 `Excel AI` 现有的设置页）

---

## 部署前提：网关证书必须已被客户端信任

Office.js 要求任务窗格的 `SourceLocation` 必须是 **https**（`localhost` 除外），
所以网关得有个被客户端认的证书。两种都行：

- 域内 PKI 统一下发的根所签的证书（企业里最常见）
- 公网证书

> **安装程序不再装任何证书。** 早先它会把内网自签 CA 装进当前用户的
> 受信任根存储。那是降低用户整台机器防护等级的操作——那个根证书能为
> **任意域**签发被这台机器信任的证书，影响远不止这一个加载项。
> 装个 Excel 插件不该有这种副作用。
>
> 只能提供自签证书的话，由 IT 用组策略统一下发根证书，
> 而不是让安装包替每个用户做这个决定。

> ⚠ **证书不受信任时，任务窗格是空白的，而且不报任何错。**
> 用户只会说"AI 打不开"，排查起来毫无线索。
> 部署前务必在一台**普通用户的机器**上用浏览器访问一次网关地址，
> 不弹证书警告才算过。

网关只需要托管静态文件 + 一个配置接口，**不需要代理大模型**——
模型地址是网关在下面那个接口里回给任务窗格的，或者由用户自己填。

---

## 网关要实现的接口

只有一个端点。

### `GET /api/ai-config?u=<用户名>`

**命中白名单：**

```json
{
  "mode": "managed",
  "baseUrl": "https://模型网关/v1",
  "model": "qwen2.5-72b-instruct",
  "apiKey": "sk-...",
  "notice": "你正在使用公司配置的模型"
}
```

**未命中：**

```json
{
  "mode": "byok",
  "notice": "未在白名单内，请配置你自己的内网模型"
}
```

**实现要点：**

| 要点 | 说明 |
|---|---|
| 白名单存哪 | 一个文本文件即可，一行一个用户名。改完不需要重启，读文件就行 |
| 大小写 | 用户名比对**必须忽略大小写**。Windows 登录名不区分大小写，`ZhangSan` 和 `zhangsan` 是同一个人 |
| 缺参数 | `u` 为空或缺失时返回 `byok`，不要报错——那只会让用户看到一个白屏 |
| CORS | 必须允许任务窗格的来源，否则浏览器直接拦掉。这是内网自建网关最常见的坑 |
| apiKey | 只在 `managed` 时下发。**它会出现在浏览器能看到的地方**，所以那个 key 应当是专供此用途、可随时轮换的，不要用主账号的 key |

### 参考实现（Nginx + 一个静态文件也能凑合）

最省事的做法：不写服务，直接放两个静态 JSON，用 Nginx 按用户名映射。
白名单几十人以内完全够用：

```nginx
location /api/ai-config {
    # $arg_u 是 URL 上的 u 参数，转成小写后查 map
    if ($ai_mode = managed) { return 200 '{"mode":"managed","baseUrl":"https://gw/v1","model":"qwen2.5-72b","apiKey":"sk-xxx"}'; }
    return 200 '{"mode":"byok"}';
}
```

```nginx
map $arg_u $ai_mode {
    default    byok;
    ~*^zhangsan$  managed;
    ~*^lisi$      managed;
}
```

改白名单 = 改一行 + `nginx -s reload`，不需要写代码。

---

## 任务窗格侧要改什么（`Excel AI` 项目）

**本仓库不包含这部分**，它属于 `Excel AI` 项目。需要的改动很小：

1. 启动时读 URL 参数并拉配置：

```ts
// src/taskpane/main.tsx 启动处
const u = new URLSearchParams(location.search).get('u') ?? '';
const cfg = await fetch(`/api/ai-config?u=${encodeURIComponent(u)}`)
  .then(r => r.json())
  .catch(() => ({ mode: 'byok' }));          // 网关挂了就退回自配置，别让用户卡在白屏

if (cfg.mode === 'managed') {
  useSettings.getState().set({
    kind: 'openai-compatible',
    baseUrl: cfg.baseUrl,
    model: cfg.model,
    apiKey: cfg.apiKey,
  });
}
```

2. `managed` 模式下把设置页的接口地址/模型/Key 三项设为只读，并显示 `notice`。
   **不要整个隐藏设置页**——用户仍然需要看到"我现在用的是哪个模型"。

3. `PRESETS` 里的 `intranet` 项保留，给 `byok` 用户当模板
   （`src/store/settings.ts` 已有）。

4. `testConnection()` **必须保留**。它会真发一次带工具的请求验证模型支持
   **function calling**——不支持工具调用的模型没法操作文档，
   等用户实际对话时才发现就太晚了。

---

## 迁移到真正的鉴权

哪天需要把它变成安全机制，改动集中在网关，客户端几乎不动：

1. 网关开 Kerberos/Negotiate（IIS 勾"Windows 身份验证"，或 Nginx 上
   `auth_gss` 模块）
2. `/api/ai-config` 忽略 URL 上的 `u`，改从认证上下文取域账号
3. 客户端那个 `?u=` 参数留着不管，或者去掉——两种都行

**客户端不需要重装**，因为身份的来源从"URL 参数"变成了"HTTP 认证头"，
而认证头是浏览器/WebView 自动带的。

---

## 已知限制

| 限制 | 说明 |
|---|---|
| 用户名可伪造 | 见开头。用户能编辑 `%LOCALAPPDATA%\ExcelToolbox\ai\manifest.xml` 改掉 `?u=` |
| 换人登录同一台机器 | manifest 是按安装时的用户生成的。换个人登录这台机器需要重跑一次安装 |
| apiKey 暴露在前端 | `managed` 模式下 Key 会到浏览器里。必须用专供此用途、可轮换的 Key |
| 网关不可达 | 任务窗格退回 `byok` 模式，用户仍可自配置——不会整个用不了 |
