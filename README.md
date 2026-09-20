# Excel 通用工具箱（ExcelToolbox.xlam）

纯 VBA 实现的 Excel 加载宏工具箱，覆盖单元格/文本处理、数据处理、工作表管理、多文件合并、
文件批处理、公式工具、数据体检、数据可视化等常用功能，并内置**统一撤销框架**。

**v1.0.0**：M1–M9 共 59 个命令，**每个命令都有回归测试**（192 个断言，命令覆盖 59/59），
并内置**内网自动更新**——同事打开 Excel 即是最新版，升级和回滚都不需要他们做任何操作。

- 功能清单：[docs/功能清单.md](docs/功能清单.md)
- 部署与自动更新：[docs/部署与自动更新.md](docs/部署与自动更新.md)

---

## 为什么源码不是 .xlam

`.xlam` 是 zip 打包的二进制，git 无法 diff、无法合并。所以仓库里**只存源码**：

| 路径 | 内容 |
|---|---|
| `src/code/` | VBA 源码（`.bas` / `.cls` / `.frm`），git 的唯一真相 |
| `src/package/` | 需要注入 xlam 包的 XML，目前是 `customUI/customUI14.xml`（Ribbon 定义） |
| `dist/` | 构建产物，不入 git |

`.xlam` 由 `build/build.ps1` 从源码生成。这套布局参考自 [byronwall/bUTL](https://github.com/byronwall/bUTL)。

---

## 一次性前置设置

构建脚本要用 COM 把 VBA 源码导入工作簿，需要 Excel 开启对 VBA 工程对象模型的信任：

> Excel → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
> → 勾选 **信任对 VBA 工程对象模型的访问**

只影响本机开发环境，最终用户安装 `.xlam` 不需要这个设置。
构建脚本会自行检测，未开启时直接中止并给出提示（脚本不会替你改注册表）。

---

## 日常流程

构建：

```bash
powershell -ExecutionPolicy Bypass -File "build/build.ps1"
```

安装到 Excel（需先完全关闭 Excel）：

```bash
powershell -ExecutionPolicy Bypass -File "build/install.ps1"
```

卸载：

```bash
powershell -ExecutionPolicy Bypass -File "build/install.ps1" -Uninstall
```

验证构建产物（整工程编译 + 自检）：

```bash
timeout 120 powershell -ExecutionPolicy Bypass -File "build/verify.ps1"
```

跑回归测试。分成两个脚本，因为性质不同——主套件全在内存里跑，文件套件要在
`%TEMP%` 下建真实的 xlsx / 图片夹具并真的改磁盘文件：

```bash
timeout 900 powershell -ExecutionPolicy Bypass -File "tests/run-tests.ps1"
```

```bash
timeout 700 powershell -ExecutionPolicy Bypass -File "tests/run-tests-files.ps1"
```

自动更新（瘦加载器）的回归测试：

```bash
timeout 700 powershell -ExecutionPolicy Bypass -File "tests/run-tests-loader.ps1"
```

验证 Ribbon 真的被 Excel 接受（需要可见 Excel，所以单独一个脚本）：

```bash
timeout 200 powershell -ExecutionPolicy Bypass -File "tests/check-ribbon.ps1"
```

> **这几个脚本都必须套 `timeout`。** VBA 编译错误会弹 VBE 模态框，而 Excel 是无界面运行的，
> 那个框谁也看不见、谁也点不掉，脚本会永久挂起。超时即视为失败。
> 超时后记得清掉残留的 Excel 进程，否则它会占着 `dist/ExcelToolbox.xlam`，下次构建直接失败。

在 VBE 里直接改完代码后，把改动回写到 `src/code`：

```bash
powershell -ExecutionPolicy Bypass -File "build/export.ps1"
```

> `export.ps1` 按组件名在 `src/code` 下递归找同名文件原地覆盖，保持分层目录；
> 新建的组件会落到 `src/code/_new`，需人工挪到对应子目录。

---

## 架构要点

### 所有命令走同一条管线

```
Ribbon 按钮 (tag = actionId)
      ↓
modRibbon.Ribbon_OnAction        只解析参数，不含业务逻辑
      ↓
modAction.RunAction(actionId)    前置校验 → 确认 → 高速模式 → 撤销事务
      ↓                          → 执行 → 提交/回滚 → 恢复环境 → 统一报错
modAction.Dispatch               唯一的 actionId → 业务过程转派点
      ↓
业务模块过程                      不碰 ScreenUpdating、不碰撤销、不弹错误框
```

**不允许 Ribbon 直接调业务过程**——那样该操作就绕过了撤销框架和环境恢复。

### 新增一个工具

1. `src/package/customUI/customUI14.xml` 加按钮，`tag` 设为 actionId
2. `modAction.RegisterAll` 里加一行 `RegisterAction`
3. `modAction.Dispatch` 的 `Select Case` 里加一行，指向业务过程

### 撤销

VBA 宏一执行，Excel 原生 Ctrl+Z 就永久失效且无法恢复，所以必须自建：

- 快照存在一个 `IsAddin = True` 的内存工作簿里，一个事务一张表。用 IsAddin 而不是普通隐藏
  工作簿，是因为后者会在用户关掉自己所有文件后**仍把 Excel 进程吊住不退出**
- 用 `Range.Copy` 做快照，一次同时保住值、公式、数字格式、字体、边框、底纹、批注、合并状态。
  手工存 Variant 数组的话这些要逐项处理，极易漏项
- **结构性操作**（删除行列、排序、拆分合并单元格）必须快照整张表的 UsedRange，否则地址错位
- 多区域收敛成外接矩形再存；同一事务内多个快照**逆序还原**
- 默认保留 5 步；超过规模阈值（默认 100 万单元格）的操作不入栈
- 只在**当前 Excel 会话**内有效，关掉 Excel 即失效

**能力边界**：跨文件写入、文件重命名、导出、打印**无法撤销**。这类操作不注册事务，
改为执行前强制确认（`clsActionDef.ConfirmBeforeRun`），UI 上如实标注，不做"假装可撤销"。

### 踩过的坑（加新模块时还会再撞）

**1. 别用和 VBA 关键字同名的标识符。** VBA 大小写不敏感，`Dim eNum As Long` 里的 `eNum`
就是保留字 `Enum`，直接报"语法错误"。同类要避开的还有 `Name`、`Type`、`IsEmpty`、`Error`、`Len`。
症状很有迷惑性：VBE 里那一行标红，但相邻行看着完全正常。

**2. VBA 是按需编译的。** 调用一个函数只编译它用到的那条路径，别的过程里的语法错误照样能
蒙混过关，直到用户点到那个按钮才炸。所以 `verify.ps1` 必须显式触发
VBE 的「编译 VBAProject」（命令 ID 578），而且要先激活加载宏的工程——
否则编译的是那个临时空工作簿，等于什么都没验证。

**3. 源文件必须转成 CRLF + 系统 ANSI 才能被 VBE 正确导入。** 踩中任何一个都不会报错，
只会静默产出一个编译不过的工程：LF 换行会让解析器认不出 `.cls` 的 `VERSION 1.0 CLASS`
头部，**把类模块当成标准模块导入**，头部行随即变成语法错误；UTF-8 则让中文注释变成乱码。
`build.ps1` 会在导入后断言组件类型，不对就直接失败。

**4. 业务过程里不许 `MsgBox`。** 无头运行时它会弹在一个不可见的 Excel 里，谁也点不到，
永久挂死。结果一律以字符串返回，由 `RunAction` 统一呈现；测试前先开
`Toolbox_SetSilent`。

**5. PowerShell 5.1 读无 BOM 的 `.ps1` 会按 GBK 解码**，中文注释被打碎导致语法报错。
`build/` 和 `tests/` 下的脚本一律存 UTF-8 **with BOM**。

**6. 不要用 `StrConv(s, vbNarrow)` / `vbWide` 做全角半角转换。** 它依赖操作系统的 DBCS 支持，
在相当多的环境里对**任意输入**都抛运行时错误 5（本机 Excel 2016 64 位中文环境实测如此，
连纯 ASCII 都抛），而且只在运行到那一行才炸。一律用 `modStr.ToHalfWidth` / `ToFullWidth`
的显式码位映射。顺带两个 VBA 陷阱：十六进制字面量超过 `&H7FFF` 会被当成负 Integer，
必须写 `&HFF5E&`；`AscW` 返回带符号 Integer，码位大于 32767 时是负数，要补回 65536。

**7. `Range.Find` 的起点落在合并单元格上会静默返回 Nothing**，把有数据的表判成空表。
`modRange.RealUsedRange` 因此有两层保护：`After` 用右下角单元格，且 Find 失败时退回
`UsedRange`——宁可多算几行残留格式，也不能让所有工具对着合并表罢工。

**8. 往单元格写值之前先把数字格式改好。** 源单元格是文本格式时（文本型数字、文本型日期
本来就是），直接写入数值或日期会被 Excel 原样存成文本字符串，**事后再设格式也救不回来**。
`text.toNumber` 和 `misc.normalizeDates` 都是先改格式再写值。

**9. 改完 VBA 一定要跑 `verify.ps1`，只跑 `build.ps1` 不够。** 构建只做导入，不编译。
我就因此漏过一次「`On Error GoTo Failed` 但没写 `Failed:` 标签」——构建通过、前面几十个
用例也照常跑，直到调用到那个函数才弹出模态框把测试挂死。

**10. 迷你图必须在 `ScreenUpdating = True` 时创建**，否则 `SparklineGroups.Add` 抛 1004。
`AddSparklines` 因此临时打开屏幕刷新再还原——注意不能调 `modPerf.FastModeOff`，
那个带嵌套计数，会把整条管线的高速模式一起关掉。

**11. `Range.Copy` 之后要清 `Application.CutCopyMode`。** 撤销框架每次快照都用 Copy，
留下的复制模式会让后续某些 API 报出和复制毫无关系的 1004。`modUndo.CaptureBox` 已统一处理。

**12. 局部变量不能和同模块的函数同名。** VBA 不区分大小写，
`Dim cacheDir As String : cacheDir = CacheDir()` 里两者是同一个标识符，
编译器把 `CacheDir()` 当成对字符串变量取下标，报"缺少数组"。
和第 9 条一样，这是编译期错误，只跑 `build.ps1` 发现不了。

**13. Excel 宏里 `CreateObject("WScript.Shell")` 可能被安全策略拦下。**
它是宏病毒的典型载体，企业 AV 和 Windows ASR 规则普遍会阻止 Office 创建它，
而且**失败是静默的**——本机实测注册表读取直接返回空值。
加载器因此完全不依赖它，也不做任何运行时配置。

**14. 多进程共享同一个缓存目录时，临时文件名必须唯一。** 几十个人同时开 Excel
是每天早上的常态。固定临时名会让两个进程互删对方的中间文件，最后留下一个
拼出来的半截文件——而它看起来和正常文件没区别。加载器的临时名带进程 ID + 随机数，
并且改名失败时先检查"是不是别人已经装好了"，而不是直接报错。

**15. 别给模块级过程起名 `Trace`。** 加了一个 `Public Sub Trace` 之后，整个加载宏的
每一次 `Application.Run` 都开始无响应，而且编译照样通过。改名即恢复。同类的名字（和
Office 对象模型隐式成员重名的）都要避开。

### 兼容性

- Excel 2010–365，32/64 位双支持：所有 `Declare` 用 `#If VBA7` + `PtrSafe` + `LongPtr` 包裹
- WPS：`clsActionDef.SupportedInWps = False` 的工具在 WPS 下自动灰显，而不是点了才报错
- `IRibbonUI` 指针在 VBA 工程重置后会失效，`modRibbon` 用 Name + `CopyMemory` 持久化 `ObjPtr` 还原

---

## 目录

```
src/code/Core/      基础设施（执行管线、撤销、高速模式、区域、字符串、IO、参数、事件）
src/code/Text/      M1 单元格/文本处理
src/code/Data/      M2 数据处理
src/code/Sheet/     M3 工作表/工作簿管理
src/code/Merge/     M4 多文件合并
src/code/File/      M5 文件批处理
src/code/Formula/   M6 公式与引用工具
src/code/Audit/     M7 数据体检/清洗报告
src/code/Viz/       M8 数据可视化
src/code/Misc/      M9 辅助增强（聚光灯、身份证、金额大写等）
src/loader/code/    瘦加载器（内网自动更新，独立构建，不含业务功能）
```

没有 `Forms/`：`.frm` 会带一个 `.frx` 二进制资源文件，放进"源码入 git、构建时导入"的流程里
既不能 diff 又容易损坏。需要参数的工具统一走 `modPrompt`——交互时弹输入框，静默模式下从
测试预设取值，两条路径走的是同一份业务代码。

完整功能清单见 [docs/功能清单.md](docs/功能清单.md)。
