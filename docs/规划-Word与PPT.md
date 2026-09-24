# 规划：Word 与 PowerPoint 支持

> 状态：**二期进行中**（第 1 步已完成，见第五节）。三期 Word 仍是规划，未开工。
> 本文只定方向和判据，不是承诺的排期。

## 为什么当初砍掉，现在能捡回来

最初的需求是 Excel + Word + PPT，中途收窄成「本期只做 Excel」。
理由是：没有 Word/PPT 的消费者时，做宿主抽象属于过度设计，
还会把已经验证过的 Excel 代码搅乱。

现在 Excel 这一侧稳定了（60 个命令、300 条断言、帮助系统、遥测、
一键安装、AI 接入都已交付），再做抽象有了真实的第二个消费者，
时机比当初合适。

---

## 一、先看清楚能复用多少

统计了现有 8146 行 VBA 对 Excel 对象（`Worksheet`/`Workbook`/`Range`/
`ScreenUpdating`/`Cells(`/`UsedRange`）的引用次数：

| 模块 | 耦合 | 归属判断 |
|---|---|---|
| `modUndo` 50、`modRange` 36 | 极高 | **各宿主自己实现**，没有复用空间 |
| `modSheets` 47、`modText` 40、`modViz` 36 | 高 | 业务逻辑 Excel 专有 |
| `modFileBatch` 31、`modFormula` 24 | 高 | 同上 |
| `modAction` 22 | 中 | **抽出宿主调用后可进 shared** |
| `modHelp` 15、`modUtils` 18 | 中低 | 拆分：通用部分进 shared |
| `modStr` 0、`modSettings` 1、`clsActionDef` 1、`modRibbon` 3、`modApp` 3 | 极低 | **直接进 shared** |

**结论：真正能原样复用的是"框架"，不是"功能"。**

可复用（约占代码量的 1/4，但是最难写的那 1/4）：
统一执行管线、命令注册表、功能区回调、帮助系统、遥测、
静默模式与参数预设、安装与自动更新、打包脚本、整套测试基建。

**不能复用的是 60 个命令本身**——它们全是 Excel 语义
（单元格、区域、工作表、公式）。Word 和 PPT 要的是另一套命令。

---

## 二、三个宿主的硬差异

### 撤销：三种完全不同的机制

| 宿主 | 机制 | 后果 |
|---|---|---|
| **Excel** | 自建快照框架（`modUndo`，50 处耦合） | VBA 一执行，原生 Ctrl+Z 永久失效，只能自建 |
| **Word** | **原生 `Application.UndoRecord`** | `StartCustomRecord(name)` / `EndCustomRecord` 把整组操作合成**一条原生撤销记录**，用户直接按 Ctrl+Z。**比 Excel 简单，体验也更好** |
| **PowerPoint** | **没有任何机制** | 无 UndoRecord，VBA 改动基本无法撤销 |

PPT 的结论是硬的：**所有命令一律 `Undoable:=False` + `ConfirmBeforeRun`**，
UI 上如实标「不可撤销」。`clsActionDef` 现有的三个标志
（`Undoable` / `ConfirmBeforeRun` / `PromptsForInput`）和自动生成的
「【可撤销】/【不可撤销】」提示**不用改**就能承载这三种策略。

### 构建产物

| 宿主 | COM | 产物 | SaveAs 格式值 |
|---|---|---|---|
| Excel | `Excel.Application` | `.xlam` | `xlOpenXMLAddIn` = 55 |
| Word | `Word.Application` | `.dotm` | `wdFormatXMLTemplateMacroEnabled` = 13 |
| PowerPoint | `PowerPoint.Application` | `.ppam` | `ppSaveAsOpenXMLAddin` = 30 |

三者都是 OOXML zip，`customUI` 注入代码可以原样复用。

### 「信任对 VBA 工程对象模型的访问」是**逐宿主独立**的开关

Excel 开了不代表 Word / PPT 开了。构建脚本要**分别检测**并给出
针对该宿主的提示。这一条不做的话，表现是"构建脚本在 Word 上莫名失败"。
（已实测踩过一次：本机 Excel 开了这个设置，PowerPoint 没开，直到用户
手动去 PowerPoint 的信任中心单独勾选才打通。）

### PowerPoint COM 自动化的实测差异（用真实 COM 调用逐条验证过）

这些都不是从文档推断的，是直接拿这台机器的 PowerPoint 16.0 通过
PowerShell COM 自动化实测出来的，构建脚本要按这些写，别照抄 Excel 那套：

- **`Application.Visible` 不能设为 `False`**——PowerPoint 直接抛异常
  "Hiding the application window is not allowed"。Excel 的
  `build.ps1` 里 `$xl.Visible = $false` 那种无界面构建方式在 PPT 上不成立，
  改用 `$p.WindowState = 2`（`ppWindowMinimized`）退而求其次。
- **`Application.EnableEvents` 属性不存在**——PowerPoint 的 Application
  对象没有这个成员，构建脚本里对应那一行要整个跳过，不是改成别的值。
- **`Application.DisplayAlerts` 是枚举不是布尔**——`PpAlertLevel`：
  `ppAlertsNone = 1`、`ppAlertsAll = 2`，不能像 Excel 那样直接赋
  `$false`/`$true`。
- **`.ppam` 不能用 `Presentations.Open()` 打开**——会抛
  "You must use Addins.Add to load Addin files"。必须用
  `Application.AddIns.Add(path)` 拿到 `AddIn` 对象，再设
  `.Loaded = $true`（这一点和 Excel `xl.AddIns.Add(...).Installed = $true`
  的套路一致，`install\Install-Toolbox.ps1` 里就是这么给 Excel 用的）。
- **新建的 `Presentation.VBProject` 默认零组件**——不像 Excel 的
  `ThisWorkbook` 那样自带一个文档模块，`build.ps1` 里"文档模块不能
  Import，只能塞 CodeModule"那一段特殊处理，PPT 这边不需要。
- **PowerShell 里 `$app.Run(...)` 直接点调用会失败**——`Application.Run`
  在 PowerPoint 的类型库里签名是
  `Run(string MacroName, [ref] Params Object[] safeArrayOfParams)`，
  这个 `ParamArray` 签名在 PowerShell 的late-bound 点调用下解析不出重载，
  报"找不到 Run 方法或属性"。必须改用
  `$app.GetType().InvokeMember('Run', [Reflection.BindingFlags]::InvokeMethod, $null, $app, @('MacroName'))`
  绕过 PowerShell 自己的 COM 绑定，VBA 代码内部互相调用不受影响，
  这纯粹是 PowerShell 测试脚本这一层的坑。
- **`%APPDATA%\Microsoft\Addins` 已经是 PowerPoint 的信任位置**——
  和 Excel 的 `%APPDATA%\Microsoft\AddIns`（大小写不同、注意是两个
  不同目录）是同一个思路，装 PPT 加载项应该放这里。

### 已解决：ribbon 版 .ppam 必须完全重建 zip，不能用 ZipFile 的 Update 模式原地改

**根因和修法**（排查过程见下，结论先说）：`build.ps1` 里 `Add-CustomUI`
函数用的是 `[System.IO.Compression.ZipFile]::Open(path, "Update")`
原地打开已有 zip、删旧条目、加新条目、改 `_rels/.rels`。这个写法
Excel 的 `.xlam` 完全能接受（用了很多轮，一直没事），但 PowerPoint
对 `.ppam` 的包完整性检查明显更严格——同样的 Update 模式产出的包，
PowerPoint 一律拒绝加载（自动化下表现为 `.Loaded = True` 卡死不返回，
手动通过「文件→选项→加载项」界面操作则表现为明确的
"抱歉，由于某种原因，PowerPoint 无法加载...加载项"）。

**修法**：不在原 zip 上做原地更新，改成**完全展开到临时目录 → 加文件 →
改 `_rels/.rels` → 从目录重新打包成新 zip**（`ExtractToDirectory` +
`CreateFromDirectory`，而不是 `ZipFile.Open(..., "Update")`）。
同样的 customUI 内容，这样打包出来的 `.ppam` 真机验证**加载成功、
功能区选项卡和按钮都正常显示**。

排查过程记录（怎么一步步定位到这的，供类似问题参考）：

- 用最小化 customUI14.xml 注入后，`AddIns.Add(path).Loaded = True`
  自动化调用会无限期挂住（实测等过 170 秒以上），`EnumWindows` 查不到
  任何可见对话框——一开始怀疑是自动化环境本身的问题（消息泵、STA
  重入之类）。
- 改成让用户在真机上手动走「文件→选项→加载项」操作，**同样的文件在
  手动操作下会明确报错**（不是挂死）——证明问题出在文件本身，
  不是自动化脚本的锅。
- 用不带功能区的 `.ppam`（A 组）秒开成功；带功能区但去掉 `onLoad`
  回调的 `.ppam`（B 组）依然报同样的错——把问题精确定位到"文件里
  只要带 customUI 关系条目就失败"，和回调代码无关。
- 怀疑 `[Content_Types].xml` 缺 `customUI14.xml` 的显式 Override
  声明（C 组：补上这条声明），**依然报错**，排除。
- 怀疑 Group Policy 限制了 PowerPoint 的功能区自定义（Click-to-Run
  安装常见），检查过 `HKCU/HKLM\Software\Policies\Microsoft\Office`
  下没有任何 PowerPoint/ribbon 相关策略键，排除。
- 最后怀疑 `ZipFile` 的 `Update` 模式本身留下了某种瑕疵（可能是
  压缩方式、条目顺序或 central directory 元数据），改成完全展开
  重新打包（D 组）——**真机验证通过，选项卡和按钮都正常出现**。

这个坑目前只在 `.ppam` 上验证过是真问题；`build.ps1` 给 Excel 用的
`Add-CustomUI`（Update 模式）暂不改动，因为它对 `.xlam` 一直工作正常，
没有回归的必要——PPT 这边的 `Add-CustomUI` 单独实现，用重建 zip 的写法。

---

## 三、怎么拆

### `modHost`：同名模块，各宿主一份实现

VBA 没有接口，所以用「每个宿主工程里有一个同名的 `modHost`」达成同样效果。
构建时 `src/shared/code` + `src/<host>/code` 一起导入，
`modAction` 只调 `modHost.*`，不再直接碰 `ActiveWorkbook` / `Selection` / `ScreenUpdating`。

契约（`modAction` 依赖的全部）：

```
Host_Kind() / Host_Version() / Host_Bitness()
Host_HasDocument()         Excel: ActiveWorkbook  Word: ActiveDocument  PPT: ActivePresentation
Host_HasValidSelection()   Excel: TypeName(Selection)="Range"，Word/PPT 各自判定
Host_FastModeOn/Off/Reset  Excel: ScreenUpdating+Calculation+EnableEvents
                           Word:  ScreenUpdating       PPT: 基本无可关
Host_SetStatus / Host_ClearStatus
Host_BeginUndo(name) / Host_CommitUndo() / Host_RollbackUndo() / Host_CanUndo() / Host_PeekLabel()
```

> **改造 `modAction.RunAction` 时不要动管线的顺序与配对**。
> 那段的注释已写明 `FastModeOn` 必须最外层、`BeginTx` 必须在 `Dispatch` 之前。
> 只把对象调用换成 `modHost.*`。

### 目录

```
src/
  shared/code/   modAction modRibbon modApp clsActionDef modStr
                 modSettings modPrompt modTelemetry modHelp
  excel/code/    modHost + 现有全部业务模块
  excel/package/ customUI14.xml
  word/code/     modHost + Word 业务模块
  ppt/code/      modHost + PPT 业务模块
```

---

## 四、Word / PPT 做哪些命令

**不要照搬 Excel 的 60 个。** 换个宿主，用户的痛点完全不同。

### Word（建议首批 10–12 个）

排版与清理类最有价值——和 Excel 侧「清除空格 / 数据体检」是同一类需求：

- 清理多余空行、多余空格、全半角混排
- 统一字体字号、统一段落间距与缩进
- 批量替换（含正则）
- 表格：统一表格样式、自动调整列宽、表头重复
- 图片：统一尺寸、居中、加题注
- 生成/更新目录、更新所有域
- 批注与修订：一键接受全部、导出批注清单
- 文档体检：字体不统一、空段落、手动换行符、超长段落

### PowerPoint（建议首批 8–10 个）

- 统一字体、统一母版/版式
- 批量替换文本
- 对齐与分布（选中对象）
- 导出所有备注为文本
- 图片批量压缩/统一尺寸
- 检查：超出版心的对象、字号过小、空占位符

> PPT 全部不可撤销，所以**破坏性命令要更少、确认要更强**。
> 只读的「检查」类在 PPT 上价值更高，优先做。

---

## 五、实施顺序与完成判据

> 项目改成三期路线：一期 Excel（已完成）、二期 PPT、三期 Word——
> 不是本节最初写的"三宿主一起做"。下表按二期 PPT 优先重排，
> Word 相关步骤挪到最后，留给三期。

| 步 | 内容 | 判据 | 状态 |
|---|---|---|---|
| 1 | 源码重组 + `modHost` 抽象（仅 Excel） | **Excel 现有断言全绿**（这是不能退的底线） | **已完成**：`src/shared/code` + `src/excel/code` 两分区，`modHost.bas` 包一层 `modPerf`/`modUndo`，409 条断言零回归（提交 `08c10c5`）。比原计划更保守——`modAction`/`modRibbon`/`modPrompt` 等实测仍有真实 Excel 耦合，没有强行塞进 shared |
| 2 | PPT 构建管线（独立脚本 `build-ppt.ps1`，未采用文档最初设想的 `build.ps1 -Host` 参数化，理由见下） | 产出 `.ppam`，能构建、装载、功能区出现 | **代码已完成，核心修复已由人工真机验证，自动化验证因环境问题搁置**：`build-ppt.ps1` + `_PptHost.ps1`（独立于 Excel 那份，不共用 `_ExcelHost.ps1`）+ `src/ppt/code/Core/{modApp,modRibbon,modPublic}.bas` + `src/ppt/package/customUI/customUI14.xml`。真机人工安装验证过——功能区选项卡、按钮回调都正常（见上方"已解决"小节的排查记录，那是本步骤真正的技术产出）。`build-ppt.ps1` 本身还没有一次完整跑通的自动化记录：这台机器的 PowerPoint COM 自动化在大量创建/强杀实例后进入了不稳定状态——**已排查出部分根因**：`HKCU\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\StartupItems` 下出现了损坏的注册表项（大概率是强杀进程时写到一半留下的垃圾数据），清掉这个键后 COM 激活恢复过一次，但随后又出现新的失败（`0x800706BF` RPC 调用失败），**重启电脑也没能根治**，Event Viewer 里既没有 PowerPoint 的崩溃事件、也没有对应 CLSID 的 DCOM 注册失败记录，说明失败发生在 Office Click-to-Run 自己的激活层，标准 Windows 诊断工具看不到。脚本本身在"启动 PowerPoint"这一步之前的全部 PowerShell 语法/路径逻辑已反复验证无误。**下次有机会时**：先试 Office 快速修复（设置→应用→Microsoft Office→修改→快速修复），如果还不行再考虑完整修复或重装；跑通后回来补一次 `build-ppt.ps1` 的自动化验证记录，不是阻断性的——人工验证已经证明代码是对的 |
| 3 | PPT 样板命令 3–5 个（只读检查类优先） | 闭环：构建 → 装载 → 功能区 → 执行；撤销一栏标「不可撤销」并强制确认 | 未开工 |
| 4 | PPT 命令集补齐 | `check-help`、`check-ribbon`、`check-imagemso` 针对 PPT 全绿 | 未开工 |
| 5 | PPT AI（Excel AI 仓库） | 任务窗格怎么接入 PPT——沿用同一 manifest 按 `Office.context.host` 分流，还是独立产品线，留到 PPT 工具箱有实际命令后再定 | 未开工，架构未定 |
| 6（三期） | Word 构建管线 + 样板命令 | 闭环同上，撤销走**原生 `Application.UndoRecord`**（比 Excel 简单） | 留给三期，不在本期范围 |

### 测试要新增的

- `tests/run-tests-word.ps1` / `run-tests-ppt.ps1`，沿用现有断言风格与
  静默模式 + 参数预设机制（`Toolbox_SetSilent` / `Toolbox_SetParam`）
- `check-ribbon.ps1` / `check-imagemso.ps1` 参数化到三宿主
- **`imageMso` 必须逐宿主验证**：同一个 ID 在 Word 里可能不存在，
  而 Office 对坏 ID 是**静默不画图标**。这正是 `docs/图标验收记录.md` 存在的理由，
  每个宿主要各出一份
- `_ExcelHost.ps1` → `_OfficeHost.ps1`，进程登记与回收那一整套
  （四态判定、PID+启动时间三重校验、锁文件）**原样复用**，
  只把进程名白名单从 `EXCEL/et/wps` 扩到含 `WINWORD/POWERPNT`

---

## 六、已知会踩的坑（Excel 侧已经踩过，别再踩一次）

- **不要引用高版本独有常量或 MSO 库符号**（`msoTrue`/`msoFalse`/`FileDialog`）——
  后果不是功能降级，是**整个工程编译不过**。查的时候搜 `\bmso[A-Z]` 扫前缀
- **`.bas/.cls` 必须 CRLF + 系统 ANSI**，否则 VBE 静默把类模块导成标准模块
- **功能区回调必须全部挂上**，漏挂 `getEnabled` 不会让任何测试变红
- **Word/PPT 的 COM 自动化同样可能被 WPS 劫持**，`New-Real*` 守卫要一并加
- **`AscW` 返回带符号 Integer**，U+8000 以上是负数——这个坑在 Excel 侧踩过两次，
  代价是静默删汉字。一律走 `modStr.CodePointOf`

---

## 七、工作量的诚实估计

| 部分 | 相对规模 |
|---|---|
| 1–2 步（重组 + 三宿主管线） | 最难，但一次性 |
| Word 命令集 | 中等，撤销比 Excel 简单 |
| PPT 命令集 | 命令少，但每个都要确认框 + 文案 |
| 帮助内容（每个命令一节 + 示例） | **容易被低估**，Excel 侧 60 条写了很久 |
| 逐宿主图标验收 | 必须肉眼看一遍总览图，「存在」不等于「画得出来」 |

**风险最高的一步是第 1 步**：动 `modAction` 和 `modUndo` 意味着碰
整个执行管线。判据必须是"Excel 300 条断言一条不少"，不达标就不往下走。
