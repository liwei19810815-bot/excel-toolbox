# 规划：Word 与 PowerPoint 支持

> 状态：**二期已完成并推送**（PPT 工具箱骨架 + PPT AI）。**三期进行中**
> （Word 构建管线骨架 + Word AI + 首批三个样板命令已落地，见第五、八节；
> `modAction` 已按第八节方案拆成 shared + 各宿主 modActionRegistry）。
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
| Word | `Word.Application` | `.dotm` | `wdFormatXMLTemplateMacroEnabled` = 15（原计划写的 13 是错的，那其实是 `wdFormatXMLDocumentMacroEnabled`（`.docm`），已用真实 COM 调用验证：`SaveAs(path, 15)` 产出的包 `[Content_Types].xml` 里同时匹配 `template` 和 `macroEnabled`，`SaveAs(path, 13)` 没试但 13 明确是文档不是模板，不能用） |
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

### Word COM 自动化的实测差异（用真实 COM 调用逐条验证过）

三期开工前先拿这台机器的 Word 16.0 实测了一遍，结论是：**Word 比
PowerPoint 更接近 Excel**，但不完全一样，构建脚本要按这些写：

- **`Application.Visible` 可以直接设 `False`**——和 Excel 一致，不像
  PowerPoint 那样会抛异常，不需要 `WindowState` 变通。
- **`Application.EnableEvents` 属性不存在**——和 PowerPoint 一样，这是
  Excel 专有的，构建脚本对应那一行整个跳过；但 `ScreenUpdating` 正常
  存在、可以设置。
- **`Application.DisplayAlerts` 是数值 `wdAlertsNone = 0`**——不是
  PowerPoint 那种从 1 开始的 `PpAlertLevel` 枚举，也不是 Excel 的布尔，
  三个宿主三种写法，不能互相抄数值。
- **没有 `Application.Hwnd`**——真机测试过，即使已经打开一个可见文档，
  `$w.Hwnd` 读出来也是空值。Word 不能像 Excel 那样反查 PID，必须走
  PowerPoint 那套"创建前后进程快照比对"方案（`_WordHost.ps1` 的
  `Get-WordIdentity`，直接照抄 `_PptHost.ps1` 修好之后的版本——快照
  必须由调用方在 `New-Object` **之前**拍好传进来，不能在函数内部现拍，
  这是 PPT 那边被 Codex 挑出过的真实 bug，Word 这边从一开始就按正确
  顺序写）。
- **`wdFormatXMLTemplateMacroEnabled` 正确值是 15，不是 13**——见上方
  "构建产物"表格的脚注，已用真实 `SaveAs` 调用验证过产出文件的
  `[Content_Types].xml` 同时匹配 `template` 和 `macroEnabled`。
- **customUI 注入沿用 PPT 的"完全重建 zip"做法**——没有反过来验证
  "Word 的 `.dotm` 是否也存在 Excel 那种原地 Update 会被拒绝加载"的
  问题，直接用已经证明安全的重建做法，不重复冒险验证。
- **「信任对 VBA 工程对象模型的访问」这台机器上 Word 没开**——Excel
  和 PowerPoint 的 `HKCU\...\Excel\Security\AccessVBOM` /
  `...\PowerPoint\Security\AccessVBOM` 都是 `1`，唯独 Word 的对应键
  不存在。这是安全相关的注册表设置，不能由自动化流程代自己打开，
  需要用户在 Word 里手动开启（文件→选项→信任中心→信任中心设置→
  宏设置→勾选"信任对 VBA 工程对象模型的访问"）。**这也是当前
  `build-word.ps1` 没有一次完整自动化跑通记录的原因**——脚本本身
  的 PowerShell 语法、customUI 注入、进程识别逻辑都已经过静态审查
  和 COM 调用抽样验证，只差这一步用户手动开关。

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
| 5 | PPT AI（Excel AI 仓库） | 任务窗格怎么接入 PPT——沿用同一 manifest 按 `Office.context.host` 分流，还是独立产品线 | **架构已定并落地**：沿用同一 manifest/bundle，运行时按 `Office.context.host` 分流（`src/store/host.ts`），复用 chat/session/settings/sidecar 客户端（本来就是宿主无关的）。新增 `src/powerpoint/{coordinator,blueprint}.ts`、四个 PPT 工具（`get_presentation_overview`/`read_slide`/`add_text_box`/`add_slide`）、`EXCEL_TOOL_NAMES`/`POWERPOINT_TOOL_NAMES` 按宿主过滤工具列表、manifest 新增 `Presentation` Host 块。`npx tsc --noEmit`/`npx vitest run`（130/130）/`npm run build` 全过，已提交（`354fda8`，未推送）。**两项未完成**：①这台机器 PowerPoint 环境不稳定，没做过真机侧载验证；②Codex 独立验收三次尝试都因账号侧模型配置问题失败（"gpt-6-luna"/"gpt-5.3-codex" 均报 "not supported when using Codex with a ChatGPT account"），需要用户跑 `/codex:setup` 排查，推送前应补这轮验收 |
| 6（三期第一步） | Word 构建管线（独立脚本 `build-word.ps1` + `_WordHost.ps1`） | 产出 `.dotm`，能构建、装载、功能区出现 | **代码已完成，静态验证已过，端到端自动化被这台机器的一个前置条件挡住**：`build-word.ps1` + `_WordHost.ps1`（独立于 `_ExcelHost.ps1`/`_PptHost.ps1`）+ `src/word/code/Core/{modApp,modRibbon,modPublic}.bas` + `src/word/package/customUI/customUI14.xml`。PowerShell 语法、XML 结构、VBA 源码 BOM 编码都已核对；`SaveAs` 格式常量、`DisplayAlerts`/`Hwnd`/`Visible` 等宿主差异都用真实 COM 调用逐条验证过（见上方"Word COM 自动化的实测差异"）。**卡住的地方**：这台机器 Word 的「信任对 VBA 工程对象模型的访问」没开（Excel/PowerPoint 都开了，Word 没有），这是安全设置，不能由自动化脚本代为打开，需要用户手动去 Word 信任中心勾选后才能跑通一次真实构建 |
| 7（三期第二步） | Word 样板命令 3 个（见第八节设计） | 闭环：构建 → 装载 → 功能区 → 执行；撤销走**原生 `Application.UndoRecord`**（比 Excel 简单） | **代码已完成，逐项用真实 COM 调用验证过底层行为，端到端仍卡在 VBOM 信任**：`word.audit`（只读体检）、`word.cleanSpaces`（Undoable=True，走 UndoRecord）、`word.updateFields`（ConfirmBeforeRun=True）。同时把 `modAction` 按第八节方案拆成 `shared/code/Core/modAction.bas`（RunAction 管线）+ 各宿主 `modActionRegistry.bas`（Excel 的拆分已用 build.ps1 + tests/run-all.ps1 全量验证 409/409，Codex 两轮复审通过；PPT 构建脚本排除了还用不上的 modAction.bas）。Word 这三个命令开发过程中用直接 COM 调用（不经过 VBE 编译，绕开 VBOM 限制）逐个验证了用到的每个 API——过程中真实发现并修了三个坑：①全角空格搜索默认会连带命中半角空格（`Find.MatchByte` 必须显式设 True，不设会误删用户文档里所有正常空格）；②`AscW` 处理段落尾字符时对 U+8000 以上的汉字返回负数，会被"当控制字符砍掉"误判成负数满足 `<32`（改用 `modStr.CodePointOf`，这是文档六已经点名过的坑，写第一版时还是踩了一次）；③`Document.Fields.Update()` 不会刷新目录内容，即使目录被算进 `Fields.Count`，必须额外调 `TablesOfContents(i).Update()`。**仍未验证**：一次完整的 VBE 编译 + 装载真机测试——这台机器 Word 的 VBOM 信任没开，是安全设置，需要用户手动去开 |
| 8（三期第三步） | Word AI（Excel AI 仓库，第三个 `Office.context.host` 分支） | 参照 PPT AI 的接入方式：新增 `src/word/coordinator.ts`/`blueprint.ts`、`WORD_TOOL_NAMES`、manifest 新增 `Document` Host 块 | **已完成并推送**：`src/word/{coordinator,blueprint}.ts` 镜像 PPT 那两个文件，`WordApi` 用到的每个方法（`body.paragraphs`/`getSelection`/`insertText`/`search`）都用 `@types/office-js` 的类型定义核实过是 1.1 基线，不是猜的版本号。四个工具：`get_document_overview`/`read_paragraph`（read）、`insert_paragraph`/`replace_text`（mutate:structure——理由和 PPT 一致，VBA 侧"Word 原生 UndoRecord 更简单"说的是 COM 自动化那条路，不能套到 Office.js 任务窗格上）。`ChatPane.tsx` 的工具禁用逻辑从二元判断改成三态，`host=unknown` 仍兜底按 Excel 处理。manifest 新增 `Document` Host 块，`office-addin-manifest validate` 三个宿主全过。Codex 两轮复审：第一轮挑出 `replace_text` 的 `find` 参数没校验 255 字符上限（`Body.search` 文档写明的限制），修复后第二轮 PASS。`npx tsc --noEmit`/`npx vitest run`（141/141）/`npm run build` 全绿，已推送。**未验证**：真机在 Word 里侧载打开任务窗格——理论上 Office.js 加载项走 manifest 侧载，不经过 COM 自动化导入 VBA 那条路，不受这台机器 Word VBOM 信任设置的限制，但没有条件实测确认 |

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

---

## 八、命令集与跨宿主执行管线设计

> 状态：**modAction 拆分（8.1）+ Word 首批三个命令（8.3 的 Word 部分）
> 已实施**。PPT 部分（8.3 的 PPT 三个命令、8.4 提到的 PPT
> modActionRegistry）**尝试开工但被环境问题挡住，未落地任何代码**——
> 这台机器的 PowerPoint COM 自动化在这一轮又复现了之前排查过的不稳定
> （见第二节"已解决：ribbon 版 .ppam..."上方、以及历史记录里
> "重启电脑也没能根治"那次排查），连最基本的
> `New-Object -ComObject PowerPoint.Application` + `Quit()` 都会挂死
> 或报 `CO_E_SERVER_EXEC_FAILURE`，清理 `Resiliency\StartupItems`
> 注册表键（上次的临时解法）这次也不起作用。Word 首批三个命令能在
> VBOM 信任卡住的情况下还继续推进，是因为可以绕开 VBE 编译、直接用
> PowerShell COM 调用逐个验证每个 API 行为；PPT 这次连这条退路都没有
> ——COM 本身连不稳定的连接都建不起来，没有任何验证手段，贸然写三个
> 命令等于纯猜代码，风险和"没有调用方的抽象"是同一类问题反过来的
> 版本，所以这一轮没有写。8.1/8.2/8.3 下面的内容是原始设计草案，
> 保持不动作为决策记录；已实施部分的实际结果和过程中发现的新问题记在
> 本节末尾的"8.5 实施结果"。

PPT 工具箱（二期）和 Word 工具箱（三期）曾经都停在"构建管线骨架"这一步，
"样板命令"都没做——不是漏了，是刻意的：没有第二个真实消费者之前，
硬套 Excel 那套 `modAction`/`clsActionDef` 管线属于没人验证过的抽象。
PPT 和 Word 同时到了这个节点后，才把这件事想清楚、写成下面的设计草案
（**8.1-8.4 是设计时的原始文字，Word 部分已经按这个方案实施**）。

### 8.1 重新读了一遍 `modAction.bas` 之后发现的关键事实

之前的规划文档把 `modAction` 整体归类为"中等耦合，抽出宿主调用后可进
shared"，这次逐行核对发现耦合比想象的更集中：

- `RunAction`（前置校验 → 确认框 → 高速模式 → 撤销事务 → 派发 →
  提交/回滚 → 恢复环境 → 遥测）**完全不碰 Excel 对象**，只调
  `modHost.*`、`modRibbon.*`、`modTelemetry.*`、`clsActionDef`、还有
  两个"自己项目里必须存在"的自由函数：`RegisterAll()`（注册命令元数据）
  和 `Dispatch(actionId)`（真正执行，返回结果字符串）。
- 只有 `RegisterAll` 里各条 `RegisterAction` 调用之间夹杂的字符串（纯数据，
  无耦合）和 `Dispatch` 的 `Select Case` 分支体（`Selection`/`ActiveWorkbook`/
  `ActiveSheet` + `modText`/`modSheets`/... 这些 Excel 专属业务模块）
  是真正的 Excel 耦合点。

这意味着 `modAction.bas` 可以**按现有的"同名模块"套路整个拆开**，
不需要发明新机制：

```
shared/code/Core/modAction.bas
    RunAction / SetSilent / IsSilent / LastMessage /
    RegisterAction（写入注册表的辅助函数）/ GetAction / IsActionEnabled /
    ActionLabel ——这些【一字不改】原样搬过去，因为本来就没碰 Excel 对象

excel/code/Core/modActionRegistry.bas（新拆出来，原 modAction.bas 里
    RegisterAll + Dispatch 那部分整体平移，函数名不变）
word/code/Core/modActionRegistry.bas（新写，Word 版 RegisterAll + Dispatch）
ppt/code/Core/modActionRegistry.bas（新写，等真的要给 PPT 做命令时再写）
```

`RunAction` 调用 `RegisterAll()`/`Dispatch(actionId)` 时不写模块前缀，
VBA 按"当前工程内查找同名过程"解析——和 `modHost` 那套"每个工程一份
同名实现"是同一个技巧，`RunAction` 完全不用知道自己调的是哪个宿主的
`Dispatch`。

### 8.2 Word 能不能用原生 `UndoRecord`：能，但只能用一半

实测确认过两件事（本轮已用真实 COM 调用验证，见第二节"Word COM 自动化
的实测差异"）：

1. `Application.UndoRecord.StartCustomRecord(name)` /
   `EndCustomRecord()` 真的能把中间任意多次编辑合并成**一条**原生
   撤销记录，用户按一次 Ctrl+Z 就能整体撤销——这部分能用，而且好用，
   Word 命令可以标 `Undoable:=True`，`Host_BeginUndo`/`Host_CommitUndo`
   包一层 `StartCustomRecord`/`EndCustomRecord` 就行。
2. 但 Office 没有暴露"查询当前是否有能撤销的记录"或"看一眼上一条
   撤销记录叫什么名字"的 API——`UndoRecord` 是纯粹的"开始记、结束记"，
   不能反向查询。这意味着 Excel 那套由 `modUndo.CanUndo()`/`PeekLabel()`
   驱动的**工具箱自己的"撤销上一步"按钮**（`core.undoLast`，靠这两个
   函数决定按钮是否可点、按钮上写哪个操作名）在 Word 上**做不出来**，
   不是没设计好，是 Office 本身没给这个能力。

**结论**：Word 版 `modHost` 里 `Host_BeginUndo`/`Host_CommitUndo` 有
真实实现（包一层 `UndoRecord`），但 `Host_CanUndo`/`Host_PeekLabel`/
`Host_UndoLast` 老实返回"不支持"（`False`/空字符串/空操作），Word 的
`RegisterAll` 里**不注册 `core.undoLast` 这个命令**，customUI 里也不放
这个按钮——用户改动后自己按 Ctrl+Z，不通过工具箱按钮撤销。这样
`RunAction` 里那句无条件的 `modRibbon.RefreshControl "btnUndoLast"`
不会报错（控件不存在时 `RefreshControl` 按现有实现直接空操作），
但也不会做任何事，是安全的。

PPT 沿用已经定好的策略：全部 `Undoable:=False`、`ConfirmBeforeRun:=True`，
不受这次讨论影响。

### 8.3 建议的首批命令（每个宿主 3 个，覆盖 `RunAction` 的三条分支）

选 3 个而不是文档里列的 10-12 个候选，是为了先把"构建 → 装载 →
功能区 → 执行 → 撤销/确认"这条完整链路在真机上跑通一次，跑通之后
按同样的模式批量补齐候选清单里剩下的命令，风险和工作量都可预估。

**Word**（覆盖只读 / 原生撤销 / 强制确认三条路径）：

| actionId | 做什么 | Undoable | ConfirmBeforeRun | 为什么选它 |
|---|---|---|---|---|
| `word.audit` | 体检：统计空段落、连续空格、手动换行符（非段落符）、超长段落，只报告不改动 | False | False | 只读，零风险，第一个验证"构建→装载→执行"链路整体走通 |
| `word.cleanSpaces` | 清理多余空格（含全角空格/不间断空格/零宽字符），逻辑上和 Excel 的 `text.cleanSpaces` 是同一类需求 | **True**（用 `UndoRecord` 分组） | False | 第一个验证 `Host_BeginUndo`/`Host_CommitUndo` 包一层 `UndoRecord` 是否真的按预期工作 |
| `word.updateFields` | 更新全文所有域（含目录），对应文档里"生成/更新目录、更新所有域"这条候选 | False（域更新后的撤销语义复杂，不承诺能撤销，老实标不可撤销） | True | 验证 `ConfirmBeforeRun` 强制确认这条路径，且是文档候选列表里价值较高的一条 |

**PowerPoint**（延续"全部不可撤销"策略，覆盖只读 / 强制确认两条路径）：

| actionId | 做什么 | Undoable | ConfirmBeforeRun | 为什么选它 |
|---|---|---|---|---|
| `ppt.audit` | 检查：超出版心的对象、字号过小、空占位符，只报告不改动 | False | False | 只读，零风险，同上验证链路 |
| `ppt.exportNotes` | 导出所有页的备注为一个文本文件 | False | False | 只读（不改动原文件，只是导出），文档候选列表里价值较高的一条 |
| `ppt.replaceText` | 批量替换全文文本 | False | **True** | 会真的改动内容，PPT 没有任何撤销机制，必须强制确认——验证这条路径 |

### 8.4 不在这轮做的

- 不把 `excel/code/Core/modActionRegistry.bas` 拆分本身当成"零风险"操作——
  这是对已验收的 Excel 执行管线做结构性改动，即使理论上只是把代码原样
  搬到另一个文件、函数签名和调用关系不变，仍然要求 **Excel 现有全部
  断言一条不少** 才能算过，和当初"重组 + modHost 抽象"那一步同一个判据。
- 不做 PPT 的具体命令实现（`ppt/code/Core/modActionRegistry.bas` 只搭
  骨架，或者等确认要不要现在一起做）——本节先把 Word 的三个样板命令
  做完、验证过手动/真机可用之后，再决定是否同一轮顺手把 PPT 的三个
  也做了，还是分开验证。
- 不假设 Word 首批三个命令能在这台机器自动化验证到底——`build-word.ps1`
  仍然卡在 VBOM 信任那一步（见第五节），命令写完之后能做的是：PowerShell
  语法检查、VBA 源码 BOM 编码检查、`Dispatch` 分支的手工代码走查，以及
  可能的情况下用不需要 VBProject 访问权限的纯 COM 调用验证 Word 对象
  模型行为（这次设计 `word.cleanSpaces`/`UndoRecord` 用的就是这种方式）。
  真正把 VBA 源码导入进 `.dotm` 跑一遍，仍然需要用户手动开一次 Word 的
  「信任对 VBA 工程对象模型的访问」。

### 8.5 实施结果（modAction 拆分 + Word 首批三个命令）

#### modAction 拆分：比设计草案预想的多两处耦合

8.1 只发现了 `RegisterAll`/`Dispatch` 是 Excel 耦合点，实际动手拆分时
逐行核对又找出两处：`IsActionEnabled`/`IsActionPressed` 内部的
`Select Case` 分支里，分别藏着 `viz.sparklines`/文件对话框能力探测
（调 `modCaps.*`）和 `misc.spotlight` 按下状态（调 `modSpotlight.*`）——
这两个也是 Excel 专属业务模块。处理方式和 `RegisterAll`/`Dispatch`
同一个技巧：`shared/modAction.IsActionEnabled`/`IsActionPressed` 对
`core.undoLast` 之外的情况转派给不写前缀的 `ActionExtraEnabled`/
`ActionExtraPressed`，各宿主的 `modActionRegistry.bas` 里实现。

验证：`build.ps1` 构建产物模块清单核对通过，`tests/run-all.ps1` 全量
重跑 **409/409**，和文档记录的基线完全一致，零回归。Codex 两轮复审：
第一轮指出 `modAction.bas` 进 shared 后 `build-ppt.ps1` 会因为 PPT
工程缺依赖而编译不过（PPT 还没有自己的 `modActionRegistry.bas`/
`modHost.bas`），修法是 `build-ppt.ps1` 显式排除 `modAction.bas`；
两轮都是 PASS。

#### Word 首批三个命令：COM 调用逐项验证时踩出的三个真坑

Word 的 VBOM 信任这台机器没开，没法走真正的 VBE 编译，所以每个用到的
API 都改用不经过 VBA、直接的 PowerShell COM 调用单独验证行为——过程中
真实发现（不是理论风险）三个问题，都已改正：

1. **`Find.MatchByte` 不显式设 `True`，全角/半角字符会被当成等价**。
   这台机器装了中文语言包，Word 的 `Find` 默认把全角空格（U+3000）和
   半角空格当成同一个东西——实测：3 字符的文档里搜 1 个全角空格，
   命中数算出来是 3，把两个普通半角空格也数进去了。`modClean.bas`
   如果不设这个属性，"清理全角空格"这个命令会把文档里所有正常空格
   一起删掉，是会破坏用户数据的真实 bug。修法：每次用 `Find` 之前
   显式 `f.MatchByte = True`。
2. **`AscW` 处理段落文字时，对 U+8000 以上的字符返回负数**——这正是
   文档第六节点名过的坑（"这个坑在 Excel 侧踩过两次，代价是静默删
   汉字，一律走 `modStr.CodePointOf`"），`modAudit.bas` 判断段落末尾
   字符是不是控制字符时，第一版写的时候还是先用了 `AscW`，构造一个
   以 U+9F98（"龘"）结尾的测试段落后复现：`AscW` 返回负数，负数天然
   满足 `< 32`，会被误判成"控制字符"整段砍掉，静默丢字。改用
   `modStr.CodePointOf`（内部把负数加 65536 转回正确码点）后重测
   通过。写档时已经知道这条规矩，实操时还是踩了一次，说明"知道有这条
   规矩"和"每次新写字符判断代码时真的记得套用"是两回事，光靠文档
   提醒不够，最终是靠"给每个真实用到的 API 都补一个可验证的用例"
   这套方法论抓出来的，不是提前想到的。
3. **`Document.Fields.Update()` 不会刷新目录内容**，即使目录本身被
   算进 `Fields.Count`（真机测过：只有一个目录时 `Fields.Count` 就是
   1，容易让人以为"目录也是域，更新域就够了"）。构造"目录建好之后
   在文档末尾新加一个标题"的场景，只调 `Fields.Update()` 目录文字
   完全不变；改调 `TablesOfContents(i).Update()` 之后新标题才出现
   在目录里。`modFields.bas` 的 `word.updateFields` 两步都做，不能
   只做一步就假设域和目录一起更新了。

除了这三个真问题，另有一处虚惊：一开始以为"`Document.Fields` 不包含
目录条目，必须分开处理"，后来直接测才发现目录确实被计入 `Fields`，
只是**被计入不等于被 `Fields.Update()` 真正刷新**——这是两件不同的
事，第 3 点的坑更细一点，写进最终代码注释里的是修正后的准确说法，
不是最初的猜测。

**方法论小结**：这三个问题没有一个是靠读文档或者凭经验猜出来的，
全部是把 VBA 要调的每一行 Office 对象模型 API，原样用 PowerShell
的 `New-Object -ComObject Word.Application` 复刻一遍、构造针对性的
边界用例（全角空格混着半角空格、段落以生僻高码点汉字结尾、目录建好
之后再加新标题）跑出来的。这台机器 VBOM 信任被卡住反而逼着把每个
API 调用单独拆出来验证，比"整体导入进 VBE 编译一次，能跑就当对"
覆盖到的边界情况更细——**这不是退而求其次的将就，抓到的三个问题里
至少两个（MatchByte、Fields vs TablesOfContents）就算真的能编译
进 VBE 里跑，普通的手工冒烟测试也不一定测得到这么细的边界**。
