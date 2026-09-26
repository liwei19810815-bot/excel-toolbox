# Claude ↔ Codex 自动复审桥接

这套协议用于让 Claude Code 自动发起 Codex 复审，并通过共享仓库文件返回结果。用户不需要转述复审内容。

## 快速使用

在仓库根目录运行一条命令即可：

```powershell
npm run review:run -- --commit HEAD
```

这条命令会创建请求、启动 Codex 复审，并等待结果写入共享文件。Claude Code 的自动化流程也可以拆成后台任务：

1. 提交待复审代码；
2. 调用 `review:request`；
3. 在后台重复调用 `review:poll`，直到 `status` 为 `passed`、`failed` 或 `error`；
4. `failed` 时读取 `review/result.json` 的 findings，修复后创建新的请求；
5. `passed` 时才执行推送。

需要让 Claude 在不阻塞主会话时运行时，使用下面的拆分方式：

```powershell
npm run review:request -- --commit HEAD
npm run review:dispatch -- --request-id <request-id>
npm run review:poll -- --request-id <request-id>
```

## 文件协议

- `review/request.json`：当前复审请求，由脚本生成。
- `review/result.json`：当前复审结果，由 Codex 调用脚本生成。
- `review/*.schema.json`：协议格式。
- `review/.gitignore`：运行时文件不进入 Git 提交。

请求和结果都包含 `requestId` 与 `commit`，避免 Claude 误读旧结果。结果文件写入采用临时文件替换，轮询过程中不会读到半截 JSON。

## 模型配置

默认模型回退链为 `gpt-5.6-sol,gpt-5.5`，首选模型不被账号支持时会自动尝试下一个。也可以通过环境变量覆盖：

```powershell
$env:CODEX_REVIEW_MODELS = "gpt-5.6-sol,gpt-5.5"
```

如果账号不支持该模型，改用当前账号可用的 Codex 模型。不要使用 ChatGPT 账号不支持的 `gpt-6-luna`。

## Claude 集成约定

Claude 不需要等待交互式终端。它可以在后台执行：

```powershell
node scripts/review/dispatch.mjs --request-id <request-id>
```

脚本会调用本机 `codex exec review`，把最终结果写入 `review/result.json`。Codex 只读复审，不修改工作区、不推送代码。
