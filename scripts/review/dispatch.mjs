import { existsSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { spawn } from 'node:child_process';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const reviewDir = resolve(root, 'review');
const requestPath = resolve(reviewDir, 'request.json');
const resultPath = resolve(reviewDir, 'result.json');
const models = (process.env.CODEX_REVIEW_MODELS || process.env.CODEX_REVIEW_MODEL || 'gpt-5.6-sol,gpt-5.5')
  .split(',')
  .map((value) => value.trim())
  .filter(Boolean);

if (!existsSync(requestPath)) throw new Error('找不到 review/request.json，请先创建复审请求。');
const request = JSON.parse(readFileSync(requestPath, 'utf8'));
const requestArgIndex = process.argv.indexOf('--request-id');
const requestedId = requestArgIndex >= 0 ? process.argv[requestArgIndex + 1] : '';
if (requestedId && requestedId !== request.requestId) {
  throw new Error(`请求 ID 不匹配：当前为 ${request.requestId}，收到 ${requestedId}`);
}

const startedAt = new Date().toISOString();
writeResult({
  requestId: request.requestId,
  commit: request.commit,
  status: 'running',
  startedAt,
  finishedAt: startedAt,
  model: models[0],
  summary: 'Codex 正在复审。',
  findings: [],
  rawOutput: '',
  exitCode: null,
});

const prompt = [
  `复审请求 ${request.requestId}。`,
  `只检查提交 ${request.commit}。`,
  request.instructions,
  '请最后明确给出 PASS 或 FAIL。列出每个问题的文件和行号。不要修改文件，不要提交，不要推送。',
].join('\n');

let currentModel = models[0];
runWithFallback(0);

function runWithFallback(index) {
  const model = models[index];
  if (!model) {
    finish('error', '所有配置的 Codex 模型都不可用。', '', null);
    return;
  }
  // win32 下用 codex.cmd（不是 codex.ps1）——.ps1 不在 PATHEXT 默认列表里，
  // 直接当命令跑不起来。不带 shell:true 直接 spawn .cmd 会报
  // "spawn EINVAL"（真实跑过验证），Windows 下执行 .cmd 必须过 shell；
  // 但 shell:true 时 Node 明确放弃参数转义、只做字符串拼接
  // （DEP0190），多行/带空格的 prompt 直接放进 argv 会被 cmd.exe 按
  // 空白重新分词，报 "unexpected argument '<prompt 里的某个词>' found"
  // （也真实跑过验证）。两难之下：commit sha 和 model 名字本身不含
  // 空格/shell 特殊字符，可以留在 argv 里；prompt 改成用 codex 自己
  // 支持的 "PROMPT 传 - 表示从 stdin 读" 这条路，整段文本走 stdin
  // 管道，不再经过 cmd.exe 的命令行分词，从根上避开转义问题。
  // 另外 -C/--cd 是 codex 顶层参数，必须出现在子命令 exec review 之前，
  // 放在后面会被 clap 拒绝——cwd 已经靠 spawn 的 cwd 选项设置了，
  // 这里不需要再传一次。
  // --commit 和自定义 PROMPT 互斥（真实跑过验证：codex 报
  // "the argument '--commit <SHA>' cannot be used with '[PROMPT]'"）。
  // prompt 文本里已经写了"只检查提交 X"，靠 codex 自己在 cwd 里
  // 跑 git 命令去看那次提交，不用 --commit 这个选项。
  const codexCommand = process.platform === 'win32' ? 'codex.cmd' : 'codex';
  const child = spawn(codexCommand, [
    'exec', 'review',
    '--model', model, '--ephemeral', '--json', '-',
  ], {
    cwd: root,
    windowsHide: true,
    shell: process.platform === 'win32',
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  child.stdin.end(prompt);

  let stdout = '';
  let stderr = '';
  child.stdout.on('data', (chunk) => { stdout += chunk.toString(); });
  child.stderr.on('data', (chunk) => { stderr += chunk.toString(); });
  child.on('error', (error) => {
    if (index + 1 < models.length) return runWithFallback(index + 1);
    finish('error', `无法启动 Codex：${error.message}`, stderr || stdout, null);
  });
  child.on('close', (code) => {
    const output = `${stdout}${stderr ? `\n${stderr}` : ''}`.trim();
    const normalized = output.toLowerCase();
    const unsupported = /not supported|unsupported model|模型.*不支持|不支持.*模型/.test(normalized);
    if (unsupported && index + 1 < models.length) return runWithFallback(index + 1);

    const failed = code !== 0 || /\b(fail|blocking|阻塞|未通过)\b/.test(normalized);
    const passed = !failed && /\b(pass|passed|通过|approved)\b/.test(normalized);
    currentModel = model;
    finish(
      passed ? 'passed' : failed ? 'failed' : 'error',
      passed ? 'Codex 复审通过。' : failed ? 'Codex 复审发现问题。' : '无法从 Codex 输出中确定复审结论。',
      output,
      code,
    );
  });
}

function writeResult(result) {
  const temp = resolve(reviewDir, 'result.json.tmp');
  writeFileSync(temp, JSON.stringify(result, null, 2) + '\n', 'utf8');
  renameSync(temp, resultPath);
}

function finish(status, summary, rawOutput, exitCode) {
  if (!existsSync(resultPath)) return;
  const current = JSON.parse(readFileSync(resultPath, 'utf8'));
  if (current.requestId !== request.requestId) return;
  writeResult({ ...current, status, summary, model: currentModel, rawOutput, exitCode, finishedAt: new Date().toISOString() });
}
