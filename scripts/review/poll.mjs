import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const requestPath = resolve(root, 'review/request.json');
const resultPath = resolve(root, 'review/result.json');
const args = process.argv.slice(2);
const i = args.indexOf('--request-id');
const expectedId = i >= 0 ? args[i + 1] : '';
const timeoutMs = Number(process.env.CODEX_REVIEW_TIMEOUT_MS || 900000);
const intervalMs = Number(process.env.CODEX_REVIEW_POLL_MS || 5000);
const started = Date.now();

if (!existsSync(requestPath)) throw new Error('找不到 review/request.json。');
const request = JSON.parse(readFileSync(requestPath, 'utf8'));
if (expectedId && request.requestId !== expectedId) throw new Error('请求 ID 与当前请求不一致。');

while (Date.now() - started <= timeoutMs) {
  if (existsSync(resultPath)) {
    const result = JSON.parse(readFileSync(resultPath, 'utf8'));
    if (result.requestId === request.requestId && result.status !== 'running') {
      console.log(JSON.stringify(result, null, 2));
      process.exit(result.status === 'passed' ? 0 : 1);
    }
  }
  await new Promise((resolvePromise) => setTimeout(resolvePromise, intervalMs));
}

console.error(`Codex 复审超过 ${timeoutMs}ms 未完成。`);
process.exit(2);
