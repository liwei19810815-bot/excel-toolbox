import { randomUUID } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const reviewDir = resolve(root, 'review');
mkdirSync(reviewDir, { recursive: true });

const args = process.argv.slice(2);
const value = (name, fallback = '') => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 ? args[i + 1] ?? fallback : fallback;
};

const commit = value('commit', 'HEAD');
const instructions = value(
  'instructions',
  '审查当前提交的正确性、安全性、宿主兼容性和回归风险。只报告有证据的问题。',
);
const resolvedCommit = execFileSync('git', ['rev-parse', commit], { cwd: root, encoding: 'utf8' }).trim();
const request = {
  requestId: randomUUID(),
  commit: resolvedCommit,
  createdAt: new Date().toISOString(),
  repo: root,
  instructions,
  attempt: Number(value('attempt', '1')),
};

const temp = resolve(reviewDir, 'request.json.tmp');
writeFileSync(temp, JSON.stringify(request, null, 2) + '\n', 'utf8');
renameSync(temp, resolve(reviewDir, 'request.json'));

// 旧结果不能被轮询器误认成当前请求的结果。
const resultPath = resolve(reviewDir, 'result.json');
if (existsSync(resultPath)) {
  const old = JSON.parse(readFileSync(resultPath, 'utf8'));
  if (old.requestId !== request.requestId) {
    const resultTemp = resolve(reviewDir, 'result.json.tmp');
    writeFileSync(resultTemp, JSON.stringify({
      requestId: request.requestId,
      commit: resolvedCommit,
      status: 'running',
      startedAt: new Date().toISOString(),
      finishedAt: new Date().toISOString(),
      summary: '等待 Codex 复审启动。',
      model: process.env.CODEX_REVIEW_MODEL || 'gpt-5.6-sol',
      findings: [],
      rawOutput: '',
      exitCode: null,
    }, null, 2) + '\n', 'utf8');
    renameSync(resultTemp, resultPath);
  }
}

console.log(JSON.stringify(request));
