import { spawn } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '../..');
const requestScript = resolve(root, 'scripts/review/request.mjs');
const dispatchScript = resolve(root, 'scripts/review/dispatch.mjs');
const requestArgs = process.argv.slice(2);

await run(process.execPath, [requestScript, ...requestArgs]);
const request = JSON.parse(readFileSync(resolve(root, 'review/request.json'), 'utf8'));
await run(process.execPath, [dispatchScript, '--request-id', request.requestId]);

function run(command, args) {
  return new Promise((resolvePromise, reject) => {
    const child = spawn(command, args, { cwd: root, windowsHide: true, stdio: 'inherit' });
    child.on('error', reject);
    child.on('close', (code) => {
      if (code === 0) resolvePromise();
      else reject(new Error(`${command} 退出码为 ${code}`));
    });
  });
}

if (!existsSync(resolve(root, 'review/result.json'))) {
  throw new Error('Codex 未生成 review/result.json。');
}
