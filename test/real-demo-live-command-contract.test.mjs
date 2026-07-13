import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const packageJson = JSON.parse(
  await readFile(new URL('../package.json', import.meta.url), 'utf8'),
);
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8');

const command = packageJson.scripts['demo:live'];
assert.equal(
  command,
  'node scripts/monitor-real-demo.mjs --prepare --auto-mark --timeout-ms=900000 --report-file=data/live-demo-report.json',
);
assert.match(command, /--prepare/);
assert.doesNotMatch(command, /--open-auth|--wait-auth|--scan-after-auth/);
assert.match(command, /--auto-mark/);
assert.match(command, /data\/live-demo-report\.json/);
const authCommand = packageJson.scripts['demo:live:auth'];
assert.match(authCommand, /--open-auth/);
assert.match(authCommand, /--wait-auth/);
assert.match(authCommand, /--scan-after-auth/);
assert.match(readme, /npm run demo:live/);
assert.match(readme, /npm run demo:live:auth/);
assert.match(readme, /data\/live-demo-report\.json/);
assert.match(readme, /直接开启飞书会议/);
assert.match(readme, /开放标注接口/);

console.log('ok real demo live command contract');
