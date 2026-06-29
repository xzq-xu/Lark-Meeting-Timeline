import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const packageJson = JSON.parse(
  await readFile(new URL('../package.json', import.meta.url), 'utf8'),
);
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8');

const command = packageJson.scripts['demo:live'];
assert.equal(
  command,
  'node scripts/monitor-real-demo.mjs --prepare --open-auth --wait-auth --scan-after-auth --auto-mark --timeout-ms=900000 --auth-timeout-ms=300000 --report-file=data/live-demo-report.json',
);
assert.match(command, /--prepare/);
assert.match(command, /--open-auth/);
assert.match(command, /--wait-auth/);
assert.match(command, /--scan-after-auth/);
assert.match(command, /--auto-mark/);
assert.match(command, /data\/live-demo-report\.json/);
assert.match(readme, /npm run demo:live/);
assert.match(readme, /data\/live-demo-report\.json/);
assert.match(readme, /直接开启飞书会议/);
assert.match(readme, /开放标注接口/);

console.log('ok real demo live command contract');
