import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const packageJson = JSON.parse(
  await readFile(new URL('../package.json', import.meta.url), 'utf8'),
);
const readme = await readFile(new URL('../README.md', import.meta.url), 'utf8');

assert.equal(packageJson.scripts.start, 'REAL_DEMO_AUTO_ARM=1 REAL_DEMO_AUTO_ANNOTATION=0 REAL_DEMO_DEVICE_SIMULATOR=0 REAL_DEMO_DEVICE_STREAM=0 node src/server.mjs');
assert.equal(packageJson.scripts['start:plain'], 'node src/server.mjs');
assert.equal(
  packageJson.scripts['auth:meeting-scan'],
  'node scripts/monitor-real-demo.mjs --open-auth --wait-auth --scan-after-auth --once --auth-timeout-ms=300000',
);
assert.equal(packageJson.scripts['event:report'], 'node scripts/event-delivery-report.mjs --report-file=data/event-delivery-report.json');
assert.equal(packageJson.scripts['onsite:status'], 'node scripts/onsite-status.mjs --report-file=data/onsite-status-report.json');
assert.equal(packageJson.scripts['auth:status'], 'node scripts/auth-meeting-scan-status.mjs --report-file=data/auth-meeting-scan-status-report.json');
assert.equal(packageJson.scripts['auth:open'], 'node scripts/auth-meeting-scan-status.mjs --open=true --wait=true --report-file=data/auth-meeting-scan-status-report.json');
assert.equal(
  packageJson.scripts['demo:live'],
  'node scripts/monitor-real-demo.mjs --prepare --open-auth --wait-auth --scan-after-auth --auto-mark --timeout-ms=900000 --auth-timeout-ms=300000 --report-file=data/live-demo-report.json',
);
assert.equal(
  packageJson.scripts['accept:real-meeting'],
  'node scripts/monitor-real-demo.mjs --prepare --auto-mark --timeout-ms=900000 --report-file=data/real-meeting-acceptance-report.json',
);
assert.equal(
  packageJson.scripts['accept:real-meeting:auth'],
  'node scripts/monitor-real-demo.mjs --prepare --open-auth --wait-auth --scan-after-auth --auto-mark --timeout-ms=900000 --auth-timeout-ms=300000 --report-file=data/real-meeting-acceptance-report.json',
);
assert.equal(
  packageJson.scripts['accept:onsite'],
  'node scripts/device-preflight.mjs --report-file=data/onsite-device-preflight-report.json && node scripts/monitor-real-demo.mjs --prepare --open-auth --wait-auth --scan-after-auth --auto-mark --timeout-ms=900000 --auth-timeout-ms=300000 --report-file=data/onsite-acceptance-report.json',
);
assert.equal(
  packageJson.scripts['accept:onsite:strict'],
  'node scripts/device-preflight.mjs --report-file=data/onsite-strict-device-preflight-report.json && node scripts/monitor-real-demo.mjs --prepare --open-auth --wait-auth --scan-after-auth --auto-mark --require-event-axis --timeout-ms=900000 --auth-timeout-ms=300000 --report-file=data/onsite-strict-event-acceptance-report.json',
);
assert.equal(packageJson.scripts['device:preflight'], 'node scripts/device-preflight.mjs --report-file=data/device-preflight-report.json');
assert.equal(packageJson.scripts['device:roundtrip'], 'node scripts/device-roundtrip.mjs');

assert.match(readme, /`npm run start` 会设置 `REAL_DEMO_AUTO_ARM=1`/);
assert.match(readme, /默认不会自动写入验收标注、虚拟墨水屏标注或设备流标注/);
assert.match(readme, /npm run start:plain/);
assert.match(readme, /npm run auth:meeting-scan/);
assert.match(readme, /npm run event:report/);
assert.match(readme, /npm run onsite:status/);
assert.match(readme, /npm run auth:status/);
assert.match(readme, /npm run auth:open/);
assert.match(readme, /npm run demo:live/);
assert.match(readme, /npm run accept:real-meeting/);
assert.match(readme, /npm run accept:real-meeting:auth/);
assert.match(readme, /npm run accept:onsite/);
assert.match(readme, /npm run accept:onsite:strict/);
assert.match(readme, /npm run device:preflight/);
assert.match(readme, /data\/device-preflight-report\.json/);
assert.match(readme, /data\/event-delivery-report\.json/);
assert.match(readme, /data\/onsite-status-report\.json/);
assert.match(readme, /data\/auth-meeting-scan-status-report\.json/);
assert.match(readme, /data\/live-demo-report\.json/);
assert.match(readme, /data\/onsite-device-preflight-report\.json/);
assert.match(readme, /data\/onsite-strict-device-preflight-report\.json/);
assert.match(readme, /npm run device:roundtrip/);
assert.match(readme, /点击“启动事件等待”或“启动验收探针”只启动等待\/验收窗口，不会创建会议轴/);

console.log('ok runtime script contract');
