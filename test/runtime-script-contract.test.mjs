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
assert.equal(packageJson.scripts['meeting-app:extension'], 'node scripts/export-meeting-app-extension.mjs --out-dir=data/meeting-app-extension');
assert.equal(packageJson.scripts['meeting-app:extension:build'], 'node scripts/export-meeting-app-extension.mjs --out-dir=data/meeting-app-extension --install=true --build=true');
assert.equal(packageJson.scripts['meeting-app:evidence-gate'], 'node scripts/meeting-app-evidence-gate.mjs');
assert.equal(packageJson.scripts['meeting-app:evidence-matrix'], 'node scripts/meeting-app-evidence-matrix.mjs --dir=data/meeting-app-evidence --report-file=data/meeting-app-evidence-matrix.json');
assert.equal(packageJson.scripts['meeting-app:fixture-tracks'], 'node scripts/meeting-app-fixture-tracks.mjs --platforms=google-meet,teams,zoom,webex,lark --report-file=data/meeting-app-fixture-track-readiness.json');
assert.equal(packageJson.scripts['meeting-app:adapter-manifest'], 'node scripts/meeting-app-adapter-manifest.mjs --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787 --out-dir=data/meeting-app-adapter-manifests --report-file=data/meeting-app-adapter-manifest-report.json');
assert.equal(packageJson.scripts['meeting-app:adapter-spec'], 'node scripts/meeting-app-adapter-spec.mjs --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787 --out-dir=data/meeting-app-adapter-specs --report-file=data/meeting-app-adapter-spec-report.json');
assert.equal(packageJson.scripts['meeting-app:adapter-runtime-config'], 'node scripts/meeting-app-adapter-runtime-config.mjs --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787 --out-dir=data/meeting-app-adapter-runtime-configs --report-file=data/meeting-app-adapter-runtime-config-report.json');
assert.equal(packageJson.scripts['meeting-platform:rollout-matrix'], 'node scripts/meeting-platform-rollout-matrix.mjs --provider-dir=data/provider-evidence --dom-dir=data/meeting-app-evidence --report-file=data/meeting-platform-rollout-matrix.json');
assert.equal(packageJson.scripts['meeting-platform:live-readiness'], 'node scripts/meeting-platform-live-readiness.mjs --dir=data/meeting-platform-evidence-packages --report-file=data/meeting-platform-live-readiness-report.json');
assert.equal(packageJson.scripts['meeting-platform:evidence-package'], 'node scripts/meeting-platform-evidence-package.mjs --dir=data/meeting-platform-evidence-packages --report-file=data/meeting-platform-evidence-package-report.json');
assert.equal(packageJson.scripts['sdk:package-smoke'], 'node test/sdk-package-smoke.test.mjs');

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
assert.match(readme, /npm run meeting-app:extension/);
assert.match(readme, /npm run meeting-app:extension:build/);
assert.match(readme, /npm run meeting-app:evidence-gate/);
assert.match(readme, /npm run meeting-app:evidence-matrix/);
assert.match(readme, /npm run meeting-app:fixture-tracks/);
assert.match(readme, /npm run meeting-app:adapter-manifest/);
assert.match(readme, /npm run meeting-app:adapter-spec/);
assert.match(readme, /npm run meeting-app:adapter-runtime-config/);
assert.match(readme, /npm run meeting-platform:rollout-matrix/);
assert.match(readme, /npm run meeting-platform:live-readiness/);
assert.match(readme, /npm run meeting-platform:evidence-package/);
assert.match(readme, /npm run sdk:package-smoke/);
assert.match(readme, /data\/meeting-app-extension/);
assert.match(readme, /data\/meeting-app-evidence/);
assert.match(readme, /data\/meeting-app-evidence-matrix\.json/);
assert.match(readme, /data\/meeting-app-fixture-track-readiness\.json/);
assert.match(readme, /data\/meeting-app-adapter-manifests/);
assert.match(readme, /data\/meeting-app-adapter-manifest-report\.json/);
assert.match(readme, /data\/meeting-app-adapter-specs/);
assert.match(readme, /data\/meeting-app-adapter-spec-report\.json/);
assert.match(readme, /data\/meeting-app-adapter-runtime-configs/);
assert.match(readme, /data\/meeting-app-adapter-runtime-config-report\.json/);
assert.match(readme, /data\/provider-evidence/);
assert.match(readme, /data\/meeting-platform-rollout-matrix\.json/);
assert.match(readme, /data\/meeting-platform-evidence-packages/);
assert.match(readme, /data\/meeting-platform-live-readiness-report\.json/);
assert.match(readme, /data\/meeting-platform-evidence-package-report\.json/);
assert.match(readme, /data\/meeting-app-live-gate-report\.json/);
assert.match(readme, /window\.__meetingTimelineLiveCapture\.captureActive\(\)/);
assert.match(readme, /点击“启动事件等待”或“启动验收探针”只启动等待\/验收窗口，不会创建会议轴/);

console.log('ok runtime script contract');
