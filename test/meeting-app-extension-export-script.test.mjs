import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const outDir = await mkdtemp(join(tmpdir(), 'meeting-app-extension-'));

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/export-meeting-app-extension.mjs',
  `--out-dir=${outDir}`,
  '--base-url=https://timeline.example.com',
  '--platforms=google-meet,teams',
  '--json=true',
], { cwd: repoRoot });

const summary = JSON.parse(stdout);
assert.equal(summary.type, 'meeting_app_extension_export');
assert.equal(summary.ok, true);
assert.equal(summary.out_dir, outDir);
assert.equal(summary.base_url, 'https://timeline.example.com');
assert.deepEqual(summary.platforms, ['google_meet', 'microsoft_teams']);
assert.equal(summary.acceptance.accepted, true);
assert.equal(summary.acceptance.uses_all_urls, false);
assert.equal(summary.build_gate.requested, false);
assert.equal(summary.build_gate.ok, true);
assert.equal(summary.build_gate.validation.reason, 'build_not_requested');

const filePaths = summary.files.map((file) => file.path);
assert.deepEqual(filePaths, [
  'package.json',
  'build.mjs',
  'manifest.json',
  'src/content-script.entry.mjs',
  'src/live-capture.entry.mjs',
  'src/background.entry.mjs',
  'src/popup.entry.mjs',
  'popup.html',
  'popup.css',
  'README.md',
]);

const generatedPackage = JSON.parse(await readFile(join(outDir, 'package.json'), 'utf8'));
assert.equal(generatedPackage.name, 'meeting-timeline-real-page-extension');
assert.match(generatedPackage.dependencies['@ai-annotation/meeting-timeline-sdk'], /^file:/);
assert.match(generatedPackage.dependencies['@ai-annotation/meeting-timeline-sdk'], /packages\/meeting-timeline-sdk$/);
assert.equal(generatedPackage.scripts.build, 'node build.mjs');

const manifest = JSON.parse(await readFile(join(outDir, 'manifest.json'), 'utf8'));
assert.equal(manifest.manifest_version, 3);
assert.deepEqual(manifest.content_scripts[0].js, ['content-script.js', 'live-capture.js']);
assert.equal(manifest.host_permissions.includes('<all_urls>'), false);
assert.equal(manifest.host_permissions.includes('https://timeline.example.com/*'), true);
assert.equal(manifest.content_scripts[0].matches.includes('https://meet.google.com/*'), true);
assert.equal(manifest.content_scripts[0].matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(manifest.content_scripts[0].matches.includes('https://*.teams.microsoft.com/*'), true);
assert.deepEqual(manifest.background, { service_worker: 'background.js', type: 'module' });
assert.equal(manifest.action.default_popup, 'popup.html');
assert.equal(manifest.optional_host_permissions.includes('http://*/*'), true);
assert.equal(manifest.optional_host_permissions.includes('https://*/*'), true);

const buildSource = await readFile(join(outDir, 'build.mjs'), 'utf8');
assert.match(buildSource, /src\/content-script\.entry\.mjs/);
assert.match(buildSource, /src\/live-capture\.entry\.mjs/);
assert.match(buildSource, /src\/background\.entry\.mjs/);
assert.match(buildSource, /src\/popup\.entry\.mjs/);

const liveCaptureSource = await readFile(join(outDir, 'src/live-capture.entry.mjs'), 'utf8');
assert.match(liveCaptureSource, /__meetingTimelineLiveCapture/);
assert.match(liveCaptureSource, /captureActive/);
assert.match(liveCaptureSource, /captureEnded/);
assert.match(liveCaptureSource, /meeting_app_dom_adaptation_diagnosis/);
assert.match(liveCaptureSource, /diagnose/);

const popupSource = await readFile(join(outDir, 'src/popup.entry.mjs'), 'utf8');
assert.match(popupSource, /checkCurrentMeeting/);
assert.match(popupSource, /meeting_timeline\.capture_evidence/);
const popupHtml = await readFile(join(outDir, 'popup.html'), 'utf8');
assert.match(popupHtml, /popup\.js/);

const readme = await readFile(join(outDir, 'README.md'), 'utf8');
assert.match(readme, /npm run build/);
assert.match(readme, /window\.__meetingTimelineLiveCapture/);
assert.match(readme, /diagnose\(\)/);

console.log('ok meeting app extension export script');
