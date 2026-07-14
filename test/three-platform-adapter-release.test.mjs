import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { buildThreePlatformReleaseReadiness } from '../scripts/three-platform-release-readiness.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const outDir = await mkdtemp(join(tmpdir(), 'three-platform-adapters-'));

const acceptedLiveRows = ['google_meet', 'microsoft_teams', 'zoom'].map((platform) => ({
  platform,
  accepted: true,
  core_accepted: true,
  speaker_accepted: true,
  production_ready: true,
  full_production_ready: true,
  meeting_id: `${platform}-real-meeting`,
  file: `/evidence/${platform}.json`,
  failed_check_ids: [],
}));
const acceptedReadiness = buildThreePlatformReleaseReadiness({
  accepted: true,
  production_ready: true,
  full_production_ready: true,
  rows: acceptedLiveRows,
});
assert.equal(acceptedReadiness.production_ready, true);
assert.equal(acceptedReadiness.speaker_ready, true);
assert.equal(acceptedReadiness.full_production_ready, true);
assert.equal(acceptedReadiness.adapters.every((row) => row.real_meeting_accepted), true);
assert.equal(acceptedReadiness.remaining_gate, null);

const coreOnlyReadiness = buildThreePlatformReleaseReadiness({
  accepted: true,
  production_ready: true,
  full_production_ready: false,
  rows: acceptedLiveRows.map((row) => ({
    ...row,
    speaker_accepted: false,
    full_production_ready: false,
    failed_speaker_check_ids: ['stable_speaker_marker'],
  })),
});
assert.equal(coreOnlyReadiness.production_ready, true);
assert.equal(coreOnlyReadiness.speaker_ready, false);
assert.equal(coreOnlyReadiness.full_production_ready, false);
assert.equal(coreOnlyReadiness.remaining_gate, null);
assert.deepEqual(coreOnlyReadiness.remaining_speaker_platforms, ['google_meet', 'microsoft_teams', 'zoom']);

const incompleteReadiness = buildThreePlatformReleaseReadiness({
  accepted: false,
  production_ready: false,
  rows: acceptedLiveRows.filter((row) => row.platform !== 'zoom'),
});
assert.equal(incompleteReadiness.production_ready, false);
assert.deepEqual(incompleteReadiness.remaining_platforms, ['zoom']);

try {
  const { stdout } = await execFileAsync(process.execPath, [
    'scripts/export-three-platform-adapters.mjs',
    `--out-dir=${outDir}`,
    '--base-url=http://localhost:8787',
    '--offline=true',
    `--live-evidence-dir=${join(outDir, 'no-live-evidence')}`,
    '--json=true',
  ], {
    cwd: repoRoot,
    env: { ...process.env, NODE_ENV: 'production' },
    maxBuffer: 30 * 1024 * 1024,
  });

  const report = JSON.parse(stdout);
  assert.equal(report.schema, 'three_platform_adapter_release');
  assert.equal(report.ok, true);
  assert.equal(report.delivery_ready, true);
  assert.equal(report.production_ready, false);
  assert.equal(report.speaker_ready, false);
  assert.equal(report.full_production_ready, false);
  assert.deepEqual(report.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
  assert.equal(report.browser_extension.build_ok, true);
  assert.equal(report.desktop_host.package_ok, true);
  assert.equal(report.desktop_host.install_start_ok, true);
  assert.equal(report.desktop_host.missing_package_files.length, 0);
  assert.equal(report.desktop_host.bundled_browser_extension.ok, true);
  assert.equal(report.desktop_host.bundled_browser_extension.manifest_version, 3);
  assert.deepEqual(report.desktop_host.bundled_browser_extension.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
  assert.equal(report.desktop_host.bundled_release_status.ok, true);
  assert.equal(report.desktop_host.bundled_release_status.schema, 'three_platform_adapter_release_status');
  assert.equal(report.desktop_host.bundled_release_status.production_ready, false);
  assert.equal(report.desktop_host.bundled_release_status.speaker_ready, false);
  assert.equal(report.desktop_host.bundled_release_status.full_production_ready, false);
  assert.equal(report.desktop_host.bundled_release_status.adapter_count, 3);
  assert.equal(report.desktop_host.required_package_files.includes('package/runtime/browser-extension/manifest.json'), true);
  assert.equal(report.desktop_host.required_package_files.includes('package/runtime/three-platform-release-status.json'), true);

  const tarballStatus = JSON.parse(await execFileAsync('tar', [
    '-xOzf',
    join(outDir, report.desktop_host.package),
    'package/runtime/three-platform-release-status.json',
  ]).then((result) => result.stdout));
  assert.equal(tarballStatus.schema, 'three_platform_adapter_release_status');
  assert.equal(tarballStatus.adapters.some((row) => String(row.evidence_file ?? '').startsWith('/')), false);

  await Promise.all([
    stat(join(outDir, 'browser-extension', 'manifest.json')),
    stat(join(outDir, 'browser-extension', 'popup.html')),
    stat(join(outDir, 'meeting-timeline-browser-extension.zip')),
    stat(join(outDir, report.desktop_host.package)),
    stat(join(outDir, 'desktop-host', 'install-macos.command')),
    stat(join(outDir, 'desktop-host', 'install-windows.cmd')),
    stat(join(outDir, 'release-manifest.json')),
  ]);

  const extensionManifest = JSON.parse(await readFile(join(outDir, 'browser-extension', 'manifest.json'), 'utf8'));
  assert.equal(extensionManifest.content_scripts[0].all_frames, false);

  const readme = await readFile(join(outDir, 'README.md'), 'utf8');
  assert.match(readme, /自动监听会议开始、结束和稳定发言人/);
  assert.match(readme, /只用于验收取证/);
  assert.match(readme, /不需要调用方自己构造 `windows\[\]`/);
  assert.match(readme, /meeting-timeline-adapters --out-dir/);
  assert.match(readme, /meeting-timeline-adapters --status=true/);
  assert.match(readme, /不需要回到源码仓库构建/);
  assert.match(readme, /--speaker-peer=true/);
  assert.match(readme, /不要求第二账号/);
  assert.match(readme, /full_production_ready/);
} finally {
  await rm(outDir, { recursive: true, force: true });
}

console.log('ok three platform adapter release');
