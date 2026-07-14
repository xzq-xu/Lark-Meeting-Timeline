import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const outDir = await mkdtemp(join(tmpdir(), 'three-platform-adapters-'));

try {
  const { stdout } = await execFileAsync(process.execPath, [
    'scripts/export-three-platform-adapters.mjs',
    `--out-dir=${outDir}`,
    '--base-url=http://localhost:8787',
    '--offline=true',
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
  assert.deepEqual(report.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
  assert.equal(report.browser_extension.build_ok, true);
  assert.equal(report.desktop_host.package_ok, true);
  assert.equal(report.desktop_host.install_start_ok, true);
  assert.equal(report.desktop_host.missing_package_files.length, 0);

  await Promise.all([
    stat(join(outDir, 'browser-extension', 'manifest.json')),
    stat(join(outDir, 'browser-extension', 'popup.html')),
    stat(join(outDir, 'meeting-timeline-browser-extension.zip')),
    stat(join(outDir, report.desktop_host.package)),
    stat(join(outDir, 'desktop-host', 'install-macos.command')),
    stat(join(outDir, 'desktop-host', 'install-windows.cmd')),
    stat(join(outDir, 'release-manifest.json')),
  ]);

  const readme = await readFile(join(outDir, 'README.md'), 'utf8');
  assert.match(readme, /自动监听会议开始、结束和稳定发言人/);
  assert.match(readme, /只用于验收取证/);
  assert.match(readme, /不需要调用方自己构造 `windows\[\]`/);
} finally {
  await rm(outDir, { recursive: true, force: true });
}

console.log('ok three platform adapter release');
