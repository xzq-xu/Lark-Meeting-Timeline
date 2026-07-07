import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-host-integration-script-'));
const outDir = join(tmpDir, 'host-scaffold');
const reportFile = join(tmpDir, 'host-integration-report.json');
const baseUrl = 'https://timeline.example.com';
const basePath = '/api/meeting-platform/events';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-host-integration.mjs',
  `--base-url=${baseUrl}`,
  `--base-path=${basePath}`,
  '--platforms=google-meet,teams,zoom',
  '--package-name=timeline-host-test',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_host_integration_scaffold_report');
assert.equal(report.ok, true);
assert.equal(report.accepted, true);
assert.equal(report.base_url, baseUrl);
assert.equal(report.base_path, basePath);
assert.equal(report.package_name, 'timeline-host-test');
assert.equal(report.platform_count, 3);
assert.equal(report.file_count > 10, true);
assert.equal(report.written_file_count, report.file_count);
assert.equal(report.adapter_runtime_ready, true);
assert.equal(report.adapter_runtime_ready_count, 3);
assert.equal(report.candidate_observation_ready, true);
assert.equal(report.meeting_track_ready, true);
assert.equal(report.platform_conformance_ready, true);
assert.equal(report.observer_plan_ready, true);
assert.equal(report.adapter_blueprint_ready, true);
assert.equal(report.generated_files.some((file) => file.path === 'src/platform-adapters/google_meet.mjs'), true);
assert.equal(report.generated_files.some((file) => file.path === 'scripts/print-platform-adapters.mjs'), true);
assert.equal(report.scaffold.files[0].content, undefined);
assert.equal(report.scaffold.plan.adapter_runtime_contract.all_ready, true);
assert.equal(report.acceptance.adapter_runtime_ready, true);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.accepted, true);
assert.equal(writtenReport.required_files.includes('src/platform-adapters/microsoft_teams.mjs'), true);

const manifest = JSON.parse(await readFile(join(outDir, 'package.json'), 'utf8'));
assert.equal(manifest.name, 'timeline-host-test');
assert.equal(manifest.scripts['meeting-platform:adapters'], 'node ./scripts/print-platform-adapters.mjs');
assert.equal(manifest.dependencies['@ai-annotation/meeting-timeline-sdk'], '^0.1.0');

const hostSource = await readFile(join(outDir, 'src', 'meeting-platform-host.mjs'), 'utf8');
assert.equal(hostSource.includes('createMeetingPlatformIntegrationRuntime'), true);
assert.equal(hostSource.includes('platformAdapter(platform'), true);

const adapterIndexSource = await readFile(join(outDir, 'src', 'platform-adapters', 'index.mjs'), 'utf8');
assert.equal(adapterIndexSource.includes('createMeetingPlatformAdapter'), true);
assert.equal(adapterIndexSource.includes('createGoogleMeetTimelineAdapter'), true);

const googleAdapterSource = await readFile(join(outDir, 'src', 'platform-adapters', 'google_meet.mjs'), 'utf8');
assert.equal(googleAdapterSource.includes('observeCandidates'), true);
assert.equal(googleAdapterSource.includes('insertAnnotation'), true);
assert.equal(googleAdapterSource.includes('captured_at_ms'), true);

const readme = await readFile(join(outDir, 'README.md'), 'utf8');
assert.equal(readme.includes('Adapter runtime entries'), true);
assert.equal(readme.includes('Meeting track contract'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-host-integration.mjs',
  '--platforms=webex',
  '--write-scaffold=false',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_host_integration_scaffold_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: adapter=src\/platform-adapters\/webex\.mjs/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-host-integration.mjs',
  '--platforms=zoom',
  '--json=true',
  '--write-scaffold=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_host_integration_scaffold_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 1);
assert.deepEqual(binReport.platforms, ['zoom']);
assert.equal(binReport.written_file_count, 0);
assert.equal(binReport.generated_files.some((file) => file.path === 'src/platform-adapters/zoom.mjs'), true);

console.log('ok meeting platform host integration script');
