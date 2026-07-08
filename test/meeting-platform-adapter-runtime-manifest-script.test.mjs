import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-runtime-manifest-script-'));
const outFile = join(tmpDir, 'adapter-runtime-manifest.json');
const reportFile = join(tmpDir, 'adapter-runtime-manifest-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-manifest.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-file=${outFile}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_runtime_manifest_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.runtime_ready_count, 3);
assert.equal(report.browser_surface_count, 1);
assert.equal(report.native_surface_count, 2);
assert.equal(report.provider_reconcile_surface_count, 0);
assert.equal(report.local_surface_count, 3);
assert.equal(report.raw_signal_runtime_event_count, 12);
assert.equal(report.host_endpoints.runtime_events, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(report.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(report.runtime_contract.observe_before_insert, true);
assert.equal(report.bridge_groups.find((row) => row.bridge_kind === 'browser_content_script').platforms.includes('google_meet'), true);
assert.equal(report.bridge_groups.find((row) => row.bridge_kind === 'native_detector_runtime_event_client').platform_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').host_kind, 'browser_extension_content_script');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').first_required_method, 'observePlatformCandidates');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').insert_method, 'insertAnnotation');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').host_kind, 'native_desktop_detector');
assert.equal(report.rows.find((row) => row.platform === 'zoom').speaker_position_markers.enabled, true);
assert.equal(report.manifest, undefined);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.rows.find((row) => row.platform === 'zoom').bridge_kind, 'native_detector_runtime_event_client');

const manifest = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(manifest.schema, 'meeting_platform_adapter_runtime_manifest');
assert.equal(manifest.runtime_ready, true);
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'google_meet').adapter_module, '@ai-annotation/meeting-timeline-sdk/adapters/google-meet');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'microsoft_teams').adapter_module, '@ai-annotation/meeting-timeline-sdk/adapters/microsoft-teams');
assert.equal(manifest.platform_registry.rows.find((row) => row.platform === 'zoom').host_kind, 'native_desktop_detector');
assert.equal(manifest.recipes, undefined);

const { stdout: includedStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-manifest.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet',
  '--json=true',
  '--include-manifest=true',
  '--include-recipes=true',
], {
  cwd: repoRoot,
});
const includedReport = JSON.parse(includedStdout);
assert.equal(includedReport.manifest.schema, 'meeting_platform_adapter_runtime_manifest');
assert.equal(includedReport.manifest.recipes.length, 1);
assert.equal(includedReport.manifest.recipes[0].startup_plan, undefined);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-manifest.mjs',
  '--platforms=webex',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_runtime_manifest_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: runtime=yes/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-runtime-manifest.mjs',
  '--platforms=google-meet,zoom',
  '--json=true',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_adapter_runtime_manifest_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.rows.find((row) => row.platform === 'google_meet').host_kind, 'browser_extension_content_script');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').host_kind, 'native_desktop_detector');

console.log('ok meeting platform adapter runtime manifest script');
