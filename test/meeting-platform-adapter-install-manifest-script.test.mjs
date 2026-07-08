import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-install-manifest-script-'));
const exportDir = join(tmpDir, 'exports');
const importDir = join(tmpDir, 'imports');
const manifestFile = join(tmpDir, 'adapter-install-manifest.json');
const reportFile = join(tmpDir, 'adapter-install-manifest-report.json');
const baseUrl = 'https://timeline.example.com';

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-export-package.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-dir=${exportDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-import-plan.mjs',
  `--dir=${exportDir}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-dir=${importDir}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-install-manifest.mjs',
  `--dir=${importDir}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--base-url=${baseUrl}`,
  '--install-target=desktop_host',
  `--out-file=${manifestFile}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_install_manifest_report');
assert.equal(report.ok, true);
assert.equal(report.plan_count, 2);
assert.equal(report.platform_count, 2);
assert.equal(report.ready_platform_count, 2);
assert.equal(report.blocked_platform_count, 0);
assert.equal(report.browser_content_script_count, 1);
assert.equal(report.adapter_preflight_platform_count, 2);
assert.equal(report.adapter_preflight_startup_ready_count, 2);
assert.equal(report.adapter_preflight_realtime_ready_count, 0);
assert.equal(report.raw_signal_validation_platform_count, 2);
assert.equal(report.raw_signal_validation_ready_count, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').raw_signal_validation_status, 'ready');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_preflight_status, 'needs_live_page_evidence');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_preflight_selected_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_preflight_startup_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_preflight_realtime_ready, false);
assert.equal(report.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(report.rows.find((row) => row.platform === 'zoom').adapter_preflight_selected_surface, 'native_detector');

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.ok, true);

const manifest = JSON.parse(await readFile(manifestFile, 'utf8'));
assert.equal(manifest.schema, 'meeting_platform_adapter_install_manifest');
assert.equal(manifest.accepted, true);
assert.equal(manifest.sdk_imports.adapter_preflight, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight');
assert.equal(manifest.sdk_imports.raw_signal, '@ai-annotation/meeting-timeline-sdk/adapters/platform-raw-signal');
assert.equal(manifest.runtime_contract.raw_signal_validation_required_before_preflight, true);
assert.equal(manifest.runtime_contract.adapter_preflight_required_before_realtime_insert, true);
assert.equal(manifest.runtime_contract.adapter_preflight_url_only_status, 'needs_live_page_evidence');
assert.equal(manifest.adapter_blueprints.enabled, true);
assert.equal(manifest.adapter_blueprints.platform_count, 2);
assert.equal(manifest.adapter_blueprints.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(manifest.raw_signal_validation.enabled, true);
assert.equal(manifest.raw_signal_validation.platform_count, 2);
assert.equal(manifest.raw_signal_validation.ready_count, 2);
assert.equal(manifest.raw_signal_validation.rows.find((row) => row.platform === 'google_meet').status, 'ready');
assert.equal(manifest.adapter_preflight.enabled, true);
assert.equal(manifest.adapter_preflight.platform_count, 2);
assert.equal(manifest.adapter_preflight.startup_ready_count, 2);
assert.equal(manifest.adapter_preflight.realtime_ready_count, 0);
assert.equal(manifest.adapter_preflight.rows.find((row) => row.platform === 'google_meet').status, 'needs_live_page_evidence');
assert.equal(manifest.adapter_preflight.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(manifest.native_detector.platform_count, 1);
assert.equal(manifest.browser_extension.host_permissions.includes('https://meet.google.com/*'), true);
assert.equal(manifest.browser_extension.message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(manifest.browser_extension.message_types.includes('meeting_timeline.preflight_candidates'), true);
assert.equal(manifest.install_sequence.find((step) => step.id === 'run_raw_signal_validation_before_preflight').sdk_method, 'platformRawSignalBatch');
assert.equal(manifest.install_sequence.find((step) => step.id === 'run_adapter_preflight_before_realtime_session').url_only_status, 'needs_live_page_evidence');
assert.equal(manifest.install_sequence.find((step) => step.id === 'insert_marks_in_realtime').sdk_method, 'insertAnnotation');
assert.equal(manifest.next_actions.includes('run_adapter_preflight_with_live_window_evidence_before_first_insert'), true);

const { stdout: missingStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-install-manifest.mjs',
  `--dir=${importDir}`,
  '--platforms=google-meet,webex',
  '--target=static',
  `--out-file=${join(tmpDir, 'missing-install-manifest.json')}`,
  '--json=true',
], {
  cwd: repoRoot,
});
const missingReport = JSON.parse(missingStdout);
assert.equal(missingReport.plan_count, 1);
assert.equal(missingReport.missing_plan_file_count, 1);
assert.equal(missingReport.ok, false);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-install-manifest.mjs',
  `--dir=${importDir}`,
  '--platforms=google-meet,zoom',
  '--target=static',
  `--out-file=${join(tmpDir, 'text-install-manifest.json')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_install_manifest_report/);
assert.match(textStdout, /google_meet: ready=yes/);
assert.match(textStdout, /raw_signal=ready/);
assert.match(textStdout, /preflight=needs_live_page_evidence/);
assert.match(textStdout, /realtime=no/);

console.log('ok meeting platform adapter install manifest script');
