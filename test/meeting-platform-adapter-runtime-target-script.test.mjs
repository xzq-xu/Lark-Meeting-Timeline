import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

import {
  buildMeetingPlatformAdapterRuntimeManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runtime-recipe.mjs';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-runtime-target-script-'));
const manifestFile = join(tmpDir, 'adapter-runtime-manifest.json');
const targetFile = join(tmpDir, 'adapter-runtime-target.json');
const reportFile = join(tmpDir, 'adapter-runtime-target-report.json');
const baseUrl = 'https://timeline.example.com';

const manifest = buildMeetingPlatformAdapterRuntimeManifest({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
await writeFile(manifestFile, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-target.mjs',
  `--base-url=${baseUrl}`,
  `--manifest-file=${manifestFile}`,
  '--url=https://meet.google.com/abc-defg-hij',
  `--out-file=${targetFile}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_runtime_target_report');
assert.equal(report.ok, true);
assert.equal(report.platform, 'google_meet');
assert.equal(report.detection_reason, 'meeting_url');
assert.equal(report.selected_surface, 'browser_extension');
assert.equal(report.host_kind, 'browser_extension_content_script');
assert.equal(report.bridge_kind, 'browser_content_script');
assert.equal(report.first_required_method, 'observePlatformCandidates');
assert.equal(report.insert_method, 'insertAnnotation');
assert.equal(report.timestamp_field, 'captured_at_ms');
assert.equal(report.runtime_action_count > 0, true);
assert.equal(report.issue_count, 0);
assert.equal(report.target.mark_template.source, 'meeting_platform_adapter_runtime_target');

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.target.host_kind, 'browser_extension_content_script');

const writtenTarget = JSON.parse(await readFile(targetFile, 'utf8'));
assert.equal(writtenTarget.schema, 'meeting_platform_adapter_runtime_target');
assert.equal(writtenTarget.accepted, true);
assert.equal(writtenTarget.host_kind, 'browser_extension_content_script');

const { stdout: zoomStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-target.mjs',
  `--base-url=${baseUrl}`,
  '--platform=zoom',
  '--platforms=google-meet,zoom',
  '--json=true',
], {
  cwd: repoRoot,
});
const zoomReport = JSON.parse(zoomStdout);
assert.equal(zoomReport.ok, true);
assert.equal(zoomReport.platform, 'zoom');
assert.equal(zoomReport.detection_reason, 'explicit_platform');
assert.equal(zoomReport.host_kind, 'native_desktop_detector');
assert.equal(zoomReport.bridge_kind, 'native_detector_runtime_event_client');

const { stdout: mismatchStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-target.mjs',
  `--manifest-file=${manifestFile}`,
  '--platform=zoom',
  '--surface=browser-extension',
  '--json=true',
], {
  cwd: repoRoot,
});
const mismatchReport = JSON.parse(mismatchStdout);
assert.equal(mismatchReport.ok, false);
assert.equal(mismatchReport.issue_codes.includes('surface_mismatch'), true);
assert.equal(mismatchReport.target.readiness.issues.find((item) => item.code === 'surface_mismatch').requested_surface, 'browser_extension');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-target.mjs',
  '--platform=webex',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_runtime_target_report/);
assert.match(textStdout, /platform=webex/);
assert.match(textStdout, /host=browser_extension_content_script/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-runtime-target.mjs',
  `--manifest-file=${manifestFile}`,
  '--platform=zoom',
  '--json=true',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.ok, true);
assert.equal(binReport.host_kind, 'native_desktop_detector');

console.log('ok meeting platform adapter runtime target script');
