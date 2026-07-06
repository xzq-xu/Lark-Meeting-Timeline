import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-adapter-handoff-package-'));
const specDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'packages');
const reportFile = join(tmpDir, 'report.json');
const templateFile = join(tmpDir, 'template.json');
await mkdir(specDir, { recursive: true });

const wherebySpec = {
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
  host_permissions: ['https://whereby.com/*'],
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
};
await writeFile(join(specDir, 'whereby.json'), `${JSON.stringify(wherebySpec, null, 2)}\n`, 'utf8');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-handoff-package.mjs',
  '--platforms=google-meet,zoom',
  `--spec-dir=${specDir}`,
  '--template-adapter-key=slack-huddle',
  '--template-matches=https://app.slack.com/*',
  `--template-file=${templateFile}`,
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_app_adapter_handoff_package_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /packages=4/);
assert.match(stdout, /accepted=4/);
assert.match(stdout, /built_in=2/);
assert.match(stdout, /custom=2/);
assert.match(stdout, /runtime_ready=4/);
assert.match(stdout, /content_script_ready=4/);
assert.match(stdout, /live_evidence_required=4/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_adapter_handoff_package_report');
assert.equal(report.ok, true);
assert.equal(report.package_count, 4);
assert.equal(report.accepted_count, 4);
assert.equal(report.matrix.packages, undefined);
assert.equal(report.rows.find((row) => row.adapter_key === 'whereby').source, 'custom_spec');
assert.equal(report.rows.find((row) => row.adapter_key === 'slack_huddle').accepted, true);
assert.equal(report.written_files.length, 22);

const wherebyRuntime = JSON.parse(await readFile(join(outDir, 'whereby', 'runtime-config.json'), 'utf8'));
assert.equal(wherebyRuntime.schema, 'meeting_app_adapter_runtime_config');
assert.equal(wherebyRuntime.accepted, true);
assert.equal(wherebyRuntime.contracts.timestamp_field, 'captured_at_ms');

const wherebyReadme = await readFile(join(outDir, 'whereby', 'integration-readme.md'), 'utf8');
assert.match(wherebyReadme, /installMeetingAppContentScriptBridge/);
assert.match(wherebyReadme, /captured_at_ms/);

const googleManifest = JSON.parse(await readFile(join(outDir, 'google_meet', 'adapter-manifest.json'), 'utf8'));
assert.equal(googleManifest.schema, 'meeting_app_adapter_manifest');

const descriptor = JSON.parse(await readFile(join(outDir, 'whereby', 'handoff-package.json'), 'utf8'));
assert.equal(descriptor.schema, 'meeting_app_adapter_handoff_package');
assert.equal(descriptor.files.some((file) => 'content' in file), false);
assert.equal(descriptor.written_files.length, 4);

const template = JSON.parse(await readFile(templateFile, 'utf8'));
assert.equal(template.adapter_key, 'slack_huddle');
assert.equal(template.matches[0], 'https://app.slack.com/*');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-handoff-package.mjs',
  '--platforms=google-meet',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.rows[0].adapter_key, 'google_meet');
assert.equal(jsonReport.matrix.packages, undefined);

console.log('ok meeting app adapter handoff package script');
