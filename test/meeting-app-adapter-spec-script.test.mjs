import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-adapter-spec-'));
const specDir = join(tmpDir, 'input');
const outDir = join(tmpDir, 'specs');
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
  'scripts/meeting-app-adapter-spec.mjs',
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

assert.match(stdout, /meeting_app_adapter_spec_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /specs=4/);
assert.match(stdout, /accepted=4/);
assert.match(stdout, /built_in=2/);
assert.match(stdout, /custom=2/);
assert.match(stdout, /written=4/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_adapter_spec_report');
assert.equal(report.ok, true);
assert.equal(report.spec_count, 4);
assert.equal(report.accepted_count, 4);
assert.equal(report.matrix.specs, undefined);
assert.equal(report.rows.find((row) => row.adapter_key === 'whereby').source, 'custom_spec');
assert.equal(report.rows.find((row) => row.adapter_key === 'slack_huddle').accepted, true);

const whereby = JSON.parse(await readFile(join(outDir, 'whereby.json'), 'utf8'));
assert.equal(whereby.schema, 'meeting_app_adapter_spec');
assert.equal(whereby.adapter_key, 'whereby');
assert.equal(whereby.accepted, true);
assert.equal(whereby.contracts.timestamp_field, 'captured_at_ms');

const template = JSON.parse(await readFile(templateFile, 'utf8'));
assert.equal(template.adapter_key, 'slack_huddle');
assert.equal(template.matches[0], 'https://app.slack.com/*');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-adapter-spec.mjs',
  '--platforms=google-meet',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.rows[0].adapter_key, 'google_meet');
assert.equal(jsonReport.matrix.specs, undefined);

console.log('ok meeting app adapter spec script');
