import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-registry-script-'));
const reportFile = join(tmpDir, 'registry-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-registry.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_registry_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.normalizer_count, 3);
assert.equal(report.runtime_ready_count, 3);
assert.equal(report.contract_accepted_count, 3);
assert.equal(report.provider_required_for_realtime_count, 0);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.acceptance.accepted, true);
assert.equal(report.acceptance.blocking_count, 0);
assert.equal(report.acceptance.manifest, undefined);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').insert_endpoint, `${baseUrl}/api/annotations`);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').runtime_event_action_count, 18);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_provider_blocks_realtime, false);
assert.equal(report.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(report.manifest.entries.length, 3);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.acceptance.accepted, true);
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'microsoft_teams').provider.security_verifier, 'verifyMicrosoftGraphClientState');
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').annotations.runtime_event_plan.supported_actions.includes('insert_annotation'), true);
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').annotations.runtime_event_plan.supported_actions.includes('run_manifest'), true);
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').sdk.imports.runtime_event, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event');
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').sdk.imports.adapter_route, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route');
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').sdk.imports.adapter_blueprint, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint');
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').adapter_route.first_route, 'local_observer_axis');
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').adapter_blueprint.ready, true);
assert.equal(writtenReport.manifest.entries.find((entry) => entry.platform === 'google_meet').host.endpoints.adapter_blueprints, '/api/meeting-platform/adapter-blueprints');

const { stdout: compactStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-registry.mjs',
  '--platforms=webex',
  '--include-entries=false',
  '--json=true',
], {
  cwd: repoRoot,
});
const compactReport = JSON.parse(compactStdout);
assert.equal(compactReport.manifest.entries, undefined);
assert.equal(compactReport.rows[0].platform, 'webex');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-registry.mjs',
  '--platforms=lark',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_registry_report/);
assert.match(textStdout, /lark: normalize=yes/);
assert.match(textStdout, /provider=/);
assert.match(textStdout, /runtime_actions=18/);

console.log('ok meeting platform registry script');
