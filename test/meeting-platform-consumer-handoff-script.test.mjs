import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-consumer-handoff-script-'));
const reportFile = join(tmpDir, 'consumer-handoff-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-consumer-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_platform_consumer_handoff/);
assert.match(stdout, /accepted=yes/);
assert.match(stdout, /consumer_ready=3/);
assert.match(stdout, /production=0/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.schema, 'meeting_platform_consumer_handoff');
assert.equal(report.accepted, true);
assert.equal(report.platform_count, 3);
assert.equal(report.consumer_ready_count, 3);
assert.equal(report.lightweight_connector_ready, true);
assert.equal(report.lightweight_connector_platform_count, 3);
assert.equal(report.candidate_observer_count, 3);
assert.equal(report.speaker_track_ready_count, 3);
assert.equal(report.participant_track_ready_count, 3);
assert.equal(report.production_ready_count, 0);
assert.equal(report.entrypoints.primary_modules.consumer_handoff, '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff');
assert.equal(report.entrypoints.primary_modules.meeting_platform_connector, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(report.lightweight_connector_handoff.content_script_bridge.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(report.lightweight_connector_handoff.host_requirements.timestamp_field, 'captured_at_ms');
assert.equal(report.sdk_facade_handoff.create_function, 'createMeetingAppTimelineSdk');
assert.equal(report.sdk_facade_handoff.required_facade_methods.includes('platformConsumerHandoff'), true);
assert.equal(report.sdk_facade_handoff.surface_wiring.browser_extension.content_script_messages.includes('meeting_timeline.observe_candidates'), true);
assert.equal(report.sdk_facade_handoff.platform_rows.find((row) => row.platform === 'google_meet').provider_path, 'google_workspace_events_pubsub');
assert.equal(report.surface_coverage_matrix.platform_count, 3);
assert.equal(report.surface_coverage_matrix.browser_extension_ready_count, 3);
assert.equal(report.surface_coverage_matrix.provider_reconcile_ready_count, 3);
assert.equal(report.surface_coverage_matrix.rows.find((row) => row.platform === 'google_meet').webview_preload.ready, true);
assert.equal(report.surface_coverage_matrix.rows.find((row) => row.platform === 'microsoft_teams').provider_reconcile.provider_path, 'microsoft_graph_change_notifications');
assert.equal(report.adaptation_roadmap.schema, 'meeting_platform_adaptation_roadmap');
assert.equal(report.adaptation_roadmap.platform_count, 3);
assert.equal(report.adaptation_roadmap.recommended_first_platform, 'google_meet');
assert.equal(report.adaptation_roadmap.recommended_first_surface, 'browser_extension');
assert.deepEqual(report.adaptation_roadmap.priority_order.slice(0, 3), ['google_meet', 'zoom', 'microsoft_teams']);
assert.equal(report.adaptation_roadmap.rows[0].platform, 'google_meet');
assert.equal(report.adaptation_roadmap.rows[0].recommended_first_surface, 'browser_extension');
assert.equal(report.adaptation_roadmap.rows.find((row) => row.platform === 'microsoft_teams').provider_permission_risk, 'tenant_admin_consent_and_subscription_renewal');
assert.equal(report.rows.find((row) => row.platform === 'microsoft_teams').consumer_ready, true);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-consumer-handoff.mjs',
  '--platforms=webex',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.rows[0].platform, 'webex');
assert.equal(jsonReport.rows[0].consumer_ready, true);

let strictStdout = '';
let strictCode = 0;
try {
  const result = await execFileAsync(process.execPath, [
    'scripts/meeting-platform-consumer-handoff.mjs',
    '--platforms=zoom',
    '--require-production-ready=true',
    '--fail-on-rejected=true',
  ], {
    cwd: repoRoot,
  });
  strictStdout = result.stdout;
} catch (error) {
  strictStdout = error.stdout;
  strictCode = error.code;
}
assert.equal(strictCode, 2);
assert.match(strictStdout, /accepted=no/);

console.log('ok meeting platform consumer handoff script');
