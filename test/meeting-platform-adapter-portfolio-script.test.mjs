import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-portfolio-script-'));
const outDir = join(tmpDir, 'portfolio');
const reportFile = join(tmpDir, 'adapter-portfolio-report.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-portfolio.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,Acme Rooms',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_portfolio_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.built_in_count, 1);
assert.equal(report.external_authoring_count, 1);
assert.equal(report.browser_surface_ready_count, 1);
assert.equal(report.browser_first_count, 1);
assert.equal(report.native_first_count, 1);
assert.equal(report.local_axis_first_count, 2);
assert.equal(report.provider_reconcile_count, 2);
assert.equal(report.provider_required_for_realtime_count, 0);
assert.equal(report.post_meeting_transcript_count, 1);
assert.equal(report.transcript_blocking_count, 0);
assert.equal(report.implementation_ready_count, 1);
assert.equal(report.pilot_ready_count, 1);
assert.equal(report.production_ready_count, 0);
assert.equal(report.written_files.length, 2);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').item_file, join(outDir, 'google_meet.json'));
assert.equal(report.rows.find((row) => row.platform === 'google_meet').official_doc_count, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').first_surface_family, 'browser_observer');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').provider_required_for_realtime, false);
assert.equal(report.rows.find((row) => row.platform === 'acme_rooms').adapter_status, 'adapter_authoring_required');
assert.equal(report.portfolio.items.some((item) => item.artifacts), false);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 2);
assert.equal(writtenReport.portfolio.items.some((item) => item.artifacts), false);

const googleItem = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleItem.schema, 'meeting_platform_adapter_portfolio_item');
assert.equal(googleItem.built_in, true);
assert.equal(googleItem.p1_provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(googleItem.p1_provider_reconcile.official_doc_count, 3);
assert.equal(googleItem.p0_realtime_axis.timestamp_field, 'captured_at_ms');
assert.equal(googleItem.adapter_strategy.realtime_axis.owner, 'local_observer');
assert.equal(googleItem.adapter_strategy.provider_reconcile.required_for_realtime, false);
assert.equal(googleItem.adapter_strategy.post_meeting_artifacts.transcript_blocks_realtime, false);

const acmeItem = JSON.parse(await readFile(join(outDir, 'acme_rooms.json'), 'utf8'));
assert.equal(acmeItem.schema, 'meeting_platform_adapter_portfolio_item');
assert.equal(acmeItem.built_in, false);
assert.equal(acmeItem.adapter_status, 'adapter_authoring_required');
assert.equal(acmeItem.next_actions.includes('add_event_normalizer'), true);

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-portfolio.mjs',
  '--platforms=webex',
  `--out-dir=${join(tmpDir, 'webex-portfolio')}`,
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_portfolio_report/);
assert.match(textStdout, /builtin=1/);
assert.match(textStdout, /browser_first=1/);
assert.match(textStdout, /local_axis=1/);
assert.match(textStdout, /webex: status=built_in_adapter_available/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-portfolio.mjs',
  '--platforms=google-meet,zoom',
  '--json=true',
  '--write-items=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_adapter_portfolio_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.browser_first_count, 1);
assert.equal(binReport.native_first_count, 1);
assert.equal(binReport.rows.find((row) => row.platform === 'google_meet').recommended_first_surface, 'browser_extension');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').recommended_first_surface, 'native_detector');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').first_surface_family, 'host_native_observer');
assert.equal(binReport.written_files.length, 0);

console.log('ok meeting platform adapter portfolio script');
