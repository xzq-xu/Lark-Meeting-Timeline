import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-app-connector-package-'));
const outDir = join(tmpDir, 'connector');
const reportFile = join(tmpDir, 'report.json');

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-connector-package.mjs',
  '--platforms=google-meet,zoom',
  '--surfaces=browser-extension,native-detector',
  '--base-url=https://timeline.example.com',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
], {
  cwd: repoRoot,
});

assert.match(stdout, /meeting_app_timeline_connector_package_report/);
assert.match(stdout, /ok=yes/);
assert.match(stdout, /platforms=2/);
assert.match(stdout, /surfaces=2/);
assert.match(stdout, /extension=yes/);
assert.match(stdout, /runtime_actions=36/);
assert.match(stdout, /blueprint_ready=2/);
assert.match(stdout, /startup_ready=2/);

const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_app_timeline_connector_package_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 2);
assert.equal(report.surface_count, 2);
assert.equal(report.handoff_count, 4);
assert.equal(report.ready_count, 4);
assert.equal(report.extension_scaffold, true);
assert.equal(report.extension_accepted, true);
assert.equal(report.runtime_event_action_count, 36);
assert.equal(report.adapter_blueprint_ready_count, 2);
assert.equal(report.startup_plan_ready_count, 2);
assert.equal(report.observer_surface_count, 2);
assert.equal(report.scheduler_surface_count, 2);
assert.equal(report.written_files.length, 19);
assert.equal(report.rows.some((row) => row.platform === 'google_meet' && row.surface === 'browser_extension'), true);
assert.equal(report.rows.some((row) => row.platform === 'zoom' && row.surface === 'native_detector'), true);
assert.equal(report.handoff.schema, 'meeting_app_timeline_connector_handoff');
assert.equal(report.handoff.startup_plans.rows.some((row) => row.platform === 'zoom' && row.selected_surface === 'native_detector'), true);
assert.equal(report.package.extension.scaffold.files.some((file) => 'content' in file), false);

const connectorPackage = JSON.parse(await readFile(join(outDir, 'connector-package.json'), 'utf8'));
assert.equal(connectorPackage.schema, 'meeting_app_timeline_connector_package');
assert.equal(connectorPackage.extension.scaffold.files.some((file) => 'content' in file), false);
assert.equal(connectorPackage.contracts.timestamp_field, 'captured_at_ms');
assert.equal(connectorPackage.contracts.provider_events_block_realtime, false);

const hostPackage = JSON.parse(await readFile(join(outDir, 'host-package.json'), 'utf8'));
assert.equal(hostPackage.schema, 'meeting_app_runtime_adapter_host_package');
assert.equal(hostPackage.handoff_count, 4);

const runtimeEventPlanMatrix = JSON.parse(await readFile(join(outDir, 'runtime-event-plan-matrix.json'), 'utf8'));
assert.equal(runtimeEventPlanMatrix.schema, 'meeting_platform_runtime_event_plan_matrix');
assert.equal(runtimeEventPlanMatrix.platform_count, 2);

const adapterBlueprintMatrix = JSON.parse(await readFile(join(outDir, 'adapter-blueprint-matrix.json'), 'utf8'));
assert.equal(adapterBlueprintMatrix.schema, 'meeting_platform_adapter_blueprint_matrix');
assert.equal(adapterBlueprintMatrix.platform_count, 2);
assert.equal(adapterBlueprintMatrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');

const startupPlanMatrix = JSON.parse(await readFile(join(outDir, 'startup-plan-matrix.json'), 'utf8'));
assert.equal(startupPlanMatrix.schema, 'meeting_platform_adapter_startup_plan_matrix');
assert.equal(startupPlanMatrix.platform_count, 2);
assert.equal(startupPlanMatrix.realtime_startup_ready_count, 2);
assert.equal(startupPlanMatrix.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(startupPlanMatrix.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');

const connectorHandoff = JSON.parse(await readFile(join(outDir, 'connector-handoff.json'), 'utf8'));
assert.equal(connectorHandoff.schema, 'meeting_app_timeline_connector_handoff');
assert.equal(connectorHandoff.startup_plans.realtime_startup_ready_count, 2);
assert.equal(connectorHandoff.adapter_blueprints.ready_count, 2);

const browserObserverPlan = JSON.parse(await readFile(join(outDir, 'observer-plan-browser_extension.json'), 'utf8'));
assert.equal(browserObserverPlan.schema, 'meeting_app_runtime_observer_plan_matrix');
assert.equal(browserObserverPlan.platform_count, 2);
assert.equal(browserObserverPlan.rows.every((row) => row.observer_factory === 'createMeetingAppBrowserRuntime'), true);

const nativeScheduler = JSON.parse(await readFile(join(outDir, 'scheduler-config-native_detector.json'), 'utf8'));
assert.equal(nativeScheduler.schema, 'meeting_app_observer_scheduler_config_matrix');
assert.equal(nativeScheduler.track_enabled_count, 2);
assert.equal(nativeScheduler.rows.every((row) => row.runtime_factory === 'createMeetingAppTrackRuntime'), true);

const manifest = JSON.parse(await readFile(join(outDir, 'extension', 'manifest.json'), 'utf8'));
assert.equal(manifest.manifest_version, 3);
assert.equal(manifest.permissions.includes('tabs'), true);
assert.equal(manifest.host_permissions.includes('https://meet.google.com/*'), true);
assert.equal(manifest.host_permissions.includes('https://zoom.us/*'), true);
assert.equal(manifest.host_permissions.includes('https://timeline.example.com/*'), true);

const contentEntry = await readFile(join(outDir, 'extension', 'src', 'content-script.entry.mjs'), 'utf8');
assert.match(contentEntry, /installMeetingPlatformIntegrationContentScriptBridge/);

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-app-connector-package.mjs',
  '--platforms=google-meet',
  '--surfaces=browser-extension',
  '--json=true',
], {
  cwd: repoRoot,
});
const jsonReport = JSON.parse(jsonStdout);
assert.equal(jsonReport.ok, true);
assert.equal(jsonReport.package.platform_count, 1);
assert.equal(jsonReport.package.extension.scaffold.files.some((file) => 'content' in file), false);

console.log('ok meeting app connector package script');
