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
assert.equal(report.written_files.length, 30);
assert.equal(report.rows.some((row) => row.platform === 'google_meet' && row.surface === 'browser_extension'), true);
assert.equal(report.rows.some((row) => row.platform === 'zoom' && row.surface === 'native_detector'), true);
assert.equal(report.handoff.schema, 'meeting_app_timeline_connector_handoff');
assert.equal(report.handoff.startup_plans.rows.some((row) => row.platform === 'zoom' && row.selected_surface === 'native_detector'), true);
assert.equal(report.adoption_index.schema, 'meeting_app_timeline_connector_adoption_index');
assert.equal(report.adoption_index.accepted, true);
assert.equal(report.adoption_index.realtime_ready_count, 2);
assert.equal(report.adoption_index.bridge_ready_count, 2);
assert.equal(report.field_intake_index.schema, 'meeting_app_timeline_connector_field_intake_index');
assert.equal(report.field_intake_index.accepted, true);
assert.equal(report.field_intake_index.rows.find((row) => row.platform === 'google_meet').provider_endpoint, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(report.bridge_handoff.schema, 'meeting_app_timeline_connector_bridge_handoff');
assert.equal(report.bridge_handoff.accepted, true);
assert.equal(report.bridge_handoff.factories.install_content_script_bridge, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(report.bridge_handoff_acceptance.schema, 'meeting_app_timeline_connector_bridge_handoff_acceptance_report');
assert.equal(report.bridge_handoff_acceptance.accepted, true);
assert.equal(report.bridge_smoke_report.schema, 'meeting_app_timeline_connector_bridge_smoke_report');
assert.equal(report.bridge_smoke_report.accepted, true);
assert.equal(report.bridge_smoke_report.runtime_event_actions.includes('observe_platform_candidates'), true);
assert.equal(report.bridge_smoke_report.runtime_event_actions.includes('insert_annotation'), true);
assert.equal(report.host_install_checklist.schema, 'meeting_app_timeline_connector_host_install_checklist');
assert.equal(report.host_install_checklist.rows.some((row) => row.platform === 'zoom' && row.selected_surface === 'native_detector'), true);
assert.equal(report.host_install_checklist_acceptance.schema, 'meeting_app_timeline_connector_host_install_checklist_acceptance_report');
assert.equal(report.host_install_checklist_acceptance.accepted, true);
assert.equal(report.smoke_plan.schema, 'meeting_app_timeline_connector_smoke_plan');
assert.equal(report.smoke_plan.accepted, true);
assert.equal(report.smoke_plan.rows.find((row) => row.platform === 'google_meet').steps.some((step) => step.action === 'observe_platform_candidates'), true);
assert.equal(report.smoke_plan.rows.find((row) => row.platform === 'zoom').steps.some((step) => step.action === 'insert_annotation'), true);
assert.equal(report.smoke_plan_acceptance.schema, 'meeting_app_timeline_connector_smoke_plan_acceptance_report');
assert.equal(report.smoke_plan_acceptance.accepted, true);
assert.equal(report.smoke_run_report.schema, 'meeting_app_timeline_connector_smoke_run_report');
assert.equal(report.smoke_run_report.accepted, true);
assert.equal(report.smoke_run_report.dry_run, true);
assert.equal(report.smoke_run_report.executed_step_count, 8);
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

const connectorAdoptionIndex = JSON.parse(await readFile(join(outDir, 'connector-adoption-index.json'), 'utf8'));
assert.equal(connectorAdoptionIndex.schema, 'meeting_app_timeline_connector_adoption_index');
assert.equal(connectorAdoptionIndex.accepted, true);
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'zoom').status, 'pilot_ready_needs_live_evidence');
assert.equal(connectorAdoptionIndex.rows.every((row) => row.production_evidence_required.includes('runtime_host_replay')), true);
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').field_intake.field_evidence_input.endsWith('/meeting-platform-field-evidence/google_meet.json'), true);

const connectorFieldIntakeIndex = JSON.parse(await readFile(join(outDir, 'connector-field-intake-index.json'), 'utf8'));
assert.equal(connectorFieldIntakeIndex.schema, 'meeting_app_timeline_connector_field_intake_index');
assert.equal(connectorFieldIntakeIndex.accepted, true);
assert.equal(connectorFieldIntakeIndex.rows.find((row) => row.platform === 'google_meet').required_provider_coverage.includes('meeting_start'), true);
assert.equal(connectorFieldIntakeIndex.rows.find((row) => row.platform === 'zoom').commands.build_field_evidence.includes('meeting-platform:field-evidence'), true);

const connectorBridgeHandoff = JSON.parse(await readFile(join(outDir, 'connector-bridge-handoff.json'), 'utf8'));
assert.equal(connectorBridgeHandoff.schema, 'meeting_app_timeline_connector_bridge_handoff');
assert.equal(connectorBridgeHandoff.accepted, true);
assert.equal(connectorBridgeHandoff.message_contract.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(connectorBridgeHandoff.message_contract.message_types.includes('meeting_timeline.observe_candidates'), true);
assert.equal(connectorBridgeHandoff.host_requirements.timestamp_field, 'captured_at_ms');
assert.equal(connectorBridgeHandoff.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');

const connectorBridgeHandoffAcceptance = JSON.parse(await readFile(join(outDir, 'connector-bridge-handoff-acceptance.json'), 'utf8'));
assert.equal(connectorBridgeHandoffAcceptance.schema, 'meeting_app_timeline_connector_bridge_handoff_acceptance_report');
assert.equal(connectorBridgeHandoffAcceptance.accepted, true);
assert.equal(connectorBridgeHandoffAcceptance.required_message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(connectorBridgeHandoffAcceptance.required_output_runtime_actions.includes('observe_platform_candidates'), true);

const connectorBridgeSmokeReport = JSON.parse(await readFile(join(outDir, 'connector-bridge-smoke-report.json'), 'utf8'));
assert.equal(connectorBridgeSmokeReport.schema, 'meeting_app_timeline_connector_bridge_smoke_report');
assert.equal(connectorBridgeSmokeReport.accepted, true);
assert.equal(connectorBridgeSmokeReport.observe_before_insert, true);
assert.equal(connectorBridgeSmokeReport.calls.some((call) => call.action === 'insert_annotation'), true);

const hostInstallChecklist = JSON.parse(await readFile(join(outDir, 'host-install-checklist.json'), 'utf8'));
assert.equal(hostInstallChecklist.schema, 'meeting_app_timeline_connector_host_install_checklist');
assert.equal(hostInstallChecklist.accepted, true);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').client_methods.insert_annotation, 'insertAnnotation');

const hostInstallChecklistAcceptance = JSON.parse(await readFile(join(outDir, 'host-install-checklist-acceptance.json'), 'utf8'));
assert.equal(hostInstallChecklistAcceptance.schema, 'meeting_app_timeline_connector_host_install_checklist_acceptance_report');
assert.equal(hostInstallChecklistAcceptance.accepted, true);
assert.equal(hostInstallChecklistAcceptance.ready_count, 2);
assert.equal(hostInstallChecklistAcceptance.issue_count, 0);

const smokePlan = JSON.parse(await readFile(join(outDir, 'connector-smoke-plan.json'), 'utf8'));
assert.equal(smokePlan.schema, 'meeting_app_timeline_connector_smoke_plan');
assert.equal(smokePlan.accepted, true);
assert.equal(smokePlan.timestamp_field, 'captured_at_ms');
assert.equal(smokePlan.rows.find((row) => row.platform === 'google_meet').steps.find((step) => step.action === 'observe_platform_candidates').input.captured_at_ms, 1_782_614_400_000);
assert.equal(smokePlan.rows.find((row) => row.platform === 'zoom').steps.find((step) => step.action === 'insert_annotation').input.annotation.captured_at_ms, 1_782_614_411_000);

const smokePlanAcceptance = JSON.parse(await readFile(join(outDir, 'connector-smoke-plan-acceptance.json'), 'utf8'));
assert.equal(smokePlanAcceptance.schema, 'meeting_app_timeline_connector_smoke_plan_acceptance_report');
assert.equal(smokePlanAcceptance.accepted, true);
assert.equal(smokePlanAcceptance.required_step_count, 4);
assert.equal(smokePlanAcceptance.optional_step_count, 4);

const smokeRunReport = JSON.parse(await readFile(join(outDir, 'connector-smoke-run-report.json'), 'utf8'));
assert.equal(smokeRunReport.schema, 'meeting_app_timeline_connector_smoke_run_report');
assert.equal(smokeRunReport.accepted, true);
assert.equal(smokeRunReport.dry_run, true);
assert.equal(smokeRunReport.call_count, 8);
assert.equal(smokeRunReport.rows.every((row) => row.observe_before_insert === true), true);
assert.equal(smokeRunReport.rows.every((row) => row.captured_at_ms_preserved === true), true);

const connectorQuickstart = await readFile(join(outDir, 'connector-quickstart.md'), 'utf8');
assert.match(connectorQuickstart, /Meeting App Timeline Connector Quickstart/);
assert.match(connectorQuickstart, /connector-adoption-index\.json/);
assert.match(connectorQuickstart, /connector-field-intake-index\.json/);
assert.match(connectorQuickstart, /connector-bridge-handoff\.json/);
assert.match(connectorQuickstart, /connector-bridge-handoff-acceptance\.json/);
assert.match(connectorQuickstart, /connector-bridge-smoke-report\.json/);
assert.match(connectorQuickstart, /host-install-checklist\.json/);
assert.match(connectorQuickstart, /host-install-checklist-acceptance\.json/);
assert.match(connectorQuickstart, /connector-smoke-plan\.json/);
assert.match(connectorQuickstart, /connector-smoke-plan-acceptance\.json/);
assert.match(connectorQuickstart, /connector-smoke-run-report\.json/);
assert.match(connectorQuickstart, /startup-plan-matrix\.json/);
assert.match(connectorQuickstart, /captured_at_ms/);
assert.match(connectorQuickstart, /createMeetingAppTimelineConnectorRuntimeClient/);
assert.match(connectorQuickstart, /google_meet/);
assert.match(connectorQuickstart, /native_detector/);

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
