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
assert.match(stdout, /release_gate=yes/);
assert.match(stdout, /adapter_matrix=yes/);
assert.match(stdout, /host_configs=yes/);
assert.match(stdout, /host_bootstrap=yes/);
assert.match(stdout, /provider_replay=yes/);
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
assert.equal(report.provider_replay_accepted, true);
assert.equal(report.provider_replay_accepted_count, 2);
assert.equal(report.provider_replay_runtime_event_count, 8);
assert.equal(report.adapter_blueprint_ready_count, 2);
assert.equal(report.startup_plan_ready_count, 2);
assert.equal(report.observer_surface_count, 2);
assert.equal(report.scheduler_surface_count, 2);
assert.equal(report.release_gate_accepted, true);
assert.equal(report.release_gate_target, 'pilot');
assert.equal(report.adapter_matrix_accepted, true);
assert.equal(report.host_adapter_config_accepted, true);
assert.equal(report.host_adapter_config_count, 2);
assert.equal(report.host_adapter_bootstrap_accepted, true);
assert.equal(report.host_adapter_bootstrap_count, 2);
assert.equal(report.host_adapter_bootstrap_issue_count, 0);
assert.equal(report.platform_onboarding_accepted, true);
assert.equal(report.platform_onboarding_target, 'static');
assert.equal(report.written_files.length, 41);
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
assert.equal(report.provider_replay.matrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(report.provider_replay.accepted, true);
assert.equal(report.provider_replay.matrix.rows.every((row) => row.provider_events_block_realtime === false), true);
assert.equal(report.release_gate.schema, 'meeting_app_timeline_connector_release_gate');
assert.equal(report.release_gate.accepted, true);
assert.equal(report.release_gate.pilot_ready_count, 2);
assert.equal(report.release_gate.production_ready_count, 0);
assert.equal(report.release_gate.required_gates.find((gate) => gate.id === 'smoke_run_report').accepted, true);
assert.equal(report.release_gate.required_gates.find((gate) => gate.id === 'host_adapter_bootstrap_plan_matrix_acceptance').accepted, true);
assert.equal(report.release_gate.files_to_read_first.includes('host-adapter-bootstrap-plan-matrix-acceptance.json'), true);
assert.equal(report.platform_roadmap.schema, 'meeting_app_timeline_connector_platform_roadmap');
assert.equal(report.platform_roadmap.accepted, true);
assert.equal(report.platform_roadmap.rows.find((row) => row.platform === 'google_meet').recommended_first_surface, 'browser_extension');
assert.equal(report.platform_roadmap.rows.find((row) => row.platform === 'zoom').recommended_first_surface, 'native_detector');
assert.equal(report.adapter_matrix.schema, 'meeting_app_timeline_connector_adapter_matrix');
assert.equal(report.adapter_matrix.accepted, true);
assert.equal(report.adapter_matrix.rows.find((row) => row.platform === 'google_meet').adapter_mode, 'browser_content_script');
assert.equal(report.adapter_matrix.rows.find((row) => row.platform === 'zoom').adapter_mode, 'native_or_desktop_observer');
assert.equal(report.adapter_matrix.rows.every((row) => row.runtime_sequence[0].action === 'observe_platform_candidates'), true);
assert.equal(report.host_adapter_config_index.schema, 'meeting_app_timeline_host_adapter_config_index');
assert.equal(report.host_adapter_config_index.accepted, true);
assert.equal(report.host_adapter_config_index.rows.find((row) => row.platform === 'google_meet').config_file, 'host-adapter-configs/google_meet.json');
assert.equal(report.host_adapter_bootstrap_plan_matrix.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix');
assert.equal(report.host_adapter_bootstrap_plan_matrix.accepted, true);
assert.equal(report.host_adapter_bootstrap_plan_matrix.plans.google_meet.startup_order[3], 'runtime_observe_platform_candidates');
assert.equal(report.host_adapter_bootstrap_plan_matrix_acceptance.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix_acceptance_report');
assert.equal(report.host_adapter_bootstrap_plan_matrix_acceptance.accepted, true);
assert.equal(report.host_adapter_bootstrap_plan_matrix_acceptance.rows.every((row) => row.startup_order_ready === true), true);
assert.equal(report.adapter_matrix_acceptance.schema, 'meeting_app_timeline_connector_adapter_matrix_acceptance_report');
assert.equal(report.adapter_matrix_acceptance.accepted, true);
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

const providerReplayMatrix = JSON.parse(await readFile(join(outDir, 'provider-replay-matrix.json'), 'utf8'));
assert.equal(providerReplayMatrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(providerReplayMatrix.accepted, true);
assert.equal(providerReplayMatrix.accepted_count, 2);
assert.equal(providerReplayMatrix.rows.every((row) => row.provider_events_block_realtime === false), true);

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
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').adapter_preflight.url_only_status, 'needs_live_page_evidence');
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').local_observer_contract.observer_mode, 'browser_dom_observer');
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').local_observer_runtime_wiring.runtime_factory, 'createMeetingAppBrowserRuntime');
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'zoom').status, 'pilot_ready_needs_live_evidence');
assert.equal(connectorAdoptionIndex.rows.every((row) => row.production_evidence_required.includes('runtime_host_replay')), true);
assert.equal(connectorAdoptionIndex.rows.find((row) => row.platform === 'google_meet').field_intake.field_evidence_input.endsWith('/meeting-platform-field-evidence/google_meet.json'), true);

const connectorFieldIntakeIndex = JSON.parse(await readFile(join(outDir, 'connector-field-intake-index.json'), 'utf8'));
assert.equal(connectorFieldIntakeIndex.schema, 'meeting_app_timeline_connector_field_intake_index');
assert.equal(connectorFieldIntakeIndex.accepted, true);
assert.equal(connectorFieldIntakeIndex.rows.find((row) => row.platform === 'google_meet').required_provider_coverage.includes('meeting_start'), true);
assert.equal(connectorFieldIntakeIndex.rows.find((row) => row.platform === 'zoom').commands.build_field_evidence.includes('meeting-platform:field-evidence'), true);

const connectorReleaseGate = JSON.parse(await readFile(join(outDir, 'connector-release-gate.json'), 'utf8'));
assert.equal(connectorReleaseGate.schema, 'meeting_app_timeline_connector_release_gate');
assert.equal(connectorReleaseGate.accepted, true);
assert.equal(connectorReleaseGate.target, 'pilot');
assert.equal(connectorReleaseGate.pilot_ready_count, 2);
assert.equal(connectorReleaseGate.production_ready_count, 0);
assert.equal(connectorReleaseGate.rows.every((row) => row.pilot_ready === true), true);
assert.equal(connectorReleaseGate.required_gates.find((gate) => gate.id === 'bridge_smoke_report').accepted, true);
assert.equal(connectorReleaseGate.required_gates.find((gate) => gate.id === 'host_adapter_bootstrap_plan_matrix_acceptance').accepted, true);
assert.equal(connectorReleaseGate.files_to_read_first.includes('connector-release-gate.json'), true);
assert.equal(connectorReleaseGate.files_to_read_first.includes('host-adapter-bootstrap-plan-matrix-acceptance.json'), true);

const connectorPlatformRoadmap = JSON.parse(await readFile(join(outDir, 'connector-platform-roadmap.json'), 'utf8'));
assert.equal(connectorPlatformRoadmap.schema, 'meeting_app_timeline_connector_platform_roadmap');
assert.equal(connectorPlatformRoadmap.accepted, true);
assert.equal(connectorPlatformRoadmap.recommended_first_platform, 'google_meet');
assert.equal(connectorPlatformRoadmap.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(connectorPlatformRoadmap.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');
assert.equal(connectorPlatformRoadmap.files_to_read_first.includes('connector-platform-roadmap.json'), true);

const connectorAdapterMatrix = JSON.parse(await readFile(join(outDir, 'connector-adapter-matrix.json'), 'utf8'));
assert.equal(connectorAdapterMatrix.schema, 'meeting_app_timeline_connector_adapter_matrix');
assert.equal(connectorAdapterMatrix.accepted, true);
assert.equal(connectorAdapterMatrix.recommended_first_platform, 'google_meet');
assert.equal(connectorAdapterMatrix.runtime_invariants.mark_timestamp_field, 'captured_at_ms');
assert.equal(connectorAdapterMatrix.runtime_invariants.provider_replay_blocks_realtime, false);
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').evidence_contract.adapter_preflight.evidence_kind, 'live_dom_snapshot');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').evidence_contract.pilot_required.includes('adapter_preflight_live_evidence'), true);
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').local_observer_contract.input_contract.kind, 'browser_live_dom');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').evidence_contract.local_observer.filters.speaker.min_stable_ms, 700);
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').local_observer_runtime_wiring.client_methods.insert_annotation, 'insertAnnotation');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').evidence_contract.local_observer_runtime_wiring.bridge_messages.observe_meeting_app, 'meeting_timeline.sample');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').runtime_sequence[1].required_field, 'captured_at_ms');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').provider_replay.accepted, true);
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'google_meet').provider_replay.file, 'provider-replay-matrix.json');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'zoom').install_step, 'install_native_desktop_observer_or_accessibility_detector');
assert.equal(connectorAdapterMatrix.rows.find((row) => row.platform === 'zoom').provider_replay.provider_events_block_realtime, false);
assert.equal(connectorAdapterMatrix.files_to_read_first.includes('connector-adapter-matrix.json'), true);
assert.equal(connectorAdapterMatrix.files_to_read_first.includes('provider-replay-matrix.json'), true);

const hostAdapterConfigIndex = JSON.parse(await readFile(join(outDir, 'host-adapter-config-index.json'), 'utf8'));
assert.equal(hostAdapterConfigIndex.schema, 'meeting_app_timeline_host_adapter_config_index');
assert.equal(hostAdapterConfigIndex.accepted, true);
assert.equal(hostAdapterConfigIndex.row_count, 2);
assert.equal(hostAdapterConfigIndex.rows.find((row) => row.platform === 'zoom').config_file, 'host-adapter-configs/zoom.json');
assert.equal(hostAdapterConfigIndex.rows.find((row) => row.platform === 'google_meet').local_observer_mode, 'browser_dom_observer');
assert.equal(hostAdapterConfigIndex.rows.find((row) => row.platform === 'zoom').local_observer_mode, 'native_window_observer');
assert.equal(hostAdapterConfigIndex.rows.find((row) => row.platform === 'google_meet').local_observer_runtime_factory, 'createMeetingAppBrowserRuntime');

const hostAdapterBootstrapPlanMatrix = JSON.parse(await readFile(join(outDir, 'host-adapter-bootstrap-plan-matrix.json'), 'utf8'));
assert.equal(hostAdapterBootstrapPlanMatrix.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix');
assert.equal(hostAdapterBootstrapPlanMatrix.accepted, true);
assert.equal(hostAdapterBootstrapPlanMatrix.accepted_count, 2);
assert.equal(hostAdapterBootstrapPlanMatrix.rows.find((row) => row.platform === 'google_meet').install_target, 'manifest_v3_content_script');
assert.equal(hostAdapterBootstrapPlanMatrix.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');
assert.equal(hostAdapterBootstrapPlanMatrix.plans.zoom.startup_order.includes('runtime_insert_annotation'), true);

const hostAdapterBootstrapPlanMatrixAcceptance = JSON.parse(await readFile(join(outDir, 'host-adapter-bootstrap-plan-matrix-acceptance.json'), 'utf8'));
assert.equal(hostAdapterBootstrapPlanMatrixAcceptance.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix_acceptance_report');
assert.equal(hostAdapterBootstrapPlanMatrixAcceptance.accepted, true);
assert.equal(hostAdapterBootstrapPlanMatrixAcceptance.issue_count, 0);
assert.equal(hostAdapterBootstrapPlanMatrixAcceptance.rows.find((row) => row.platform === 'google_meet').startup_order_ready, true);

const googleHostAdapterConfig = JSON.parse(await readFile(join(outDir, 'host-adapter-configs', 'google_meet.json'), 'utf8'));
assert.equal(googleHostAdapterConfig.schema, 'meeting_app_timeline_host_adapter_config');
assert.equal(googleHostAdapterConfig.platform, 'google_meet');
assert.equal(googleHostAdapterConfig.selected_surface, 'browser_extension');
assert.equal(googleHostAdapterConfig.realtime_contract.mark_timestamp_field, 'captured_at_ms');
assert.equal(googleHostAdapterConfig.evidence_contract.adapter_preflight.required, true);
assert.equal(googleHostAdapterConfig.evidence_contract.adapter_preflight.bridge_messages.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(googleHostAdapterConfig.local_observer_contract.sampling.speaker_sample_interval_ms, 300);
assert.equal(googleHostAdapterConfig.local_observer_contract.input_contract.selector_groups.participants.some((selector) => selector.includes('speaking')), true);
assert.equal(googleHostAdapterConfig.local_observer_runtime_wiring.runtime_guards.require_preflight_before_first_insert, true);
assert.equal(googleHostAdapterConfig.local_observer_runtime_wiring.trigger_order.find((step) => step.trigger === 'annotation_captured').required_field, 'captured_at_ms');
assert.equal(googleHostAdapterConfig.provider_replay.accepted, true);
assert.equal(googleHostAdapterConfig.runtime_sequence[0].action, 'observe_platform_candidates');

const zoomHostAdapterConfig = JSON.parse(await readFile(join(outDir, 'host-adapter-configs', 'zoom.json'), 'utf8'));
assert.equal(zoomHostAdapterConfig.platform, 'zoom');
assert.equal(zoomHostAdapterConfig.selected_surface, 'native_detector');
assert.equal(zoomHostAdapterConfig.evidence_contract.adapter_preflight.evidence_kind, 'native_window_or_process_state');
assert.equal(zoomHostAdapterConfig.local_observer_contract.input_contract.required_inputs.includes('candidate_windows'), true);
assert.equal(zoomHostAdapterConfig.local_observer_runtime_wiring.runtime_factory, 'createMeetingAppTrackRuntime');
assert.equal(zoomHostAdapterConfig.provider_replay.provider_events_block_realtime, false);

const connectorAdapterMatrixAcceptance = JSON.parse(await readFile(join(outDir, 'connector-adapter-matrix-acceptance.json'), 'utf8'));
assert.equal(connectorAdapterMatrixAcceptance.schema, 'meeting_app_timeline_connector_adapter_matrix_acceptance_report');
assert.equal(connectorAdapterMatrixAcceptance.accepted, true);
assert.equal(connectorAdapterMatrixAcceptance.issue_count, 0);

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
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').required_host_steps.includes('apply_local_observer_sampling_and_filter_contract'), true);

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
assert.match(connectorQuickstart, /connector-release-gate\.json/);
assert.match(connectorQuickstart, /connector-adapter-matrix\.json/);
assert.match(connectorQuickstart, /connector-adapter-matrix-acceptance\.json/);
assert.match(connectorQuickstart, /platform-onboarding-checklist-matrix\.json/);
assert.match(connectorQuickstart, /Platform Onboarding Matrix/);
assert.match(connectorQuickstart, /host-adapter-bootstrap-plan-matrix\.json/);
assert.match(connectorQuickstart, /host-adapter-bootstrap-plan-matrix-acceptance\.json/);
assert.match(connectorQuickstart, /connector-platform-roadmap\.json/);
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

const platformOnboarding = JSON.parse(await readFile(join(outDir, 'platform-onboarding-checklist-matrix.json'), 'utf8'));
assert.equal(platformOnboarding.schema, 'meeting_platform_adapter_acceptance_checklist_matrix');
assert.equal(platformOnboarding.rows.find((row) => row.platform === 'google_meet').runtime_contract_timestamp_field, 'captured_at_ms');
assert.equal(platformOnboarding.rows.find((row) => row.platform === 'zoom').local_observer_install_step, 'wire_native_window_or_accessibility_detector');

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
