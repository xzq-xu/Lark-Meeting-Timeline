import assert from 'node:assert/strict';

import {
  MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_ADAPTER_MATRIX_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_MATRIX_ACCEPTANCE_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_MATRIX_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_INDEX_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_RESOLUTION_SCHEMA,
  MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_PLATFORM_ROADMAP_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_RELEASE_GATE_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA,
  MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA,
  assertMeetingAppTimelineConnectorBridgeHandoff,
  assertMeetingAppTimelineConnectorAdoptionIndex,
  assertMeetingAppTimelineConnectorAdapterMatrix,
  assertMeetingAppTimelineConnectorFieldIntakeIndex,
  assertMeetingAppTimelineHostAdapterBootstrapPlan,
  assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  assertMeetingAppTimelineHostAdapterConfig,
  assertMeetingAppTimelineHostAdapterConfigIndex,
  assertMeetingAppTimelineResolvedHostAdapterConfig,
  assertMeetingAppTimelineConnectorBridgeSmoke,
  assertMeetingAppTimelineConnectorHostInstallChecklist,
  assertMeetingAppTimelineConnectorPackage,
  assertMeetingAppTimelineConnectorPlatformRoadmap,
  assertMeetingAppTimelineConnectorReleaseGate,
  assertMeetingAppTimelineConnectorSmokePlan,
  assertMeetingAppTimelineConnectorSmokeRun,
  buildMeetingAppTimelineConnectorBridgeHandoff,
  buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport,
  buildMeetingAppTimelineConnectorAdoptionIndex,
  buildMeetingAppTimelineConnectorAdapterMatrix,
  buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport,
  buildMeetingAppTimelineConnectorFieldIntakeIndex,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineHostAdapterBootstrapPlan,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport,
  buildMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineHostAdapterConfigIndex,
  resolveMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineConnectorHostInstallChecklist,
  buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport,
  buildMeetingAppTimelineConnectorPlatformRoadmap,
  buildMeetingAppTimelineConnectorReleaseGate,
  buildMeetingAppTimelineConnectorSmokePlan,
  buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport,
  createMeetingAppTimelineSdk,
  createMeetingAppTimelineConnectorRuntimeClient,
  runMeetingAppTimelineConnectorBridgeSmoke,
  runMeetingAppTimelineConnectorSmokePlan,
  stripMeetingAppTimelineConnectorPackageFileContents,
} from '../packages/meeting-timeline-sdk/index.mjs';
import {
  buildMeetingAppTimelineConnectorPackageAcceptanceReport as buildAcceptanceReportFromSubpath,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-connector-package.mjs';

const sdk = createMeetingAppTimelineSdk({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'zoom'],
});
const connectorPackage = sdk.connectorPackage({
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});

const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage);
assert.equal(acceptance.schema, 'meeting_app_timeline_connector_package_acceptance_report');
assert.equal(acceptance.accepted, true);
assert.equal(acceptance.target, 'realtime');
assert.deepEqual(acceptance.platforms, ['google_meet', 'zoom']);
assert.deepEqual(acceptance.surfaces, ['browser_extension', 'native_detector']);
assert.equal(acceptance.require_tracks, true);
assert.equal(acceptance.issue_count, 0);
assert.equal(acceptance.missing_runtime_actions.length, 0);
assert.equal(acceptance.surface_reports.length, 2);
assert.equal(acceptance.surface_reports.every((row) => row.accepted), true);
assert.equal(acceptance.surface_reports.find((row) => row.surface === 'browser_extension').track_enabled_count, 2);
assert.equal(acceptance.extension.required, true);
assert.equal(acceptance.extension.accepted, true);
assert.equal(acceptance.extension.manifest_version, 3);
assert.equal(connectorPackage.provider_replay.accepted, true);
assert.equal(connectorPackage.provider_replay.matrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(connectorPackage.provider_replay.accepted_count, 2);
assert.equal(connectorPackage.provider_replay.matrix.rows.every((row) => row.provider_events_block_realtime === false), true);

assert.equal(buildAcceptanceReportFromSubpath(connectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorPackage(connectorPackage), connectorPackage);

const handoff = buildMeetingAppTimelineConnectorHandoff(connectorPackage);
assert.equal(handoff.schema, MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA);
assert.equal(handoff.accepted, true);
assert.equal(handoff.runtime_event_endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(handoff.timestamp_field, 'captured_at_ms');
assert.equal(handoff.provider_events_block_realtime, false);
assert.equal(handoff.transcript_blocks_realtime, false);
assert.equal(handoff.provider_replay.accepted, true);
assert.equal(handoff.provider_replay.accepted_count, 2);
assert.equal(handoff.provider_replay.provider_events_block_realtime, false);
assert.equal(handoff.adapter_blueprints.ready_count, 2);
assert.equal(handoff.adapter_blueprints.platform_count, 2);
assert.equal(handoff.adapter_blueprints.sdk_method, 'sdk.platformAdapterBlueprint(platform)');
assert.equal(handoff.adapter_blueprints.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(handoff.startup_plans.realtime_startup_ready_count, 2);
assert.equal(handoff.startup_plans.platform_count, 2);
assert.equal(handoff.startup_plans.sdk_method, 'sdk.platformAdapterStartupPlan(input)');
assert.equal(handoff.startup_plans.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(handoff.startup_plans.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(handoff.surface_matrix.length, 2);
assert.equal(handoff.surface_matrix.find((row) => row.surface === 'browser_extension').ready_count, 2);
assert.equal(handoff.extension.file_paths.includes('manifest.json'), true);
assert.equal(handoff.extension.file_paths.includes('src/content-script.entry.mjs'), true);
assert.equal(handoff.ci_gates.includes('require_captured_at_ms_for_realtime_annotations'), true);

const hostInstallChecklist = buildMeetingAppTimelineConnectorHostInstallChecklist(connectorPackage);
assert.equal(hostInstallChecklist.schema, 'meeting_app_timeline_connector_host_install_checklist');
assert.equal(hostInstallChecklist.accepted, true);
assert.equal(hostInstallChecklist.ready_count, 2);
assert.equal(hostInstallChecklist.timestamp_field, 'captured_at_ms');
assert.equal(hostInstallChecklist.provider_replay.accepted, true);
assert.equal(hostInstallChecklist.provider_replay.rows.find((row) => row.platform === 'google_meet').accepted, true);
assert.equal(hostInstallChecklist.files_to_read_first.includes('provider-replay-matrix.json'), true);
assert.equal(hostInstallChecklist.files_to_read_first.includes('startup-plan-matrix.json'), true);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').client_methods.insert_annotation, 'insertAnnotation');
assert.equal(hostInstallChecklist.contracts.adapter_preflight_required_before_realtime_insert, true);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').adapter_preflight.evidence_kind, 'live_dom_snapshot');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').adapter_preflight.current_window_mode_supported, true);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').adapter_preflight.url_only_status, 'needs_live_page_evidence');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').required_host_steps.includes('run_adapter_preflight_with_live_evidence_before_first_annotation'), true);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').adapter_preflight.evidence_kind, 'native_window_or_process_state');
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'zoom').adapter_preflight.current_window_mode_supported, false);
assert.equal(hostInstallChecklist.rows.find((row) => row.platform === 'google_meet').required_host_steps.includes('insert_annotation_with_captured_at_ms'), true);
const hostInstallChecklistAcceptance = buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(hostInstallChecklist);
assert.equal(hostInstallChecklistAcceptance.schema, 'meeting_app_timeline_connector_host_install_checklist_acceptance_report');
assert.equal(hostInstallChecklistAcceptance.accepted, true);
assert.equal(hostInstallChecklistAcceptance.ready_count, 2);
assert.equal(hostInstallChecklistAcceptance.issue_count, 0);
assert.equal(hostInstallChecklistAcceptance.rows.find((row) => row.platform === 'google_meet').adapter_preflight_required, true);
assert.equal(hostInstallChecklistAcceptance.rows.find((row) => row.platform === 'google_meet').adapter_preflight_url_only_status, 'needs_live_page_evidence');
assert.equal(assertMeetingAppTimelineConnectorHostInstallChecklist(hostInstallChecklist), hostInstallChecklist);
assert.equal(buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(connectorPackage).accepted, true);

const adoptionIndex = buildMeetingAppTimelineConnectorAdoptionIndex(hostInstallChecklist);
assert.equal(adoptionIndex.schema, MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA);
assert.equal(adoptionIndex.accepted, true);
assert.equal(adoptionIndex.realtime_ready_count, 2);
assert.equal(adoptionIndex.bridge_ready_count, 2);
assert.equal(adoptionIndex.production_evidence_pending_count, 2);
assert.equal(adoptionIndex.rows.find((row) => row.platform === 'google_meet').status, 'pilot_ready_needs_live_evidence');
assert.equal(adoptionIndex.rows.find((row) => row.platform === 'google_meet').p0_axis_bootstrap.must_precede, 'insert_annotation');
assert.equal(adoptionIndex.rows.find((row) => row.platform === 'google_meet').adapter_preflight.blocks_realtime_if_missing, true);
assert.equal(adoptionIndex.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(adoptionIndex.rows.every((row) => row.can_start_axis_before_provider === true), true);
assert.equal(adoptionIndex.rows.every((row) => row.production_evidence_required.includes('runtime_host_replay')), true);
assert.equal(assertMeetingAppTimelineConnectorAdoptionIndex(connectorPackage).accepted, true);

const fieldIntakeIndex = buildMeetingAppTimelineConnectorFieldIntakeIndex(hostInstallChecklist);
assert.equal(fieldIntakeIndex.schema, MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA);
assert.equal(fieldIntakeIndex.accepted, true);
assert.equal(fieldIntakeIndex.platform_count, 2);
assert.equal(fieldIntakeIndex.rows.find((row) => row.platform === 'google_meet').provider_endpoint, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(fieldIntakeIndex.rows.find((row) => row.platform === 'google_meet').field_evidence_input.endsWith('/meeting-platform-field-evidence/google_meet.json'), true);
assert.equal(fieldIntakeIndex.rows.find((row) => row.platform === 'zoom').commands.validate_real_intake.includes('meeting-platform:real-intake'), true);
assert.equal(fieldIntakeIndex.rows.every((row) => row.required_local_snapshots.includes('active_speaker')), true);
assert.equal(fieldIntakeIndex.next_actions.includes('capture_real_provider_start_end_events'), true);
assert.equal(assertMeetingAppTimelineConnectorFieldIntakeIndex(connectorPackage).accepted, true);
assert.equal(adoptionIndex.rows.find((row) => row.platform === 'google_meet').field_intake.commands.build_field_evidence.includes('meeting-platform:field-evidence'), true);

const bridgeHandoff = buildMeetingAppTimelineConnectorBridgeHandoff(connectorPackage);
assert.equal(bridgeHandoff.schema, MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA);
assert.equal(bridgeHandoff.accepted, true);
assert.equal(bridgeHandoff.platform_count, 2);
assert.equal(bridgeHandoff.module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(bridgeHandoff.factories.install_content_script_bridge, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(bridgeHandoff.message_contract.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(bridgeHandoff.message_contract.message_types.includes('meeting_timeline.observe_candidates'), true);
assert.equal(bridgeHandoff.host_requirements.timestamp_field, 'captured_at_ms');
assert.equal(bridgeHandoff.host_requirements.provider_events_block_realtime, false);
assert.equal(bridgeHandoff.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(bridgeHandoff.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');
assert.equal(bridgeHandoff.sample_messages[0].payload.captured_at_ms, 1_782_614_400_000);
assert.equal(bridgeHandoff.issue_count, 0);
const bridgeHandoffAcceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(bridgeHandoff);
assert.equal(bridgeHandoffAcceptance.schema, 'meeting_app_timeline_connector_bridge_handoff_acceptance_report');
assert.equal(bridgeHandoffAcceptance.accepted, true);
assert.equal(bridgeHandoffAcceptance.required_message_types.includes('meeting_timeline.observe_candidates'), true);
assert.equal(bridgeHandoffAcceptance.required_output_runtime_actions.includes('insert_annotation'), true);
assert.equal(bridgeHandoffAcceptance.issue_count, 0);
assert.equal(assertMeetingAppTimelineConnectorBridgeHandoff(bridgeHandoff), bridgeHandoff);
assert.equal(buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(connectorPackage).accepted, true);
const bridgeSmokeReport = await runMeetingAppTimelineConnectorBridgeSmoke(bridgeHandoff);
assert.equal(bridgeSmokeReport.schema, 'meeting_app_timeline_connector_bridge_smoke_report');
assert.equal(bridgeSmokeReport.accepted, true);
assert.equal(bridgeSmokeReport.bridge_handoff_acceptance_accepted, true);
assert.equal(bridgeSmokeReport.observe_before_insert, true);
assert.equal(bridgeSmokeReport.runtime_event_actions.includes('observe_platform_candidates'), true);
assert.equal(bridgeSmokeReport.runtime_event_actions.includes('insert_annotation'), true);
assert.equal(bridgeSmokeReport.steps.find((step) => step.id === 'insert_mark_message').handled, true);
assert.equal(bridgeSmokeReport.steps.find((step) => step.id === 'sample_tracks_message').accepted, true);
assert.equal(bridgeSmokeReport.calls.find((call) => call.action === 'insert_annotation').annotation_id, 'google_meet-bridge-smoke-mark');
assert.equal((await assertMeetingAppTimelineConnectorBridgeSmoke(bridgeHandoff)).accepted, true);

const smokePlan = buildMeetingAppTimelineConnectorSmokePlan(hostInstallChecklist, {
  baseCapturedAtMs: 1_782_614_400_000,
});
assert.equal(smokePlan.schema, MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA);
assert.equal(smokePlan.accepted, true);
assert.equal(smokePlan.platform_count, 2);
assert.equal(smokePlan.row_count, 2);
assert.equal(smokePlan.timestamp_field, 'captured_at_ms');
assert.equal(smokePlan.rows.every((row) => row.required_action_count === 2), true);
assert.equal(smokePlan.rows.every((row) => row.optional_action_count === 2), true);
const googleSmokeRow = smokePlan.rows.find((row) => row.platform === 'google_meet');
const zoomSmokeRow = smokePlan.rows.find((row) => row.platform === 'zoom');
const googleObserveSmokeStep = googleSmokeRow.steps.find((step) => step.action === 'observe_platform_candidates');
const zoomInsertSmokeStep = zoomSmokeRow.steps.find((step) => step.action === 'insert_annotation');
assert.equal(googleObserveSmokeStep.client_method, 'observePlatformCandidates');
assert.equal(googleObserveSmokeStep.input.captured_at_ms, 1_782_614_400_000);
assert.equal(googleObserveSmokeStep.input.tabs[0].url, 'https://meet.google.com/abc-defg-hij');
assert.equal(zoomInsertSmokeStep.client_method, 'insertAnnotation');
assert.equal(zoomInsertSmokeStep.input.annotation.captured_at_ms, 1_782_614_411_000);
const smokePlanAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(smokePlan);
assert.equal(smokePlanAcceptance.schema, 'meeting_app_timeline_connector_smoke_plan_acceptance_report');
assert.equal(smokePlanAcceptance.accepted, true);
assert.equal(smokePlanAcceptance.required_step_count, 4);
assert.equal(smokePlanAcceptance.optional_step_count, 4);
assert.equal(assertMeetingAppTimelineConnectorSmokePlan(smokePlan), smokePlan);
assert.equal(buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(connectorPackage).accepted, true);

const smokeRunReport = await runMeetingAppTimelineConnectorSmokePlan(smokePlan);
assert.equal(smokeRunReport.schema, MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA);
assert.equal(smokeRunReport.accepted, true);
assert.equal(smokeRunReport.dry_run, true);
assert.equal(smokeRunReport.plan_accepted, true);
assert.equal(smokeRunReport.required_step_count, 4);
assert.equal(smokeRunReport.optional_step_count, 4);
assert.equal(smokeRunReport.executed_step_count, 8);
assert.equal(smokeRunReport.failed_step_count, 0);
assert.equal(smokeRunReport.call_count, 8);
assert.equal(smokeRunReport.rows.every((row) => row.observe_before_insert === true), true);
assert.equal(smokeRunReport.rows.every((row) => row.captured_at_ms_preserved === true), true);
assert.deepEqual(smokeRunReport.calls.slice(0, 2).map((call) => call.method), ['observePlatformCandidates', 'insertAnnotation']);
const assertedSmokeRun = await assertMeetingAppTimelineConnectorSmokeRun(smokePlan);
assert.equal(assertedSmokeRun.accepted, true);

const releaseGate = buildMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance,
  bridgeSmokeReport,
  hostInstallChecklistAcceptance,
  smokePlan,
  smokePlanAcceptance,
  smokeRunReport,
});
assert.equal(releaseGate.schema, MEETING_APP_TIMELINE_CONNECTOR_RELEASE_GATE_SCHEMA);
assert.equal(releaseGate.accepted, true);
assert.equal(releaseGate.target, 'pilot');
assert.equal(releaseGate.pilot_ready_count, 2);
assert.equal(releaseGate.production_ready_count, 0);
assert.equal(releaseGate.realtime_ready_count, 2);
assert.equal(releaseGate.bridge_ready_count, 2);
assert.equal(releaseGate.field_intake_ready_count, 2);
assert.equal(releaseGate.rows.every((row) => row.pilot_ready === true), true);
assert.equal(releaseGate.rows.every((row) => row.production_ready === false), true);
assert.equal(releaseGate.required_gates.find((gate) => gate.id === 'smoke_run_report').accepted, true);
assert.equal(releaseGate.files_to_read_first.includes('connector-release-gate.json'), true);
assert.equal(assertMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance,
  bridgeSmokeReport,
  hostInstallChecklistAcceptance,
  smokePlan,
  smokePlanAcceptance,
  smokeRunReport,
}).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorReleaseGate(connectorPackage).accepted, true);
const productionReleaseGate = buildMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
  target: 'production',
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance,
  bridgeSmokeReport,
  hostInstallChecklistAcceptance,
  smokePlan,
  smokePlanAcceptance,
  smokeRunReport,
});
assert.equal(productionReleaseGate.accepted, false);
assert.equal(productionReleaseGate.issues.some((issue) => issue.code.includes('production_evidence_accepted')), true);
assert.throws(
  () => assertMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
    target: 'production',
    adoptionIndex,
    fieldIntakeIndex,
    bridgeHandoff,
    bridgeHandoffAcceptance,
    bridgeSmokeReport,
    hostInstallChecklistAcceptance,
    smokePlan,
    smokePlanAcceptance,
    smokeRunReport,
  }),
  /release gate is not accepted/,
);

const platformRoadmap = buildMeetingAppTimelineConnectorPlatformRoadmap(hostInstallChecklist, {
  releaseGate,
});
assert.equal(platformRoadmap.schema, MEETING_APP_TIMELINE_CONNECTOR_PLATFORM_ROADMAP_SCHEMA);
assert.equal(platformRoadmap.accepted, true);
assert.equal(platformRoadmap.pilot_ready_count, 2);
assert.equal(platformRoadmap.production_ready_count, 0);
assert.equal(platformRoadmap.recommended_first_platform, 'google_meet');
assert.equal(platformRoadmap.rows.find((row) => row.platform === 'google_meet').recommended_first_surface, 'browser_extension');
assert.equal(platformRoadmap.rows.find((row) => row.platform === 'zoom').recommended_first_surface, 'native_detector');
assert.equal(platformRoadmap.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(platformRoadmap.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');
assert.equal(platformRoadmap.rows.every((row) => row.validation_sequence.includes('connector-release-gate.json')), true);
assert.equal(platformRoadmap.rows.find((row) => row.platform === 'google_meet').sdk_facade_methods.connector_release_gate, 'sdk.connectorReleaseGate(connectorPackage)');
assert.equal(assertMeetingAppTimelineConnectorPlatformRoadmap(hostInstallChecklist, { releaseGate }).accepted, true);

const adapterMatrix = buildMeetingAppTimelineConnectorAdapterMatrix(hostInstallChecklist, {
  releaseGate,
  platformRoadmap,
  fieldIntakeIndex,
  bridgeHandoff,
  smokePlan,
});
assert.equal(adapterMatrix.schema, MEETING_APP_TIMELINE_CONNECTOR_ADAPTER_MATRIX_SCHEMA);
assert.equal(adapterMatrix.accepted, true);
assert.equal(adapterMatrix.host_wiring_ready_count, 2);
assert.equal(adapterMatrix.pilot_ready_count, 2);
assert.equal(adapterMatrix.production_ready_count, 0);
assert.equal(adapterMatrix.recommended_first_platform, 'google_meet');
assert.equal(adapterMatrix.runtime_invariants.first_runtime_action, 'observe_platform_candidates');
assert.equal(adapterMatrix.runtime_invariants.mark_timestamp_field, 'captured_at_ms');
assert.equal(adapterMatrix.runtime_invariants.provider_replay_blocks_realtime, false);
assert.equal(adapterMatrix.files_to_read_first.includes('provider-replay-matrix.json'), true);
const googleAdapterRow = adapterMatrix.rows.find((row) => row.platform === 'google_meet');
const zoomAdapterRow = adapterMatrix.rows.find((row) => row.platform === 'zoom');
assert.equal(googleAdapterRow.selected_surface, 'browser_extension');
assert.equal(googleAdapterRow.adapter_mode, 'browser_content_script');
assert.equal(googleAdapterRow.install_step, 'install_manifest_v3_content_script_or_web_extension');
assert.equal(googleAdapterRow.runtime_sequence[0].action, 'observe_platform_candidates');
assert.equal(googleAdapterRow.runtime_sequence[1].action, 'insert_annotation');
assert.equal(googleAdapterRow.runtime_sequence[1].required_field, 'captured_at_ms');
assert.equal(googleAdapterRow.can_start_axis_before_provider, true);
assert.equal(googleAdapterRow.provider_reconcile_blocks_realtime, false);
assert.equal(googleAdapterRow.provider_replay.accepted, true);
assert.equal(googleAdapterRow.provider_replay.coverage.meeting_start, true);
assert.equal(googleAdapterRow.provider_replay.coverage.meeting_end, true);
assert.equal(googleAdapterRow.provider_replay.provider_events_block_realtime, false);
assert.equal(googleAdapterRow.input_sources.find((source) => source.id === 'provider_reconcile').blocks_realtime_annotation, false);
assert.equal(googleAdapterRow.bridge_contract.install_bridge_factory, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(googleAdapterRow.evidence_contract.pilot_required.includes('adapter_preflight_live_evidence'), true);
assert.equal(googleAdapterRow.evidence_contract.adapter_preflight.bridge_messages.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(googleAdapterRow.evidence_contract.pilot_required.includes('candidate_observation'), true);
assert.equal(googleAdapterRow.validation_files.includes('connector-smoke-run-report.json'), true);
assert.equal(zoomAdapterRow.selected_surface, 'native_detector');
assert.equal(zoomAdapterRow.adapter_mode, 'native_or_desktop_observer');
assert.equal(zoomAdapterRow.install_step, 'install_native_desktop_observer_or_accessibility_detector');
assert.equal(zoomAdapterRow.provider_replay.accepted, true);
assert.equal(zoomAdapterRow.evidence_contract.production_required.includes('runtime_host_replay'), true);
const adapterMatrixAcceptance = buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport(adapterMatrix);
assert.equal(adapterMatrixAcceptance.schema, 'meeting_app_timeline_connector_adapter_matrix_acceptance_report');
assert.equal(adapterMatrixAcceptance.accepted, true);
assert.equal(adapterMatrixAcceptance.issue_count, 0);
assert.equal(assertMeetingAppTimelineConnectorAdapterMatrix(hostInstallChecklist, {
  releaseGate,
  platformRoadmap,
  fieldIntakeIndex,
  bridgeHandoff,
  smokePlan,
}).accepted, true);

const hostAdapterConfigIndex = buildMeetingAppTimelineHostAdapterConfigIndex(adapterMatrix);
assert.equal(hostAdapterConfigIndex.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_INDEX_SCHEMA);
assert.equal(hostAdapterConfigIndex.accepted, true);
assert.equal(hostAdapterConfigIndex.row_count, 2);
assert.equal(hostAdapterConfigIndex.runtime_event_endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(hostAdapterConfigIndex.files_to_read_first.includes('host-adapter-config-index.json'), true);
assert.equal(hostAdapterConfigIndex.rows.find((row) => row.platform === 'google_meet').config_file, 'host-adapter-configs/google_meet.json');
assert.equal(hostAdapterConfigIndex.configs.google_meet.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_SCHEMA);
assert.equal(hostAdapterConfigIndex.configs.zoom.selected_surface, 'native_detector');
assert.equal(assertMeetingAppTimelineHostAdapterConfigIndex(adapterMatrix).accepted, true);

const googleHostAdapterConfig = buildMeetingAppTimelineHostAdapterConfig(adapterMatrix, 'google-meet');
assert.equal(googleHostAdapterConfig.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_SCHEMA);
assert.equal(googleHostAdapterConfig.accepted, true);
assert.equal(googleHostAdapterConfig.platform, 'google_meet');
assert.equal(googleHostAdapterConfig.selected_surface, 'browser_extension');
assert.equal(googleHostAdapterConfig.config_file, 'host-adapter-configs/google_meet.json');
assert.equal(googleHostAdapterConfig.realtime_contract.first_runtime_action, 'observe_platform_candidates');
assert.equal(googleHostAdapterConfig.realtime_contract.mark_timestamp_field, 'captured_at_ms');
assert.equal(googleHostAdapterConfig.realtime_contract.provider_replay_blocks_realtime, false);
assert.equal(googleHostAdapterConfig.realtime_contract.can_start_axis_before_provider, true);
assert.equal(googleHostAdapterConfig.runtime_sequence[0].action, 'observe_platform_candidates');
assert.equal(googleHostAdapterConfig.runtime_sequence[1].required_field, 'captured_at_ms');
assert.equal(googleHostAdapterConfig.evidence_contract.adapter_preflight.required_before, 'observe_platform_candidates_or_insert_annotation');
assert.equal(googleHostAdapterConfig.evidence_contract.adapter_preflight.required_live_inputs.includes('current_window_document_or_live_dom_snapshot'), true);
assert.equal(googleHostAdapterConfig.sdk_facade_methods.host_adapter_config, "sdk.connectorHostAdapterConfig('google_meet')");
assert.equal(googleHostAdapterConfig.provider_replay.accepted, true);
assert.equal(googleHostAdapterConfig.issue_count, 0);
assert.equal(assertMeetingAppTimelineHostAdapterConfig(connectorPackage, 'google-meet').accepted, true);

const sdkGoogleHostAdapterConfig = sdk.connectorHostAdapterConfig('google-meet', {
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});
assert.equal(sdkGoogleHostAdapterConfig.platform, 'google_meet');
assert.equal(sdkGoogleHostAdapterConfig.accepted, true);
assert.equal(sdkGoogleHostAdapterConfig.runtime_sequence[1].action, 'insert_annotation');
assert.equal(sdk.assertConnectorHostAdapterConfig(connectorPackage, 'zoom').platform, 'zoom');
const sdkHostAdapterConfigIndex = sdk.connectorHostAdapterConfigIndex(connectorPackage);
assert.equal(sdkHostAdapterConfigIndex.accepted, true);
assert.equal(sdkHostAdapterConfigIndex.configs.zoom.provider_replay.provider_events_block_realtime, false);
assert.equal(sdk.assertConnectorHostAdapterConfigIndex(adapterMatrix).row_count, 2);

const resolvedGoogleHostAdapter = resolveMeetingAppTimelineHostAdapterConfig(hostAdapterConfigIndex, {
  tabs: [
    { url: 'https://zoom.us/j/987654321', title: 'Zoom Meeting', active: false, in_meeting: true },
    { url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true, audible: true, in_meeting: true },
  ],
});
assert.equal(resolvedGoogleHostAdapter.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_CONFIG_RESOLUTION_SCHEMA);
assert.equal(resolvedGoogleHostAdapter.accepted, true);
assert.equal(resolvedGoogleHostAdapter.platform, 'google_meet');
assert.equal(resolvedGoogleHostAdapter.resolution.reason, 'meeting_url');
assert.equal(resolvedGoogleHostAdapter.host_config.selected_surface, 'browser_extension');
assert.equal(resolvedGoogleHostAdapter.config_file, 'host-adapter-configs/google_meet.json');
assert.equal(assertMeetingAppTimelineResolvedHostAdapterConfig(adapterMatrix, {
  tab: { url: 'https://meet.google.com/abc-defg-hij', active: true },
}).accepted, true);

const resolvedZoomHostAdapter = resolveMeetingAppTimelineHostAdapterConfig(connectorPackage, 'https://zoom.us/j/987654321');
assert.equal(resolvedZoomHostAdapter.accepted, true);
assert.equal(resolvedZoomHostAdapter.platform, 'zoom');
assert.equal(resolvedZoomHostAdapter.host_config.adapter_mode, 'native_or_desktop_observer');
assert.equal(resolvedZoomHostAdapter.runtime_event_endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');

const unresolvedTeamsHostAdapter = resolveMeetingAppTimelineHostAdapterConfig(hostAdapterConfigIndex, {
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
});
assert.equal(unresolvedTeamsHostAdapter.accepted, false);
assert.equal(unresolvedTeamsHostAdapter.platform, 'microsoft_teams');
assert.equal(unresolvedTeamsHostAdapter.issues.some((issue) => issue.code === 'meeting_platform_not_supported'), true);

const googleBootstrapPlan = buildMeetingAppTimelineHostAdapterBootstrapPlan(hostAdapterConfigIndex, {
  tab: { url: 'https://meet.google.com/abc-defg-hij', active: true, audible: true, in_meeting: true },
});
assert.equal(googleBootstrapPlan.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_SCHEMA);
assert.equal(googleBootstrapPlan.accepted, true);
assert.equal(googleBootstrapPlan.platform, 'google_meet');
assert.deepEqual(googleBootstrapPlan.startup_order.slice(0, 5), [
  'resolve_meeting_platform',
  'load_host_adapter_config',
  'install_host_adapter',
  'runtime_observe_platform_candidates',
  'runtime_insert_annotation',
]);
assert.equal(googleBootstrapPlan.steps.find((step) => step.id === 'install_host_adapter').selected_surface, 'browser_extension');
assert.equal(googleBootstrapPlan.steps.find((step) => step.id === 'install_host_adapter').install_target, 'manifest_v3_content_script');
assert.equal(googleBootstrapPlan.steps.find((step) => step.action === 'insert_annotation').captured_at_ms_required, true);
assert.equal(googleBootstrapPlan.required_runtime_actions.includes('observe_platform_candidates'), true);
assert.equal(googleBootstrapPlan.required_runtime_actions.includes('insert_annotation'), true);
assert.equal(googleBootstrapPlan.local_axis_contract.provider_reconcile_blocks_realtime, false);
assert.equal(assertMeetingAppTimelineHostAdapterBootstrapPlan(connectorPackage, 'https://meet.google.com/abc-defg-hij').accepted, true);

const zoomBootstrapPlan = buildMeetingAppTimelineHostAdapterBootstrapPlan(connectorPackage, 'https://zoom.us/j/987654321');
assert.equal(zoomBootstrapPlan.accepted, true);
assert.equal(zoomBootstrapPlan.platform, 'zoom');
assert.equal(zoomBootstrapPlan.install_target, 'native_or_desktop_observer');
assert.equal(zoomBootstrapPlan.steps.find((step) => step.id === 'install_host_adapter').adapter_mode, 'native_or_desktop_observer');

const bootstrapPlanMatrix = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix(connectorPackage);
assert.equal(bootstrapPlanMatrix.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_MATRIX_SCHEMA);
assert.equal(bootstrapPlanMatrix.accepted, true);
assert.equal(bootstrapPlanMatrix.platform_count, 2);
assert.equal(bootstrapPlanMatrix.accepted_count, 2);
assert.equal(bootstrapPlanMatrix.rows.find((row) => row.platform === 'google_meet').install_target, 'manifest_v3_content_script');
assert.equal(bootstrapPlanMatrix.rows.find((row) => row.platform === 'zoom').install_target, 'native_or_desktop_observer');
assert.equal(bootstrapPlanMatrix.plans.google_meet.startup_order[3], 'runtime_observe_platform_candidates');
assert.equal(assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix(connectorPackage).accepted, true);
const bootstrapPlanMatrixAcceptance = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport(bootstrapPlanMatrix);
assert.equal(bootstrapPlanMatrixAcceptance.schema, MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_MATRIX_ACCEPTANCE_SCHEMA);
assert.equal(bootstrapPlanMatrixAcceptance.accepted, true);
assert.equal(bootstrapPlanMatrixAcceptance.issue_count, 0);
assert.equal(bootstrapPlanMatrixAcceptance.rows.every((row) => row.startup_order_ready === true), true);
const releaseGateWithBootstrapAcceptance = buildMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance,
  bridgeSmokeReport,
  hostInstallChecklistAcceptance,
  smokePlan,
  smokePlanAcceptance,
  smokeRunReport,
  hostAdapterBootstrapPlanMatrixAcceptanceReport: bootstrapPlanMatrixAcceptance,
});
assert.equal(releaseGateWithBootstrapAcceptance.accepted, true);
assert.equal(
  releaseGateWithBootstrapAcceptance.required_gates
    .find((gate) => gate.id === 'host_adapter_bootstrap_plan_matrix_acceptance')
    .accepted,
  true,
);
assert.equal(
  releaseGateWithBootstrapAcceptance.files_to_read_first.includes('host-adapter-bootstrap-plan-matrix-acceptance.json'),
  true,
);
assert.equal(
  releaseGateWithBootstrapAcceptance.source_schemas.host_adapter_bootstrap_plan_matrix_acceptance,
  MEETING_APP_TIMELINE_HOST_ADAPTER_BOOTSTRAP_PLAN_MATRIX_ACCEPTANCE_SCHEMA,
);

const allPlatformSdk = createMeetingAppTimelineSdk({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
const allPlatformPackage = allPlatformSdk.connectorPackage({
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});
const allPlatformBootstrapMatrix = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix(allPlatformPackage);
assert.equal(allPlatformBootstrapMatrix.accepted, true);
assert.equal(allPlatformBootstrapMatrix.platform_count, 5);
assert.deepEqual(
  new Set(allPlatformBootstrapMatrix.rows.map((row) => row.platform)),
  new Set(['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']),
);
assert.equal(allPlatformBootstrapMatrix.rows.every((row) => row.startup_order.includes('runtime_insert_annotation')), true);
assert.equal(allPlatformBootstrapMatrix.plans.microsoft_teams.resolution.platform, 'microsoft_teams');
assert.equal(allPlatformBootstrapMatrix.plans.webex.accepted, true);
assert.equal(allPlatformBootstrapMatrix.plans.lark.accepted, true);

const teamsBootstrapPlan = buildMeetingAppTimelineHostAdapterBootstrapPlan(hostAdapterConfigIndex, {
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
});
assert.equal(teamsBootstrapPlan.accepted, false);
assert.equal(teamsBootstrapPlan.issues.includes('host_adapter_config_resolution_not_accepted'), true);

assert.equal(connectorPackage.adapter_blueprints.ready_count, 2);
assert.equal(connectorPackage.adapter_blueprints.matrix.schema, 'meeting_platform_adapter_blueprint_matrix');
assert.equal(connectorPackage.adapter_blueprints.matrix.blueprints, undefined);
assert.equal(connectorPackage.startup_plans.realtime_startup_ready_count, 2);
assert.equal(connectorPackage.startup_plans.matrix.schema, 'meeting_platform_adapter_startup_plan_matrix');
assert.equal(connectorPackage.startup_plans.matrix.plans, undefined);
assert.equal(connectorPackage.entrypoints.some((entry) => entry.id === 'adapter-blueprint'), true);
assert.equal(connectorPackage.entrypoints.some((entry) => entry.id === 'startup-plan'), true);
assert.equal(connectorPackage.contracts.adapter_blueprint_required_before_host_wiring, true);
assert.equal(connectorPackage.contracts.startup_plan_required_before_runtime_install, true);

let capturedRuntimeRequest = null;
const connectorRuntimeClient = createMeetingAppTimelineConnectorRuntimeClient(connectorPackage, {
  now: () => 1_782_614_401_000,
  fetch: async (url, init) => {
    capturedRuntimeRequest = {
      url,
      method: init.method,
      headers: init.headers,
      body: JSON.parse(init.body),
    };
    return new Response(JSON.stringify({
      ok: true,
      accepted: capturedRuntimeRequest.body,
    }), {
      status: 201,
      headers: { 'content-type': 'application/json' },
    });
  },
});
assert.equal(connectorRuntimeClient.schema, 'meeting_app_timeline_connector_runtime_client');
assert.equal(connectorRuntimeClient.endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(connectorRuntimeClient.supports('insert_annotation', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('provider_event', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('adapter_blueprint', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('missing_action', 'google-meet'), false);
assert.equal(connectorRuntimeClient.supported_actions_by_platform.google_meet.includes('speaker_track'), true);
const builtInsertEvent = connectorRuntimeClient.buildEvent({
  action: 'insert_annotation',
  platform: 'zoom',
  annotation: {
    id: 'note-1',
    captured_at_ms: 1_782_614_402_000,
  },
});
assert.equal(builtInsertEvent.schema, 'meeting_platform_runtime_event');
assert.equal(builtInsertEvent.action, 'insert_annotation');
assert.equal(builtInsertEvent.platform, 'zoom');

const insertResult = await connectorRuntimeClient.insertAnnotation('zoom', {
  id: 'note-2',
  label: 'why?',
  captured_at_ms: 1_782_614_403_000,
});
assert.equal(capturedRuntimeRequest.url, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(capturedRuntimeRequest.method, 'POST');
assert.equal(capturedRuntimeRequest.body.action, 'insert_annotation');
assert.equal(capturedRuntimeRequest.body.platform, 'zoom');
assert.equal(insertResult.accepted.annotation.id, 'note-2');

await connectorRuntimeClient.observePlatformCandidates({
  tabs: [{
    active: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
  }],
});
assert.equal(capturedRuntimeRequest.body.action, 'observe_platform_candidates');
assert.equal(capturedRuntimeRequest.body.platform, undefined);
assert.equal(capturedRuntimeRequest.body.tabs[0].url, 'https://meet.google.com/abc-defg-hij');
await connectorRuntimeClient.runManifest({ platforms: ['google-meet'], target: 'realtime' });
assert.equal(capturedRuntimeRequest.body.action, 'run_manifest');

await connectorRuntimeClient.adapterBlueprint('google-meet');
assert.equal(capturedRuntimeRequest.body.action, 'adapter_blueprint');
assert.equal(capturedRuntimeRequest.body.platform, 'google_meet');

await connectorRuntimeClient.adapterBlueprints({ platforms: ['google-meet', 'zoom'] });
assert.equal(capturedRuntimeRequest.body.action, 'adapter_blueprints');
assert.deepEqual(capturedRuntimeRequest.body.platforms, ['google-meet', 'zoom']);

const stripped = stripMeetingAppTimelineConnectorPackageFileContents(connectorPackage);
assert.equal(stripped.extension.scaffold.files.some((file) => 'content' in file), false);
assert.equal(connectorPackage.extension.scaffold.files.some((file) => 'content' in file), true);

const missingRuntimeActionPackage = {
  ...connectorPackage,
  runtime_events: {
    ...connectorPackage.runtime_events,
    plan_matrix: {
      ...connectorPackage.runtime_events.plan_matrix,
      rows: connectorPackage.runtime_events.plan_matrix.rows.filter((row) => row.action !== 'insert_annotation'),
    },
  },
};
const missingRuntimeActionAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(
  missingRuntimeActionPackage,
);
assert.equal(missingRuntimeActionAcceptance.accepted, false);
assert.equal(
  missingRuntimeActionAcceptance.missing_runtime_actions.some((item) => item.action === 'insert_annotation'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineConnectorPackage(missingRuntimeActionPackage),
  /Meeting app timeline connector package is not accepted/,
);
assert.throws(
  () => createMeetingAppTimelineConnectorRuntimeClient(missingRuntimeActionPackage, {
    fetch: async () => new Response('{}'),
  }),
  /not accepted for runtime client/,
);

const missingBlueprintPackage = {
  ...connectorPackage,
  adapter_blueprints: undefined,
};
const missingBlueprintAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(missingBlueprintPackage);
assert.equal(missingBlueprintAcceptance.accepted, false);
assert.equal(missingBlueprintAcceptance.issues.some((issue) => issue.code === 'missing_adapter_blueprint_matrix'), true);
const missingStartupPackage = {
  ...connectorPackage,
  startup_plans: undefined,
};
const missingStartupAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(missingStartupPackage);
assert.equal(missingStartupAcceptance.accepted, false);
assert.equal(missingStartupAcceptance.issues.some((issue) => issue.code === 'missing_startup_plan_matrix'), true);
const unsafeRuntimeClient = createMeetingAppTimelineConnectorRuntimeClient(missingRuntimeActionPackage, {
  assertPackage: false,
  fetch: async () => new Response('{}'),
});
assert.throws(
  () => unsafeRuntimeClient.buildEvent({
    action: 'insert_annotation',
    platform: 'zoom',
    annotation: { id: 'blocked', captured_at_ms: 1_782_614_404_000 },
  }),
  /Runtime action insert_annotation is not present in connector package/,
);

const blockingProviderPackage = {
  ...connectorPackage,
  contracts: {
    ...connectorPackage.contracts,
    provider_events_block_realtime: true,
  },
};
const blockingProviderAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(blockingProviderPackage);
assert.equal(blockingProviderAcceptance.accepted, false);
assert.equal(blockingProviderAcceptance.issues.some((issue) => issue.code === 'provider_events_block_realtime'), true);

const brokenHostInstallChecklist = {
  ...hostInstallChecklist,
  rows: hostInstallChecklist.rows.map((row) => row.platform === 'zoom'
    ? { ...row, selected_surface: undefined }
    : row),
};
const brokenHostInstallChecklistAcceptance = buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(brokenHostInstallChecklist);
assert.equal(brokenHostInstallChecklistAcceptance.accepted, false);
assert.equal(
  brokenHostInstallChecklistAcceptance.issues.some((issue) => issue.code === 'row_missing_selected_surface'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineConnectorHostInstallChecklist(brokenHostInstallChecklist),
  /host install checklist is not accepted/,
);

const brokenBridgeHandoff = structuredClone(bridgeHandoff);
brokenBridgeHandoff.message_contract.message_types = brokenBridgeHandoff.message_contract.message_types.filter((item) => item !== 'meeting_timeline.insert_mark');
const brokenBridgeHandoffAcceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(brokenBridgeHandoff);
assert.equal(brokenBridgeHandoffAcceptance.accepted, false);
assert.equal(
  brokenBridgeHandoffAcceptance.issues.some((issue) => issue.code === 'missing_message_type' && issue.message_type === 'meeting_timeline.insert_mark'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineConnectorBridgeHandoff(brokenBridgeHandoff),
  /bridge handoff is not accepted/,
);
const brokenBridgeSmoke = await runMeetingAppTimelineConnectorBridgeSmoke(brokenBridgeHandoff);
assert.equal(brokenBridgeSmoke.accepted, false);
assert.equal(brokenBridgeSmoke.issues.some((issue) => issue === 'acceptance:missing_message_type'), true);
await assert.rejects(
  () => assertMeetingAppTimelineConnectorBridgeSmoke(brokenBridgeHandoff),
  /bridge smoke failed/,
);

const brokenSmokePlan = structuredClone(smokePlan);
delete brokenSmokePlan.rows
  .find((row) => row.platform === 'zoom')
  .steps.find((step) => step.action === 'insert_annotation')
  .input.annotation.captured_at_ms;
const brokenSmokePlanAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(brokenSmokePlan);
assert.equal(brokenSmokePlanAcceptance.accepted, false);
assert.equal(
  brokenSmokePlanAcceptance.issues.some((issue) => issue.code === 'row_insert_annotation_missing_captured_at_ms'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineConnectorSmokePlan(brokenSmokePlan),
  /connector smoke plan is not accepted/,
);
await assert.rejects(
  () => assertMeetingAppTimelineConnectorSmokeRun(brokenSmokePlan),
  /connector smoke run failed/,
);

const brokenBootstrapPlanMatrix = structuredClone(bootstrapPlanMatrix);
brokenBootstrapPlanMatrix.rows.find((row) => row.platform === 'zoom').startup_order = [
  'resolve_meeting_platform',
  'load_host_adapter_config',
  'install_host_adapter',
  'runtime_insert_annotation',
  'runtime_observe_platform_candidates',
];
const brokenBootstrapPlanMatrixAcceptance = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport(brokenBootstrapPlanMatrix);
assert.equal(brokenBootstrapPlanMatrixAcceptance.accepted, false);
assert.equal(
  brokenBootstrapPlanMatrixAcceptance.issues.some((issue) => issue.code === 'row_invalid_startup_order'),
  true,
);
const brokenBootstrapReleaseGate = buildMeetingAppTimelineConnectorReleaseGate(hostInstallChecklist, {
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance,
  bridgeSmokeReport,
  hostInstallChecklistAcceptance,
  smokePlan,
  smokePlanAcceptance,
  smokeRunReport,
  hostAdapterBootstrapPlanMatrixAcceptanceReport: brokenBootstrapPlanMatrixAcceptance,
});
assert.equal(brokenBootstrapReleaseGate.accepted, false);
assert.equal(
  brokenBootstrapReleaseGate.required_gates
    .find((gate) => gate.id === 'host_adapter_bootstrap_plan_matrix_acceptance')
    .accepted,
  false,
);
assert.equal(
  brokenBootstrapReleaseGate.issues.some((issue) => issue.gate === 'host_adapter_bootstrap_plan_matrix_acceptance'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix(brokenBootstrapPlanMatrix),
  /bootstrap plan matrix is not accepted/,
);

const productionAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage, {
  target: 'production',
});
assert.equal(productionAcceptance.accepted, false);
assert.equal(
  productionAcceptance.issues.some((issue) => issue.code === 'production_requires_live_snapshot_evidence'),
  true,
);

const productionEvidenceAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage, {
  target: 'production',
  liveEvidenceAccepted: true,
});
assert.equal(productionEvidenceAcceptance.accepted, true);

console.log('ok meeting app connector package adapter');
