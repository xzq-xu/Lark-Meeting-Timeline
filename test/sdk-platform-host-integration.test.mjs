import assert from 'node:assert/strict';
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { pathToFileURL } from 'node:url';

import {
  MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA,
  MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA,
  MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA,
  assertMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationPlan,
  buildMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-host-integration.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const platforms = ['google-meet', 'teams', 'zoom'];

function file(scaffold, path) {
  return scaffold.files.find((item) => item.path === path);
}

const plan = buildMeetingPlatformHostIntegrationPlan({
  baseUrl,
  platforms,
  packageName: 'timeline-host-consumer',
});

assert.equal(plan.schema, MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA);
assert.deepEqual(plan.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
assert.equal(plan.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(plan.runtime_contract.provider_events_block_realtime, false);
assert.equal(plan.runtime_contract.candidate_observation_required_for_host_axis_binding, true);
assert.equal(plan.runtime_contract.speaker_track_required_for_realtime_timeline, true);
assert.equal(plan.runtime_contract.participant_track_required_for_realtime_timeline, true);
assert.equal(plan.endpoints.platform_events, '/api/platform-events');
assert.equal(plan.endpoints.adapter_contracts, '/api/meeting-platform/contracts');
assert.equal(plan.endpoints.runtime_bundles, '/api/meeting-platform/runtime-bundles');
assert.equal(plan.endpoints.platform_conformance, '/api/meeting-platform/conformance');
assert.equal(plan.endpoints.consumer_handoff, '/api/meeting-platform/consumer-handoff');
assert.equal(plan.endpoints.observer_plans, '/api/meeting-platform/observer-plans');
assert.equal(plan.endpoints.runtime_event_plans, '/api/meeting-platform/runtime-event-plans');
assert.equal(plan.endpoints.adapter_routes, '/api/meeting-platform/adapter-routes');
assert.equal(plan.endpoints.adapter_blueprints, '/api/meeting-platform/adapter-blueprints');
assert.equal(plan.endpoints.strategy, '/api/meeting-platform/strategy');
assert.equal(plan.endpoints.platform_resolution, '/api/meeting-platform/resolve');
assert.equal(plan.endpoints.platform_candidate_resolution, '/api/meeting-platform/resolve-candidates');
assert.equal(plan.endpoints.platform_candidate_observation, '/api/meeting-platform/observe-candidates');
assert.equal(plan.endpoints.extension_plan, '/api/meeting-platform/extension-plan');
assert.equal(plan.endpoints.integration_runtime, '/api/meeting-platform/integration-runtime');
assert.equal(plan.endpoints.integration_runtime_manifest, '/api/meeting-platform/integration-runtime/manifest');
assert.equal(plan.endpoints.integration_runtime_run_manifest, '/api/meeting-platform/integration-runtime/run-manifest');
assert.equal(plan.endpoints.handoff_readiness, '/api/meeting-platform/handoff-readiness');
assert.equal(plan.endpoints.runtime_events, '/api/meeting-platform/runtime-events');
assert.equal(plan.commands.validate_integration_runtime_manifest, 'npm run meeting-platform:integration-runtime-run-manifest');
assert.equal(plan.commands.validate_handoff_readiness, 'npm run meeting-platform:handoff-readiness');
assert.equal(plan.commands.validate_platform_conformance, 'npm run meeting-platform:conformance');
assert.equal(plan.commands.export_consumer_handoff, 'npm run meeting-platform:consumer-handoff');
assert.equal(plan.sdk.runtime_event_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event');
assert.equal(plan.sdk.platform_conformance_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-conformance');
assert.equal(plan.sdk.consumer_handoff_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff');
assert.equal(plan.handoff_bundle.platform_count, 3);
assert.equal(plan.platform_conformance_report.accepted, true);
assert.equal(plan.platform_conformance_report.accepted_count, 3);
assert.equal(plan.platform_conformance_report.candidate_observer_count, 3);
assert.equal(plan.runtime_bundle_matrix.platform_count, 3);
assert.equal(plan.runtime_bundle_matrix.provider_required_for_realtime_count, 0);
assert.equal(plan.runtime_bundle_matrix.candidate_observer_count, 3);
assert.equal(plan.runtime_bundle_matrix.observer_plan_ready_count, 3);
assert.equal(plan.observer_plan_matrix.platform_count, 3);
assert.equal(plan.observer_plan_matrix.sdk_ready_count, 3);
assert.equal(plan.observer_plan_matrix.preflight_accepted_count, 0);
assert.equal(plan.observer_plan_matrix.rows.every((row) => row.preflight_status === 'not_run'), true);
assert.equal(plan.candidate_observation_contract.platform_count, 3);
assert.equal(plan.candidate_observation_contract.ready_count, 3);
assert.equal(plan.candidate_observation_contract.all_ready, true);
assert.equal(plan.candidate_observation_contract.endpoint, '/api/meeting-platform/observe-candidates');
assert.equal(plan.candidate_observation_contract.runtime_event_action, 'observe_platform_candidates');
assert.deepEqual(plan.candidate_observation_contract.required_permissions, ['tabs']);
assert.equal(plan.candidate_observation_contract.rows.every((row) => row.ready === true), true);
assert.equal(plan.candidate_observation_contract.rows.some((row) => row.platform === 'google_meet'), true);
assert.equal(plan.speaker_track_matrix.platform_count, 3);
assert.equal(plan.speaker_track_matrix.provider_blocking_count, 0);
assert.equal(plan.speaker_track_matrix.transcript_blocking_count, 0);
assert.equal(plan.participant_track_matrix.platform_count, 3);
assert.equal(plan.participant_track_matrix.provider_blocking_count, 0);
assert.equal(plan.participant_track_matrix.transcript_blocking_count, 0);
assert.equal(plan.meeting_track_contract.platform_count, 3);
assert.equal(plan.meeting_track_contract.speaker_ready_count, 3);
assert.equal(plan.meeting_track_contract.participant_ready_count, 3);
assert.equal(plan.meeting_track_contract.provider_blocking_count, 0);
assert.equal(plan.meeting_track_contract.transcript_blocking_count, 0);
assert.equal(plan.meeting_track_contract.all_ready, true);
assert.equal(plan.meeting_track_contract.rows.every((row) => row.ready === true), true);
assert.equal(plan.runtime_event_plan_matrix.platform_count, 3);
assert.equal(plan.runtime_event_plan_matrix.realtime_provider_dependency_count, 0);
assert.equal(plan.runtime_event_plan_matrix.transcript_realtime_dependency_count, 0);
assert.equal(plan.adapter_route_matrix.platform_count, 3);
assert.equal(plan.adapter_route_matrix.rows.every((row) => row.first_route === 'local_observer_axis'), true);
assert.equal(plan.adapter_blueprint_matrix.platform_count, 3);
assert.equal(plan.adapter_blueprint_matrix.ready_count, 3);
assert.equal(plan.adapter_blueprint_matrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(plan.adapter_blueprint_matrix.rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
assert.equal(plan.adapter_runtime_contract.platform_count, 3);
assert.equal(plan.adapter_runtime_contract.ready_count, 3);
assert.equal(plan.adapter_runtime_contract.all_ready, true);
assert.equal(plan.adapter_runtime_contract.index_file, 'src/platform-adapters/index.mjs');
assert.equal(plan.adapter_runtime_contract.rows.find((row) => row.platform === 'google_meet').adapter_file, 'src/platform-adapters/google_meet.mjs');
assert.equal(plan.adapter_runtime_contract.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(plan.adapter_runtime_contract.rows.every((row) => row.timestamp_field === 'captured_at_ms'), true);
assert.equal(plan.adaptation_strategy_matrix.strategy_count, 3);
assert.equal(plan.adaptation_strategy_matrix.rows.every((row) => row.provider_blocks_realtime === false), true);
assert.equal(plan.adapter_contract_matrix.platform_count, 3);
assert.equal(plan.adapter_contract_acceptance_matrix.accepted_count, 3);
assert.equal(plan.extension_install_plan.platforms.includes('google_meet'), true);
assert.equal(plan.integration_plans.google_meet.provider_events.endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(plan.next_actions.includes('run_meeting_platform_contract_acceptance_before_enabling_new_platform'), true);
assert.equal(plan.next_actions.includes('run_meeting_platform_conformance_before_host_handoff'), true);
assert.equal(plan.next_actions.includes('export_meeting_platform_consumer_handoff_for_downstream_project'), true);
assert.equal(plan.next_actions.includes('wire_host_routes_to_handleMeetingPlatformRequest'), true);
assert.equal(plan.next_actions.includes('wire_observer_plans_to_host_scheduler'), true);
assert.equal(plan.next_actions.includes('publish_adapter_blueprints_endpoint_for_downstream_hosts'), true);
assert.equal(plan.next_actions.includes('wire_platform_adapter_runtime_entries_into_host_surface'), true);
assert.equal(plan.next_actions.includes('run_meeting_platform_integration_runtime_manifest_before_host_handoff'), true);
assert.equal(plan.next_actions.includes('run_meeting_platform_handoff_readiness_before_host_handoff'), true);

const scaffold = buildMeetingPlatformHostIntegrationScaffold({
  baseUrl,
  platforms,
  packageName: 'timeline-host-consumer',
});

assert.equal(scaffold.schema, MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA);
assert.equal(scaffold.files.length, 30);
assert.equal(file(scaffold, 'package.json').mime, 'application/json');
const manifest = JSON.parse(file(scaffold, 'package.json').content);
assert.equal(manifest.name, 'timeline-host-consumer');
assert.equal(manifest.scripts['meeting-platform:strategy'], 'node ./scripts/print-strategy.mjs');
assert.equal(manifest.scripts['meeting-platform:contracts'], 'node ./scripts/print-contracts.mjs');
assert.equal(manifest.scripts['meeting-platform:contract-acceptance'], 'node ./scripts/verify-contracts.mjs');
assert.equal(manifest.scripts['meeting-platform:conformance'], 'node ./scripts/verify-conformance.mjs');
assert.equal(manifest.scripts['meeting-platform:consumer-handoff'], 'node ./scripts/print-consumer-handoff.mjs');
assert.equal(manifest.scripts['meeting-platform:runtime-bundles'], 'node ./scripts/print-runtime-bundles.mjs');
assert.equal(manifest.scripts['meeting-platform:observer-plans'], 'node ./scripts/print-observer-plans.mjs');
assert.equal(manifest.scripts['meeting-platform:runtime-event-plans'], 'node ./scripts/print-runtime-event-plans.mjs');
assert.equal(manifest.scripts['meeting-platform:adapter-routes'], 'node ./scripts/print-adapter-routes.mjs');
assert.equal(manifest.scripts['meeting-platform:adapter-blueprints'], 'node ./scripts/print-adapter-blueprints.mjs');
assert.equal(manifest.scripts['meeting-platform:adapters'], 'node ./scripts/print-platform-adapters.mjs');
assert.equal(manifest.scripts['meeting-platform:resolve'], 'node ./scripts/resolve-platform.mjs');
assert.equal(manifest.scripts['meeting-platform:resolve-candidates'], 'node ./scripts/resolve-platform-candidates.mjs');
assert.equal(manifest.scripts['meeting-platform:observe-candidates'], 'node ./scripts/observe-platform-candidates.mjs');
assert.equal(manifest.scripts['meeting-platform:extension-plan'], 'node ./scripts/print-extension-plan.mjs');
assert.equal(manifest.scripts['meeting-platform:integration-runtime'], 'node ./scripts/print-integration-runtime.mjs');
assert.equal(manifest.scripts['meeting-platform:integration-runtime-manifest'], 'node ./scripts/print-integration-runtime-manifest.mjs');
assert.equal(manifest.scripts['meeting-platform:integration-runtime-run-manifest'], 'node ./scripts/run-integration-runtime-manifest.mjs');
assert.equal(manifest.scripts['meeting-platform:handoff-readiness'], 'node ./scripts/run-handoff-readiness.mjs');
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('createMeetingPlatformIntegrationRuntime'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('integrationRuntimeManifest'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('runIntegrationRuntimeManifest'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('assertRunIntegrationRuntimeManifest'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('integrationRuntimeSummary'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('runHandoffReadiness'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('handleRuntimeEvent'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('liveAdapters'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterContractMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterContractAcceptanceMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformConformance'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('consumerHandoff'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformRuntimeBundleMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('meetingAppRuntimeObserverPlanMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformRuntimeEventPlanMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterRouteMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('adapterRoutes'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterBlueprintMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('adapterBlueprints'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdaptationStrategyMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('resolvePlatform'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('resolvePlatformCandidates'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('observePlatformCandidates'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('meetingAppExtensionInstallPlan'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('capturedAtMs'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('speakerTrack'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('participantTrack'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('createMeetingPlatformAdapters'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapter(platform'), true);
assert.equal(file(scaffold, 'src/platform-adapters/index.mjs').content.includes('createMeetingPlatformAdapter'), true);
assert.equal(file(scaffold, 'src/platform-adapters/index.mjs').content.includes('createGoogleMeetTimelineAdapter'), true);
assert.equal(file(scaffold, 'src/platform-adapters/google_meet.mjs').content.includes('createGoogleMeetTimelineAdapter'), true);
assert.equal(file(scaffold, 'src/platform-adapters/google_meet.mjs').content.includes('observeCandidates'), true);
assert.equal(file(scaffold, 'src/platform-adapters/google_meet.mjs').content.includes('insertAnnotation'), true);
assert.equal(file(scaffold, 'src/platform-adapters/google_meet.mjs').content.includes('captured_at_ms'), true);
assert.equal(file(scaffold, 'src/platform-adapters/microsoft_teams.mjs').content.includes('createMicrosoftTeamsTimelineAdapter'), true);
assert.equal(file(scaffold, 'src/platform-adapters/zoom.mjs').content.includes('createZoomTimelineAdapter'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('handleFetchRequest'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/contracts'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/conformance'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/consumer-handoff'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/runtime-bundles'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/observer-plans'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/runtime-event-plans'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/adapter-routes'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/adapter-blueprints'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/strategy'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/resolve'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/resolve-candidates'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/observe-candidates'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/extension-plan'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/integration-runtime'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/integration-runtime/manifest'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/integration-runtime/run-manifest'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/handoff-readiness'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/runtime-events'), true);
assert.equal(file(scaffold, 'scripts/print-handoff.mjs').content.includes('host.handoffBundle()'), true);
assert.equal(file(scaffold, 'scripts/print-contracts.mjs').content.includes('host.adapterContracts()'), true);
assert.equal(file(scaffold, 'scripts/verify-contracts.mjs').content.includes('adapterContractAcceptance'), true);
assert.equal(file(scaffold, 'scripts/verify-conformance.mjs').content.includes('host.platformConformance()'), true);
assert.equal(file(scaffold, 'scripts/print-consumer-handoff.mjs').content.includes('host.consumerHandoff'), true);
assert.equal(file(scaffold, 'scripts/print-runtime-bundles.mjs').content.includes('host.runtimeBundles()'), true);
assert.equal(file(scaffold, 'scripts/print-observer-plans.mjs').content.includes('host.observerPlans()'), true);
assert.equal(file(scaffold, 'scripts/print-runtime-event-plans.mjs').content.includes('host.runtimeEventPlans()'), true);
assert.equal(file(scaffold, 'scripts/print-adapter-routes.mjs').content.includes('host.adapterRoutes()'), true);
assert.equal(file(scaffold, 'scripts/print-adapter-blueprints.mjs').content.includes('host.adapterBlueprints()'), true);
assert.equal(file(scaffold, 'scripts/print-platform-adapters.mjs').content.includes('host.platformAdapters()'), true);
assert.equal(file(scaffold, 'scripts/print-strategy.mjs').content.includes('host.adaptationStrategyMatrix()'), true);
assert.equal(file(scaffold, 'scripts/resolve-platform.mjs').content.includes('host.resolvePlatform'), true);
assert.equal(file(scaffold, 'scripts/resolve-platform.mjs').content.includes('MEETING_PLATFORM_URL'), true);
assert.equal(file(scaffold, 'scripts/resolve-platform-candidates.mjs').content.includes('host.resolvePlatformCandidates'), true);
assert.equal(file(scaffold, 'scripts/resolve-platform-candidates.mjs').content.includes('MEETING_PLATFORM_CANDIDATES_JSON'), true);
assert.equal(file(scaffold, 'scripts/observe-platform-candidates.mjs').content.includes('host.observePlatformCandidates'), true);
assert.equal(file(scaffold, 'scripts/observe-platform-candidates.mjs').content.includes('MEETING_PLATFORM_CANDIDATES_JSON'), true);
assert.equal(file(scaffold, 'scripts/print-extension-plan.mjs').content.includes('host.extensionInstallPlan()'), true);
assert.equal(file(scaffold, 'scripts/print-integration-runtime.mjs').content.includes('host.integrationRuntimeSummary()'), true);
assert.equal(file(scaffold, 'scripts/print-integration-runtime-manifest.mjs').content.includes('host.integrationRuntimeManifest()'), true);
assert.equal(file(scaffold, 'scripts/run-integration-runtime-manifest.mjs').content.includes('host.runIntegrationRuntimeManifest'), true);
assert.equal(file(scaffold, 'scripts/run-integration-runtime-manifest.mjs').content.includes('MEETING_PLATFORM_REQUIRE_HANDOFF_READY'), true);
assert.equal(file(scaffold, 'scripts/run-handoff-readiness.mjs').content.includes('host.runHandoffReadiness'), true);
assert.equal(file(scaffold, 'scripts/run-handoff-readiness.mjs').content.includes('handoff_ready_count'), true);
assert.equal(file(scaffold, 'README.md').content.includes('captured_at_ms'), true);
assert.equal(file(scaffold, 'README.md').content.includes('adaptationStrategyMatrix'), true);
assert.equal(file(scaffold, 'README.md').content.includes('runtimeEventPlans'), true);
assert.equal(file(scaffold, 'README.md').content.includes('adapterRoutes'), true);
assert.equal(file(scaffold, 'README.md').content.includes('adapterBlueprints'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Adapter blueprints'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Adapter runtime entries'), true);
assert.equal(file(scaffold, 'README.md').content.includes('platformAdapter'), true);
assert.equal(file(scaffold, 'README.md').content.includes('observerPlans'), true);
assert.equal(file(scaffold, 'README.md').content.includes('resolvePlatform'), true);
assert.equal(file(scaffold, 'README.md').content.includes('resolvePlatformCandidates'), true);
assert.equal(file(scaffold, 'README.md').content.includes('observePlatformCandidates'), true);
assert.equal(file(scaffold, 'README.md').content.includes('meeting_timeline.observe_candidates'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Candidate observation contract'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Observer plan contract'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Meeting track contract'), true);
assert.equal(file(scaffold, 'README.md').content.includes('speaker_track_ready'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Platform conformance'), true);
assert.equal(file(scaffold, 'README.md').content.includes('platformConformance'), true);
assert.equal(file(scaffold, 'README.md').content.includes('Consumer handoff'), true);
assert.equal(file(scaffold, 'README.md').content.includes('consumerHandoff'), true);
assert.equal(file(scaffold, 'README.md').content.includes('extensionInstallPlan'), true);
assert.equal(file(scaffold, 'README.md').content.includes('integrationRuntimeSummary'), true);
assert.equal(file(scaffold, 'README.md').content.includes('runIntegrationRuntimeManifest'), true);

const acceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(scaffold);
assert.equal(acceptance.schema, MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA);
assert.equal(acceptance.accepted, true);
assert.equal(acceptance.blocking_count, 0);
assert.equal(acceptance.candidate_observation_ready, true);
assert.equal(acceptance.adapter_blueprint_ready, true);
assert.equal(acceptance.adapter_blueprint_ready_count, 3);
assert.equal(acceptance.adapter_blueprint_matrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(acceptance.adapter_runtime_ready, true);
assert.equal(acceptance.adapter_runtime_ready_count, 3);
assert.equal(acceptance.adapter_runtime_contract.rows.find((row) => row.platform === 'zoom').adapter_file, 'src/platform-adapters/zoom.mjs');
assert.equal(acceptance.candidate_observer_count, 3);
assert.equal(acceptance.candidate_observer_missing_count, 0);
assert.equal(acceptance.candidate_observation_contract.all_ready, true);
assert.equal(acceptance.platform_conformance_ready, true);
assert.equal(acceptance.platform_conformance_accepted_count, 3);
assert.equal(acceptance.platform_conformance_blocking_count, 0);
assert.equal(acceptance.platform_conformance_report.accepted, true);
assert.equal(acceptance.observer_plan_ready, true);
assert.equal(acceptance.observer_plan_ready_count, 3);
assert.equal(acceptance.observer_plan_preflight_accepted_count, 0);
assert.equal(acceptance.observer_plan_matrix.platform_count, 3);
assert.equal(acceptance.meeting_track_ready, true);
assert.equal(acceptance.speaker_track_ready_count, 3);
assert.equal(acceptance.participant_track_ready_count, 3);
assert.equal(acceptance.meeting_track_contract.all_ready, true);
assert.equal(acceptance.required_files.includes('src/http-routes.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-strategy.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/verify-contracts.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/verify-conformance.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-consumer-handoff.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-runtime-bundles.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-observer-plans.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-runtime-event-plans.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-adapter-routes.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-platform-adapters.mjs'), true);
assert.equal(acceptance.required_files.includes('src/platform-adapters/index.mjs'), true);
assert.equal(acceptance.required_files.includes('src/platform-adapters/google_meet.mjs'), true);
assert.equal(acceptance.required_files.includes('src/platform-adapters/microsoft_teams.mjs'), true);
assert.equal(acceptance.required_files.includes('src/platform-adapters/zoom.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/resolve-platform.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/resolve-platform-candidates.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/observe-platform-candidates.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-extension-plan.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-integration-runtime.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/print-integration-runtime-manifest.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/run-integration-runtime-manifest.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/run-handoff-readiness.mjs'), true);
assert.equal(assertMeetingPlatformHostIntegrationScaffold(scaffold).accepted, true);

const broken = {
  ...scaffold,
  files: scaffold.files.filter((item) => item.path !== 'src/http-routes.mjs'),
};
const brokenAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(broken);
assert.equal(brokenAcceptance.accepted, false);
assert.equal(brokenAcceptance.issues.some((item) => item.code === 'missing_required_file'), true);
assert.throws(
  () => assertMeetingPlatformHostIntegrationScaffold(broken),
  /Meeting platform host integration scaffold failed acceptance/,
);

const missingContractGate = {
  ...scaffold,
  files: scaffold.files.filter((item) => item.path !== 'scripts/verify-contracts.mjs'),
};
const missingContractGateAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(missingContractGate);
assert.equal(missingContractGateAcceptance.accepted, false);
assert.equal(missingContractGateAcceptance.issues.some((item) => item.path === 'scripts/verify-contracts.mjs'), true);

const missingPlatformAdapter = {
  ...scaffold,
  files: scaffold.files.filter((item) => item.path !== 'src/platform-adapters/google_meet.mjs'),
};
const missingPlatformAdapterAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(missingPlatformAdapter);
assert.equal(missingPlatformAdapterAcceptance.accepted, false);
assert.equal(missingPlatformAdapterAcceptance.issues.some((item) => item.path === 'src/platform-adapters/google_meet.mjs'), true);

const missingCandidateObservationContract = {
  ...scaffold,
  plan: {
    ...scaffold.plan,
    candidate_observation_contract: {
      ...scaffold.plan.candidate_observation_contract,
      ready_count: 2,
      missing_count: 1,
      all_ready: false,
    },
  },
};
const missingCandidateObservationContractAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(missingCandidateObservationContract);
assert.equal(missingCandidateObservationContractAcceptance.accepted, false);
assert.equal(
  missingCandidateObservationContractAcceptance.issues.some((item) => item.code === 'candidate_observation_contract_not_ready'),
  true,
);

const missingMeetingTrackContract = {
  ...scaffold,
  plan: {
    ...scaffold.plan,
    meeting_track_contract: {
      ...scaffold.plan.meeting_track_contract,
      speaker_ready_count: 2,
      missing_count: 1,
      all_ready: false,
    },
  },
};
const missingMeetingTrackContractAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(missingMeetingTrackContract);
assert.equal(missingMeetingTrackContractAcceptance.accepted, false);
assert.equal(
  missingMeetingTrackContractAcceptance.issues.some((item) => item.code === 'meeting_track_contract_not_ready'),
  true,
);

const adapterTmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapters-'));
try {
  for (const item of scaffold.files.filter((entry) => entry.path.startsWith('src/platform-adapters/'))) {
    const target = join(adapterTmpDir, item.path);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, item.content);
  }
  const adapterModule = await import(pathToFileURL(join(adapterTmpDir, 'src/platform-adapters/index.mjs')).href);
  const adapterCalls = [];
  const fakeHost = {
    resolvePlatform(input) { adapterCalls.push(['resolvePlatform', input]); return { accepted: true, input }; },
    resolvePlatformCandidates(input) { adapterCalls.push(['resolvePlatformCandidates', input]); return { accepted: true, input }; },
    observePlatformCandidates(input) { adapterCalls.push(['observePlatformCandidates', input]); return { accepted: true, input }; },
    insertAnnotation(platform, input) { adapterCalls.push(['insertAnnotation', platform, input]); return { accepted: true, platform, input }; },
    speakerTrack(platform, input) { adapterCalls.push(['speakerTrack', platform, input]); return { accepted: true, platform, input }; },
    participantTrack(platform, input) { adapterCalls.push(['participantTrack', platform, input]); return { accepted: true, platform, input }; },
    ingestProvider(platform, requestOrPayload, payload) {
      adapterCalls.push(['ingestProvider', platform, requestOrPayload, payload]);
      return { accepted: true, platform, requestOrPayload, payload };
    },
  };
  const googleAdapter = adapterModule.createMeetingPlatformAdapter(fakeHost, 'google-meet');
  assert.equal(googleAdapter.platform, 'google_meet');
  assert.equal(googleAdapter.selected_surface, 'browser_extension');
  assert.throws(
    () => googleAdapter.insertAnnotation({ label: 'why?' }),
    /captured_at_ms/,
  );
  const inserted = googleAdapter.insertAnnotation({ label: 'why?', captured_at_ms: 1234 });
  assert.equal(inserted.platform, 'google_meet');
  assert.equal(inserted.input.capturedAtMs, 1234);
  assert.equal(adapterCalls.at(-1)[0], 'insertAnnotation');
  assert.equal(adapterModule.createMeetingPlatformAdapters(fakeHost).zoom.selected_surface, 'native_detector');
} finally {
  await rm(adapterTmpDir, { recursive: true, force: true });
}

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  verify: false,
});

assert.equal(kit.platformHostIntegrationPlan({ platforms: ['zoom'] }).platforms[0], 'zoom');
assert.equal(kit.platformHostIntegrationScaffold({ platforms: ['zoom'] }).files.some((item) => item.path === 'src/meeting-platform-host.mjs'), true);
assert.equal(kit.platformHostIntegrationAcceptance(scaffold).accepted, true);
assert.equal(kit.assertPlatformHostIntegrationScaffold(scaffold).accepted, true);
assert.equal(kit.report({ platforms }).platform_host_integration.platforms.includes('microsoft_teams'), true);
assert.equal(kit.report({ platforms }).platform_host_integration_acceptance.accepted, true);

console.log('ok meeting platform host integration scaffold');
