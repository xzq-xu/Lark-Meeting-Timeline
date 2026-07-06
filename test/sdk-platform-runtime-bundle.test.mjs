import assert from 'node:assert/strict';

import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-runtime-bundle.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformRuntimeBundle('google-meet', { baseUrl });
assert.equal(google.schema, 'meeting_platform_runtime_bundle');
assert.equal(google.platform, 'google_meet');
assert.equal(google.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.deepEqual(google.browser.matches, ['https://meet.google.com/*']);
assert.equal(google.browser.permissions.includes('tabs'), true);
assert.equal(google.browser.candidate_observation.runtime_event_action, 'observe_platform_candidates');
assert.equal(google.browser.manifest.content_scripts[0].matches.includes('https://meet.google.com/*'), true);
assert.equal(google.modules.platform_integration_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(google.modules.content_script_bridge, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(google.modules.runtime_event, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event');
assert.equal(google.modules.meeting_platform_connector, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(google.modules.observer_plan, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile');
assert.equal(google.modules.observer_scheduler, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-observer-scheduler');
assert.equal(google.modules.runtime_host, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host');
assert.equal(google.adapter_route.schema, 'meeting_platform_adapter_route');
assert.equal(google.adapter_route.routes[0].route, 'local_observer_axis');
assert.equal(google.adapter_route.realtime_invariants.provider_events_block_realtime, false);
assert.equal(google.adapter_route.realtime_invariants.transcript_blocks_realtime, false);
assert.equal(google.runtime.preset, 'google_meet');
assert.equal(google.runtime.start_options.runtimePreset, 'google_meet');
assert.equal(google.runtime.start_options.captureOptions.captureProfile, 'google_meet');
assert.equal(google.runtime.content_script_bridge.install_function, 'installMeetingPlatformIntegrationContentScriptBridge');
assert.deepEqual(google.runtime.content_script_bridge.options.platforms, ['google_meet']);
assert.equal(google.runtime.content_script_bridge.options.startOptions.runtimePreset, 'google_meet');
assert.equal(google.runtime.lightweight_connector_bridge.module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(google.runtime.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(google.runtime.lightweight_connector_bridge.create_browser_runtime_function, 'createMeetingPlatformConnectorBrowserRuntime');
assert.deepEqual(google.runtime.lightweight_connector_bridge.options.platforms, ['google_meet']);
assert.equal(google.runtime.lightweight_connector_bridge.options.startOptions.runtimePreset, 'google_meet');
assert.equal(google.runtime.lightweight_connector_bridge.route_source_priority.includes('current location.href'), true);
assert.equal(google.runtime.mutation_observer.enabled, true);
assert.equal(google.runtime.mutation_observer.debounce_ms, 150);
assert.equal(google.runtime.speaker_filter.min_stable_ms, 700);
assert.equal(google.runtime.observer_plan.schema, 'meeting_app_runtime_observer_plan');
assert.equal(google.runtime.observer_plan.platform, 'google_meet');
assert.equal(google.runtime.observer_plan.preflight_status, 'not_run');
assert.equal(google.runtime.observation_loop.factory, 'createMeetingAppBrowserRuntime');
assert.equal(google.runtime.observation_loop.timestamp_field, 'captured_at_ms');
assert.equal(google.runtime.observation_loop.cadence.fallback_poll_interval_ms, 10000);
assert.equal(google.runtime.observation_loop.trigger_policy.some((item) => item.trigger === 'dom_mutation'), true);
assert.equal(google.runtime.observer_scheduler.create_function, 'createMeetingAppObserverScheduler');
assert.equal(google.runtime.observer_scheduler.config.schema, 'meeting_app_observer_scheduler_config');
assert.equal(google.runtime.observer_scheduler.config.platform, 'google_meet');
assert.equal(google.runtime.runtime_host.create_function, 'createMeetingPlatformRuntimeHost');
assert.equal(google.runtime.runtime_host.config_function, 'buildMeetingPlatformRuntimeHostConfig');
assert.equal(google.messaging.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(google.messaging.bridge_message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(google.messaging.bridge_message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(google.messaging.lightweight_connector_message_types.includes('meeting_timeline.sample_tracks'), true);
assert.equal(google.messaging.lightweight_connector_message_types.includes('meeting_timeline.observe_candidates'), true);
assert.equal(google.messaging.lightweight_connector_message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(google.messaging.background_message_types.includes('meeting_timeline.observe_candidates'), true);
assert.equal(google.messaging.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.messaging.accepted_methods.includes('insertMark'), true);
assert.equal(google.messaging.accepted_methods.includes('runtimeEvents'), true);
assert.equal(google.messaging.runtime_event.schema, 'meeting_platform_runtime_event');
assert.equal(google.messaging.runtime_event.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.messaging.runtime_event.client_factory, 'createMeetingPlatformRuntimeEventClient');
assert.equal(google.messaging.runtime_event.plan_schema, 'meeting_platform_runtime_event_plan');
assert.equal(google.messaging.runtime_event.plan.platform, 'google_meet');
assert.equal(google.messaging.runtime_event.plan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(google.messaging.runtime_event.plan.actions.find((row) => row.action === 'speaker_track').client_method, 'speakerTrack');
assert.equal(google.messaging.examples.insert_annotation.method, 'insertMark');
assert.equal(google.messaging.examples.observe_candidates.type, 'meeting_timeline.observe_candidates');
assert.equal(google.messaging.examples.observe_candidates.tabs[0].url, 'https://meet.google.com/');
assert.equal(google.messaging.examples.preflight_current_window.type, 'meeting_timeline.preflight_current_window');
assert.equal(google.messaging.examples.preflight_current_window.options.requireSpeakerTrack, true);
assert.equal(google.messaging.examples.insert_annotation.input.captured_at_ms, 1_782_614_400_000);
assert.equal(google.messaging.examples.content_script_insert_annotation.type, 'meeting_timeline.insert_mark');
assert.equal(google.messaging.examples.content_script_insert_annotation.payload.mark.captured_at_ms, 1_782_614_400_000);
assert.equal(google.host.endpoints.insertMark, `${baseUrl}/api/annotations`);
assert.equal(google.host.endpoints.runtimeEvents, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.host.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.provider_reconcile.required_for_realtime, false);
assert.equal(google.transcript.blocks_realtime_annotation, false);
assert.equal(google.readiness.runtime_ready, true);
assert.equal(google.readiness.observer_plan_ready, true);
assert.equal(google.readiness.observer_scheduler_ready, true);
assert.equal(google.readiness.runtime_host_ready, true);
assert.equal(google.readiness.lightweight_connector_ready, true);
assert.equal(google.readiness.observer_preflight_status, 'not_run');
assert.equal(google.readiness.provider_required_for_realtime, false);
assert.equal(google.readiness.transcript_blocks_realtime, false);
assert.equal(google.adaptation_package.schema, 'meeting_platform_adaptation_package');

const localDetector = buildMeetingPlatformRuntimeBundle('local-detector', { baseUrl });
assert.equal(localDetector.platform, 'local_detector');
assert.equal(localDetector.browser.matches.length, 0);
assert.equal(localDetector.browser.manifest.content_scripts.length, 0);
assert.deepEqual(localDetector.runtime.content_script_bridge.options.platforms, []);
assert.deepEqual(localDetector.runtime.lightweight_connector_bridge.options.platforms, []);
assert.equal(localDetector.readiness.lightweight_connector_ready, false);
assert.equal(localDetector.runtime.observer_plan, null);
assert.equal(localDetector.observer_plan, null);
assert.equal(localDetector.readiness.runtime_ready, true);

const matrix = buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_runtime_bundle_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.runtime_ready_count, 5);
assert.equal(matrix.sdk_wiring_ready_count, 5);
assert.equal(matrix.observer_plan_ready_count, 5);
assert.equal(matrix.runtime_host_ready_count, 5);
assert.equal(matrix.lightweight_connector_ready_count, 5);
assert.equal(matrix.adapter_route_ready_count, 5);
assert.equal(matrix.local_observer_first_count, 5);
assert.equal(matrix.provider_non_blocking_route_count, 5);
assert.equal(matrix.transcript_non_blocking_route_count, 5);
assert.equal(matrix.candidate_observer_count, 5);
assert.equal(matrix.provider_required_for_realtime_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').browser_match_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_first_route, 'local_observer_axis');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').adapter_recommended_mode, 'local_observer_first_provider_reconcile');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').candidate_observation_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').candidate_observer_message_type, 'meeting_timeline.observe_candidates');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').candidate_observer_permission, 'tabs');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').observer_plan_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').observer_scheduler_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').runtime_host_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').lightweight_connector_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').lightweight_connector_install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').observer_preflight_status, 'not_run');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').observer_factory, 'createMeetingAppBrowserRuntime');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').observer_meeting_end_grace_ms, 4000);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'webex').transcript_blocks_realtime, false);
assert.equal(matrix.next_actions.includes('install_meeting_platform_integration_content_script_bridge'), true);
assert.equal(matrix.next_actions.includes('install_meeting_platform_connector_content_script_bridge'), true);
assert.equal(matrix.next_actions.includes('start_meeting_app_content_script_bridge'), true);
assert.equal(matrix.next_actions.includes('wire_observer_plan_to_host_scheduler'), true);

const client = {
  async startMeeting(input) {
    return { ok: true, input };
  },
  async endMeeting(input) {
    return { ok: true, input };
  },
  async insertMark(input) {
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformRuntimeBundle('google-meet').browser.matches[0], 'https://meet.google.com/*');
assert.equal(kit.platformRuntimeBundle('google-meet').runtime.observation_loop.timestamp_field, 'captured_at_ms');
assert.equal(kit.platformRuntimeBundleMatrix().platform_count, 2);
assert.equal(kit.platformRuntimeBundleMatrix().observer_plan_ready_count, 2);
assert.equal(kit.platformRuntimeBundleMatrix().lightweight_connector_ready_count, 2);
assert.equal(kit.platformRuntimeEventPlan('google-meet').examples.insert_annotation.annotation.label, 'why?');
assert.equal(kit.platformRuntimeEventPlanMatrix().platform_count, 2);
assert.equal(kit.report().platform_runtime_bundle_matrix.provider_required_for_realtime_count, 0);
assert.equal(kit.report().platform_runtime_event_plan_matrix.realtime_provider_dependency_count, 0);

console.log('ok meeting platform runtime bundle');
