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
assert.equal(google.browser.manifest.content_scripts[0].matches.includes('https://meet.google.com/*'), true);
assert.equal(google.modules.platform_integration_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(google.modules.content_script_bridge, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(google.modules.runtime_event, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event');
assert.equal(google.runtime.preset, 'google_meet');
assert.equal(google.runtime.start_options.runtimePreset, 'google_meet');
assert.equal(google.runtime.start_options.captureOptions.captureProfile, 'google_meet');
assert.equal(google.runtime.content_script_bridge.install_function, 'installMeetingPlatformIntegrationContentScriptBridge');
assert.deepEqual(google.runtime.content_script_bridge.options.platforms, ['google_meet']);
assert.equal(google.runtime.content_script_bridge.options.startOptions.runtimePreset, 'google_meet');
assert.equal(google.runtime.mutation_observer.enabled, true);
assert.equal(google.runtime.mutation_observer.debounce_ms, 150);
assert.equal(google.runtime.speaker_filter.min_stable_ms, 700);
assert.equal(google.messaging.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(google.messaging.bridge_message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(google.messaging.accepted_methods.includes('insertMark'), true);
assert.equal(google.messaging.accepted_methods.includes('runtimeEvents'), true);
assert.equal(google.messaging.runtime_event.schema, 'meeting_platform_runtime_event');
assert.equal(google.messaging.runtime_event.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.messaging.runtime_event.client_factory, 'createMeetingPlatformRuntimeEventClient');
assert.equal(google.messaging.examples.insert_annotation.method, 'insertMark');
assert.equal(google.messaging.examples.insert_annotation.input.captured_at_ms, 1_782_614_400_000);
assert.equal(google.messaging.examples.content_script_insert_annotation.type, 'meeting_timeline.insert_mark');
assert.equal(google.messaging.examples.content_script_insert_annotation.payload.mark.captured_at_ms, 1_782_614_400_000);
assert.equal(google.host.endpoints.insertMark, `${baseUrl}/api/annotations`);
assert.equal(google.host.endpoints.runtimeEvents, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.host.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.provider_reconcile.required_for_realtime, false);
assert.equal(google.transcript.blocks_realtime_annotation, false);
assert.equal(google.readiness.runtime_ready, true);
assert.equal(google.readiness.provider_required_for_realtime, false);
assert.equal(google.readiness.transcript_blocks_realtime, false);
assert.equal(google.adaptation_package.schema, 'meeting_platform_adaptation_package');

const localDetector = buildMeetingPlatformRuntimeBundle('local-detector', { baseUrl });
assert.equal(localDetector.platform, 'local_detector');
assert.equal(localDetector.browser.matches.length, 0);
assert.equal(localDetector.browser.manifest.content_scripts.length, 0);
assert.deepEqual(localDetector.runtime.content_script_bridge.options.platforms, []);
assert.equal(localDetector.readiness.runtime_ready, true);

const matrix = buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_runtime_bundle_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.runtime_ready_count, 5);
assert.equal(matrix.sdk_wiring_ready_count, 5);
assert.equal(matrix.provider_required_for_realtime_count, 0);
assert.equal(matrix.transcript_blocking_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').browser_match_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'webex').transcript_blocks_realtime, false);
assert.equal(matrix.next_actions.includes('install_meeting_platform_integration_content_script_bridge'), true);
assert.equal(matrix.next_actions.includes('start_meeting_app_content_script_bridge'), true);

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
assert.equal(kit.platformRuntimeBundleMatrix().platform_count, 2);
assert.equal(kit.report().platform_runtime_bundle_matrix.provider_required_for_realtime_count, 0);

console.log('ok meeting platform runtime bundle');
