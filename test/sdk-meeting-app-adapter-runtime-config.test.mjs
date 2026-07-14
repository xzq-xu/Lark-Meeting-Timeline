import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA,
  assertMeetingAppAdapterRuntimeConfig,
  assertMeetingAppAdapterRuntimeConfigMatrix,
  buildMeetingAppAdapterRuntimeConfig,
  buildMeetingAppAdapterRuntimeConfigMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config.mjs';
import { buildMeetingAppAdapterSpec } from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-spec.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const wherebySpec = {
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
  host_permissions: ['https://whereby.com/*'],
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
};

const whereby = buildMeetingAppAdapterRuntimeConfig(wherebySpec, {
  windowMessaging: true,
  allowedOrigins: ['https://whereby.com'],
});
assert.equal(whereby.schema, MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA);
assert.equal(whereby.accepted, true);
assert.equal(whereby.adapter_key, 'whereby');
assert.equal(whereby.browser_runtime_options.runtimePreset, false);
assert.equal(whereby.browser_runtime_options.observeMutations, true);
assert.equal(whereby.browser_runtime_options.captureOptions.captureProfile, false);
assert.equal(whereby.browser_runtime_options.captureOptions.controlSelectors.includes('[aria-label*="Leave" i]'), true);
assert.equal(whereby.browser_runtime_options.mutationTrackSelectors.includes('[data-participant-id]'), true);
assert.equal(whereby.content_script_options.windowMessaging, true);
assert.equal(whereby.content_script_options.allowedOrigins[0], 'https://whereby.com');
assert.equal(whereby.contracts.timestamp_field, 'captured_at_ms');
assert.equal(whereby.contracts.provider_events_block_realtime, false);
assert.equal(whereby.entrypoints.browser_runtime_factory.includes('createMeetingAppBrowserRuntime'), true);
assert.equal(assertMeetingAppAdapterRuntimeConfig(whereby).adapter_key, 'whereby');

const google = buildMeetingAppAdapterRuntimeConfig('google-meet');
assert.equal(google.accepted, true);
assert.equal(google.source_spec.source, 'built_in_manifest');
assert.equal(google.extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.browser_runtime_options.runtimePreset, false);
assert.equal(google.capture_options.participantSelectors.length > 0, true);

const brokenSpec = buildMeetingAppAdapterSpec({
  adapter_key: 'broken_app',
  matches: ['https://broken.example/*'],
});
const brokenRuntime = buildMeetingAppAdapterRuntimeConfig(brokenSpec);
assert.equal(brokenRuntime.accepted, false);
assert.equal(brokenRuntime.issues.some((item) => item.code === 'adapter_spec_not_accepted'), true);
assert.throws(() => assertMeetingAppAdapterRuntimeConfig(brokenRuntime), /Meeting app adapter runtime config acceptance failed/);

const matrix = buildMeetingAppAdapterRuntimeConfigMatrix({
  platforms: ['google-meet', 'zoom'],
  adapters: [wherebySpec],
});
assert.equal(matrix.schema, MEETING_APP_ADAPTER_RUNTIME_CONFIG_MATRIX_SCHEMA);
assert.equal(matrix.accepted, true);
assert.equal(matrix.config_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.built_in_count, 2);
assert.equal(matrix.custom_count, 1);
assert.equal(matrix.capture_ready_count, 3);
assert.equal(matrix.mutation_ready_count, 3);
assert.equal(assertMeetingAppAdapterRuntimeConfigMatrix(matrix).accepted, true);

const kit = createMeetingPlatformTimelineKit({ baseUrl: 'http://localhost:8787' });
assert.equal(kit.meetingAppAdapterRuntimeConfig(wherebySpec).adapter_key, 'whereby');
assert.equal(kit.meetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet'], adapters: [wherebySpec] }).config_count, 2);
assert.equal(kit.assertMeetingAppAdapterRuntimeConfig(wherebySpec).accepted, true);
assert.equal(kit.assertMeetingAppAdapterRuntimeConfigMatrix({ platforms: ['zoom'], adapters: [wherebySpec] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_runtime_config_matrix.config_count, 1);

console.log('ok meeting app adapter runtime config');
