import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_SPEC_SCHEMA,
  assertMeetingAppAdapterSpec,
  assertMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpec,
  buildMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpecTemplate,
  builtInMeetingAppAdapterSpecMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-spec.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const google = buildMeetingAppAdapterSpec('google-meet');
assert.equal(google.schema, MEETING_APP_ADAPTER_SPEC_SCHEMA);
assert.equal(google.accepted, true);
assert.equal(google.adapter_key, 'google_meet');
assert.equal(google.source, 'built_in_manifest');
assert.equal(google.url_detection.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.contracts.timestamp_field, 'captured_at_ms');
assert.equal(google.contracts.provider_events_block_realtime, false);
assert.equal(google.contracts.transcript_blocks_realtime, false);
assert.equal(google.readiness.requires_live_snapshot_before_production, true);

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
const whereby = buildMeetingAppAdapterSpec(wherebySpec);
assert.equal(whereby.accepted, true);
assert.equal(whereby.adapter_key, 'whereby');
assert.equal(whereby.source, 'custom_spec');
assert.equal(whereby.readiness.capture_selector_ready, true);
assert.equal(assertMeetingAppAdapterSpec(whereby).adapter_key, 'whereby');

const rejected = buildMeetingAppAdapterSpec({
  adapter_key: 'broken_app',
  matches: ['https://broken.example/*'],
});
assert.equal(rejected.accepted, false);
assert.equal(rejected.issues.some((item) => item.code === 'missing_control_selectors'), true);
assert.throws(() => assertMeetingAppAdapterSpec(rejected), /Meeting app adapter spec acceptance failed/);

const matrix = buildMeetingAppAdapterSpecMatrix({
  platforms: ['google-meet', 'zoom'],
  adapters: [wherebySpec],
});
assert.equal(matrix.schema, MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA);
assert.equal(matrix.accepted, true);
assert.equal(matrix.spec_count, 3);
assert.equal(matrix.built_in_count, 2);
assert.equal(matrix.custom_count, 1);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.rows.find((row) => row.adapter_key === 'whereby').source, 'custom_spec');
assert.equal(assertMeetingAppAdapterSpecMatrix(matrix).accepted, true);

const builtInMatrix = builtInMeetingAppAdapterSpecMatrix({ platforms: ['google-meet', 'teams'] });
assert.equal(builtInMatrix.accepted, true);
assert.equal(builtInMatrix.spec_count, 2);
assert.equal(builtInMatrix.built_in_count, 2);

const template = buildMeetingAppAdapterSpecTemplate({ adapter_key: 'slack-huddle', matches: ['https://app.slack.com/*'] });
assert.equal(template.adapter_key, 'slack_huddle');
assert.equal(template.matches[0], 'https://app.slack.com/*');
assert.equal(buildMeetingAppAdapterSpec(template).accepted, true);

const kit = createMeetingPlatformTimelineKit({ baseUrl: 'http://localhost:8787' });
assert.equal(kit.meetingAppAdapterSpec('webex').adapter_key, 'webex');
assert.equal(kit.meetingAppAdapterSpecMatrix({ platforms: ['google-meet'], adapters: [wherebySpec] }).spec_count, 2);
assert.equal(kit.meetingAppAdapterSpecTemplate({ adapter_key: 'around' }).adapter_key, 'around');
assert.equal(kit.assertMeetingAppAdapterSpec(wherebySpec).accepted, true);
assert.equal(kit.assertMeetingAppAdapterSpecMatrix({ platforms: ['zoom'], adapters: [wherebySpec] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_spec_matrix.spec_count, 1);

console.log('ok meeting app adapter spec');
