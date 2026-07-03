import assert from 'node:assert/strict';

import {
  MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
  MEETING_APP_INTEGRATION_PROFILE_SCHEMA,
  buildAllMeetingAppIntegrationProfiles,
  buildMeetingAppIntegrationMatrix,
  buildMeetingAppIntegrationProfile,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';

assert.deepEqual(MEETING_APP_INTEGRATION_PROFILE_PLATFORMS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);
assert.equal(MEETING_APP_INTEGRATION_PROFILE_SCHEMA, 'meeting_app_integration_profile');

const googleProfile = buildMeetingAppIntegrationProfile('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
assert.equal(googleProfile.type, 'meeting_app_integration_profile');
assert.equal(googleProfile.platform, 'google_meet');
assert.equal(googleProfile.display_name, 'Google Meet');
assert.equal(googleProfile.recommended_mode, 'browser_extension_local_observer_first');
assert.deepEqual(googleProfile.extension.matches, ['https://meet.google.com/*']);
assert.equal(googleProfile.extension.recommended_permissions.includes('storage'), true);
assert.equal(googleProfile.extension.adapters.browser_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
assert.equal(googleProfile.extension.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(googleProfile.extension.timeline_endpoints.insertMark, '/api/annotations');
assert.equal(googleProfile.runtime.runtimePreset, 'google_meet');
assert.equal(googleProfile.runtime.observeMutations, true);
assert.equal(googleProfile.runtime.mutation_track_selector_count > 0, true);
assert.equal(googleProfile.capture.profile.platform, 'google_meet');
assert.equal(googleProfile.capture.participant_selector_count > 0, true);
assert.equal(googleProfile.event_model.realtime_axis.primary, 'browser_extension_local_observer');
assert.equal(googleProfile.event_model.speaker_activity.primary, 'local_dom_active_speaker_observer');
assert.equal(googleProfile.provider_reconciliation.enabled, true);
assert.equal(googleProfile.provider_reconciliation.endpoint, 'https://timeline.example.com/api/platform-events/google-meet');
assert.equal(googleProfile.readiness.sdk_wiring_ready, true);
assert.equal(googleProfile.readiness.production_ready, false);
assert.equal(googleProfile.readiness.next_actions.includes('capture_live_dom_snapshots_for_this_platform'), true);
assert.equal(googleProfile.launch_gate.passed, true);
assert.equal(googleProfile.launch_gate.production_ready, false);
assert.equal(googleProfile.implementation_steps.includes('insert_annotations_with_absolute_captured_at_ms'), true);

const teamsProfile = buildMeetingAppIntegrationProfile({
  platform: 'teams',
  includeLaunchGate: false,
});
assert.equal(teamsProfile.platform, 'microsoft_teams');
assert.equal(teamsProfile.extension.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal('launch_gate' in teamsProfile, false);
assert.equal('readiness' in teamsProfile, false);

const selectedProfiles = buildAllMeetingAppIntegrationProfiles({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedProfiles), ['zoom', 'webex']);
assert.equal(selectedProfiles.zoom.extension.matches.includes('https://zoom.us/*'), true);
assert.equal(selectedProfiles.webex.capture.profile.platform, 'webex');

const matrix = buildMeetingAppIntegrationMatrix({
  baseUrl: 'https://timeline.example.com',
});
assert.equal(matrix.type, 'meeting_app_integration_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.platforms.includes('lark'), true);
const larkRow = matrix.rows.find((row) => row.platform === 'lark');
assert.equal(larkRow.recommended_mode, 'browser_extension_local_observer_first');
assert.equal(larkRow.provider_reconciliation, true);
assert.equal(larkRow.sdk_wiring_ready, true);

assert.throws(
  () => buildMeetingAppIntegrationProfile('local-detector'),
  /Unsupported meeting app integration profile platform/,
);

console.log('ok meeting app integration profile');
