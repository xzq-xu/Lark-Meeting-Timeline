import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
  MEETING_APP_INTEGRATION_PROFILE_SCHEMA,
  MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
  buildAllMeetingAppIntegrationProfiles,
  buildAllMeetingAppRuntimeAdapterAcceptanceReports,
  buildAllMeetingAppRuntimeAdapterConfigs,
  buildAllMeetingAppRuntimeAdapterValidationReports,
  buildMeetingAppIntegrationMatrix,
  buildMeetingAppIntegrationProfile,
  buildMeetingAppRuntimeAdapterAcceptanceReport,
  buildMeetingAppRuntimeAdapterConfig,
  buildMeetingAppRuntimeAdapterValidationReport,
  assertMeetingAppRuntimeAdapterConfig,
  assertMeetingAppRuntimeAdapterValidation,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';

assert.deepEqual(MEETING_APP_INTEGRATION_PROFILE_PLATFORMS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);
assert.equal(MEETING_APP_INTEGRATION_PROFILE_SCHEMA, 'meeting_app_integration_profile');
assert.equal(MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA, 'meeting_app_runtime_adapter_config');

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

const googleRuntimeConfig = buildMeetingAppRuntimeAdapterConfig('google-meet');
assert.equal(googleRuntimeConfig.type, 'meeting_app_runtime_adapter_config');
assert.equal(googleRuntimeConfig.platform, 'google_meet');
assert.deepEqual(googleRuntimeConfig.extension.matches, ['https://meet.google.com/*']);
assert.equal(googleRuntimeConfig.extension.permissions.includes('storage'), true);
assert.equal(googleRuntimeConfig.bridge_options.browser_runtime_preset, 'google_meet');
assert.equal(googleRuntimeConfig.bridge_options.extensionMessaging, true);
assert.equal(googleRuntimeConfig.bridge_options.windowMessaging, true);
assert.equal(googleRuntimeConfig.runtime_options.runtimePreset, 'google_meet');
assert.equal(googleRuntimeConfig.runtime_options.observeMutations, true);
assert.equal(googleRuntimeConfig.capture_options.captureProfile, 'google_meet');
assert.equal(googleRuntimeConfig.capture_options.participantSelectors.length > 0, true);
assert.equal(googleRuntimeConfig.startup.attached_message_type, 'meeting_timeline.extension_attached');
assert.equal(googleRuntimeConfig.supported_client_methods.includes('insertMark'), true);
assert.equal(googleRuntimeConfig.readiness.runtime_ready, true);
const googleRuntimeAcceptance = buildMeetingAppRuntimeAdapterAcceptanceReport(googleRuntimeConfig);
assert.equal(googleRuntimeAcceptance.type, 'meeting_app_runtime_adapter_acceptance_report');
assert.equal(googleRuntimeAcceptance.accepted, true);
assert.equal(googleRuntimeAcceptance.coverage.bridge_preset, true);
assert.equal(googleRuntimeAcceptance.coverage.participant_selectors, true);
assert.equal(assertMeetingAppRuntimeAdapterConfig(googleRuntimeConfig).accepted, true);

const missingLiveValidation = buildMeetingAppRuntimeAdapterValidationReport(googleRuntimeConfig);
assert.equal(missingLiveValidation.accepted, false);
assert.equal(missingLiveValidation.production_ready, false);
assert.equal(missingLiveValidation.evidence_level, 'none');
assert.equal(missingLiveValidation.issues.some((item) => item.code === 'missing_meeting_app_evidence'), true);

const liveSnapshots = [
  buildMeetingAppFixtureSnapshot('google-meet', {
    state: 'active',
    observedAtMs: 1_783_356_000_000,
  }),
  buildMeetingAppFixtureSnapshot('google-meet', {
    state: 'prejoin',
    observedAtMs: 1_783_356_600_000,
  }),
];
const liveValidation = buildMeetingAppRuntimeAdapterValidationReport(googleRuntimeConfig, {
  snapshots: liveSnapshots,
});
assert.equal(liveValidation.accepted, true);
assert.equal(liveValidation.production_ready, true);
assert.equal(liveValidation.evidence_level, 'captured_dom');
assert.equal(liveValidation.launch_gate.production_ready, true);
assert.equal(assertMeetingAppRuntimeAdapterValidation(googleRuntimeConfig, {
  snapshots: liveSnapshots,
}).production_ready, true);

const unsafeRuntimeConfig = {
  ...googleRuntimeConfig,
  extension: {
    ...googleRuntimeConfig.extension,
    host_permissions: ['<all_urls>'],
  },
};
const unsafeAcceptance = buildMeetingAppRuntimeAdapterAcceptanceReport(unsafeRuntimeConfig);
assert.equal(unsafeAcceptance.accepted, false);
assert.equal(unsafeAcceptance.issues.some((item) => item.code === 'overbroad_host_permission'), true);
assert.throws(
  () => assertMeetingAppRuntimeAdapterConfig(unsafeRuntimeConfig),
  /Meeting app runtime adapter config acceptance failed/,
);

const stoppedRuntimeConfig = buildMeetingAppRuntimeAdapterConfig('google-meet', {
  startRuntime: false,
  windowMessaging: false,
});
assert.equal(stoppedRuntimeConfig.bridge_options.startRuntime, false);
assert.equal(stoppedRuntimeConfig.bridge_options.windowMessaging, false);
assert.equal(stoppedRuntimeConfig.startup.start_runtime, false);

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

const selectedRuntimeConfigs = buildAllMeetingAppRuntimeAdapterConfigs({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedRuntimeConfigs), ['zoom', 'webex']);
assert.equal(selectedRuntimeConfigs.zoom.bridge_options.browser_runtime_preset, 'zoom');
assert.equal(selectedRuntimeConfigs.webex.extension.matches.includes('https://*.webex.com/*'), true);

const selectedRuntimeAcceptance = buildAllMeetingAppRuntimeAdapterAcceptanceReports({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedRuntimeAcceptance), ['zoom', 'webex']);
assert.equal(selectedRuntimeAcceptance.zoom.accepted, true);
assert.equal(selectedRuntimeAcceptance.webex.coverage.storage_permission, true);

const selectedRuntimeValidation = buildAllMeetingAppRuntimeAdapterValidationReports({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedRuntimeValidation), ['zoom', 'webex']);
assert.equal(selectedRuntimeValidation.zoom.accepted, false);
assert.equal(selectedRuntimeValidation.webex.evidence_level, 'none');

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
