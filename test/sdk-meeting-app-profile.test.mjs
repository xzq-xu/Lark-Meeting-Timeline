import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
  MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA,
  MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA,
  MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
  MEETING_APP_INTEGRATION_PROFILE_SCHEMA,
  MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA,
  MEETING_APP_LIVE_SNAPSHOT_CAPTURE_PLAN_SCHEMA,
  MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
  MEETING_APP_RUNTIME_ADAPTER_PROFILE_MATRIX_SCHEMA,
  MEETING_APP_RUNTIME_ADAPTER_PROFILE_RESOLUTION_SCHEMA,
  assertMeetingAppDeploymentManifest,
  buildAllMeetingAppDeploymentManifests,
  buildAllMeetingAppDomAdaptationDiagnoses,
  buildMeetingAppDomAdaptationDiagnosisMatrix,
  buildAllMeetingAppDeploymentManifestAcceptanceReports,
  buildAllMeetingAppIntegrationProfiles,
  buildAllMeetingAppLiveSnapshotCapturePlans,
  buildAllMeetingAppRuntimeAdapterAcceptanceReports,
  buildAllMeetingAppRuntimeAdapterConfigs,
  buildAllMeetingAppRuntimeAdapterValidationReports,
  buildMeetingAppDeploymentManifest,
  buildMeetingAppDeploymentManifestAcceptanceReport,
  buildMeetingAppDeploymentManifestAcceptanceSummary,
  buildMeetingAppDomAdaptationDiagnosis,
  buildMeetingAppIntegrationMatrix,
  buildMeetingAppIntegrationProfile,
  buildMeetingAppLiveEvidencePackage,
  buildMeetingAppLiveEvidencePackageSummary,
  buildMeetingAppLiveSnapshotCapturePlan,
  buildMeetingAppRuntimeAdapterAcceptanceReport,
  buildMeetingAppRuntimeAdapterConfig,
  buildMeetingAppRuntimeAdapterProfileMatrix,
  buildMeetingAppRuntimeAdapterValidationReport,
  resolveMeetingAppRuntimeAdapterProfile,
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
assert.equal(MEETING_APP_RUNTIME_ADAPTER_PROFILE_RESOLUTION_SCHEMA, 'meeting_app_runtime_adapter_profile_resolution');
assert.equal(MEETING_APP_RUNTIME_ADAPTER_PROFILE_MATRIX_SCHEMA, 'meeting_app_runtime_adapter_profile_matrix');
assert.equal(MEETING_APP_LIVE_SNAPSHOT_CAPTURE_PLAN_SCHEMA, 'meeting_app_live_snapshot_capture_plan');
assert.equal(MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA, 'meeting_app_deployment_manifest');
assert.equal(MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA, 'meeting_app_live_evidence_package');
assert.equal(MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA, 'meeting_app_dom_adaptation_diagnosis');
assert.equal(MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA, 'meeting_app_dom_adaptation_diagnosis_matrix');

const googleProfile = buildMeetingAppIntegrationProfile('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
assert.equal(googleProfile.type, 'meeting_app_integration_profile');
assert.equal(googleProfile.platform, 'google_meet');
assert.equal(googleProfile.display_name, 'Google Meet');
assert.equal(googleProfile.recommended_mode, 'browser_extension_local_observer_first');
assert.deepEqual(googleProfile.extension.matches, ['https://meet.google.com/*']);
assert.equal(googleProfile.extension.recommended_permissions.includes('storage'), true);
assert.equal(googleProfile.extension.recommended_permissions.includes('tabs'), true);
assert.equal(googleProfile.extension.adapters.browser_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
assert.equal(googleProfile.extension.message_types.client_call, 'meeting_timeline.client_call');
assert.equal(googleProfile.extension.message_types.observe_candidates, 'meeting_timeline.observe_candidates');
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
assert.equal(googleRuntimeConfig.extension.permissions.includes('tabs'), true);
assert.equal(googleRuntimeConfig.extension.message_types.observe_candidates, 'meeting_timeline.observe_candidates');
assert.equal(googleRuntimeConfig.readiness.runtime_ready, true);

const googleResolved = resolveMeetingAppRuntimeAdapterProfile({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
});
assert.equal(googleResolved.type, 'meeting_app_runtime_adapter_profile_resolution');
assert.equal(googleResolved.schema, MEETING_APP_RUNTIME_ADAPTER_PROFILE_RESOLUTION_SCHEMA);
assert.equal(googleResolved.detected, true);
assert.equal(googleResolved.platform, 'google_meet');
assert.equal(googleResolved.reason, 'capture_profile');
assert.deepEqual(googleResolved.extension.matches, ['https://meet.google.com/*']);
assert.equal(googleResolved.capture.profile.platform, 'google_meet');
assert.equal(googleResolved.runtime.options.runtimePreset, 'google_meet');
assert.equal(googleResolved.tracks.output_intents.includes('speaker_track'), true);
assert.equal(googleResolved.tracks.runtime_options.speakerTrackOptions.minStableMs >= 500, true);
assert.equal(googleResolved.tracks.content_policy, 'position_markers_only_no_transcript_text_required');
assert.equal(googleResolved.readiness.runtime_ready, true);
assert.equal(googleResolved.next_actions.includes('enable_observeTracks_or_trackMutations_when_speaker_position_marks_are_needed'), true);

const teamsResolved = resolveMeetingAppRuntimeAdapterProfile('https://teams.microsoft.com/l/meetup-join/19%3ameeting');
assert.equal(teamsResolved.detected, true);
assert.equal(teamsResolved.platform, 'microsoft_teams');
assert.equal(teamsResolved.runtime_config.extension.matches.includes('https://teams.microsoft.com/*'), true);

const unknownResolved = resolveMeetingAppRuntimeAdapterProfile({
  url: 'https://example.com/not-a-meeting',
  title: 'Not a meeting',
});
assert.equal(unknownResolved.detected, false);
assert.equal(unknownResolved.platform, null);
assert.equal(unknownResolved.issues.some((item) => item.code === 'platform_not_detected'), true);

const profileMatrix = buildMeetingAppRuntimeAdapterProfileMatrix({
  platforms: ['google-meet', 'teams', 'zoom'],
  inputs: {
    google_meet: 'https://meet.google.com/abc-defg-hij',
    microsoft_teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting',
  },
});
assert.equal(profileMatrix.type, 'meeting_app_runtime_adapter_profile_matrix');
assert.equal(profileMatrix.schema, MEETING_APP_RUNTIME_ADAPTER_PROFILE_MATRIX_SCHEMA);
assert.deepEqual(profileMatrix.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
assert.equal(profileMatrix.platform_count, 3);
assert.equal(profileMatrix.detected_count, 3);
assert.equal(profileMatrix.runtime_ready_count, 3);
assert.equal(profileMatrix.track_profile_count, 3);
assert.equal(profileMatrix.rows.find((row) => row.platform === 'google_meet').extension_match_count, 1);
assert.equal(profileMatrix.rows.find((row) => row.platform === 'microsoft_teams').runtime_preset, 'microsoft_teams');
assert.equal(profileMatrix.rows.find((row) => row.platform === 'zoom').capture_profile, 'zoom');
assert.equal(profileMatrix.next_actions.includes('enable_observeTracks_or_trackMutations_when_speaker_position_marks_are_needed'), true);

const googleCapturePlan = buildMeetingAppLiveSnapshotCapturePlan('google-meet');
assert.equal(googleCapturePlan.type, 'meeting_app_live_snapshot_capture_plan');
assert.equal(googleCapturePlan.platform, 'google_meet');
assert.equal(googleCapturePlan.runtime_config.platform, 'google_meet');
assert.equal(googleCapturePlan.validation.method, 'buildMeetingAppRuntimeAdapterValidationReport');
assert.equal(googleCapturePlan.validation.production_ready_requires, 'captured_dom');
assert.equal(googleCapturePlan.required_snapshots.some((item) => item.id === 'active_speaker'), true);
assert.equal(googleCapturePlan.required_snapshots.some((item) => item.id === 'meeting_ended'), true);
assert.equal(googleCapturePlan.required_snapshots.find((item) => item.id === 'active_speaker').required_coverage.includes('speaker_started'), true);
assert.equal(googleCapturePlan.minimum_record_count, 2);
assert.match(googleCapturePlan.handoff.success_condition, /production_ready/);

const googleDeploymentManifest = buildMeetingAppDeploymentManifest('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
assert.equal(googleDeploymentManifest.type, 'meeting_app_deployment_manifest');
assert.equal(googleDeploymentManifest.schema, MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA);
assert.equal(googleDeploymentManifest.platform, 'google_meet');
assert.equal(googleDeploymentManifest.profile.platform, 'google_meet');
assert.equal(googleDeploymentManifest.runtime_config.platform, 'google_meet');
assert.deepEqual(googleDeploymentManifest.extension_install_plan.platforms, ['google_meet']);
assert.equal(googleDeploymentManifest.live_snapshot_capture_plan.minimum_record_count, 2);
assert.equal(googleDeploymentManifest.validation_report.accepted, false);
assert.equal(googleDeploymentManifest.production_gate.requires_captured_dom, true);
assert.equal(googleDeploymentManifest.production_gate.minimum_live_record_count, 2);
assert.equal(googleDeploymentManifest.integration_targets.some((item) => item.surface === 'chrome_or_edge_extension'), true);
assert.equal(googleDeploymentManifest.integration_targets.find((item) => item.surface === 'chrome_or_edge_extension').permissions.includes('tabs'), true);
assert.equal(googleDeploymentManifest.integration_targets.find((item) => item.surface === 'chrome_or_edge_extension').candidate_observer.runtime_event_action, 'observe_platform_candidates');
assert.equal(googleDeploymentManifest.integration_targets.some((item) => item.surface === 'electron_or_embedded_webview'), true);
assert.equal(googleDeploymentManifest.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(googleDeploymentManifest.runtime_contract.required_signals.includes('speaker_started'), true);
assert.equal(googleDeploymentManifest.runtime_contract.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(googleDeploymentManifest.runtime_contract.candidate_observation.required_permission, 'tabs');
assert.equal(googleDeploymentManifest.handoff.kit_methods.includes('meetingAppDeploymentManifest'), true);
assert.equal(googleDeploymentManifest.rollout_checklist.includes('pass_runtime_validation_with_production_ready_true'), true);
const googleDeploymentAcceptance = buildMeetingAppDeploymentManifestAcceptanceReport(googleDeploymentManifest);
assert.equal(googleDeploymentAcceptance.accepted, true);
assert.equal(googleDeploymentAcceptance.production_ready, false);
assert.equal(googleDeploymentAcceptance.coverage.runtime_config, true);
assert.equal(googleDeploymentAcceptance.coverage.live_dom_verified, false);
assert.equal(googleDeploymentAcceptance.issues.some((item) => item.code === 'live_dom_not_verified'), true);
assert.equal(assertMeetingAppDeploymentManifest(googleDeploymentManifest).accepted, true);
const unsafeDeploymentManifest = {
  ...googleDeploymentManifest,
  extension_install_plan: {
    ...googleDeploymentManifest.extension_install_plan,
    host_permissions: ['<all_urls>'],
  },
};
const unsafeDeploymentAcceptance = buildMeetingAppDeploymentManifestAcceptanceReport(unsafeDeploymentManifest);
assert.equal(unsafeDeploymentAcceptance.accepted, false);
assert.equal(unsafeDeploymentAcceptance.issues.some((item) => item.code === 'overbroad_host_permission'), true);
assert.throws(
  () => assertMeetingAppDeploymentManifest(unsafeDeploymentManifest),
  /Meeting app deployment manifest acceptance failed/,
);
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
const liveDeploymentAcceptance = buildMeetingAppDeploymentManifestAcceptanceReport('google-meet', {
  snapshots: liveSnapshots,
});
assert.equal(liveDeploymentAcceptance.accepted, true);
assert.equal(liveDeploymentAcceptance.production_ready, true);
assert.equal(liveDeploymentAcceptance.coverage.live_dom_verified, true);
assert.equal(liveDeploymentAcceptance.issues.some((item) => item.code === 'live_dom_not_verified'), false);
const liveEvidencePackage = buildMeetingAppLiveEvidencePackage({
  packageId: 'evidence-google-001',
  snapshots: {
    'google-meet': liveSnapshots,
  },
});
assert.equal(liveEvidencePackage.type, 'meeting_app_live_evidence_package');
assert.equal(liveEvidencePackage.schema, MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA);
assert.equal(liveEvidencePackage.id, 'evidence-google-001');
assert.deepEqual(liveEvidencePackage.platforms, ['google_meet']);
assert.equal(liveEvidencePackage.record_count, 2);
assert.equal(liveEvidencePackage.accepted, true);
assert.equal(liveEvidencePackage.production_ready, true);
assert.equal(liveEvidencePackage.records_by_platform.google_meet.record_count, 2);
assert.equal(liveEvidencePackage.manifest_acceptance.google_meet.production_ready, true);
assert.equal(liveEvidencePackage.summary.production_ready_count, 1);
assert.equal(liveEvidencePackage.summary.rows[0].missing_required_coverage.length, 0);
const liveEvidenceSummary = buildMeetingAppLiveEvidencePackageSummary(liveEvidencePackage);
assert.equal(liveEvidenceSummary.production_ready, true);
assert.equal(liveEvidenceSummary.record_count, 2);
const liveDomDiagnosis = buildMeetingAppDomAdaptationDiagnosis('google-meet', {
  snapshots: liveSnapshots,
});
assert.equal(liveDomDiagnosis.type, 'meeting_app_dom_adaptation_diagnosis');
assert.equal(liveDomDiagnosis.schema, MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA);
assert.equal(liveDomDiagnosis.platform, 'google_meet');
assert.equal(liveDomDiagnosis.accepted, true);
assert.equal(liveDomDiagnosis.production_ready, true);
assert.equal(liveDomDiagnosis.selector_probe.matched.controls, true);
assert.equal(liveDomDiagnosis.selector_probe.matched.participants, true);
assert.equal(liveDomDiagnosis.selector_probe.matched.active_speaker, true);
assert.equal(liveDomDiagnosis.observer_probe.signal_types.includes('meeting_started'), true);
assert.equal(liveDomDiagnosis.observer_probe.signal_types.includes('speaker_started'), true);
assert.equal(liveDomDiagnosis.observer_probe.signal_types.includes('meeting_ended'), true);
assert.deepEqual(liveDomDiagnosis.recommended_capture.required_snapshot_ids, ['active_speaker', 'meeting_ended']);

const noEvidenceDomDiagnosis = buildMeetingAppDomAdaptationDiagnosis('google-meet');
assert.equal(noEvidenceDomDiagnosis.accepted, false);
assert.equal(noEvidenceDomDiagnosis.production_ready, false);
assert.equal(noEvidenceDomDiagnosis.issues.some((item) => item.code === 'missing_live_snapshots'), true);
assert.equal(noEvidenceDomDiagnosis.next_actions.includes('capture_required_live_snapshots'), true);

for (const [platformAlias, expectedPlatform] of [
  ['google-meet', 'google_meet'],
  ['teams', 'microsoft_teams'],
  ['zoom', 'zoom'],
  ['webex', 'webex'],
]) {
  const platformSnapshots = [
    buildMeetingAppFixtureSnapshot(platformAlias, {
      state: 'active',
      observedAtMs: 1_783_356_000_000,
    }),
    buildMeetingAppFixtureSnapshot(platformAlias, {
      state: 'prejoin',
      observedAtMs: 1_783_356_600_000,
    }),
  ];
  const diagnosis = buildMeetingAppDomAdaptationDiagnosis(platformAlias, {
    snapshots: platformSnapshots,
  });
  assert.equal(diagnosis.platform, expectedPlatform);
  assert.equal(diagnosis.accepted, true);
  assert.equal(diagnosis.production_ready, true);
  assert.equal(diagnosis.selector_probe.matched.controls, true);
  assert.equal(diagnosis.selector_probe.matched.participants, true);
  assert.equal(diagnosis.observer_probe.coverage.meeting_started, true);
  assert.equal(diagnosis.observer_probe.coverage.speaker_started, true);
  assert.equal(diagnosis.observer_probe.coverage.meeting_ended, true);
}

const selectedDomDiagnoses = buildAllMeetingAppDomAdaptationDiagnoses({
  platforms: ['zoom'],
  snapshots: {
    zoom: [
      buildMeetingAppFixtureSnapshot('zoom', {
        state: 'active',
        observedAtMs: 1_783_356_000_000,
      }),
      buildMeetingAppFixtureSnapshot('zoom', {
        state: 'prejoin',
        observedAtMs: 1_783_356_600_000,
      }),
    ],
  },
});
assert.deepEqual(Object.keys(selectedDomDiagnoses), ['zoom']);
assert.equal(selectedDomDiagnoses.zoom.production_ready, true);
const domDiagnosisMatrix = buildMeetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet', 'zoom'],
  snapshots: {
    'google-meet': liveSnapshots,
    zoom: [
      buildMeetingAppFixtureSnapshot('zoom', {
        state: 'active',
        observedAtMs: 1_783_356_000_000,
      }),
      buildMeetingAppFixtureSnapshot('zoom', {
        state: 'prejoin',
        observedAtMs: 1_783_356_600_000,
      }),
    ],
  },
});
assert.equal(domDiagnosisMatrix.type, 'meeting_app_dom_adaptation_diagnosis_matrix');
assert.equal(domDiagnosisMatrix.schema, MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA);
assert.equal(domDiagnosisMatrix.platform_count, 2);
assert.equal(domDiagnosisMatrix.accepted_count, 2);
assert.equal(domDiagnosisMatrix.production_ready_count, 2);
assert.equal(domDiagnosisMatrix.active_speaker_ready_count, 2);
assert.equal(domDiagnosisMatrix.meeting_start_ready_count, 2);
assert.equal(domDiagnosisMatrix.meeting_end_ready_count, 2);
assert.equal(domDiagnosisMatrix.rows.find((row) => row.platform === 'google_meet').active_speaker_matched, true);
const missingEvidencePackage = buildMeetingAppLiveEvidencePackage({
  platforms: ['google-meet'],
  snapshots: [],
});
assert.equal(missingEvidencePackage.accepted, false);
assert.equal(missingEvidencePackage.production_ready, false);
assert.equal(missingEvidencePackage.issues.some((item) => item.code === 'missing_platform_evidence_records'), true);

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

const selectedCapturePlans = buildAllMeetingAppLiveSnapshotCapturePlans({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedCapturePlans), ['zoom', 'webex']);
assert.equal(selectedCapturePlans.zoom.required_snapshots.length >= 2, true);
assert.equal(selectedCapturePlans.webex.runtime_config.platform, 'webex');

const selectedDeploymentManifests = buildAllMeetingAppDeploymentManifests({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedDeploymentManifests), ['zoom', 'webex']);
assert.equal(selectedDeploymentManifests.zoom.extension_install_plan.matches.includes('https://zoom.us/*'), true);
assert.equal(selectedDeploymentManifests.webex.production_gate.requires_captured_dom, true);
const selectedDeploymentAcceptance = buildAllMeetingAppDeploymentManifestAcceptanceReports({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedDeploymentAcceptance), ['zoom', 'webex']);
assert.equal(selectedDeploymentAcceptance.zoom.accepted, true);
assert.equal(selectedDeploymentAcceptance.webex.production_ready, false);
const deploymentSummary = buildMeetingAppDeploymentManifestAcceptanceSummary({
  platforms: ['zoom', 'webex'],
});
assert.equal(deploymentSummary.accepted, true);
assert.equal(deploymentSummary.production_ready, false);
assert.equal(deploymentSummary.platform_count, 2);
assert.equal(deploymentSummary.accepted_count, 2);
assert.equal(deploymentSummary.production_ready_count, 0);
assert.equal(deploymentSummary.rows.every((row) => row.warning_count >= 1), true);

const selectedRuntimeAcceptance = buildAllMeetingAppRuntimeAdapterAcceptanceReports({
  platforms: ['zoom', 'webex'],
});
assert.deepEqual(Object.keys(selectedRuntimeAcceptance), ['zoom', 'webex']);
assert.equal(selectedRuntimeAcceptance.zoom.accepted, true);
assert.equal(selectedRuntimeAcceptance.webex.coverage.storage_permission, true);
assert.equal(selectedRuntimeAcceptance.webex.coverage.tabs_permission, true);
assert.equal(selectedRuntimeAcceptance.webex.coverage.candidate_observer_message_type, true);

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
