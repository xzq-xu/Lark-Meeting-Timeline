import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_FIT_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_FIT_SCHEMA,
  buildMeetingAppAdapterFitMatrix,
  buildMeetingAppAdapterFitReport,
  createMeetingAppObserver,
  detectMeetingAppPreset,
  normalizeMeetingAppSnapshot,
  normalizeMeetingAppSnapshots,
  observeMeetingAppSample,
} from '../packages/meeting-timeline-sdk/adapters/meeting-apps.mjs';
import {
  MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA,
  MEETING_APP_ADAPTER_EXECUTION_PLAN_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_EXECUTION_PLAN_SCHEMA,
  buildMeetingAppAdapterCapabilityMatrix,
  buildMeetingAppAdapterCapabilityReport,
  buildMeetingAppAdapterExecutionPlan,
  buildMeetingAppAdapterExecutionPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-capability.mjs';
import {
  MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA,
  buildMeetingAppAdapterIntegrationPackage,
  buildMeetingAppAdapterIntegrationPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package.mjs';
import { createMeetingSourceAggregator } from '../packages/meeting-timeline-sdk/adapters/meeting-source.mjs';

const startMs = 1_783_010_400_000;

function googleMeetDomSnapshot(atMs = startMs) {
  return {
    browser: { name: 'Chrome' },
    windows: [{
      id: 'chrome-win',
      focused: true,
      tabs: [{
        id: 'meet-tab',
        active: true,
        audible: true,
        url: 'https://meet.google.com/abc-defg-hij',
        title: 'Design review - Google Meet',
        page: {
          documentVisible: true,
          buttons: [
            { ariaLabel: 'Turn off microphone' },
            { ariaLabel: 'Leave call' },
          ],
          tiles: [
            { dataset: { participantId: 'ada' }, ariaLabel: 'Ada Lovelace is speaking', audioLevel: 0.74 },
            { dataset: { participantId: 'grace' }, ariaLabel: 'Grace Hopper, muted' },
          ],
        },
      }],
    }],
    observedAtMs: atMs,
  };
}

let detected = detectMeetingAppPreset({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review - Google Meet',
});
assert.equal(detected.platform, 'google_meet');
assert.equal(detected.reason, 'url');

const normalized = normalizeMeetingAppSnapshots(googleMeetDomSnapshot());
assert.equal(normalized.length, 1);
assert.equal(normalized[0].platform, 'google_meet');
assert.equal(normalized[0].meeting_id, 'abc-defg-hij');
assert.equal(normalized[0].inMeeting, true);
assert.equal(normalized[0].activeSpeaker.name, 'Ada Lovelace');
assert.equal(normalized[0].activeSpeaker.speaking, true);
assert.equal(normalized[0].participants.length, 2);

const googleFit = buildMeetingAppAdapterFitReport(googleMeetDomSnapshot(), {
  platform: 'google-meet',
});
assert.equal(googleFit.type, 'meeting_app_adapter_fit_report');
assert.equal(googleFit.schema, MEETING_APP_ADAPTER_FIT_SCHEMA);
assert.equal(googleFit.platform, 'google_meet');
assert.deepEqual(googleFit.detected_platforms, ['google_meet']);
assert.equal(googleFit.accepted, true);
assert.equal(googleFit.ready_for_realtime_axis, true);
assert.equal(googleFit.ready_for_speaker_track, true);
assert.equal(googleFit.ready_for_participant_track, true);
assert.equal(googleFit.recommended_surface, 'browser_extension_or_webview');
assert.equal(googleFit.coverage.meeting_identity, true);
assert.equal(googleFit.coverage.meeting_start_candidate, true);
assert.equal(googleFit.rows[0].active_speaker_name, 'Ada Lovelace');
assert.equal(googleFit.issues.some((item) => item.code === 'missing_meeting_end_candidate'), true);

const fullEvidence = {
  live_dom_snapshot: true,
  candidate_observation: true,
  speaker_track: true,
  participant_track: true,
  annotation_insert_current_axis: true,
};
const googleCapability = buildMeetingAppAdapterCapabilityReport('google-meet', {
  input: googleMeetDomSnapshot(),
  evidence: fullEvidence,
});
assert.equal(googleCapability.schema, MEETING_APP_ADAPTER_CAPABILITY_REPORT_SCHEMA);
assert.equal(googleCapability.platform, 'google_meet');
assert.equal(googleCapability.static_ready, true);
assert.equal(googleCapability.pilot_ready, true);
assert.equal(googleCapability.production_ready, true);
assert.equal(googleCapability.recommended_mode, 'hybrid_local_observer_first');
assert.equal(googleCapability.timeline_capabilities.realtime_axis.status, 'local_ready');
assert.equal(googleCapability.timeline_capabilities.speaker_track.status, 'local_ready');
assert.equal(googleCapability.timeline_capabilities.post_meeting_transcript.provider_declared, true);
const googleExecutionPlan = buildMeetingAppAdapterExecutionPlan(googleCapability);
assert.equal(googleExecutionPlan.schema, MEETING_APP_ADAPTER_EXECUTION_PLAN_SCHEMA);
assert.equal(googleExecutionPlan.realtime_ready, true);
assert.equal(googleExecutionPlan.first_blocked_step, undefined);
assert.equal(googleExecutionPlan.steps.find((step) => step.id === 'insert_annotation_on_current_axis').status, 'ready');
assert.equal(googleExecutionPlan.gates.find((gate) => gate.id === 'live_evidence_package').status, 'passed');
const googleIntegrationPackage = buildMeetingAppAdapterIntegrationPackage(googleCapability);
assert.equal(googleIntegrationPackage.schema, MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_SCHEMA);
assert.equal(googleIntegrationPackage.platform, 'google_meet');
assert.equal(googleIntegrationPackage.accepted, true);
assert.equal(googleIntegrationPackage.realtime_ready, true);
assert.equal(googleIntegrationPackage.production_ready, true);
assert.equal(googleIntegrationPackage.file_paths.includes('capability-report.json'), true);
assert.equal(googleIntegrationPackage.file_paths.includes('execution-plan.json'), true);
assert.equal(googleIntegrationPackage.entrypoints.adapter_integration_package, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package');
assert.equal(googleIntegrationPackage.evidence_contract.required_for_realtime.includes('annotation_insert_current_axis'), true);
assert.equal(googleIntegrationPackage.integration_steps.some((step) => step.id === 'capture_meeting_app_snapshot'), true);

const explicitPlatformNormalized = normalizeMeetingAppSnapshot({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review - Google Meet',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, {
  platform: 'google_meet',
});
assert.equal(explicitPlatformNormalized.platform, 'google_meet');
assert.equal(explicitPlatformNormalized.meeting_id, 'abc-defg-hij');
assert.equal(explicitPlatformNormalized.meeting.meeting_id, 'abc-defg-hij');

let observed = observeMeetingAppSample(null, googleMeetDomSnapshot(), {
  source: 'browser_extension',
  minStableMs: 0,
  observedAtMs: startMs,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observed.signals[0].meeting.platform, 'google_meet');
assert.equal(observed.signals[0].meeting.meeting_id, 'abc-defg-hij');
assert.equal(observed.signals[1].speaker_name, 'Ada Lovelace');

const teams = normalizeMeetingAppSnapshot({
  application: { name: 'Microsoft Teams' },
  window: { title: 'Weekly sync | Microsoft Teams', focused: true },
  accessibility: {
    controls: [{ label: 'Leave' }, { label: 'Mute microphone' }],
    participants: [
      { id: 'sam', label: 'Sam Carter speaking' },
      { id: 'lin', label: 'Lin Zhang muted' },
    ],
  },
  meeting_id: 'teams-local-window',
  observedAtMs: startMs + 1_000,
});
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.inMeeting, true);
assert.equal(teams.activeSpeaker.id, 'sam');
assert.equal(teams.activeSpeaker.name, 'Sam Carter');

const fitMatrix = buildMeetingAppAdapterFitMatrix({
  platforms: ['google-meet', 'teams', 'zoom'],
  inputs: {
    google_meet: googleMeetDomSnapshot(),
    microsoft_teams: {
      application: { name: 'Microsoft Teams' },
      window: { title: 'Weekly sync | Microsoft Teams', focused: true },
      accessibility: {
        controls: [{ label: 'Leave' }, { label: 'Mute microphone' }],
        participants: [
          { id: 'sam', label: 'Sam Carter speaking' },
          { id: 'lin', label: 'Lin Zhang muted' },
        ],
      },
      meeting_id: 'teams-local-window',
      observedAtMs: startMs + 1_000,
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: {
        title: 'Zoom Meeting',
        focused: true,
        controls: [{ label: 'Leave Meeting' }, { label: 'Participants' }],
      },
      meeting_id: 'zoom-local-123',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
      observedAtMs: startMs + 2_000,
    },
  },
});
assert.equal(fitMatrix.schema, MEETING_APP_ADAPTER_FIT_MATRIX_SCHEMA);
assert.equal(fitMatrix.platform_count, 3);
assert.equal(fitMatrix.accepted_count, 3);
assert.equal(fitMatrix.realtime_axis_ready_count, 3);
assert.equal(fitMatrix.speaker_track_ready_count, 3);
assert.equal(fitMatrix.participant_track_ready_count, 3);
assert.equal(fitMatrix.rows.find((row) => row.platform === 'microsoft_teams').recommended_surface, 'native_detector');

const capabilityMatrix = buildMeetingAppAdapterCapabilityMatrix({
  platforms: ['google-meet', 'teams', 'zoom'],
  inputs: {
    google_meet: googleMeetDomSnapshot(),
    microsoft_teams: {
      application: { name: 'Microsoft Teams' },
      window: { title: 'Weekly sync | Microsoft Teams', focused: true },
      accessibility: {
        controls: [{ label: 'Leave' }, { label: 'Mute microphone' }],
        participants: [
          { id: 'sam', label: 'Sam Carter speaking' },
          { id: 'lin', label: 'Lin Zhang muted' },
        ],
      },
      meeting_id: 'teams-local-window',
      observedAtMs: startMs + 1_000,
    },
    zoom: {
      app: { name: 'Zoom Workplace' },
      window: {
        title: 'Zoom Meeting',
        focused: true,
        controls: [{ label: 'Leave Meeting' }, { label: 'Participants' }],
      },
      meeting_id: 'zoom-local-123',
      tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
      observedAtMs: startMs + 2_000,
    },
  },
});
assert.equal(capabilityMatrix.schema, MEETING_APP_ADAPTER_CAPABILITY_MATRIX_SCHEMA);
assert.equal(capabilityMatrix.platform_count, 3);
assert.equal(capabilityMatrix.accepted_count, 3);
assert.equal(capabilityMatrix.pilot_ready_count, 3);
assert.equal(capabilityMatrix.local_axis_ready_count, 3);
assert.equal(capabilityMatrix.production_ready_count, 0);
assert.equal(capabilityMatrix.rows.find((row) => row.platform === 'zoom').recommended_mode, 'hybrid_local_observer_first');
const executionPlanMatrix = buildMeetingAppAdapterExecutionPlanMatrix({
  capabilityMatrix,
});
assert.equal(executionPlanMatrix.schema, MEETING_APP_ADAPTER_EXECUTION_PLAN_MATRIX_SCHEMA);
assert.equal(executionPlanMatrix.platform_count, 3);
assert.equal(executionPlanMatrix.realtime_ready_count, 3);
assert.equal(executionPlanMatrix.rows.find((row) => row.platform === 'microsoft_teams').first_blocked_step, undefined);
const integrationPackageMatrix = buildMeetingAppAdapterIntegrationPackageMatrix({
  capabilityMatrix,
});
assert.equal(integrationPackageMatrix.schema, MEETING_APP_ADAPTER_INTEGRATION_PACKAGE_MATRIX_SCHEMA);
assert.equal(integrationPackageMatrix.platform_count, 3);
assert.equal(integrationPackageMatrix.pilot_ready_count, 3);
assert.equal(integrationPackageMatrix.realtime_ready_count, 3);
assert.equal(integrationPackageMatrix.rows.find((row) => row.platform === 'zoom').recommended_mode, 'hybrid_local_observer_first');

const zoomObserver = createMeetingAppObserver({
  source: 'desktop_observer',
  speakerOptions: { minStableMs: 0 },
});
const zoomControlsOnly = normalizeMeetingAppSnapshot({
  app: { name: 'Zoom Workplace' },
  window: {
    title: 'Zoom Meeting',
    controls: [{ label: 'Leave Meeting' }, { label: 'Participants' }],
  },
  meeting_id: 'zoom-controls-only',
});
assert.equal(zoomControlsOnly.platform, 'zoom');
assert.equal(zoomControlsOnly.inMeeting, true);

observed = zoomObserver.observe({
  app: { name: 'Zoom Workplace' },
  window: {
    id: 'zoom-win',
    title: 'Zoom Meeting',
    focused: true,
    controls: [{ label: 'Leave Meeting' }, { label: 'Participants' }],
  },
  meeting_id: 'zoom-local-123',
  tiles: [{ id: 'mira', ariaLabel: 'Mira Patel is speaking' }],
  observedAtMs: startMs + 2_000,
}, {
  observedAtMs: startMs + 2_000,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observed.signals[0].meeting.platform, 'zoom');
assert.equal(observed.signals[1].speaker_name, 'Mira Patel');

const lark = normalizeMeetingAppSnapshot({
  url: 'https://vc.feishu.cn/j/123456789',
  title: '会议进展实时可视化 - 飞书',
  page: {
    documentVisible: true,
    controls: [
      { label: '挂断' },
      { label: 'AI 视图' },
      { label: '共享屏幕' },
    ],
    participants: [
      { id: 'xzq', ariaLabel: '徐智强 正在发言' },
      { id: 'hold21', ariaLabel: 'hold21 已静音' },
    ],
  },
  observedAtMs: startMs + 2_500,
});
assert.equal(lark.platform, 'lark');
assert.equal(lark.meeting_id, '123456789');
assert.equal(lark.inMeeting, true);
assert.equal(lark.activeSpeaker.id, 'xzq');
assert.equal(lark.activeSpeaker.name, '徐智强');

const webex = normalizeMeetingAppSnapshot({
  url: 'https://example.webex.com/meet/product-review',
  title: 'Product review - Webex',
  page: {
    controls: [
      { label: 'Leave meeting' },
      { label: 'Unmute' },
      { label: 'Chat' },
    ],
    participants: [
      { id: 'maya', ariaLabel: 'Maya Chen, active speaker' },
      { id: 'noah', ariaLabel: 'Noah Smith, muted' },
    ],
  },
  observedAtMs: startMs + 2_800,
});
assert.equal(webex.platform, 'webex');
assert.equal(webex.meeting_id, 'meet-product-review');
assert.equal(webex.inMeeting, true);
assert.equal(webex.activeSpeaker.name, 'Maya Chen');

const calls = [];
const source = createMeetingSourceAggregator({
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, input };
  },
}, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
});

const applied = await source.observeMeetingApp(googleMeetDomSnapshot(startMs + 3_000), {
  observedAtMs: startMs + 3_000,
});
assert.equal(applied.source, 'meeting_app');
assert.deepEqual(applied.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(calls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(calls[0].input.platform, 'google_meet');
assert.equal(calls[1].input.kind, 'speaker_started');

console.log('ok meeting app presets');
