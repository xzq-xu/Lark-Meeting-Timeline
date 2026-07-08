import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-import-plan.mjs';
import {
  buildMeetingPlatformAdapterInstallManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-install-manifest.mjs';
import {
  buildMeetingPlatformAdapterMessageBridgeHandoff,
  createMeetingPlatformAdapterMessageBridge,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-message-bridge.mjs';
import {
  createMeetingPlatformAdapterRunner,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runner.mjs';
import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';
const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}, {
  baseUrl,
  target: 'static',
  includeArtifacts: true,
});
const availableFiles = exportMatrix.packages.flatMap((pkg) => pkg.host_files.map((file) => file.path));
const importMatrix = buildMeetingPlatformAdapterImportPlanMatrix(exportMatrix.packages, {
  availableFiles,
});
const manifest = buildMeetingPlatformAdapterInstallManifest(importMatrix.plans, {
  baseUrl,
});

const calls = [];
const adapterClient = {
  async platformRawSignalBatch(payload, options) {
    calls.push(['platformRawSignalBatch', payload, options]);
    return { ok: true, runtime_event_count: 1 };
  },
  async observePlatformCandidates(payload, options) {
    calls.push(['observePlatformCandidates', payload, options]);
    return { ok: true, observed: payload.platform };
  },
  async insertAnnotation(platform, payload, options) {
    calls.push(['insertAnnotation', platform, payload, options]);
    return { ok: true, platform, label: payload.label };
  },
  async speakerTrack(platform, payload, options) {
    calls.push(['speakerTrack', platform, payload, options]);
    return { ok: true, platform, speaker: payload.speaker_id };
  },
  async participantTrack(platform, payload, options) {
    calls.push(['participantTrack', platform, payload, options]);
    return { ok: true, platform, participant: payload.participant_id };
  },
  async ingestProvider(platform, payload, options) {
    calls.push(['ingestProvider', platform, payload, options]);
    return { ok: true, platform, event_type: payload.event_type };
  },
};

const bridge = createMeetingPlatformAdapterMessageBridge(manifest, adapterClient, {
  clock: () => 1_782_800_000_123,
});

assert.equal(bridge.schema, 'meeting_platform_adapter_message_bridge');
assert.equal((await bridge.handleMessage({ type: 'unsupported' })).handled, false);

const googleActiveSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_782_800_000_123,
});
const candidatePlan = await bridge.handleMessage({
  type: 'meeting_timeline.candidate_launch_plan',
  request_id: 'candidate-plan-1',
  payload: {
    tabs: [
      { id: 6, url: 'https://zoom.us/j/987654321', title: 'Zoom Meeting' },
      {
        id: 7,
        url: 'https://meet.google.com/abc-defg-hij',
        title: 'Google Meet',
        snapshots: [googleActiveSnapshot],
      },
    ],
  },
}, {
  requireSpeakerTrack: true,
});
assert.equal(candidatePlan.action, 'candidate_launch_plan');
assert.equal(candidatePlan.request_id, 'candidate-plan-1');
assert.equal(candidatePlan.result.accepted, true);
assert.equal(candidatePlan.result.status, 'ready_for_realtime_launch');
assert.equal(candidatePlan.result.launch_plan.platform, 'google_meet');
assert.equal(candidatePlan.adapter_selection_axis_source, 'local_observer_axis');
assert.equal(candidatePlan.adapter_selection_axis_surface, 'browser_extension');
assert.equal(candidatePlan.adapter_selection_timestamp_field, 'captured_at_ms');
assert.equal(candidatePlan.adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(candidatePlan.raw_signal_validation_status, 'ready');
assert.equal(candidatePlan.result.selected_candidate.tab_id, 7);

const urlOnlyCandidatePlan = await bridge.handleMessage({
  type: 'meeting_timeline.candidate_launch_plan',
  payload: {
    tabs: [
      { id: 8, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' },
    ],
  },
}, {
  requireSpeakerTrack: true,
});
assert.equal(urlOnlyCandidatePlan.action, 'candidate_launch_plan');
assert.equal(urlOnlyCandidatePlan.result.accepted, false);
assert.equal(urlOnlyCandidatePlan.result.status, 'needs_live_page_evidence');

const strictCalls = [];
const strictBridge = createMeetingPlatformAdapterMessageBridge(manifest, {
  async observePlatformCandidates(payload, options) {
    strictCalls.push(['observePlatformCandidates', payload, options]);
    return { ok: true, observed: payload.platform };
  },
  async insertAnnotation(platform, payload, options) {
    strictCalls.push(['insertAnnotation', platform, payload, options]);
    return { ok: true, platform, label: payload.label };
  },
}, {
  clock: () => 1_782_800_000_456,
});
const strictOpened = await strictBridge.handleMessage({
  type: 'meeting_timeline.open_candidate_session',
  payload: {
    tabs: [{
      id: 9,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
      snapshots: [googleActiveSnapshot],
    }],
  },
}, {
  requireSpeakerTrack: true,
});
assert.equal(strictOpened.action, 'open_candidate_session');
assert.equal(strictOpened.platform, 'google_meet');
assert.equal(strictOpened.result.payload.launch_plan.platform, 'google_meet');
assert.equal(strictOpened.adapter_selection_axis_surface, 'browser_extension');
assert.equal(strictOpened.adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(strictOpened.raw_signal_validation_status, 'ready');
assert.equal(strictOpened.result.payload.adapter_selection_event.action, 'read_adapter_selection');
assert.equal(strictOpened.result.payload.raw_signal_event.action, 'validate_raw_signal');
assert.equal(strictCalls[0][0], 'observePlatformCandidates');
assert.equal(strictCalls[0][1].candidates[0].tab_id, 9);

const opened = await bridge.handleMessage({
  type: 'meeting_timeline.observe_candidates',
  request_id: 'observe-1',
  payload: {
    tabs: [
      { url: 'https://example.com', title: 'Other', active: false },
      { url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true, in_meeting: true },
    ],
  },
});
assert.equal(opened.handled, true);
assert.equal(opened.action, 'open_session');
assert.equal(opened.request_id, 'observe-1');
assert.equal(opened.platform, 'google_meet');
assert.equal(opened.adapter_selection_axis_source, 'local_observer_axis');
assert.equal(opened.adapter_selection_axis_surface, 'browser_extension');
assert.equal(opened.adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(opened.adapter_blueprint_first_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(opened.raw_signal_validation_status, 'ready');
assert.equal(opened.result.payload.raw_signal_event.action, 'validate_raw_signal');
assert.equal(opened.result.payload.observe_event.action, 'observe_axis');
assert.equal(opened.result.payload.observe_event.captured_at_ms, 1_782_800_000_123);
assert.equal(calls[0][0], 'platformRawSignalBatch');
assert.equal(calls[1][0], 'observePlatformCandidates');

const inserted = await bridge.handleMessage({
  type: 'meeting_timeline.insert_mark',
  request_id: 'mark-1',
  payload: {
    mark: {
      label: 'why?',
      kind: 'question',
    },
  },
});
assert.equal(inserted.action, 'insert_annotation');
assert.equal(inserted.request_id, 'mark-1');
assert.equal(inserted.result.payload.label, 'why?');
assert.equal(inserted.result.payload.platform, 'google_meet');
assert.equal(calls.find((call) => call[0] === 'insertAnnotation' && call[2].label === 'why?')[1], 'google_meet');

assert.equal((await bridge.handleMessage({
  type: 'meeting_timeline.speaker_track',
  payload: { speaker: { speaker_id: 'speaker-1' } },
})).action, 'speaker_track');
assert.equal((await bridge.handleMessage({
  type: 'meeting_timeline.participant_track',
  payload: { participant: { participant_id: 'participant-1' } },
})).action, 'participant_track');
assert.equal((await bridge.handleMessage({
  type: 'meeting_timeline.provider_event',
  payload: { event: { event_type: 'meeting_ended' } },
})).action, 'provider_reconcile');
assert.equal((await bridge.handleMessage({ type: 'meeting_timeline.status' })).action, 'status');
assert.equal((await bridge.handleMessage({ type: 'meeting_timeline.reset' })).action, 'reset');

const autoOpenBridge = createMeetingPlatformAdapterMessageBridge(manifest, adapterClient, {
  clock: () => 222,
});
const autoInserted = await autoOpenBridge.handleMessage({
  type: 'meeting_timeline.insert_annotation',
  payload: {
    url: 'https://meet.google.com/abc-defg-hij',
    annotation: {
      label: 'auto open mark',
    },
  },
});
assert.equal(autoInserted.action, 'insert_annotation');
assert.equal(autoInserted.result.payload.captured_at_ms, 222);
assert.equal(autoOpenBridge.getState().runner.opened, true);

const runner = createMeetingPlatformAdapterRunner(manifest, adapterClient, {
  clock: () => 333,
});
const runnerBridge = createMeetingPlatformAdapterMessageBridge(runner, {
  clock: () => 333,
});
assert.equal((await runnerBridge.handleMessage({
  type: 'meeting_timeline.open_session',
  payload: { url: 'https://meet.google.com/abc-defg-hij' },
})).platform, 'google_meet');

const handoff = buildMeetingPlatformAdapterMessageBridgeHandoff(manifest);
assert.equal(handoff.schema, 'meeting_platform_adapter_message_bridge_handoff');
assert.equal(handoff.message_types.includes('meeting_timeline.insert_mark'), true);
assert.equal(handoff.message_types.includes('meeting_timeline.candidate_launch_plan'), true);
assert.equal(handoff.message_types.includes('meeting_timeline.open_candidate_session'), true);
assert.equal(handoff.supported_surfaces.includes('electron_webview_preload'), true);
assert.equal(handoff.runtime_sequence.includes('runner_reads_adapter_selection_before_runtime_wiring'), true);

const timelineCalls = [];
const timelineClient = {
  async startMeeting(input) {
    timelineCalls.push(['startMeeting', input]);
    return { ok: true, input };
  },
  async endMeeting(input) {
    timelineCalls.push(['endMeeting', input]);
    return { ok: true, input };
  },
  async insertMark(input) {
    timelineCalls.push(['insertMark', input]);
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(timelineClient, {
  baseUrl,
});
const kitBridge = kit.platformAdapterMessageBridge(manifest, {
  clock: () => 444,
});
await kitBridge.handleMessage({
  type: 'meeting_timeline.open_session',
  payload: { url: 'https://meet.google.com/abc-defg-hij' },
});
await kitBridge.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: { mark: { label: 'kit bridge mark' } },
});
assert.equal(timelineCalls.some((call) => call[0] === 'startMeeting'), true);
assert.equal(timelineCalls.some((call) => call[0] === 'insertMark' && call[1].label === 'kit bridge mark'), true);
assert.equal(kit.platformAdapterMessageBridgeHandoff(manifest).schema, 'meeting_platform_adapter_message_bridge_handoff');

const rootSdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet'],
});
const rootBridge = rootSdk.platformAdapterMessageBridge(manifest, adapterClient, {
  clock: () => 555,
});
assert.equal((await rootBridge.handleMessage({
  type: 'meeting_timeline.open_session',
  payload: { url: 'https://meet.google.com/abc-defg-hij' },
})).result.payload.observe_event.captured_at_ms, 555);
assert.equal(rootSdk.adapterMessageBridgeHandoff(manifest).schema, 'meeting_platform_adapter_message_bridge_handoff');

console.log('ok meeting platform adapter message bridge');
