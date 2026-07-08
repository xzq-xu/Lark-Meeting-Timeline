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
  buildMeetingPlatformAdapterLaunchPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-launch-plan.mjs';
import {
  buildMeetingPlatformAdapterRuntimeManifest,
  buildMeetingPlatformAdapterRuntimeTarget,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runtime-recipe.mjs';
import {
  buildMeetingPlatformAdapterSessionHandoff,
  createMeetingPlatformAdapterSession,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-session.mjs';
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
});
const availableFiles = exportMatrix.packages.flatMap((pkg) => pkg.host_files.map((file) => file.path));
const importMatrix = buildMeetingPlatformAdapterImportPlanMatrix(exportMatrix.packages, {
  availableFiles,
});
const manifest = buildMeetingPlatformAdapterInstallManifest(importMatrix.plans, {
  baseUrl,
});
const googlePlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Adapter session review',
});

assert.equal(googlePlan.accepted, true);

const calls = [];
const adapterClient = {
  async platformRawSignalBatch(payload, options) {
    calls.push(['platformRawSignalBatch', payload, options]);
    return { ok: true, runtime_event_count: 1 };
  },
  async observePlatformCandidates(payload, options) {
    calls.push(['observePlatformCandidates', payload, options]);
    return { ok: true, axis: payload.candidates?.[0]?.platform };
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
    return { ok: true, platform, event: payload.event_type };
  },
};

const session = createMeetingPlatformAdapterSession(googlePlan, adapterClient, {
  clock: () => 1_782_614_400_321,
});

assert.equal(session.schema, 'meeting_platform_adapter_session');
assert.equal(session.platform, 'google_meet');
assert.equal(session.selected_surface, 'browser_extension');
assert.equal(session.adapter_selection.axis_surface, 'browser_extension');
assert.equal(session.getState().adapter_selection_read, false);
assert.equal(session.getState().axis_observed, false);
await assert.rejects(
  () => session.insertAnnotation({ label: 'why?' }),
  (error) => error.details?.code === 'adapter_session_raw_signal_not_validated',
);

const validated = await session.validateRawSignal({
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.equal(validated.action, 'validate_raw_signal');
assert.equal(validated.payload.raw_signals[0].platform, 'google_meet');
assert.equal(session.getState().adapter_selection_read, true);
assert.equal(session.getState().last_adapter_selection_event.action, 'read_adapter_selection');
assert.equal(session.getState().last_adapter_selection_event.payload.axis_surface, 'browser_extension');
assert.equal(session.getState().raw_signal_validated, true);
assert.equal(calls[0][0], 'platformRawSignalBatch');

const observed = await session.observeAxis({
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.equal(observed.schema, 'meeting_platform_adapter_session_event');
assert.equal(observed.action, 'observe_axis');
assert.equal(observed.payload.candidates[0].platform, 'google_meet');
assert.equal(observed.payload.captured_at_ms, 1_782_614_400_321);
assert.equal(session.getState().axis_observed, true);
assert.equal(calls[1][0], 'observePlatformCandidates');

const inserted = await session.insertAnnotation({
  label: 'why?',
  kind: 'handwriting',
});
assert.equal(inserted.action, 'insert_annotation');
assert.equal(inserted.payload.platform, 'google_meet');
assert.equal(inserted.payload.surface, 'browser_extension');
assert.equal(inserted.payload.captured_at_ms, 1_782_614_400_321);
assert.equal(calls[2][0], 'insertAnnotation');
assert.equal(calls[2][1], 'google_meet');
assert.equal(calls[2][2].label, 'why?');

assert.equal((await session.insertMark({ label: 'alias mark' })).action, 'insert_annotation');
assert.equal((await session.speakerTrack({ speaker_id: 's1', active: true })).action, 'speaker_track');
assert.equal((await session.participantTrack({ participant_id: 'p1', state: 'joined' })).action, 'participant_track');
assert.equal((await session.providerReconcile({ event_type: 'meeting_ended' })).action, 'provider_reconcile');
assert.equal(calls.map((call) => call[0]).includes('speakerTrack'), true);
assert.equal(calls.map((call) => call[0]).includes('participantTrack'), true);
assert.equal(calls.map((call) => call[0]).includes('ingestProvider'), true);

const handoff = buildMeetingPlatformAdapterSessionHandoff(googlePlan);
assert.equal(handoff.schema, 'meeting_platform_adapter_session_handoff');
assert.equal(handoff.session_factory, 'createMeetingPlatformAdapterSession');
assert.equal(handoff.plan_kind, 'launch_plan');
assert.equal(handoff.input_schema, 'meeting_platform_adapter_launch_plan');
assert.equal(handoff.required_client_methods.includes('observePlatformCandidates'), true);
assert.equal(handoff.optional_client_methods.includes('platformAdapterSelection'), true);
assert.equal(handoff.next_actions.includes('call_session_readAdapterSelection_before_runtime_wiring'), true);
assert.equal(handoff.next_actions.includes('call_session_insertAnnotation_for_realtime_marks'), true);

const runtimeManifest = buildMeetingPlatformAdapterRuntimeManifest({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
const googleTarget = buildMeetingPlatformAdapterRuntimeTarget(runtimeManifest, {
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runtime target review',
});
assert.equal(googleTarget.accepted, true);

const targetCalls = [];
const targetClient = {
  async observePlatformCandidates(payload, options) {
    targetCalls.push(['observePlatformCandidates', payload, options]);
    return { ok: true, observed_platform: payload.platform };
  },
  async insertAnnotation(platform, payload, options) {
    targetCalls.push(['insertAnnotation', platform, payload, options]);
    return { ok: true, platform, label: payload.label };
  },
};
const targetSession = createMeetingPlatformAdapterSession(googleTarget, targetClient, {
  clock: () => 1_782_614_401_000,
});
assert.equal(targetSession.plan_kind, 'runtime_target');
assert.equal(targetSession.platform, 'google_meet');
assert.equal(targetSession.selected_surface, 'browser_extension');
assert.equal(targetSession.host_kind, 'browser_extension_content_script');
assert.equal(targetSession.bridge_kind, 'browser_content_script');
assert.equal(targetSession.launch_plan, undefined);
assert.equal(targetSession.runtime_target.schema, 'meeting_platform_adapter_runtime_target');
assert.equal(targetSession.axis_contract.timestamp_field, 'captured_at_ms');
assert.equal(targetSession.getState().plan_kind, 'runtime_target');

const targetObserved = await targetSession.observeAxis({
  title: 'Runtime target review',
});
assert.equal(targetObserved.action, 'observe_axis');
assert.equal(targetObserved.payload.platform, 'google_meet');
assert.equal(targetObserved.payload.surface, 'browser_extension');
assert.equal(targetObserved.payload.current_url, 'https://meet.google.com/abc-defg-hij');
assert.equal(targetCalls[0][0], 'observePlatformCandidates');

const targetInserted = await targetSession.insertAnnotation({
  label: 'runtime target mark',
});
assert.equal(targetInserted.action, 'insert_annotation');
assert.equal(targetInserted.payload.platform, 'google_meet');
assert.equal(targetInserted.payload.surface, 'browser_extension');
assert.equal(targetInserted.payload.source, 'meeting_platform_adapter_session');
assert.equal(targetCalls[1][0], 'insertAnnotation');
assert.equal(targetCalls[1][1], 'google_meet');

const targetHandoff = buildMeetingPlatformAdapterSessionHandoff(googleTarget);
assert.equal(targetHandoff.plan_kind, 'runtime_target');
assert.equal(targetHandoff.input_schema, 'meeting_platform_adapter_runtime_target');
assert.equal(targetHandoff.runtime_target_schema, 'meeting_platform_adapter_runtime_target');
assert.equal(targetHandoff.launch_plan_schema, undefined);
assert.equal(targetHandoff.host_kind, 'browser_extension_content_script');
assert.equal(targetHandoff.next_actions.includes('create_session_from_runtime_target'), true);

const bridgeCalls = [];
const bridgeStyleClient = {
  async observe(payload) {
    bridgeCalls.push(['observe', payload]);
    return { ok: true };
  },
  async insertAnnotation(payload) {
    bridgeCalls.push(['insertAnnotation', payload]);
    return { ok: true };
  },
};
const bridgeSession = createMeetingPlatformAdapterSession(googlePlan, bridgeStyleClient, {
  clock: () => 88,
  platformArgument: false,
});
await bridgeSession.observeAxis();
const bridgeInserted = await bridgeSession.insertAnnotation({ label: 'bridge mark' });
assert.equal(bridgeCalls[1][0], 'insertAnnotation');
assert.equal(bridgeCalls[1][1].platform, 'google_meet');
assert.equal(bridgeInserted.payload.captured_at_ms, 88);

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
const kitSession = kit.platformAdapterSession(googlePlan, {
  clock: () => 99,
});
assert.equal((await kitSession.validateRawSignal()).action, 'validate_raw_signal');
await kitSession.observeAxis({}, { validateRawSignal: false });
await kitSession.insertAnnotation({ label: 'kit mark' });
assert.equal(timelineCalls.some((call) => call[0] === 'insertMark' && call[1].platform === 'google_meet'), true);
assert.equal(kit.platformAdapterSessionHandoff(googlePlan).schema, 'meeting_platform_adapter_session_handoff');

const rootSdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet'],
});
const rootSession = rootSdk.platformAdapterSession(googlePlan, adapterClient, {
  clock: () => 123,
});
await rootSession.observeAxis();
assert.equal((await rootSession.insertAnnotation({ label: 'root mark' })).payload.captured_at_ms, 123);
assert.equal(rootSdk.adapterSessionHandoff(googlePlan).schema, 'meeting_platform_adapter_session_handoff');

console.log('ok meeting platform adapter session');
