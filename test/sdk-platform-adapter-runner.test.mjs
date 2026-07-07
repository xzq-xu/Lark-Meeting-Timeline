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
  buildMeetingPlatformAdapterRunnerHandoff,
  createMeetingPlatformAdapterRunner,
  openMeetingPlatformAdapterSession,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runner.mjs';
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

const runner = createMeetingPlatformAdapterRunner(manifest, adapterClient, {
  clock: () => 1_782_700_000_123,
});

assert.equal(runner.schema, 'meeting_platform_adapter_runner');
assert.equal(runner.getState().opened, false);
await assert.rejects(
  () => runner.insertAnnotation({ label: 'before open' }),
  (error) => error.details?.code === 'adapter_runner_session_not_open',
);

const launchPlan = runner.launchPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runner review',
});
assert.equal(launchPlan.accepted, true);
assert.equal(launchPlan.platform, 'google_meet');
assert.equal(launchPlan.adapter_blueprint.primary_surface, 'browser_extension');

const launchPlanRunner = createMeetingPlatformAdapterRunner(launchPlan, adapterClient, {
  clock: () => 321,
});
assert.equal((await launchPlanRunner.open()).payload.observe_event.captured_at_ms, 321);

const opened = await runner.open({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runner review',
});
assert.equal(opened.schema, 'meeting_platform_adapter_open_session_event');
assert.equal(opened.action, 'open_session');
assert.equal(opened.platform, 'google_meet');
assert.equal(opened.payload.launch_plan.platform, 'google_meet');
assert.equal(opened.payload.launch_plan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(opened.payload.observe_event.action, 'observe_axis');
assert.equal(opened.payload.observe_event.captured_at_ms, 1_782_700_000_123);
assert.equal(runner.getState().opened, true);
assert.equal(runner.currentSession().platform, 'google_meet');
assert.equal(calls[0][0], 'observePlatformCandidates');

const mark = await runner.insertAnnotation({
  label: 'why?',
});
assert.equal(mark.action, 'insert_annotation');
assert.equal(mark.payload.platform, 'google_meet');
assert.equal(mark.payload.captured_at_ms, 1_782_700_000_123);
const firstInsertCall = calls.find((call) => call[0] === 'insertAnnotation' && call[2]?.label === 'why?');
assert.equal(firstInsertCall[1], 'google_meet');

assert.equal((await runner.insertMark({ label: 'alias mark' })).action, 'insert_annotation');
assert.equal((await runner.speakerTrack({ speaker_id: 'speaker-1' })).action, 'speaker_track');
assert.equal((await runner.participantTrack({ participant_id: 'participant-1' })).action, 'participant_track');
assert.equal((await runner.providerReconcile({ event_type: 'meeting_ended' })).action, 'provider_reconcile');

runner.reset();
assert.equal(runner.getState().opened, false);

const zoomPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  platform: 'zoom',
});
const openedWithoutObserve = await runner.open(zoomPlan, {
  observe: false,
});
assert.equal(openedWithoutObserve.payload.observe_event, undefined);
await assert.rejects(
  () => runner.insertAnnotation({ label: 'needs observed axis' }),
  (error) => error.details?.code === 'adapter_session_axis_not_observed',
);
await runner.observeAxis({ platform: 'zoom' });
assert.equal((await runner.insertAnnotation({ label: 'after manual observe' })).payload.platform, 'zoom');

const directOpened = await openMeetingPlatformAdapterSession(manifest, adapterClient, {
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  clock: () => 123,
});
assert.equal(directOpened.schema, 'meeting_platform_adapter_open_session_event');
assert.equal(directOpened.payload.observe_event.captured_at_ms, 123);

const handoff = buildMeetingPlatformAdapterRunnerHandoff(manifest);
assert.equal(handoff.schema, 'meeting_platform_adapter_runner_handoff');
assert.equal(handoff.runner_factory, 'createMeetingPlatformAdapterRunner');
assert.equal(handoff.runtime_sequence.includes('build_launch_plan_from_current_url_or_platform'), true);

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
const kitRunner = kit.platformAdapterRunner(manifest, {
  clock: () => 456,
});
await kitRunner.open({
  url: 'https://meet.google.com/abc-defg-hij',
});
await kitRunner.insertAnnotation({ label: 'kit runner mark' });
assert.equal(timelineCalls.some((call) => call[0] === 'startMeeting'), true);
assert.equal(timelineCalls.some((call) => call[0] === 'insertMark' && call[1].platform === 'google_meet'), true);
assert.equal((await kit.openPlatformAdapterSession(manifest, { url: 'https://meet.google.com/abc-defg-hij' })).schema, 'meeting_platform_adapter_open_session_event');
assert.equal(kit.platformAdapterRunnerHandoff(manifest).schema, 'meeting_platform_adapter_runner_handoff');

const rootSdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet'],
});
const rootRunner = rootSdk.platformAdapterRunner(manifest, adapterClient, {
  clock: () => 789,
});
await rootRunner.open({
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.equal((await rootRunner.insertAnnotation({ label: 'root runner mark' })).payload.captured_at_ms, 789);
assert.equal(rootSdk.adapterRunnerHandoff(manifest).schema, 'meeting_platform_adapter_runner_handoff');

console.log('ok meeting platform adapter runner');
