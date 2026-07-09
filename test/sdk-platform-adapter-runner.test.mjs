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
  buildMeetingPlatformAdapterCandidateLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-launch-plan.mjs';
import {
  buildMeetingPlatformAdapterRuntimeManifest,
  buildMeetingPlatformAdapterRuntimeTarget,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runtime-recipe.mjs';
import {
  buildMeetingPlatformAdapterRunnerHandoff,
  createMeetingPlatformAdapterRunner,
  openMeetingPlatformAdapterSession,
  openMeetingPlatformAdapterRuntimeSession,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-runner.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

function node(tagName, attrs = {}, text = '') {
  return {
    tagName: tagName.toUpperCase(),
    attributes: attrs,
    dataset: Object.fromEntries(Object.entries(attrs)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [
        key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()),
        value,
      ])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attrs[name] ?? null;
    },
  };
}

function selectorAttrMatches(item, selector) {
  if (selector === 'button') return item.tagName === 'BUTTON';
  const attrParts = [...String(selector).matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (!attrParts.length) return false;
  return attrParts.every((match) => {
    const [, attrName, operator, doubleQuoted, singleQuoted, bare] = match;
    const actual = item.attributes[attrName];
    if (operator == null) return actual != null;
    if (actual == null) return false;
    const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
    if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
    return String(actual) === String(expected);
  });
}

function queryNodes(nodes, selector) {
  const text = String(selector);
  if (text === '*') return nodes;
  const generic = nodes.filter((item) => selectorAttrMatches(item, text));
  if (generic.length) return generic;
  if (text.includes('speaking')) {
    return nodes.filter((item) => /speaking|active speaker|正在发言|正在讲话|正在说话/i.test(item.attributes['aria-label'] ?? ''));
  }
  if (text.includes('aria-live')) return nodes.filter((item) => item.attributes['aria-live']);
  if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
  return [];
}

function fakeDocument({ url, title, nodes = [] }) {
  return {
    nodeType: 9,
    title,
    hidden: false,
    location: { href: url },
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
}

function liveMeetCandidate(title = 'Candidate runner review') {
  return {
    active: true,
    document: fakeDocument({
      url: 'https://meet.google.com/abc-defg-hij',
      title,
      nodes: [
        node('button', { 'aria-label': 'Leave call' }),
        node('button', { 'aria-label': 'Turn off microphone' }),
        node('div', {
          'data-participant-id': 'ada',
          'aria-label': 'Ada Lovelace is speaking',
        }),
      ],
    }),
  };
}

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
assert.equal(launchPlanRunner.getState().last_open_event.payload.adapter_selection_event.action, 'read_adapter_selection');
assert.equal(launchPlanRunner.getState().last_open_event.payload.adapter_selection_event.payload.axis_surface, 'browser_extension');
assert.equal(launchPlanRunner.getState().last_open_event.payload.raw_signal_event.action, 'validate_raw_signal');

const opened = await runner.open({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runner review',
});
assert.equal(opened.schema, 'meeting_platform_adapter_open_session_event');
assert.equal(opened.action, 'open_session');
assert.equal(opened.platform, 'google_meet');
assert.equal(opened.payload.launch_plan.platform, 'google_meet');
assert.equal(opened.payload.launch_plan.adapter_selection.axis_surface, 'browser_extension');
assert.equal(opened.payload.launch_plan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(opened.payload.adapter_selection_event.action, 'read_adapter_selection');
assert.equal(opened.payload.adapter_selection_event.payload.axis_source, 'local_observer_axis');
assert.equal(opened.payload.raw_signal_event.action, 'validate_raw_signal');
assert.equal(opened.payload.observe_event.action, 'observe_axis');
assert.equal(opened.payload.observe_event.captured_at_ms, 1_782_700_000_123);
assert.equal(runner.getState().opened, true);
assert.equal(runner.currentSession().platform, 'google_meet');
assert.equal(calls[0][0], 'platformRawSignalBatch');
assert.equal(calls[1][0], 'observePlatformCandidates');

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
assert.equal(openedWithoutObserve.payload.raw_signal_event.action, 'validate_raw_signal');
assert.equal(openedWithoutObserve.payload.observe_event, undefined);
await assert.rejects(
  () => runner.insertAnnotation({ label: 'needs observed axis' }),
  (error) => error.details?.code === 'adapter_session_axis_not_observed',
);
await runner.observeAxis({ platform: 'zoom' });
assert.equal((await runner.insertAnnotation({ label: 'after manual observe' })).payload.platform, 'zoom');

const candidateCalls = [];
const candidateClient = {
  async platformRawSignalBatch(payload) {
    candidateCalls.push(['platformRawSignalBatch', payload]);
    return { ok: true };
  },
  async observePlatformCandidates(payload) {
    candidateCalls.push(['observePlatformCandidates', payload]);
    return { ok: true };
  },
  async insertAnnotation(platform, payload) {
    candidateCalls.push(['insertAnnotation', platform, payload]);
    return { ok: true };
  },
};
const candidateRunner = createMeetingPlatformAdapterRunner(manifest, candidateClient, {
  clock: () => 1_782_700_000_456,
});
const candidateOpened = await candidateRunner.openCandidate({
  candidates: [liveMeetCandidate()],
}, {
  requireSpeakerTrack: true,
});
assert.equal(candidateOpened.plan_kind, 'candidate_launch_plan');
assert.equal(candidateOpened.payload.candidate_launch_plan.schema, 'meeting_platform_adapter_candidate_launch_plan');
assert.equal(candidateOpened.payload.session.plan_kind, 'candidate_launch_plan');
assert.equal(candidateOpened.payload.session.selected_evidence.interaction_can_leave, true);
assert.equal(candidateOpened.payload.observe_event.payload.preflight_evidence.active_speaker_candidate.name, 'Ada Lovelace');
assert.equal(candidateOpened.payload.observe_event.payload.candidates[0].semantic_signal_types.includes('meeting_leave_available'), true);
assert.equal(candidateOpened.payload.observe_event.payload.candidates[0].control_signal_summary.leave_available, true);
assert.equal(candidateRunner.getState().current_session.plan_kind, 'candidate_launch_plan');
assert.equal(candidateCalls[0][0], 'platformRawSignalBatch');
assert.equal(candidateCalls[1][0], 'observePlatformCandidates');
const candidateInserted = await candidateRunner.insertAnnotation({ label: 'candidate runner mark' });
assert.equal(candidateInserted.payload.preflight_evidence.interaction_can_leave, true);
assert.equal(candidateCalls[2][0], 'insertAnnotation');

const initialCandidateLaunchPlan = buildMeetingPlatformAdapterCandidateLaunchPlan(manifest, {
  candidates: [liveMeetCandidate('Initial candidate runner review')],
}, {
  requireSpeakerTrack: true,
  capturedAtMs: 1_782_700_000_789,
});
const initializedCandidateRunner = createMeetingPlatformAdapterRunner(initialCandidateLaunchPlan, candidateClient, {
  clock: () => 1_782_700_000_790,
});
const initializedCandidateOpened = await initializedCandidateRunner.open();
assert.equal(initializedCandidateOpened.plan_kind, 'candidate_launch_plan');
assert.equal(initializedCandidateOpened.payload.session.plan_kind, 'candidate_launch_plan');
assert.equal(initializedCandidateOpened.payload.observe_event.payload.preflight_evidence.accepted, true);

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
assert.equal(handoff.convenience_method, 'openMeetingPlatformAdapterSession');
assert.equal(handoff.optional_client_methods.includes('platformAdapterSelection'), true);
assert.equal(handoff.runtime_sequence.includes('read_adapter_selection_before_runtime_wiring'), true);
assert.equal(handoff.runtime_sequence.includes('build_launch_plan_from_current_url_or_platform'), true);

const runtimeManifest = buildMeetingPlatformAdapterRuntimeManifest({}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
const runtimeRunner = createMeetingPlatformAdapterRunner(runtimeManifest, adapterClient, {
  clock: () => 1_782_700_001_000,
});
assert.equal(runtimeRunner.runtime_manifest.schema, 'meeting_platform_adapter_runtime_manifest');
const runtimeTarget = runtimeRunner.runtimeTarget({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runtime runner review',
});
assert.equal(runtimeTarget.schema, 'meeting_platform_adapter_runtime_target');
assert.equal(runtimeTarget.host_kind, 'browser_extension_content_script');
const runtimeOpened = await runtimeRunner.open({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Runtime runner review',
});
assert.equal(runtimeOpened.plan_kind, 'runtime_target');
assert.equal(runtimeOpened.platform, 'google_meet');
assert.equal(runtimeOpened.payload.launch_plan, undefined);
assert.equal(runtimeOpened.payload.runtime_target.schema, 'meeting_platform_adapter_runtime_target');
assert.equal(runtimeOpened.payload.session.plan_kind, 'runtime_target');
assert.equal(runtimeOpened.payload.session.host_kind, 'browser_extension_content_script');
assert.equal(runtimeOpened.payload.observe_event.payload.current_url, 'https://meet.google.com/abc-defg-hij');
assert.equal(runtimeRunner.getState().current_runtime_target.host_kind, 'browser_extension_content_script');
assert.equal(runtimeRunner.currentSession().runtime_target.schema, 'meeting_platform_adapter_runtime_target');
assert.equal((await runtimeRunner.insertAnnotation({ label: 'runtime runner mark' })).payload.surface, 'browser_extension');

const runtimeTargetRunner = createMeetingPlatformAdapterRunner(runtimeTarget, adapterClient, {
  clock: () => 1_782_700_002_000,
});
assert.equal((await runtimeTargetRunner.open()).payload.runtime_target.host_kind, 'browser_extension_content_script');

const directRuntimeOpened = await openMeetingPlatformAdapterRuntimeSession(runtimeManifest, adapterClient, {
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  clock: () => 1_782_700_003_000,
});
assert.equal(directRuntimeOpened.payload.session.plan_kind, 'runtime_target');
assert.equal(directRuntimeOpened.payload.observe_event.captured_at_ms, 1_782_700_003_000);
assert.equal((await openMeetingPlatformAdapterSession(runtimeTarget, adapterClient, {}, {
  clock: () => 1_782_700_004_000,
})).payload.runtime_target.schema, 'meeting_platform_adapter_runtime_target');

const runtimeHandoff = buildMeetingPlatformAdapterRunnerHandoff(runtimeManifest);
assert.equal(runtimeHandoff.convenience_method, 'openMeetingPlatformAdapterRuntimeSession');
assert.equal(runtimeHandoff.runtime_manifest_schema, 'meeting_platform_adapter_runtime_manifest');
assert.equal(runtimeHandoff.runtime_sequence.includes('build_runtime_target_from_current_url_or_platform'), true);
assert.equal(runtimeHandoff.runtime_sequence.includes('create_session_from_runtime_target'), true);

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
assert.equal((await kit.openPlatformAdapterSession(runtimeManifest, { url: 'https://meet.google.com/abc-defg-hij' })).payload.session.plan_kind, 'runtime_target');
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
assert.equal((await rootSdk.openPlatformAdapterSession(runtimeManifest, { url: 'https://meet.google.com/abc-defg-hij' }, {
  client: adapterClient,
  clock: () => 790,
})).payload.session.plan_kind, 'runtime_target');

console.log('ok meeting platform adapter runner');
