import assert from 'node:assert/strict';

import {
  createNativeMeetingObserver,
  createNativeMeetingTimelineObserver,
  normalizeNativeActiveSpeakerSample,
  normalizeNativeMeetingCandidate,
  normalizeNativeMeetingCandidates,
  observeNativeMeetingSample,
  selectNativeMeetingCandidate,
} from '../packages/meeting-timeline-sdk/adapters/native-meeting.mjs';

const startMs = 1_782_787_200_000;

function zoomNative(observedAtMs, activeSpeaker = null) {
  return {
    applications: [{
      name: 'zoom.us',
      bundleId: 'us.zoom.xos',
      windows: [{
        id: 'zoom-call-1',
        title: 'Daily Standup - Zoom Meeting',
        focused: true,
        visible: true,
        inMeeting: true,
        accessibility: {
          callActive: true,
          activeSpeaker,
        },
      }],
    }],
    observedAtMs,
  };
}

let candidates = normalizeNativeMeetingCandidates(zoomNative(startMs, { id: 'ada', name: 'Ada', speaking: true }), {
  observedAtMs: startMs,
});
assert.equal(candidates.length, 1);
assert.equal(candidates[0].platform, 'zoom');
assert.equal(candidates[0].meeting_id, 'native-zoom-daily-standup');
assert.equal(candidates[0].in_meeting, true);
assert.equal(candidates[0].discovery.platform_reason, 'bundle_id');

let selected = selectNativeMeetingCandidate(zoomNative(startMs), { observedAtMs: startMs });
assert.equal(selected.detectedMeeting.platform, 'zoom');
assert.equal(selected.detectedMeeting.meeting_id, 'native-zoom-daily-standup');

const teamsCandidate = normalizeNativeMeetingCandidate({
  application: {
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
  },
  window: {
    id: 'teams-call-1',
    title: 'Roadmap review | Microsoft Teams',
    focused: true,
  },
  accessibility: {
    inMeeting: true,
    meetingId: 'teams-roadmap-001',
  },
  observedAtMs: startMs,
});
assert.equal(teamsCandidate.platform, 'microsoft_teams');
assert.equal(teamsCandidate.meeting_id, 'teams-roadmap-001');

const mainWindows = normalizeNativeMeetingCandidates({
  applications: [{
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
    windows: [{ id: 'teams-main', title: 'Microsoft Teams', visible: true }],
  }, {
    name: 'Cisco Webex',
    windows: [{ id: 'webex-main', title: 'Webex', visible: true }],
  }],
});
assert.equal(mainWindows.length, 0);

const mixed = normalizeNativeMeetingCandidates({
  applications: [{
    name: 'Lark',
    windows: [{
      id: 'lark-call-1',
      title: '产品周会 - 飞书会议',
      inMeeting: true,
    }],
  }, {
    name: 'Cisco Webex',
    bundleId: 'Cisco-Systems.Spark',
    windows: [{
      id: 'webex-call-1',
      title: 'Weekly sync - Webex Meeting',
      inMeeting: true,
    }],
  }],
}, { observedAtMs: startMs + 1_000 });
assert.deepEqual(mixed.map((item) => item.platform), ['lark', 'webex']);
assert.equal(mixed[0].meeting_id, 'native-lark-lark-call-1');
assert.equal(mixed[1].meeting_id, 'native-webex-weekly-sync');

const speakerSample = normalizeNativeActiveSpeakerSample({
  application: { name: 'zoom.us', bundleId: 'us.zoom.xos' },
  window: {
    id: 'zoom-call-1',
    title: 'Daily Standup - Zoom Meeting',
    inMeeting: true,
  },
  accessibility: {
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
  },
  observedAtMs: startMs + 100,
});
assert.equal(speakerSample.meeting.platform, 'zoom');
assert.equal(speakerSample.speaker_id, 'ada');
assert.equal(speakerSample.speaking, true);

let observed = observeNativeMeetingSample(null, zoomNative(startMs, { id: 'ada', name: 'Ada', speaking: true }), {
  minStableMs: 300,
  source: 'mac_accessibility_observer',
  observedAtMs: startMs,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started']);
assert.equal(observed.session.signals[0].meeting.platform, 'zoom');

observed = observeNativeMeetingSample(observed.state, zoomNative(startMs + 350, { id: 'ada', name: 'Ada', speaking: true }), {
  minStableMs: 300,
  source: 'mac_accessibility_observer',
  observedAtMs: startMs + 350,
});
assert.deepEqual(observed.signals.map((item) => item.type), ['speaker_started']);
assert.equal(observed.speaker.signals[0].speaker_id, 'ada');
assert.equal(observed.speaker.signals[0].occurred_at_ms, startMs);

observed = observeNativeMeetingSample(observed.state, zoomNative(startMs + 600, { id: 'grace', name: 'Grace', speaking: true }), {
  switchStableMs: 500,
  source: 'mac_accessibility_observer',
  observedAtMs: startMs + 600,
});
assert.equal(observed.speaker.signals.length, 0);

observed = observeNativeMeetingSample(observed.state, zoomNative(startMs + 750, { id: 'ada', name: 'Ada', speaking: true }), {
  switchStableMs: 500,
  source: 'mac_accessibility_observer',
  observedAtMs: startMs + 750,
});
assert.equal(observed.speaker.signals.length, 0);

const observer = createNativeMeetingObserver({
  source: 'native_desktop_observer',
  speakerOptions: {
    minStableMs: 0,
    endIdleMs: 500,
  },
});
const nativeStart = observer.observe(zoomNative(startMs + 2_000, { id: 'lin', name: 'Lin', speaking: true }), {
  observedAtMs: startMs + 2_000,
});
assert.deepEqual(nativeStart.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.equal(observer.getState().sessionState.activeMeeting.platform, 'zoom');

let nativeEnd = observer.observe({
  applications: [{
    name: 'zoom.us',
    bundleId: 'us.zoom.xos',
    windows: [{ id: 'zoom-main', title: 'Zoom Workplace', visible: true }],
  }],
  observedAtMs: startMs + 3_000,
}, {
  observedAtMs: startMs + 3_000,
});
assert.deepEqual(nativeEnd.signals.map((item) => item.type), ['meeting_ended']);
nativeEnd = observer.observe({
  applications: [{
    name: 'zoom.us',
    bundleId: 'us.zoom.xos',
    windows: [{ id: 'zoom-main', title: 'Zoom Workplace', visible: true }],
  }],
  observedAtMs: startMs + 3_600,
}, {
  observedAtMs: startMs + 3_600,
});
assert.deepEqual(nativeEnd.signals.map((item) => item.type), ['speaker_ended']);

const calls = [];
const timeline = createNativeMeetingTimelineObserver({
  async startMeeting(payload) {
    calls.push({ method: 'startMeeting', payload });
    return { ok: true, payload };
  },
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
}, {
  source: 'native_desktop_observer',
  speakerOptions: { minStableMs: 0 },
  applyOptions: { speakerAsAnnotation: true },
});

const applied = await timeline.observe({
  application: { name: 'Lark' },
  window: {
    id: 'lark-call-2',
    title: '产品评审 - 飞书会议',
    inMeeting: true,
    focused: true,
  },
  accessibility: {
    activeSpeaker: { name: 'Ada', speaking: true },
  },
  observedAtMs: startMs + 5_000,
}, {
  observedAtMs: startMs + 5_000,
});
assert.deepEqual(applied.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(applied.results.map((item) => item.action), ['startMeeting', 'insertSpeakerMark']);
assert.equal(calls[0].payload.platform, 'lark');
assert.equal(calls[0].payload.meeting_id, 'native-lark-lark-call-2');
assert.equal(calls[1].input.kind, 'speaker_started');
assert.equal(calls[1].input.captured_at_ms, startMs + 5_000);

console.log('ok native meeting adapter');
