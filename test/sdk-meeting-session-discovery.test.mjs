import assert from 'node:assert/strict';

import {
  createMeetingSessionDiscoveryObserver,
  createMeetingSessionTimelineDiscovery,
  detectMeetingApplication,
  normalizeMeetingSessionCandidate,
  normalizeMeetingSessionCandidates,
  selectMeetingSessionCandidate,
} from '../packages/meeting-timeline-sdk/adapters/meeting-session-discovery.mjs';

const startMs = 1_782_528_000_000;

const browserEnv = {
  windows: [{
    id: 'browser-win-1',
    focused: true,
    application: {
      name: 'Google Chrome',
      bundleId: 'com.google.Chrome',
    },
    tabs: [
      {
        id: 'mail-tab',
        active: false,
        url: 'https://mail.google.com/mail/u/0/#inbox',
        title: 'Inbox',
      },
      {
        id: 'meet-tab',
        active: true,
        audible: true,
        url: 'https://meet.google.com/abc-defg-hij',
        title: 'Design review - Google Meet',
      },
    ],
  }],
};

let candidates = normalizeMeetingSessionCandidates(browserEnv, { observedAtMs: startMs });
assert.equal(candidates.length, 1);
assert.equal(candidates[0].platform, 'google_meet');
assert.equal(candidates[0].meeting_id, 'abc-defg-hij');
assert.equal(candidates[0].meeting.title, 'Design review - Google Meet');
assert.equal(candidates[0].discovery.platform_reason, 'url');

let selected = selectMeetingSessionCandidate(browserEnv, { observedAtMs: startMs });
assert.equal(selected.detectedMeeting.platform, 'google_meet');
assert.equal(selected.detectedMeeting.meeting_id, 'abc-defg-hij');
assert.equal(selected.normalizedCandidates.length, 1);

const zoomScheme = detectMeetingApplication({
  url: 'zoommtg://zoom.us/join?confno=987654321&pwd=secret',
  title: 'Zoom Meeting',
});
assert.equal(zoomScheme.platform, 'zoom');
assert.equal(zoomScheme.meeting.meeting_id, '987654321');
assert.equal(zoomScheme.reason, 'url_scheme');

const zoomCandidate = normalizeMeetingSessionCandidate({
  id: 'zoom-window-1',
  title: 'Zoom Meeting',
  application: {
    name: 'zoom.us',
    bundleId: 'us.zoom.xos',
  },
  inMeeting: true,
  visible: true,
}, { observedAtMs: startMs + 10_000 });
assert.equal(zoomCandidate.platform, 'zoom');
assert.equal(zoomCandidate.meeting_id, 'native-zoom-zoom-window-1');
assert.equal(zoomCandidate.in_meeting, true);
assert.equal(zoomCandidate.discovery.platform_reason, 'bundle_id');

const teamsMainWindow = normalizeMeetingSessionCandidates({
  applications: [{
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
    windows: [{
      id: 'teams-main',
      title: 'Microsoft Teams',
      visible: true,
    }],
  }],
});
assert.equal(teamsMainWindow.length, 0);

candidates = normalizeMeetingSessionCandidates({
  applications: [{
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
    windows: [{
      id: 'teams-call-1',
      title: 'Design Review | Microsoft Teams',
      inMeeting: true,
      focused: true,
    }],
  }, {
    name: 'Lark',
    windows: [{
      id: 'lark-call-1',
      title: '产品周会 - 飞书会议',
      in_meeting: true,
    }],
  }, {
    name: 'Cisco Webex',
    windows: [{
      id: 'webex-call-1',
      title: 'Weekly sync - Webex Meeting',
      inMeeting: true,
    }],
  }],
}, { observedAtMs: startMs + 20_000 });
assert.deepEqual(candidates.map((item) => item.platform), ['microsoft_teams', 'lark', 'webex']);
assert.equal(candidates[0].meeting_id, 'native-microsoft_teams-design-review');
assert.equal(candidates[1].platform, 'lark');
assert.equal(candidates[2].platform, 'webex');

const observer = createMeetingSessionDiscoveryObserver({ source: 'desktop_session_discovery' });
const observedStart = observer.observeEnvironment(browserEnv, { observedAtMs: startMs + 30_000 });
assert.equal(observedStart.signals.length, 1);
assert.equal(observedStart.signals[0].type, 'meeting_started');
assert.equal(observedStart.signals[0].meeting.platform, 'google_meet');
assert.equal(observedStart.signals[0].source, 'desktop_session_discovery');

const observedRepeat = observer.observeEnvironment(browserEnv, { observedAtMs: startMs + 35_000 });
assert.equal(observedRepeat.signals.length, 0);
assert.equal(observer.getState().activeMeeting.meeting_id, 'abc-defg-hij');

const observedEnd = observer.observeEnvironment({
  windows: [{
    id: 'browser-win-1',
    focused: true,
    tabs: [{ id: 'mail-tab', active: true, url: 'https://mail.google.com', title: 'Inbox' }],
  }],
}, { observedAtMs: startMs + 40_000 });
assert.equal(observedEnd.signals.length, 1);
assert.equal(observedEnd.signals[0].type, 'meeting_ended');
assert.equal(observedEnd.signals[0].meeting.meeting_id, 'abc-defg-hij');
assert.equal(observer.getState().activeMeeting, null);

const clientCalls = [];
const timeline = createMeetingSessionTimelineDiscovery({
  async startMeeting(payload) {
    clientCalls.push({ method: 'startMeeting', payload });
    return { ok: true, payload };
  },
  async endMeeting(payload) {
    clientCalls.push({ method: 'endMeeting', payload });
    return { ok: true, payload };
  },
}, {
  source: 'native_app_detector',
});

const nativeStart = await timeline.observeEnvironment({
  applications: [{
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
    windows: [{
      id: 'teams-call-2',
      title: 'Roadmap review | Microsoft Teams',
      inMeeting: true,
      focused: true,
    }],
  }],
}, { observedAtMs: startMs + 50_000 });
assert.equal(nativeStart.signals[0].meeting.platform, 'microsoft_teams');
assert.equal(nativeStart.results[0].action, 'startMeeting');
assert.equal(clientCalls[0].payload.meeting_id, 'native-microsoft_teams-roadmap-review');

const nativeEnd = await timeline.observeEnvironment({
  applications: [{
    name: 'Microsoft Teams',
    bundleId: 'com.microsoft.teams2',
    windows: [{ id: 'teams-main', title: 'Microsoft Teams' }],
  }],
}, { observedAtMs: startMs + 60_000 });
assert.equal(nativeEnd.signals[0].type, 'meeting_ended');
assert.equal(nativeEnd.results[0].action, 'endMeeting');
assert.equal(clientCalls[1].payload.end_time_ms, startMs + 60_000);

console.log('ok meeting session discovery adapter');
