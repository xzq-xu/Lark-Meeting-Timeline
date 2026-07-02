import assert from 'node:assert/strict';

import {
  createMeetingAppObserver,
  detectMeetingAppPreset,
  normalizeMeetingAppSnapshot,
  normalizeMeetingAppSnapshots,
  observeMeetingAppSample,
} from '../packages/meeting-timeline-sdk/adapters/meeting-apps.mjs';
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
