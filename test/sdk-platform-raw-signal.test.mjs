import assert from 'node:assert/strict';
import {
  buildMeetingPlatformRawSignal,
  buildMeetingPlatformRawSignalBatch,
  buildMeetingPlatformRawSignalExampleBatch,
  buildMeetingPlatformRuntimeEventsFromRawSignal,
  createMeetingAppTimelineSdk,
  normalizeMeetingPlatformRawSignalKind,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseMs = 1_782_614_400_000;

assert.equal(normalizeMeetingPlatformRawSignalKind('dom-snapshot'), 'meeting_app_snapshot');
assert.equal(normalizeMeetingPlatformRawSignalKind('insert_mark'), 'annotation');

const googleSnapshot = buildMeetingPlatformRawSignal({
  kind: 'meeting_app_snapshot',
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Weekly sync',
  observed_at_ms: baseMs,
  in_meeting: true,
  active_speaker: {
    id: 'alex',
    display_name: 'Alex',
  },
}, {
  source: 'browser_dom_observer',
});
assert.equal(googleSnapshot.platform, 'google_meet');
assert.equal(googleSnapshot.kind, 'meeting_app_snapshot');
assert.equal(googleSnapshot.runtime_event_count, 2);
assert.equal(googleSnapshot.runtime_events[0].action, 'observe_meeting_app');
assert.equal(googleSnapshot.runtime_events[0].snapshot.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(googleSnapshot.runtime_events[1].action, 'speaker_track');
assert.equal(googleSnapshot.runtime_events[1].signals[0].speaker.display_name, 'Alex');

const inferredSnapshot = buildMeetingPlatformRawSignal({
  url: 'https://meet.google.com/abc-defg-hij',
  observed_at_ms: baseMs,
  active_speaker: {
    id: 'alex',
    display_name: 'Alex',
  },
});
assert.equal(inferredSnapshot.kind, 'meeting_app_snapshot');
assert.equal(inferredSnapshot.runtime_events[0].action, 'observe_meeting_app');

const zoomCandidates = buildMeetingPlatformRawSignal({
  kind: 'platform_candidates',
  observed_at_ms: baseMs + 1_000,
  windows: [{
    id: 'zoom-window-1',
    focused: true,
    title: 'Zoom Meeting',
    tabs: [],
  }],
});
assert.equal(zoomCandidates.kind, 'platform_candidates');
assert.equal(zoomCandidates.platform, undefined);
assert.equal(zoomCandidates.runtime_events[0].action, 'observe_platform_candidates');
assert.equal(zoomCandidates.runtime_events[0].windows.length, 1);

const larkProviderEvents = buildMeetingPlatformRuntimeEventsFromRawSignal({
  kind: 'provider_event',
  platform: 'lark',
  event_id: 'evt-1',
  provider_event: {
    event_type: 'vc.meeting.meeting_started_v1',
    event_ts: baseMs,
    meeting: {
      meeting_id: 'lark-001',
      title: 'Demo meeting',
    },
  },
});
assert.equal(larkProviderEvents.length, 1);
assert.equal(larkProviderEvents[0].action, 'provider_event');
assert.equal(larkProviderEvents[0].platform, 'lark');
assert.equal(larkProviderEvents[0].provider_event.event_type, 'vc.meeting.meeting_started_v1');

const annotationSignal = buildMeetingPlatformRawSignal({
  kind: 'annotation',
  platform: 'google-meet',
  label: 'why?',
  captured_at_ms: baseMs + 30_000,
  current_meeting: {
    meeting_id: 'google_meet-sample-meeting',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
  },
});
assert.equal(annotationSignal.runtime_events[0].action, 'insert_annotation');
assert.equal(annotationSignal.runtime_events[0].annotation.label, 'why?');
assert.equal(annotationSignal.runtime_events[0].annotation.captured_at_ms, baseMs + 30_000);
assert.equal(annotationSignal.runtime_events[0].current_meeting.meeting_id, 'google_meet-sample-meeting');

const batch = buildMeetingPlatformRawSignalBatch([
  {
    kind: 'meeting_app_snapshot',
    platform: 'zoom',
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: baseMs,
  },
  {
    kind: 'speaker_track',
    platform: 'zoom',
    speaker: 'Morgan',
    occurred_at_ms: baseMs + 5_000,
  },
  {
    kind: 'annotation',
    platform: 'zoom',
    label: 'follow up',
    captured_at_ms: baseMs + 8_000,
  },
]);
assert.equal(batch.signal_count, 3);
assert.equal(batch.runtime_event_count, 3);
assert.deepEqual(batch.kinds, ['meeting_app_snapshot', 'speaker_track', 'annotation']);
assert.equal(batch.runtime_events.map((event) => event.action).join(','), 'observe_meeting_app,speaker_track,insert_annotation');

const filteredSpeakerBatch = buildMeetingPlatformRawSignalBatch([
  {
    kind: 'meeting_app_snapshot',
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    kind: 'meeting_app_snapshot',
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs + 200,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    kind: 'meeting_app_snapshot',
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly sync',
    observed_at_ms: baseMs + 700,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
], {
  filterActiveSpeakerSamples: true,
  minStableMs: 500,
});
assert.equal(filteredSpeakerBatch.signal_count, 3);
assert.equal(filteredSpeakerBatch.filtered_speaker_event_count, 1);
assert.equal(filteredSpeakerBatch.runtime_event_count, 4);
assert.equal(filteredSpeakerBatch.runtime_events.filter((event) => event.action === 'speaker_track').length, 1);
assert.equal(filteredSpeakerBatch.runtime_events.find((event) => event.action === 'speaker_track').signals[0].speaker_name, 'Alex');

const filteredFlickerBatch = buildMeetingPlatformRawSignalBatch([
  {
    kind: 'meeting_app_snapshot',
    url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    observed_at_ms: baseMs,
    active_speaker: { id: 'alex', display_name: 'Alex' },
  },
  {
    kind: 'meeting_app_snapshot',
    url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    observed_at_ms: baseMs + 100,
    active_speaker: { id: 'morgan', display_name: 'Morgan' },
  },
], {
  filterActiveSpeakerSamples: true,
  minStableMs: 500,
});
assert.equal(filteredFlickerBatch.filtered_speaker_event_count, 0);
assert.equal(filteredFlickerBatch.runtime_events.some((event) => event.action === 'speaker_track'), false);

const examples = buildMeetingPlatformRawSignalExampleBatch({
  platforms: ['google-meet', 'teams', 'zoom'],
  sample_at_ms: baseMs,
});
assert.equal(examples.type, 'meeting_platform_raw_signal_example_batch');
assert.equal(examples.platform_count, 3);
assert.equal(examples.signal_count, 9);
assert.equal(examples.platforms.includes('microsoft_teams'), true);
assert.equal(examples.runtime_events.some((event) => event.action === 'speaker_track'), true);

const sent = [];
const sdk = createMeetingAppTimelineSdk({
  baseUrl: 'https://timeline.example.test',
  runtimeEvent: true,
  fetch: async (_url, init = {}) => {
    sent.push(JSON.parse(init.body));
    return {
      ok: true,
      status: 200,
      json: async () => ({ ok: true }),
      text: async () => '{"ok":true}',
    };
  },
});
const sendResult = await sdk.sendRawSignal({
  kind: 'meeting_app_snapshot',
  platform: 'google-meet',
  url: 'https://meet.google.com/abc-defg-hij',
  observed_at_ms: baseMs,
  active_speaker: { display_name: 'Alex' },
});
assert.equal(sendResult.runtime_event_count, 2);
assert.equal(sent.length, 2);
assert.equal(sent[0].action, 'observe_meeting_app');
assert.equal(sent[1].action, 'speaker_track');

console.log('ok meeting platform raw signal');
