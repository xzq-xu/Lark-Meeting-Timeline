import assert from 'node:assert/strict';

import {
  MEETING_APP_SNAPSHOT_RECORD_SCHEMA,
  MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA,
  buildMeetingAppLaunchGateInputFromRecords,
  buildMeetingAppSnapshotRecord,
  buildMeetingAppSnapshotRecordSet,
  createMeetingAppSnapshotRecorder,
  meetingAppSnapshotRecords,
  meetingAppSnapshotsFromRecords,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { buildMeetingAppLaunchGate } from '../packages/meeting-timeline-sdk/adapters/meeting-app-gate.mjs';

const startMs = 1_783_356_000_000;

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

function queryNodes(nodes, selector) {
  const text = String(selector);
  if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
  if (text.includes('data-participant-id')) return nodes.filter((item) => item.attributes['data-participant-id']);
  if (text.includes('speaking')) return nodes.filter((item) => /speaking/i.test(item.attributes['aria-label'] ?? ''));
  if (text.includes('aria-label*="Leave call"')) return nodes.filter((item) => /Leave call/i.test(item.attributes['aria-label'] ?? ''));
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

const activeDocument = fakeDocument({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design Review - Google Meet',
  nodes: [
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': 'ada',
      'aria-label': 'Ada Lovelace is speaking',
      'data-audio-level': '0.81',
    }),
  ],
});

const endedSnapshot = {
  platform: 'google_meet',
  provider: 'google_meet',
  url: 'https://meet.google.com/abc-defg-hij',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  meeting_id: 'abc-defg-hij',
  observedAtMs: startMs + 1_000,
  title: 'Ready to join - Google Meet',
  page: {
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Ready to join - Google Meet',
    buttons: [{ label: 'Join now', ariaLabel: 'Join now' }],
    controls: [{ label: 'Join now', ariaLabel: 'Join now' }],
    participants: [],
    tiles: [],
  },
};

const activeRecord = buildMeetingAppSnapshotRecord({ document: activeDocument }, {
  capturedAtMs: startMs,
  phase: 'active',
  label: 'live-active',
  captureProfile: 'google_meet',
});
assert.equal(activeRecord.schema, MEETING_APP_SNAPSHOT_RECORD_SCHEMA);
assert.equal(activeRecord.platform, 'google_meet');
assert.equal(activeRecord.phase, 'active');
assert.equal(activeRecord.snapshot.capture.profile, 'google_meet');
assert.equal(activeRecord.snapshot.page.tiles[0].id, 'ada');
assert.equal(activeRecord.meeting_id, 'abc-defg-hij');

const endedRecord = buildMeetingAppSnapshotRecord({ snapshot: endedSnapshot }, {
  capturedAtMs: startMs + 1_000,
  phase: 'ended',
  label: 'live-ended',
});
assert.equal(endedRecord.phase, 'ended');
assert.equal(endedRecord.platform, 'google_meet');

const recordSet = buildMeetingAppSnapshotRecordSet([activeRecord, endedRecord], {
  id: 'google-meet-record-set',
  createdAtMs: startMs + 2_000,
});
assert.equal(recordSet.schema, MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA);
assert.equal(recordSet.record_count, 2);
assert.deepEqual(recordSet.platforms, ['google_meet']);
assert.equal(meetingAppSnapshotRecords(recordSet).length, 2);
assert.equal(meetingAppSnapshotsFromRecords(recordSet).length, 2);

const gateInput = buildMeetingAppLaunchGateInputFromRecords(recordSet);
assert.equal(gateInput.snapshots.google_meet.length, 2);
assert.equal(gateInput.snapshot_records[0].label, 'live-active');

const gateFromRecords = buildMeetingAppLaunchGate('google-meet', {
  recordSet,
});
assert.equal(gateFromRecords.evidence_level, 'captured_dom');
assert.equal(gateFromRecords.production_ready, true);
assert.equal(gateFromRecords.coverage.meeting_started, true);
assert.equal(gateFromRecords.coverage.speaker_started, true);
assert.equal(gateFromRecords.coverage.meeting_ended, true);

const recorder = createMeetingAppSnapshotRecorder({
  captureProfile: 'google_meet',
});
recorder.capture({ document: activeDocument }, {
  capturedAtMs: startMs,
  phase: 'active',
  label: 'capture-active',
});
recorder.add(endedRecord);
assert.equal(recorder.getState().record_count, 2);
assert.deepEqual(recorder.getState().platforms, ['google_meet']);
assert.equal(recorder.findByLabel('capture-active').length, 1);
assert.equal(recorder.gateInput().snapshots.google_meet.length, 2);
assert.equal(recorder.exportRecords({ id: 'recorder-export' }).id, 'recorder-export');
assert.deepEqual(recorder.reset(), { removed: 2 });
assert.equal(recorder.getState().record_count, 0);

console.log('ok meeting app snapshot recorder');
