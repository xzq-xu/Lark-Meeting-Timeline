import assert from 'node:assert/strict';

import {
  buildDesktopMeetingAdapterReadiness,
  createDesktopMeetingAdapterHost,
  normalizeDesktopMeetingScan,
  scanDesktopMeetingApps,
} from '../packages/meeting-timeline-sdk/adapters/desktop-meeting-host.mjs';

function teamsScan(atMs, options = {}) {
  return {
    schema: 'desktop_meeting_scan',
    os: 'darwin',
    observed_at_ms: atMs,
    accessibility_trusted: true,
    applications: [{
      name: 'Microsoft Teams',
      bundle_id: 'com.microsoft.teams2',
      pid: 51,
      frontmost: true,
      windows: [{
        title: 'Weekly product review | Microsoft Teams',
        focused: true,
        visible: true,
        controls: options.controls ?? [
          { role: 'AXButton', name: 'Leave' },
          { role: 'AXButton', name: 'Mute microphone' },
          { role: 'AXButton', name: 'Turn camera off' },
          { role: 'AXStaticText', name: 'Alex Chen is speaking' },
        ],
      }],
    }],
  };
}

const normalized = normalizeDesktopMeetingScan(teamsScan(1_784_000_000_000));
assert.equal(normalized.application_count, 1);
assert.equal(normalized.candidate_count, 1);
assert.equal(normalized.active_candidate_count, 1);
assert.equal(normalized.selected_candidate.platform, 'microsoft_teams');
assert.equal(normalized.selected_candidate.in_meeting, true);
assert.equal(normalized.selected_candidate.active_speaker.name, 'Alex Chen');
assert.equal(normalized.selected_candidate.matched_controls.leave.length, 1);

const falsePositive = normalizeDesktopMeetingScan(teamsScan(1_784_000_000_100, {
  controls: [{ role: 'AXButton', name: 'Open Teams settings' }],
}));
assert.equal(falsePositive.active_candidate_count, 0);

const readiness = buildDesktopMeetingAdapterReadiness(normalized);
assert.equal(readiness.ready, true);
assert.equal(readiness.target_application_running, true);
assert.equal(readiness.active_meeting_detected, true);

const permissionMissing = buildDesktopMeetingAdapterReadiness({
  ...teamsScan(1_784_000_000_200),
  accessibility_trusted: false,
});
assert.equal(permissionMissing.ready, false);
assert.equal(permissionMissing.issues[0].code, 'accessibility_permission_required');

const calls = { starts: [], ends: [], marks: [] };
const events = [];
const host = createDesktopMeetingAdapterHost({
  platforms: ['teams', 'zoom'],
  startStableSamples: 2,
  endMissingSamples: 2,
  speakerStableSamples: 2,
  client: {
    async startMeeting(value) { calls.starts.push(value); return { ok: true }; },
    async endMeeting(value) { calls.ends.push(value); return { ok: true }; },
    async insertMark(value) { calls.marks.push(value); return { ok: true }; },
  },
  onEvent(event) { events.push(event); },
});

const tick1 = await host.tick({ scan: teamsScan(1_784_000_001_000) });
assert.equal(tick1.events.length, 0);
const tick2 = await host.tick({ scan: teamsScan(1_784_000_001_750) });
assert.equal(tick2.events[0].type, 'meeting_started');
assert.equal(calls.starts.length, 1);
assert.equal(calls.starts[0].platform, 'microsoft_teams');
assert.equal(calls.starts[0].observer_surface, 'native_detector');
assert.equal(calls.starts[0].meeting_app_record.phase, 'active');

await host.tick({ scan: teamsScan(1_784_000_002_500) });
const speakerTick = await host.tick({ scan: teamsScan(1_784_000_003_250) });
assert.equal(speakerTick.events[0].type, 'speaker_started');
assert.equal(calls.marks.length, 1);
assert.equal(calls.marks[0].label, 'Speaker: Alex Chen');
assert.equal(calls.marks[0].kind, 'speaker_started');
assert.equal(calls.marks[0].intent, 'speaker_track');
assert.equal(calls.marks[0].speaker_name, 'Alex Chen');
assert.equal(calls.marks[0].captured_at_ms, 1_784_000_002_500);

const emptyScan = (atMs) => ({
  schema: 'desktop_meeting_scan',
  os: 'darwin',
  observed_at_ms: atMs,
  accessibility_trusted: true,
  applications: [],
});
const missing1 = await host.tick({ scan: emptyScan(1_784_000_004_000) });
assert.equal(missing1.events.length, 0);
const missing2 = await host.tick({ scan: emptyScan(1_784_000_004_750) });
assert.equal(missing2.events[0].type, 'meeting_ended');
assert.equal(calls.ends.length, 1);
assert.equal(calls.ends[0].meeting_app_record.phase, 'ended');
assert.equal(host.getState().active_meeting, null);
assert.deepEqual(events.map((event) => event.type), ['meeting_started', 'speaker_started', 'meeting_ended']);

if (process.platform === 'darwin') {
  const realScan = await scanDesktopMeetingApps({ timeoutMs: 5000 });
  assert.equal(realScan.schema, 'desktop_meeting_scan');
  assert.equal(realScan.supported, true);
  assert.equal(Array.isArray(realScan.applications), true);
}

console.log('ok desktop meeting host');
