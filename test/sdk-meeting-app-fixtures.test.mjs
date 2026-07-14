import assert from 'node:assert/strict';

import {
  MEETING_APP_FIXTURE_PLATFORMS,
  buildAllMeetingAppFixtureSnapshots,
  buildMeetingAppFixtureAcceptanceReport,
  buildMeetingAppFixtureSnapshot,
  diagnoseMeetingAppFixture,
  diagnoseMeetingAppFixtureLifecycle,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  buildMeetingAppFixtureTrackReadinessReport,
  diagnoseMeetingAppFixtureTrackReadiness,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks.mjs';
import { createMeetingAppTimelineRuntime } from '../packages/meeting-timeline-sdk/adapters/meeting-app-runtime.mjs';
import { normalizeMeetingAppSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-apps.mjs';

const observedAtMs = 1_783_356_000_000;

assert.deepEqual(MEETING_APP_FIXTURE_PLATFORMS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);

const allSnapshots = buildAllMeetingAppFixtureSnapshots({ observedAtMs });
assert.deepEqual(Object.keys(allSnapshots), MEETING_APP_FIXTURE_PLATFORMS);
assert.equal(allSnapshots.google_meet.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(allSnapshots.lark.page.buttons.some((item) => item.label === 'AI 视图'), true);

for (const platform of MEETING_APP_FIXTURE_PLATFORMS) {
  const snapshot = buildMeetingAppFixtureSnapshot(platform, { observedAtMs });
  const normalized = normalizeMeetingAppSnapshot(snapshot);
  assert.equal(snapshot.fixture_state, 'active');
  assert.equal(normalized.platform, platform);
  assert.equal(normalized.inMeeting, true, `${platform} fixture should be in meeting`);
  assert.equal(Boolean(normalized.meeting_id), true, `${platform} fixture should expose meeting id`);
  assert.equal(normalized.activeSpeaker.name, 'Ada Lovelace');

  const diagnosis = diagnoseMeetingAppFixture(platform, { observedAtMs });
  assert.equal(diagnosis.platform, platform);
  assert.deepEqual(diagnosis.signal_types, ['meeting_started', 'speaker_started']);
  assert.equal(diagnosis.coverage.platform_detected, true);
  assert.equal(diagnosis.coverage.meeting_started, true);
  assert.equal(diagnosis.coverage.speaker_started, true);

  const prejoinSnapshot = buildMeetingAppFixtureSnapshot(platform, { observedAtMs, state: 'prejoin' });
  const prejoin = normalizeMeetingAppSnapshot(prejoinSnapshot);
  assert.equal(prejoinSnapshot.fixture_state, 'prejoin');
  assert.equal(prejoin.platform, platform);
  assert.equal(prejoin.inMeeting, false, `${platform} prejoin fixture should be outside meeting`);

  const lifecycle = diagnoseMeetingAppFixtureLifecycle(platform, { observedAtMs });
  assert.deepEqual(lifecycle.signal_types, ['meeting_started', 'speaker_started', 'meeting_ended']);
  assert.equal(lifecycle.coverage.meeting_ended, true);
  assert.equal(lifecycle.coverage.ended_in_meeting_false, true);

  const trackReadiness = diagnoseMeetingAppFixtureTrackReadiness(platform, { observedAtMs });
  assert.equal(trackReadiness.coverage.participant_roster_snapshot, true);
  assert.equal(trackReadiness.coverage.participant_roster_count >= 2, true);
  assert.equal(trackReadiness.coverage.speaker_track_mark, true);
  assert.equal(trackReadiness.coverage.participant_track_mark, true);
  assert.equal(trackReadiness.speaker_track.marks[0].intent, 'speaker_track');
  assert.equal(trackReadiness.participant_track.marks[0].intent, 'participant_track');
}

const report = buildMeetingAppFixtureAcceptanceReport({ observedAtMs });
assert.equal(report.type, 'meeting_app_fixture_acceptance_report');
assert.equal(report.accepted, true);
assert.equal(report.platform_count, MEETING_APP_FIXTURE_PLATFORMS.length);
assert.equal(report.accepted_count, MEETING_APP_FIXTURE_PLATFORMS.length);
assert.deepEqual(report.missing, []);
assert.equal(report.coverage_by_platform.webex.active_speaker, true);
assert.equal(report.coverage_by_platform.webex.meeting_ended, true);
assert.equal(report.lifecycle_reports.length, MEETING_APP_FIXTURE_PLATFORMS.length);

const trackReadinessReport = buildMeetingAppFixtureTrackReadinessReport({ observedAtMs });
assert.equal(trackReadinessReport.schema, 'meeting_app_fixture_track_readiness_report');
assert.equal(trackReadinessReport.accepted, true);
assert.equal(trackReadinessReport.accepted_count, MEETING_APP_FIXTURE_PLATFORMS.length);
assert.deepEqual(trackReadinessReport.missing, []);
assert.equal(trackReadinessReport.coverage_by_platform.webex.participant_roster_snapshot, true);
assert.equal(trackReadinessReport.coverage_by_platform.webex.speaker_track_mark, true);
assert.equal(trackReadinessReport.coverage_by_platform.webex.participant_track_mark, true);

const calls = [];
const runtime = createMeetingAppTimelineRuntime({
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, input };
  },
  async insertMark(input, options) {
    calls.push({ method: 'insertMark', input, options });
    return { ok: true, input };
  },
}, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
});

const webexSnapshot = buildMeetingAppFixtureSnapshot('webex', { observedAtMs });
const observed = await runtime.observeMeetingApp(webexSnapshot, { observedAtMs });
assert.deepEqual(observed.signals.map((item) => item.type), ['meeting_started', 'speaker_started']);
assert.deepEqual(calls.map((item) => item.method), ['startMeeting', 'insertMark']);
assert.equal(calls[0].input.platform, 'webex');
assert.equal(calls[0].input.meeting_id, 'meet-sdk-fixture');
assert.equal(calls[1].input.kind, 'speaker_started');
assert.equal(calls[1].input.payload.speaker_name, 'Ada Lovelace');

assert.throws(
  () => buildMeetingAppFixtureSnapshot('local-detector'),
  /Unsupported meeting app fixture platform/,
);

console.log('ok meeting app fixture acceptance');
