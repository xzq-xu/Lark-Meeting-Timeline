import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import {
  buildThreePlatformLiveAcceptanceReport,
  evaluateThreePlatformLiveEvidence,
} from '../scripts/three-platform-live-acceptance-core.mjs';
import { verifyThreePlatformLiveAcceptance } from '../scripts/verify-three-platform-live-acceptance.mjs';

const startMs = 1_784_100_000_000;

const LIVE_URLS = Object.freeze({
  google_meet: 'https://meet.google.com/abc-defg-hij',
  microsoft_teams: 'https://teams.microsoft.com/v2/?meetingjoin=true',
  zoom: 'https://app.zoom.us/wc/987654321/join?pwd=test-token',
});

function liveEvidence(platform, meetingId = `${platform}-real-1`) {
  return {
    schema: 'meeting_platform_field_evidence_input',
    schema_version: 1,
    platform,
    meeting_id: meetingId,
    run_id: `${platform}:${meetingId}`,
    meeting: {
      platform,
      meeting_id: meetingId,
      start_time: new Date(startMs).toISOString(),
      end_time: new Date(startMs + 20_000).toISOString(),
      source: 'open_meeting_session',
    },
    meetingAppRecords: [
      {
        phase: 'active',
        source: 'meeting_app_extension_auto_evidence',
        captured_at_ms: startMs,
        snapshot: {
          schema: 'meeting_app_dom_capture',
          source: 'meeting_app_extension_auto_evidence',
          url: LIVE_URLS[platform],
          inMeeting: true,
          interaction: { in_call: true, can_leave: true },
          semanticSignalTypes: ['meeting_leave_available', 'active_speaker_candidate'],
          capture: { profile: platform },
        },
      },
      {
        phase: 'ended',
        source: 'meeting_app_extension_auto_evidence',
        captured_at_ms: startMs + 20_000,
        snapshot: {
          schema: 'meeting_app_dom_capture',
          source: 'meeting_app_extension_auto_evidence',
          url: LIVE_URLS[platform],
          inMeeting: false,
          interaction: { pre_join: true },
          semanticSignalTypes: ['meeting_join_available'],
          capture: { profile: platform },
        },
      },
    ],
    annotations: [{
      id: `three-platform-live-acceptance-${platform}-${startMs + 5_000}`,
      source: 'three_platform_live_acceptance',
      captured_at_ms: startMs + 5_000,
      visible_latency_ms: 45,
      timeline_error_ms: 0,
      visible_instance_count: 1,
    }],
    speaker_markers: [{
      id: `speaker-${platform}`,
      source: `${platform}_speaker`,
      kind: 'speaker_started',
      intent: 'speaker_track',
      speaker_name: 'Alex Chen',
      captured_at_ms: startMs + 2_000,
      visible_latency_ms: 50,
      timeline_error_ms: 0,
    }],
    measurements: {
      active_observer_to_axis_latency_ms: 20,
      ended_observer_to_axis_latency_ms: 30,
      previous_meeting_annotation_count_on_new_axis: 0,
    },
  };
}

const evaluations = ['google_meet', 'microsoft_teams', 'zoom'].map((platform) => (
  evaluateThreePlatformLiveEvidence(liveEvidence(platform), { platform })
));
assert.equal(evaluations.every((row) => row.accepted), true);
assert.equal(evaluations.every((row) => row.core_accepted), true);
assert.equal(evaluations.every((row) => row.speaker_accepted), true);
assert.equal(buildThreePlatformLiveAcceptanceReport(evaluations).accepted, true);

const coreOnly = liveEvidence('zoom', 'zoom-core-only');
coreOnly.speaker_markers = [];
const coreOnlyEvaluation = evaluateThreePlatformLiveEvidence(coreOnly);
assert.equal(coreOnlyEvaluation.accepted, true);
assert.equal(coreOnlyEvaluation.core_accepted, true);
assert.equal(coreOnlyEvaluation.production_ready, true);
assert.equal(coreOnlyEvaluation.speaker_accepted, false);
assert.equal(coreOnlyEvaluation.full_production_ready, false);
assert.deepEqual(coreOnlyEvaluation.failed_check_ids, []);
assert.deepEqual(coreOnlyEvaluation.warning_check_ids, [
  'stable_speaker_marker',
  'speaker_marker_visible_latency',
]);

const strictCoreOnlyEvaluation = evaluateThreePlatformLiveEvidence(coreOnly, { requireSpeaker: true });
assert.equal(strictCoreOnlyEvaluation.accepted, false);
assert.equal(strictCoreOnlyEvaluation.production_ready, true);
assert.equal(strictCoreOnlyEvaluation.full_production_ready, false);
assert.deepEqual(strictCoreOnlyEvaluation.failed_check_ids, [
  'stable_speaker_marker',
  'speaker_marker_visible_latency',
]);

const falsePositive = liveEvidence('microsoft_teams', 'teams.microsoft.com-v2');
falsePositive.meetingAppRecords[0].snapshot = {
  schema: 'meeting_app_dom_capture',
  source: 'meeting_app_extension_auto_evidence',
  url: 'https://teams.microsoft.com/v2/',
  title: 'Microsoft Teams',
  interaction: {},
  semanticSignals: [],
  semanticSignalTypes: [],
  page: { controls: [{ label: 'Retry' }, { label: 'Clear cache and retry' }] },
};
const rejectedFalsePositive = evaluateThreePlatformLiveEvidence(falsePositive);
assert.equal(rejectedFalsePositive.accepted, false);
assert.equal(rejectedFalsePositive.failed_check_ids.includes('real_active_snapshot'), true);

const untrustedDomain = liveEvidence('zoom', 'fake-zoom-domain');
untrustedDomain.meetingAppRecords.forEach((record) => {
  record.snapshot.url = 'https://example.test/wc/987654321/join';
});
const rejectedUntrustedDomain = evaluateThreePlatformLiveEvidence(untrustedDomain);
assert.equal(rejectedUntrustedDomain.accepted, false);
assert.equal(rejectedUntrustedDomain.failed_check_ids.includes('real_active_snapshot'), true);

const contaminated = liveEvidence('zoom');
contaminated.measurements.previous_meeting_annotation_count_on_new_axis = 1;
const rejectedContamination = evaluateThreePlatformLiveEvidence(contaminated);
assert.equal(rejectedContamination.accepted, false);
assert.equal(rejectedContamination.failed_check_ids.includes('cross_meeting_isolation'), true);

const tempDir = await mkdtemp(join(tmpdir(), 'three-platform-live-acceptance-'));
try {
  for (const platform of ['google_meet', 'microsoft_teams', 'zoom']) {
    await writeFile(join(tempDir, `${platform}.json`), `${JSON.stringify(liveEvidence(platform), null, 2)}\n`, 'utf8');
  }
  const report = await verifyThreePlatformLiveAcceptance({ inputDir: tempDir });
  assert.equal(report.accepted, true);
  assert.equal(report.accepted_platform_count, 3);
  assert.equal(report.production_ready, true);
  assert.equal(report.full_production_ready, true);

  await writeFile(join(tempDir, 'zoom.json'), `${JSON.stringify(coreOnly, null, 2)}\n`, 'utf8');
  const coreReport = await verifyThreePlatformLiveAcceptance({
    inputDir: tempDir,
    platforms: ['zoom'],
  });
  assert.equal(coreReport.accepted, true);
  assert.equal(coreReport.production_ready, true);
  assert.equal(coreReport.full_production_ready, false);
  assert.deepEqual(coreReport.remaining_speaker_platforms, ['zoom']);

  const strictReport = await verifyThreePlatformLiveAcceptance({
    inputDir: tempDir,
    platforms: ['zoom'],
    requireSpeaker: true,
  });
  assert.equal(strictReport.accepted, false);
  assert.equal(strictReport.production_ready, true);
  assert.equal(strictReport.full_production_ready, false);
} finally {
  await rm(tempDir, { recursive: true, force: true });
}

console.log('ok three platform live acceptance');
