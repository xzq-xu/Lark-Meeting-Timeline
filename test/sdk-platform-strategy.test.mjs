import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import {
  MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA,
  buildAllMeetingPlatformAdaptationStrategies,
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-strategy.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_356_000_000;

function meetingAppRecordSet(platform) {
  const key = platform.replaceAll('-', '_');
  return buildMeetingAppSnapshotRecordSet([
    {
      platform: key,
      phase: 'active',
      capturedAtMs: observedAtMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs,
        state: 'active',
      }),
    },
    {
      platform: key,
      phase: 'ended',
      capturedAtMs: observedAtMs + 60_000,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: observedAtMs + 60_000,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${key}-strategy-record-set`,
    createdAtMs: observedAtMs + 61_000,
  });
}

const googleStrategy = buildMeetingPlatformAdaptationStrategy('google-meet', {
  baseUrl,
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  },
});
assert.equal(googleStrategy.schema, MEETING_PLATFORM_ADAPTATION_STRATEGY_SCHEMA);
assert.equal(googleStrategy.platform, 'google_meet');
assert.equal(googleStrategy.realtime_axis.primary_source, 'local_observer');
assert.equal(googleStrategy.realtime_axis.provider_events_block_realtime, false);
assert.equal(googleStrategy.realtime_axis.transcript_blocks_realtime, false);
assert.equal(googleStrategy.provider_events.role, 'non_blocking_reconcile_and_backfill');
assert.equal(googleStrategy.provider_events.required_for_production, true);
assert.equal(googleStrategy.provider_events.required_for_pilot, false);
assert.equal(googleStrategy.provider_events.event_types.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(googleStrategy.speaker_activity.realtime_primary, 'local_observer_or_detector');
assert.equal(googleStrategy.post_meeting_transcript.realtime_dependency, false);
assert.equal(googleStrategy.evidence_contract.production.requires.includes('meetingAppRecordSet'), true);
assert.equal(googleStrategy.evidence_contract.production.requires.includes('providerRecords'), true);

const zoomPilot = buildMeetingPlatformAdaptationStrategy('zoom', {
  baseUrl,
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
});
assert.equal(zoomPilot.rollout_status, 'realtime_ready_provider_pending');
assert.equal(zoomPilot.ready_for_realtime_annotations, true);
assert.equal(zoomPilot.production_ready, false);
assert.equal(zoomPilot.recommendation, 'enable_local_observer_pilot_collect_provider_evidence');
assert.equal(zoomPilot.provider_events.latency.realtime_blocking, false);
assert.equal(zoomPilot.evidence_contract.pilot.current_passed, true);
assert.equal(zoomPilot.evidence_contract.production.current_passed, false);

const localDetector = buildMeetingPlatformAdaptationStrategy('local-detector', { baseUrl });
assert.equal(localDetector.platform, 'local_detector');
assert.equal(localDetector.realtime_axis.primary_source, 'local_detector');
assert.equal(localDetector.provider_events, undefined);
assert.equal(localDetector.local_observer.role, 'primary_axis_source');
assert.equal(localDetector.evidence_contract.production.requires.includes('local_detector_start_end_records'), true);

const strategies = buildAllMeetingPlatformAdaptationStrategies({
  baseUrl,
  platforms: ['google-meet', 'teams', 'webex'],
});
assert.deepEqual(strategies.map((item) => item.platform), ['google_meet', 'microsoft_teams', 'webex']);
assert.equal(strategies.every((item) => item.realtime_axis.transcript_blocks_realtime === false), true);

const matrix = buildMeetingPlatformAdaptationStrategyMatrix({
  baseUrl,
  platforms: ['google-meet', 'zoom', 'local-detector'],
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
});
assert.equal(matrix.type, 'meeting_platform_adaptation_strategy_matrix');
assert.equal(matrix.strategy_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').ready_for_realtime_annotations, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').provider_blocks_realtime, false);
assert.equal(matrix.rows.find((row) => row.platform === 'local_detector').primary_axis_source, 'local_detector');

const kit = createMeetingPlatformTimelineKit({ baseUrl, verify: false });
assert.equal(kit.platformAdaptationStrategy('google-meet').platform, 'google_meet');
assert.equal(kit.allPlatformAdaptationStrategies({ platforms: ['zoom'] }).length, 1);
assert.equal(kit.platformAdaptationStrategyMatrix({ platforms: ['webex'] }).rows[0].platform, 'webex');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_adaptation_strategy.type, 'meeting_platform_adaptation_strategy_matrix');

console.log('ok meeting platform adaptation strategy');
