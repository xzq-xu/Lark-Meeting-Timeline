import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA,
  buildMeetingPlatformAdapterSelection,
  buildMeetingPlatformAdapterSelectionMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-selection.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const observedAtMs = 1_783_356_000_000;

function meetingAppRecordSet(platform = 'zoom') {
  return {
    schema: 'meeting_app_snapshot_record_set',
    platform,
    records: [
      {
        phase: 'active',
        captured_at_ms: observedAtMs,
        snapshot: {
          platform,
          in_meeting: true,
          url: platform === 'google_meet' ? 'https://meet.google.com/abc-defg-hij' : undefined,
        },
      },
      {
        phase: 'ended',
        captured_at_ms: observedAtMs + 60_000,
        snapshot: {
          platform,
          in_meeting: false,
        },
      },
    ],
  };
}

const google = buildMeetingPlatformAdapterSelection('google-meet', {}, { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.selection.axis_source, 'local_observer_axis');
assert.equal(google.selection.axis_surface, 'browser_extension');
assert.equal(google.selection.annotation_source, 'annotation_insert');
assert.equal(google.selection.timestamp_field, 'captured_at_ms');
assert.equal(google.selection.provider_reconcile_source, 'provider_reconcile');
assert.equal(google.selection.provider_reconcile_required_for_production, true);
assert.equal(google.selection.speaker_track_source, 'local_observer_or_detector');
assert.equal(google.selection.post_meeting_artifact_source, 'post_meeting_artifact_import');
assert.equal(google.runtime_policy.runtime_actions[0], 'validate_raw_signal');
assert.equal(google.runtime_policy.startup_order[1], 'local_observer_axis');
assert.equal(google.runtime_policy.provider_events_block_realtime, false);
assert.equal(google.runtime_policy.transcript_blocks_realtime, false);
assert.equal(google.current_evidence.realtime_axis_evidence, false);
assert.equal(google.readiness.selection_ready, true);
assert.equal(google.readiness.pilot_evidence_ready, false);
assert.equal(google.readiness.production_evidence_ready, false);
assert.equal(google.sources.find((item) => item.id === 'provider_reconcile').references.vendor, 'Google Workspace Events API');
assert.equal(google.sources.find((item) => item.id === 'post_meeting_artifact_import').realtime_dependency, false);
assert.equal(google.next_actions.includes('capture_live_local_observer_or_detector_axis_evidence'), true);

const zoomPilot = buildMeetingPlatformAdapterSelection('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, { baseUrl });
assert.equal(zoomPilot.platform, 'zoom');
assert.equal(zoomPilot.selection.axis_source, 'local_observer_axis');
assert.equal(zoomPilot.selection.axis_surface, 'native_detector');
assert.equal(zoomPilot.current_evidence.has_meeting_app_record_set, true);
assert.equal(zoomPilot.readiness.pilot_evidence_ready, true);
assert.equal(zoomPilot.readiness.production_evidence_ready, false);
assert.equal(zoomPilot.next_actions.includes('collect_provider_records_for_production_reconcile'), true);

const teamsProduction = buildMeetingPlatformAdapterSelection('teams', {
  meetingAppRecordSet: meetingAppRecordSet('microsoft_teams'),
  providerRecords: [{ type: 'meeting_started' }, { type: 'meeting_ended' }],
}, { baseUrl });
assert.equal(teamsProduction.platform, 'microsoft_teams');
assert.equal(teamsProduction.selection.axis_source, 'local_observer_axis');
assert.equal(teamsProduction.current_evidence.has_provider_records, true);
assert.equal(teamsProduction.readiness.production_evidence_ready, true);
assert.equal(teamsProduction.sources.find((item) => item.id === 'provider_reconcile').blocks_realtime_if_missing, false);

const local = buildMeetingPlatformAdapterSelection('local-detector', {
  localDetectorRecords: [{ type: 'meeting_started' }, { type: 'meeting_ended' }],
}, { baseUrl });
assert.equal(local.platform, 'local_detector');
assert.equal(local.selection.axis_source, 'host_detector_axis');
assert.equal(local.selection.axis_surface, 'host_sdk');
assert.equal(local.selection.provider_reconcile_source, undefined);
assert.equal(local.selection.provider_reconcile_required_for_production, false);
assert.equal(local.runtime_policy.runtime_actions.includes('reconcile_provider_event_when_available'), false);
assert.equal(local.readiness.pilot_evidence_ready, true);
assert.equal(local.readiness.production_evidence_ready, true);

const matrix = buildMeetingPlatformAdapterSelectionMatrix({
  platforms: ['google-meet', 'zoom', 'teams', 'local-detector'],
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
  providerRecords: [{ type: 'meeting_started' }, { type: 'meeting_ended' }],
}, { baseUrl });
assert.equal(matrix.type, 'meeting_platform_adapter_selection_matrix');
assert.equal(matrix.platform_count, 4);
assert.equal(matrix.selection_ready_count, 4);
assert.equal(matrix.pilot_evidence_ready_count, 1);
assert.equal(matrix.provider_reconcile_count, 3);
assert.equal(matrix.local_axis_selected_count, 4);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').axis_source, 'local_observer_axis');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').pilot_evidence_ready, false);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').pilot_evidence_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'local_detector').axis_source, 'host_detector_axis');
assert.equal(matrix.rows.every((row) => row.provider_events_block_realtime === false), true);
assert.equal(matrix.rows.every((row) => row.transcript_blocks_realtime === false), true);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.platformAdapterSelection('google-meet').selection.axis_source, 'local_observer_axis');
assert.equal(kit.platformAdapterSelectionMatrix({ platforms: ['webex'] }).rows[0].platform, 'webex');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_adapter_selection_matrix.platform_count, 1);

console.log('ok meeting platform adapter selection');
