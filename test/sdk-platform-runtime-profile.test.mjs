import assert from 'node:assert/strict';
import {
  MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA,
  buildMeetingPlatformRuntimeProfile,
  buildMeetingPlatformRuntimeProfileMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-runtime-profile.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformRuntimeProfile('google-meet', {
  baseUrl,
  speakerFilter: {
    min_stable_ms: 650,
    google_meet: {
      switch_stable_ms: 1_100,
    },
  },
});

assert.equal(google.schema, MEETING_PLATFORM_RUNTIME_PROFILE_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.runtime_contract.provider_events_block_realtime, false);
assert.equal(google.runtime_contract.transcript_blocks_realtime, false);
assert.equal(google.runtime_contract.can_insert_annotation_before_provider_start_event, true);
assert.equal(google.adapter_surfaces.primary, 'browser_extension');
assert.deepEqual(google.adapter_surfaces.recommended_order, ['browser_extension', 'desktop_observer', 'provider_reconcile']);
assert.equal(google.adapter_surfaces.provider_reconcile_surface, 'google_workspace_events_pubsub');
assert.equal(google.launch_requirements.required_message_types.includes('meeting_timeline.candidate_launch_plan'), true);
assert.equal(google.launch_requirements.provider_events_block_launch, false);
assert.equal(google.evidence_thresholds.pilot.provider_records_required, false);
assert.equal(google.evidence_thresholds.production.provider_records_required, true);
assert.equal(google.fallback_policy.provider_late_or_missing, 'keep_realtime_marks_on_captured_at_ms_and_reconcile_axis_bounds_later');
assert.equal(google.axis.start.create_on, 'local_observer_active_meeting_detected');
assert.deepEqual(google.axis.start.provider_reconcile_events, ['google.workspace.meet.conference.v2.started']);
assert.deepEqual(google.axis.end.provider_reconcile_events, ['google.workspace.meet.conference.v2.ended']);
assert.equal(google.axis.end.fallbacks.includes('browser_tab_leaves_meet_url_or_meeting_dom_inactive'), true);
assert.equal(google.speaker_markers.content_policy, 'speaker_position_only_no_transcript_text_required');
assert.equal(google.speaker_markers.provider_realtime_required, false);
assert.equal(google.speaker_markers.filter.min_stable_ms, 650);
assert.equal(google.speaker_markers.filter.switch_stable_ms, 1_100);
assert.equal(google.speaker_markers.filter.google_meet, undefined);
assert.equal(google.transcript.realtime_dependency, false);
assert.equal(google.implementation_order[0], 'open_or_detect_meeting_axis_from_local_observer');

const teams = buildMeetingPlatformRuntimeProfile('teams', { baseUrl });
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.adapter_surfaces.primary, 'desktop_or_browser_observer');
assert.deepEqual(teams.adapter_surfaces.recommended_order, ['desktop_observer', 'browser_extension', 'provider_reconcile']);
assert.equal(teams.axis.end.fallbacks.includes('teams_call_stage_leaves_active_meeting'), true);
assert.equal(teams.speaker_markers.filter.switch_stable_ms, 1_200);
assert.equal(teams.provider_events.participant_events.includes('meetingCallEvents.updated:rosterUpdated'), true);

const zoom = buildMeetingPlatformRuntimeProfile('zoom', {
  baseUrl,
  env: { ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret' },
});
assert.deepEqual(zoom.axis.start.provider_reconcile_events, ['meeting.started']);
assert.deepEqual(zoom.axis.end.provider_reconcile_events, ['meeting.ended']);
assert.equal(zoom.adapter_surfaces.primary, 'native_detector');
assert.equal(zoom.fallback_policy.next_surface_on_preflight_failure, 'browser_extension');
assert.equal(zoom.provider_events.missing_env.length, 0);
assert.equal(zoom.speaker_markers.filter.min_stable_ms, 600);

const webex = buildMeetingPlatformRuntimeProfile('webex', {
  baseUrl,
  env: { WEBEX_WEBHOOK_SECRET: 'webex-secret' },
});
assert.equal(webex.provider_events.artifact_events.includes('meetingTranscripts.created'), true);
assert.equal(webex.axis.provider_reconcile.required_for_realtime, false);

const matrix = buildMeetingPlatformRuntimeProfileMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex'],
});
assert.equal(matrix.type, 'meeting_platform_runtime_profile_matrix');
assert.equal(matrix.platform_count, 4);
assert.equal(matrix.non_blocking_provider_count, 4);
assert.equal(matrix.transcript_non_blocking_count, 4);
assert.equal(matrix.speaker_marker_enabled_count, 4);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').start_create_on, 'local_observer_active_meeting_detected');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').surface_order[0], 'desktop_observer');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'webex').production_provider_records_required, true);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').provider_events_block_realtime, false);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.platformRuntimeProfile('google-meet').axis.end.provider_reconcile_events[0], 'google.workspace.meet.conference.v2.ended');
assert.equal(kit.platformRuntimeProfileMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_runtime_profile_matrix.platform_count, 1);

console.log('ok meeting platform runtime profile');
