import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA,
  MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA,
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-contract.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterContract('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.mode, 'hybrid_browser_observer_and_provider_events');
assert.equal(google.supported_surfaces.browser_observer, true);
assert.equal(google.supported_surfaces.provider_webhook_or_event_subscription, true);
assert.equal(google.supported_surfaces.realtime_transcript_required, false);
assert.equal(google.timebase.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.timebase.provider_events_block_realtime, false);
assert.equal(google.timebase.transcript_blocks_realtime, false);
assert.equal(google.realtime_axis.start.create_on, 'local_observer_active_meeting_detected');
assert.equal(google.realtime_axis.end.create_on, 'local_observer_meeting_inactive_or_manual_stop');
assert.equal(google.realtime_axis.rules.includes('insert_annotation_by_captured_at_ms_on_the_active_axis'), true);
assert.equal(google.annotations.endpoints.insertMark, `${baseUrl}/api/annotations`);
assert.equal(google.local_observer.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.local_observer.snapshot_collector.required_snapshots.includes('active_speaker'), true);
assert.equal(google.provider_observer.required_for_realtime, false);
assert.equal(google.provider_observer.required_for_production_evidence, true);
assert.equal(google.provider_observer.events.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.provider_observer.events.end_events.includes('google.workspace.meet.conference.v2.ended'), true);
assert.equal(google.transcript.realtime_dependency, false);
assert.equal(google.evidence.required_provider_coverage.includes('meeting_start'), true);
assert.equal(google.evidence.required_local_snapshots.includes('active_speaker'), true);
assert.equal(google.implementation.imports.contract, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract');
assert.equal(google.implementation.kit_methods.includes('platformAdapterContract'), true);
assert.equal(google.readiness.missing_items.includes('provider_missing:meeting_start'), true);

const teams = buildMeetingPlatformAdapterContract('teams', { baseUrl });
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.local_observer.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(teams.provider_observer.events.participant_events.includes('meetingCallEvents.updated:rosterUpdated'), true);
assert.equal(teams.provider_observer.security.verifier, 'verifyMicrosoftGraphClientState');
assert.equal(teams.local_observer.speaker_markers.filter.min_stable_ms, 800);

const detector = buildMeetingPlatformAdapterContract('local-detector', { baseUrl });
assert.equal(detector.platform, 'local_detector');
assert.equal(detector.mode, 'provider_or_host_detector_only');
assert.equal(detector.supported_surfaces.browser_observer, false);
assert.equal(detector.supported_surfaces.provider_webhook_or_event_subscription, false);
assert.equal(detector.realtime_axis.start.create_on, 'local_detector_meeting_started');
assert.equal(detector.provider_observer.required_for_production_evidence, false);

const matrix = buildMeetingPlatformAdapterContractMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'local-detector'],
});
assert.equal(matrix.schema, MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA);
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.browser_observer_count, 2);
assert.equal(matrix.provider_observer_count, 2);
assert.deepEqual(matrix.platforms, ['google_meet', 'microsoft_teams', 'local_detector']);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').start_create_on, 'local_observer_active_meeting_detected');
assert.equal(matrix.rows.find((row) => row.platform === 'local_detector').browser_observer, false);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
  async importTranscript(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, { baseUrl, verify: false });
assert.equal(kit.platformAdapterContract('google-meet').schema, MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA);
assert.equal(kit.platformAdapterContractMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_adapter_contract_matrix.platform_count, 1);

console.log('ok meeting platform adapter contract');
