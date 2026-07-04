import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdaptationPackage,
  buildMeetingPlatformAdaptationPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adaptation-package.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdaptationPackage('google-meet', { baseUrl });
assert.equal(google.schema, 'meeting_platform_adaptation_package');
assert.equal(google.platform, 'google_meet');
assert.equal(google.display_name, 'Google Meet');
assert.equal(google.readiness.sdk_wiring_ready, true);
assert.equal(google.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(google.runtime_contract.provider_events_block_realtime, false);
assert.equal(google.runtime_contract.transcript_blocks_realtime, false);
assert.equal(google.local_observer.required_for_realtime_axis, true);
assert.equal(google.local_observer.candidate_observation.runtime_event_action, 'observe_platform_candidates');
assert.deepEqual(google.extension.matches, ['https://meet.google.com/*']);
assert.equal(google.extension.permissions.includes('tabs'), true);
assert.equal(google.extension.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.candidate_observation.runtime_event_client_method, 'observePlatformCandidates');
assert.equal(google.candidate_observation.example.action, 'observe_platform_candidates');
assert.equal(google.annotation_pipeline.insert_endpoint, `${baseUrl}/api/annotations`);
assert.equal(google.annotation_pipeline.runtime_event_endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.annotation_pipeline.runtime_event.client_factory, 'createMeetingPlatformRuntimeEventClient');
assert.equal(google.annotation_pipeline.runtime_event_plan.schema, 'meeting_platform_runtime_event_plan');
assert.equal(google.annotation_pipeline.runtime_event_actions.includes('insert_annotation'), true);
assert.equal(google.annotation_pipeline.sdk_module, '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation');
assert.equal(google.runtime_event_plan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(google.runtime_event_plan.examples.insert_annotation.action, 'insert_annotation');
assert.equal(google.provider_observer.required_for_realtime, false);
assert.equal(google.provider_observer.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.provider_observer.end_events.includes('google.workspace.meet.conference.v2.ended'), true);
assert.equal(google.speaker_markers.content_policy, 'speaker_position_only_no_transcript_text_required');
assert.equal(google.speaker_markers.filter.min_stable_ms, 700);
assert.equal(google.transcript.blocks_realtime_annotation, false);
assert.equal(google.implementation.kit_methods.includes('platformAdaptationPackage'), true);
assert.equal(google.implementation.kit_methods.includes('platformRuntimeEventPlan'), true);
assert.equal(google.implementation.kit_methods.includes('insertAnnotation'), true);
assert.equal(google.commands.print_package.includes('meeting-platform:adaptation-package'), true);
assert.equal(google.commands.print_runtime_event_plan.includes('meeting-platform:runtime-event-plan'), true);
assert.equal(google.contract_acceptance.accepted, true);
assert.equal(google.contract.realtime_axis.rules.includes('use_provider_events_only_for_reconcile_and_backfill'), true);

const matrix = buildMeetingPlatformAdaptationPackageMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_adaptation_package_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.sdk_wiring_ready_count, 5);
assert.equal(matrix.browser_observer_count, 5);
assert.equal(matrix.candidate_observer_count, 5);
assert.equal(matrix.provider_observer_count, 5);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').provider_start_event_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').runtime_event_action_count, 14);
assert.equal(google.runtime_event_plan.supported_actions.includes('run_manifest'), true);
assert.equal(google.runtime_event_plan.supported_actions.includes('run_handoff_readiness'), true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').candidate_observation_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').candidate_observer_message_type, 'meeting_timeline.observe_candidates');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').candidate_observer_permission, 'tabs');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').candidate_observer_client_method, 'observePlatformCandidates');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'webex').transcript_blocks_realtime, false);
assert.equal(matrix.rows.find((row) => row.platform === 'lark').sdk_wiring_ready, true);
assert.equal(matrix.next_actions.includes('wire_host_to_platform_live_adapter'), true);

const client = {
  async startMeeting(input) {
    return { ok: true, input };
  },
  async endMeeting(input) {
    return { ok: true, input };
  },
  async insertMark(input) {
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'teams'],
});
assert.equal(kit.platformAdaptationPackage('google-meet').extension.matches[0], 'https://meet.google.com/*');
assert.equal(kit.platformAdaptationPackageMatrix().platform_count, 2);
assert.equal(kit.report().platform_adaptation_package_matrix.platform_count, 2);
assert.equal(kit.report().platform_adaptation_package_matrix.rows.every((row) => row.sdk_wiring_ready), true);

console.log('ok meeting platform adaptation package');
