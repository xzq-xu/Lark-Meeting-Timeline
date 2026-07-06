import assert from 'node:assert/strict';

import {
  assertMeetingPlatformConsumerHandoff,
  buildMeetingPlatformConsumerHandoff,
} from '../packages/meeting-timeline-sdk/adapters/platform-consumer-handoff.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const handoff = buildMeetingPlatformConsumerHandoff({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

assert.equal(handoff.schema, 'meeting_platform_consumer_handoff');
assert.equal(handoff.accepted, true);
assert.equal(handoff.platform_count, 5);
assert.equal(handoff.consumer_ready_count, 5);
assert.equal(handoff.conformance_accepted_count, 5);
assert.equal(handoff.runtime_ready_count, 5);
assert.equal(handoff.adapter_route_ready_count, 5);
assert.equal(handoff.candidate_observer_count, 5);
assert.equal(handoff.speaker_track_ready_count, 5);
assert.equal(handoff.participant_track_ready_count, 5);
assert.equal(handoff.production_ready_count, 0);
assert.equal(handoff.blocking_count, 0);
assert.equal(handoff.warning_count, 5);
assert.equal(handoff.entrypoints.primary_modules.consumer_handoff, '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff');
assert.equal(handoff.entrypoints.primary_modules.app_adapter_manifest, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest');
assert.equal(handoff.entrypoints.kit_methods.includes('platformConsumerHandoff'), true);
assert.equal(handoff.entrypoints.kit_methods.includes('meetingAppAdapterManifestMatrix'), true);
assert.equal(handoff.entrypoints.host_methods.includes('consumerHandoff'), true);
assert.match(handoff.entrypoints.commands.app_adapter_manifest, /meeting-app:adapter-manifest/);
assert.equal(handoff.endpoints.consumer_handoff, '/api/meeting-platform/consumer-handoff');
assert.equal(handoff.hard_contracts.timestamp_field, 'captured_at_ms');
assert.equal(handoff.hard_contracts.provider_events_block_realtime, false);
assert.equal(handoff.hard_contracts.transcript_blocks_realtime, false);
assert.equal(handoff.boot_order[0].action, 'run_static_consumer_handoff');
assert.equal(handoff.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(handoff.rows.every((row) => row.consumer_ready === true), true);
assert.equal(handoff.rows.every((row) => row.production_ready === false), true);
assert.equal(handoff.conformance_report.accepted, true);
assert.equal(handoff.host_integration_plan.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(handoff.adaptation_package_matrix.packages, undefined);
assert.equal(handoff.handoff_readiness_matrix.reports, undefined);
assert.equal(assertMeetingPlatformConsumerHandoff(handoff).accepted, true);

assert.throws(() => assertMeetingPlatformConsumerHandoff({
  baseUrl,
  platforms: ['google-meet'],
  requireProductionReady: true,
}), /Meeting platform consumer handoff is not ready/);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
const zoomHandoff = kit.platformConsumerHandoff({ platforms: ['zoom'] });
assert.equal(zoomHandoff.platform_count, 1);
assert.equal(zoomHandoff.rows[0].platform, 'zoom');
assert.equal(kit.assertPlatformConsumerHandoff({ platforms: ['zoom'] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_consumer_handoff.platform_count, 1);

console.log('ok meeting platform consumer handoff');
