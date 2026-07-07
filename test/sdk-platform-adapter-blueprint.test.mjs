import assert from 'node:assert/strict';

import { createMeetingAppTimelineSdk } from '../packages/meeting-timeline-sdk/index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA,
  assertMeetingPlatformAdapterBlueprint,
  assertMeetingPlatformAdapterBlueprintMatrix,
  buildMeetingPlatformAdapterBlueprint,
  buildMeetingPlatformAdapterBlueprintMatrix,
  verifyMeetingPlatformAdapterBlueprint,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-blueprint.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterBlueprint('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_ADAPTER_BLUEPRINT_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.readiness.ready, true);
assert.equal(google.primary_surface, 'browser_extension');
assert.deepEqual(google.surface_order, ['browser_extension', 'desktop_observer', 'provider_reconcile']);
assert.equal(google.surfaces.browser_extension.recommended, true);
assert.equal(google.surfaces.browser_extension.priority, 1);
assert.equal(google.surfaces.browser_extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.surfaces.browser_extension.evidence.evidence_kind, 'dom_snapshot');
assert.equal(google.surfaces.native_detector.entrypoint.app_names.includes('Google Chrome'), true);
assert.equal(google.surfaces.provider_reconcile.blocks_realtime, false);
assert.equal(google.surfaces.provider_reconcile.required_for_realtime, false);
assert.equal(google.surfaces.provider_reconcile.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.surfaces.provider_reconcile.end_events.includes('google.workspace.meet.conference.v2.ended'), true);
assert.equal(google.surfaces.post_meeting_artifact.blocks_realtime, false);
assert.equal(google.realtime_axis_contract.timestamp_field, 'captured_at_ms');
assert.equal(google.realtime_axis_contract.may_insert_before_provider_start_event, true);
assert.equal(google.runtime_contract.provider_events_block_realtime, false);
assert.equal(google.runtime_contract.transcript_blocks_realtime, false);
assert.equal(google.annotation_contract.per_meeting_isolation_required, true);
assert.equal(google.acceptance_gates.realtime_pilot.includes('annotation_insert_by_captured_at_ms_before_provider_event'), true);
assert.equal(google.platform_notes.production_focus.includes('workspace_events_subscription'), true);

const googleReadiness = verifyMeetingPlatformAdapterBlueprint(google);
assert.equal(googleReadiness.ready, true);
assert.equal(googleReadiness.missing.length, 0);
assert.equal(assertMeetingPlatformAdapterBlueprint(google).platform, 'google_meet');

const blocking = verifyMeetingPlatformAdapterBlueprint({
  ...google,
  runtime_contract: {
    ...google.runtime_contract,
    provider_events_block_realtime: true,
  },
});
assert.equal(blocking.ready, false);
assert.equal(blocking.missing.includes('provider_must_not_block_realtime'), true);

const teams = buildMeetingPlatformAdapterBlueprint('teams', { baseUrl });
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.primary_surface, 'desktop_or_browser_observer');
assert.equal(teams.surfaces.native_detector.recommended, true);
assert.equal(teams.surfaces.native_detector.priority, 1);
assert.equal(teams.surfaces.browser_extension.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(teams.platform_notes.production_focus.includes('graph_change_notification_reconcile'), true);

const zoom = buildMeetingPlatformAdapterBlueprint('zoom', { baseUrl });
assert.equal(zoom.primary_surface, 'native_detector');
assert.equal(zoom.surfaces.native_detector.recommended, true);
assert.equal(zoom.surfaces.native_detector.priority, 1);
assert.equal(zoom.surfaces.native_detector.evidence.evidence_kind, 'native_window');
assert.equal(zoom.surfaces.provider_reconcile.end_events.includes('meeting.ended'), true);
assert.equal(zoom.runtime_contract.provider_events_block_realtime, false);

const matrix = buildMeetingPlatformAdapterBlueprintMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_blueprint_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.ready_count, 5);
assert.equal(matrix.provider_non_blocking_count, 5);
assert.equal(matrix.transcript_non_blocking_count, 5);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').native_recommended, true);
assert.equal(assertMeetingPlatformAdapterBlueprintMatrix(matrix).ready_count, 5);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.platformAdapterBlueprint('google-meet').primary_surface, 'browser_extension');
assert.equal(kit.platformAdapterBlueprintMatrix({ platforms: ['zoom'] }).rows[0].primary_surface, 'native_detector');
assert.equal(kit.assertPlatformAdapterBlueprint('teams').readiness.ready, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_adapter_blueprint_matrix.platform_count, 1);

const sdk = createMeetingAppTimelineSdk({ baseUrl, platforms: ['google-meet', 'zoom'] });
assert.equal(sdk.adapterBlueprint('google-meet').surfaces.provider_reconcile.blocks_realtime, false);
assert.equal(sdk.adapterBlueprintMatrix().platform_count, 2);
assert.equal(sdk.assertAdapterBlueprint('zoom').primary_surface, 'native_detector');

console.log('ok meeting platform adapter blueprint');
