import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA,
  buildMeetingPlatformAdapterRoute,
  buildMeetingPlatformAdapterRouteMatrix,
  summarizeMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-route.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterRoute('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_ADAPTER_ROUTE_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.recommended_mode, 'local_observer_first_provider_reconcile');
assert.equal(google.adapter_surfaces.primary, 'browser_extension');
assert.deepEqual(google.adapter_surfaces.recommended_order, ['browser_extension', 'desktop_observer', 'provider_reconcile']);
assert.equal(google.launch_requirements.provider_events_block_launch, false);
assert.equal(google.evidence_thresholds.production.provider_records_required, true);
assert.equal(google.routes[0].route, 'local_observer_axis');
assert.equal(google.routes[0].required_for_realtime, true);
assert.equal(google.routes[0].browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(google.routes[0].native_entrypoint.app_names.includes('Google Chrome'), true);
assert.equal(google.routes.find((route) => route.route === 'annotation_insert').timestamp_field, 'captured_at_ms');
assert.equal(google.routes.find((route) => route.route === 'annotation_insert').supported_actions.includes('insert_annotation'), true);
assert.equal(google.routes.find((route) => route.route === 'provider_reconcile').blocks_realtime_if_missing, false);
assert.equal(google.routes.find((route) => route.route === 'provider_reconcile').start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.routes.find((route) => route.route === 'provider_reconcile').references.vendor, 'Google Workspace Events API');
assert.equal(google.routes.find((route) => route.route === 'post_meeting_artifact_import').blocks_realtime_if_missing, false);
assert.equal(google.realtime_invariants.annotations_use_absolute_captured_at_ms, true);
assert.equal(google.realtime_invariants.provider_events_block_realtime, false);
assert.equal(google.realtime_invariants.transcript_blocks_realtime, false);
assert.equal(google.gates.production, 'meetingAppRecordSet_plus_providerRecords');
const googleRouteSummary = summarizeMeetingPlatformAdapterRoute(google);
assert.equal(googleRouteSummary.first_route, 'local_observer_axis');
assert.equal(googleRouteSummary.primary_surface, 'browser_extension');
assert.equal(googleRouteSummary.provider_events_block_realtime, false);
const googleRouteReadiness = verifyMeetingPlatformAdapterRouteReadiness(google);
assert.equal(googleRouteReadiness.type, 'meeting_platform_adapter_route_readiness');
assert.equal(googleRouteReadiness.ready, true);
assert.equal(googleRouteReadiness.required_axis_route, 'local_observer_axis');
assert.equal(googleRouteReadiness.missing.length, 0);
const blockingRouteReadiness = verifyMeetingPlatformAdapterRouteReadiness({
  ...google,
  realtime_invariants: {
    ...google.realtime_invariants,
    provider_events_block_realtime: true,
  },
});
assert.equal(blockingRouteReadiness.ready, false);
assert.equal(blockingRouteReadiness.missing.includes('provider_events_block_realtime'), true);

const teams = buildMeetingPlatformAdapterRoute('teams', { baseUrl });
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.adapter_surfaces.recommended_order[0], 'desktop_observer');
assert.equal(teams.entrypoints.native_detector.app_names.includes('Microsoft Teams'), true);
assert.equal(teams.entrypoints.browser_extension.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(teams.routes.find((route) => route.route === 'provider_reconcile').participant_events.includes('meetingCallEvents.updated:rosterUpdated'), true);

const zoom = buildMeetingPlatformAdapterRoute('zoom', { baseUrl });
assert.equal(zoom.adapter_surfaces.primary, 'native_detector');
assert.equal(zoom.entrypoints.native_detector.app_names.includes('Zoom Workplace'), true);
assert.equal(zoom.routes.find((route) => route.route === 'provider_reconcile').end_events.includes('meeting.ended'), true);
assert.equal(zoom.routes.find((route) => route.route === 'speaker_position_markers').required_for_realtime, false);

const local = buildMeetingPlatformAdapterRoute('local-detector', { baseUrl });
assert.equal(local.platform, 'local_detector');
assert.equal(local.recommended_mode, 'trusted_host_detector_only');
assert.equal(local.routes[0].route, 'host_detector_axis');
assert.equal(local.entrypoints.browser_extension, undefined);
assert.equal(local.routes.some((route) => route.route === 'provider_reconcile'), false);

const matrix = buildMeetingPlatformAdapterRouteMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.type, 'meeting_platform_adapter_route_matrix');
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.route_ready_count, 5);
assert.equal(matrix.local_observer_first_count, 5);
assert.equal(matrix.provider_non_blocking_count, 5);
assert.equal(matrix.transcript_non_blocking_count, 5);
assert.deepEqual(matrix.platforms, ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').first_route, 'local_observer_axis');
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
assert.equal(matrix.rows.find((row) => row.platform === 'webex').production_provider_records_required, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').route_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').provider_blocks_realtime, false);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.platformAdapterRoute('google-meet').routes[0].route, 'local_observer_axis');
assert.equal(kit.platformAdapterRouteMatrix({ platforms: ['zoom'] }).rows[0].platform, 'zoom');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_adapter_route_matrix.platform_count, 1);

console.log('ok meeting platform adapter route');
