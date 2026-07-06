import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_MANIFEST_SCHEMA,
  MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA,
  assertMeetingAppAdapterManifest,
  assertMeetingAppAdapterManifestMatrix,
  buildMeetingAppAdapterManifest,
  buildMeetingAppAdapterManifestMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'http://localhost:8787';

const google = buildMeetingAppAdapterManifest('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_APP_ADAPTER_MANIFEST_SCHEMA);
assert.equal(google.accepted, true);
assert.equal(google.platform, 'google_meet');
assert.equal(google.surface, 'browser_extension_or_webview');
assert.equal(google.extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.extension.permissions.includes('tabs'), true);
assert.equal(google.extension.permissions.includes('storage'), true);
assert.equal(google.extension.message_types.observe_candidates, 'meeting_timeline.observe_candidates');
assert.equal(google.runtime.observe_mutations, true);
assert.equal(google.capture.selector_counts.participant > 0, true);
assert.equal(google.capture.selector_counts.control > 0, true);
assert.equal(google.contracts.timestamp_field, 'captured_at_ms');
assert.equal(google.contracts.provider_events_block_realtime, false);
assert.equal(google.contracts.transcript_blocks_realtime, false);
assert.equal(google.readiness.requires_live_snapshot_before_production, true);
assert.match(google.commands.print_manifest, /meeting-app:adapter-manifest/);
assert.equal(assertMeetingAppAdapterManifest(google).platform, 'google_meet');
assert.equal(assertMeetingAppAdapterManifest('zoom', { baseUrl }).platform, 'zoom');

const matrix = buildMeetingAppAdapterManifestMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA);
assert.equal(matrix.accepted, true);
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.accepted_count, 5);
assert.equal(matrix.mutation_observer_count, 5);
assert.equal(matrix.candidate_observer_count, 5);
assert.equal(matrix.participant_selector_ready_count, 5);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').accepted, true);
assert.equal(assertMeetingAppAdapterManifestMatrix(matrix).accepted_count, 5);

const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.meetingAppAdapterManifest('webex').platform, 'webex');
assert.equal(kit.meetingAppAdapterManifestMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.assertMeetingAppAdapterManifest('lark').accepted, true);
assert.equal(kit.assertMeetingAppAdapterManifestMatrix({ platforms: ['zoom'] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_manifest_matrix.platform_count, 1);

assert.throws(() => buildMeetingAppAdapterManifest('local-detector'), /meeting-app adapter manifest requires a meeting app platform/);

console.log('ok meeting app adapter manifest');
