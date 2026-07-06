import assert from 'node:assert/strict';

import {
  MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA,
  MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA,
  assertMeetingAppAdapterHandoffPackage,
  assertMeetingAppAdapterHandoffPackageMatrix,
  buildMeetingAppAdapterHandoffPackage,
  buildMeetingAppAdapterHandoffPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package.mjs';
import { buildMeetingAppAdapterSpec } from '../packages/meeting-timeline-sdk/adapters/meeting-app-adapter-spec.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const wherebySpec = {
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
  host_permissions: ['https://whereby.com/*'],
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
};

const whereby = buildMeetingAppAdapterHandoffPackage(wherebySpec, {
  windowMessaging: true,
  allowedOrigins: ['https://whereby.com'],
});
assert.equal(whereby.schema, MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA);
assert.equal(whereby.accepted, true);
assert.equal(whereby.adapter_key, 'whereby');
assert.equal(whereby.adapter_spec.accepted, true);
assert.equal(whereby.runtime_config.accepted, true);
assert.equal(whereby.adapter_manifest, undefined);
assert.equal(whereby.file_paths.includes('adapter-spec.json'), true);
assert.equal(whereby.file_paths.includes('runtime-config.json'), true);
assert.equal(whereby.file_paths.includes('extension-manifest-fragment.json'), true);
assert.equal(whereby.file_paths.includes('integration-readme.md'), true);
assert.equal(whereby.contracts.timestamp_field, 'captured_at_ms');
assert.equal(whereby.contracts.provider_events_block_realtime, false);
assert.equal(whereby.validation.required_live_evidence.includes('annotation_insert_current_axis'), true);
assert.equal(
  String(whereby.files.find((file) => file.path === 'integration-readme.md').content).includes('installMeetingAppContentScriptBridge'),
  true,
);
assert.equal(assertMeetingAppAdapterHandoffPackage(whereby).adapter_key, 'whereby');

const google = buildMeetingAppAdapterHandoffPackage('google-meet');
assert.equal(google.accepted, true);
assert.equal(google.source, 'built_in_manifest');
assert.equal(google.adapter_manifest.schema, 'meeting_app_adapter_manifest');
assert.equal(google.file_paths.includes('adapter-manifest.json'), true);
assert.equal(google.extension_manifest_fragment.matches.includes('https://meet.google.com/*'), true);

const brokenSpec = buildMeetingAppAdapterSpec({
  adapter_key: 'broken_app',
  matches: ['https://broken.example/*'],
});
const broken = buildMeetingAppAdapterHandoffPackage(brokenSpec);
assert.equal(broken.accepted, false);
assert.equal(broken.issues.some((item) => item.code === 'adapter_spec_not_accepted'), true);
assert.throws(() => assertMeetingAppAdapterHandoffPackage(broken), /Meeting app adapter handoff package acceptance failed/);

const matrix = buildMeetingAppAdapterHandoffPackageMatrix({
  platforms: ['google-meet', 'zoom'],
  adapters: [wherebySpec],
});
assert.equal(matrix.schema, MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA);
assert.equal(matrix.accepted, true);
assert.equal(matrix.package_count, 3);
assert.equal(matrix.accepted_count, 3);
assert.equal(matrix.built_in_count, 2);
assert.equal(matrix.custom_count, 1);
assert.equal(matrix.runtime_config_ready_count, 3);
assert.equal(matrix.content_script_ready_count, 3);
assert.equal(matrix.live_evidence_required_count, 3);
assert.equal(assertMeetingAppAdapterHandoffPackageMatrix(matrix).accepted, true);

const kit = createMeetingPlatformTimelineKit({ baseUrl: 'http://localhost:8787' });
assert.equal(kit.meetingAppAdapterHandoffPackage(wherebySpec).adapter_key, 'whereby');
assert.equal(kit.meetingAppAdapterHandoffPackageMatrix({ platforms: ['google-meet'], adapters: [wherebySpec] }).package_count, 2);
assert.equal(kit.assertMeetingAppAdapterHandoffPackage(wherebySpec).accepted, true);
assert.equal(kit.assertMeetingAppAdapterHandoffPackageMatrix({ platforms: ['zoom'], adapters: [wherebySpec] }).accepted, true);
assert.equal(kit.report({ platforms: ['google-meet'] }).meeting_app_adapter_handoff_package_matrix.package_count, 1);

console.log('ok meeting app adapter handoff package');
