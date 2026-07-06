import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterAuthoringMatrix,
  buildMeetingPlatformAdapterAuthoringPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-authoring.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterAuthoringPlan('google-meet', { baseUrl });
assert.equal(google.schema, 'meeting_platform_adapter_authoring_plan');
assert.equal(google.platform, 'google_meet');
assert.equal(google.built_in, true);
assert.equal(google.current_sdk_status, 'built_in_static_sdk_path_available');
assert.equal(google.recommended_first_surface, 'browser_extension');
assert.equal(google.provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(google.provider_reconcile.required_for_realtime, false);
assert.equal(google.required_contracts.timestamp_field, 'captured_at_ms');
assert.equal(google.required_contracts.local_axis_first, true);
assert.equal(google.required_contracts.provider_events_block_realtime, false);
assert.equal(google.required_contracts.transcript_blocks_realtime, false);
assert.equal(google.browser_surface.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.built_in_artifacts.implementation_handoff.schema, 'meeting_platform_implementation_handoff');
assert.equal(google.built_in_artifacts.runtime_bundle_summary.runtime_ready, true);
assert.equal(google.built_in_artifacts.adapter_contract_summary.timestamp_field, 'captured_at_ms');
assert.equal(google.commands.implementation_handoff.includes('meeting-platform:implementation-handoff'), true);
assert.equal(google.authoring_steps.some((step) => step.id === 'author_runtime_preset'), true);

const custom = buildMeetingPlatformAdapterAuthoringPlan('Acme Rooms', {
  baseUrl,
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
  providerTransport: 'webhook',
});
assert.equal(custom.schema, 'meeting_platform_adapter_authoring_plan');
assert.equal(custom.platform, 'acme_rooms');
assert.equal(custom.built_in, false);
assert.equal(custom.current_sdk_status, 'external_adapter_authoring_required');
assert.equal(custom.display_name, 'Acme Rooms');
assert.equal(custom.recommended_first_surface, 'browser_extension');
assert.equal(custom.provider_reconcile.path, 'acme_rooms_webhooks');
assert.equal(custom.provider_reconcile.transport, 'webhook');
assert.equal(custom.browser_surface.matches[0], 'https://meet.acme.example/*');
assert.match(custom.normalizer_template, /normalizeAcmeRoomsEvent/);
assert.equal(custom.risks.includes('platform_not_registered_in_static_sdk'), true);
assert.equal(custom.next_actions.includes('add_event_normalizer'), true);

const matrix = buildMeetingPlatformAdapterAuthoringMatrix({
  baseUrl,
  platforms: ['google-meet', 'Acme Rooms'],
});
assert.equal(matrix.schema, 'meeting_platform_adapter_authoring_matrix');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.built_in_count, 1);
assert.equal(matrix.external_authoring_count, 1);
assert.equal(matrix.browser_surface_ready_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').browser_match_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'acme_rooms').next_action, 'add_platform_setup_entry');
assert.equal(matrix.next_actions.includes('add_browser_runtime_preset'), true);

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
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformAdapterAuthoringPlan('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformAdapterAuthoringMatrix().built_in_count, 2);
assert.equal(kit.report().platform_adapter_authoring_matrix.platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterAuthoringPlan('google-meet').built_in, true);
assert.equal(sdk.adapterAuthoringPlan('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(sdk.platformAdapterAuthoringMatrix().built_in_count, 2);
assert.equal(sdk.adapterAuthoringMatrix({ platforms: ['google-meet', 'Acme Rooms'] }).external_authoring_count, 1);

console.log('ok meeting platform adapter authoring');
