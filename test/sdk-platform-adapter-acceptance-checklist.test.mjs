import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterAcceptanceChecklist,
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const staticGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'static',
});
assert.equal(staticGoogle.schema, 'meeting_platform_adapter_acceptance_checklist');
assert.equal(staticGoogle.platform, 'google_meet');
assert.equal(staticGoogle.target, 'static');
assert.equal(staticGoogle.accepted, true);
assert.equal(staticGoogle.summary.blocking_count, 0);
assert.equal(staticGoogle.checklist.find((item) => item.id === 'captured_at_ms_contract').passed, true);
assert.equal(staticGoogle.checklist.find((item) => item.id === 'real_local_observer_evidence').status, 'skip');
assert.equal(staticGoogle.next_actions.includes('capture_real_meeting_app_snapshots'), true);

const pilotGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'pilot',
});
assert.equal(pilotGoogle.accepted, false);
assert.equal(pilotGoogle.summary.failed_required_ids.includes('real_local_observer_evidence'), true);
assert.equal(pilotGoogle.summary.failed_required_ids.includes('handoff_ready'), true);
assert.equal(pilotGoogle.checklist.find((item) => item.id === 'provider_reconcile_evidence').status, 'skip');

const productionGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'production',
});
assert.equal(productionGoogle.accepted, false);
assert.equal(productionGoogle.summary.failed_required_ids.includes('production_ready'), true);
assert.equal(productionGoogle.summary.failed_required_ids.includes('provider_reconcile_evidence'), true);

const custom = buildMeetingPlatformAdapterAcceptanceChecklist('Acme Rooms', {}, {
  baseUrl,
  target: 'static',
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
});
assert.equal(custom.platform, 'acme_rooms');
assert.equal(custom.accepted, false);
assert.equal(custom.summary.failed_required_ids.includes('adapter_registered_or_authorable'), true);
assert.equal(custom.handoff_readiness, undefined);
assert.equal(custom.next_actions.includes('add_platform_setup_entry'), true);

const matrix = buildMeetingPlatformAdapterAcceptanceChecklistMatrix({
  platforms: ['google-meet', 'Acme Rooms'],
}, {
  baseUrl,
  target: 'static',
});
assert.equal(matrix.schema, 'meeting_platform_adapter_acceptance_checklist_matrix');
assert.equal(matrix.target, 'static');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.accepted_count, 1);
assert.equal(matrix.blocked_count, 1);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').accepted, true);
assert.equal(matrix.rows.find((row) => row.platform === 'acme_rooms').failed_required_ids.includes('adapter_registered_or_authorable'), true);

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
assert.equal(kit.platformAdapterAcceptanceChecklist('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(kit.platformAdapterAcceptanceChecklistMatrix({}, { target: 'static' }).accepted_count, 2);
assert.equal(kit.report().platform_adapter_acceptance_checklist_matrix.platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterAcceptanceChecklist('google-meet', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.adapterAcceptanceChecklist('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.platformAdapterAcceptanceChecklistMatrix({}, { target: 'static' }).accepted_count, 2);
assert.equal(sdk.adapterAcceptanceChecklistMatrix({ platforms: ['google-meet', 'Acme Rooms'] }, { target: 'static' }).blocked_count, 1);

console.log('ok meeting platform adapter acceptance checklist');
