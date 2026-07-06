import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterPortfolio,
  buildMeetingPlatformAdapterPortfolioItem,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-portfolio.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const google = buildMeetingPlatformAdapterPortfolioItem('google-meet', { baseUrl });
assert.equal(google.schema, 'meeting_platform_adapter_portfolio_item');
assert.equal(google.platform, 'google_meet');
assert.equal(google.built_in, true);
assert.equal(google.adapter_status, 'built_in_adapter_available');
assert.equal(google.recommended_first_surface, 'browser_extension');
assert.equal(google.p0_realtime_axis.local_axis_first, true);
assert.equal(google.p0_realtime_axis.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(google.p0_realtime_axis.timestamp_field, 'captured_at_ms');
assert.equal(google.p0_realtime_axis.provider_events_block_realtime, false);
assert.equal(google.p1_provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(google.p1_provider_reconcile.required_for_realtime, false);
assert.equal(google.p1_provider_reconcile.official_doc_count, 3);
assert.equal(google.p1_provider_reconcile.event_mapping_count, 6);
assert.equal(google.p2_post_meeting_backfill.realtime_dependency, false);
assert.equal(google.implementation.implementation_ready, true);
assert.equal(google.implementation.pilot_ready, true);
assert.equal(google.commands.adapter_authoring.includes('meeting-platform:adapter-authoring'), true);
assert.equal(google.commands.provider_connection.includes('meeting-platform:provider-connection'), true);
assert.equal(google.evidence_requirements.required_local_evidence.includes('realtime_annotation_sample'), true);

const acme = buildMeetingPlatformAdapterPortfolioItem('Acme Rooms', {
  baseUrl,
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
  providerTransport: 'webhook',
});
assert.equal(acme.schema, 'meeting_platform_adapter_portfolio_item');
assert.equal(acme.platform, 'acme_rooms');
assert.equal(acme.built_in, false);
assert.equal(acme.adapter_status, 'adapter_authoring_required');
assert.equal(acme.recommended_first_surface, 'browser_extension');
assert.equal(acme.p1_provider_reconcile.path, 'acme_rooms_webhooks');
assert.equal(acme.p1_provider_reconcile.official_doc_count, 0);
assert.equal(acme.implementation, undefined);
assert.equal(acme.next_actions.includes('add_platform_setup_entry'), true);

const portfolio = buildMeetingPlatformAdapterPortfolio({
  baseUrl,
  platforms: ['google-meet', 'Acme Rooms'],
});
assert.equal(portfolio.schema, 'meeting_platform_adapter_portfolio');
assert.equal(portfolio.platform_count, 2);
assert.equal(portfolio.built_in_count, 1);
assert.equal(portfolio.external_authoring_count, 1);
assert.equal(portfolio.browser_surface_ready_count, 1);
assert.equal(portfolio.provider_reconcile_count, 2);
assert.equal(portfolio.implementation_ready_count, 1);
assert.equal(portfolio.pilot_ready_count, 1);
assert.equal(portfolio.production_ready_count, 0);
assert.equal(portfolio.rows.find((row) => row.platform === 'google_meet').official_doc_count, 3);
assert.equal(portfolio.rows.find((row) => row.platform === 'acme_rooms').adapter_status, 'adapter_authoring_required');

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
assert.equal(kit.platformAdapterPortfolioItem('zoom').p1_provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformAdapterPortfolio().built_in_count, 2);
assert.equal(kit.report().platform_adapter_portfolio.platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterPortfolioItem('google-meet').built_in, true);
assert.equal(sdk.adapterPortfolioItem('zoom').p1_provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(sdk.platformAdapterPortfolio().pilot_ready_count, 2);
assert.equal(sdk.adapterPortfolio({ platforms: ['google-meet', 'Acme Rooms'] }).external_authoring_count, 1);

console.log('ok meeting platform adapter portfolio');
