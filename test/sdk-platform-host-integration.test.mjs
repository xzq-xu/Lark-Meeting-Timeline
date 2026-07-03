import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA,
  MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA,
  MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA,
  assertMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationPlan,
  buildMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-host-integration.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const platforms = ['google-meet', 'teams', 'zoom'];

function file(scaffold, path) {
  return scaffold.files.find((item) => item.path === path);
}

const plan = buildMeetingPlatformHostIntegrationPlan({
  baseUrl,
  platforms,
  packageName: 'timeline-host-consumer',
});

assert.equal(plan.schema, MEETING_PLATFORM_HOST_INTEGRATION_SCHEMA);
assert.deepEqual(plan.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
assert.equal(plan.runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(plan.runtime_contract.provider_events_block_realtime, false);
assert.equal(plan.endpoints.platform_events, '/api/platform-events');
assert.equal(plan.endpoints.adapter_contracts, '/api/meeting-platform/contracts');
assert.equal(plan.handoff_bundle.platform_count, 3);
assert.equal(plan.adapter_contract_matrix.platform_count, 3);
assert.equal(plan.adapter_contract_acceptance_matrix.accepted_count, 3);
assert.equal(plan.extension_install_plan.platforms.includes('google_meet'), true);
assert.equal(plan.integration_plans.google_meet.provider_events.endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(plan.next_actions.includes('run_meeting_platform_contract_acceptance_before_enabling_new_platform'), true);
assert.equal(plan.next_actions.includes('wire_host_routes_to_handleMeetingPlatformRequest'), true);

const scaffold = buildMeetingPlatformHostIntegrationScaffold({
  baseUrl,
  platforms,
  packageName: 'timeline-host-consumer',
});

assert.equal(scaffold.schema, MEETING_PLATFORM_HOST_INTEGRATION_SCAFFOLD_SCHEMA);
assert.equal(scaffold.files.length, 9);
assert.equal(file(scaffold, 'package.json').mime, 'application/json');
const manifest = JSON.parse(file(scaffold, 'package.json').content);
assert.equal(manifest.name, 'timeline-host-consumer');
assert.equal(manifest.scripts['meeting-platform:contracts'], 'node ./scripts/print-contracts.mjs');
assert.equal(manifest.scripts['meeting-platform:contract-acceptance'], 'node ./scripts/verify-contracts.mjs');
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('createMeetingPlatformTimelineKit'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformLiveAdapterSuite'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterContractMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('platformAdapterContractAcceptanceMatrix'), true);
assert.equal(file(scaffold, 'src/meeting-platform-host.mjs').content.includes('capturedAtMs'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('handleFetchRequest'), true);
assert.equal(file(scaffold, 'src/http-routes.mjs').content.includes('/api/meeting-platform/contracts'), true);
assert.equal(file(scaffold, 'scripts/print-handoff.mjs').content.includes('host.handoffBundle()'), true);
assert.equal(file(scaffold, 'scripts/print-contracts.mjs').content.includes('host.adapterContracts()'), true);
assert.equal(file(scaffold, 'scripts/verify-contracts.mjs').content.includes('adapterContractAcceptance'), true);
assert.equal(file(scaffold, 'README.md').content.includes('captured_at_ms'), true);

const acceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(scaffold);
assert.equal(acceptance.schema, MEETING_PLATFORM_HOST_INTEGRATION_ACCEPTANCE_SCHEMA);
assert.equal(acceptance.accepted, true);
assert.equal(acceptance.blocking_count, 0);
assert.equal(acceptance.required_files.includes('src/http-routes.mjs'), true);
assert.equal(acceptance.required_files.includes('scripts/verify-contracts.mjs'), true);
assert.equal(assertMeetingPlatformHostIntegrationScaffold(scaffold).accepted, true);

const broken = {
  ...scaffold,
  files: scaffold.files.filter((item) => item.path !== 'src/http-routes.mjs'),
};
const brokenAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(broken);
assert.equal(brokenAcceptance.accepted, false);
assert.equal(brokenAcceptance.issues.some((item) => item.code === 'missing_required_file'), true);
assert.throws(
  () => assertMeetingPlatformHostIntegrationScaffold(broken),
  /Meeting platform host integration scaffold failed acceptance/,
);

const missingContractGate = {
  ...scaffold,
  files: scaffold.files.filter((item) => item.path !== 'scripts/verify-contracts.mjs'),
};
const missingContractGateAcceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(missingContractGate);
assert.equal(missingContractGateAcceptance.accepted, false);
assert.equal(missingContractGateAcceptance.issues.some((item) => item.path === 'scripts/verify-contracts.mjs'), true);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  verify: false,
});

assert.equal(kit.platformHostIntegrationPlan({ platforms: ['zoom'] }).platforms[0], 'zoom');
assert.equal(kit.platformHostIntegrationScaffold({ platforms: ['zoom'] }).files.some((item) => item.path === 'src/meeting-platform-host.mjs'), true);
assert.equal(kit.platformHostIntegrationAcceptance(scaffold).accepted, true);
assert.equal(kit.assertPlatformHostIntegrationScaffold(scaffold).accepted, true);
assert.equal(kit.report({ platforms }).platform_host_integration.platforms.includes('microsoft_teams'), true);
assert.equal(kit.report({ platforms }).platform_host_integration_acceptance.accepted, true);

console.log('ok meeting platform host integration scaffold');
