import assert from 'node:assert/strict';

import {
  MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA,
  assertMeetingAppTimelineConnectorPackage,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport,
  createMeetingAppTimelineSdk,
  createMeetingAppTimelineConnectorRuntimeClient,
  stripMeetingAppTimelineConnectorPackageFileContents,
} from '../packages/meeting-timeline-sdk/index.mjs';
import {
  buildMeetingAppTimelineConnectorPackageAcceptanceReport as buildAcceptanceReportFromSubpath,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-connector-package.mjs';

const sdk = createMeetingAppTimelineSdk({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'zoom'],
});
const connectorPackage = sdk.connectorPackage({
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});

const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage);
assert.equal(acceptance.schema, 'meeting_app_timeline_connector_package_acceptance_report');
assert.equal(acceptance.accepted, true);
assert.equal(acceptance.target, 'realtime');
assert.deepEqual(acceptance.platforms, ['google_meet', 'zoom']);
assert.deepEqual(acceptance.surfaces, ['browser_extension', 'native_detector']);
assert.equal(acceptance.require_tracks, true);
assert.equal(acceptance.issue_count, 0);
assert.equal(acceptance.missing_runtime_actions.length, 0);
assert.equal(acceptance.surface_reports.length, 2);
assert.equal(acceptance.surface_reports.every((row) => row.accepted), true);
assert.equal(acceptance.surface_reports.find((row) => row.surface === 'browser_extension').track_enabled_count, 2);
assert.equal(acceptance.extension.required, true);
assert.equal(acceptance.extension.accepted, true);
assert.equal(acceptance.extension.manifest_version, 3);

assert.equal(buildAcceptanceReportFromSubpath(connectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorPackage(connectorPackage), connectorPackage);

const handoff = buildMeetingAppTimelineConnectorHandoff(connectorPackage);
assert.equal(handoff.schema, MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA);
assert.equal(handoff.accepted, true);
assert.equal(handoff.runtime_event_endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(handoff.timestamp_field, 'captured_at_ms');
assert.equal(handoff.provider_events_block_realtime, false);
assert.equal(handoff.transcript_blocks_realtime, false);
assert.equal(handoff.adapter_blueprints.ready_count, 2);
assert.equal(handoff.adapter_blueprints.platform_count, 2);
assert.equal(handoff.adapter_blueprints.sdk_method, 'sdk.platformAdapterBlueprint(platform)');
assert.equal(handoff.adapter_blueprints.rows.find((row) => row.platform === 'google_meet').primary_surface, 'browser_extension');
assert.equal(handoff.startup_plans.realtime_startup_ready_count, 2);
assert.equal(handoff.startup_plans.platform_count, 2);
assert.equal(handoff.startup_plans.sdk_method, 'sdk.platformAdapterStartupPlan(input)');
assert.equal(handoff.startup_plans.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(handoff.startup_plans.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(handoff.surface_matrix.length, 2);
assert.equal(handoff.surface_matrix.find((row) => row.surface === 'browser_extension').ready_count, 2);
assert.equal(handoff.extension.file_paths.includes('manifest.json'), true);
assert.equal(handoff.extension.file_paths.includes('src/content-script.entry.mjs'), true);
assert.equal(handoff.ci_gates.includes('require_captured_at_ms_for_realtime_annotations'), true);
assert.equal(connectorPackage.adapter_blueprints.ready_count, 2);
assert.equal(connectorPackage.adapter_blueprints.matrix.schema, 'meeting_platform_adapter_blueprint_matrix');
assert.equal(connectorPackage.adapter_blueprints.matrix.blueprints, undefined);
assert.equal(connectorPackage.startup_plans.realtime_startup_ready_count, 2);
assert.equal(connectorPackage.startup_plans.matrix.schema, 'meeting_platform_adapter_startup_plan_matrix');
assert.equal(connectorPackage.startup_plans.matrix.plans, undefined);
assert.equal(connectorPackage.entrypoints.some((entry) => entry.id === 'adapter-blueprint'), true);
assert.equal(connectorPackage.entrypoints.some((entry) => entry.id === 'startup-plan'), true);
assert.equal(connectorPackage.contracts.adapter_blueprint_required_before_host_wiring, true);
assert.equal(connectorPackage.contracts.startup_plan_required_before_runtime_install, true);

let capturedRuntimeRequest = null;
const connectorRuntimeClient = createMeetingAppTimelineConnectorRuntimeClient(connectorPackage, {
  now: () => 1_782_614_401_000,
  fetch: async (url, init) => {
    capturedRuntimeRequest = {
      url,
      method: init.method,
      headers: init.headers,
      body: JSON.parse(init.body),
    };
    return new Response(JSON.stringify({
      ok: true,
      accepted: capturedRuntimeRequest.body,
    }), {
      status: 201,
      headers: { 'content-type': 'application/json' },
    });
  },
});
assert.equal(connectorRuntimeClient.schema, 'meeting_app_timeline_connector_runtime_client');
assert.equal(connectorRuntimeClient.endpoint, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(connectorRuntimeClient.supports('insert_annotation', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('provider_event', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('adapter_blueprint', 'google-meet'), true);
assert.equal(connectorRuntimeClient.supports('missing_action', 'google-meet'), false);
assert.equal(connectorRuntimeClient.supported_actions_by_platform.google_meet.includes('speaker_track'), true);
const builtInsertEvent = connectorRuntimeClient.buildEvent({
  action: 'insert_annotation',
  platform: 'zoom',
  annotation: {
    id: 'note-1',
    captured_at_ms: 1_782_614_402_000,
  },
});
assert.equal(builtInsertEvent.schema, 'meeting_platform_runtime_event');
assert.equal(builtInsertEvent.action, 'insert_annotation');
assert.equal(builtInsertEvent.platform, 'zoom');

const insertResult = await connectorRuntimeClient.insertAnnotation('zoom', {
  id: 'note-2',
  label: 'why?',
  captured_at_ms: 1_782_614_403_000,
});
assert.equal(capturedRuntimeRequest.url, 'https://timeline.example.com/api/meeting-platform/runtime-events');
assert.equal(capturedRuntimeRequest.method, 'POST');
assert.equal(capturedRuntimeRequest.body.action, 'insert_annotation');
assert.equal(capturedRuntimeRequest.body.platform, 'zoom');
assert.equal(insertResult.accepted.annotation.id, 'note-2');

await connectorRuntimeClient.observePlatformCandidates({
  tabs: [{
    active: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
  }],
});
assert.equal(capturedRuntimeRequest.body.action, 'observe_platform_candidates');
assert.equal(capturedRuntimeRequest.body.platform, undefined);
assert.equal(capturedRuntimeRequest.body.tabs[0].url, 'https://meet.google.com/abc-defg-hij');
await connectorRuntimeClient.runManifest({ platforms: ['google-meet'], target: 'realtime' });
assert.equal(capturedRuntimeRequest.body.action, 'run_manifest');

await connectorRuntimeClient.adapterBlueprint('google-meet');
assert.equal(capturedRuntimeRequest.body.action, 'adapter_blueprint');
assert.equal(capturedRuntimeRequest.body.platform, 'google_meet');

await connectorRuntimeClient.adapterBlueprints({ platforms: ['google-meet', 'zoom'] });
assert.equal(capturedRuntimeRequest.body.action, 'adapter_blueprints');
assert.deepEqual(capturedRuntimeRequest.body.platforms, ['google-meet', 'zoom']);

const stripped = stripMeetingAppTimelineConnectorPackageFileContents(connectorPackage);
assert.equal(stripped.extension.scaffold.files.some((file) => 'content' in file), false);
assert.equal(connectorPackage.extension.scaffold.files.some((file) => 'content' in file), true);

const missingRuntimeActionPackage = {
  ...connectorPackage,
  runtime_events: {
    ...connectorPackage.runtime_events,
    plan_matrix: {
      ...connectorPackage.runtime_events.plan_matrix,
      rows: connectorPackage.runtime_events.plan_matrix.rows.filter((row) => row.action !== 'insert_annotation'),
    },
  },
};
const missingRuntimeActionAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(
  missingRuntimeActionPackage,
);
assert.equal(missingRuntimeActionAcceptance.accepted, false);
assert.equal(
  missingRuntimeActionAcceptance.missing_runtime_actions.some((item) => item.action === 'insert_annotation'),
  true,
);
assert.throws(
  () => assertMeetingAppTimelineConnectorPackage(missingRuntimeActionPackage),
  /Meeting app timeline connector package is not accepted/,
);
assert.throws(
  () => createMeetingAppTimelineConnectorRuntimeClient(missingRuntimeActionPackage, {
    fetch: async () => new Response('{}'),
  }),
  /not accepted for runtime client/,
);

const missingBlueprintPackage = {
  ...connectorPackage,
  adapter_blueprints: undefined,
};
const missingBlueprintAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(missingBlueprintPackage);
assert.equal(missingBlueprintAcceptance.accepted, false);
assert.equal(missingBlueprintAcceptance.issues.some((issue) => issue.code === 'missing_adapter_blueprint_matrix'), true);
const missingStartupPackage = {
  ...connectorPackage,
  startup_plans: undefined,
};
const missingStartupAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(missingStartupPackage);
assert.equal(missingStartupAcceptance.accepted, false);
assert.equal(missingStartupAcceptance.issues.some((issue) => issue.code === 'missing_startup_plan_matrix'), true);
const unsafeRuntimeClient = createMeetingAppTimelineConnectorRuntimeClient(missingRuntimeActionPackage, {
  assertPackage: false,
  fetch: async () => new Response('{}'),
});
assert.throws(
  () => unsafeRuntimeClient.buildEvent({
    action: 'insert_annotation',
    platform: 'zoom',
    annotation: { id: 'blocked', captured_at_ms: 1_782_614_404_000 },
  }),
  /Runtime action insert_annotation is not present in connector package/,
);

const blockingProviderPackage = {
  ...connectorPackage,
  contracts: {
    ...connectorPackage.contracts,
    provider_events_block_realtime: true,
  },
};
const blockingProviderAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(blockingProviderPackage);
assert.equal(blockingProviderAcceptance.accepted, false);
assert.equal(blockingProviderAcceptance.issues.some((issue) => issue.code === 'provider_events_block_realtime'), true);

const productionAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage, {
  target: 'production',
});
assert.equal(productionAcceptance.accepted, false);
assert.equal(
  productionAcceptance.issues.some((issue) => issue.code === 'production_requires_live_snapshot_evidence'),
  true,
);

const productionEvidenceAcceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage, {
  target: 'production',
  liveEvidenceAccepted: true,
});
assert.equal(productionEvidenceAcceptance.accepted, true);

console.log('ok meeting app connector package adapter');
