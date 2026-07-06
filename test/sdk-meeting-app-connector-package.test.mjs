import assert from 'node:assert/strict';

import {
  MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA,
  assertMeetingAppTimelineConnectorPackage,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport,
  createMeetingAppTimelineSdk,
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
assert.equal(handoff.surface_matrix.length, 2);
assert.equal(handoff.surface_matrix.find((row) => row.surface === 'browser_extension').ready_count, 2);
assert.equal(handoff.extension.file_paths.includes('manifest.json'), true);
assert.equal(handoff.extension.file_paths.includes('src/content-script.entry.mjs'), true);
assert.equal(handoff.ci_gates.includes('require_captured_at_ms_for_realtime_annotations'), true);

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
