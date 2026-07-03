import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { buildMeetingAppSnapshotRecordSet } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  MEETING_PLATFORM_FIELD_EVIDENCE_BUNDLE_SCHEMA,
  MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA,
  buildMeetingPlatformFieldEvidenceBundle,
  buildMeetingPlatformFieldEvidenceMatrix,
  buildMeetingPlatformFieldCaptureMatrix,
  buildMeetingPlatformFieldCapturePlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-capture.mjs';
import { buildMeetingPlatformEvidencePackage } from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_783_702_000_000;
const endMs = startMs + 120_000;

function meetingAppRecordSet(platform) {
  return buildMeetingAppSnapshotRecordSet([
    {
      platform,
      phase: 'active',
      capturedAtMs: startMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: startMs,
        state: 'active',
      }),
    },
    {
      platform,
      phase: 'ended',
      capturedAtMs: endMs,
      snapshot: buildMeetingAppFixtureSnapshot(platform, {
        observedAtMs: endMs,
        state: 'prejoin',
      }),
    },
  ], {
    id: `${platform}-field-capture-dom`,
    createdAtMs: endMs + 1_000,
  });
}

function providerRecords(platform) {
  return [
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_start', {
      startMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: startMs,
      label: `${platform} start`,
    }),
    capturePlatformWebhookEvent(platform, buildPlatformFixtureEvent(platform, 'meeting_end', {
      startMs,
      durationMs: 120_000,
    }), {
      capturedAtMs: endMs,
      label: `${platform} end`,
    }),
  ];
}

const emptyGoogle = buildMeetingPlatformFieldCapturePlan('google-meet', { baseUrl });
assert.equal(emptyGoogle.schema, MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA);
assert.equal(emptyGoogle.platform, 'google_meet');
assert.equal(emptyGoogle.status, 'needs_local_and_provider_evidence');
assert.equal(emptyGoogle.production_ready, false);
assert.equal(emptyGoogle.local_observer.minimum_record_count, 2);
assert.equal(emptyGoogle.provider_events.minimum_record_count, 2);
assert.equal(emptyGoogle.provider_events.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(emptyGoogle.provider_events.end_events.includes('google.workspace.meet.conference.v2.ended'), true);
assert.equal(emptyGoogle.checklist.some((item) => item.id === 'insert_live_annotation_sample'), true);
assert.equal(emptyGoogle.missing_items.includes('capture_required_snapshot:active_speaker'), true);
assert.equal(emptyGoogle.missing_items.includes('provider_missing:meeting_start'), true);

const zoomDomOnlyPackage = buildMeetingPlatformEvidencePackage('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, {
  baseUrl,
  includeRunbook: false,
});
const zoomDomOnly = buildMeetingPlatformFieldCapturePlan('zoom', {
  baseUrl,
  evidencePackage: zoomDomOnlyPackage,
});
assert.equal(zoomDomOnly.status, 'pilot_ready_provider_pending');
assert.equal(zoomDomOnly.ready_for_realtime_annotations, true);
assert.equal(zoomDomOnly.production_ready, false);
assert.equal(zoomDomOnly.missing_items.includes('capture_live_dom_snapshots_for_local_observer'), false);
assert.equal(zoomDomOnly.missing_items.includes('capture_real_provider_start_end_events'), true);
assert.equal(zoomDomOnly.provider_events.security.missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);

const webexPackage = buildMeetingPlatformEvidencePackage('webex', {
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex'),
}, {
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
  includeRunbook: false,
});
const webexReady = buildMeetingPlatformFieldCapturePlan('webex', {
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
  evidencePackage: webexPackage,
});
assert.equal(webexReady.status, 'production_ready');
assert.equal(webexReady.production_ready, true);
assert.equal(webexReady.missing_items.length, 0);
assert.deepEqual(webexReady.next_actions, ['ship_with_monitoring_and_keep_collecting_regression_evidence']);

const matrix = buildMeetingPlatformFieldCaptureMatrix({
  baseUrl,
  platforms: ['google-meet', 'zoom', 'webex'],
  evidencePackage: {
    zoom: zoomDomOnlyPackage,
    webex: webexPackage,
  },
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(matrix.type, 'meeting_platform_field_capture_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.production_ready_count, 1);
assert.equal(matrix.realtime_ready_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'webex').status, 'production_ready');
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').status, 'pilot_ready_provider_pending');
assert.equal(matrix.missing_item_count > 0, true);

const zoomDomOnlyBundle = buildMeetingPlatformFieldEvidenceBundle('zoom', {
  meetingAppRecordSet: meetingAppRecordSet('zoom'),
}, { baseUrl });
assert.equal(zoomDomOnlyBundle.schema, MEETING_PLATFORM_FIELD_EVIDENCE_BUNDLE_SCHEMA);
assert.equal(zoomDomOnlyBundle.platform, 'zoom');
assert.equal(zoomDomOnlyBundle.status, 'pilot_ready_provider_pending');
assert.equal(zoomDomOnlyBundle.ready_for_realtime_annotations, true);
assert.equal(zoomDomOnlyBundle.production_ready, false);
assert.equal(zoomDomOnlyBundle.verification.passed, false);
assert.equal(zoomDomOnlyBundle.evidence_package.schema, 'meeting_platform_evidence_package');
assert.equal(zoomDomOnlyBundle.evidence_summary.meeting_app_record_count, 2);
assert.equal(zoomDomOnlyBundle.field_capture_plan.missing_items.includes('capture_real_provider_start_end_events'), true);

const webexBundle = buildMeetingPlatformFieldEvidenceBundle('webex', {
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex'),
}, {
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(webexBundle.status, 'production_ready');
assert.equal(webexBundle.production_ready, true);
assert.equal(webexBundle.verification.passed, true);
assert.equal(webexBundle.evidence_summary.provider_record_count, 2);
assert.equal(webexBundle.evidence_summary.meeting_app_record_count, 2);
assert.deepEqual(webexBundle.next_actions, ['ship_with_monitoring_and_keep_collecting_regression_evidence']);

const evidenceMatrix = buildMeetingPlatformFieldEvidenceMatrix({
  baseUrl,
  platforms: ['zoom', 'webex'],
  inputs: {
    zoom: {
      meetingAppRecordSet: meetingAppRecordSet('zoom'),
    },
    webex: {
      providerRecords: providerRecords('webex'),
      meetingAppRecordSet: meetingAppRecordSet('webex'),
    },
  },
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(evidenceMatrix.schema, 'meeting_platform_field_evidence_matrix');
assert.equal(evidenceMatrix.platform_count, 2);
assert.equal(evidenceMatrix.production_ready_count, 1);
assert.equal(evidenceMatrix.realtime_ready_count, 2);
assert.equal(evidenceMatrix.verified_count, 1);
assert.equal(evidenceMatrix.rows.find((row) => row.platform === 'webex').verified, true);

const kit = createMeetingPlatformTimelineKit({
  baseUrl,
  env: {
    WEBEX_WEBHOOK_SECRET: 'secret',
  },
});
assert.equal(kit.platformFieldCapturePlan('webex', { evidencePackage: webexPackage }).production_ready, true);
assert.equal(kit.platformFieldCaptureMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.platformFieldEvidenceBundle('webex', {
  providerRecords: providerRecords('webex'),
  meetingAppRecordSet: meetingAppRecordSet('webex'),
}).production_ready, true);
assert.equal(kit.platformFieldEvidenceMatrix({
  platforms: ['webex'],
  inputs: {
    webex: {
      providerRecords: providerRecords('webex'),
      meetingAppRecordSet: meetingAppRecordSet('webex'),
    },
  },
}).verified_count, 1);
assert.equal(kit.report({ platforms: ['webex'], evidencePackage: { webex: webexPackage } }).platform_field_capture_matrix.production_ready_count, 1);

console.log('ok meeting platform field capture');
