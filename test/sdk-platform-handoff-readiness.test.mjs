import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import { capturePlatformWebhookEvent } from '../packages/meeting-timeline-sdk/adapters/platform-capture.mjs';
import {
  buildMeetingPlatformEvidencePackage,
} from '../packages/meeting-timeline-sdk/adapters/platform-evidence-package.mjs';
import {
  assertMeetingPlatformHandoffReadiness,
  assertMeetingPlatformHandoffReadinessMatrix,
  buildMeetingPlatformHandoffReadiness,
  buildMeetingPlatformHandoffReadinessMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-handoff-readiness.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const startMs = 1_784_010_000_000;
const durationMs = 120_000;
const productionEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  MICROSOFT_GRAPH_CLIENT_STATE: 'real-teams-client-state',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'real-zoom-secret-token',
  WEBEX_WEBHOOK_SECRET: 'real-webex-secret',
};

function realSnapshot(platform, state, observedAtMs) {
  return {
    ...buildMeetingAppFixtureSnapshot(platform, {
      state,
      observedAtMs,
      title: state === 'active' ? 'Handoff readiness meeting' : 'Ready to join',
      speakerId: 'real-speaker-001',
      speakerName: 'Ada Real',
      mutedId: 'real-muted-001',
      mutedName: 'Grace Real',
    }),
    id: `handoff-${platform}-${state}-${observedAtMs}`,
    source: 'chrome_extension_capture',
    fixture_state: undefined,
  };
}

function realProviderRecord(platform, signalType, capturedAtMs) {
  const body = buildPlatformFixtureEvent(platform, signalType, {
    startMs,
    durationMs,
    title: 'Handoff readiness meeting',
    participantId: 'real-user-001',
    participantName: 'Ada Real',
    speakerId: 'real-speaker-001',
    speakerName: 'Ada Real',
    googleRecordId: 'real-google-record-001',
    teamsMeetingId: 'real-teams-meeting-001',
    zoomUuid: 'real-zoom-uuid-001',
    webexMeetingId: 'real-webex-meeting-001',
    larkMeetingId: 'real-lark-meeting-001',
    transcriptId: 'real-transcript-001',
    recordingId: 'real-recording-001',
  });
  return capturePlatformWebhookEvent({
    platform,
    method: 'POST',
    url: `${baseUrl}/api/platform-events/${platform}`,
    body,
    id: `handoff-${platform}-${signalType}`,
  }, undefined, {
    capturedAtMs,
  });
}

function realInput(platform) {
  return {
    providerRecords: [
      realProviderRecord(platform, 'meeting_start', startMs),
      realProviderRecord(platform, 'participant_join', startMs + 30_000),
      realProviderRecord(platform, 'participant_left', startMs + 90_000),
      realProviderRecord(platform, 'meeting_end', startMs + durationMs),
      realProviderRecord(platform, 'transcript_ready', startMs + durationMs + 120_000),
    ],
    meetingAppRecords: [
      realSnapshot(platform, 'active', startMs),
      realSnapshot(platform, 'prejoin', startMs + durationMs),
    ],
  };
}

const emptyGoogle = buildMeetingPlatformHandoffReadiness('google-meet', {}, {
  baseUrl,
  env: productionEnv,
});
assert.equal(emptyGoogle.schema, 'meeting_platform_handoff_readiness');
assert.equal(emptyGoogle.platform, 'google_meet');
assert.equal(emptyGoogle.status, 'needs_local_observer_evidence');
assert.equal(emptyGoogle.handoff_ready, false);
assert.equal(emptyGoogle.local_observer_ready, false);
assert.equal(emptyGoogle.adapter_contract_accepted, true);
assert.equal(emptyGoogle.next_actions.includes('capture_real_meeting_app_snapshots'), true);
assert.equal(emptyGoogle.commands.validate_real_intake.includes('meeting-platform:real-intake'), true);

const googlePackage = buildMeetingPlatformEvidencePackage('google-meet', realInput('google-meet'), {
  baseUrl,
  env: productionEnv,
  includeRunbook: false,
});
const productionGoogle = buildMeetingPlatformHandoffReadiness('google-meet', {
  evidencePackage: googlePackage,
}, {
  baseUrl,
  env: productionEnv,
  target: 'production',
});
assert.equal(productionGoogle.status, 'production_ready');
assert.equal(productionGoogle.handoff_ready, true);
assert.equal(productionGoogle.pilot_ready, true);
assert.equal(productionGoogle.production_ready, true);
assert.equal(productionGoogle.local_observer_ready, true);
assert.equal(productionGoogle.provider_reconcile_ready, true);
assert.equal(productionGoogle.evidence_counts.provider_records, 5);
assert.equal(productionGoogle.evidence_counts.dom_records, 2);
assert.equal(productionGoogle.required_host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(assertMeetingPlatformHandoffReadiness('google-meet', {
  evidencePackage: googlePackage,
}, {
  baseUrl,
  env: productionEnv,
  target: 'production',
}).production_ready, true);

const matrix = buildMeetingPlatformHandoffReadinessMatrix({
  platforms: ['google-meet', 'zoom'],
  google_meet: {
    evidencePackage: googlePackage,
  },
  zoom: {
    evidencePackage: buildMeetingPlatformEvidencePackage('zoom', realInput('zoom'), {
      baseUrl,
      env: productionEnv,
      includeRunbook: false,
    }),
  },
}, {
  baseUrl,
  env: productionEnv,
  target: 'production',
});
assert.equal(matrix.schema, 'meeting_platform_handoff_readiness_matrix');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.handoff_ready_count, 2);
assert.equal(matrix.production_ready_count, 2);
assert.equal(matrix.provider_setup_needed_count, 0);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').production_ready, true);
assert.equal(assertMeetingPlatformHandoffReadinessMatrix({
  platforms: ['google-meet'],
  google_meet: {
    evidencePackage: googlePackage,
  },
}, {
  baseUrl,
  env: productionEnv,
  target: 'production',
}).handoff_ready_count, 1);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
  async insertMarks(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  env: productionEnv,
  verify: false,
});
assert.equal(kit.platformHandoffReadiness('google-meet', {
  evidencePackage: googlePackage,
}, {
  target: 'production',
}).production_ready, true);
assert.equal(kit.platformHandoffReadinessMatrix({
  platforms: ['google-meet'],
  google_meet: {
    evidencePackage: googlePackage,
  },
}, {
  target: 'production',
}).handoff_ready_count, 1);
assert.equal(kit.report({
  platforms: ['google-meet'],
}).platform_handoff_readiness_matrix.platform_count, 1);

console.log('ok meeting platform handoff readiness');
