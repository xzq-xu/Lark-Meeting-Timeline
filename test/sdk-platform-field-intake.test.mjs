import assert from 'node:assert/strict';

import {
  buildMeetingPlatformFieldIntakeMatrix,
  buildMeetingPlatformFieldIntakePlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-intake.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';
const productionEnv = {
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'meet-events@example.iam.gserviceaccount.com',
  MICROSOFT_GRAPH_CLIENT_STATE: 'real-teams-client-state',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'real-zoom-secret-token',
  WEBEX_WEBHOOK_SECRET: 'real-webex-secret',
};

const googlePlan = buildMeetingPlatformFieldIntakePlan('google-meet', {
  baseUrl,
  evidenceDir: 'tmp/evidence',
  env: productionEnv,
});
assert.equal(googlePlan.schema, 'meeting_platform_field_intake_plan');
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.base_url, baseUrl);
assert.equal(googlePlan.provider_endpoint, `${baseUrl}/api/platform-events/google-meet`);
assert.equal(googlePlan.provider_connection.security.missing_env.length, 0);
assert.equal(googlePlan.provider_connection.realtime_annotation_policy.provider_events_block_realtime, false);
assert.equal(googlePlan.acceptance.timestamp_field, 'captured_at_ms');
assert.equal(googlePlan.acceptance.required_provider_coverage.includes('meeting_start'), true);
assert.equal(googlePlan.acceptance.required_provider_coverage.includes('meeting_end'), true);
assert.equal(googlePlan.acceptance.required_local_snapshots.includes('active_speaker'), true);
assert.equal(googlePlan.commands.validate_real_intake.includes('meeting-platform:real-intake'), true);
assert.equal(googlePlan.files.evidence_package, 'tmp/evidence/meeting-platform-evidence-packages/google_meet.json');
assert.equal(googlePlan.operator_steps.some((step) => step.id === 'insert_live_annotation_sample'), true);
assert.equal(googlePlan.operator_steps.some((step) => step.id === 'run_real_intake_gate'), true);
assert.equal(googlePlan.next_actions.includes('capture_live_dom_snapshots_for_local_observer'), true);

const missingZoomPlan = buildMeetingPlatformFieldIntakePlan('zoom', {
  baseUrl,
  env: {},
});
assert.equal(missingZoomPlan.status, 'provider_setup_missing_env');
assert.equal(missingZoomPlan.next_actions.includes('configure_env:ZOOM_WEBHOOK_SECRET_TOKEN'), true);

const matrix = buildMeetingPlatformFieldIntakeMatrix({
  baseUrl,
  env: productionEnv,
});
assert.equal(matrix.schema, 'meeting_platform_field_intake_matrix');
assert.equal(matrix.platform_count, 5);
assert.deepEqual(matrix.platforms, ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').provider_endpoint, `${baseUrl}/api/platform-events/teams`);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').missing_env.length, 0);
assert.equal(matrix.next_actions.includes('capture_real_provider_start_end_events'), true);

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
assert.equal(kit.platformFieldIntakePlan('webex').platform, 'webex');
assert.equal(kit.platformFieldIntakeMatrix({ platforms: ['google-meet', 'lark'] }).platform_count, 2);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_field_intake_matrix.platform_count, 1);

console.log('ok meeting platform field intake');
