import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA,
  MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
  assertMeetingPlatformIntegrationRuntimeManifest,
  buildMeetingPlatformIntegrationRuntimeManifest,
  createMeetingPlatformIntegrationRuntime,
} from '../packages/meeting-timeline-sdk/adapters/platform-integration-runtime.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const baseUrl = 'https://timeline.example.com';
const manifest = buildMeetingPlatformIntegrationRuntimeManifest({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(manifest.schema, MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA);
assert.equal(manifest.platform_count, 2);
assert.deepEqual(manifest.platforms, ['google_meet', 'zoom']);
assert.equal(manifest.host_integration_ready, true);
assert.equal(manifest.blocking_count, 0);
assert.equal(manifest.warning_count, 1);
assert.equal(manifest.registry_acceptance.accepted, true);
assert.equal(manifest.runtime_bundle_matrix.runtime_ready_count, 2);
assert.equal(manifest.runtime_bundle_matrix.provider_required_for_realtime_count, 0);
assert.equal(manifest.adaptation_package_matrix.sdk_wiring_ready_count, 2);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').browser_match_count, 1);
assert.equal(manifest.rows.find((row) => row.platform === 'zoom').sdk_wiring_ready, true);
assert.equal(assertMeetingPlatformIntegrationRuntimeManifest(manifest).host_integration_ready, true);

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, method: 'startMeeting', input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, method: 'endMeeting', input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, method: 'insertMark', input };
  },
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true, method: 'insertMarks', input };
  },
  async importTranscript(input) {
    calls.push({ method: 'importTranscript', input });
    return { ok: true, method: 'importTranscript', input };
  },
};

const runtime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  verify: false,
});
assert.equal(runtime.schema, MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA);
assert.deepEqual(runtime.platforms, ['google_meet', 'zoom']);
assert.equal(runtime.manifest().host_integration_ready, true);
assert.equal(runtime.summary().host_integration_ready, true);
assert.equal(runtime.registry().acceptance.accepted, true);
assert.equal(runtime.runtimeBundle('google-meet').browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(runtime.runtimeBundles().platform_count, 2);
assert.equal(runtime.adaptationPackages().sdk_wiring_ready_count, 2);
assert.equal(runtime.adapter('google-meet').platform, 'google_meet');

const googleSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_782_614_400_000,
});
const observeResult = await runtime.observeMeetingApp('google-meet', googleSnapshot);
assert.equal(observeResult.action, 'observe_meeting_app');
assert.equal(observeResult.live_evidence.meeting_app_record_count, 1);
assert.equal(calls.some((call) => call.method === 'startMeeting'), true);

const insertResult = await runtime.insertAnnotation('google-meet', {
  annotation: {
    id: 'note-runtime-1',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
});
assert.equal(insertResult.action, 'insert_annotation');
assert.equal(insertResult.result.actions.includes('insert_mark'), true);
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'note-runtime-1'), true);

const providerResult = await runtime.handleEvent({
  ...buildPlatformFixtureEvent('google-meet', 'meeting_end'),
  action: 'provider_event',
  platform: 'google-meet',
});
assert.equal(providerResult.action, 'ingest_provider');
assert.equal(providerResult.live_evidence.provider_record_count, 1);
assert.equal(providerResult.diagnostic.raw_signal_count, 1);

const routedInsert = await runtime.handleEvent({
  action: 'insert_annotation',
  platform: 'zoom',
  annotation: {
    id: 'zoom-note-1',
    label: 'follow up',
    captured_at_ms: 1_782_614_405_000,
  },
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
});
assert.equal(routedInsert.result.status, 'ready_to_insert');
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'zoom-note-1'), true);

const speakerTrack = runtime.speakerTrack('google-meet', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      speaker_name: 'Ada',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_402_000,
      speaker_name: 'Ada',
    },
  ],
});
assert.equal(speakerTrack.mark_count, 1);
assert.equal(speakerTrack.marks[0].intent, 'speaker_track');

const participantTrack = await runtime.handleEvent({
  action: 'participant_track',
  platform: 'zoom',
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_403_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
});
assert.equal(participantTrack.mark_count, 1);

const view = runtime.timelineView('zoom', {
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'zoom-note-1',
      label: 'follow up',
      captured_at_ms: 1_782_614_405_000,
    },
  ],
  speakerTrack,
});
assert.equal(view.diagnostics.marker_count, 2);
assert.equal(view.markers.some((marker) => marker.rail === 'speaker'), true);
assert.equal((await runtime.handleEvent({ action: 'runtime_bundles' })).platform_count, 2);
assert.equal((await runtime.handleEvent({ action: 'manifest' })).host_integration_ready, true);
await assert.rejects(
  () => runtime.handleEvent({ action: 'insert_annotation' }),
  /platform is required/,
);

console.log('ok meeting platform integration runtime');
