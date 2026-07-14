import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import {
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  allPlatformCapabilityContracts,
  allPlatformIntegrationPlans,
  allPlatformSetupManifests,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
  platformEventEndpoint,
  platformSetupManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import {
  MEETING_PLATFORM_EVENT_ADAPTERS,
  meetingPlatformEventAdapterFor,
} from '../packages/meeting-timeline-sdk/adapters/platform-registry.mjs';
import {
  assertMeetingPlatformConformanceReport,
  buildMeetingPlatformConformanceReport,
} from '../packages/meeting-timeline-sdk/adapters/platform-conformance.mjs';
import { createMeetingPlatformTimelineKit } from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import * as transcriptAdapters from '../packages/meeting-timeline-sdk/adapters/transcript.mjs';

const baseUrl = 'https://timeline.example.com';
const expectedPlatforms = ['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex'];

assert.deepEqual([...MEETING_PLATFORM_KEYS], expectedPlatforms);
assert.equal(MEETING_PLATFORM_ALIASES['local-detector'], 'local_detector');
assert.equal(MEETING_PLATFORM_ALIASES.feishu, 'lark');
assert.equal(MEETING_PLATFORM_ALIASES['google-meet'], 'google_meet');
assert.equal(MEETING_PLATFORM_ALIASES['microsoft-teams'], 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('meet'), 'google_meet');
assert.equal(normalizeMeetingPlatform('desktop-observer'), 'local_detector');
assert.equal(normalizeMeetingPlatform('teams'), 'microsoft_teams');
assert.equal(normalizeMeetingPlatform('cisco-webex'), 'webex');
assert.equal(normalizeMeetingPlatform('lark-suite'), 'lark');
assert.throws(() => normalizeMeetingPlatform('unknown-meeting-platform'), /Unsupported meeting platform/);
assert.equal(meetingPlatformEventAdapterFor('unknown-meeting-platform'), null);

const conformance = buildMeetingPlatformConformanceReport({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(conformance.schema, 'meeting_platform_conformance_report');
assert.equal(conformance.accepted, true);
assert.equal(conformance.platform_count, 5);
assert.equal(conformance.accepted_count, 5);
assert.equal(conformance.normalizer_count, 5);
assert.equal(conformance.adapter_contract_accepted_count, 5);
assert.equal(conformance.adapter_route_ready_count, 5);
assert.equal(conformance.runtime_ready_count, 5);
assert.equal(conformance.candidate_observer_count, 5);
assert.equal(conformance.provider_required_for_realtime_count, 0);
assert.equal(conformance.transcript_blocking_count, 0);
assert.equal(conformance.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(conformance.rows.every((row) => row.timestamp_field === 'captured_at_ms'), true);
assert.equal(assertMeetingPlatformConformanceReport({
  baseUrl,
  platforms: ['google-meet'],
}).accepted, true);
const kit = createMeetingPlatformTimelineKit({ baseUrl });
assert.equal(kit.platformConformance({ platforms: ['zoom'] }).rows[0].platform, 'zoom');
assert.equal(kit.assertPlatformConformance({ platforms: ['zoom'] }).accepted, true);

const setupRows = allPlatformSetupManifests({ baseUrl });
const capabilityRows = allPlatformCapabilityContracts({ baseUrl });
const integrationRows = allPlatformIntegrationPlans({ baseUrl });
assert.deepEqual(setupRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(capabilityRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(integrationRows.map((item) => item.platform), expectedPlatforms);
assert.deepEqual(MEETING_PLATFORM_EVENT_ADAPTERS.map((item) => item.key), expectedPlatforms);

const envExample = readFileSync(new URL('../.env.example', import.meta.url), 'utf8');

for (const platform of MEETING_PLATFORM_KEYS) {
  const eventAdapter = meetingPlatformEventAdapterFor(platform);
  const manifest = platformSetupManifest(platform, { baseUrl });
  const capability = platformCapabilityContract(platform, { baseUrl });
  const plan = buildPlatformIntegrationPlan(platform, { baseUrl });
  assert.equal(eventAdapter.key, platform);
  assert.equal(eventAdapter.source, platform === 'local_detector' ? 'local_detector' : `${platform}_webhook`);
  assert.equal(typeof eventAdapter.normalize, 'function');
  assert.equal(eventAdapter.aliases.includes(platform), true);
  for (const alias of eventAdapter.aliases) {
    assert.equal(meetingPlatformEventAdapterFor(alias).key, platform);
  }
  assert.equal(manifest.platform, platform);
  assert.equal(manifest.capabilities.platform, platform);
  assert.equal(capability.platform, platform);
  assert.equal(capability.endpoints.platform_events, platformEventEndpoint(baseUrl, platform));
  assert.equal(capability.endpoints.transcript_import, `${baseUrl}/api/import/transcript`);
  assert.equal(plan.platform, platform);
  assert.equal(plan.endpoints.platform_events, platformEventEndpoint(baseUrl, platform));
  assert.equal(plan.modules.ingest, '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest');
  assert.equal(plan.realtime_axis.signal_types.includes('meeting_started'), true);
  assert.equal(plan.realtime_axis.signal_types.includes('meeting_ended'), true);
  assert.equal(plan.realtime_axis.local_observer.session_discovery_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-session-discovery');
  assert.equal(plan.realtime_axis.local_observer.browser_meeting_module, '@ai-annotation/meeting-timeline-sdk/adapters/browser-meeting');
  assert.equal(plan.realtime_axis.local_observer.native_meeting_module, '@ai-annotation/meeting-timeline-sdk/adapters/native-meeting');
  assert.equal(plan.realtime_axis.local_observer.meeting_apps_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_capture_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_monitor_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-monitor');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_runtime_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_browser_runtime_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_content_script_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_extension_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_snapshot_recorder_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_fixtures_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures');
  assert.equal(plan.realtime_axis.local_observer.meeting_app_gate_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate');
  assert.equal(plan.realtime_axis.local_observer.meeting_source_module, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-source');
  assert.equal(plan.speaker_activity.active_speaker_module, '@ai-annotation/meeting-timeline-sdk/adapters/active-speaker');
  assert.equal(plan.realtime_annotations.required_field, 'captured_at_ms');
  assert.equal(typeof plan.post_meeting_transcript.strategy, 'string');
  assert.equal(typeof plan.readiness.ready, 'boolean');
  assert.equal(typeof capability.sdk_modules.events, 'string');
  assert.equal(capability.sdk_modules.url_detection, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url');
  assert.equal(capability.sdk_modules.local_observer, '@ai-annotation/meeting-timeline-sdk/adapters/local-observer');
  assert.equal(capability.sdk_modules.session_discovery, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-session-discovery');
  assert.equal(capability.sdk_modules.active_speaker, '@ai-annotation/meeting-timeline-sdk/adapters/active-speaker');
  assert.equal(capability.sdk_modules.browser_meeting, '@ai-annotation/meeting-timeline-sdk/adapters/browser-meeting');
  assert.equal(capability.sdk_modules.native_meeting, '@ai-annotation/meeting-timeline-sdk/adapters/native-meeting');
  assert.equal(capability.sdk_modules.meeting_apps, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps');
  assert.equal(capability.sdk_modules.meeting_app_capture, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture');
  assert.equal(capability.sdk_modules.meeting_app_monitor, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-monitor');
  assert.equal(capability.sdk_modules.meeting_app_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime');
  assert.equal(capability.sdk_modules.meeting_app_browser_runtime, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
  assert.equal(capability.sdk_modules.meeting_app_content_script, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
  assert.equal(capability.sdk_modules.meeting_app_extension, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension');
  assert.equal(capability.sdk_modules.meeting_app_snapshot_recorder, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder');
  assert.equal(capability.sdk_modules.meeting_app_fixtures, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures');
  assert.equal(capability.sdk_modules.meeting_app_gate, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate');
  assert.equal(capability.sdk_modules.meeting_source, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-source');
  assert.equal(capability.sdk_modules.ingest, '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest');
  assert.equal(capability.sdk_modules.timeline_bridge, '@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge');
  assert.equal(capability.sdk_modules.signal_reconciler, '@ai-annotation/meeting-timeline-sdk/adapters/signal-reconciler');
  assert.equal(capability.sdk_modules.webhook_handler, '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler');
  assert.equal(capability.sdk_modules.webhook_router, '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router');
  assert.equal(capability.sdk_modules.acceptance, '@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance');
  assert.equal(capability.sdk_modules.artifact_plan, '@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan');
  assert.equal(capability.sdk_modules.artifact_fetch, '@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch');
  assert.equal(capability.sdk_modules.onboarding, '@ai-annotation/meeting-timeline-sdk/adapters/platform-onboarding');
  assert.equal(capability.sdk_modules.platform_http, '@ai-annotation/meeting-timeline-sdk/adapters/platform-http');
  assert.equal(capability.sdk_modules.platform_node, '@ai-annotation/meeting-timeline-sdk/adapters/platform-node');
  assert.equal(capability.sdk_modules.platform_capture, '@ai-annotation/meeting-timeline-sdk/adapters/platform-capture');
  assert.equal(capability.sdk_modules.platform_gate, '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate');
  assert.equal(capability.sdk_modules.fixtures, '@ai-annotation/meeting-timeline-sdk/adapters/platform-fixtures');
  assert.equal(capability.sdk_modules.platform_kit, '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit');
  assert.equal(typeof capability.sdk_modules.transcript, 'string');
  assert.equal(typeof capability.sdk_modules.setup, 'string');
  assert.equal(typeof capability.sdk_modules.security, 'string');
  assert.equal(Array.isArray(capability.realtime_axis.signal_types), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_started'), true);
  assert.equal(capability.realtime_axis.signal_types.includes('meeting_ended'), true);
  assert.equal(Array.isArray(capability.speaker_activity.signal_types), true);
  assert.equal(capability.speaker_activity.signal_types.includes('speaker_started'), true);
  assert.equal(capability.speaker_activity.signal_types.includes('speaker_ended'), true);

  for (const envName of manifest.required_security_env ?? []) {
    assert.match(envExample, new RegExp(`^${envName}=`, 'm'), `${platform} missing ${envName} in .env.example`);
  }

  const transcriptNormalizer = capability.post_meeting_transcript?.sdk_normalizer;
  if (transcriptNormalizer) {
    assert.equal(typeof transcriptAdapters[transcriptNormalizer], 'function', `${platform} transcript normalizer missing`);
  }
}

console.log('ok meeting platform SDK conformance');
