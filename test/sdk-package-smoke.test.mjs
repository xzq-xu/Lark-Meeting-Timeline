import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const sdkDir = new URL('../packages/meeting-timeline-sdk', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-timeline-sdk-package-'));
const packDir = join(tmpDir, 'pack');
const consumerDir = join(tmpDir, 'consumer');
await mkdir(packDir, { recursive: true });
await mkdir(consumerDir, { recursive: true });

const { stdout: packStdout } = await execFileAsync('npm', [
  'pack',
  '--json',
  '--pack-destination',
  packDir,
], {
  cwd: sdkDir,
});
const packInfo = JSON.parse(packStdout)[0];
assert.equal(packInfo.name, '@ai-annotation/meeting-timeline-sdk');
const packedFiles = packInfo.files.map((item) => item.path).sort();
assert.equal(packedFiles.includes('index.mjs'), true);
assert.equal(packedFiles.includes('index.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-kit.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-kit.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-integration-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-integration-runtime.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-event.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-event.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-registry.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-registry.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-host-integration.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-host-integration.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-provider-connection.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-provider-connection.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-subscription-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-subscription-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-speaker-track.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-speaker-track.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-participant-track.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-participant-track.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-timeline-view.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-timeline-view.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-annotation-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-annotation-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-clock-sync.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-clock-sync.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-session-binding.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-session-binding.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-realtime-annotation.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-realtime-annotation.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-artifact-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-artifact-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adaptation-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adaptation-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-bundle.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-bundle.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-contract.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-contract.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-sample.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-sample.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-real-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-real-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-field-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-field-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-handoff-readiness.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-handoff-readiness.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.d.ts'), true);
assert.equal(packedFiles.includes('README.md'), true);
assert.equal(packedFiles.some((item) => item.startsWith('test/')), false);
assert.equal(packedFiles.some((item) => item.startsWith('scripts/')), false);

await writeFile(join(consumerDir, 'package.json'), JSON.stringify({
  type: 'module',
  private: true,
  dependencies: {
    '@ai-annotation/meeting-timeline-sdk': sdkDir.pathname,
  },
}, null, 2), 'utf8');

await execFileAsync('npm', [
  'install',
  '--ignore-scripts',
  '--no-audit',
  '--no-fund',
], {
  cwd: consumerDir,
});

await writeFile(join(consumerDir, 'smoke.mjs'), `
import assert from 'node:assert/strict';
import {
  SDK_VERSION,
  createMeetingTimelineClient,
} from '@ai-annotation/meeting-timeline-sdk';
import {
  createMeetingPlatformTimelineKit,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit';
import {
  assertMeetingPlatformIntegrationRuntimeManifest,
  buildMeetingPlatformIntegrationRuntimeManifest,
  createMeetingPlatformIntegrationBrowserRuntime,
  createMeetingPlatformIntegrationContentScriptBridge,
  createMeetingPlatformIntegrationRuntime,
  detectMeetingPlatformForBrowser,
  installMeetingPlatformIntegrationContentScriptBridge,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  buildMeetingPlatformAnnotationRuntimeEvent,
  createMeetingPlatformRuntimeEventClient,
  meetingPlatformRuntimeEventEndpoint,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event';
import {
  assertMeetingPlatformRegistryManifest,
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryEntry,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry';
import {
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformRolloutPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';
import {
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy';
import {
  buildMeetingPlatformEvidenceCorrelation,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-correlation';
import {
  createMeetingPlatformEvidenceSession,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session';
import {
  assertMeetingPlatformLiveAdapterReadiness,
  assertMeetingPlatformLiveAdapterReadinessMatrix,
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterHandoffBundle,
  buildMeetingPlatformLiveAdapterMatrix,
  buildMeetingPlatformLiveAdapterPlan,
  buildMeetingPlatformLiveAdapterReadiness,
  buildMeetingPlatformLiveAdapterReadinessMatrix,
  createMeetingPlatformLiveAdapter,
  createMeetingPlatformLiveAdapterSuite,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter';
import {
  buildMeetingPlatformEvidencePackage,
  verifyMeetingPlatformEvidencePackage,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package';
import {
  buildMeetingPlatformHostIntegrationPlan,
  buildMeetingPlatformHostIntegrationScaffold,
  assertMeetingPlatformHostIntegrationScaffold,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-host-integration';
import {
  buildMeetingPlatformProviderConnectionMatrix,
  buildMeetingPlatformProviderConnectionPack,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection';
import {
  buildMeetingPlatformSubscriptionHandoff,
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-subscription-handoff';
import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
  buildMeetingPlatformSpeakerTrackPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track';
import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
  buildMeetingPlatformParticipantTrackPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track';
import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
  buildMeetingPlatformTimelineViewPlan,
  zoomMeetingPlatformTimelineViewport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view';
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
  buildMeetingPlatformAnnotationIntakePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake';
import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncPlan,
  buildMeetingPlatformClockSyncReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync';
import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
  buildMeetingPlatformSessionBindingPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding';
import {
  buildMeetingPlatformRealtimeAnnotation,
  buildMeetingPlatformRealtimeAnnotationMatrix,
  buildMeetingPlatformRealtimeAnnotationPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation';
import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
  buildMeetingPlatformArtifactHandoffPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-artifact-handoff';
import {
  buildMeetingPlatformAdaptationPackage,
  buildMeetingPlatformAdaptationPackageMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package';
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle';
import {
  assertMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractAcceptanceReport,
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract';
import {
  buildMeetingPlatformAdapterSamplePlan,
  runMeetingPlatformAdapterSample,
  runMeetingPlatformAdapterSampleMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-sample';
import {
  buildMeetingPlatformRealEvidenceIntakePlan,
  buildMeetingPlatformRealEvidenceIntakeReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-real-intake';
import {
  buildMeetingPlatformFieldIntakeMatrix,
  buildMeetingPlatformFieldIntakePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake';
import {
  buildMeetingPlatformHandoffReadiness,
  buildMeetingPlatformHandoffReadinessMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness';
import {
  buildMeetingPlatformFieldCaptureManifest,
  buildMeetingPlatformFieldCollectorConfig,
  buildMeetingPlatformFieldEvidenceBundle,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';
import {
  normalizeGoogleMeetEvent,
} from '@ai-annotation/meeting-timeline-sdk/adapters/google-meet';
import {
  buildMeetingAppLaunchGate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate';
import {
  buildMeetingAppDomAdaptationDiagnosisMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';

assert.equal(SDK_VERSION, '0.1.0');
const client = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true }), {
    headers: { 'content-type': 'application/json' },
  }),
});
assert.equal(typeof client.startMeeting, 'function');

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'http://localhost:8787',
  verify: false,
});
assert.equal(kit.platformRolloutPlan('google-meet').platform, 'google_meet');
assert.equal(kit.platformAdaptationRunbook('zoom').platform, 'zoom');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_rollout.type, 'meeting_platform_rollout_summary');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_registry_manifest.platform_count, 1);
const integrationRuntimeManifest = buildMeetingPlatformIntegrationRuntimeManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
});
assert.equal(integrationRuntimeManifest.host_integration_ready, true);
assert.equal(assertMeetingPlatformIntegrationRuntimeManifest(integrationRuntimeManifest).platform_count, 2);
const integrationRuntime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
});
assert.equal(integrationRuntime.manifest().host_integration_ready, true);
assert.equal(integrationRuntime.runtimeBundles().platform_count, 1);
assert.equal(detectMeetingPlatformForBrowser({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(createMeetingPlatformIntegrationBrowserRuntime(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  url: 'https://meet.google.com/abc-defg-hij',
}).detect().platform, 'google_meet');
assert.equal(createMeetingPlatformIntegrationContentScriptBridge(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  url: 'https://meet.google.com/abc-defg-hij',
  extensionMessaging: false,
}).detect().platform, 'google_meet');
assert.equal(typeof installMeetingPlatformIntegrationContentScriptBridge, 'function');
assert.equal(integrationRuntime.timelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [{
    id: 'runtime-smoke-note',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  }],
}).diagnostics.marker_count, 1);
const runtimeEvent = buildMeetingPlatformAnnotationRuntimeEvent('google-meet', {
  annotation: {
    id: 'package-runtime-event-note-1',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
}, {
  now: () => 1_782_614_402_000,
});
assert.equal(runtimeEvent.schema, MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA);
assert.equal(runtimeEvent.action, 'insert_annotation');
assert.equal(meetingPlatformRuntimeEventEndpoint({
  baseUrl: 'http://localhost:8787',
}), 'http://localhost:8787/api/meeting-platform/runtime-events');
assert.equal(typeof createMeetingPlatformRuntimeEventClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response('{}'),
}).send, 'function');
assert.equal(kit.platformRegistryEntry('google-meet').annotations.timestamp_field, 'captured_at_ms');
assert.equal(kit.platformRegistryManifest({ platforms: ['zoom'] }).rows[0].platform, 'zoom');
assert.equal(meetingPlatformEventAdapterFor('teams').key, 'microsoft_teams');
assert.equal(buildMeetingPlatformRegistryEntry('google-meet', {
  baseUrl: 'http://localhost:8787',
}).runtime.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(buildMeetingPlatformRegistryManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).provider_required_for_realtime_count, 0);
assert.equal(buildMeetingPlatformRegistryAcceptanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted, true);
assert.equal(assertMeetingPlatformRegistryManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted, true);
assert.equal(kit.platformLiveAdapterHandoff('zoom').sdk.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(kit.platformLiveAdapterHandoffBundle({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformHostIntegrationPlan({
  platforms: ['zoom'],
}).runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(kit.platformHostIntegrationPlan({
  platforms: ['zoom'],
}).runtime_bundle_matrix.platform_count, 1);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'src/meeting-platform-host.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-runtime-bundles.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-integration-runtime.mjs'), true);
assert.equal(kit.platformProviderConnectionPack('zoom').security.verifier, 'verifyZoomWebhookEvent');
assert.equal(kit.platformProviderConnectionMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSubscriptionHandoff('zoom').schema, 'meeting_platform_subscription_handoff');
assert.equal(kit.platformSubscriptionHandoffMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSpeakerTrackPlan('zoom').schema, 'meeting_platform_speaker_track_plan');
assert.equal(kit.platformSpeakerTrackMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSpeakerTrack('zoom', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      speaker_name: 'Ada',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_402_000,
      speaker_name: 'Ada',
    },
  ],
}).mark_count, 1);
assert.equal(kit.platformParticipantTrackPlan('zoom').schema, 'meeting_platform_participant_track_plan');
assert.equal(kit.platformParticipantTrackMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformParticipantTrack('zoom', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).mark_count, 1);
assert.equal(kit.platformTimelineViewPlan('zoom').schema, 'meeting_platform_timeline_view_plan');
assert.equal(kit.platformTimelineViewMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformTimelineView('zoom', {
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'note-1',
      label: 'why?',
      captured_at_ms: 1_782_614_401_000,
    },
  ],
}).diagnostics.marker_count, 1);
assert.equal(kit.zoomPlatformTimelineViewport({
  start_ms: 0,
  duration_ms: 600_000,
  full_duration_ms: 600_000,
}, 2).duration_ms, 300_000);
assert.equal(kit.platformAnnotationIntakePlan('zoom').schema, 'meeting_platform_annotation_intake_plan');
assert.equal(kit.platformAnnotationIntakeMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformAnnotationIntake('zoom', {
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-2',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'ready_to_insert_current_axis');
assert.equal(kit.platformClockSyncPlan('zoom').schema, 'meeting_platform_clock_sync_plan');
assert.equal(kit.platformClockSyncMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformClockSync('zoom', {
  clock_sync: {
    offset_ms: 10,
    rtt_ms: 40,
  },
}).status, 'clock_sync_ready');
assert.equal(kit.platformSessionBindingPlan('zoom').schema, 'meeting_platform_session_binding_plan');
assert.equal(kit.platformSessionBindingMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSessionBinding('zoom', {
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: 1_782_614_400_000,
  },
}).status, 'open_axis_from_local_observer');
assert.equal(kit.platformRealtimeAnnotationPlan('zoom').schema, 'meeting_platform_realtime_annotation_plan');
assert.equal(kit.platformRealtimeAnnotationMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformRealtimeAnnotation('zoom', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-realtime-smoke',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'start_axis_then_insert');
assert.equal(kit.platformArtifactHandoffPlan('zoom').schema, 'meeting_platform_artifact_handoff_plan');
assert.equal(kit.platformArtifactHandoffMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformArtifactHandoff('zoom', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      artifact_kind: 'transcript',
      artifact_url: 'https://zoom.us/transcript.vtt',
    },
  ],
}).fetch_request_count, 1);
assert.equal(kit.platformAdaptationPackage('zoom').schema, 'meeting_platform_adaptation_package');
assert.equal(kit.platformAdaptationPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformAdaptationPackage('google-meet', {
  baseUrl: 'http://localhost:8787',
}).extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(buildMeetingPlatformAdaptationPackageMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].sdk_wiring_ready, true);
assert.equal(kit.platformRuntimeBundle('google-meet').schema, 'meeting_platform_runtime_bundle');
assert.equal(kit.platformRuntimeBundleMatrix({
  platforms: ['google-meet'],
}).platform_count, 1);
const smokeGoogleRuntimeBundle = buildMeetingPlatformRuntimeBundle('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(smokeGoogleRuntimeBundle.browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(smokeGoogleRuntimeBundle.runtime.content_script_bridge.install_function, 'installMeetingPlatformIntegrationContentScriptBridge');
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.content_script_insert_annotation.type, 'meeting_timeline.insert_mark');
assert.equal(buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].provider_required_for_realtime, false);

const rollout = buildMeetingPlatformRolloutPlan('teams', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(rollout.platform, 'microsoft_teams');

const runbook = buildMeetingPlatformAdaptationRunbook('webex', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(runbook.type, 'meeting_platform_adaptation_runbook');
assert.equal(runbook.steps.some((item) => item.id === 'validate_rollout'), true);

const strategy = buildMeetingPlatformAdaptationStrategy('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(strategy.platform, 'google_meet');
assert.equal(strategy.realtime_axis.provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformAdaptationStrategyMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].platform, 'zoom');
assert.equal(buildMeetingPlatformEvidenceCorrelation('zoom', {}).status, 'single_source');
const session = createMeetingPlatformEvidenceSession('zoom', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(session.summary().platform, 'zoom');
assert.equal(session.summary().package_ready_for_handoff, false);
assert.equal(kit.platformEvidenceSession('google-meet').platform, 'google_meet');
const liveAdapter = createMeetingPlatformLiveAdapter('google-meet', client, {
  baseUrl: 'http://localhost:8787',
});
assert.equal(liveAdapter.summary().platform, 'google_meet');
assert.equal(buildMeetingPlatformLiveAdapterPlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).live_adapter.kit_method, 'platformLiveAdapter');
assert.equal(buildMeetingPlatformLiveAdapterMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].platform, 'zoom');
assert.equal(buildMeetingPlatformLiveAdapterReadiness('zoom', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_live_adapter_readiness');
assert.equal(buildMeetingPlatformLiveAdapterReadinessMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(buildMeetingPlatformLiveAdapterHandoff('zoom', {
  baseUrl: 'http://localhost:8787',
}).host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingPlatformLiveAdapterHandoffBundle({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_live_readiness, 'npm run meeting-platform:live-readiness');
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadiness('zoom', {
    baseUrl: 'http://localhost:8787',
  }),
  /Meeting platform live adapter readiness failed/,
);
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadinessMatrix({
    baseUrl: 'http://localhost:8787',
    platforms: ['zoom'],
  }),
  /Meeting platform live adapter readiness matrix failed/,
);
assert.equal(createMeetingPlatformLiveAdapterSuite(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).summary().platform_count, 1);
assert.equal(kit.platformLiveAdapter('zoom').platform, 'zoom');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platforms[0], 'zoom');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.runtime_bundles, '/api/meeting-platform/runtime-bundles');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.extension_plan, '/api/meeting-platform/extension-plan');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.integration_runtime, '/api/meeting-platform/integration-runtime');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.runtime_events, '/api/meeting-platform/runtime-events');
assert.equal(assertMeetingPlatformHostIntegrationScaffold(buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
})).accepted, true);
assert.equal(buildMeetingPlatformProviderConnectionPack('zoom', {
  baseUrl: 'http://localhost:8787',
}).realtime_annotation_policy.annotation_timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingPlatformProviderConnectionMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).packs[0].official_docs.some((doc) => doc.url.includes('zoom.us')), true);
assert.equal(buildMeetingPlatformSubscriptionHandoff('zoom', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_subscription_handoff');
assert.equal(buildMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(buildMeetingPlatformSpeakerTrackPlan('google-meet').output_contract.transcript_required, false);
assert.equal(buildMeetingPlatformSpeakerTrackMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformSpeakerTrack('google-meet', {
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
}).marks[0].intent, 'speaker_track');
assert.equal(buildMeetingPlatformParticipantTrackPlan('google-meet').transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformParticipantTrackMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformParticipantTrack('google-meet', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).marks[0].intent, 'participant_track');
assert.equal(buildMeetingPlatformTimelineViewPlan('google-meet').output_contract.renderer_agnostic, true);
assert.equal(buildMeetingPlatformTimelineViewMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformTimelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'note-1',
      label: 'why?',
      captured_at_ms: 1_782_614_401_000,
    },
  ],
}).markers[0].x_ratio, 0.001667);
assert.equal(zoomMeetingPlatformTimelineViewport({
  start_ms: 0,
  duration_ms: 600_000,
  full_duration_ms: 600_000,
}, 2).duration_ms, 300_000);
assert.equal(buildMeetingPlatformAnnotationIntakePlan('google-meet').realtime_policy.transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformAnnotationIntakeMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-intake-1',
    label: 'why?',
    captured_at_ms: 1_782_614_403_000,
  },
}).normalized_time_ms, 3_000);
assert.equal(buildMeetingPlatformClockSyncPlan('google-meet').required, true);
assert.equal(buildMeetingPlatformClockSyncMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformClockSyncReport('google-meet', {
  samples: [
    {
      client_send_at_ms: 1_782_614_400_000,
      server_time_ms: 1_782_614_400_060,
      client_receive_at_ms: 1_782_614_400_100,
    },
  ],
}).recommended_offset_ms, 10);
assert.equal(buildMeetingPlatformSessionBindingPlan('google-meet').realtime_policy.provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformSessionBindingMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformSessionBinding('google-meet', {
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_400_000,
  },
}).status, 'open_axis_from_local_observer');
assert.equal(buildMeetingPlatformRealtimeAnnotationPlan('google-meet').pipeline.includes('clock_sync'), true);
assert.equal(buildMeetingPlatformRealtimeAnnotationMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-realtime-direct',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'start_axis_then_insert');
assert.equal(buildMeetingPlatformArtifactHandoffPlan('google-meet').transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformArtifactHandoffMatrix({
  platforms: ['google-meet'],
}).realtime_blocking_count, 0);
assert.equal(buildMeetingPlatformArtifactHandoff('google-meet', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      artifact_kind: 'transcript',
      artifact_id: 'transcript-1',
    },
  ],
}).transcript_import_count, 1);
assert.equal(buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'http://localhost:8787',
}).supported_surfaces.browser_observer, true);
assert.equal(buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'http://localhost:8787',
}).annotations.endpoints.runtimeEvents, 'http://localhost:8787/api/meeting-platform/runtime-events');
assert.equal(buildMeetingPlatformAdapterContractMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).contracts[0].platform, 'zoom');
assert.equal(buildMeetingPlatformAdapterContractAcceptanceReport('google-meet', {
  baseUrl: 'http://localhost:8787',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterContractAcceptanceMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted_count, 1);
assert.equal(assertMeetingPlatformAdapterContract('zoom', {
  baseUrl: 'http://localhost:8787',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterSamplePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).platform, 'google_meet');
assert.equal((await runMeetingPlatformAdapterSample('zoom', {
  baseUrl: 'http://localhost:8787',
})).accepted, true);
assert.equal((await runMeetingPlatformAdapterSampleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['webex'],
})).accepted_count, 1);
assert.equal(buildMeetingPlatformRealEvidenceIntakePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).provider_endpoint, 'http://localhost:8787/api/platform-events/google-meet');
assert.equal(buildMeetingPlatformRealEvidenceIntakeReport('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).accepted, false);
assert.equal(buildMeetingPlatformFieldIntakePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_field_intake_plan');
assert.equal(buildMeetingPlatformFieldIntakeMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformHandoffReadiness('google-meet', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_handoff_readiness');
assert.equal(buildMeetingPlatformHandoffReadinessMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).platform_count, 1);

const evidencePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: [],
  meetingAppRecords: [],
}, {
  baseUrl: 'http://localhost:8787',
  includeRunbook: false,
});
assert.equal(evidencePackage.schema, 'meeting_platform_evidence_package');
assert.equal(evidencePackage.platform, 'google_meet');
assert.equal(verifyMeetingPlatformEvidencePackage(evidencePackage, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).type, 'meeting_platform_evidence_package_verification');
assert.equal(buildMeetingPlatformFieldEvidenceBundle('google-meet', evidencePackage, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).schema, 'meeting_platform_field_evidence_bundle');
assert.equal(buildMeetingPlatformFieldCaptureManifest('google-meet', {
  baseUrl: 'http://localhost:8787',
  evidencePackage,
}).schema, 'meeting_platform_field_capture_manifest');
assert.equal(buildMeetingPlatformFieldCollectorConfig('google-meet', {
  baseUrl: 'http://localhost:8787',
}).browser_observer.matches.includes('https://meet.google.com/*'), true);
assert.equal(kit.platformFieldCaptureManifest('google-meet', {
  evidencePackage,
}).file_contract.files.evidence_package.endsWith('/google_meet.json'), true);
assert.equal(kit.platformAdapterContract('google-meet').schema, 'meeting_platform_adapter_contract');
assert.equal(kit.platformAdapterContractMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformAdapterContractAcceptance('google-meet').accepted, true);
assert.equal(kit.platformAdapterContractAcceptanceMatrix({ platforms: ['zoom'] }).accepted_count, 1);
assert.equal(kit.platformRealEvidenceIntakePlan('google-meet').schema, 'meeting_platform_real_evidence_intake_plan');
assert.equal(kit.platformRealEvidenceIntake('google-meet', {}, {
  requireProductionReady: false,
}).accepted, false);
assert.equal(kit.platformFieldCollectorConfig('google-meet').schema, 'meeting_platform_field_collector_config');
assert.equal(kit.platformFieldEvidenceBundle('google-meet', evidencePackage, {
  requireProductionReady: false,
}).field_capture_plan.schema, 'meeting_platform_field_capture_plan');
assert.equal(kit.platformFieldIntakePlan('google-meet').schema, 'meeting_platform_field_intake_plan');
assert.equal(kit.platformFieldIntakeMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformHandoffReadiness('google-meet').schema, 'meeting_platform_handoff_readiness');
assert.equal(kit.platformHandoffReadinessMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(buildMeetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet'],
}).schema, 'meeting_app_dom_adaptation_diagnosis_matrix');
assert.equal(kit.meetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet'],
}).platform_count, 1);

const gate = buildMeetingAppLaunchGate('google-meet', {
  allowFixtureEvidence: true,
  requireProductionReady: false,
});
assert.equal(gate.platform, 'google_meet');

const signals = normalizeGoogleMeetEvent({
  type: 'google.workspace.meet.conference.v2.started',
  subject: '//meet.googleapis.com/conferenceRecords/demo',
  time: '2026-06-26T02:00:00.000Z',
  data: {
    conferenceRecord: {
      name: 'conferenceRecords/demo',
      space: 'spaces/demo',
      startTime: '2026-06-26T02:00:00.000Z',
    },
  },
});
assert.equal(signals.some((signal) => signal.type === 'meeting_started'), true);

console.log('ok consumer package imports');
`, 'utf8');

const { stdout: smokeStdout } = await execFileAsync(process.execPath, ['smoke.mjs'], {
  cwd: consumerDir,
});
assert.match(smokeStdout, /ok consumer package imports/);

const installedPackage = JSON.parse(await readFile(
  join(consumerDir, 'node_modules/@ai-annotation/meeting-timeline-sdk/package.json'),
  'utf8',
));
assert.equal(installedPackage.version, '0.1.0');

console.log('ok meeting timeline SDK package smoke');
