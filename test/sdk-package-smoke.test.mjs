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
assert.equal(packedFiles.includes('adapters/platform-conformance.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-conformance.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-consumer-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-consumer-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-route.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-route.d.ts'), true);
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
assert.equal(packedFiles.includes('adapters/meeting-app-fixture-tracks.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-fixture-tracks.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-manifest.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-manifest.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-spec.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-spec.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-runtime-config.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-runtime-config.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-runtime.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-profile.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-profile.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-observer-scheduler.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-observer-scheduler.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host-verifier.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host-verifier.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-pipeline.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-pipeline.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-runtime.d.ts'), true);
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
	  resolveMeetingPlatformCandidates,
	  resolveMeetingPlatformForInput,
	  runMeetingPlatformIntegrationRuntimeManifest,
		} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  buildMeetingPlatformAnnotationRuntimeEvent,
  buildMeetingPlatformCandidateObservationRuntimeEvent,
  buildMeetingPlatformRunHandoffReadinessRuntimeEvent,
  buildMeetingPlatformRunManifestRuntimeEvent,
  buildMeetingPlatformRuntimeEventPlan,
  buildMeetingPlatformRuntimeEventPlanMatrix,
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
  assertMeetingPlatformConformanceReport,
  buildMeetingPlatformConformanceReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-conformance';
import {
  assertMeetingPlatformConsumerHandoff,
  buildMeetingPlatformConsumerHandoff,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff';
import {
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformRolloutPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';
import {
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy';
import {
  buildMeetingPlatformAdapterRoute,
  buildMeetingPlatformAdapterRouteMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route';
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
  buildMeetingAppAdapterFitMatrix,
  buildMeetingAppAdapterFitReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps';
import {
  assertMeetingAppRuntimeAdapterHandoff,
  buildMeetingAppDomAdaptationDiagnosisMatrix,
  buildMeetingAppRuntimeAdapterHandoffAcceptanceReport,
  buildMeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport,
  buildMeetingAppRuntimeAdapterHandoff,
  buildMeetingAppRuntimeAdapterHandoffMatrix,
  buildMeetingAppRuntimeAdapterHostPackage,
  buildMeetingAppRuntimeAdapterProfileMatrix,
  buildMeetingAppRuntimeObserverPlan,
  buildMeetingAppRuntimeObserverPlanMatrix,
  resolveMeetingAppRuntimeAdapterProfile,
  selectMeetingAppRuntimeAdapter,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';
import {
  assertMeetingAppAdapterManifest,
  assertMeetingAppAdapterManifestMatrix,
  buildMeetingAppAdapterManifest,
  buildMeetingAppAdapterManifestMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest';
import {
  assertMeetingAppAdapterSpec,
  assertMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpec,
  buildMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpecTemplate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec';
import {
  assertMeetingAppAdapterRuntimeConfig,
  assertMeetingAppAdapterRuntimeConfigMatrix,
  buildMeetingAppAdapterRuntimeConfig,
  buildMeetingAppAdapterRuntimeConfigMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config';
import {
  buildMeetingAppFixtureTrackReadinessReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks';
import {
  buildMeetingAppTrackPipeline,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-pipeline';
import {
  createMeetingAppTrackRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime';
import {
  createMeetingAppTimelineRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime';
import {
  buildMeetingAppObserverSchedulerConfig,
  buildMeetingAppObserverSchedulerConfigMatrix,
  createMeetingAppObserverScheduler,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-observer-scheduler';
import {
  buildMeetingPlatformRuntimeHostConfig,
  buildMeetingPlatformRuntimeHostConfigMatrix,
  buildMeetingPlatformRuntimeHostHandoff,
  buildMeetingPlatformRuntimeHostHandoffMatrix,
  createMeetingPlatformRuntimeHost,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host';
import {
  createMeetingPlatformRuntimeHostFixtureEnvironment,
  createMeetingPlatformRuntimeHostVerificationClient,
  runMeetingPlatformRuntimeHostReplay,
  runMeetingPlatformRuntimeHostReplayMatrix,
  runMeetingPlatformRuntimeHostVerification,
  runMeetingPlatformRuntimeHostVerificationMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier';

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
assert.equal(kit.meetingAppRuntimeAdapterProfile('https://meet.google.com/abc-defg-hij').platform, 'google_meet');
assert.equal(kit.meetingAppRuntimeAdapterProfileMatrix({ platforms: ['google-meet', 'teams'] }).runtime_ready_count, 2);
assert.equal(kit.meetingAppAdapterFit({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).ready_for_speaker_track, true);
assert.equal(kit.meetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).observer_runtime.factory, 'createMeetingAppBrowserRuntime');
assert.equal(kit.meetingAppObserverSchedulerConfig('google-meet').schema, 'meeting_app_observer_scheduler_config');
assert.equal(kit.meetingAppObserverSchedulerConfigMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.selectMeetingAppRuntimeAdapter('https://meet.google.com/abc-defg-hij').launch.runtime_options.runtimePreset, 'google_meet');
assert.equal(kit.meetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').readiness.ready_to_start, true);
assert.equal(kit.meetingAppRuntimeAdapterHandoffMatrix({ platforms: ['google-meet', 'teams'], surfaces: ['browser-extension'] }).handoff_count, 2);
assert.equal(kit.meetingAppRuntimeAdapterHostPackage({ platforms: ['google-meet'], surfaces: ['browser-extension'] }).accepted, true);
assert.equal(kit.meetingAppRuntimeAdapterHandoffAcceptance('https://meet.google.com/abc-defg-hij').accepted, true);
const integrationRuntimeManifest = buildMeetingPlatformIntegrationRuntimeManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
});
assert.equal(integrationRuntimeManifest.host_integration_ready, true);
assert.equal(integrationRuntimeManifest.platform_conformance_report.accepted, true);
assert.equal(assertMeetingPlatformIntegrationRuntimeManifest(integrationRuntimeManifest).platform_count, 2);
const integrationRuntimeRunManifest = await runMeetingPlatformIntegrationRuntimeManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
});
assert.equal(integrationRuntimeRunManifest.host_integration_ready, true);
assert.equal(integrationRuntimeRunManifest.handoff_readiness_matrix.runtime_host_replay_ready_count, 0);
const integrationRuntime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
});
assert.equal(integrationRuntime.manifest().host_integration_ready, true);
assert.equal(integrationRuntime.conformance().accepted, true);
assert.equal(integrationRuntime.assertConformance().accepted_count, 1);
assert.equal(integrationRuntime.manifest().adapter_route_matrix.platform_count, 1);
assert.equal((await integrationRuntime.runManifest()).host_integration_ready, true);
assert.equal(integrationRuntime.runtimeBundles().platform_count, 1);
assert.equal(integrationRuntime.adapterRoutes().rows[0].first_route, 'local_observer_axis');
assert.equal(integrationRuntime.adaptationStrategyMatrix().strategy_count, 1);
assert.equal(integrationRuntime.resolvePlatform({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(detectMeetingPlatformForBrowser({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(resolveMeetingPlatformForInput({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).supported, true);
assert.equal(resolveMeetingPlatformCandidates({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).selected_resolution.platform, 'google_meet');
assert.equal((await integrationRuntime.observePlatformCandidates({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  observedAtMs: 1_782_614_400_000,
})).signals[0].type, 'meeting_started');
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
assert.equal(buildMeetingPlatformCandidateObservationRuntimeEvent({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  now: () => 1_782_614_402_000,
}).action, 'observe_platform_candidates');
assert.equal(buildMeetingPlatformRunManifestRuntimeEvent({
  platforms: ['google-meet'],
  requireHandoffReady: true,
}, {
  now: () => 1_782_614_402_000,
}).action, 'run_manifest');
assert.equal(buildMeetingPlatformRunHandoffReadinessRuntimeEvent({
  platforms: ['zoom'],
}, {
  now: () => 1_782_614_402_000,
}).action, 'run_handoff_readiness');
assert.equal(meetingPlatformRuntimeEventEndpoint({
  baseUrl: 'http://localhost:8787',
}), 'http://localhost:8787/api/meeting-platform/runtime-events');
const smokeRuntimeEventClient = createMeetingPlatformRuntimeEventClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response('{}'),
});
assert.equal(typeof smokeRuntimeEventClient.send, 'function');
assert.equal(typeof smokeRuntimeEventClient.runManifest, 'function');
assert.equal(typeof smokeRuntimeEventClient.runHandoffReadiness, 'function');
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
assert.equal(buildMeetingPlatformConformanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(assertMeetingPlatformConformanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).rows[0].adapter_first_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformConsumerHandoff({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).consumer_ready_count, 1);
assert.equal(assertMeetingPlatformConsumerHandoff({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).rows[0].adapter_first_route, 'local_observer_axis');
assert.equal(kit.platformConsumerHandoff({ platforms: ['zoom'] }).entrypoints.kit_methods.includes('platformConsumerHandoff'), true);
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
}).files.some((file) => file.path === 'scripts/print-strategy.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-integration-runtime.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-consumer-handoff.mjs'), true);
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
assert.equal(buildMeetingAppFixtureTrackReadinessReport({
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(buildMeetingAppTrackPipeline([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).speaker_track.mark_count, 1);
assert.equal(kit.meetingAppTrackPipeline([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).marks[0].intent, 'speaker_track');
const trackRuntime = createMeetingAppTrackRuntime(client, {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
});
assert.equal((await trackRuntime.observe([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal((await kit.meetingAppTrackRuntime({
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).observe([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal((await createMeetingAppTimelineRuntime(client, {
  trackRuntimeOptions: {
    speakerTrackOptions: {
      minStableMs: 250,
      minSegmentMs: 0,
      closeOpenSegmentsAtMs: 1783356001000,
    },
  },
}).observeMeetingAppTracks([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal(resolveMeetingAppRuntimeAdapterProfile('https://meet.google.com/abc-defg-hij').platform, 'google_meet');
assert.equal(buildMeetingAppAdapterFitReport({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).accepted, true);
assert.equal(buildMeetingAppAdapterFitMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_adapter_fit_matrix');
assert.equal(buildMeetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).schema, 'meeting_app_runtime_observer_plan');
assert.equal(buildMeetingAppRuntimeObserverPlanMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_runtime_observer_plan_matrix');
assert.equal(buildMeetingAppAdapterManifest('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterManifest('google-meet').contracts.timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingAppAdapterManifestMatrix({ platforms: ['google-meet', 'zoom'] }).accepted_count, 2);
assert.equal(assertMeetingAppAdapterManifest('zoom').runtime.observe_mutations, true);
assert.equal(assertMeetingAppAdapterManifestMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.meetingAppAdapterManifest('google-meet').capture.selector_counts.participant > 0, true);
assert.equal(kit.meetingAppAdapterManifestMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_manifest_matrix');
assert.equal(buildMeetingAppAdapterSpec('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterSpecTemplate({ adapter_key: 'whereby' }).adapter_key, 'whereby');
assert.equal(buildMeetingAppAdapterSpecMatrix({
  platforms: ['google-meet'],
  adapters: [buildMeetingAppAdapterSpecTemplate({ adapter_key: 'whereby', matches: ['https://whereby.com/*'] })],
}).accepted_count, 2);
assert.equal(assertMeetingAppAdapterSpec('zoom').adapter_key, 'zoom');
assert.equal(assertMeetingAppAdapterSpecMatrix({ platforms: ['google-meet'] }).spec_count, 1);
assert.equal(kit.meetingAppAdapterSpec('google-meet').source, 'built_in_manifest');
assert.equal(kit.meetingAppAdapterSpecMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_spec_matrix');
assert.equal(buildMeetingAppAdapterRuntimeConfig('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterRuntimeConfig('google-meet').browser_runtime_options.runtimePreset, false);
assert.equal(buildMeetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet', 'zoom'] }).accepted_count, 2);
assert.equal(assertMeetingAppAdapterRuntimeConfig('zoom').contracts.timestamp_field, 'captured_at_ms');
assert.equal(assertMeetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet'] }).config_count, 1);
assert.equal(kit.meetingAppAdapterRuntimeConfig('google-meet').capture_options.participantSelectors.length > 0, true);
assert.equal(kit.meetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_runtime_config_matrix');
assert.equal(buildMeetingAppObserverSchedulerConfig('google-meet').schema, 'meeting_app_observer_scheduler_config');
assert.equal(buildMeetingAppObserverSchedulerConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_observer_scheduler_config_matrix');
assert.equal(typeof createMeetingAppObserverScheduler({
  async sample() {
    return { ok: true };
  },
}, 'google-meet').handleTrigger, 'function');
assert.equal(buildMeetingPlatformRuntimeHostConfig('google-meet').schema, 'meeting_platform_runtime_host_config');
assert.equal(buildMeetingPlatformRuntimeHostConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_platform_runtime_host_config_matrix');
assert.equal(buildMeetingPlatformRuntimeHostHandoff('google-meet').schema, 'meeting_platform_runtime_host_handoff');
assert.equal(buildMeetingPlatformRuntimeHostHandoffMatrix({ platforms: ['google-meet'] }).schema, 'meeting_platform_runtime_host_handoff_matrix');
assert.equal(createMeetingPlatformRuntimeHostFixtureEnvironment('google-meet').platform, 'google_meet');
assert.equal(createMeetingPlatformRuntimeHostVerificationClient().getState().call_count, 0);
assert.equal((await runMeetingPlatformRuntimeHostVerification('google-meet')).accepted, true);
assert.equal((await runMeetingPlatformRuntimeHostVerificationMatrix({ platforms: ['google-meet', 'zoom'] })).accepted_count, 2);
const replayInput = {
  records: [
    {
      platform: 'google_meet',
      phase: 'active',
      captured_at_ms: 1_783_356_000_000,
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
    {
      platform: 'google_meet',
      phase: 'ended',
      captured_at_ms: 1_783_356_002_000,
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Join now' }],
      },
    },
  ],
};
assert.equal((await runMeetingPlatformRuntimeHostReplay('google-meet', replayInput)).accepted, true);
assert.equal((await runMeetingPlatformRuntimeHostReplayMatrix({ platforms: ['google-meet'], input: replayInput })).accepted_count, 1);
assert.equal(typeof createMeetingPlatformRuntimeHost({
  async sample() {
    return { ok: true };
  },
}, 'google-meet', {
  setInterval() {
    return {};
  },
  clearInterval() {},
}).changed, 'function');
assert.equal(resolveMeetingAppRuntimeAdapterProfile('https://example.com/not-a-meeting').detected, false);
assert.equal(selectMeetingAppRuntimeAdapter('https://meet.google.com/abc-defg-hij').selected, true);
assert.equal(buildMeetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').surface, 'browser_extension');
assert.equal(buildMeetingAppRuntimeAdapterHandoffMatrix({ platforms: ['google-meet'], surfaces: ['native-detector'] }).ready_count, 1);
assert.equal(buildMeetingAppRuntimeAdapterHostPackage({ platforms: ['google-meet'], surfaces: ['browser-extension'] }).schema, 'meeting_app_runtime_adapter_host_package');
assert.equal(buildMeetingAppRuntimeAdapterHandoffAcceptanceReport('https://meet.google.com/abc-defg-hij').accepted, true);
assert.equal(buildMeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport({ platforms: ['google-meet'], surfaces: ['native-detector'] }).accepted, true);
assert.equal(assertMeetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').accepted, true);
assert.equal(buildMeetingAppRuntimeAdapterProfileMatrix({ platforms: ['google-meet', 'teams'] }).platform_count, 2);
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
assert.equal(kit.platformRuntimeHostConfig('google-meet').schema, 'meeting_platform_runtime_host_config');
assert.equal(kit.platformRuntimeHostConfigMatrix({
  platforms: ['google-meet'],
}).host_ready_count, 1);
assert.equal(kit.platformRuntimeHostHandoff('google-meet').acceptance.accepted, true);
assert.equal(kit.platformRuntimeHostHandoffMatrix({
  platforms: ['google-meet'],
}).accepted_count, 1);
assert.equal((await kit.replayPlatformRuntimeHost('google-meet', replayInput)).accepted, true);
assert.equal((await kit.replayPlatformRuntimeHostMatrix({
  platforms: ['google-meet'],
  input: replayInput,
})).accepted_count, 1);
assert.equal(kit.platformRuntimeEventPlan('google-meet').schema, 'meeting_platform_runtime_event_plan');
assert.equal(kit.platformRuntimeEventPlan('google-meet').realtime_contract.provider_events_required_for_realtime, false);
assert.equal(kit.platformRuntimeEventPlanMatrix({
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformRuntimeEventPlan('zoom', {
  baseUrl: 'http://localhost:8787',
}).examples.insert_annotation.action, 'insert_annotation');
assert.equal(buildMeetingPlatformRuntimeEventPlanMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).transcript_realtime_dependency_count, 0);
const smokeGoogleRuntimeBundle = buildMeetingPlatformRuntimeBundle('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(smokeGoogleRuntimeBundle.browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(smokeGoogleRuntimeBundle.adapter_route.routes[0].route, 'local_observer_axis');
assert.equal(smokeGoogleRuntimeBundle.runtime.content_script_bridge.install_function, 'installMeetingPlatformIntegrationContentScriptBridge');
assert.equal(smokeGoogleRuntimeBundle.messaging.runtime_event.plan.realtime_contract.transcript_required_for_realtime, false);
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.content_script_insert_annotation.type, 'meeting_timeline.insert_mark');
assert.equal(buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].adapter_first_route, 'local_observer_axis');
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
}).adapter_route_matrix.rows[0].first_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.platform_conformance, '/api/meeting-platform/conformance');
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
}).endpoints.integration_runtime_run_manifest, '/api/meeting-platform/integration-runtime/run-manifest');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.handoff_readiness, '/api/meeting-platform/handoff-readiness');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.runtime_events, '/api/meeting-platform/runtime-events');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_platform_conformance, 'npm run meeting-platform:conformance');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_integration_runtime_manifest, 'npm run meeting-platform:integration-runtime-run-manifest');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_handoff_readiness, 'npm run meeting-platform:handoff-readiness');
const smokeHostScaffold = buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
});
assert.equal(assertMeetingPlatformHostIntegrationScaffold(buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
})).accepted, true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/run-integration-runtime-manifest.mjs'), true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/verify-conformance.mjs'), true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/run-handoff-readiness.mjs'), true);
assert.equal(smokeHostScaffold.files.find((file) => file.path === 'README.md').content.includes('runIntegrationRuntimeManifest'), true);
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
assert.equal(buildMeetingPlatformAdapterRoute('google-meet', {
  baseUrl: 'http://localhost:8787',
}).routes[0].route, 'local_observer_axis');
assert.equal(buildMeetingPlatformAdapterRouteMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformAdapterRoute('google-meet').realtime_invariants.provider_events_block_realtime, false);
assert.equal(kit.platformAdapterRouteMatrix({ platforms: ['zoom'] }).rows[0].first_route, 'local_observer_axis');
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
