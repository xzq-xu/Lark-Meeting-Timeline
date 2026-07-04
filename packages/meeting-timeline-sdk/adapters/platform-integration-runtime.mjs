import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import {
  createMeetingAppBrowserRuntime,
  meetingAppBrowserInput,
  meetingAppBrowserRuntimePreset,
} from './meeting-app-browser-runtime.mjs';
import { createMeetingAppContentScriptBridge } from './meeting-app-content-script.mjs';
import { createMeetingPlatformTimelineKit } from './platform-kit.mjs';
import {
  buildMeetingPlatformAdaptationPackageMatrix,
} from './platform-adaptation-package.mjs';
import {
  buildMeetingPlatformHandoffReadinessMatrix,
} from './platform-handoff-readiness.mjs';
import {
  buildMeetingPlatformLiveAdapterMatrix,
} from './platform-live-adapter.mjs';
import {
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryManifest,
} from './platform-registry.mjs';
import {
  buildMeetingPlatformRuntimeBundleMatrix,
} from './platform-runtime-bundle.mjs';

export const MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA = 'meeting_platform_integration_runtime';
export const MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA = 'meeting_platform_integration_runtime_manifest';
export const MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function withDefaults(defaults = {}, options = {}) {
  return compactObject({
    ...defaults,
    ...options,
    clientOptions: {
      ...(defaults.clientOptions ?? defaults.client_options ?? {}),
      ...(options.clientOptions ?? options.client_options ?? {}),
    },
    env: {
      ...(defaults.env ?? {}),
      ...(options.env ?? {}),
    },
  });
}

function runtimeBaseOptions(options = {}) {
  return compactObject({
    ...options,
    baseUrl: firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787'),
    basePath: firstNonEmpty(options.basePath, options.base_path, '/api/platform-events'),
    platforms: selectedPlatforms(options),
  });
}

function byPlatform(rows = []) {
  return Object.fromEntries(asArray(rows).map((row) => [row.platform, row]));
}

function runtimeIssues(manifest = {}) {
  const issues = [];
  if (manifest.registry_acceptance?.accepted !== true) {
    issues.push({
      severity: 'error',
      code: 'registry_not_accepted',
      message: 'Platform registry must pass SDK acceptance before a reusable runtime can be handed off.',
    });
  }
  if (manifest.runtime_bundle_matrix?.runtime_ready_count !== manifest.platform_count) {
    issues.push({
      severity: 'error',
      code: 'runtime_bundle_not_ready',
      message: 'Every selected platform must have a runtime bundle.',
    });
  }
  if (manifest.runtime_bundle_matrix?.sdk_wiring_ready_count !== manifest.platform_count) {
    issues.push({
      severity: 'error',
      code: 'runtime_sdk_wiring_not_ready',
      message: 'Every selected platform runtime bundle must expose SDK wiring.',
    });
  }
  if ((manifest.runtime_bundle_matrix?.provider_required_for_realtime_count ?? 0) > 0) {
    issues.push({
      severity: 'error',
      code: 'provider_blocks_realtime',
      message: 'Provider events must not be required for realtime annotation insertion.',
    });
  }
  if ((manifest.runtime_bundle_matrix?.transcript_blocking_count ?? 0) > 0) {
    issues.push({
      severity: 'error',
      code: 'transcript_blocks_realtime',
      message: 'Transcript import must not block realtime annotation insertion.',
    });
  }
  if (manifest.adaptation_package_matrix?.sdk_wiring_ready_count !== manifest.platform_count) {
    issues.push({
      severity: 'error',
      code: 'adaptation_package_sdk_wiring_not_ready',
      message: 'Every selected platform must have an adaptation package with SDK wiring.',
    });
  }
  if ((manifest.handoff_readiness_matrix?.pilot_ready_count ?? 0) < manifest.platform_count) {
    issues.push({
      severity: 'warning',
      code: 'real_evidence_not_ready',
      message: 'One or more platforms still need real DOM/provider evidence before pilot or production rollout.',
    });
  }
  return issues;
}

function runtimeAccepted(issues = []) {
  return issues.every((item) => item.severity !== 'error');
}

function actionName(input = {}, options = {}) {
  return String(firstNonEmpty(
    options.action,
    options.kind,
    input.action,
    input.kind,
    input.event_kind,
    input.eventKind,
    input.type,
  ) ?? '').toLowerCase();
}

function platformFrom(input = {}, options = {}) {
  const platform = firstNonEmpty(options.platform, input.platform, input.provider, input.adapter);
  if (!platform) {
    throw new MeetingTimelineSdkError('platform is required for meeting platform runtime events');
  }
  return normalizeMeetingPlatform(platform);
}

function snapshotFrom(input = {}) {
  return firstNonEmpty(input.snapshot, input.meeting_app_snapshot, input.meetingAppSnapshot, input.dom, input);
}

function providerInputFrom(input = {}, payload) {
  return firstNonEmpty(input.request, input.provider_event, input.providerEvent, input.event, input.raw, input, payload);
}

function runtimeEventPayloadFrom(input = {}, payload) {
  return payload === undefined
    ? firstNonEmpty(input.payload, input.provider_event, input.providerEvent, input.snapshot, input.annotation, input.mark, input.raw)
    : payload;
}

function runtimeEventInputFrom(input = {}, payload) {
  const eventPayload = runtimeEventPayloadFrom(input, payload);
  if (!isPlainObject(eventPayload)) {
    return {
      input,
      payload: eventPayload,
    };
  }
  return {
    input: compactObject({
      ...eventPayload,
      ...input,
      payload: input.payload ?? eventPayload,
      snapshot: firstNonEmpty(
        input.snapshot,
        input.meeting_app_snapshot,
        input.meetingAppSnapshot,
        eventPayload.snapshot,
        eventPayload.meeting_app_snapshot,
        eventPayload.meetingAppSnapshot,
      ),
      provider_event: firstNonEmpty(
        input.provider_event,
        input.providerEvent,
        eventPayload.provider_event,
        eventPayload.providerEvent,
        input.raw,
        eventPayload.raw,
        eventPayload,
      ),
      annotation: firstNonEmpty(input.annotation, input.mark, eventPayload.annotation, eventPayload.mark),
      current_meeting: firstNonEmpty(
        input.current_meeting,
        input.currentMeeting,
        eventPayload.current_meeting,
        eventPayload.currentMeeting,
      ),
      signals: firstNonEmpty(input.signals, eventPayload.signals),
      meeting: firstNonEmpty(input.meeting, eventPayload.meeting),
      annotations: firstNonEmpty(input.annotations, eventPayload.annotations),
    }),
    payload: eventPayload,
  };
}

function maybeNormalizePlatform(platform) {
  if (platform == null || platform === '') return undefined;
  try {
    return normalizeMeetingPlatform(platform);
  } catch {
    return undefined;
  }
}

function detectionReason(explicit, urlDetected, presetPlatform) {
  if (explicit) return 'explicit';
  if (urlDetected?.platform) return 'url';
  if (presetPlatform) return 'runtime_preset';
  return 'none';
}

function emptyObservationResult(reason, details = {}) {
  return {
    source: 'meeting_app',
    signals: [],
    rawSignals: [],
    raw_signals: [],
    reconciliation: {
      signals: [],
      skipped: [],
      decisions: [],
      state: {},
    },
    results: [],
    diagnostic: {
      source: 'meeting_app',
      raw_signal_count: 0,
      applied_signal_count: 0,
      skipped_count: 0,
      result_count: 0,
      reason,
    },
    ...details,
  };
}

export function detectMeetingPlatformForBrowser(input = {}, options = {}) {
  const browserInput = meetingAppBrowserInput({
    ...options,
    ...input,
  });
  const explicitPlatform = maybeNormalizePlatform(firstNonEmpty(options.platform, input.platform, input.provider, input.adapter));
  const urlDetected = detectMeetingFromUrl({
    ...browserInput,
    ...input,
  });
  const runtimePreset = meetingAppBrowserRuntimePreset({
    ...browserInput,
    ...input,
    ...options,
  });
  const presetPlatform = maybeNormalizePlatform(firstNonEmpty(
    runtimePreset?.captureOptions?.platform,
    runtimePreset?.capture_options?.platform,
  ));
  const platform = explicitPlatform ?? maybeNormalizePlatform(urlDetected?.platform) ?? presetPlatform;
  return compactObject({
    type: 'meeting_platform_browser_detection',
    detected: Boolean(platform),
    platform,
    reason: detectionReason(explicitPlatform, urlDetected, presetPlatform),
    meeting: urlDetected,
    browser: {
      url: browserInput.url,
      title: browserInput.title,
    },
    runtime_preset_platform: presetPlatform,
  });
}

export function buildMeetingPlatformIntegrationRuntimeManifest(options = {}) {
  const merged = runtimeBaseOptions(options);
  const platforms = selectedPlatforms(merged);
  const registryManifest = buildMeetingPlatformRegistryManifest({
    ...merged,
    platforms,
  });
  const registryAcceptance = buildMeetingPlatformRegistryAcceptanceReport(registryManifest, merged);
  const runtimeBundleMatrix = buildMeetingPlatformRuntimeBundleMatrix({
    ...merged,
    platforms,
  });
  const adaptationPackageMatrix = buildMeetingPlatformAdaptationPackageMatrix({
    ...merged,
    platforms,
  });
  const liveAdapterMatrix = buildMeetingPlatformLiveAdapterMatrix({
    ...merged,
    platforms,
  });
  const handoffReadinessMatrix = buildMeetingPlatformHandoffReadinessMatrix({
    ...merged,
    platforms,
  });
  const runtimeRows = byPlatform(runtimeBundleMatrix.rows);
  const adaptationRows = byPlatform(adaptationPackageMatrix.rows);
  const liveRows = byPlatform(liveAdapterMatrix.rows);
  const handoffRows = byPlatform(handoffReadinessMatrix.rows);
  const rows = platforms.map((platform) => compactObject({
    platform,
    display_name: runtimeRows[platform]?.display_name ?? adaptationRows[platform]?.display_name,
    runtime_ready: runtimeRows[platform]?.runtime_ready === true,
    sdk_wiring_ready: runtimeRows[platform]?.sdk_wiring_ready === true && adaptationRows[platform]?.sdk_wiring_ready === true,
    browser_match_count: runtimeRows[platform]?.browser_match_count,
    recommended_mode: liveRows[platform]?.recommended_mode ?? adaptationRows[platform]?.recommended_mode,
    provider_required_for_realtime: runtimeRows[platform]?.provider_required_for_realtime === true,
    transcript_blocks_realtime: runtimeRows[platform]?.transcript_blocks_realtime === true,
    speaker_min_stable_ms: runtimeRows[platform]?.speaker_min_stable_ms,
    handoff_ready: handoffRows[platform]?.handoff_ready === true,
    pilot_ready: handoffRows[platform]?.pilot_ready === true,
    production_ready: handoffRows[platform]?.production_ready === true,
    first_next_action: handoffRows[platform]?.first_next_action ?? adaptationRows[platform]?.first_next_action,
  }));
  const partial = {
    type: 'meeting_platform_integration_runtime_manifest',
    schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
    base_url: merged.baseUrl,
    base_path: merged.basePath,
    platform_count: platforms.length,
    platforms,
    registry_acceptance: registryAcceptance,
    runtime_bundle_matrix: runtimeBundleMatrix,
    adaptation_package_matrix: adaptationPackageMatrix,
    live_adapter_matrix: liveAdapterMatrix,
    handoff_readiness_matrix: handoffReadinessMatrix,
    rows,
  };
  const issues = runtimeIssues(partial);
  return {
    ...partial,
    host_integration_ready: runtimeAccepted(issues),
    blocking_count: issues.filter((item) => item.severity === 'error').length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: unique([
      ...issues.map((item) => item.code),
      ...handoffReadinessMatrix.rows.flatMap((row) => row.next_actions ?? []),
    ]),
  };
}

export function assertMeetingPlatformIntegrationRuntimeManifest(manifestOrOptions = {}, options = {}) {
  const manifest = manifestOrOptions?.schema === MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA
    ? manifestOrOptions
    : buildMeetingPlatformIntegrationRuntimeManifest({
      ...manifestOrOptions,
      ...options,
    });
  if (manifest.host_integration_ready !== true) {
    throw new MeetingTimelineSdkError('Meeting platform integration runtime is not ready for host handoff', {
      issues: manifest.issues,
      manifest,
    });
  }
  return manifest;
}

export function createMeetingPlatformIntegrationRuntime(clientOrOptions, options = {}) {
  const defaults = runtimeBaseOptions({
    ...(clientOrOptions && typeof clientOrOptions === 'object' && typeof clientOrOptions.startMeeting !== 'function' ? clientOrOptions : {}),
    ...options,
  });
  const kit = createMeetingPlatformTimelineKit(clientOrOptions, defaults);
  const platforms = selectedPlatforms(defaults);
  const suite = kit.platformLiveAdapterSuite({ platforms });

  function adapter(platform, adapterOptions = {}) {
    return suite.adapter(normalizeMeetingPlatform(platform), adapterOptions);
  }

  const runtime = {
    type: 'meeting_platform_integration_runtime',
    schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
    schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
    platforms,
    client: kit.client,
    kit,
    liveAdapters: suite,
    live_adapters: suite,
    adapter,
    platformAdapter: adapter,
    manifest(manifestOptions = {}) {
      return buildMeetingPlatformIntegrationRuntimeManifest(withDefaults(defaults, manifestOptions));
    },
    assertManifest(manifestOptions = {}) {
      return assertMeetingPlatformIntegrationRuntimeManifest(withDefaults(defaults, manifestOptions));
    },
    registry(registryOptions = {}) {
      const manifest = kit.platformRegistryManifest({
        ...registryOptions,
        platforms: registryOptions.platforms ?? registryOptions.platform_keys ?? platforms,
      });
      return {
        manifest,
        acceptance: kit.platformRegistryAcceptance(manifest, registryOptions),
      };
    },
    runtimeBundle(platform, bundleOptions = {}) {
      return kit.platformRuntimeBundle(platform, bundleOptions);
    },
    runtimeBundles(bundleOptions = {}) {
      return kit.platformRuntimeBundleMatrix({
        ...bundleOptions,
        platforms: bundleOptions.platforms ?? bundleOptions.platform_keys ?? platforms,
      });
    },
    adaptationPackages(packageOptions = {}) {
      return kit.platformAdaptationPackageMatrix({
        ...packageOptions,
        platforms: packageOptions.platforms ?? packageOptions.platform_keys ?? platforms,
      });
    },
    readiness(readinessOptions = {}) {
      return suite.readinessMatrix({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    handoffReadiness(readinessOptions = {}) {
      return kit.platformHandoffReadinessMatrix({
        ...readinessOptions,
        platforms: readinessOptions.platforms ?? readinessOptions.platform_keys ?? platforms,
      });
    },
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      return adapter(platform, observeOptions.adapterOptions ?? observeOptions.adapter_options ?? {})
        .observeMeetingApp(snapshot, observeOptions);
    },
    observeApp(platform, snapshot = {}, observeOptions = {}) {
      return runtime.observeMeetingApp(platform, snapshot, observeOptions);
    },
    ingestProvider(platform, requestOrPayload = {}, payload, ingestOptions = {}) {
      const providerPayload = payload === undefined ? requestOrPayload : payload;
      return adapter(platform, ingestOptions.adapterOptions ?? ingestOptions.adapter_options ?? {})
        .ingestProvider(requestOrPayload, providerPayload, ingestOptions);
    },
    insertAnnotation(platform, input = {}, markOptions = {}) {
      return adapter(platform, markOptions.adapterOptions ?? markOptions.adapter_options ?? {})
        .insertAnnotation(input, markOptions);
    },
    insertMark(platform, input = {}, markOptions = {}) {
      return runtime.insertAnnotation(platform, input, markOptions);
    },
    speakerTrack(platform, input = {}, trackOptions = {}) {
      return kit.platformSpeakerTrack(platform, input, trackOptions);
    },
    participantTrack(platform, input = {}, trackOptions = {}) {
      return kit.platformParticipantTrack(platform, input, trackOptions);
    },
    timelineView(platform, input = {}, viewOptions = {}) {
      return kit.platformTimelineView(platform, input, viewOptions);
    },
    async handleEvent(input = {}, payload, eventOptions = {}) {
      const runtimeEvent = runtimeEventInputFrom(input, payload);
      const eventInput = runtimeEvent.input;
      const eventPayload = runtimeEvent.payload;
      const action = actionName(eventInput, eventOptions);
      if (['registry', 'platform_registry'].includes(action)) return runtime.registry(eventOptions);
      if (['manifest', 'runtime_manifest', 'integration_manifest'].includes(action)) return runtime.manifest(eventOptions);
      if (['runtime_bundles', 'runtime_bundle_matrix'].includes(action)) return runtime.runtimeBundles(eventOptions);
      if (['readiness', 'live_readiness'].includes(action)) return runtime.readiness(eventOptions);
      if (['handoff_readiness'].includes(action)) return runtime.handoffReadiness(eventOptions);
      const platform = platformFrom(eventInput, eventOptions);
      if (['observe', 'observe_app', 'observe_meeting_app', 'meeting_app_snapshot', 'snapshot'].includes(action)) {
        return runtime.observeMeetingApp(platform, snapshotFrom(eventInput), eventOptions);
      }
      if (['provider', 'provider_event', 'ingest_provider', 'webhook'].includes(action)) {
        return runtime.ingestProvider(platform, providerInputFrom(eventInput, eventPayload), eventPayload, eventOptions);
      }
      if (['annotation', 'mark', 'insert_annotation', 'insert_mark'].includes(action)) {
        return runtime.insertAnnotation(platform, eventInput, eventOptions);
      }
      if (['speaker_track', 'speaker'].includes(action)) {
        return runtime.speakerTrack(platform, eventInput, eventOptions);
      }
      if (['participant_track', 'participant'].includes(action)) {
        return runtime.participantTrack(platform, eventInput, eventOptions);
      }
      if (['timeline_view', 'view'].includes(action)) {
        return runtime.timelineView(platform, eventInput, eventOptions);
      }
      throw new MeetingTimelineSdkError('unsupported meeting platform runtime event action', {
        action,
        supported_actions: [
          'observe_meeting_app',
          'ingest_provider',
          'insert_annotation',
          'speaker_track',
          'participant_track',
          'timeline_view',
          'runtime_bundles',
          'registry',
          'manifest',
          'readiness',
          'handoff_readiness',
        ],
      });
    },
    summary(summaryOptions = {}) {
      const manifest = runtime.manifest(summaryOptions);
      return compactObject({
        type: 'meeting_platform_integration_runtime_summary',
        schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
        schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
        platform_count: manifest.platform_count,
        platforms: manifest.platforms,
        host_integration_ready: manifest.host_integration_ready,
        blocking_count: manifest.blocking_count,
        warning_count: manifest.warning_count,
        handoff_ready_count: manifest.handoff_readiness_matrix?.handoff_ready_count,
        pilot_ready_count: manifest.handoff_readiness_matrix?.pilot_ready_count,
        production_ready_count: manifest.handoff_readiness_matrix?.production_ready_count,
        next_actions: manifest.next_actions,
      });
    },
    getState() {
      return {
        type: 'meeting_platform_integration_runtime_state',
        schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
        schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
        platforms,
        kit: kit.getState(),
        live_adapters: suite.getState(),
      };
    },
    reset(nextState = {}) {
      return {
        kit: kit.reset(nextState.kit ?? {}),
        live_adapters: suite.reset(nextState.live_adapters ?? nextState.liveAdapters ?? {}),
      };
    },
  };
  return runtime;
}

export function createMeetingPlatformIntegrationBrowserRuntime(clientOrOptions, options = {}) {
  const integrationRuntime = options.integrationRuntime
    ?? options.integration_runtime
    ?? createMeetingPlatformIntegrationRuntime(clientOrOptions, options);
  const baseInputOptions = {
    window: options.window,
    win: options.win,
    document: options.document,
    doc: options.doc,
    location: options.location,
    navigator: options.navigator,
  };

  function detect(input = {}, detectOptions = {}) {
    return detectMeetingPlatformForBrowser(input, {
      ...options,
      ...baseInputOptions,
      ...detectOptions,
    });
  }

  function platformFor(input = {}, platformOptions = {}) {
    const detection = detect(input, platformOptions);
    if (!detection.platform) {
      throw new MeetingTimelineSdkError('Unable to detect meeting platform for browser runtime event', {
        detection,
      });
    }
    return detection.platform;
  }

  const sources = {
    client: integrationRuntime.client,
    async observeMeetingApp(input = {}, observeOptions = {}) {
      const detection = detect(input, observeOptions);
      if (!detection.platform) {
        return emptyObservationResult('unsupported_browser_page', {
          browser_detection: detection,
        });
      }
      const result = await integrationRuntime.observeMeetingApp(detection.platform, {
        ...input,
        platform: detection.platform,
      }, observeOptions);
      return {
        ...result,
        browser_detection: detection,
      };
    },
    observeApp(input = {}, observeOptions = {}) {
      return this.observeMeetingApp(input, observeOptions);
    },
    async insertMark(input = {}, markOptions = {}) {
      const platform = platformFor(input, markOptions);
      return integrationRuntime.insertAnnotation(platform, input, markOptions);
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      return this.insertMark(input, markOptions);
    },
    async insertMarks(inputs = [], markOptions = {}) {
      const rows = Array.isArray(inputs) ? inputs : [inputs];
      const results = [];
      for (const input of rows) {
        results.push(await this.insertMark(input, markOptions));
      }
      return {
        ok: results.every((item) => item?.result?.ok !== false),
        results,
      };
    },
    async ingestProvider(platformOrInput, payload, ingestOptions = {}) {
      const platform = maybeNormalizePlatform(typeof platformOrInput === 'string'
        ? platformOrInput
        : firstNonEmpty(platformOrInput?.platform, platformOrInput?.provider, ingestOptions.platform))
        ?? platformFor(platformOrInput, ingestOptions);
      return integrationRuntime.ingestProvider(platform, platformOrInput, payload, ingestOptions);
    },
    ingestSignals(signals = [], ingestOptions = {}) {
      return integrationRuntime.kit.ingest('local-detector', signals, ingestOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return integrationRuntime.kit.importTranscript(input, transcriptOptions);
    },
    startMeeting(input = {}) {
      return integrationRuntime.kit.startMeeting(input);
    },
    endMeeting(input = {}) {
      return integrationRuntime.kit.endMeeting(input);
    },
    getState() {
      return integrationRuntime.getState();
    },
    reset(nextState = {}) {
      return integrationRuntime.reset(nextState);
    },
  };

  const browserRuntime = createMeetingAppBrowserRuntime({}, {
    ...options,
    sources,
  });

  return {
    ...browserRuntime,
    type: 'meeting_platform_integration_browser_runtime',
    schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
    schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
    integrationRuntime,
    integration_runtime: integrationRuntime,
    detect,
    platformFor,
    observeMeetingApp(input = {}, observeOptions = {}) {
      const detection = detect(input, observeOptions);
      if (!detection.platform) {
        return Promise.resolve(emptyObservationResult('unsupported_browser_page', {
          browser_detection: detection,
        }));
      }
      return integrationRuntime.observeMeetingApp(detection.platform, {
        ...input,
        platform: detection.platform,
      }, observeOptions);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return sources.insertMark(input, markOptions);
    },
    getState() {
      return {
        ...browserRuntime.getState(),
        integration_runtime: integrationRuntime.getState(),
        browser_detection: detect(),
      };
    },
    reset(nextState = {}) {
      return {
        browser_runtime: browserRuntime.reset?.(nextState.browser_runtime ?? nextState.browserRuntime ?? {}),
        integration_runtime: integrationRuntime.reset(nextState.integration_runtime ?? nextState.integrationRuntime ?? {}),
      };
    },
  };
}

export function createMeetingPlatformIntegrationContentScriptBridge(clientOrOptions, options = {}) {
  const runtime = options.runtime
    ?? options.browserRuntime
    ?? options.browser_runtime
    ?? createMeetingPlatformIntegrationBrowserRuntime(clientOrOptions, options);
  const bridge = createMeetingAppContentScriptBridge(clientOrOptions, {
    ...options,
    runtime,
  });
  const integrationRuntime = runtime.integrationRuntime ?? runtime.integration_runtime;

  return {
    ...bridge,
    type: 'meeting_platform_integration_content_script_bridge',
    schema: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
    schema_version: MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA_VERSION,
    runtime,
    integrationRuntime,
    integration_runtime: integrationRuntime,
    detect(input = {}, detectOptions = {}) {
      return runtime.detect?.(input, detectOptions);
    },
    getState() {
      return {
        ...bridge.getState(),
        integration_runtime: integrationRuntime?.getState?.(),
        browser_detection: runtime.detect?.(),
      };
    },
    dispose() {
      return bridge.dispose();
    },
  };
}

export function installMeetingPlatformIntegrationContentScriptBridge(clientOrOptions, options = {}) {
  const bridge = createMeetingPlatformIntegrationContentScriptBridge(clientOrOptions, options);
  bridge.start(options.startOptions ?? options.start_options ?? {});
  return bridge;
}
