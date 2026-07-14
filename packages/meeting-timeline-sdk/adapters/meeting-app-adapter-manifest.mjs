import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_APP_BROWSER_RUNTIME_PRESETS,
  meetingAppBrowserRuntimePreset,
} from './meeting-app-browser-runtime.mjs';
import {
  meetingAppDomCaptureProfile,
} from './meeting-app-capture.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
  buildMeetingAppContentScriptManifest,
  buildMeetingAppExtensionInstallPlan,
} from './meeting-app-extension.mjs';
import {
  buildMeetingAppIntegrationProfile,
  buildMeetingAppRuntimeAdapterAcceptanceReport,
  buildMeetingAppRuntimeAdapterConfig,
  buildMeetingAppRuntimeAdapterHandoff,
  buildMeetingAppRuntimeObserverPlan,
} from './meeting-app-profile.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_APP_ADAPTER_MANIFEST_SCHEMA = 'meeting_app_adapter_manifest';
export const MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA = 'meeting_app_adapter_manifest_matrix';
export const MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION = 1;

const DEFAULT_APP_ADAPTER_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_APP_ADAPTER_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)))
    .filter((platform) => platform !== 'local_detector');
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function selectorCounts(capture = {}, runtime = {}) {
  return {
    control: asArray(capture.controlSelectors ?? capture.control_selectors).length,
    participant: asArray(capture.participantSelectors ?? capture.participant_selectors).length,
    text: asArray(capture.textSelectors ?? capture.text_selectors).length,
    mutation_track: asArray(runtime.mutationTrackSelectors ?? runtime.mutation_track_selectors).length,
    mutation_ignore: asArray(runtime.mutationIgnoreSelectors ?? runtime.mutation_ignore_selectors).length,
  };
}

function manifestCommands(platform, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const baseUrlArg = baseUrl ? ` --base-url=${baseUrl}` : '';
  return {
    print_manifest: `npm run meeting-app:adapter-manifest -- --platforms=${platform}${baseUrlArg}`,
    print_extension: `npm run meeting-app:extension -- --platforms=${platform}`,
    verify_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${platform}${baseUrlArg}`,
    run_runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${platform}${baseUrlArg}`,
  };
}

function staticIssues({
  platform,
  config,
  acceptance,
  observerPlan,
  handoff,
  captureProfile,
  runtimePreset,
  counts,
} = {}) {
  return [
    acceptance.accepted === true
      ? undefined
      : issue('error', 'runtime_adapter_acceptance_failed', 'Runtime adapter config must pass static acceptance.', {
        platform,
        issue_count: acceptance.issues?.length,
      }),
    observerPlan.sdk_ready === true
      ? undefined
      : issue('error', 'observer_plan_not_sdk_ready', 'Observer plan must be SDK-ready.', { platform }),
    asArray(config.extension?.host_permissions).includes('<all_urls>')
      ? issue('error', 'overbroad_host_permission', 'Adapter manifest must not require <all_urls>.', { platform })
      : undefined,
    asArray(config.extension?.permissions).includes('tabs')
      ? undefined
      : issue('error', 'missing_tabs_permission', 'Adapter manifest needs tabs for candidate observation.', { platform }),
    asArray(config.extension?.permissions).includes('storage')
      ? undefined
      : issue('error', 'missing_storage_permission', 'Adapter manifest needs storage for diagnostics/status.', { platform }),
    config.extension?.message_types?.observe_candidates === MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates
      ? undefined
      : issue('error', 'missing_observe_candidates_message', 'Adapter manifest must expose observe_candidates.', { platform }),
    runtimePreset?.observeMutations === true
      ? undefined
      : issue('error', 'mutation_observer_disabled', 'Runtime preset must enable mutation observation.', { platform }),
    counts.control > 0
      ? undefined
      : issue('error', 'missing_control_selectors', 'Capture profile must include meeting control selectors.', { platform }),
    counts.participant > 0
      ? undefined
      : issue('error', 'missing_participant_selectors', 'Capture profile must include participant/speaker selectors.', { platform }),
    counts.text > 0
      ? undefined
      : issue('warning', 'missing_text_selectors', 'Capture profile should include title/status text selectors.', { platform }),
    captureProfile?.platform === platform
      ? undefined
      : issue('error', 'capture_profile_mismatch', 'Capture profile platform must match manifest platform.', { platform }),
    runtimePreset?.platform === platform
      ? undefined
      : issue('error', 'runtime_preset_mismatch', 'Runtime preset platform must match manifest platform.', { platform }),
    handoff.readiness?.ready_to_start === true || handoff.readiness?.runtime_ready === true
      ? undefined
      : issue('warning', 'runtime_handoff_needs_live_validation', 'Runtime handoff should be validated with live snapshots before production.', { platform }),
  ].filter(Boolean);
}

export function buildMeetingAppAdapterManifest(platformOrOptions = {}, options = {}) {
  const objectInput = platformOrOptions && typeof platformOrOptions === 'object' && !Array.isArray(platformOrOptions);
  const rawPlatform = objectInput
    ? firstNonEmpty(platformOrOptions.platform, platformOrOptions.provider, platformOrOptions.key, platformOrOptions.name)
    : platformOrOptions;
  const merged = objectInput ? { ...platformOrOptions, ...options } : { ...options, platform: rawPlatform };
  const platform = normalizeMeetingPlatform(firstNonEmpty(rawPlatform, merged.platform));
  if (platform === 'local_detector') {
    throw new MeetingTimelineSdkError('meeting-app adapter manifest requires a meeting app platform', {
      code: 'local_detector_not_supported_for_app_adapter_manifest',
      platform,
    });
  }
  const integrationProfile = buildMeetingAppIntegrationProfile(platform, merged);
  const runtimeConfig = buildMeetingAppRuntimeAdapterConfig(platform, merged);
  const acceptance = buildMeetingAppRuntimeAdapterAcceptanceReport(runtimeConfig, merged);
  const observerPlan = buildMeetingAppRuntimeObserverPlan(platform, merged);
  const handoff = buildMeetingAppRuntimeAdapterHandoff({ platform }, merged);
  const extensionInstallPlan = buildMeetingAppExtensionInstallPlan({
    ...merged,
    platforms: [platform],
  });
  const contentScriptManifest = buildMeetingAppContentScriptManifest({
    ...merged,
    platforms: [platform],
  });
  const captureProfile = meetingAppDomCaptureProfile(platform, merged);
  const runtimePreset = meetingAppBrowserRuntimePreset(platform, merged) ?? MEETING_APP_BROWSER_RUNTIME_PRESETS[platform];
  const counts = selectorCounts(captureProfile, runtimePreset);
  const issues = staticIssues({
    platform,
    config: runtimeConfig,
    acceptance,
    observerPlan,
    handoff,
    captureProfile,
    runtimePreset,
    counts,
  });
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    type: 'meeting_app_adapter_manifest',
    schema: MEETING_APP_ADAPTER_MANIFEST_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    platform,
    display_name: integrationProfile.display_name,
    recommended_mode: integrationProfile.recommended_mode,
    surface: 'browser_extension_or_webview',
    extension: {
      matches: runtimeConfig.extension?.matches ?? [],
      host_permissions: runtimeConfig.extension?.host_permissions ?? [],
      permissions: runtimeConfig.extension?.permissions ?? [],
      message_types: runtimeConfig.extension?.message_types,
      status_storage_key: runtimeConfig.extension?.status_storage_key,
      content_scripts: contentScriptManifest.content_scripts ?? [],
      install_plan_schema: extensionInstallPlan.schema,
    },
    runtime: {
      preset: runtimePreset?.platform,
      observe_mutations: runtimePreset?.observeMutations === true,
      mutation_debounce_ms: runtimePreset?.mutationDebounceMs,
      speaker_stable_followup_ms: runtimePreset?.speakerStableFollowupMs,
      sample_interval_ms: runtimePreset?.sampleIntervalMs,
      unchanged_observe_every_ms: runtimePreset?.unchangedObserveEveryMs,
      mutation_track_selectors: runtimePreset?.mutationTrackSelectors ?? [],
      mutation_ignore_selectors: runtimePreset?.mutationIgnoreSelectors ?? [],
    },
    capture: {
      profile: captureProfile?.platform,
      display_name: captureProfile?.displayName,
      control_selectors: captureProfile?.controlSelectors ?? [],
      participant_selectors: captureProfile?.participantSelectors ?? [],
      text_selectors: captureProfile?.textSelectors ?? [],
      selector_counts: counts,
    },
    bridge: runtimeConfig.bridge_options,
    observer_plan: {
      schema: observerPlan.schema,
      accepted: observerPlan.accepted,
      sdk_ready: observerPlan.sdk_ready,
      preflight_status: observerPlan.preflight_status,
      cadence: observerPlan.cadence,
      trigger_policy: observerPlan.trigger_policy,
      signal_contract: observerPlan.signal_contract,
    },
    contracts: {
      timestamp_field: 'captured_at_ms',
      candidate_observation_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      required_signals: ['meeting_started', 'meeting_ended', 'speaker_started'],
      output_intents: ['realtime_axis', 'speaker_track', 'participant_track', 'annotation_insert'],
    },
    host_endpoints: MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
    implementation_order: [
      'install_extension_or_webview_script_for_platform_matches',
      'send_extension_attached_and_observe_candidates',
      'start_browser_runtime_with_platform_preset',
      'capture_meeting_start_end_speaker_participant_snapshots',
      'insert_marks_with_captured_at_ms',
      'run_live_snapshot_validation_before_production',
    ],
    readiness: {
      accepted: blocking.length === 0,
      runtime_adapter_accepted: acceptance.accepted === true,
      observer_sdk_ready: observerPlan.sdk_ready === true,
      extension_match_count: asArray(runtimeConfig.extension?.matches).length,
      host_permission_count: asArray(runtimeConfig.extension?.host_permissions).length,
      mutation_track_selector_count: counts.mutation_track,
      participant_selector_count: counts.participant,
      requires_live_snapshot_before_production: true,
    },
    commands: manifestCommands(platform, merged),
    acceptance_report: {
      accepted: acceptance.accepted,
      coverage: acceptance.coverage,
      issue_count: acceptance.issues?.length ?? 0,
    },
    handoff: {
      schema: handoff.schema,
      platform: handoff.platform,
      readiness: handoff.readiness,
      runtime: handoff.runtime,
      tracks: handoff.tracks,
    },
    issue_count: issues.length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: unique([
      ...(blocking.map((item) => item.code)),
      ...(observerPlan.next_actions ?? []),
      'capture_live_dom_snapshots_before_production_rollout',
    ]),
  };
}

export function buildMeetingAppAdapterManifestMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const manifests = platforms.map((platform) => buildMeetingAppAdapterManifest(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_app_adapter_manifest_matrix',
    schema: MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_MANIFEST_SCHEMA_VERSION,
    accepted: manifests.every((manifest) => manifest.accepted === true),
    platform_count: manifests.length,
    accepted_count: manifests.filter((manifest) => manifest.accepted === true).length,
    extension_match_count: manifests.reduce((total, manifest) => total + manifest.readiness.extension_match_count, 0),
    mutation_observer_count: manifests.filter((manifest) => manifest.runtime.observe_mutations === true).length,
    candidate_observer_count: manifests.filter((manifest) => (
      manifest.contracts.candidate_observation_message_type === MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates
    )).length,
    participant_selector_ready_count: manifests.filter((manifest) => manifest.readiness.participant_selector_count > 0).length,
    platforms,
    rows: manifests.map((manifest) => ({
      platform: manifest.platform,
      display_name: manifest.display_name,
      accepted: manifest.accepted,
      surface: manifest.surface,
      extension_match_count: manifest.readiness.extension_match_count,
      host_permission_count: manifest.readiness.host_permission_count,
      mutation_track_selector_count: manifest.readiness.mutation_track_selector_count,
      participant_selector_count: manifest.readiness.participant_selector_count,
      observer_sdk_ready: manifest.readiness.observer_sdk_ready,
      runtime_adapter_accepted: manifest.readiness.runtime_adapter_accepted,
      first_next_action: manifest.next_actions?.[0],
      blocking_count: manifest.blocking_count,
      warning_count: manifest.warning_count,
    })),
    manifests,
    next_actions: unique(manifests.flatMap((manifest) => manifest.next_actions ?? [])),
  };
}

export function assertMeetingAppAdapterManifest(manifestOrPlatform = {}, options = {}) {
  const manifest = manifestOrPlatform?.schema === MEETING_APP_ADAPTER_MANIFEST_SCHEMA
    ? manifestOrPlatform
    : buildMeetingAppAdapterManifest(manifestOrPlatform, options);
  if (manifest.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter manifest acceptance failed', {
      code: 'meeting_app_adapter_manifest_rejected',
      platform: manifest.platform,
      issues: manifest.issues,
      next_actions: manifest.next_actions,
    });
  }
  return manifest;
}

export function assertMeetingAppAdapterManifestMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_MANIFEST_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterManifestMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter manifest matrix acceptance failed', {
      code: 'meeting_app_adapter_manifest_matrix_rejected',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}
