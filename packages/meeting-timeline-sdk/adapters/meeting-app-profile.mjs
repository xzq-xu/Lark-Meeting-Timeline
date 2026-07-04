import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { meetingAppBrowserRuntimePreset } from './meeting-app-browser-runtime.mjs';
import { meetingAppDomCaptureProfile } from './meeting-app-capture.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLaunchGate } from './meeting-app-gate.mjs';
import {
  buildMeetingAppSnapshotRecordSet,
  meetingAppSnapshotRecords,
} from './meeting-app-snapshot-recorder.mjs';
import { normalizeMeetingAppSnapshot } from './meeting-apps.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  MEETING_APP_EXTENSION_STATUS_STORAGE_KEY,
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
  buildMeetingAppExtensionInstallPlan,
  meetingAppExtensionProfile,
} from './meeting-app-extension.mjs';
import {
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';

export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA = 'meeting_app_integration_profile';
export const MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION = 1;
export const MEETING_APP_INTEGRATION_PROFILE_PLATFORMS = MEETING_APP_FIXTURE_PLATFORMS;
export const MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA = 'meeting_app_runtime_adapter_config';
export const MEETING_APP_LIVE_SNAPSHOT_CAPTURE_PLAN_SCHEMA = 'meeting_app_live_snapshot_capture_plan';
export const MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA = 'meeting_app_deployment_manifest';
export const MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA = 'meeting_app_live_evidence_package';
export const MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA = 'meeting_app_dom_adaptation_diagnosis';
export const MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA = 'meeting_app_dom_adaptation_diagnosis_matrix';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function issue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function normalizeAppPlatform(platform) {
  const key = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_INTEGRATION_PROFILE_PLATFORMS.includes(key)) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app integration profile platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
    });
  }
  return key;
}

function platformList(options = {}) {
  const raw = firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    options.platform,
    options.provider,
    MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
  );
  return uniqueList(asArray(raw).map((platform) => normalizeAppPlatform(platform)));
}

function platformListFromEvidence(recordSet = {}, options = {}) {
  const evidencePlatforms = uniqueList([
    ...(recordSet.platforms ?? []),
    ...meetingAppSnapshotRecords(recordSet).map((record) => record.platform ?? record.provider),
  ]);
  const raw = firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    options.platform,
    options.provider,
    evidencePlatforms.length > 0 ? evidencePlatforms : undefined,
    MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
  );
  return uniqueList(asArray(raw).map((platform) => normalizeAppPlatform(platform)));
}

function selectorSummary(captureProfile = {}) {
  return {
    control_selector_count: captureProfile.controlSelectors?.length ?? 0,
    participant_selector_count: captureProfile.participantSelectors?.length ?? 0,
    text_selector_count: captureProfile.textSelectors?.length ?? 0,
  };
}

function runtimeSummary(runtimePreset = {}) {
  return compactObject({
    runtimePreset: runtimePreset.platform,
    captureOptions: runtimePreset.captureOptions,
    observeMutations: runtimePreset.observeMutations,
    mutationDebounceMs: runtimePreset.mutationDebounceMs,
    speakerStableFollowupMs: runtimePreset.speakerStableFollowupMs,
    sampleIntervalMs: runtimePreset.sampleIntervalMs,
    unchangedObserveEveryMs: runtimePreset.unchangedObserveEveryMs,
    mutation_track_selector_count: runtimePreset.mutationTrackSelectors?.length ?? 0,
    mutation_ignore_selector_count: runtimePreset.mutationIgnoreSelectors?.length ?? 0,
  });
}

function runtimeOptions(runtimePreset = {}) {
  return compactObject({
    runtimePreset: runtimePreset.platform,
    runtime_preset: runtimePreset.platform,
    browserRuntimePreset: runtimePreset.platform,
    browser_runtime_preset: runtimePreset.platform,
    captureOptions: runtimePreset.captureOptions,
    capture_options: runtimePreset.capture_options ?? runtimePreset.captureOptions,
    observeMutations: runtimePreset.observeMutations,
    observe_mutations: runtimePreset.observe_mutations,
    mutationDebounceMs: runtimePreset.mutationDebounceMs,
    mutation_debounce_ms: runtimePreset.mutation_debounce_ms,
    speakerStableFollowupMs: runtimePreset.speakerStableFollowupMs,
    speaker_stable_followup_ms: runtimePreset.speaker_stable_followup_ms,
    sampleIntervalMs: runtimePreset.sampleIntervalMs,
    sample_interval_ms: runtimePreset.sample_interval_ms,
    unchangedObserveEveryMs: runtimePreset.unchangedObserveEveryMs,
    unchanged_observe_every_ms: runtimePreset.unchanged_observe_every_ms,
    mutationTrackSelectors: runtimePreset.mutationTrackSelectors,
    mutation_track_selectors: runtimePreset.mutation_track_selectors,
    mutationIgnoreSelectors: runtimePreset.mutationIgnoreSelectors,
    mutation_ignore_selectors: runtimePreset.mutation_ignore_selectors,
  });
}

function captureOptions(captureProfile = {}) {
  return compactObject({
    platform: captureProfile.platform,
    captureProfile: captureProfile.platform,
    capture_profile: captureProfile.platform,
    controlSelectors: captureProfile.controlSelectors,
    control_selectors: captureProfile.controlSelectors,
    participantSelectors: captureProfile.participantSelectors,
    participant_selectors: captureProfile.participantSelectors,
    textSelectors: captureProfile.textSelectors,
    text_selectors: captureProfile.textSelectors,
  });
}

function bridgeOptions(platform, options = {}) {
  return compactObject({
    browser_runtime_preset: platform,
    source: firstNonEmpty(options.source, options.detectorSource, options.detector_source, 'meeting_app_extension'),
    extensionMessaging: firstNonEmpty(options.extensionMessaging, options.extension_messaging, true),
    extension_messaging: firstNonEmpty(options.extensionMessaging, options.extension_messaging, true),
    windowMessaging: firstNonEmpty(options.windowMessaging, options.window_messaging, true),
    window_messaging: firstNonEmpty(options.windowMessaging, options.window_messaging, true),
    startRuntime: firstNonEmpty(options.startRuntime, options.start_runtime, true),
    start_runtime: firstNonEmpty(options.startRuntime, options.start_runtime, true),
  });
}

function runtimeConfigFromInput(configOrPlatform = {}, options = {}) {
  if (configOrPlatform?.type === 'meeting_app_runtime_adapter_config') return configOrPlatform;
  return buildMeetingAppRuntimeAdapterConfig(configOrPlatform, options);
}

function deploymentManifestFromInput(manifestOrPlatform = {}, options = {}) {
  if (manifestOrPlatform?.type === 'meeting_app_deployment_manifest') return manifestOrPlatform;
  return buildMeetingAppDeploymentManifest(manifestOrPlatform, options);
}

function launchGateOptions(options = {}) {
  return {
    allowFixtureProduction: true,
    requireProductionReady: false,
    ...(options.launchGateOptions ?? {}),
    ...(options.launch_gate_options ?? {}),
    ...options,
  };
}

function readinessFor(gate = {}, runtimePreset = {}, captureProfile = {}) {
  const nextActions = [];
  if (gate.production_ready !== true) nextActions.push('capture_live_dom_snapshots_for_this_platform');
  if ((runtimePreset.mutationTrackSelectors?.length ?? 0) === 0) nextActions.push('add_mutation_track_selectors');
  if (!captureProfile?.platform) nextActions.push('add_dom_capture_profile');
  return {
    sdk_wiring_ready: gate.passed === true,
    production_ready: gate.production_ready === true,
    evidence_level: gate.evidence_level,
    runtime_ready: gate.runtime_ready,
    missing_required_coverage: gate.missing_required_coverage ?? [],
    next_actions: uniqueList([
      ...nextActions,
      ...(gate.next_actions ?? []),
    ]),
  };
}

function eventModel(platform, capability = {}, integrationPlan = {}) {
  return {
    realtime_axis: {
      primary: 'browser_extension_local_observer',
      provider_reconciliation: integrationPlan.provider_events?.enabled === true,
      signal_types: ['meeting_started', 'meeting_ended'],
      timestamp_invariant: 'annotations_use_absolute_captured_at_ms',
    },
    speaker_activity: {
      primary: 'local_dom_active_speaker_observer',
      provider_support: capability.speaker_activity?.status,
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: capability.speaker_activity?.fallback,
    },
    participant_track: {
      primary: capability.participant_track?.status?.startsWith('supported')
        ? `${platform}_provider_events`
        : 'local_dom_participant_observer',
      signal_types: ['participant_joined', 'participant_left'],
      provider_status: capability.participant_track?.status,
    },
    post_meeting_transcript: capability.post_meeting_transcript,
  };
}

export function buildMeetingAppIntegrationProfile(platformOrInput = {}, options = {}) {
  const input = typeof platformOrInput === 'string'
    ? { platform: platformOrInput }
    : (platformOrInput ?? {});
  const merged = { ...input, ...options };
  const platform = normalizeAppPlatform(firstNonEmpty(merged.platform, merged.provider, merged.key, merged.name));
  const capability = platformCapabilityContract(platform, merged);
  const integrationPlan = buildPlatformIntegrationPlan(platform, merged);
  const extensionProfile = meetingAppExtensionProfile(platform);
  const extensionPlan = buildMeetingAppExtensionInstallPlan({
    ...merged,
    platforms: [platform],
  });
  const captureProfile = meetingAppDomCaptureProfile(platform, merged);
  const runtimePreset = meetingAppBrowserRuntimePreset(platform, merged);
  const gate = merged.includeLaunchGate === false || merged.include_launch_gate === false
    ? undefined
    : buildMeetingAppLaunchGate(platform, launchGateOptions(merged));

  return compactObject({
    type: 'meeting_app_integration_profile',
    schema: MEETING_APP_INTEGRATION_PROFILE_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform,
    display_name: capability.display_name,
    recommended_mode: 'browser_extension_local_observer_first',
    extension: {
      profile: extensionProfile,
      matches: extensionProfile.matches,
      host_permissions: extensionProfile.host_permissions,
      manifest_permissions: extensionPlan.manifest?.permissions ?? [],
      recommended_permissions: uniqueList([
        ...(extensionPlan.manifest?.permissions ?? []),
        'storage',
        'tabs',
      ]),
      content_scripts: extensionPlan.content_scripts,
      adapters: {
        content_script: extensionPlan.content_script_adapter,
        browser_runtime: extensionPlan.browser_runtime_adapter,
        snapshot_recorder: extensionPlan.snapshot_recorder_adapter,
        launch_gate: extensionPlan.launch_gate_adapter,
      },
      message_types: { ...MEETING_APP_EXTENSION_MESSAGE_TYPES },
      status_storage_key: MEETING_APP_EXTENSION_STATUS_STORAGE_KEY,
      timeline_endpoints: { ...MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS },
    },
    runtime: runtimeSummary(runtimePreset),
    capture: {
      profile: captureProfile,
      ...selectorSummary(captureProfile),
    },
    event_model: eventModel(platform, capability, integrationPlan),
    provider_reconciliation: {
      enabled: integrationPlan.provider_events?.enabled === true,
      transport: integrationPlan.provider_events?.transport,
      endpoint: integrationPlan.provider_events?.endpoint,
      event_types: integrationPlan.provider_events?.event_types,
      lifecycle_event_types: integrationPlan.provider_events?.lifecycle_event_types,
    },
    readiness: gate ? readinessFor(gate, runtimePreset, captureProfile) : undefined,
    launch_gate: gate ? {
      status: gate.status,
      passed: gate.passed,
      production_ready: gate.production_ready,
      evidence_level: gate.evidence_level,
      runtime_ready: gate.runtime_ready,
      missing_required_coverage: gate.missing_required_coverage,
      warnings: gate.warnings,
      blocking_issues: gate.blocking_issues,
    } : undefined,
    implementation_steps: [
      'install_browser_extension_with_platform_host_permissions',
      'run_content_script_bridge_with_platform_runtime_preset',
      'create_or_restore_timeline_axis_from_local_meeting_started_signal',
      'insert_annotations_with_absolute_captured_at_ms',
      'reconcile_provider_start_end_and_artifacts_when_official_events_arrive',
      'import_post_meeting_transcript_after_meeting_end_when_available',
    ],
    limitations: uniqueList([
      ...(capability.limitations ?? []),
      'browser_dom_selectors_are_best_effort_and_must_be_verified_with_live_snapshots',
      'provider_transcripts_are_post_meeting_backfill_and_must_not_block_realtime_annotation',
    ]),
  });
}

export function buildAllMeetingAppIntegrationProfiles(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppIntegrationProfile(platform, options),
  ]));
}

export function buildMeetingAppIntegrationMatrix(options = {}) {
  const profiles = Object.values(buildAllMeetingAppIntegrationProfiles(options));
  return {
    type: 'meeting_app_integration_matrix',
    schema: MEETING_APP_INTEGRATION_PROFILE_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform_count: profiles.length,
    platforms: profiles.map((profile) => profile.platform),
    rows: profiles.map((profile) => ({
      platform: profile.platform,
      display_name: profile.display_name,
      recommended_mode: profile.recommended_mode,
      extension_match_count: profile.extension.matches.length,
      runtime_ready: profile.readiness?.runtime_ready ?? true,
      sdk_wiring_ready: profile.readiness?.sdk_wiring_ready ?? true,
      production_ready: profile.readiness?.production_ready ?? false,
      evidence_level: profile.readiness?.evidence_level,
      provider_reconciliation: profile.provider_reconciliation.enabled,
      post_meeting_transcript: profile.event_model.post_meeting_transcript?.status,
      next_actions: profile.readiness?.next_actions ?? [],
    })),
  };
}

export function buildMeetingAppRuntimeAdapterConfig(platformOrInput = {}, options = {}) {
  const input = typeof platformOrInput === 'string'
    ? { platform: platformOrInput }
    : (platformOrInput ?? {});
  const merged = { ...input, ...options };
  const profile = buildMeetingAppIntegrationProfile({
    ...merged,
    includeLaunchGate: firstNonEmpty(merged.includeLaunchGate, merged.include_launch_gate, false),
  });
  const platform = profile.platform;
  const captureProfile = profile.capture.profile;
  const runtimePreset = meetingAppBrowserRuntimePreset(platform, merged);
  const resolvedBridgeOptions = bridgeOptions(platform, merged);
  return compactObject({
    type: 'meeting_app_runtime_adapter_config',
    schema: MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform,
    display_name: profile.display_name,
    source: firstNonEmpty(merged.source, merged.detectorSource, merged.detector_source, 'meeting_app_extension'),
    extension: {
      matches: profile.extension.matches,
      host_permissions: profile.extension.host_permissions,
      permissions: profile.extension.recommended_permissions,
      message_types: profile.extension.message_types,
      status_storage_key: profile.extension.status_storage_key,
      timeline_endpoints: profile.extension.timeline_endpoints,
    },
    bridge_options: resolvedBridgeOptions,
    runtime_options: runtimeOptions(runtimePreset),
    capture_options: captureOptions(captureProfile),
    startup: {
      start_runtime: resolvedBridgeOptions.startRuntime,
      send_attached_message: true,
      attached_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
      timestamp_field: 'captured_at_ms',
    },
    supported_client_methods: Object.keys(MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS),
    readiness: compactObject({
      runtime_ready: (runtimePreset?.mutationTrackSelectors?.length ?? 0) > 0
        && (runtimePreset?.mutationIgnoreSelectors?.length ?? 0) > 0
        && Boolean(captureProfile?.platform),
      requires_live_snapshot_before_production: true,
      validation_gate: 'buildMeetingAppLaunchGate',
    }),
  });
}

export function buildAllMeetingAppRuntimeAdapterConfigs(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppRuntimeAdapterConfig(platform, options),
  ]));
}

export function buildMeetingAppRuntimeAdapterAcceptanceReport(configOrPlatform = {}, options = {}) {
  const config = runtimeConfigFromInput(configOrPlatform, options);
  const issues = [];
  let platform;
  try {
    platform = normalizeAppPlatform(config.platform);
  } catch (error) {
    issues.push(issue('error', 'unsupported_platform', 'Runtime adapter config platform is unsupported.', {
      platform: config.platform,
      error: String(error?.message ?? error),
    }));
  }
  if (config.type !== 'meeting_app_runtime_adapter_config') {
    issues.push(issue('error', 'invalid_type', 'Runtime adapter config type must be meeting_app_runtime_adapter_config.', {
      actual: config.type,
    }));
  }
  if (config.schema !== MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA) {
    issues.push(issue('error', 'invalid_schema', 'Runtime adapter config schema is invalid.', {
      actual: config.schema,
      expected: MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
    }));
  }
  if (asArray(config.extension?.host_permissions).includes('<all_urls>')) {
    issues.push(issue('error', 'overbroad_host_permission', 'Runtime adapter config must not request <all_urls>.'));
  }
  if (asArray(config.extension?.matches).length === 0) {
    issues.push(issue('error', 'missing_extension_matches', 'Runtime adapter config is missing extension match patterns.'));
  }
  if (!asArray(config.extension?.permissions).includes('storage')) {
    issues.push(issue('error', 'missing_storage_permission', 'Runtime adapter config must request storage for extension diagnostics.'));
  }
  if (!asArray(config.extension?.permissions).includes('tabs')) {
    issues.push(issue('error', 'missing_tabs_permission', 'Runtime adapter config must request tabs for browser candidate observation.'));
  }
  if (config.extension?.message_types?.client_call !== MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call) {
    issues.push(issue('error', 'missing_client_call_message_type', 'Runtime adapter config is missing the client_call message type.'));
  }
  if (config.extension?.message_types?.extension_attached !== MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached) {
    issues.push(issue('error', 'missing_attached_message_type', 'Runtime adapter config is missing the extension_attached message type.'));
  }
  if (config.extension?.message_types?.observe_candidates !== MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates) {
    issues.push(issue('error', 'missing_candidate_observer_message_type', 'Runtime adapter config is missing the observe_candidates message type.'));
  }
  if (config.extension?.status_storage_key !== MEETING_APP_EXTENSION_STATUS_STORAGE_KEY) {
    issues.push(issue('error', 'invalid_status_storage_key', 'Runtime adapter config has an unexpected status storage key.', {
      actual: config.extension?.status_storage_key,
    }));
  }
  for (const method of Object.keys(MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS)) {
    if (!asArray(config.supported_client_methods).includes(method)) {
      issues.push(issue('error', 'missing_supported_client_method', `Runtime adapter config is missing ${method}.`, {
        method,
      }));
    }
  }
  if (platform && config.bridge_options?.browser_runtime_preset !== platform) {
    issues.push(issue('error', 'bridge_preset_mismatch', 'Bridge runtime preset must match the config platform.', {
      platform,
      actual: config.bridge_options?.browser_runtime_preset,
    }));
  }
  if (platform && config.runtime_options?.runtimePreset !== platform) {
    issues.push(issue('error', 'runtime_preset_mismatch', 'Runtime preset must match the config platform.', {
      platform,
      actual: config.runtime_options?.runtimePreset,
    }));
  }
  if (platform && config.capture_options?.captureProfile !== platform) {
    issues.push(issue('error', 'capture_profile_mismatch', 'Capture profile must match the config platform.', {
      platform,
      actual: config.capture_options?.captureProfile,
    }));
  }
  if (config.runtime_options?.observeMutations !== true) {
    issues.push(issue('error', 'mutation_observer_disabled', 'Runtime adapter config must enable mutation observation.'));
  }
  if (asArray(config.runtime_options?.mutationTrackSelectors).length === 0) {
    issues.push(issue('error', 'missing_mutation_track_selectors', 'Runtime adapter config has no mutation track selectors.'));
  }
  if (asArray(config.runtime_options?.mutationIgnoreSelectors).length === 0) {
    issues.push(issue('error', 'missing_mutation_ignore_selectors', 'Runtime adapter config has no mutation ignore selectors.'));
  }
  if (asArray(config.capture_options?.participantSelectors).length === 0) {
    issues.push(issue('error', 'missing_participant_selectors', 'Runtime adapter config has no participant selectors.'));
  }
  if (config.startup?.attached_message_type !== MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached) {
    issues.push(issue('error', 'invalid_attached_startup_message', 'Startup attached message type is invalid.', {
      actual: config.startup?.attached_message_type,
    }));
  }
  if (config.readiness?.requires_live_snapshot_before_production !== true) {
    issues.push(issue('warning', 'missing_live_snapshot_requirement', 'Runtime adapter config should require live snapshot validation before production.'));
  }

  const accepted = issues.every((item) => item.severity !== 'error');
  return {
    type: 'meeting_app_runtime_adapter_acceptance_report',
    schema: MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    accepted,
    platform: config.platform,
    config,
    coverage: {
      extension_matches: asArray(config.extension?.matches).length > 0,
      storage_permission: asArray(config.extension?.permissions).includes('storage'),
      tabs_permission: asArray(config.extension?.permissions).includes('tabs'),
      message_types: Boolean(
        config.extension?.message_types?.client_call
        && config.extension?.message_types?.extension_attached
        && config.extension?.message_types?.observe_candidates
      ),
      candidate_observer_message_type: config.extension?.message_types?.observe_candidates === MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      bridge_preset: platform ? config.bridge_options?.browser_runtime_preset === platform : false,
      runtime_preset: platform ? config.runtime_options?.runtimePreset === platform : false,
      capture_profile: platform ? config.capture_options?.captureProfile === platform : false,
      mutation_observer: config.runtime_options?.observeMutations === true,
      mutation_selectors: asArray(config.runtime_options?.mutationTrackSelectors).length > 0,
      participant_selectors: asArray(config.capture_options?.participantSelectors).length > 0,
    },
    issues,
  };
}

export function assertMeetingAppRuntimeAdapterConfig(configOrPlatform = {}, options = {}) {
  const report = buildMeetingAppRuntimeAdapterAcceptanceReport(configOrPlatform, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app runtime adapter config acceptance failed', report);
  }
  return report;
}

export function buildAllMeetingAppRuntimeAdapterAcceptanceReports(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppRuntimeAdapterAcceptanceReport(platform, options),
  ]));
}

function sourcedIssues(source, issues = []) {
  return asArray(issues).map((item) => ({
    source,
    ...item,
  }));
}

function validationOptions(options = {}) {
  return {
    allowFixtureEvidence: false,
    allow_fixture_evidence: false,
    allowFixtureProduction: false,
    allow_fixture_production: false,
    requireProductionReady: false,
    require_production_ready: false,
    ...options,
  };
}

function snapshotCaptureStep(id, phase, label, requiredCoverage = [], options = {}) {
  return compactObject({
    id,
    phase,
    label,
    required_coverage: requiredCoverage,
    min_count: firstNonEmpty(options.minCount, options.min_count, 1),
    capture_timing: firstNonEmpty(options.captureTiming, options.capture_timing),
    notes: options.notes,
  });
}

export function buildMeetingAppLiveSnapshotCapturePlan(platformOrInput = {}, options = {}) {
  const config = runtimeConfigFromInput(platformOrInput, options);
  const platform = normalizeAppPlatform(config.platform);
  const emptyGate = buildMeetingAppLaunchGate(platform, validationOptions(options));
  const activeCoverage = [
    'platform_detected',
    'meeting_id',
    'in_meeting',
    'meeting_started',
    'active_speaker',
    'speaker_started',
  ];
  const endedCoverage = [
    'meeting_ended',
    'ended_in_meeting_false',
  ].filter((item) => emptyGate.required_coverage?.includes(item));
  return compactObject({
    type: 'meeting_app_live_snapshot_capture_plan',
    schema: MEETING_APP_LIVE_SNAPSHOT_CAPTURE_PLAN_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform,
    display_name: config.display_name,
    runtime_config: config,
    recorder: {
      adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder',
      capture_method: 'createMeetingAppSnapshotRecorder().capture',
      export_method: 'createMeetingAppSnapshotRecorder().exportRecords',
      timestamp_field: 'captured_at_ms',
    },
    validation: {
      adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile',
      method: 'buildMeetingAppRuntimeAdapterValidationReport',
      production_ready_requires: 'captured_dom',
      required_coverage: emptyGate.required_coverage,
      default_options: {
        allowFixtureEvidence: false,
        requireMeetingEnd: options.requireMeetingEnd ?? options.require_meeting_end ?? true,
      },
    },
    required_snapshots: [
      snapshotCaptureStep(
        'active_speaker',
        'active',
        'Capture while the meeting is active, at least one participant is visible, and one speaker is active.',
        activeCoverage.filter((item) => emptyGate.required_coverage?.includes(item)),
        { captureTiming: 'after_join_and_audio_activity_is_visible' },
      ),
      snapshotCaptureStep(
        'meeting_ended',
        'ended',
        'Capture after leaving or ending the same meeting so the observer can emit meeting_ended.',
        endedCoverage,
        { captureTiming: 'after_leave_or_end_meeting' },
      ),
    ].filter((step) => step.required_coverage.length > 0),
    recommended_snapshots: [
      snapshotCaptureStep(
        'participants_panel',
        'active',
        'Capture with the participants panel or roster visible when the platform exposes richer participant DOM.',
        ['participant_selectors'],
        { notes: 'Useful for selector tuning even if launch gate does not require participant_joined.' },
      ),
      snapshotCaptureStep(
        'screen_share_or_presentation',
        'active',
        'Capture during screen sharing when this is a common meeting state for the product.',
        ['platform_detected', 'in_meeting'],
        { notes: 'Prevents screen-share UI from hiding controls or speaker markers unexpectedly.' },
      ),
    ],
    minimum_record_count: emptyGate.required_coverage?.includes('meeting_ended') ? 2 : 1,
    handoff: {
      validation_input: 'Pass recorder.exportRecords() as { records } or pass snapshots directly as { snapshots }.',
      success_condition: 'meetingAppRuntimeAdapterValidation(...).production_ready === true',
    },
  });
}

export function buildAllMeetingAppLiveSnapshotCapturePlans(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppLiveSnapshotCapturePlan(platform, options),
  ]));
}

export function buildMeetingAppDeploymentManifest(platformOrInput = {}, options = {}) {
  const config = runtimeConfigFromInput(platformOrInput, options);
  const platform = normalizeAppPlatform(config.platform);
  const profile = buildMeetingAppIntegrationProfile({
    ...options,
    platform,
  });
  const extensionInstallPlan = buildMeetingAppExtensionInstallPlan({
    ...options,
    platforms: [platform],
  });
  const capturePlan = buildMeetingAppLiveSnapshotCapturePlan(config, options);
  const validationReport = buildMeetingAppRuntimeAdapterValidationReport(config, options);
  return compactObject({
    type: 'meeting_app_deployment_manifest',
    schema: MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform,
    display_name: config.display_name,
    recommended_mode: profile.recommended_mode,
    profile,
    runtime_config: config,
    extension_install_plan: extensionInstallPlan,
    live_snapshot_capture_plan: capturePlan,
    validation_report: validationReport,
    runtime_contract: {
      timestamp_field: 'captured_at_ms',
      annotation_time_invariant: 'write_annotations_with_absolute_capture_time_not_meeting_offset',
      required_signals: ['meeting_started', 'speaker_started', 'meeting_ended'],
      supported_client_methods: config.supported_client_methods,
      message_types: config.extension?.message_types,
      timeline_endpoints: config.extension?.timeline_endpoints,
      candidate_observation: {
        message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
        runtime_event_action: 'observe_platform_candidates',
        required_permission: 'tabs',
        producer: 'browser_extension_background_or_native_host',
      },
    },
    integration_targets: [
      {
        surface: 'chrome_or_edge_extension',
        injection: 'manifest_v3_content_script',
        matches: extensionInstallPlan.matches,
        host_permissions: extensionInstallPlan.host_permissions,
        permissions: config.extension?.permissions,
        content_script_adapter: extensionInstallPlan.content_script_adapter,
        background_bridge: 'meeting_app_extension_background_forwarder',
        candidate_observer: {
          message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
          runtime_event_action: 'observe_platform_candidates',
        },
      },
      {
        surface: 'electron_or_embedded_webview',
        injection: 'preload_bridge',
        runtime_preset: config.runtime_options?.runtimePreset,
        browser_runtime_adapter: extensionInstallPlan.browser_runtime_adapter,
      },
      {
        surface: 'native_desktop_accessibility_observer',
        injection: 'dom_or_accessibility_snapshot_adapter',
        capture_profile: config.capture_options?.captureProfile,
        snapshot_recorder_adapter: extensionInstallPlan.snapshot_recorder_adapter,
      },
    ],
    production_gate: {
      accepted: validationReport.accepted,
      production_ready: validationReport.production_ready,
      requires_captured_dom: true,
      evidence_level: validationReport.evidence_level,
      minimum_live_record_count: capturePlan.minimum_record_count,
      required_coverage: capturePlan.validation?.required_coverage,
      next_actions: validationReport.next_actions,
    },
    handoff: {
      package: '@ai-annotation/meeting-timeline-sdk',
      primary_entry: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      kit_methods: [
        'meetingAppDeploymentManifest',
        'meetingAppRuntimeAdapterConfig',
        'meetingAppExtensionScaffold',
        'meetingAppLiveSnapshotCapturePlan',
        'meetingAppRuntimeAdapterValidation',
      ],
      validation_input: capturePlan.handoff?.validation_input,
      success_condition: capturePlan.handoff?.success_condition,
    },
    rollout_checklist: [
      'install_or_embed_runtime_for_this_platform_only',
      'verify_extension_or_preload_attached_status',
      'capture_required_live_snapshots_from_a_real_meeting',
      'pass_runtime_validation_with_production_ready_true',
      'insert_live_annotations_with_captured_at_ms_from_device_or_mark_source',
      'use_provider_events_or_post_meeting_artifacts_only_for_reconciliation',
    ],
  });
}

export function buildAllMeetingAppDeploymentManifests(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppDeploymentManifest(platform, options),
  ]));
}

export function buildMeetingAppDeploymentManifestAcceptanceReport(manifestOrPlatform = {}, options = {}) {
  let manifest;
  const issues = [];
  try {
    manifest = deploymentManifestFromInput(manifestOrPlatform, options);
  } catch (error) {
    return {
      type: 'meeting_app_deployment_manifest_acceptance_report',
      schema: MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
      version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
      accepted: false,
      production_ready: false,
      platform: manifestOrPlatform?.platform,
      manifest: manifestOrPlatform,
      coverage: {},
      issues: [issue('error', 'manifest_build_failed', 'Deployment manifest could not be built.', {
        error: String(error?.message ?? error),
      })],
    };
  }

  let platform;
  try {
    platform = normalizeAppPlatform(manifest.platform);
  } catch (error) {
    issues.push(issue('error', 'unsupported_platform', 'Deployment manifest platform is unsupported.', {
      platform: manifest.platform,
      error: String(error?.message ?? error),
    }));
  }

  if (manifest.type !== 'meeting_app_deployment_manifest') {
    issues.push(issue('error', 'invalid_type', 'Deployment manifest type must be meeting_app_deployment_manifest.', {
      actual: manifest.type,
    }));
  }
  if (manifest.schema !== MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA) {
    issues.push(issue('error', 'invalid_schema', 'Deployment manifest schema is invalid.', {
      actual: manifest.schema,
      expected: MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
    }));
  }

  const platformMismatch = (source, value) => {
    if (!platform || !value) return;
    try {
      const normalized = normalizeAppPlatform(value);
      if (normalized !== platform) {
        issues.push(issue('error', `${source}_platform_mismatch`, `${source} platform must match manifest platform.`, {
          expected: platform,
          actual: normalized,
        }));
      }
    } catch (error) {
      issues.push(issue('error', `${source}_unsupported_platform`, `${source} platform is unsupported.`, {
        actual: value,
        error: String(error?.message ?? error),
      }));
    }
  };

  if (!manifest.profile) {
    issues.push(issue('error', 'missing_profile', 'Deployment manifest is missing profile.'));
  } else {
    if (manifest.profile.type !== 'meeting_app_integration_profile') {
      issues.push(issue('error', 'invalid_profile_type', 'Deployment manifest profile type is invalid.', {
        actual: manifest.profile.type,
      }));
    }
    platformMismatch('profile', manifest.profile.platform);
  }

  let runtimeAcceptance = null;
  if (!manifest.runtime_config) {
    issues.push(issue('error', 'missing_runtime_config', 'Deployment manifest is missing runtime_config.'));
  } else {
    platformMismatch('runtime_config', manifest.runtime_config.platform);
    runtimeAcceptance = buildMeetingAppRuntimeAdapterAcceptanceReport(manifest.runtime_config, options);
    issues.push(...sourcedIssues('runtime_config', runtimeAcceptance.issues));
  }

  const extensionPlan = manifest.extension_install_plan ?? {};
  if (!manifest.extension_install_plan) {
    issues.push(issue('error', 'missing_extension_install_plan', 'Deployment manifest is missing extension_install_plan.'));
  }
  if (platform && !asArray(extensionPlan.platforms).some((item) => {
    try {
      return normalizeAppPlatform(item) === platform;
    } catch {
      return false;
    }
  })) {
    issues.push(issue('error', 'extension_platform_missing', 'Extension install plan does not include the manifest platform.', {
      platform,
      platforms: extensionPlan.platforms,
    }));
  }
  if (asArray(extensionPlan.matches).length === 0) {
    issues.push(issue('error', 'missing_extension_matches', 'Deployment manifest extension install plan has no matches.'));
  }
  if (asArray(extensionPlan.host_permissions).length === 0) {
    issues.push(issue('error', 'missing_extension_host_permissions', 'Deployment manifest extension install plan has no host permissions.'));
  }
  if (asArray(extensionPlan.host_permissions).includes('<all_urls>')) {
    issues.push(issue('error', 'overbroad_host_permission', 'Deployment manifest must not request <all_urls>.'));
  }
  if (!extensionPlan.content_script_adapter || !extensionPlan.browser_runtime_adapter) {
    issues.push(issue('error', 'missing_extension_adapters', 'Deployment manifest extension install plan is missing SDK adapters.'));
  }

  const capturePlan = manifest.live_snapshot_capture_plan ?? {};
  if (!manifest.live_snapshot_capture_plan) {
    issues.push(issue('error', 'missing_live_snapshot_capture_plan', 'Deployment manifest is missing live_snapshot_capture_plan.'));
  } else {
    platformMismatch('live_snapshot_capture_plan', capturePlan.platform);
  }
  const requiredSnapshots = asArray(capturePlan.required_snapshots);
  const requiredSnapshotIds = requiredSnapshots.map((item) => item?.id);
  if (requiredSnapshots.length === 0) {
    issues.push(issue('error', 'missing_required_snapshots', 'Live snapshot capture plan has no required snapshots.'));
  }
  if (!requiredSnapshotIds.includes('active_speaker')) {
    issues.push(issue('error', 'missing_active_speaker_snapshot', 'Live snapshot capture plan must require an active_speaker snapshot.'));
  }
  const requireMeetingEnd = options.requireMeetingEnd !== false
    && options.require_meeting_end !== false
    && asArray(capturePlan.validation?.required_coverage).includes('meeting_ended');
  if (requireMeetingEnd && !requiredSnapshotIds.includes('meeting_ended')) {
    issues.push(issue('error', 'missing_meeting_ended_snapshot', 'Live snapshot capture plan must require a meeting_ended snapshot.'));
  }
  if (Number(capturePlan.minimum_record_count ?? 0) < requiredSnapshots.length) {
    issues.push(issue('error', 'minimum_record_count_too_low', 'Live snapshot minimum_record_count is lower than required_snapshots length.', {
      minimum_record_count: capturePlan.minimum_record_count,
      required_snapshot_count: requiredSnapshots.length,
    }));
  }

  const contract = manifest.runtime_contract ?? {};
  const requiredSignals = asArray(contract.required_signals);
  if (contract.timestamp_field !== 'captured_at_ms') {
    issues.push(issue('error', 'invalid_runtime_timestamp_field', 'Runtime contract timestamp_field must be captured_at_ms.', {
      actual: contract.timestamp_field,
    }));
  }
  for (const signal of ['meeting_started', 'speaker_started', ...(requireMeetingEnd ? ['meeting_ended'] : [])]) {
    if (!requiredSignals.includes(signal)) {
      issues.push(issue('error', 'missing_runtime_required_signal', `Runtime contract is missing ${signal}.`, { signal }));
    }
  }
  if (contract.annotation_time_invariant !== 'write_annotations_with_absolute_capture_time_not_meeting_offset') {
    issues.push(issue('error', 'invalid_annotation_time_invariant', 'Runtime contract must preserve absolute annotation capture time.', {
      actual: contract.annotation_time_invariant,
    }));
  }

  const integrationSurfaces = asArray(manifest.integration_targets).map((item) => item?.surface);
  if (!integrationSurfaces.includes('chrome_or_edge_extension')) {
    issues.push(issue('error', 'missing_extension_target', 'Deployment manifest must describe a browser extension integration target.'));
  }
  if (!integrationSurfaces.includes('electron_or_embedded_webview')) {
    issues.push(issue('warning', 'missing_embedded_webview_target', 'Deployment manifest should describe an Electron/WebView integration target.'));
  }
  if (!integrationSurfaces.includes('native_desktop_accessibility_observer')) {
    issues.push(issue('warning', 'missing_native_accessibility_target', 'Deployment manifest should describe a native accessibility observer target.'));
  }

  const validationReport = manifest.validation_report ?? {};
  if (!manifest.validation_report) {
    issues.push(issue('error', 'missing_validation_report', 'Deployment manifest is missing validation_report.'));
  } else {
    platformMismatch('validation_report', validationReport.platform);
  }
  if (runtimeAcceptance?.accepted === false) {
    issues.push(issue('error', 'runtime_config_not_accepted', 'Deployment manifest runtime config is not accepted.'));
  }

  const productionGate = manifest.production_gate ?? {};
  if (productionGate.requires_captured_dom !== true) {
    issues.push(issue('error', 'production_gate_must_require_captured_dom', 'Deployment manifest production gate must require captured DOM evidence.'));
  }
  if (Number(productionGate.minimum_live_record_count ?? 0) < Number(capturePlan.minimum_record_count ?? 1)) {
    issues.push(issue('error', 'production_gate_record_count_too_low', 'Production gate minimum live record count is lower than capture plan.', {
      production_gate_minimum: productionGate.minimum_live_record_count,
      capture_plan_minimum: capturePlan.minimum_record_count,
    }));
  }
  if (productionGate.production_ready !== validationReport.production_ready) {
    issues.push(issue('error', 'production_gate_validation_mismatch', 'Production gate production_ready must mirror validation_report.production_ready.', {
      production_gate: productionGate.production_ready,
      validation_report: validationReport.production_ready,
    }));
  }
  if (validationReport.production_ready !== true) {
    issues.push(issue('warning', 'live_dom_not_verified', 'Deployment manifest is handoff-ready but still needs captured DOM snapshots before production.', {
      evidence_level: validationReport.evidence_level ?? productionGate.evidence_level,
    }));
  }

  if (manifest.handoff?.primary_entry !== '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit') {
    issues.push(issue('error', 'invalid_handoff_primary_entry', 'Deployment manifest handoff must point to platform-kit.', {
      actual: manifest.handoff?.primary_entry,
    }));
  }
  if (!asArray(manifest.rollout_checklist).includes('pass_runtime_validation_with_production_ready_true')) {
    issues.push(issue('error', 'missing_rollout_validation_step', 'Rollout checklist must include the runtime validation production gate.'));
  }

  const accepted = issues.every((item) => item.severity !== 'error');
  const productionReady = accepted
    && productionGate.production_ready === true
    && validationReport.production_ready === true;
  return {
    type: 'meeting_app_deployment_manifest_acceptance_report',
    schema: MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    accepted,
    production_ready: productionReady,
    platform,
    manifest,
    coverage: {
      profile: Boolean(manifest.profile),
      runtime_config: runtimeAcceptance?.accepted === true,
      extension_install_plan: Boolean(extensionPlan.content_script_adapter && extensionPlan.browser_runtime_adapter),
      scoped_host_permissions: !asArray(extensionPlan.host_permissions).includes('<all_urls>'),
      live_snapshot_capture_plan: requiredSnapshots.length > 0,
      runtime_contract: contract.timestamp_field === 'captured_at_ms',
      production_gate: productionGate.requires_captured_dom === true,
      handoff: manifest.handoff?.primary_entry === '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      live_dom_verified: validationReport.production_ready === true,
    },
    issues,
  };
}

export function assertMeetingAppDeploymentManifest(manifestOrPlatform = {}, options = {}) {
  const report = buildMeetingAppDeploymentManifestAcceptanceReport(manifestOrPlatform, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app deployment manifest acceptance failed', report);
  }
  return report;
}

export function buildAllMeetingAppDeploymentManifestAcceptanceReports(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppDeploymentManifestAcceptanceReport(platform, options),
  ]));
}

export function buildMeetingAppDeploymentManifestAcceptanceSummary(options = {}) {
  const reports = Object.values(buildAllMeetingAppDeploymentManifestAcceptanceReports(options));
  return {
    type: 'meeting_app_deployment_manifest_acceptance_summary',
    schema: MEETING_APP_DEPLOYMENT_MANIFEST_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    accepted: reports.every((report) => report.accepted),
    production_ready: reports.every((report) => report.production_ready),
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    production_ready_count: reports.filter((report) => report.production_ready).length,
    rows: reports.map((report) => ({
      platform: report.platform,
      accepted: report.accepted,
      production_ready: report.production_ready,
      evidence_level: report.manifest?.validation_report?.evidence_level ?? report.manifest?.production_gate?.evidence_level,
      error_count: report.issues.filter((item) => item.severity === 'error').length,
      warning_count: report.issues.filter((item) => item.severity === 'warning').length,
      next_actions: uniqueList([
        ...(report.manifest?.production_gate?.next_actions ?? []),
        ...report.issues.map((item) => item.code),
      ]),
    })),
  };
}

function evidenceInputOptions(input = {}, options = {}) {
  if (Array.isArray(input)) return { ...options, records: input };
  if (input && typeof input === 'object') return { ...input, ...options };
  return options;
}

function evidenceRows(value) {
  if (value == null) return [];
  if (Array.isArray(value)) return value;
  if (value?.records || value?.schema) return meetingAppSnapshotRecords(value);
  if (typeof value === 'object') {
    const rows = [];
    for (const [key, rowsForPlatform] of Object.entries(value)) {
      let platform;
      try {
        platform = normalizeAppPlatform(key);
      } catch {
        continue;
      }
      rows.push(...asArray(rowsForPlatform).map((row) => (
        row && typeof row === 'object' && !row.platform && !row.provider
          ? { ...row, platform }
          : row
      )));
    }
    if (rows.length > 0) return rows;
  }
  return asArray(value);
}

function recordsForPlatform(recordSet = {}, platform) {
  return meetingAppSnapshotRecords(recordSet).filter((record) => {
    const candidate = record?.platform ?? record?.provider ?? record?.snapshot?.platform ?? record?.snapshot?.provider ?? record?.snapshot?.capture?.profile;
    if (!candidate) return false;
    try {
      return normalizeAppPlatform(candidate) === platform;
    } catch {
      return false;
    }
  });
}

function arrayAtPath(input = {}, path) {
  const value = path.split('.').reduce((node, part) => node?.[part], input);
  return Array.isArray(value) ? value : [];
}

function maxNumberAtPaths(input = {}, paths = []) {
  const values = paths.map((path) => path.split('.').reduce((node, part) => node?.[part], input));
  const numbers = values.map((value) => Number(value)).filter((value) => Number.isFinite(value));
  return numbers.length > 0 ? Math.max(...numbers) : 0;
}

function collectionCountAtPaths(input = {}, paths = []) {
  const direct = paths.reduce((count, path) => count + arrayAtPath(input, path).length, 0);
  if (direct > 0) return direct;
  if (paths.some((path) => path.includes('control') || path.includes('button'))) {
    return maxNumberAtPaths(input, ['capture.control_count']);
  }
  if (paths.some((path) => path.includes('participant') || path.includes('tile'))) {
    return maxNumberAtPaths(input, ['capture.participant_count']);
  }
  if (paths.some((path) => path.includes('text'))) {
    return maxNumberAtPaths(input, ['capture.text_count']);
  }
  return 0;
}

function sumSnapshotCollections(snapshots = [], paths = []) {
  return snapshots.reduce((count, snapshot) => count + collectionCountAtPaths(snapshot, paths), 0);
}

function diagnosisPlatform(input = {}, recordSet = {}) {
  const explicit = firstNonEmpty(input.platform, input.provider, input.key, input.name);
  if (explicit) return normalizeAppPlatform(explicit);
  const candidates = uniqueList([
    ...(recordSet.platforms ?? []),
    ...meetingAppSnapshotRecords(recordSet).map((record) => (
      record?.platform
      ?? record?.provider
      ?? record?.snapshot?.platform
      ?? record?.snapshot?.provider
      ?? record?.snapshot?.capture?.profile
    )),
  ]);
  if (candidates.length === 0) {
    throw new MeetingTimelineSdkError('Meeting app DOM adaptation diagnosis requires platform or evidence with platform.', {
      supported_platforms: MEETING_APP_INTEGRATION_PROFILE_PLATFORMS,
    });
  }
  return normalizeAppPlatform(candidates[0]);
}

function diagnosisIssue(severity, code, message, details = {}) {
  return issue(severity, code, message, details);
}

function diagnosisNextActions(issues = [], launchGate = {}) {
  const mapped = issues.map((item) => {
    if (item.code === 'missing_live_snapshots') return 'capture_required_live_snapshots';
    if (item.code === 'missing_controls') return 'tune_control_selectors_or_capture_toolbar_visible_state';
    if (item.code === 'missing_participants') return 'capture_participants_visible_state_or_tune_participant_selectors';
    if (item.code === 'missing_active_speaker') return 'capture_active_speaker_snapshot_with_audio_activity';
    if (item.code === 'missing_meeting_end') return 'capture_after_leave_or_end_meeting';
    if (item.code === 'missing_meeting_identity') return 'include_meeting_url_or_title_in_snapshot';
    return item.code;
  });
  return uniqueList([
    ...mapped,
    ...(launchGate.next_actions ?? []),
  ]);
}

export function buildMeetingAppDomAdaptationDiagnosis(platformOrInput = {}, options = {}) {
  const input = typeof platformOrInput === 'string'
    ? { platform: platformOrInput }
    : (platformOrInput ?? {});
  const merged = evidenceInputOptions(input, options);
  const recordSet = buildEvidenceRecordSet(input, options);
  const platform = diagnosisPlatform(merged, recordSet);
  const config = buildMeetingAppRuntimeAdapterConfig(platform, {
    ...merged,
    includeLaunchGate: false,
  });
  const records = recordsForPlatform(recordSet, platform);
  const snapshots = records.map((record) => record.snapshot).filter(Boolean);
  const normalizedSnapshots = snapshots
    .map((snapshot) => normalizeMeetingAppSnapshot(snapshot, { ...merged, platform }))
    .filter(Boolean);
  const launchGate = buildMeetingAppLaunchGate(platform, {
    ...merged,
    records: recordSet,
    snapshotRecords: recordSet,
    snapshot_records: recordSet,
    allowFixtureEvidence: false,
    allow_fixture_evidence: false,
    requireProductionReady: false,
    require_production_ready: false,
  });
  const selectorCounts = {
    controls: sumSnapshotCollections(snapshots, [
      'page.controls',
      'page.buttons',
      'dom.controls',
      'controls',
      'buttons',
    ]),
    participants: sumSnapshotCollections(snapshots, [
      'page.participants',
      'page.tiles',
      'dom.participants',
      'dom.tiles',
      'participants',
      'tiles',
    ]),
    texts: sumSnapshotCollections(snapshots, [
      'page.texts',
      'dom.texts',
      'texts',
    ]),
    normalized_active_speaker_snapshots: normalizedSnapshots.filter((snapshot) => (
      Boolean(snapshot.activeSpeaker?.id || snapshot.activeSpeaker?.name)
    )).length,
    normalized_in_meeting_snapshots: normalizedSnapshots.filter((snapshot) => snapshot.inMeeting === true).length,
    normalized_ended_snapshots: normalizedSnapshots.filter((snapshot) => snapshot.inMeeting === false).length,
  };
  const signalTypes = launchGate.reports?.captured?.signal_types ?? [];
  const requireMeetingEnd = merged.requireMeetingEnd !== false && merged.require_meeting_end !== false;
  const capturePlan = buildMeetingAppLiveSnapshotCapturePlan(platform, merged);
  const issues = [];
  if (records.length === 0) {
    issues.push(diagnosisIssue('error', 'missing_live_snapshots', 'No live DOM snapshots were supplied for this platform.', {
      platform,
    }));
  }
  if (selectorCounts.controls === 0) {
    issues.push(diagnosisIssue('error', 'missing_controls', 'No meeting controls were detected; lifecycle detection is likely unreliable.', {
      expected_selector_count: config.capture_options?.controlSelectors?.length ?? 0,
    }));
  }
  if (selectorCounts.participants === 0) {
    issues.push(diagnosisIssue('error', 'missing_participants', 'No participant tiles or roster rows were detected.', {
      expected_selector_count: config.capture_options?.participantSelectors?.length ?? 0,
    }));
  }
  if (launchGate.coverage?.meeting_id !== true) {
    issues.push(diagnosisIssue('error', 'missing_meeting_identity', 'No stable meeting id or meeting URL was inferred from the snapshots.'));
  }
  if (launchGate.coverage?.meeting_started !== true) {
    issues.push(diagnosisIssue('error', 'missing_meeting_start', 'The snapshots did not produce a meeting_started signal.'));
  }
  if (launchGate.coverage?.active_speaker !== true || launchGate.coverage?.speaker_started !== true) {
    issues.push(diagnosisIssue('error', 'missing_active_speaker', 'The snapshots did not produce an active speaker / speaker_started signal.'));
  }
  if (requireMeetingEnd && launchGate.coverage?.meeting_ended !== true) {
    issues.push(diagnosisIssue('error', 'missing_meeting_end', 'The snapshots did not prove a meeting_ended transition.'));
  }
  if (selectorCounts.texts === 0) {
    issues.push(diagnosisIssue('warning', 'missing_text_surfaces', 'No status/title text surfaces were detected; provider-specific selector tuning may still be needed.'));
  }
  const accepted = issues.every((item) => item.severity !== 'error');
  return compactObject({
    type: 'meeting_app_dom_adaptation_diagnosis',
    schema: MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform,
    display_name: config.display_name,
    accepted,
    production_ready: accepted && launchGate.production_ready === true,
    evidence_level: launchGate.evidence_level,
    evidence_count: launchGate.evidence_count,
    record_count: records.length,
    phases: uniqueList(records.map((record) => record.phase)),
    selector_probe: {
      capture_profile: config.capture_options?.captureProfile,
      configured_selector_counts: {
        controls: config.capture_options?.controlSelectors?.length ?? 0,
        participants: config.capture_options?.participantSelectors?.length ?? 0,
        texts: config.capture_options?.textSelectors?.length ?? 0,
      },
      matched_counts: selectorCounts,
      matched: {
        controls: selectorCounts.controls > 0,
        participants: selectorCounts.participants > 0,
        texts: selectorCounts.texts > 0,
        active_speaker: selectorCounts.normalized_active_speaker_snapshots > 0,
      },
    },
    observer_probe: {
      signal_types: signalTypes,
      coverage: launchGate.coverage,
      missing_required_coverage: launchGate.missing_required_coverage ?? [],
    },
    runtime_probe: {
      runtime_ready: launchGate.runtime_ready,
      mutation_track_selector_count: config.runtime_options?.mutationTrackSelectors?.length ?? 0,
      mutation_ignore_selector_count: config.runtime_options?.mutationIgnoreSelectors?.length ?? 0,
      debounce_ms: config.runtime_options?.mutationDebounceMs,
      speaker_stable_followup_ms: config.runtime_options?.speakerStableFollowupMs,
    },
    recommended_capture: {
      minimum_record_count: capturePlan.minimum_record_count,
      required_snapshot_ids: capturePlan.required_snapshots.map((step) => step.id),
    },
    next_actions: diagnosisNextActions(issues, launchGate),
    issues,
  });
}

export function buildAllMeetingAppDomAdaptationDiagnoses(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppDomAdaptationDiagnosis(platform, options),
  ]));
}

export function buildMeetingAppDomAdaptationDiagnosisMatrix(options = {}) {
  const diagnoses = Object.values(buildAllMeetingAppDomAdaptationDiagnoses(options));
  return {
    type: 'meeting_app_dom_adaptation_diagnosis_matrix',
    schema: MEETING_APP_DOM_ADAPTATION_DIAGNOSIS_MATRIX_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    platform_count: diagnoses.length,
    accepted_count: diagnoses.filter((diagnosis) => diagnosis.accepted).length,
    production_ready_count: diagnoses.filter((diagnosis) => diagnosis.production_ready).length,
    active_speaker_ready_count: diagnoses.filter((diagnosis) => diagnosis.selector_probe?.matched?.active_speaker === true).length,
    meeting_start_ready_count: diagnoses.filter((diagnosis) => diagnosis.observer_probe?.coverage?.meeting_started === true).length,
    meeting_end_ready_count: diagnoses.filter((diagnosis) => diagnosis.observer_probe?.coverage?.meeting_ended === true).length,
    platforms: diagnoses.map((diagnosis) => diagnosis.platform),
    rows: diagnoses.map((diagnosis) => ({
      platform: diagnosis.platform,
      display_name: diagnosis.display_name,
      accepted: diagnosis.accepted,
      production_ready: diagnosis.production_ready,
      evidence_level: diagnosis.evidence_level,
      evidence_count: diagnosis.evidence_count,
      record_count: diagnosis.record_count,
      controls_matched: diagnosis.selector_probe?.matched?.controls === true,
      participants_matched: diagnosis.selector_probe?.matched?.participants === true,
      active_speaker_matched: diagnosis.selector_probe?.matched?.active_speaker === true,
      meeting_started: diagnosis.observer_probe?.coverage?.meeting_started === true,
      speaker_started: diagnosis.observer_probe?.coverage?.speaker_started === true,
      meeting_ended: diagnosis.observer_probe?.coverage?.meeting_ended === true,
      missing_required_coverage: diagnosis.observer_probe?.missing_required_coverage ?? [],
      issue_codes: (diagnosis.issues ?? []).map((item) => item.code),
      next_actions: diagnosis.next_actions ?? [],
    })),
    diagnoses,
    next_actions: uniqueList(diagnoses.flatMap((diagnosis) => diagnosis.next_actions ?? [])),
  };
}

function buildEvidenceRecordSet(input = {}, options = {}) {
  const merged = evidenceInputOptions(input, options);
  const primary = firstNonEmpty(
    merged.recordSet,
    merged.record_set,
    merged.snapshotRecords,
    merged.snapshot_records,
    merged.records,
    Array.isArray(input) ? input : undefined,
  );
  const snapshots = firstNonEmpty(
    merged.snapshots,
    merged.domSnapshots,
    merged.dom_snapshots,
  );
  const rows = [
    ...evidenceRows(primary),
    ...evidenceRows(snapshots),
  ];
  return buildMeetingAppSnapshotRecordSet(rows, {
    id: firstNonEmpty(merged.recordSetId, merged.record_set_id, merged.id),
    label: firstNonEmpty(merged.label),
    source: firstNonEmpty(merged.source, 'meeting_app_live_evidence_package'),
    createdAtMs: firstNonEmpty(merged.createdAtMs, merged.created_at_ms, merged.capturedAtMs, merged.captured_at_ms),
  });
}

function recordsByPlatform(recordSet = {}, platforms = []) {
  const output = Object.fromEntries(platforms.map((platform) => [platform, []]));
  for (const record of meetingAppSnapshotRecords(recordSet)) {
    if (!record?.platform && !record?.provider) continue;
    let platform;
    try {
      platform = normalizeAppPlatform(record.platform ?? record.provider);
    } catch {
      continue;
    }
    if (!output[platform]) output[platform] = [];
    output[platform].push(record);
  }
  return output;
}

export function buildMeetingAppLiveEvidencePackage(input = {}, options = {}) {
  const merged = evidenceInputOptions(input, options);
  const recordSet = buildEvidenceRecordSet(input, options);
  const platforms = platformListFromEvidence(recordSet, merged);
  const groupedRecords = recordsByPlatform(recordSet, platforms);
  const packageIssues = [];
  const requireEvidence = firstNonEmpty(merged.requireEvidence, merged.require_evidence, true);
  for (const platform of platforms) {
    if (requireEvidence && (groupedRecords[platform]?.length ?? 0) === 0) {
      packageIssues.push(issue('error', 'missing_platform_evidence_records', 'No live evidence records were supplied for platform.', {
        platform,
      }));
    }
  }
  const evidenceOptionsForValidation = {
    ...merged,
    records: recordSet,
    snapshotRecords: recordSet,
    snapshot_records: recordSet,
    allowFixtureEvidence: false,
    allow_fixture_evidence: false,
  };
  const manifestAcceptance = Object.fromEntries(platforms.map((platform) => [
    platform,
    buildMeetingAppDeploymentManifestAcceptanceReport(platform, evidenceOptionsForValidation),
  ]));
  const rows = platforms.map((platform) => {
    const report = manifestAcceptance[platform];
    const records = groupedRecords[platform] ?? [];
    return {
      platform,
      record_count: records.length,
      accepted: report.accepted,
      production_ready: report.production_ready,
      evidence_level: report.manifest?.validation_report?.evidence_level,
      evidence_count: report.manifest?.validation_report?.evidence_count ?? records.length,
      missing_required_coverage: report.manifest?.validation_report?.launch_gate?.missing_required_coverage ?? [],
      error_count: report.issues.filter((item) => item.severity === 'error').length
        + packageIssues.filter((item) => item.platform === platform && item.severity === 'error').length,
      warning_count: report.issues.filter((item) => item.severity === 'warning').length
        + packageIssues.filter((item) => item.platform === platform && item.severity === 'warning').length,
      next_actions: uniqueList([
        ...(report.manifest?.validation_report?.next_actions ?? []),
        ...report.issues.map((item) => item.code),
        ...packageIssues.filter((item) => item.platform === platform).map((item) => item.code),
      ]),
    };
  });
  const accepted = packageIssues.every((item) => item.severity !== 'error')
    && Object.values(manifestAcceptance).every((report) => report.accepted);
  const productionReady = accepted && Object.values(manifestAcceptance).every((report) => report.production_ready);
  const createdAtMs = normalizeAbsoluteMs(firstNonEmpty(
    merged.createdAtMs,
    merged.created_at_ms,
    recordSet.createdAtMs,
    recordSet.created_at_ms,
    Date.now(),
  ), 'meeting_app_live_evidence_package_time');
  return compactObject({
    type: 'meeting_app_live_evidence_package',
    schema: MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    id: firstNonEmpty(merged.packageId, merged.package_id, `meeting-app-live-evidence-${createdAtMs}`),
    createdAtMs,
    created_at_ms: createdAtMs,
    source: firstNonEmpty(merged.source, recordSet.source, 'meeting_app_live_evidence_package'),
    label: firstNonEmpty(merged.label, recordSet.label),
    accepted,
    production_ready: productionReady,
    platform_count: platforms.length,
    platforms,
    record_count: recordSet.record_count,
    record_set: recordSet,
    records_by_platform: Object.fromEntries(Object.entries(groupedRecords).map(([platform, records]) => [
      platform,
      {
        record_count: records.length,
        phases: uniqueList(records.map((record) => record.phase)),
        latest_record_id: records.at(-1)?.id,
      },
    ])),
    manifest_acceptance: manifestAcceptance,
    summary: {
      accepted,
      production_ready: productionReady,
      accepted_count: rows.filter((row) => row.accepted).length,
      production_ready_count: rows.filter((row) => row.production_ready).length,
      missing_evidence_count: rows.filter((row) => row.record_count === 0).length,
      rows,
    },
    issues: packageIssues,
    handoff: {
      validation_input: 'Pass record_set as { records } to meetingAppRuntimeAdapterValidation or meetingAppDeploymentManifestAcceptance.',
      success_condition: 'summary.production_ready === true',
    },
  });
}

export function buildMeetingAppLiveEvidencePackageSummary(input = {}, options = {}) {
  const evidencePackage = input?.type === 'meeting_app_live_evidence_package'
    ? input
    : buildMeetingAppLiveEvidencePackage(input, options);
  return {
    type: 'meeting_app_live_evidence_package_summary',
    schema: MEETING_APP_LIVE_EVIDENCE_PACKAGE_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    id: evidencePackage.id,
    accepted: evidencePackage.accepted,
    production_ready: evidencePackage.production_ready,
    platform_count: evidencePackage.platform_count,
    record_count: evidencePackage.record_count,
    accepted_count: evidencePackage.summary?.accepted_count ?? 0,
    production_ready_count: evidencePackage.summary?.production_ready_count ?? 0,
    rows: evidencePackage.summary?.rows ?? [],
    issues: evidencePackage.issues ?? [],
  };
}

export function buildMeetingAppRuntimeAdapterValidationReport(configOrPlatform = {}, options = {}) {
  const config = runtimeConfigFromInput(configOrPlatform, options);
  const configAcceptance = buildMeetingAppRuntimeAdapterAcceptanceReport(config, options);
  const platform = configAcceptance.platform;
  let launchGate = null;
  const issues = [...sourcedIssues('runtime_config', configAcceptance.issues)];
  if (configAcceptance.accepted && platform) {
    launchGate = buildMeetingAppLaunchGate(platform, validationOptions(options));
    issues.push(...sourcedIssues('launch_gate', [
      ...(launchGate.blocking_issues ?? []),
      ...(launchGate.warnings ?? []),
    ]));
  } else {
    issues.push(issue(
      'error',
      'runtime_config_not_accepted',
      'Runtime adapter config must pass acceptance before live validation can run.',
    ));
  }

  const accepted = configAcceptance.accepted === true && launchGate?.passed === true;
  const productionReady = configAcceptance.accepted === true && launchGate?.production_ready === true;
  return compactObject({
    type: 'meeting_app_runtime_adapter_validation_report',
    schema: MEETING_APP_RUNTIME_ADAPTER_CONFIG_SCHEMA,
    version: MEETING_APP_INTEGRATION_PROFILE_SCHEMA_VERSION,
    accepted,
    production_ready: productionReady,
    platform,
    evidence_level: launchGate?.evidence_level ?? 'none',
    evidence_count: launchGate?.evidence_count ?? 0,
    config_acceptance: configAcceptance,
    launch_gate: launchGate,
    next_actions: uniqueList([
      ...(productionReady ? [] : ['capture_live_dom_snapshots_for_this_platform']),
      ...(launchGate?.next_actions ?? []),
      ...issues.map((item) => item.code),
    ]),
    issues,
  });
}

export function assertMeetingAppRuntimeAdapterValidation(configOrPlatform = {}, options = {}) {
  const report = buildMeetingAppRuntimeAdapterValidationReport(configOrPlatform, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app runtime adapter validation failed', report);
  }
  return report;
}

export function buildAllMeetingAppRuntimeAdapterValidationReports(options = {}) {
  return Object.fromEntries(platformList(options).map((platform) => [
    platform,
    buildMeetingAppRuntimeAdapterValidationReport(platform, options),
  ]));
}

export default buildMeetingAppIntegrationProfile;
