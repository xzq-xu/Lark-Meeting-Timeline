import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { meetingAppBrowserRuntimePreset } from './meeting-app-browser-runtime.mjs';
import { meetingAppDomCaptureProfile } from './meeting-app-capture.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLaunchGate } from './meeting-app-gate.mjs';
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
  if (config.extension?.message_types?.client_call !== MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call) {
    issues.push(issue('error', 'missing_client_call_message_type', 'Runtime adapter config is missing the client_call message type.'));
  }
  if (config.extension?.message_types?.extension_attached !== MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached) {
    issues.push(issue('error', 'missing_attached_message_type', 'Runtime adapter config is missing the extension_attached message type.'));
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
      message_types: Boolean(config.extension?.message_types?.client_call && config.extension?.message_types?.extension_attached),
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
    },
    integration_targets: [
      {
        surface: 'chrome_or_edge_extension',
        injection: 'manifest_v3_content_script',
        matches: extensionInstallPlan.matches,
        host_permissions: extensionInstallPlan.host_permissions,
        content_script_adapter: extensionInstallPlan.content_script_adapter,
        background_bridge: 'meeting_app_extension_background_forwarder',
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
