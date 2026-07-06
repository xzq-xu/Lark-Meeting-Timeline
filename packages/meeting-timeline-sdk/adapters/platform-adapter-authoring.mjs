import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterContract,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformImplementationHandoff,
} from './platform-implementation-handoff.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_AUTHORING_PLAN_SCHEMA = 'meeting_platform_adapter_authoring_plan';
export const MEETING_PLATFORM_ADAPTER_AUTHORING_MATRIX_SCHEMA = 'meeting_platform_adapter_authoring_matrix';
export const MEETING_PLATFORM_ADAPTER_AUTHORING_SCHEMA_VERSION = 1;

const DEFAULT_AUTHORING_PLATFORMS = Object.freeze([
  'google_meet',
  'zoom',
  'microsoft_teams',
  'webex',
  'lark',
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

function slug(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function platformResolution(platform) {
  try {
    const key = normalizeMeetingPlatform(platform);
    return {
      platform: key,
      built_in: MEETING_PLATFORM_KEYS.includes(key),
      normalized_from: platform,
    };
  } catch {
    const key = slug(platform);
    if (!key) throw new Error('platform is required for meeting platform adapter authoring');
    return {
      platform: key,
      built_in: false,
      normalized_from: platform,
    };
  }
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_AUTHORING_PLATFORMS,
  )).map((platform) => platformResolution(platform).platform));
}

function displayName(platform, options = {}, builtInHandoff = {}) {
  const explicit = firstNonEmpty(options.displayName, options.display_name, options.name);
  if (explicit) return String(explicit);
  if (builtInHandoff.display_name) return builtInHandoff.display_name;
  return platform.split('_').map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join(' ');
}

function browserMatches(platform, options = {}, runtimeBundle = {}) {
  return unique(asArray(firstNonEmpty(
    options.browserMatches,
    options.browser_matches,
    options.matches,
    runtimeBundle.browser_matches,
    runtimeBundle.browser?.matches,
    [],
  )));
}

function providerPath(platform, options = {}, handoff = {}) {
  return firstNonEmpty(
    options.providerPath,
    options.provider_path,
    handoff.provider_reconcile?.path,
    platform === 'local_detector' ? 'host_detector_runtime' : `${platform}_provider_events`,
  );
}

function providerTransport(options = {}, handoff = {}) {
  return firstNonEmpty(
    options.providerTransport,
    options.provider_transport,
    options.transport,
    handoff.provider_reconcile?.transport,
    'webhook_or_polling',
  );
}

function authoringFiles(platform, builtIn = false) {
  const dashed = platform.replaceAll('_', '-');
  return [
    {
      path: `packages/meeting-timeline-sdk/adapters/${dashed}.mjs`,
      purpose: builtIn ? 'review_or_extend_existing_event_normalizer' : 'implement_provider_event_normalizer',
    },
    {
      path: `packages/meeting-timeline-sdk/adapters/${dashed}.d.ts`,
      purpose: builtIn ? 'review_or_extend_public_event_types' : 'declare_provider_event_normalizer_types',
    },
    {
      path: 'packages/meeting-timeline-sdk/adapters/platform-setup.mjs',
      purpose: 'add_aliases_capabilities_provider_events_and_post_meeting_artifact_contracts',
    },
    {
      path: 'packages/meeting-timeline-sdk/adapters/meeting-app-browser-runtime.mjs',
      purpose: 'add_runtime_preset_selectors_cadence_and_mutation_filters',
    },
    {
      path: 'packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs',
      purpose: 'add_active_prejoin_ended_dom_fixture_snapshots',
    },
    {
      path: 'packages/meeting-timeline-sdk/adapters/platform-registry.mjs',
      purpose: 'register_aliases_and_normalizer',
    },
    {
      path: `test/sdk-${dashed}.test.mjs`,
      purpose: 'verify_event_normalization_and_timestamp_mapping',
    },
    {
      path: `test/sdk-platform-${dashed}-runtime.test.mjs`,
      purpose: 'verify_candidate_observation_runtime_bundle_and_adapter_route',
    },
  ];
}

function authoringSteps(platform, builtIn) {
  return [
    {
      step: 1,
      id: 'define_platform_identity',
      output: 'normalized_platform_key_aliases_display_name',
      required_fields: ['platform', 'display_name', 'aliases'],
    },
    {
      step: 2,
      id: 'capture_local_surface',
      output: 'active_prejoin_ended_dom_snapshots_or_native_window_samples',
      required_for_realtime: true,
      evidence: 'meetingAppRecordSet',
    },
    {
      step: 3,
      id: 'author_runtime_preset',
      output: 'browser_or_webview_runtime_profile_with_candidate_observation',
      required_runtime_event: 'observe_platform_candidates',
    },
    {
      step: 4,
      id: 'author_provider_normalizer',
      output: 'provider_start_end_participant_artifact_events_normalized_to_timeline_signals',
      required_for_realtime: false,
      required_for_production_reconcile: true,
    },
    {
      step: 5,
      id: 'wire_nonblocking_annotation_flow',
      output: 'insert_annotation_uses_captured_at_ms_and_does_not_wait_for_provider_or_transcript',
      required_contracts: [
        'provider_events_block_realtime=false',
        'transcript_blocks_realtime=false',
        'timestamp_field=captured_at_ms',
      ],
    },
    {
      step: 6,
      id: 'add_tests_and_acceptance',
      output: builtIn ? 'extended_platform_tests_still_pass' : 'new_platform_contract_runtime_and_fixture_tests',
      required_commands: [
        'npm run meeting-platform:adapter-contract',
        'npm run meeting-platform:runtime-bundle',
        'npm run meeting-platform:implementation-handoff',
        'npm run sdk:test',
      ],
    },
  ];
}

function normalizerTemplate(platform) {
  return [
    `export function normalize${platform.split('_').map((part) => part.charAt(0).toUpperCase() + part.slice(1)).join('')}Event(input = {}) {`,
    '  const body = input.body ?? input;',
    '  return {',
    `    source: '${platform}_webhook',`,
    '    platform: body.platform,',
    '    type: body.type,',
    '    meeting: {',
    '      platform: body.platform,',
    '      meeting_id: body.meeting_id ?? body.id,',
    '      title: body.title,',
    '    },',
    '    occurred_at_ms: body.occurred_at_ms ?? body.time_ms,',
    '    raw: input,',
    '  };',
    '}',
  ].join('\n');
}

function commands(platform, options = {}, builtIn = false) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const baseArg = baseUrl ? ` --base-url=${baseUrl}` : '';
  const platformArg = `--platforms=${platform}`;
  return {
    authoring_plan: `npm run meeting-platform:adapter-authoring -- ${platformArg}${baseArg}`,
    adapter_contract: builtIn
      ? `npm run meeting-platform:adapter-contract -- ${platformArg}${baseArg}`
      : 'add platform to registry before running meeting-platform:adapter-contract',
    runtime_bundle: builtIn
      ? `npm run meeting-platform:runtime-bundle -- ${platformArg}${baseArg}`
      : 'add runtime preset before running meeting-platform:runtime-bundle',
    implementation_handoff: builtIn
      ? `npm run meeting-platform:implementation-handoff -- ${platformArg}${baseArg}`
      : 'add platform setup and runtime bundle before implementation handoff can pass',
    sdk_test: 'npm run sdk:test',
  };
}

function builtInArtifacts(platform, options = {}, builtIn = false) {
  if (!builtIn) return {};
  const handoff = buildMeetingPlatformImplementationHandoff(platform, options);
  const runtimeBundle = buildMeetingPlatformRuntimeBundle(platform, options);
  const contract = buildMeetingPlatformAdapterContract(platform, options);
  return {
    implementation_handoff: handoff,
    runtime_bundle_summary: {
      schema: runtimeBundle.schema,
      browser_matches: runtimeBundle.browser?.matches ?? [],
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      lightweight_connector_ready: runtimeBundle.readiness?.lightweight_connector_ready === true,
      runtime_event_endpoint: runtimeBundle.messaging?.runtime_event?.endpoint,
    },
    adapter_contract_summary: {
      schema: contract.schema,
      supported_surfaces: contract.supported_surfaces,
      timestamp_field: contract.timebase?.annotation_timestamp_field,
      provider_events_block_realtime: contract.timebase?.provider_events_block_realtime,
      transcript_blocks_realtime: contract.timebase?.transcript_blocks_realtime,
    },
  };
}

export function buildMeetingPlatformAdapterAuthoringPlan(platform, options = {}) {
  const resolution = platformResolution(platform);
  const builtIn = resolution.built_in;
  const artifacts = builtInArtifacts(resolution.platform, options, builtIn);
  const handoff = artifacts.implementation_handoff ?? {};
  const matches = browserMatches(resolution.platform, options, artifacts.runtime_bundle_summary);
  const provider = providerPath(resolution.platform, options, handoff);
  const transport = providerTransport(options, handoff);

  return compactObject({
    type: 'meeting_platform_adapter_authoring_plan',
    schema: MEETING_PLATFORM_ADAPTER_AUTHORING_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_AUTHORING_SCHEMA_VERSION,
    platform: resolution.platform,
    display_name: displayName(resolution.platform, options, handoff),
    built_in: builtIn,
    objective: builtIn
      ? 'extend_or_verify_an_existing_meeting_platform_adapter'
      : 'author_a_new_meeting_platform_adapter_for_realtime_annotation_timeline',
    current_sdk_status: builtIn ? 'built_in_static_sdk_path_available' : 'external_adapter_authoring_required',
    recommended_first_surface: firstNonEmpty(
      handoff.recommended_first_surface,
      matches.length > 0 ? 'browser_extension' : undefined,
      'native_detector',
    ),
    provider_reconcile: {
      path: provider,
      transport,
      required_for_realtime: false,
      role: 'reconcile_or_backfill_after_local_axis_exists',
    },
    required_contracts: {
      timestamp_field: 'captured_at_ms',
      local_axis_first: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      candidate_observation_message_type: 'meeting_timeline.observe_candidates',
    },
    browser_surface: {
      matches,
      needs_match_patterns: matches.length === 0,
      candidate_observation_required: true,
      required_permission: 'tabs',
    },
    files: authoringFiles(resolution.platform, builtIn),
    authoring_steps: authoringSteps(resolution.platform, builtIn),
    normalizer_template: builtIn ? undefined : normalizerTemplate(resolution.platform),
    commands: commands(resolution.platform, options, builtIn),
    built_in_artifacts: builtIn ? artifacts : undefined,
    production_evidence: {
      required_local_evidence: ['active_meeting_snapshot', 'ended_meeting_snapshot', 'realtime_annotation_sample'],
      required_provider_evidence: ['meeting_start', 'meeting_end'],
      required_runtime_evidence: ['candidate_observation', 'insert_annotation_with_captured_at_ms'],
    },
    risks: unique([
      matches.length === 0 ? 'missing_browser_match_patterns' : undefined,
      builtIn ? undefined : 'platform_not_registered_in_static_sdk',
      'provider_events_may_lag_and_must_not_block_realtime_marks',
      'transcript_or_recording_artifacts_are_post_meeting_only',
    ]),
    next_actions: builtIn
      ? ['collect_real_dom_and_provider_evidence', 'run_handoff_readiness_for_production']
      : ['add_platform_setup_entry', 'add_browser_runtime_preset', 'add_event_normalizer', 'add_fixture_and_contract_tests'],
  });
}

export function buildMeetingPlatformAdapterAuthoringMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterAuthoringPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_authoring_matrix',
    schema: MEETING_PLATFORM_ADAPTER_AUTHORING_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_AUTHORING_SCHEMA_VERSION,
    platform_count: plans.length,
    built_in_count: plans.filter((plan) => plan.built_in).length,
    external_authoring_count: plans.filter((plan) => !plan.built_in).length,
    browser_surface_ready_count: plans.filter((plan) => plan.browser_surface?.matches?.length > 0).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      built_in: plan.built_in,
      current_sdk_status: plan.current_sdk_status,
      recommended_first_surface: plan.recommended_first_surface,
      provider_path: plan.provider_reconcile?.path,
      provider_transport: plan.provider_reconcile?.transport,
      browser_match_count: plan.browser_surface?.matches?.length ?? 0,
      file_count: plan.files?.length ?? 0,
      risk_count: plan.risks?.length ?? 0,
      next_action: plan.next_actions?.[0],
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}
