import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
} from './meeting-app-extension.mjs';
import {
  buildMeetingAppAdapterManifest,
  buildMeetingAppAdapterManifestMatrix,
} from './meeting-app-adapter-manifest.mjs';

export const MEETING_APP_ADAPTER_SPEC_SCHEMA = 'meeting_app_adapter_spec';
export const MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA = 'meeting_app_adapter_spec_matrix';
export const MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION = 1;

const DEFAULT_BUILT_IN_PLATFORMS = Object.freeze([
  'google-meet',
  'teams',
  'zoom',
  'webex',
  'lark',
]);

const DEFAULT_EXTENSION_PERMISSIONS = Object.freeze(['storage', 'tabs']);
const DEFAULT_REQUIRED_SIGNALS = Object.freeze(['meeting_started', 'meeting_ended', 'speaker_started']);
const DEFAULT_MUTATION_IGNORE_SELECTORS = Object.freeze([
  '.caption-line',
  '.captions-line',
  '[data-caption-line]',
  '[data-transcript-line]',
  '[data-chat-message]',
  '[class*="caption" i]',
  '[class*="transcript" i]',
  '[class*="chat" i]',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function plainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function boolValue(value, fallback = false) {
  if (typeof value === 'boolean') return value;
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'on'].includes(text)) return true;
  if (['false', '0', 'no', 'off'].includes(text)) return false;
  return fallback;
}

function numberValue(value, fallback) {
  if (value == null || value === '') return fallback;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function adapterKey(value) {
  const text = String(value ?? '').trim().toLowerCase();
  if (!text) return '';
  return text
    .replace(/https?:\/\//g, '')
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .replace(/_{2,}/g, '_');
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function maybeBuiltInManifest(input = {}, options = {}) {
  const platform = typeof input === 'string'
    ? input
    : firstNonEmpty(input.platform, input.provider, input.key, input.adapter_key, input.name);
  if (!platform) return null;
  if (plainObject(input) && (
    input.matches
    || input.url_matches
    || input.control_selectors
    || input.controlSelectors
    || input.participant_selectors
    || input.participantSelectors
  )) {
    return null;
  }
  try {
    return buildMeetingAppAdapterManifest(platform, options);
  } catch {
    return null;
  }
}

function fromBuiltInManifest(manifest = {}, options = {}) {
  return {
    adapter_key: manifest.platform,
    platform: manifest.platform,
    display_name: manifest.display_name,
    source: 'built_in_manifest',
    surface: manifest.surface,
    matches: manifest.extension?.matches,
    host_permissions: manifest.extension?.host_permissions,
    permissions: manifest.extension?.permissions,
    control_selectors: manifest.capture?.control_selectors,
    participant_selectors: manifest.capture?.participant_selectors,
    text_selectors: manifest.capture?.text_selectors,
    observe_mutations: manifest.runtime?.observe_mutations,
    mutation_debounce_ms: manifest.runtime?.mutation_debounce_ms,
    speaker_stable_followup_ms: manifest.runtime?.speaker_stable_followup_ms,
    sample_interval_ms: manifest.runtime?.sample_interval_ms,
    unchanged_observe_every_ms: manifest.runtime?.unchanged_observe_every_ms,
    mutation_track_selectors: manifest.runtime?.mutation_track_selectors,
    mutation_ignore_selectors: manifest.runtime?.mutation_ignore_selectors,
    provider_events_block_realtime: manifest.contracts?.provider_events_block_realtime,
    transcript_blocks_realtime: manifest.contracts?.transcript_blocks_realtime,
    required_signals: manifest.contracts?.required_signals,
    host_endpoints: manifest.host_endpoints,
    built_in_manifest: options.includeBuiltInManifest || options.include_built_in_manifest ? manifest : undefined,
  };
}

function selectorCounts(spec = {}) {
  return {
    control: asArray(spec.control_selectors ?? spec.controlSelectors).length,
    participant: asArray(spec.participant_selectors ?? spec.participantSelectors).length,
    text: asArray(spec.text_selectors ?? spec.textSelectors).length,
    mutation_track: asArray(spec.mutation_track_selectors ?? spec.mutationTrackSelectors).length,
    mutation_ignore: asArray(spec.mutation_ignore_selectors ?? spec.mutationIgnoreSelectors).length,
  };
}

function normalizedSpecInput(specOrPlatform = {}, options = {}) {
  const builtIn = maybeBuiltInManifest(specOrPlatform, options);
  if (builtIn) return fromBuiltInManifest(builtIn, options);
  const input = typeof specOrPlatform === 'string'
    ? { platform: specOrPlatform }
    : { ...(specOrPlatform ?? {}) };
  return {
    ...input,
    ...options,
    adapter_key: firstNonEmpty(
      input.adapter_key,
      input.adapterKey,
      input.platform,
      input.provider,
      input.key,
      input.name,
      options.adapter_key,
      options.adapterKey,
      options.platform,
      options.provider,
      options.key,
      options.name,
    ),
  };
}

function buildIssues(spec = {}, counts = {}) {
  const matches = asArray(spec.matches ?? spec.url_matches ?? spec.urlMatches);
  const hostPermissions = asArray(spec.host_permissions ?? spec.hostPermissions ?? matches);
  const permissions = asArray(spec.permissions ?? DEFAULT_EXTENSION_PERMISSIONS);
  const requiredSignals = asArray(spec.required_signals ?? spec.requiredSignals ?? DEFAULT_REQUIRED_SIGNALS);
  const timestampField = firstNonEmpty(spec.timestamp_field, spec.timestampField, 'captured_at_ms');
  const providerBlocksRealtime = boolValue(firstNonEmpty(
    spec.provider_events_block_realtime,
    spec.providerEventsBlockRealtime,
    spec.provider?.blocks_realtime,
    false,
  ));
  const transcriptBlocksRealtime = boolValue(firstNonEmpty(
    spec.transcript_blocks_realtime,
    spec.transcriptBlocksRealtime,
    spec.transcript?.blocks_realtime,
    false,
  ));
  return [
    adapterKey(spec.adapter_key) ? undefined : issue('error', 'missing_adapter_key', 'Adapter spec requires a stable adapter_key.'),
    matches.length > 0 ? undefined : issue('error', 'missing_url_matches', 'Adapter spec requires at least one URL match pattern.'),
    matches.includes('<all_urls>') || hostPermissions.includes('<all_urls>')
      ? issue('error', 'overbroad_host_permission', 'Adapter spec must not require <all_urls>.')
      : undefined,
    permissions.includes('tabs') ? undefined : issue('error', 'missing_tabs_permission', 'Adapter spec needs tabs for active meeting candidate observation.'),
    permissions.includes('storage') ? undefined : issue('error', 'missing_storage_permission', 'Adapter spec needs storage for diagnostics and status.'),
    counts.control > 0 ? undefined : issue('error', 'missing_control_selectors', 'Adapter spec requires control selectors that indicate in-meeting state.'),
    counts.participant > 0 ? undefined : issue('error', 'missing_participant_selectors', 'Adapter spec requires participant or speaker selectors.'),
    counts.text > 0 ? undefined : issue('warning', 'missing_text_selectors', 'Adapter spec should include title/status text selectors.'),
    counts.mutation_track > 0 ? undefined : issue('error', 'missing_mutation_track_selectors', 'Adapter spec requires mutation track selectors for low-latency local observation.'),
    timestampField === 'captured_at_ms'
      ? undefined
      : issue('error', 'wrong_timestamp_field', 'Realtime annotations must use captured_at_ms.', { timestamp_field: timestampField }),
    providerBlocksRealtime
      ? issue('error', 'provider_blocks_realtime', 'Provider events must not block realtime annotation placement.')
      : undefined,
    transcriptBlocksRealtime
      ? issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation placement.')
      : undefined,
    requiredSignals.includes('meeting_started') ? undefined : issue('error', 'missing_meeting_started_signal', 'Adapter spec must produce or infer meeting_started.'),
    requiredSignals.includes('meeting_ended') ? undefined : issue('error', 'missing_meeting_ended_signal', 'Adapter spec must produce or infer meeting_ended.'),
    requiredSignals.includes('speaker_started') ? undefined : issue('warning', 'missing_speaker_started_signal', 'Adapter spec should support speaker_started for timeline speaker marks.'),
  ].filter(Boolean);
}

export function buildMeetingAppAdapterSpec(specOrPlatform = {}, options = {}) {
  const input = normalizedSpecInput(specOrPlatform, options);
  const key = adapterKey(input.adapter_key);
  const matches = unique(asArray(input.matches ?? input.url_matches ?? input.urlMatches));
  const hostPermissions = unique(asArray(input.host_permissions ?? input.hostPermissions ?? matches));
  const permissions = unique(asArray(input.permissions ?? DEFAULT_EXTENSION_PERMISSIONS));
  const controlSelectors = unique(asArray(input.control_selectors ?? input.controlSelectors));
  const participantSelectors = unique(asArray(input.participant_selectors ?? input.participantSelectors));
  const textSelectors = unique(asArray(input.text_selectors ?? input.textSelectors));
  const mutationTrackSelectors = unique(asArray(firstNonEmpty(
    input.mutation_track_selectors,
    input.mutationTrackSelectors,
    [...controlSelectors, ...participantSelectors, ...textSelectors],
  )));
  const mutationIgnoreSelectors = unique(asArray(firstNonEmpty(
    input.mutation_ignore_selectors,
    input.mutationIgnoreSelectors,
    DEFAULT_MUTATION_IGNORE_SELECTORS,
  )));
  const counts = selectorCounts({
    control_selectors: controlSelectors,
    participant_selectors: participantSelectors,
    text_selectors: textSelectors,
    mutation_track_selectors: mutationTrackSelectors,
    mutation_ignore_selectors: mutationIgnoreSelectors,
  });
  const issues = buildIssues({
    ...input,
    adapter_key: key,
    matches,
    host_permissions: hostPermissions,
    permissions,
    required_signals: input.required_signals ?? input.requiredSignals ?? DEFAULT_REQUIRED_SIGNALS,
  }, counts);
  const blocking = issues.filter((item) => item.severity === 'error');
  const requiredSignals = unique(asArray(input.required_signals ?? input.requiredSignals ?? DEFAULT_REQUIRED_SIGNALS));
  const hostEndpoints = input.host_endpoints ?? input.hostEndpoints ?? MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS;
  return {
    type: 'meeting_app_adapter_spec',
    schema: MEETING_APP_ADAPTER_SPEC_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    adapter_key: key,
    platform: key,
    display_name: firstNonEmpty(input.display_name, input.displayName, input.name, key),
    source: input.source ?? 'custom_spec',
    surface: firstNonEmpty(input.surface, 'browser_extension_or_webview'),
    url_detection: {
      matches,
      match_count: matches.length,
    },
    extension: {
      matches,
      host_permissions: hostPermissions,
      permissions,
      message_types: {
        observe_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
        extension_attached: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
        extension_status: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
        client_call: MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
      },
      content_script_entry: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script',
    },
    capture: {
      control_selectors: controlSelectors,
      participant_selectors: participantSelectors,
      text_selectors: textSelectors,
      selector_counts: counts,
    },
    runtime: {
      observe_mutations: boolValue(firstNonEmpty(input.observe_mutations, input.observeMutations, true), true),
      mutation_debounce_ms: numberValue(firstNonEmpty(input.mutation_debounce_ms, input.mutationDebounceMs), 180),
      speaker_stable_followup_ms: numberValue(firstNonEmpty(input.speaker_stable_followup_ms, input.speakerStableFollowupMs), 350),
      sample_interval_ms: numberValue(firstNonEmpty(input.sample_interval_ms, input.sampleIntervalMs), 10_000),
      unchanged_observe_every_ms: numberValue(firstNonEmpty(input.unchanged_observe_every_ms, input.unchangedObserveEveryMs), 10_000),
      mutation_track_selectors: mutationTrackSelectors,
      mutation_ignore_selectors: mutationIgnoreSelectors,
    },
    contracts: {
      timestamp_field: 'captured_at_ms',
      candidate_observation_message_type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
      local_observer_may_start_axis: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      required_signals: requiredSignals,
      output_intents: ['realtime_axis', 'speaker_track', 'participant_track', 'annotation_insert'],
    },
    host_endpoints: hostEndpoints,
    implementation_steps: [
      'match_active_meeting_url_or_window',
      'capture_control_participant_and_status_fields_without_screenshots',
      'emit_observe_candidates_with_captured_at_ms',
      'write_meeting_started_and_meeting_ended_from_local_state',
      'write_speaker_started_marks_after_stability_filter',
      'use_provider_events_and_transcripts_only_for_reconcile_or_backfill',
      'collect_live_dom_snapshots_before_production_rollout',
    ],
    readiness: {
      accepted: blocking.length === 0,
      url_match_ready: matches.length > 0,
      extension_permission_ready: permissions.includes('tabs') && permissions.includes('storage') && !hostPermissions.includes('<all_urls>'),
      capture_selector_ready: counts.control > 0 && counts.participant > 0,
      mutation_observer_ready: mutationTrackSelectors.length > 0,
      timestamp_ready: true,
      provider_nonblocking: true,
      transcript_nonblocking: true,
      requires_live_snapshot_before_production: true,
    },
    built_in_manifest: input.built_in_manifest,
    issue_count: issues.length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: unique([
      ...(blocking.map((item) => item.code)),
      'capture_live_dom_snapshots_before_production_rollout',
    ]),
  };
}

function specInputs(options = {}) {
  const adapters = asArray(firstNonEmpty(options.adapters, options.adapter_specs, options.adapterSpecs, options.specs));
  const builtInPlatforms = asArray(firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys, DEFAULT_BUILT_IN_PLATFORMS));
  const includeBuiltIns = firstNonEmpty(options.includeBuiltIns, options.include_built_ins, builtInPlatforms.length > 0) !== false;
  return [
    ...(includeBuiltIns ? builtInPlatforms : []),
    ...adapters,
  ];
}

export function buildMeetingAppAdapterSpecMatrix(options = {}) {
  const specs = specInputs(options).map((spec) => buildMeetingAppAdapterSpec(spec, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
    platformKeys: undefined,
    adapters: undefined,
    adapter_specs: undefined,
    adapterSpecs: undefined,
    specs: undefined,
  }));
  return {
    type: 'meeting_app_adapter_spec_matrix',
    schema: MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_SPEC_SCHEMA_VERSION,
    accepted: specs.every((spec) => spec.accepted === true),
    spec_count: specs.length,
    accepted_count: specs.filter((spec) => spec.accepted === true).length,
    custom_count: specs.filter((spec) => spec.source !== 'built_in_manifest').length,
    built_in_count: specs.filter((spec) => spec.source === 'built_in_manifest').length,
    url_match_ready_count: specs.filter((spec) => spec.readiness.url_match_ready).length,
    capture_selector_ready_count: specs.filter((spec) => spec.readiness.capture_selector_ready).length,
    mutation_observer_ready_count: specs.filter((spec) => spec.readiness.mutation_observer_ready).length,
    rows: specs.map((spec) => ({
      adapter_key: spec.adapter_key,
      display_name: spec.display_name,
      source: spec.source,
      accepted: spec.accepted,
      url_match_count: spec.url_detection.match_count,
      control_selector_count: spec.capture.selector_counts.control,
      participant_selector_count: spec.capture.selector_counts.participant,
      mutation_track_selector_count: spec.capture.selector_counts.mutation_track,
      blocking_count: spec.blocking_count,
      warning_count: spec.warning_count,
      first_next_action: spec.next_actions?.[0],
    })),
    specs,
    next_actions: unique(specs.flatMap((spec) => spec.next_actions ?? [])),
  };
}

export function buildMeetingAppAdapterSpecTemplate(input = {}) {
  const key = adapterKey(firstNonEmpty(input.adapter_key, input.adapterKey, input.platform, input.name, 'custom_meeting_app'));
  return {
    adapter_key: key,
    display_name: firstNonEmpty(input.display_name, input.displayName, input.name, 'Custom Meeting App'),
    matches: asArray(input.matches ?? input.url_matches ?? [`https://example.com/${key}/*`]),
    host_permissions: asArray(input.host_permissions ?? input.hostPermissions ?? input.matches ?? input.url_matches ?? [`https://example.com/${key}/*`]),
    permissions: [...DEFAULT_EXTENSION_PERMISSIONS],
    control_selectors: ['[aria-label*="Leave" i]', '[role="button"]'],
    participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
    text_selectors: ['[role="status"]', '[aria-live]'],
    mutation_track_selectors: ['[role="button"]', '[data-participant-id]', '[aria-live]'],
    mutation_ignore_selectors: [...DEFAULT_MUTATION_IGNORE_SELECTORS],
    required_signals: [...DEFAULT_REQUIRED_SIGNALS],
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
  };
}

export function assertMeetingAppAdapterSpec(specOrPlatform = {}, options = {}) {
  const spec = specOrPlatform?.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
    ? specOrPlatform
    : buildMeetingAppAdapterSpec(specOrPlatform, options);
  if (spec.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter spec acceptance failed', {
      code: 'meeting_app_adapter_spec_rejected',
      adapter_key: spec.adapter_key,
      issues: spec.issues,
      next_actions: spec.next_actions,
    });
  }
  return spec;
}

export function assertMeetingAppAdapterSpecMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_SPEC_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterSpecMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter spec matrix acceptance failed', {
      code: 'meeting_app_adapter_spec_matrix_rejected',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}

export function builtInMeetingAppAdapterSpecMatrix(options = {}) {
  const manifestMatrix = buildMeetingAppAdapterManifestMatrix({
    ...options,
    platforms: firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_BUILT_IN_PLATFORMS),
  });
  return buildMeetingAppAdapterSpecMatrix({
    ...options,
    includeBuiltIns: false,
    adapters: manifestMatrix.manifests.map((manifest) => fromBuiltInManifest(manifest, options)),
  });
}
