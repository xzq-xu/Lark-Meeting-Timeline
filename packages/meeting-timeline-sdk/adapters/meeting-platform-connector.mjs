import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryEntry,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from './platform-registry.mjs';
import {
  buildMeetingPlatformRuntimeEvent,
  createMeetingPlatformRuntimeEventClient,
  normalizeMeetingPlatformRuntimeEventAction,
} from './platform-runtime-event.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import {
  createMeetingAppBrowserRuntime,
  meetingAppBrowserInput,
} from './meeting-app-browser-runtime.mjs';
import { createMeetingAppContentScriptBridge } from './meeting-app-content-script.mjs';

export const MEETING_PLATFORM_CONNECTOR_SCHEMA = 'meeting_platform_connector';
export const MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA = 'meeting_platform_connector_matrix';
export const MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA = 'meeting_platform_connector_acceptance';
export const MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA = 'meeting_platform_connector_runtime';
export const MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA = 'meeting_platform_connector_hub';
export const MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA = 'meeting_platform_connector_resolution';
export const MEETING_PLATFORM_CONNECTOR_BROWSER_RUNTIME_SCHEMA = 'meeting_platform_connector_browser_runtime';
export const MEETING_PLATFORM_CONNECTOR_CONTENT_SCRIPT_BRIDGE_SCHEMA = 'meeting_platform_connector_content_script_bridge';
export const MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION = 1;

const DEFAULT_CONNECTOR_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

const GLOBAL_RUNTIME_ACTIONS = new Set([
  'observe_platform_candidates',
  'adapter_routes',
  'runtime_bundles',
  'registry',
  'manifest',
  'run_manifest',
  'readiness',
  'handoff_readiness',
  'run_handoff_readiness',
]);

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_CONNECTOR_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function runtimeActionSet(connector = {}) {
  return new Set(connector.runtime_events?.supported_actions ?? []);
}

function maybeNormalizePlatform(value) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return undefined;
  }
}

function connectorPlatform(input, options = {}) {
  const value = firstNonEmpty(
    typeof input === 'string' ? input : undefined,
    input?.platform,
    input?.key,
    input?.provider,
    options.platform,
    options.platform_key,
  );
  if (!value) {
    throw new MeetingTimelineSdkError('Meeting platform connector requires platform', {
      required_field: 'platform',
    });
  }
  return normalizeMeetingPlatform(value);
}

function isConnector(value) {
  return value?.schema === MEETING_PLATFORM_CONNECTOR_SCHEMA;
}

function connectorFrom(input, options = {}) {
  if (isConnector(input)) return input;
  if (isConnector(options.connector)) return options.connector;
  return buildMeetingPlatformConnector(input, options);
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths = []) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function explicitPlatformFrom(input = {}, options = {}) {
  return maybeNormalizePlatform(firstNonEmpty(
    options.platform,
    options.platform_key,
    options.provider,
    typeof input === 'string' ? undefined : input?.platform,
    typeof input === 'string' ? undefined : input?.platform_key,
    typeof input === 'string' ? undefined : input?.provider,
    typeof input === 'string' ? undefined : input?.adapter,
    typeof input === 'string' ? undefined : input?.meeting?.platform,
    typeof input === 'string' ? undefined : input?.current_meeting?.platform,
    typeof input === 'string' ? undefined : input?.currentMeeting?.platform,
    typeof input === 'string' ? undefined : input?.detectedMeeting?.platform,
    typeof input === 'string' ? undefined : input?.detected_meeting?.platform,
    typeof input === 'string' ? undefined : input?.snapshot?.platform,
  ));
}

function scoreCandidate(candidate = {}, index = 0) {
  const active = candidate.active === true || candidate.current === true || candidate.selected === true ? 100 : 0;
  const inMeeting = candidate.in_meeting === true || candidate.inMeeting === true || candidate.meeting?.in_meeting === true ? 40 : 0;
  const audible = candidate.audible === true || candidate.has_audio === true || candidate.hasAudio === true ? 10 : 0;
  return active + inMeeting + audible - index;
}

function compactCandidate(value = {}) {
  if (typeof value === 'string' || value instanceof URL) return { url: String(value) };
  if (!isPlainObject(value)) return {};
  return compactObject({
    platform: firstPath(value, ['platform', 'provider', 'adapter', 'meeting.platform']),
    url: firstPath(value, [
      'meeting.meeting_url',
      'meeting.meetingUrl',
      'meeting.url',
      'meeting.join_url',
      'meeting.joinUrl',
      'meeting_url',
      'meetingUrl',
      'join_url',
      'joinUrl',
      'url',
      'href',
      'window.url',
      'browser.url',
      'tab.url',
    ]),
    title: firstPath(value, [
      'meeting.title',
      'meeting.topic',
      'meeting.name',
      'title',
      'topic',
      'name',
      'window.title',
      'browser.title',
      'tab.title',
    ]),
    active: firstPath(value, ['active', 'current', 'selected', 'tab.active', 'window.active']),
    audible: firstPath(value, ['audible', 'has_audio', 'hasAudio', 'tab.audible']),
    in_meeting: firstPath(value, ['in_meeting', 'inMeeting', 'meeting.in_meeting', 'meeting.inMeeting']),
  });
}

function routeCandidates(input = {}, options = {}) {
  const candidates = [];
  const push = (value) => {
    const candidate = compactCandidate(value);
    if (candidate.url || candidate.platform || candidate.title) candidates.push(candidate);
  };
  push(input);
  push(options);
  if (isPlainObject(input)) {
    push(input.tab);
    push(input.window);
    push(input.browser);
    push(input.meeting);
    push(input.current_meeting);
    push(input.currentMeeting);
    for (const tab of asArray(input.tabs)) push(tab);
    for (const window of asArray(input.windows)) {
      push(window);
      for (const tab of asArray(window?.tabs)) push({
        ...tab,
        window: compactCandidate(window),
      });
    }
  }
  return candidates
    .map((candidate, index) => ({ ...candidate, score: scoreCandidate(candidate, index) }))
    .sort((left, right) => right.score - left.score);
}

function supportedPlatformSet(options = {}) {
  return new Set(selectedPlatforms(options));
}

export function resolveMeetingPlatformConnectorInput(input = {}, options = {}) {
  const supportedPlatforms = supportedPlatformSet(options);
  const explicitPlatform = explicitPlatformFrom(input, options);
  if (explicitPlatform) {
    return compactObject({
      type: 'meeting_platform_connector_resolution',
      schema: MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA,
      schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
      detected: true,
      supported: supportedPlatforms.has(explicitPlatform),
      platform: explicitPlatform,
      reason: 'explicit_platform',
      current_platforms: [...supportedPlatforms],
    });
  }

  if (typeof input === 'string') {
    const stringPlatform = maybeNormalizePlatform(input);
    if (stringPlatform) {
      return compactObject({
        type: 'meeting_platform_connector_resolution',
        schema: MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA,
        schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
        detected: true,
        supported: supportedPlatforms.has(stringPlatform),
        platform: stringPlatform,
        reason: 'platform_string',
        current_platforms: [...supportedPlatforms],
      });
    }
  }

  const candidates = routeCandidates(input, options);
  for (const candidate of candidates) {
    const detectedMeeting = detectMeetingFromUrl(candidate);
    const platform = maybeNormalizePlatform(detectedMeeting?.platform);
    if (!platform) continue;
    return compactObject({
      type: 'meeting_platform_connector_resolution',
      schema: MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA,
      schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
      detected: true,
      supported: supportedPlatforms.has(platform),
      platform,
      reason: 'meeting_url',
      current_platforms: [...supportedPlatforms],
      meeting: detectedMeeting,
      browser: {
        url: candidate.url,
        title: candidate.title,
        active: candidate.active,
        audible: candidate.audible,
        in_meeting: candidate.in_meeting,
      },
      candidate_count: candidates.length,
    });
  }

  return compactObject({
    type: 'meeting_platform_connector_resolution',
    schema: MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    detected: false,
    supported: false,
    current_platforms: [...supportedPlatforms],
    candidate_count: candidates.length,
    next_actions: ['pass_explicit_platform_or_supported_meeting_url'],
  });
}

function runtimeEventEndpoint(entry = {}, options = {}) {
  return firstNonEmpty(
    options.endpoint,
    options.runtimeEventEndpoint,
    options.runtime_event_endpoint,
    entry.annotations?.runtime_event_endpoint,
    entry.annotations?.runtime_event_plan?.endpoint,
  );
}

function connectorAcceptanceIssues(connector = {}, options = {}) {
  const issues = [];
  const requireProviderReady = options.requireProviderReady === true || options.require_provider_ready === true;
  const actions = runtimeActionSet(connector);
  if (connector.schema !== MEETING_PLATFORM_CONNECTOR_SCHEMA) {
    issues.push(issue('error', 'invalid_connector_schema', 'Connector schema must be meeting_platform_connector.', {
      actual_schema: connector.schema,
    }));
  }
  if (!connector.platform) {
    issues.push(issue('error', 'missing_platform', 'Connector must include a normalized platform key.'));
  }
  if (connector.event_adapter?.normalize_available !== true) {
    issues.push(issue('error', 'missing_event_normalizer', 'Connector must expose a provider event normalizer.', {
      platform: connector.platform,
    }));
  }
  if (connector.runtime_events?.endpoint == null) {
    issues.push(issue('error', 'missing_runtime_event_endpoint', 'Connector must include runtime event endpoint.', {
      platform: connector.platform,
    }));
  }
  for (const action of ['observe_meeting_app', 'observe_platform_candidates', 'insert_annotation', 'speaker_track', 'participant_track', 'provider_event']) {
    if (!actions.has(action)) {
      issues.push(issue('error', 'missing_runtime_action', `Connector runtime event plan must include ${action}.`, {
        platform: connector.platform,
        action,
      }));
    }
  }
  if (connector.timeline_ingest?.timestamp_field !== 'captured_at_ms') {
    issues.push(issue('error', 'invalid_timestamp_field', 'Connector must use captured_at_ms for realtime annotation insertion.', {
      platform: connector.platform,
      timestamp_field: connector.timeline_ingest?.timestamp_field,
    }));
  }
  if (!connector.timeline_ingest?.insert_endpoint) {
    issues.push(issue('error', 'missing_insert_endpoint', 'Connector must include annotation insert endpoint.', {
      platform: connector.platform,
    }));
  }
  if (connector.realtime_policy?.provider_events_block_realtime === true) {
    issues.push(issue('error', 'provider_blocks_realtime', 'Provider events must not block realtime annotation insertion.', {
      platform: connector.platform,
    }));
  }
  if (connector.realtime_policy?.transcript_blocks_realtime === true) {
    issues.push(issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion.', {
      platform: connector.platform,
    }));
  }
  if (connector.browser_observer?.enabled === true && (connector.browser_observer?.matches?.length ?? 0) === 0) {
    issues.push(issue('error', 'missing_browser_matches', 'Browser observer connector must include URL match patterns.', {
      platform: connector.platform,
    }));
  }
  if (connector.browser_observer?.candidate_observation_ready !== true) {
    issues.push(issue('error', 'candidate_observation_not_ready', 'Connector must expose candidate observation for host-level axis binding.', {
      platform: connector.platform,
    }));
  }
  if (connector.browser_observer?.candidate_message_type !== 'meeting_timeline.observe_candidates') {
    issues.push(issue('error', 'invalid_candidate_observation_message_type', 'Connector must use meeting_timeline.observe_candidates for candidate observation.', {
      platform: connector.platform,
      message_type: connector.browser_observer?.candidate_message_type,
    }));
  }
  if (connector.browser_observer?.required_permission !== 'tabs') {
    issues.push(issue('error', 'invalid_candidate_observation_permission', 'Connector must declare tabs permission for candidate observation.', {
      platform: connector.platform,
      required_permission: connector.browser_observer?.required_permission,
    }));
  }
  if (!connector.adapter_route?.first_route || !connector.adapter_route?.recommended_mode) {
    issues.push(issue('error', 'missing_adapter_route', 'Connector must include adapter route planning.', {
      platform: connector.platform,
    }));
  }
  if (requireProviderReady && connector.provider?.ready !== true) {
    issues.push(issue('error', 'provider_not_ready', 'Provider setup must be ready when requireProviderReady=true.', {
      platform: connector.platform,
      missing_env: connector.provider?.missing_env ?? [],
    }));
  }
  return issues;
}

export function buildMeetingPlatformConnector(platformOrOptions = {}, options = {}) {
  const inputOptions = typeof platformOrOptions === 'string' ? options : { ...platformOrOptions, ...options };
  const platform = connectorPlatform(platformOrOptions, inputOptions);
  const entry = buildMeetingPlatformRegistryEntry(platform, inputOptions);
  const registryAcceptance = buildMeetingPlatformRegistryAcceptanceReport(entry, inputOptions);
  const actions = entry.annotations?.runtime_event_plan?.supported_actions ?? [];
  const runtimeEndpoint = runtimeEventEndpoint(entry, inputOptions);

  return compactObject({
    type: 'meeting_platform_connector',
    schema: MEETING_PLATFORM_CONNECTOR_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    platform,
    display_name: entry.display_name,
    aliases: entry.aliases,
    base_url: firstNonEmpty(inputOptions.baseUrl, inputOptions.base_url),
    objective: 'single_platform_connector_for_realtime_meeting_timeline_annotations',
    event_adapter: entry.event_adapter,
    provider: {
      transport: entry.provider?.transport,
      endpoint: entry.provider?.endpoint,
      status_endpoint: entry.provider?.status_endpoint,
      role: entry.provider?.role,
      ready: entry.provider?.ready,
      required_for_realtime: false,
      security_verifier: entry.provider?.security_verifier,
      missing_env: entry.provider?.missing_env ?? [],
      start_events: entry.provider?.start_events ?? [],
      end_events: entry.provider?.end_events ?? [],
      participant_events: entry.provider?.participant_events ?? [],
      artifact_events: entry.provider?.artifact_events ?? [],
      lifecycle_events: entry.provider?.lifecycle_events ?? [],
      official_docs: entry.provider?.official_docs,
    },
    browser_observer: {
      enabled: (entry.runtime?.browser_matches?.length ?? 0) > 0,
      matches: entry.runtime?.browser_matches ?? [],
      host_permissions: entry.runtime?.host_permissions ?? [],
      content_script_count: entry.runtime?.content_script_count ?? 0,
      runtime_preset: entry.runtime?.preset,
      sample_interval_ms: entry.runtime?.sample_interval_ms,
      mutation_debounce_ms: entry.runtime?.mutation_debounce_ms,
      speaker_min_stable_ms: entry.runtime?.speaker_min_stable_ms,
      candidate_observation_ready: entry.runtime?.candidate_observation?.ready === true,
      candidate_message_type: entry.runtime?.candidate_observation?.message_type,
      required_permission: entry.runtime?.candidate_observation?.required_permission,
      candidate_endpoint: entry.runtime?.candidate_observation?.endpoint,
      candidate_runtime_action: entry.runtime?.candidate_observation?.runtime_event_action,
      candidate_runtime_method: entry.runtime?.candidate_observation?.runtime_event_client_method,
    },
    runtime_events: {
      endpoint: runtimeEndpoint,
      event_schema: entry.annotations?.runtime_event_plan?.event_schema,
      client_factory: entry.annotations?.runtime_event_plan?.client_factory,
      supported_actions: actions,
      action_count: actions.length,
      realtime_contract: entry.annotations?.runtime_event_plan?.realtime_contract,
    },
    timeline_ingest: {
      insert_endpoint: entry.annotations?.insert_endpoint,
      timestamp_field: entry.annotations?.timestamp_field,
      runtime_action: 'insert_annotation',
      annotation_should_use_device_capture_time: true,
    },
    realtime_policy: {
      route: entry.adapter_route?.recommended_mode,
      primary_axis: entry.adapter_route?.first_route,
      provider_events_block_realtime: entry.annotations?.provider_events_block_realtime === true,
      transcript_blocks_realtime: entry.annotations?.transcript_blocks_realtime === true,
      local_observer_may_start_axis: entry.annotations?.local_observer_may_start_axis === true,
      provider_events_role: 'reconcile_and_backfill_after_local_axis',
      transcript_role: 'post_meeting_import_only',
    },
    adapter_route: entry.adapter_route,
    host: entry.host,
    sdk: {
      ...entry.sdk,
      imports: {
        ...(entry.sdk?.imports ?? {}),
        connector: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
      },
      factories: {
        build_connector: 'buildMeetingPlatformConnector',
        create_runtime: 'createMeetingPlatformConnectorRuntime',
        assert_connector: 'assertMeetingPlatformConnector',
      },
    },
    readiness: {
      accepted: registryAcceptance.accepted,
      realtime_annotation_ready: registryAcceptance.accepted
        && runtimeEndpoint != null
        && actions.includes('insert_annotation')
        && actions.includes('observe_platform_candidates')
        && entry.annotations?.provider_events_block_realtime !== true
        && entry.annotations?.transcript_blocks_realtime !== true,
      registry_acceptance_schema: registryAcceptance.schema,
      registry_blocking_count: registryAcceptance.blocking_count,
      provider_ready: entry.provider?.ready === true,
      provider_required_for_realtime: entry.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: entry.readiness?.transcript_blocks_realtime === true,
      candidate_observation_ready: entry.readiness?.candidate_observation_ready === true,
      missing_items: entry.readiness?.missing_items ?? [],
    },
    commands: entry.commands,
    next_actions: unique([
      'create_platform_connector_runtime_in_host_project',
      'wire_observe_platform_candidates_before_relying_on_provider_events',
      'send_insert_annotation_with_captured_at_ms',
      ...(entry.next_actions ?? []),
    ]),
  });
}

export function buildMeetingPlatformConnectorAcceptanceReport(connectorOrOptions = {}, options = {}) {
  const connector = isConnector(connectorOrOptions)
    ? connectorOrOptions
    : buildMeetingPlatformConnector(connectorOrOptions, options);
  const issues = connectorAcceptanceIssues(connector, options);
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    type: 'meeting_platform_connector_acceptance',
    schema: MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    platform: connector.platform,
    blocking_count: blocking.length,
    warning_count: issues.length - blocking.length,
    issues,
    connector,
    next_actions: unique([
      ...blocking.map((item) => item.code),
      ...(connector.next_actions ?? []),
    ]),
  };
}

export function assertMeetingPlatformConnector(connectorOrOptions = {}, options = {}) {
  const report = buildMeetingPlatformConnectorAcceptanceReport(connectorOrOptions, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform connector failed acceptance', {
      issues: report.issues,
      report,
    });
  }
  return report;
}

export function buildMeetingPlatformConnectorMatrix(options = {}) {
  const connectors = selectedPlatforms(options).map((platform) => buildMeetingPlatformConnector(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  const acceptanceReports = connectors.map((connector) => buildMeetingPlatformConnectorAcceptanceReport(connector, options));
  return {
    type: 'meeting_platform_connector_matrix',
    schema: MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    platform_count: connectors.length,
    accepted_count: acceptanceReports.filter((report) => report.accepted).length,
    realtime_ready_count: connectors.filter((connector) => connector.readiness?.realtime_annotation_ready).length,
    provider_ready_count: connectors.filter((connector) => connector.provider?.ready).length,
    candidate_observer_count: connectors.filter((connector) => connector.browser_observer?.candidate_observation_ready).length,
    platforms: connectors.map((connector) => connector.platform),
    rows: connectors.map((connector) => ({
      platform: connector.platform,
      display_name: connector.display_name,
      accepted: acceptanceReports.find((report) => report.platform === connector.platform)?.accepted === true,
      realtime_annotation_ready: connector.readiness?.realtime_annotation_ready === true,
      provider_ready: connector.provider?.ready === true,
      provider_transport: connector.provider?.transport,
      browser_observer_enabled: connector.browser_observer?.enabled === true,
      browser_match_count: connector.browser_observer?.matches?.length ?? 0,
      candidate_observation_ready: connector.browser_observer?.candidate_observation_ready === true,
      runtime_event_endpoint: connector.runtime_events?.endpoint,
      runtime_action_count: connector.runtime_events?.action_count ?? 0,
      insert_endpoint: connector.timeline_ingest?.insert_endpoint,
      adapter_route: connector.adapter_route?.recommended_mode,
      first_route: connector.adapter_route?.first_route,
    })),
    connectors,
    acceptance_reports: acceptanceReports,
    registry_manifest: buildMeetingPlatformRegistryManifest(options),
    next_actions: unique(connectors.flatMap((connector) => connector.next_actions ?? [])),
  };
}

function connectorsByPlatform(connectors = []) {
  return new Map(asArray(connectors).map((connector) => [connector.platform, connector]));
}

function connectorHubIssues(hub = {}) {
  const issues = [];
  if (hub.matrix?.accepted_count !== hub.matrix?.platform_count) {
    issues.push(issue('error', 'connector_matrix_not_accepted', 'Every selected platform connector must pass acceptance.', {
      accepted_count: hub.matrix?.accepted_count,
      platform_count: hub.matrix?.platform_count,
    }));
  }
  if (hub.matrix?.realtime_ready_count !== hub.matrix?.platform_count) {
    issues.push(issue('error', 'realtime_connector_not_ready', 'Every selected platform connector must be ready for realtime annotation insertion.', {
      realtime_ready_count: hub.matrix?.realtime_ready_count,
      platform_count: hub.matrix?.platform_count,
    }));
  }
  if (hub.matrix?.candidate_observer_count !== hub.matrix?.platform_count) {
    issues.push(issue('error', 'candidate_observer_not_ready', 'Every selected platform connector must expose candidate observation.', {
      candidate_observer_count: hub.matrix?.candidate_observer_count,
      platform_count: hub.matrix?.platform_count,
    }));
  }
  return issues;
}

export function buildMeetingPlatformConnectorHub(options = {}) {
  const matrix = buildMeetingPlatformConnectorMatrix(options);
  const defaultPlatform = firstNonEmpty(options.defaultPlatform, options.default_platform, matrix.platforms[0]);
  const hub = compactObject({
    type: 'meeting_platform_connector_hub',
    schema: MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    objective: 'multi_platform_connector_router_for_realtime_meeting_timeline_annotations',
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    realtime_ready_count: matrix.realtime_ready_count,
    candidate_observer_count: matrix.candidate_observer_count,
    platforms: matrix.platforms,
    default_platform: maybeNormalizePlatform(defaultPlatform) ?? matrix.platforms[0],
    runtime_event_endpoint: firstNonEmpty(
      options.endpoint,
      options.runtimeEventEndpoint,
      options.runtime_event_endpoint,
      matrix.connectors[0]?.runtime_events?.endpoint,
    ),
    routing: {
      input_priority: [
        'explicit platform/provider',
        'meeting URL on input',
        'active tab URL',
        'active window tab URL',
      ],
      supported_input_fields: [
        'platform',
        'provider',
        'meeting.platform',
        'url',
        'meeting_url',
        'tab.url',
        'tabs[].url',
        'windows[].tabs[].url',
      ],
      url_detection_ready: true,
      detected_platforms: ['google_meet', 'microsoft_teams', 'zoom', 'webex', 'lark']
        .filter((platform) => matrix.platforms.includes(platform)),
      browser_observer_required_permission: 'tabs',
    },
    matrix,
    connectors: matrix.connectors,
    readiness: {
      accepted: matrix.accepted_count === matrix.platform_count,
      realtime_annotation_ready: matrix.realtime_ready_count === matrix.platform_count,
      candidate_observation_ready: matrix.candidate_observer_count === matrix.platform_count,
      provider_required_for_realtime: false,
      transcript_blocks_realtime: false,
    },
    next_actions: unique([
      'create_connector_hub_in_host_project',
      'route_insert_annotation_by_current_meeting_url_or_platform',
      'wire_observe_platform_candidates_for_browser_tabs',
      ...matrix.next_actions,
    ]),
  });
  const issues = connectorHubIssues(hub);
  return {
    ...hub,
    accepted: issues.filter((item) => item.severity === 'error').length === 0,
    blocking_count: issues.filter((item) => item.severity === 'error').length,
    warning_count: issues.filter((item) => item.severity !== 'error').length,
    issues,
  };
}

function runtimeEndpointFrom(connector = {}, options = {}) {
  return firstNonEmpty(
    options.endpoint,
    options.runtimeEventEndpoint,
    options.runtime_event_endpoint,
    connector.runtime_events?.endpoint,
  );
}

function actionFor(input = {}, options = {}) {
  const action = firstNonEmpty(
    input.action,
    input.kind,
    input.type,
    input.runtime_action,
    input.runtimeAction,
    options.action,
    options.kind,
  );
  return normalizeMeetingPlatformRuntimeEventAction(action);
}

function withDefaultPlatform(connector, input = {}, options = {}) {
  const action = actionFor(input, options);
  const platform = firstNonEmpty(input.platform, options.platform, connector.platform);
  return {
    ...input,
    action,
    platform: GLOBAL_RUNTIME_ACTIONS.has(action) ? input.platform : platform,
  };
}

export function createMeetingPlatformConnectorRuntime(platformOrConnector, options = {}) {
  const connector = connectorFrom(platformOrConnector, options);
  if (options.assertConnector !== false && options.assert_connector !== false) {
    assertMeetingPlatformConnector(connector, options);
  }
  const endpoint = runtimeEndpointFrom(connector, options);
  if (!endpoint) {
    throw new MeetingTimelineSdkError('Meeting platform connector runtime requires endpoint', {
      platform: connector.platform,
    });
  }
  const runtimeEventClient = createMeetingPlatformRuntimeEventClient({
    ...options,
    endpoint,
  });
  const adapter = meetingPlatformEventAdapterFor(connector.platform);
  const actionSet = runtimeActionSet(connector);

  function supports(action) {
    try {
      return actionSet.has(normalizeMeetingPlatformRuntimeEventAction(action));
    } catch {
      return false;
    }
  }

  function assertSupported(action) {
    const normalized = normalizeMeetingPlatformRuntimeEventAction(action);
    if (!actionSet.has(normalized)) {
      throw new MeetingTimelineSdkError('Runtime action is not present in meeting platform connector', {
        platform: connector.platform,
        action: normalized,
        supported_actions: connector.runtime_events?.supported_actions ?? [],
      });
    }
    return normalized;
  }

  return {
    type: 'meeting_platform_connector_runtime',
    schema: MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    platform: connector.platform,
    endpoint,
    connector,
    runtime_event_client: runtimeEventClient,
    supported_actions: connector.runtime_events?.supported_actions ?? [],
    supports,
    assertSupported,
    normalizeProviderEvent(raw = {}, normalizeOptions = {}) {
      if (typeof adapter?.normalize !== 'function') {
        throw new MeetingTimelineSdkError('Meeting platform connector has no provider normalizer', {
          platform: connector.platform,
        });
      }
      return adapter.normalize(raw, normalizeOptions);
    },
    buildEvent(input = {}, eventOptions = {}) {
      const eventInput = withDefaultPlatform(connector, input, eventOptions);
      assertSupported(eventInput.action);
      return buildMeetingPlatformRuntimeEvent(eventInput, {
        ...options,
        ...eventOptions,
      });
    },
    send(input = {}, sendOptions = {}) {
      const event = this.buildEvent(input, sendOptions);
      return runtimeEventClient.send(event, sendOptions);
    },
    observeMeetingApp(snapshot = {}, observeOptions = {}) {
      assertSupported('observe_meeting_app');
      return runtimeEventClient.observeMeetingApp(connector.platform, snapshot, observeOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      assertSupported('observe_platform_candidates');
      return runtimeEventClient.observePlatformCandidates(input, observeOptions);
    },
    ingestProvider(payload = {}, ingestOptions = {}) {
      assertSupported('provider_event');
      return runtimeEventClient.ingestProvider(connector.platform, payload, ingestOptions);
    },
    insertAnnotation(annotationInput = {}, markOptions = {}) {
      assertSupported('insert_annotation');
      return runtimeEventClient.insertAnnotation(connector.platform, annotationInput, markOptions);
    },
    insertMark(annotationInput = {}, markOptions = {}) {
      return this.insertAnnotation(annotationInput, markOptions);
    },
    speakerTrack(input = {}, trackOptions = {}) {
      assertSupported('speaker_track');
      return runtimeEventClient.speakerTrack(connector.platform, input, trackOptions);
    },
    participantTrack(input = {}, trackOptions = {}) {
      assertSupported('participant_track');
      return runtimeEventClient.participantTrack(connector.platform, input, trackOptions);
    },
    timelineView(input = {}, viewOptions = {}) {
      assertSupported('timeline_view');
      return runtimeEventClient.timelineView(connector.platform, input, viewOptions);
    },
    adapterRoute(routeOptions = {}) {
      assertSupported('adapter_route');
      return runtimeEventClient.adapterRoute(connector.platform, routeOptions);
    },
    adapterRoutes(routeOptions = {}) {
      assertSupported('adapter_routes');
      return runtimeEventClient.adapterRoutes({ platforms: [connector.platform], ...routeOptions });
    },
    runtimeBundles(bundleOptions = {}) {
      assertSupported('runtime_bundles');
      return runtimeEventClient.runtimeBundles({ platforms: [connector.platform], ...bundleOptions });
    },
    registry(registryOptions = {}) {
      assertSupported('registry');
      return runtimeEventClient.registry({ platforms: [connector.platform], ...registryOptions });
    },
    manifest(manifestOptions = {}) {
      assertSupported('manifest');
      return runtimeEventClient.manifest({ platforms: [connector.platform], ...manifestOptions });
    },
    readiness(readinessOptions = {}) {
      assertSupported('readiness');
      return runtimeEventClient.readiness({ platforms: [connector.platform], ...readinessOptions });
    },
    handoffReadiness(readinessOptions = {}) {
      assertSupported('handoff_readiness');
      return runtimeEventClient.handoffReadiness({ platforms: [connector.platform], ...readinessOptions });
    },
    runManifest(manifestOptions = {}) {
      assertSupported('run_manifest');
      return runtimeEventClient.runManifest({ platforms: [connector.platform], ...manifestOptions });
    },
    runHandoffReadiness(readinessOptions = {}) {
      assertSupported('run_handoff_readiness');
      return runtimeEventClient.runHandoffReadiness({ platforms: [connector.platform], ...readinessOptions });
    },
  };
}

function runtimeOptionsWithoutRoute(options = {}) {
  const {
    platform,
    platform_key: platformKey,
    provider,
    adapter,
    connector,
    ...rest
  } = options;
  void platform;
  void platformKey;
  void provider;
  void adapter;
  void connector;
  return rest;
}

export function createMeetingPlatformConnectorHub(options = {}) {
  const hub = buildMeetingPlatformConnectorHub(options);
  const connectorMap = connectorsByPlatform(hub.connectors);
  const runtimeCache = new Map();

  function resolutionFor(input = {}, resolveOptions = {}) {
    return resolveMeetingPlatformConnectorInput(input, {
      ...options,
      ...resolveOptions,
      platforms: hub.platforms,
      platform_keys: undefined,
    });
  }

  function connectorFor(input = {}, connectorOptions = {}) {
    if (isConnector(input)) return input;
    const platform = maybeNormalizePlatform(typeof input === 'string' ? input : undefined)
      ?? resolutionFor(input, connectorOptions).platform
      ?? maybeNormalizePlatform(connectorOptions.platform);
    if (!platform) {
      throw new MeetingTimelineSdkError('Unable to resolve meeting platform connector', {
        input,
        resolution: resolutionFor(input, connectorOptions),
      });
    }
    const connector = connectorMap.get(platform);
    if (!connector) {
      throw new MeetingTimelineSdkError('Meeting platform connector is not enabled in hub', {
        platform,
        platforms: hub.platforms,
      });
    }
    return connector;
  }

  function runtimeFor(input = {}, runtimeOptions = {}) {
    const connector = connectorFor(input, runtimeOptions);
    const cacheable = Object.keys(runtimeOptions).length === 0;
    if (cacheable && runtimeCache.has(connector.platform)) return runtimeCache.get(connector.platform);
    const runtime = createMeetingPlatformConnectorRuntime(connector, {
      ...runtimeOptionsWithoutRoute(options),
      ...runtimeOptionsWithoutRoute(runtimeOptions),
    });
    if (cacheable) runtimeCache.set(connector.platform, runtime);
    return runtime;
  }

  function defaultRuntime(runtimeOptions = {}) {
    return runtimeFor(hub.default_platform, runtimeOptions);
  }

  return {
    type: 'meeting_platform_connector_hub_runtime',
    schema: MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    hub,
    matrix: hub.matrix,
    platforms: hub.platforms,
    default_platform: hub.default_platform,
    endpoint: hub.runtime_event_endpoint,
    connectors: hub.connectors,
    resolvePlatform: resolutionFor,
    connectorFor,
    runtimeFor,
    supports(input = {}, action) {
      const connector = connectorFor(input);
      return runtimeActionSet(connector).has(normalizeMeetingPlatformRuntimeEventAction(action));
    },
    normalizeProviderEvent(input = {}, raw = {}, normalizeOptions = {}) {
      return runtimeFor(input, normalizeOptions).normalizeProviderEvent(raw, normalizeOptions);
    },
    buildEvent(input = {}, eventOptions = {}) {
      return runtimeFor(input, eventOptions).buildEvent(input, eventOptions);
    },
    send(input = {}, sendOptions = {}) {
      return runtimeFor(input, sendOptions).send(input, sendOptions);
    },
    observeMeetingApp(snapshot = {}, observeOptions = {}) {
      return runtimeFor(snapshot, observeOptions).observeMeetingApp(snapshot, observeOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return defaultRuntime(observeOptions).observePlatformCandidates(input, observeOptions);
    },
    ingestProvider(input = {}, payload = undefined, ingestOptions = {}) {
      const providerPayload = payload === undefined ? input : payload;
      return runtimeFor(input, ingestOptions).ingestProvider(providerPayload, ingestOptions);
    },
    insertAnnotation(input = {}, annotation = undefined, markOptions = {}) {
      const annotationInput = annotation === undefined ? input : annotation;
      return runtimeFor(input, markOptions).insertAnnotation(annotationInput, markOptions);
    },
    insertMark(input = {}, annotation = undefined, markOptions = {}) {
      return this.insertAnnotation(input, annotation, markOptions);
    },
    speakerTrack(input = {}, track = undefined, trackOptions = {}) {
      const trackInput = track === undefined ? input : track;
      return runtimeFor(input, trackOptions).speakerTrack(trackInput, trackOptions);
    },
    participantTrack(input = {}, track = undefined, trackOptions = {}) {
      const trackInput = track === undefined ? input : track;
      return runtimeFor(input, trackOptions).participantTrack(trackInput, trackOptions);
    },
    timelineView(input = {}, view = undefined, viewOptions = {}) {
      const viewInput = view === undefined ? input : view;
      return runtimeFor(input, viewOptions).timelineView(viewInput, viewOptions);
    },
    adapterRoute(input = {}, routeOptions = {}) {
      return runtimeFor(input, routeOptions).adapterRoute(routeOptions);
    },
    adapterRoutes(routeOptions = {}) {
      return defaultRuntime(routeOptions).adapterRoutes(routeOptions);
    },
    runtimeBundles(bundleOptions = {}) {
      return defaultRuntime(bundleOptions).runtimeBundles(bundleOptions);
    },
    registry(registryOptions = {}) {
      return defaultRuntime(registryOptions).registry(registryOptions);
    },
    manifest(manifestOptions = {}) {
      return defaultRuntime(manifestOptions).manifest(manifestOptions);
    },
    readiness(readinessOptions = {}) {
      return defaultRuntime(readinessOptions).readiness(readinessOptions);
    },
    handoffReadiness(readinessOptions = {}) {
      return defaultRuntime(readinessOptions).handoffReadiness(readinessOptions);
    },
    runManifest(manifestOptions = {}) {
      return defaultRuntime(manifestOptions).runManifest(manifestOptions);
    },
    runHandoffReadiness(readinessOptions = {}) {
      return defaultRuntime(readinessOptions).runHandoffReadiness(readinessOptions);
    },
  };
}

function browserBaseOptions(options = {}) {
  return {
    window: options.window,
    win: options.win,
    document: options.document,
    doc: options.doc,
    location: options.location,
    navigator: options.navigator,
  };
}

function connectorRouteInput(input = {}, options = {}) {
  const browserInput = meetingAppBrowserInput(browserBaseOptions(options));
  if (!isPlainObject(input)) return browserInput;
  return compactObject({
    ...browserInput,
    ...input,
    meeting: {
      ...(browserInput.meeting ?? {}),
      ...(input.meeting ?? {}),
    },
  });
}

function resultOk(result) {
  return result?.ok !== false && result?.error == null;
}

function createConnectorBrowserSources(hub, options = {}) {
  let lastResult = null;
  let insertCount = 0;
  let observeCount = 0;

  async function insertOne(input = {}, insertOptions = {}) {
    const route = connectorRouteInput(input, options);
    const result = await hub.insertAnnotation(route, input, insertOptions);
    insertCount += 1;
    lastResult = result;
    return result;
  }

  const client = {
    insertMark: insertOne,
    insertAnnotation: insertOne,
    async insertMarks(inputs = [], insertOptions = {}) {
      const rows = asArray(inputs);
      const results = [];
      for (const input of rows) results.push(await insertOne(input, insertOptions));
      return {
        ok: results.every(resultOk),
        result_count: results.length,
        results,
      };
    },
  };

  const sources = {
    client,
    async observeMeetingApp(input = {}, observeOptions = {}) {
      const snapshot = connectorRouteInput(input, options);
      const result = await hub.observeMeetingApp(snapshot, observeOptions);
      observeCount += 1;
      lastResult = result;
      return compactObject({
        ok: resultOk(result),
        runtime_event_result: result,
        platform_resolution: hub.resolvePlatform(snapshot, observeOptions),
      });
    },
    observeApp(input = {}, observeOptions = {}) {
      return this.observeMeetingApp(input, observeOptions);
    },
    observeBrowser(input = {}, observeOptions = {}) {
      return this.observeMeetingApp(input, observeOptions);
    },
    observeNative(input = {}, observeOptions = {}) {
      return this.observeMeetingApp(input, observeOptions);
    },
    observeLocal(input = {}, observeOptions = {}) {
      return this.observeMeetingApp(input, observeOptions);
    },
    insertMark: client.insertMark,
    insertAnnotation: client.insertAnnotation,
    insertMarks: client.insertMarks,
    async ingestProvider(platformOrInput = {}, payload = undefined, ingestOptions = {}) {
      const route = typeof platformOrInput === 'string'
        ? platformOrInput
        : connectorRouteInput(platformOrInput, options);
      const providerPayload = payload === undefined ? platformOrInput : payload;
      const result = await hub.ingestProvider(route, providerPayload, ingestOptions);
      lastResult = result;
      return result;
    },
    ingestSignals(signals = [], ingestOptions = {}) {
      return this.ingestProvider(connectorRouteInput(ingestOptions, options), {
        source: 'local_detector',
        signals: asArray(signals),
      }, ingestOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return this.ingestProvider(connectorRouteInput(input, options), {
        source: 'transcript_import',
        transcript: input,
      }, transcriptOptions);
    },
    startMeeting(input = {}) {
      return this.observeMeetingApp({
        ...input,
        in_meeting: true,
        lifecycle_hint: 'meeting_started',
      });
    },
    endMeeting(input = {}) {
      return this.observeMeetingApp({
        ...input,
        in_meeting: false,
        lifecycle_hint: 'meeting_ended',
      });
    },
    getState() {
      return {
        observe_count: observeCount,
        insert_count: insertCount,
        last_result: lastResult,
      };
    },
    reset() {
      lastResult = null;
      insertCount = 0;
      observeCount = 0;
      return this.getState();
    },
  };
  return sources;
}

export function createMeetingPlatformConnectorBrowserRuntime(options = {}) {
  const hub = options.hub
    ?? options.connectorHub
    ?? options.connector_hub
    ?? createMeetingPlatformConnectorHub(options);
  const sources = options.sources
    ?? createConnectorBrowserSources(hub, options);
  const browserRuntime = createMeetingAppBrowserRuntime({}, {
    ...options,
    sources,
  });

  return {
    ...browserRuntime,
    type: 'meeting_platform_connector_browser_runtime',
    schema: MEETING_PLATFORM_CONNECTOR_BROWSER_RUNTIME_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    hub,
    connector_hub: hub,
    sources,
    browserInput(input = {}) {
      return connectorRouteInput(input, options);
    },
    resolvePlatform(input = {}, resolveOptions = {}) {
      return hub.resolvePlatform(connectorRouteInput(input, options), resolveOptions);
    },
    currentWindowPreflight(input = {}, preflightOptions = {}) {
      return browserRuntime.currentWindowPreflight?.(connectorRouteInput(input, options), preflightOptions);
    },
    preflightCurrentWindow(input = {}, preflightOptions = {}) {
      return this.currentWindowPreflight(input, preflightOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return hub.observePlatformCandidates(connectorRouteInput(input, options), observeOptions);
    },
    observeMeetingApp(input = {}, observeOptions = {}) {
      return sources.observeMeetingApp(input, observeOptions);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return sources.insertAnnotation(input, markOptions);
    },
    insertMark(input = {}, markOptions = {}) {
      return sources.insertMark(input, markOptions);
    },
    insertMarks(inputs = [], markOptions = {}) {
      return sources.insertMarks(inputs, markOptions);
    },
    getState() {
      return {
        ...browserRuntime.getState(),
        connector_hub: {
          platforms: hub.platforms,
          default_platform: hub.default_platform,
          endpoint: hub.endpoint,
        },
        browser_resolution: hub.resolvePlatform(connectorRouteInput({}, options)),
        connector_sources: sources.getState?.(),
      };
    },
    reset(nextState = {}) {
      return {
        browser_runtime: browserRuntime.reset?.(nextState.browser_runtime ?? nextState.browserRuntime ?? {}),
        connector_sources: sources.reset?.(nextState.sources ?? nextState.connector_sources ?? {}),
      };
    },
  };
}

export function createMeetingPlatformConnectorContentScriptBridge(options = {}) {
  const runtime = options.runtime
    ?? options.browserRuntime
    ?? options.browser_runtime
    ?? createMeetingPlatformConnectorBrowserRuntime(options);
  const bridge = createMeetingAppContentScriptBridge({}, {
    ...options,
    runtime,
  });

  return {
    ...bridge,
    type: 'meeting_platform_connector_content_script_bridge',
    schema: MEETING_PLATFORM_CONNECTOR_CONTENT_SCRIPT_BRIDGE_SCHEMA,
    schema_version: MEETING_PLATFORM_CONNECTOR_SCHEMA_VERSION,
    runtime,
    hub: runtime.hub,
    connector_hub: runtime.hub,
    resolvePlatform(input = {}, resolveOptions = {}) {
      return runtime.resolvePlatform?.(input, resolveOptions);
    },
    currentWindowPreflight(input = {}, preflightOptions = {}) {
      return runtime.currentWindowPreflight?.(input, preflightOptions);
    },
    preflightCurrentWindow(input = {}, preflightOptions = {}) {
      return this.currentWindowPreflight(input, preflightOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return runtime.observePlatformCandidates?.(input, observeOptions);
    },
    getState() {
      return {
        ...bridge.getState(),
        connector_hub: {
          platforms: runtime.hub?.platforms,
          default_platform: runtime.hub?.default_platform,
          endpoint: runtime.hub?.endpoint,
        },
        browser_resolution: runtime.resolvePlatform?.(),
      };
    },
    dispose() {
      return bridge.dispose();
    },
  };
}

export function installMeetingPlatformConnectorContentScriptBridge(options = {}) {
  const bridge = createMeetingPlatformConnectorContentScriptBridge(options);
  bridge.start(options.startOptions ?? options.start_options ?? {});
  return bridge;
}

export function buildDefaultMeetingPlatformConnectorMatrix(options = {}) {
  return buildMeetingPlatformConnectorMatrix({
    platforms: MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector'),
    ...options,
  });
}

export function buildDefaultMeetingPlatformConnectorHub(options = {}) {
  return buildMeetingPlatformConnectorHub({
    platforms: MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector'),
    ...options,
  });
}
