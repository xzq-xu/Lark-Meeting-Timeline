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

export const MEETING_PLATFORM_CONNECTOR_SCHEMA = 'meeting_platform_connector';
export const MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA = 'meeting_platform_connector_matrix';
export const MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA = 'meeting_platform_connector_acceptance';
export const MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA = 'meeting_platform_connector_runtime';
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

export function buildDefaultMeetingPlatformConnectorMatrix(options = {}) {
  return buildMeetingPlatformConnectorMatrix({
    platforms: MEETING_PLATFORM_KEYS.filter((platform) => platform !== 'local_detector'),
    ...options,
  });
}
