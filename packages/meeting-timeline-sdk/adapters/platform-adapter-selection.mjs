import { compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import {
  buildMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from './platform-adapter-route.mjs';
import {
  buildMeetingPlatformAdaptationStrategy,
} from './platform-strategy.mjs';

export const MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA = 'meeting_platform_adapter_selection';
export const MEETING_PLATFORM_ADAPTER_SELECTION_MATRIX_SCHEMA = 'meeting_platform_adapter_selection_matrix';
export const MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA_VERSION = 1;

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function routeStep(route = {}, name) {
  return asArray(route.routes).find((item) => item.route === name);
}

function nonEmptyArray(value) {
  return Array.isArray(value) && value.length > 0;
}

function recordPlatform(record = {}) {
  return firstNonEmpty(
    record.platform,
    record.snapshot?.platform,
    record.candidate?.platform,
    record.raw?.platform,
  );
}

function recordSetMatchesPlatform(recordSet = {}, platform) {
  const rawPlatform = firstNonEmpty(recordSet.platform, recordSet.meeting?.platform, recordSet.meta?.platform);
  if (rawPlatform) return normalizeMeetingPlatform(rawPlatform) === platform;
  const records = firstNonEmpty(recordSet.records, recordSet.items, recordSet.snapshots);
  if (!Array.isArray(records) || records.length === 0) return true;
  return records.some((record) => {
    const value = recordPlatform(record);
    if (!value) return true;
    try {
      return normalizeMeetingPlatform(value) === platform;
    } catch {
      return false;
    }
  });
}

function hasRecordSet(platform, input = {}, options = {}) {
  const recordSet = firstNonEmpty(
    input.meetingAppRecordSet,
    input.meeting_app_record_set,
    input.record_set,
    options.meetingAppRecordSet,
    options.meeting_app_record_set,
  );
  return Boolean(recordSet && (
    nonEmptyArray(recordSet.records)
    || nonEmptyArray(recordSet.items)
    || nonEmptyArray(recordSet.snapshots)
  ) && recordSetMatchesPlatform(recordSet, platform));
}

function hasProviderRecords(input = {}, options = {}) {
  const records = firstNonEmpty(
    input.providerRecords,
    input.provider_records,
    input.provider_events,
    options.providerRecords,
    options.provider_records,
    options.provider_events,
  );
  return nonEmptyArray(records) || Boolean(records && nonEmptyArray(records.records));
}

function hasLocalDetectorRecords(input = {}, options = {}) {
  const records = firstNonEmpty(
    input.localDetectorRecords,
    input.local_detector_records,
    input.detectorRecords,
    input.detector_records,
    options.localDetectorRecords,
    options.local_detector_records,
  );
  return nonEmptyArray(records) || Boolean(records && nonEmptyArray(records.records));
}

function evidenceState(platform, input = {}, options = {}) {
  const hasMeetingAppRecordSet = hasRecordSet(platform, input, options);
  const hasProvider = hasProviderRecords(input, options);
  const hasLocalDetector = hasLocalDetectorRecords(input, options);
  return {
    has_meeting_app_record_set: hasMeetingAppRecordSet,
    has_provider_records: hasProvider,
    has_local_detector_records: hasLocalDetector,
    realtime_axis_evidence: platform === 'local_detector' ? hasLocalDetector : hasMeetingAppRecordSet,
    production_evidence: platform === 'local_detector' ? hasLocalDetector : hasMeetingAppRecordSet && hasProvider,
  };
}

function axisRouteName(platform) {
  return platform === 'local_detector' ? 'host_detector_axis' : 'local_observer_axis';
}

function selectedSurface(platform, route = {}) {
  if (platform === 'local_detector') return 'host_sdk';
  return route.adapter_surfaces?.primary
    ?? route.entrypoints?.browser_extension?.runtime_preset
    ?? route.routes?.[0]?.surfaces?.[0];
}

function sourcePlan(platform, route = {}, strategy = {}) {
  const axisName = axisRouteName(platform);
  const axis = routeStep(route, axisName);
  const annotation = routeStep(route, 'annotation_insert');
  const speaker = routeStep(route, 'speaker_position_markers');
  const provider = routeStep(route, 'provider_reconcile');
  const artifact = routeStep(route, 'post_meeting_artifact_import');
  return [
    compactObject({
      id: axisName,
      priority: 1,
      selected: true,
      role: axis?.role,
      surface: selectedSurface(platform, route),
      source: axis?.source,
      required_for_realtime: true,
      blocks_realtime_if_missing: true,
      timestamp_field: axis?.timestamp_field,
      evidence_input: axis?.evidence_input,
      start_condition: axis?.start_condition,
      end_condition: axis?.end_condition,
    }),
    compactObject({
      id: 'annotation_insert',
      priority: 2,
      selected: true,
      role: annotation?.role,
      required_for_realtime: true,
      blocks_realtime_if_missing: true,
      timestamp_field: annotation?.timestamp_field ?? 'captured_at_ms',
      invariant: annotation?.invariant,
      runtime_event_endpoint: annotation?.runtime_event_endpoint,
      supported_actions: annotation?.supported_actions,
    }),
    compactObject({
      id: 'speaker_position_markers',
      priority: 3,
      selected: Boolean(speaker),
      role: speaker?.role,
      required_for_realtime: false,
      blocks_realtime_if_missing: false,
      source: strategy.speaker_activity?.realtime_primary,
      fallback: strategy.speaker_activity?.backfill,
      filter: speaker?.filter,
      output: speaker?.output,
    }),
    compactObject({
      id: 'provider_reconcile',
      priority: 4,
      selected: Boolean(provider),
      role: provider?.role,
      required_for_realtime: false,
      required_for_production: provider?.required_for_production,
      blocks_realtime_if_missing: false,
      transport: provider?.transport,
      endpoint: provider?.endpoint,
      status_endpoint: provider?.status_endpoint,
      start_events: provider?.start_events,
      end_events: provider?.end_events,
      participant_events: provider?.participant_events,
      lifecycle_events: provider?.lifecycle_events,
      references: provider?.references,
    }),
    compactObject({
      id: 'post_meeting_artifact_import',
      priority: 5,
      selected: Boolean(artifact),
      role: artifact?.role,
      required_for_realtime: false,
      blocks_realtime_if_missing: false,
      realtime_dependency: false,
      transcript: artifact?.transcript,
      artifact_events: artifact?.artifact_events,
      import_endpoint: artifact?.import_endpoint,
      normalizer: artifact?.normalizer,
    }),
  ].filter(Boolean);
}

function runtimeActions(platform) {
  const actions = [
    'validate_raw_signal',
    platform === 'local_detector' ? 'receive_host_detector_signal' : 'observe_platform_candidates',
    platform === 'local_detector' ? 'open_host_detector_axis' : 'open_adapter_session',
    'insert_annotation',
    'append_speaker_position_marker_if_available',
    'reconcile_provider_event_when_available',
    'import_post_meeting_artifact_when_available',
  ];
  return platform === 'local_detector'
    ? actions.filter((action) => action !== 'reconcile_provider_event_when_available')
    : actions;
}

function readiness(platform, routeReadiness = {}, evidence = {}) {
  const selectionReady = routeReadiness.ready === true;
  return {
    selection_ready: selectionReady,
    route_ready: routeReadiness.ready === true,
    pilot_evidence_ready: selectionReady && evidence.realtime_axis_evidence === true,
    production_evidence_ready: selectionReady && evidence.production_evidence === true,
    provider_reconcile_required_for_production: platform !== 'local_detector',
    missing: routeReadiness.missing ?? [],
    issues: routeReadiness.issues ?? [],
  };
}

export function buildMeetingPlatformAdapterSelection(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const mergedOptions = { ...input, ...options };
  const route = buildMeetingPlatformAdapterRoute(key, mergedOptions);
  const routeReadiness = verifyMeetingPlatformAdapterRouteReadiness(route);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, mergedOptions);
  const capabilities = platformCapabilityContract(key, mergedOptions);
  const evidence = evidenceState(key, input, options);
  const sources = sourcePlan(key, route, strategy);
  const axisSource = sources.find((item) => item.id === axisRouteName(key));
  const providerSource = sources.find((item) => item.id === 'provider_reconcile');
  const artifactSource = sources.find((item) => item.id === 'post_meeting_artifact_import');
  const speakerSource = sources.find((item) => item.id === 'speaker_position_markers');

  return compactObject({
    type: 'meeting_platform_adapter_selection',
    schema: MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA_VERSION,
    platform: key,
    display_name: route.display_name ?? strategy.display_name ?? capabilities.display_name,
    recommended_mode: route.recommended_mode,
    objective: 'select_the_adapter_sources_a_host_should_wire_for_realtime_meeting_timeline_annotations',
    selection: {
      axis_source: axisSource?.id,
      axis_surface: axisSource?.surface,
      axis_source_role: axisSource?.role,
      annotation_source: 'annotation_insert',
      timestamp_field: axisSource?.timestamp_field ?? 'captured_at_ms',
      provider_reconcile_source: providerSource?.selected ? 'provider_reconcile' : undefined,
      provider_reconcile_required_for_production: key !== 'local_detector',
      speaker_track_source: speakerSource?.source,
      post_meeting_artifact_source: artifactSource?.selected ? 'post_meeting_artifact_import' : undefined,
    },
    runtime_policy: {
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      annotations_use_absolute_captured_at_ms: true,
      per_meeting_annotation_isolation_required: true,
      source_priority: strategy.source_priority,
      route_order: route.route_order,
      runtime_actions: runtimeActions(key),
      startup_order: [
        'validate_raw_signal',
        axisSource?.id,
        'annotation_insert',
        speakerSource?.selected ? 'speaker_position_markers' : undefined,
        providerSource?.selected ? 'provider_reconcile_non_blocking' : undefined,
        artifactSource?.selected ? 'post_meeting_artifact_import_non_blocking' : undefined,
      ].filter(Boolean),
    },
    sources,
    current_evidence: evidence,
    readiness: readiness(key, routeReadiness, evidence),
    gates: {
      pilot: strategy.evidence_contract?.pilot,
      production: strategy.evidence_contract?.production,
      handoff_package: strategy.evidence_contract?.handoff_package,
    },
    risk_profile: strategy.adaptation_playbook?.risk_profile,
    references: providerSource?.references,
    next_actions: unique([
      ...(strategy.next_actions ?? []),
      evidence.realtime_axis_evidence ? undefined : 'capture_live_local_observer_or_detector_axis_evidence',
      key !== 'local_detector' && !evidence.has_provider_records ? 'collect_provider_records_for_production_reconcile' : undefined,
      'wire_selected_adapter_sources_into_host_runtime',
    ]),
    route,
    strategy: options.includeStrategy === true || options.include_strategy === true ? strategy : undefined,
  });
}

export function buildMeetingPlatformAdapterSelectionMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms({ ...options, ...input });
  const selections = platforms.map((platform) => buildMeetingPlatformAdapterSelection(platform, input, options));
  return {
    type: 'meeting_platform_adapter_selection_matrix',
    schema: MEETING_PLATFORM_ADAPTER_SELECTION_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SELECTION_SCHEMA_VERSION,
    platform_count: selections.length,
    selection_ready_count: selections.filter((item) => item.readiness?.selection_ready === true).length,
    pilot_evidence_ready_count: selections.filter((item) => item.readiness?.pilot_evidence_ready === true).length,
    production_evidence_ready_count: selections.filter((item) => item.readiness?.production_evidence_ready === true).length,
    local_axis_selected_count: selections.filter((item) => (
      item.selection?.axis_source === 'local_observer_axis'
      || item.selection?.axis_source === 'host_detector_axis'
    )).length,
    provider_reconcile_count: selections.filter((item) => item.selection?.provider_reconcile_source === 'provider_reconcile').length,
    post_meeting_artifact_count: selections.filter((item) => item.selection?.post_meeting_artifact_source === 'post_meeting_artifact_import').length,
    platforms: selections.map((item) => item.platform),
    rows: selections.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      recommended_mode: item.recommended_mode,
      axis_source: item.selection?.axis_source,
      axis_surface: item.selection?.axis_surface,
      timestamp_field: item.selection?.timestamp_field,
      provider_reconcile_source: item.selection?.provider_reconcile_source,
      provider_reconcile_required_for_production: item.selection?.provider_reconcile_required_for_production,
      speaker_track_source: item.selection?.speaker_track_source,
      post_meeting_artifact_source: item.selection?.post_meeting_artifact_source,
      selection_ready: item.readiness?.selection_ready,
      pilot_evidence_ready: item.readiness?.pilot_evidence_ready,
      production_evidence_ready: item.readiness?.production_evidence_ready,
      provider_events_block_realtime: item.runtime_policy?.provider_events_block_realtime,
      transcript_blocks_realtime: item.runtime_policy?.transcript_blocks_realtime,
    })),
    selections,
  };
}
