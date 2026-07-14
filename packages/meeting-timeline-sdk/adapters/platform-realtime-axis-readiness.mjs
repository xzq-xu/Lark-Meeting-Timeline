import { MeetingTimelineSdkError, compactObject } from './internal-utils.mjs';
import {
  buildMeetingPlatformAdapterBlueprint,
} from './platform-adapter-blueprint.mjs';
import {
  buildMeetingPlatformAdapterPreflight,
} from './platform-adapter-preflight.mjs';
import {
  buildMeetingPlatformAdapterSelection,
} from './platform-adapter-selection.mjs';
import {
  buildMeetingPlatformProviderReplayReport,
} from './platform-ingest.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA = 'meeting_platform_realtime_axis_readiness';
export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_MATRIX_SCHEMA = 'meeting_platform_realtime_axis_readiness_matrix';
export const MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA_VERSION = 1;

const DEFAULT_REALTIME_AXIS_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
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

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function normalizeMaybePlatform(value) {
  if (!value) return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return String(value);
  }
}

function selectedPlatforms(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    input.platforms,
    input.platform_keys,
    DEFAULT_REALTIME_AXIS_PLATFORMS,
  )).map((platform) => normalizeMaybePlatform(platform)).filter(Boolean));
}

function platformAliases(platform) {
  return unique([
    platform,
    String(platform).replaceAll('_', '-'),
    platform === 'microsoft_teams' ? 'teams' : undefined,
    platform === 'google_meet' ? 'google-meet' : undefined,
  ]);
}

function inputForPlatform(platform, input = {}, options = {}) {
  const sources = [
    input.inputs,
    input.inputByPlatform,
    input.input_by_platform,
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
  ];
  for (const source of sources) {
    if (!isPlainObject(source)) continue;
    for (const key of platformAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(input.input, input.snapshot, input.sample, options.input, options.snapshot, options.sample, {});
}

function evidenceInput(input = {}) {
  return compactObject({
    snapshots: input.snapshots,
    domSnapshots: input.domSnapshots,
    dom_snapshots: input.dom_snapshots,
    recordSet: input.recordSet,
    record_set: input.record_set,
    snapshotRecords: input.snapshotRecords,
    snapshot_records: input.snapshot_records,
    records: input.records,
    meetingAppRecordSet: input.meetingAppRecordSet,
    meeting_app_record_set: input.meeting_app_record_set,
    meetingAppRecords: input.meetingAppRecords,
    meeting_app_records: input.meeting_app_records,
  });
}

function providerRecordsForPlatform(platform, input = {}, options = {}) {
  const sources = [
    input.providerRecordsByPlatform,
    input.provider_records_by_platform,
    input.providerEventsByPlatform,
    input.provider_events_by_platform,
    options.providerRecordsByPlatform,
    options.provider_records_by_platform,
    options.providerEventsByPlatform,
    options.provider_events_by_platform,
  ];
  for (const source of sources) {
    if (!isPlainObject(source)) continue;
    for (const key of platformAliases(platform)) {
      if (source[key] != null) return source[key];
    }
  }
  return firstNonEmpty(
    input.providerRecords,
    input.provider_records,
    input.providerEvents,
    input.provider_events,
    options.providerRecords,
    options.provider_records,
    options.providerEvents,
    options.provider_events,
  );
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function blockingRealtime(preflight = {}, selection = {}, blueprint = {}, providerReplay = {}) {
  return preflight.readiness?.provider_reconcile_required_for_realtime === true
    || selection.runtime_policy?.provider_events_block_realtime === true
    || blueprint.runtime_contract?.provider_events_block_realtime === true
    || providerReplay.runtime_contract?.provider_events_block_realtime === true
    || providerReplay.provider_events_block_realtime === true;
}

function transcriptBlocking(preflight = {}, selection = {}, blueprint = {}, providerReplay = {}) {
  return preflight.readiness?.transcript_blocks_realtime === true
    || selection.runtime_policy?.transcript_blocks_realtime === true
    || blueprint.runtime_contract?.transcript_blocks_realtime === true
    || providerReplay.runtime_contract?.transcript_blocks_realtime === true
    || providerReplay.transcript_blocks_realtime === true;
}

function readinessIssues(platform, preflight = {}, selection = {}, blueprint = {}, providerReplay = {}) {
  const issues = [
    selection.readiness?.selection_ready === true
      ? undefined
      : issue('error', 'adapter_selection_not_ready', 'Adapter selection must be ready before realtime axis handoff.', { platform }),
    blueprint.readiness?.ready === true
      ? undefined
      : issue('error', 'adapter_blueprint_not_ready', 'Adapter blueprint must be ready before realtime axis handoff.', { platform }),
    preflight.readiness?.static_startup_ready === true
      ? undefined
      : issue('error', 'static_startup_not_ready', 'Static startup plan is not ready for the selected platform.', { platform }),
    preflight.readiness?.live_evidence_ready === true
      ? undefined
      : issue('error', 'missing_live_meeting_evidence', 'Realtime axis readiness requires live meeting page or native-window evidence.', { platform }),
    preflight.readiness?.meeting_start_ready === true
      ? undefined
      : issue('error', 'missing_realtime_meeting_start', 'Live evidence did not prove a current meeting_started signal.', { platform }),
    preflight.readiness?.realtime_annotation_ready === true
      ? undefined
      : issue('error', 'realtime_axis_not_ready', 'Realtime annotation insertion cannot be opened on the current axis yet.', { platform }),
    providerReplay.accepted === true
      ? undefined
      : issue('error', 'provider_replay_not_accepted', 'Provider replay evidence must be accepted for production reconcile/backfill.', {
        platform,
        provider_replay_issues: providerReplay.issues ?? [],
      }),
    blockingRealtime(preflight, selection, blueprint, providerReplay)
      ? issue('error', 'provider_events_block_realtime', 'Provider events must not block realtime annotation insertion.', { platform })
      : undefined,
    transcriptBlocking(preflight, selection, blueprint, providerReplay)
      ? issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion.', { platform })
      : undefined,
    ...(preflight.issues ?? []).filter((item) => item.severity === 'error').map((item) => ({
      ...item,
      source: firstNonEmpty(item.source, 'adapter_preflight'),
    })),
  ].filter(Boolean);
  const seen = new Set();
  return issues.filter((item) => {
    const key = `${item.severity}:${item.code}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function statusFor(preflight = {}, issues = [], providerReplay = {}) {
  if (issues.length === 0) return 'ready_for_realtime_axis';
  if (issues.some((item) => item.code === 'static_startup_not_ready' || item.code === 'adapter_selection_not_ready' || item.code === 'adapter_blueprint_not_ready')) {
    return 'static_adapter_not_ready';
  }
  if (issues.some((item) => item.code === 'missing_live_meeting_evidence')) return 'waiting_for_live_meeting_evidence';
  if (issues.some((item) => item.code === 'missing_realtime_meeting_start')) return 'waiting_for_meeting_start';
  if (providerReplay.accepted !== true) return 'provider_replay_not_ready';
  if (preflight.readiness?.realtime_annotation_ready !== true) return 'realtime_axis_not_ready';
  return 'realtime_axis_partial';
}

function providerReplaySummary(providerReplay = {}) {
  return compactObject({
    schema: providerReplay.schema,
    accepted: providerReplay.accepted === true,
    target: providerReplay.target,
    source: providerReplay.source,
    record_count: providerReplay.record_count,
    accepted_record_count: providerReplay.accepted_record_count,
    runtime_event_count: providerReplay.runtime_event_count,
    signal_count: providerReplay.signal_count,
    signal_types: providerReplay.signal_types,
    coverage: providerReplay.coverage,
    required_coverage: providerReplay.required_coverage,
    provider_events_block_realtime: providerReplay.runtime_contract?.provider_events_block_realtime ?? providerReplay.provider_events_block_realtime,
    transcript_blocks_realtime: providerReplay.runtime_contract?.transcript_blocks_realtime ?? providerReplay.transcript_blocks_realtime,
    issue_count: providerReplay.issue_count,
    issues: providerReplay.issues,
  });
}

function selectionSummary(selection = {}) {
  return compactObject({
    schema: selection.schema,
    ready: selection.readiness?.selection_ready === true,
    recommended_mode: selection.recommended_mode,
    axis_source: selection.selection?.axis_source,
    axis_surface: selection.selection?.axis_surface,
    timestamp_field: selection.selection?.timestamp_field,
    provider_reconcile_source: selection.selection?.provider_reconcile_source,
    provider_reconcile_required_for_production: selection.selection?.provider_reconcile_required_for_production,
    speaker_track_source: selection.selection?.speaker_track_source,
    provider_events_block_realtime: selection.runtime_policy?.provider_events_block_realtime,
    transcript_blocks_realtime: selection.runtime_policy?.transcript_blocks_realtime,
  });
}

function blueprintSummary(blueprint = {}) {
  return compactObject({
    schema: blueprint.schema,
    ready: blueprint.readiness?.ready === true,
    primary_surface: blueprint.primary_surface,
    surface_order: blueprint.surface_order,
    provider_events_block_realtime: blueprint.runtime_contract?.provider_events_block_realtime,
    transcript_blocks_realtime: blueprint.runtime_contract?.transcript_blocks_realtime,
    timestamp_field: blueprint.realtime_axis_contract?.timestamp_field,
    create_on: blueprint.realtime_axis_contract?.create_on,
    end_on: blueprint.realtime_axis_contract?.end_on,
  });
}

function nextActions(preflight = {}, providerReplay = {}, issues = []) {
  return unique([
    ...(preflight.next_actions ?? []),
    ...(providerReplay.next_actions ?? []),
    ...issues.map((item) => `fix_${item.code}`),
    issues.length === 0 ? 'open_realtime_axis_session_and_insert_marks_by_captured_at_ms' : undefined,
    issues.length === 0 ? 'keep_provider_replay_for_reconcile_and_backfill_only' : undefined,
  ]);
}

export function buildMeetingPlatformRealtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
  const objectInput = typeof platformOrInput === 'string' || platformOrInput instanceof URL
    ? { platform: platformOrInput, ...(isPlainObject(inputOrOptions) ? inputOrOptions : {}) }
    : (platformOrInput ?? {});
  const options = typeof platformOrInput === 'string' || platformOrInput instanceof URL
    ? maybeOptions
    : inputOrOptions;
  const platform = normalizeMeetingPlatform(firstNonEmpty(
    objectInput.platform,
    objectInput.provider,
    options.platform,
    options.provider,
  ));
  const platformInput = {
    ...evidenceInput(objectInput),
    ...inputForPlatform(platform, objectInput, options),
    platform,
  };
  const preflight = buildMeetingPlatformAdapterPreflight(platformInput, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  });
  const selection = buildMeetingPlatformAdapterSelection(platform, objectInput, options);
  const blueprint = buildMeetingPlatformAdapterBlueprint(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  });
  const providerRecords = providerRecordsForPlatform(platform, objectInput, options);
  const providerReplay = buildMeetingPlatformProviderReplayReport(platform, asArray(providerRecords).length > 0 ? providerRecords : undefined, options);
  const issues = readinessIssues(platform, preflight, selection, blueprint, providerReplay);
  const realtimeAxisReady = preflight.readiness?.realtime_annotation_ready === true
    && selection.readiness?.selection_ready === true
    && blueprint.readiness?.ready === true
    && !blockingRealtime(preflight, selection, blueprint, providerReplay)
    && !transcriptBlocking(preflight, selection, blueprint, providerReplay);
  const productionHandoffReady = realtimeAxisReady && providerReplay.accepted === true;
  const accepted = productionHandoffReady && issues.length === 0;
  return compactObject({
    type: 'meeting_platform_realtime_axis_readiness',
    schema: MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA,
    schema_version: MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA_VERSION,
    accepted,
    status: statusFor(preflight, issues, providerReplay),
    platform,
    display_name: preflight.display_name ?? selection.display_name ?? blueprint.display_name,
    realtime_axis_ready: realtimeAxisReady,
    production_handoff_ready: productionHandoffReady,
    axis_source: selection.selection?.axis_source,
    axis_surface: selection.selection?.axis_surface,
    selected_surface: preflight.startup?.selected_surface,
    timestamp_field: selection.selection?.timestamp_field,
    local_observer: {
      preflight_schema: preflight.schema,
      accepted: preflight.accepted === true,
      status: preflight.status,
      static_startup_ready: preflight.readiness?.static_startup_ready === true,
      live_evidence_ready: preflight.readiness?.live_evidence_ready === true,
      meeting_start_ready: preflight.readiness?.meeting_start_ready === true,
      meeting_end_ready: preflight.readiness?.meeting_end_ready === true,
      speaker_track_ready: preflight.readiness?.speaker_track_ready === true,
      realtime_annotation_ready: preflight.readiness?.realtime_annotation_ready === true,
      evidence_kind: preflight.summary?.evidence_kind,
      evidence_count: preflight.summary?.evidence_count,
      record_count: preflight.summary?.record_count,
    },
    provider_replay: providerReplaySummary(providerReplay),
    adapter_selection: selectionSummary(selection),
    adapter_blueprint: blueprintSummary(blueprint),
    runtime_policy: {
      provider_events_block_realtime: blockingRealtime(preflight, selection, blueprint, providerReplay),
      transcript_blocks_realtime: transcriptBlocking(preflight, selection, blueprint, providerReplay),
      provider_replay_required_for_production_handoff: true,
      provider_replay_required_for_realtime_axis: false,
      local_observer_required_for_realtime_axis: true,
    },
    issue_count: issues.length,
    issues,
    next_actions: nextActions(preflight, providerReplay, issues),
  });
}

export function buildMeetingPlatformRealtimeAxisReadinessMatrix(input = {}, options = {}) {
  const objectInput = input ?? {};
  const platforms = selectedPlatforms(objectInput, options);
  const reports = platforms.map((platform) => buildMeetingPlatformRealtimeAxisReadiness({
    ...objectInput,
    platform,
  }, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_realtime_axis_readiness_matrix',
    schema: MEETING_PLATFORM_REALTIME_AXIS_READINESS_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted === true).length,
    realtime_axis_ready_count: reports.filter((report) => report.realtime_axis_ready === true).length,
    production_handoff_ready_count: reports.filter((report) => report.production_handoff_ready === true).length,
    local_observer_ready_count: reports.filter((report) => report.local_observer?.realtime_annotation_ready === true).length,
    live_evidence_ready_count: reports.filter((report) => report.local_observer?.live_evidence_ready === true).length,
    meeting_start_ready_count: reports.filter((report) => report.local_observer?.meeting_start_ready === true).length,
    speaker_track_ready_count: reports.filter((report) => report.local_observer?.speaker_track_ready === true).length,
    provider_replay_accepted_count: reports.filter((report) => report.provider_replay?.accepted === true).length,
    provider_blocking_count: reports.filter((report) => report.runtime_policy?.provider_events_block_realtime === true).length,
    transcript_blocking_count: reports.filter((report) => report.runtime_policy?.transcript_blocks_realtime === true).length,
    waiting_for_live_evidence_count: reports.filter((report) => report.status === 'waiting_for_live_meeting_evidence').length,
    platforms: reports.map((report) => report.platform),
    rows: reports.map((report) => ({
      platform: report.platform,
      display_name: report.display_name,
      accepted: report.accepted,
      status: report.status,
      realtime_axis_ready: report.realtime_axis_ready,
      production_handoff_ready: report.production_handoff_ready,
      axis_source: report.axis_source,
      axis_surface: report.axis_surface,
      selected_surface: report.selected_surface,
      local_observer_ready: report.local_observer?.realtime_annotation_ready === true,
      live_evidence_ready: report.local_observer?.live_evidence_ready === true,
      meeting_start_ready: report.local_observer?.meeting_start_ready === true,
      meeting_end_ready: report.local_observer?.meeting_end_ready === true,
      speaker_track_ready: report.local_observer?.speaker_track_ready === true,
      provider_replay_accepted: report.provider_replay?.accepted === true,
      provider_events_block_realtime: report.runtime_policy?.provider_events_block_realtime,
      transcript_blocks_realtime: report.runtime_policy?.transcript_blocks_realtime,
      issue_codes: (report.issues ?? []).map((item) => item.code),
      first_next_action: report.next_actions?.[0],
    })),
    reports,
    next_actions: unique(reports.flatMap((report) => report.next_actions ?? [])),
  };
}

export function assertMeetingPlatformRealtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
  const report = platformOrInput?.schema === MEETING_PLATFORM_REALTIME_AXIS_READINESS_SCHEMA
    ? platformOrInput
    : buildMeetingPlatformRealtimeAxisReadiness(platformOrInput, inputOrOptions, maybeOptions);
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform realtime axis readiness is not accepted', {
      platform: report.platform,
      status: report.status,
      issues: report.issues,
      next_actions: report.next_actions,
      report,
    });
  }
  return report;
}

export function assertMeetingPlatformRealtimeAxisReadinessMatrix(input = {}, options = {}) {
  const matrix = input?.schema === MEETING_PLATFORM_REALTIME_AXIS_READINESS_MATRIX_SCHEMA
    ? input
    : buildMeetingPlatformRealtimeAxisReadinessMatrix(input, options);
  if (matrix.accepted_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform realtime axis readiness matrix is not accepted', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      failed_platforms: matrix.rows.filter((row) => row.accepted !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
