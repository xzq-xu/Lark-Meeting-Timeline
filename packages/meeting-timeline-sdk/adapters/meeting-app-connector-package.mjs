import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA = 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA = 'meeting_app_timeline_connector_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION = 1;

const REALTIME_RUNTIME_ACTIONS = Object.freeze([
  'observe_meeting_app',
  'observe_platform_candidates',
  'insert_annotation',
]);

const TRACK_RUNTIME_ACTIONS = Object.freeze([
  'speaker_track',
  'participant_track',
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

function normalizeKey(value) {
  return String(value ?? '').trim().replaceAll('-', '_');
}

function selectedPlatforms(pkg = {}, options = {}) {
  const explicit = firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys);
  return unique(asArray(explicit ?? pkg.platforms).map((item) => normalizeKey(item)).filter(Boolean));
}

function selectedSurfaces(pkg = {}, options = {}) {
  const explicit = firstNonEmpty(options.surfaces, options.surface);
  return unique(asArray(explicit ?? pkg.surfaces).map((item) => normalizeKey(item)).filter(Boolean));
}

function acceptanceTarget(options = {}) {
  return String(firstNonEmpty(
    options.target,
    options.acceptance_target,
    options.acceptanceTarget,
    options.requireProductionReady === true || options.require_production_ready === true ? 'production' : undefined,
    options.requireRealtimeReady === true || options.require_realtime_ready === true ? 'realtime' : undefined,
    'realtime',
  ));
}

function shouldRequireTracks(options = {}) {
  return options.requireTracks !== false && options.require_tracks !== false;
}

function addIssue(issues, code, message, details = {}) {
  issues.push(compactObject({
    code,
    message,
    ...details,
  }));
}

function runtimeActionsByPlatform(pkg = {}) {
  const result = new Map();
  for (const row of pkg.runtime_events?.plan_matrix?.rows ?? []) {
    const platform = normalizeKey(row.platform);
    if (!platform) continue;
    const actions = result.get(platform) ?? new Set();
    if (row.action) actions.add(String(row.action));
    result.set(platform, actions);
  }
  return result;
}

function missingRuntimeActions(pkg = {}, platforms = [], options = {}) {
  const actionMap = runtimeActionsByPlatform(pkg);
  const requiredActions = shouldRequireTracks(options)
    ? [...REALTIME_RUNTIME_ACTIONS, ...TRACK_RUNTIME_ACTIONS]
    : [...REALTIME_RUNTIME_ACTIONS];
  return platforms.flatMap((platform) => {
    const actions = actionMap.get(platform) ?? new Set();
    return requiredActions
      .filter((action) => !actions.has(action))
      .map((action) => ({ platform, action }));
  });
}

function rowReadyCount(rows = [], key = 'ready_to_start') {
  return rows.filter((row) => row?.[key] === true).length;
}

function finiteMin(values = []) {
  const finite = values.filter((value) => Number.isFinite(value));
  return finite.length > 0 ? Math.min(...finite) : undefined;
}

function surfaceHandoffRows(pkg = {}, surface) {
  return (pkg.handoff_matrix?.rows ?? []).filter((row) => normalizeKey(row.surface) === surface);
}

function surfaceAcceptance(pkg = {}, surface, platforms = [], options = {}) {
  const observer = pkg.observer_plan_by_surface?.[surface];
  const scheduler = pkg.scheduler_config_by_surface?.[surface];
  const rows = surfaceHandoffRows(pkg, surface);
  const issues = [];

  if (!observer) {
    addIssue(issues, 'missing_observer_plan', `Missing observer plan for surface ${surface}`, { surface });
  } else if ((observer.accepted_count ?? 0) < platforms.length) {
    addIssue(issues, 'observer_plan_not_accepted', `Observer plan is not accepted for every platform on ${surface}`, {
      surface,
      accepted_count: observer.accepted_count ?? 0,
      platform_count: platforms.length,
    });
  }

  if (!scheduler) {
    addIssue(issues, 'missing_scheduler_config', `Missing scheduler config for surface ${surface}`, { surface });
  } else {
    if ((scheduler.sdk_ready_count ?? 0) < platforms.length) {
      addIssue(issues, 'scheduler_not_sdk_ready', `Scheduler is not SDK-ready for every platform on ${surface}`, {
        surface,
        sdk_ready_count: scheduler.sdk_ready_count ?? 0,
        platform_count: platforms.length,
      });
    }
    if (shouldRequireTracks(options) && (scheduler.track_enabled_count ?? 0) < platforms.length) {
      addIssue(issues, 'tracks_not_enabled', `Speaker/participant tracks are not enabled for every platform on ${surface}`, {
        surface,
        track_enabled_count: scheduler.track_enabled_count ?? 0,
        platform_count: platforms.length,
      });
    }
  }

  if (rows.length < platforms.length) {
    addIssue(issues, 'missing_handoff_rows', `Missing handoff rows for surface ${surface}`, {
      surface,
      row_count: rows.length,
      platform_count: platforms.length,
    });
  }

  if (rowReadyCount(rows, 'ready_to_start') < rows.length) {
    addIssue(issues, 'handoff_not_ready', `Not every handoff row is ready to start on ${surface}`, {
      surface,
      ready_count: rowReadyCount(rows, 'ready_to_start'),
      row_count: rows.length,
    });
  }

  return {
    surface,
    accepted: issues.length === 0,
    handoff_count: rows.length,
    ready_count: rowReadyCount(rows, 'ready_to_start'),
    realtime_annotation_ready_count: rowReadyCount(rows, 'realtime_annotation_ready'),
    speaker_track_ready_count: rowReadyCount(rows, 'speaker_track_ready'),
    participant_track_ready_count: rowReadyCount(rows, 'participant_track_ready'),
    observer_accepted_count: observer?.accepted_count ?? 0,
    scheduler_sdk_ready_count: scheduler?.sdk_ready_count ?? 0,
    track_enabled_count: scheduler?.track_enabled_count ?? 0,
    install_targets: unique(rows.map((row) => row.install_target)),
    start_modes: unique(rows.map((row) => row.start_mode)),
    runtime_factories: unique(rows.map((row) => row.runtime_factory)),
    observer_factories: unique(observer?.rows?.map((row) => row.observer_factory) ?? []),
    issues,
  };
}

function extensionAcceptance(pkg = {}, surfaces = [], options = {}) {
  if (!surfaces.includes('browser_extension') || options.requireExtensionScaffold === false || options.require_extension_scaffold === false) {
    return { required: false, accepted: true, issues: [] };
  }
  const issues = [];
  const extension = pkg.extension ?? {};
  if (!extension.scaffold) addIssue(issues, 'missing_extension_scaffold', 'Browser extension surface requires an extension scaffold');
  if (extension.acceptance?.accepted !== true) addIssue(issues, 'extension_acceptance_failed', 'Browser extension scaffold acceptance failed');
  if (extension.manifest?.manifest_version !== 3) addIssue(issues, 'extension_manifest_v3_missing', 'Browser extension manifest_version must be 3');
  if ((extension.manifest?.host_permissions ?? []).includes('<all_urls>')) {
    addIssue(issues, 'extension_all_urls_forbidden', 'Browser extension host permissions must not include <all_urls>');
  }
  return {
    required: true,
    accepted: issues.length === 0,
    file_count: extension.file_count ?? extension.scaffold?.files?.length ?? 0,
    manifest_version: extension.manifest?.manifest_version,
    issues,
  };
}

export function buildMeetingAppTimelineConnectorPackageAcceptanceReport(pkg = {}, options = {}) {
  const target = acceptanceTarget(options);
  const platforms = selectedPlatforms(pkg, options);
  const surfaces = selectedSurfaces(pkg, options);
  const issues = [];

  if (pkg.schema !== MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA) {
    addIssue(issues, 'invalid_schema', 'Expected a meeting app timeline connector package', {
      expected_schema: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA,
      actual_schema: pkg.schema,
    });
  }
  if (pkg.accepted !== true) addIssue(issues, 'connector_package_not_accepted', 'Connector package accepted flag is not true');
  if (pkg.host_package?.accepted !== true) addIssue(issues, 'host_package_not_accepted', 'Host package accepted flag is not true');
  if (pkg.handoff_acceptance?.accepted !== true) addIssue(issues, 'handoff_acceptance_failed', 'Handoff matrix acceptance failed');
  if (platforms.length === 0) addIssue(issues, 'missing_platforms', 'Connector package must contain at least one platform');
  if (surfaces.length === 0) addIssue(issues, 'missing_surfaces', 'Connector package must contain at least one runtime surface');

  const timestampField = firstNonEmpty(
    pkg.contracts?.timestamp_field,
    pkg.host_package?.runtime_contract?.annotation_timestamp_field,
  );
  if (timestampField !== 'captured_at_ms') {
    addIssue(issues, 'invalid_timestamp_contract', 'Realtime annotation insertion must use captured_at_ms', {
      timestamp_field: timestampField,
    });
  }
  if (pkg.contracts?.provider_events_block_realtime !== false) {
    addIssue(issues, 'provider_events_block_realtime', 'Provider events must not block realtime annotation insertion');
  }
  if (pkg.contracts?.transcript_blocks_realtime !== false) {
    addIssue(issues, 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion');
  }
  if (!pkg.runtime_events?.endpoint) addIssue(issues, 'missing_runtime_event_endpoint', 'Runtime event endpoint is required');
  if ((pkg.runtime_events?.action_count ?? 0) <= 0) addIssue(issues, 'missing_runtime_event_actions', 'Runtime event plan has no actions');

  const missingActions = missingRuntimeActions(pkg, platforms, options);
  for (const missing of missingActions) {
    addIssue(issues, 'missing_runtime_action', `Missing runtime action ${missing.action} for ${missing.platform}`, missing);
  }

  const surface_reports = surfaces.map((surface) => surfaceAcceptance(pkg, surface, platforms, options));
  for (const report of surface_reports) issues.push(...report.issues);

  const extension = extensionAcceptance(pkg, surfaces, options);
  issues.push(...extension.issues);

  if (target === 'production' && options.liveEvidenceAccepted !== true && options.live_evidence_accepted !== true) {
    addIssue(issues, 'production_requires_live_snapshot_evidence', 'Production acceptance requires live meeting snapshot evidence');
  }

  return {
    type: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target,
    package_id: pkg.id,
    platform_count: platforms.length,
    surface_count: surfaces.length,
    platforms,
    surfaces,
    require_tracks: shouldRequireTracks(options),
    missing_runtime_actions: missingActions,
    surface_reports,
    extension,
    issue_count: issues.length,
    issues,
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue.code}`),
        ...(pkg.next_actions ?? []),
      ])
      : pkg.next_actions ?? [],
  };
}

export function assertMeetingAppTimelineConnectorPackage(pkg = {}, options = {}) {
  const report = buildMeetingAppTimelineConnectorPackageAcceptanceReport(pkg, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector package is not accepted', {
      report,
      issues: report.issues,
    });
  }
  return pkg;
}

export function stripMeetingAppTimelineConnectorPackageFileContents(pkg = {}) {
  return {
    ...pkg,
    extension: pkg.extension
      ? {
        ...pkg.extension,
        scaffold: pkg.extension.scaffold
          ? {
            ...pkg.extension.scaffold,
            files: (pkg.extension.scaffold.files ?? []).map(({ content, ...file }) => file),
          }
          : undefined,
      }
      : undefined,
  };
}

export function buildMeetingAppTimelineConnectorHandoff(pkg = {}, options = {}) {
  const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(pkg, options);
  const platforms = acceptance.platforms;
  const surfaces = acceptance.surfaces;
  const handoffRows = pkg.handoff_matrix?.rows ?? [];
  const surface_matrix = surfaces.map((surface) => {
    const rows = handoffRows.filter((row) => normalizeKey(row.surface) === surface);
    const scheduler = pkg.scheduler_config_by_surface?.[surface];
    return compactObject({
      surface,
      platform_count: rows.length,
      ready_count: rowReadyCount(rows, 'ready_to_start'),
      realtime_annotation_ready_count: rowReadyCount(rows, 'realtime_annotation_ready'),
      speaker_track_ready_count: rowReadyCount(rows, 'speaker_track_ready'),
      participant_track_ready_count: rowReadyCount(rows, 'participant_track_ready'),
      install_targets: unique(rows.map((row) => row.install_target)),
      start_modes: unique(rows.map((row) => row.start_mode)),
      runtime_factories: unique(rows.map((row) => row.runtime_factory)),
      track_runtime_factories: unique(rows.map((row) => row.track_runtime_factory)),
      scheduler_track_enabled_count: scheduler?.track_enabled_count,
      scheduler_trigger_count: Math.max(0, ...(scheduler?.rows ?? []).map((row) => row.trigger_count ?? 0)),
      fallback_poll_interval_ms: finiteMin((scheduler?.rows ?? []).map((row) => row.fallback_poll_interval_ms)),
    });
  });
  const extensionFiles = (pkg.extension?.scaffold?.files ?? []).map((file) => file.path);

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    package_id: pkg.id,
    accepted: acceptance.accepted,
    target: acceptance.target,
    base_url: pkg.base_url,
    platforms,
    surfaces,
    platform_count: platforms.length,
    surface_count: surfaces.length,
    runtime_event_endpoint: pkg.runtime_events?.endpoint,
    timestamp_field: firstNonEmpty(pkg.contracts?.timestamp_field, pkg.host_package?.runtime_contract?.annotation_timestamp_field),
    provider_events_block_realtime: pkg.contracts?.provider_events_block_realtime,
    transcript_blocks_realtime: pkg.contracts?.transcript_blocks_realtime,
    host_entrypoints: pkg.host_package?.host_entrypoints,
    package_entrypoints: pkg.entrypoints,
    ci_gates: pkg.host_package?.ci_gates,
    rollout_checklist: pkg.host_package?.rollout_checklist,
    surface_matrix,
    extension: pkg.extension ? {
      required: surfaces.includes('browser_extension'),
      file_count: extensionFiles.length,
      file_paths: extensionFiles,
      manifest_version: pkg.extension.manifest?.manifest_version,
      permissions: pkg.extension.manifest?.permissions,
      host_permissions: pkg.extension.manifest?.host_permissions,
    } : undefined,
    acceptance,
    next_actions: acceptance.next_actions,
  });
}
