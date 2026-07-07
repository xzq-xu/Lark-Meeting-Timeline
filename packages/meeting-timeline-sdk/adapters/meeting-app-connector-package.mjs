import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformRuntimeEvent,
  createMeetingPlatformRuntimeEventClient,
  normalizeMeetingPlatformRuntimeEventAction,
} from './platform-runtime-event.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA = 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA = 'meeting_app_timeline_connector_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA = 'meeting_app_timeline_connector_host_install_checklist';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA = 'meeting_app_timeline_connector_smoke_plan';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_smoke_plan_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_RUNTIME_CLIENT_SCHEMA = 'meeting_app_timeline_connector_runtime_client';
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

function runtimeActionRows(pkg = {}) {
  return (pkg.runtime_events?.plan_matrix?.rows ?? []).map((row) => compactObject({
    platform: normalizeKey(row.platform),
    action: normalizeMeetingPlatformRuntimeEventAction(row.action),
    client_method: row.client_method,
    producer: row.producer,
    realtime_role: row.realtime_role,
    provider_dependency: row.provider_dependency,
    transcript_dependency: row.transcript_dependency,
  }));
}

function runtimeActionSet(pkg = {}) {
  return new Set(runtimeActionRows(pkg).map((row) => row.action));
}

function runtimePlatformActionMap(pkg = {}) {
  const result = new Map();
  for (const row of runtimeActionRows(pkg)) {
    if (!row.platform) continue;
    const actions = result.get(row.platform) ?? new Set();
    actions.add(row.action);
    result.set(row.platform, actions);
  }
  return result;
}

function runtimeActionRowsForPlatform(pkg = {}, platform) {
  const normalized = normalizeKey(platform);
  return runtimeActionRows(pkg).filter((row) => row.platform === normalized);
}

function runtimeActionByName(rows = [], action) {
  return rows.find((row) => row.action === action);
}

function sampleUrlForPlatform(platform) {
  return {
    google_meet: 'https://meet.google.com/abc-defg-hij',
    teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    microsoft_teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    zoom: 'https://zoom.us/j/987654321',
    webex: 'https://example.webex.com/meet/sample',
    lark: 'https://vc.feishu.cn/j/123456789',
  }[normalizeKey(platform)] ?? 'local://meeting-window/sample';
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
  if (!pkg.adapter_blueprints?.matrix) {
    addIssue(issues, 'missing_adapter_blueprint_matrix', 'Connector package must include adapter blueprint matrix for external host wiring');
  } else if ((pkg.adapter_blueprints.ready_count ?? 0) < platforms.length) {
    addIssue(issues, 'adapter_blueprints_not_ready', 'Every selected platform must expose a ready adapter blueprint', {
      ready_count: pkg.adapter_blueprints.ready_count,
      platform_count: platforms.length,
    });
  }
  if (!pkg.startup_plans?.matrix) {
    addIssue(issues, 'missing_startup_plan_matrix', 'Connector package must include startup plan matrix for host runtime install');
  } else if ((pkg.startup_plans.realtime_startup_ready_count ?? 0) < platforms.length) {
    addIssue(issues, 'startup_plans_not_ready', 'Every selected platform must expose a realtime-ready startup plan', {
      realtime_startup_ready_count: pkg.startup_plans.realtime_startup_ready_count,
      platform_count: platforms.length,
    });
  }

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
    adapter_blueprints: pkg.adapter_blueprints ? {
      ready_count: pkg.adapter_blueprints.ready_count,
      platform_count: pkg.adapter_blueprints.platform_count,
      sdk_method: pkg.adapter_blueprints.sdk_method,
      command: pkg.adapter_blueprints.command,
      rows: pkg.adapter_blueprints.matrix?.rows,
    } : undefined,
    startup_plans: pkg.startup_plans ? {
      realtime_startup_ready_count: pkg.startup_plans.realtime_startup_ready_count,
      platform_count: pkg.startup_plans.platform_count,
      sdk_method: pkg.startup_plans.sdk_method,
      matrix_sdk_method: pkg.startup_plans.matrix_sdk_method,
      rows: pkg.startup_plans.matrix?.rows,
    } : undefined,
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

export function buildMeetingAppTimelineConnectorHostInstallChecklist(pkg = {}, options = {}) {
  const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(pkg, options);
  const startupRows = pkg.startup_plans?.matrix?.rows ?? [];
  const blueprintRows = pkg.adapter_blueprints?.matrix?.rows ?? [];
  const startupByPlatform = new Map(startupRows.map((row) => [normalizeKey(row.platform), row]));
  const blueprintByPlatform = new Map(blueprintRows.map((row) => [normalizeKey(row.platform), row]));
  const rows = acceptance.platforms.map((platform) => {
    const startup = startupByPlatform.get(platform) ?? {};
    const blueprint = blueprintByPlatform.get(platform) ?? {};
    const actions = runtimeActionRowsForPlatform(pkg, platform);
    const observeCandidates = runtimeActionByName(actions, 'observe_platform_candidates');
    const observeMeetingApp = runtimeActionByName(actions, 'observe_meeting_app');
    const insertAnnotation = runtimeActionByName(actions, 'insert_annotation');
    const speakerTrack = runtimeActionByName(actions, 'speaker_track');
    const participantTrack = runtimeActionByName(actions, 'participant_track');
    return compactObject({
      platform,
      display_name: startup.display_name ?? blueprint.display_name,
      selected_surface: startup.selected_surface,
      install_target: startup.install_target,
      runtime_preset: startup.runtime_preset,
      realtime_startup_ready: startup.realtime_startup_ready === true,
      adapter_blueprint_ready: startup.adapter_blueprint_ready === true || blueprint.ready === true,
      adapter_blueprint_primary_surface: startup.adapter_blueprint_primary_surface ?? blueprint.primary_surface,
      adapter_blueprint_first_acceptance_gate: startup.adapter_blueprint_first_acceptance_gate ?? blueprint.first_acceptance_gate,
      first_action: startup.first_action,
      observe_action: startup.observe_action,
      insert_action: startup.insert_action,
      provider_events_block_realtime: startup.provider_events_block_realtime,
      transcript_blocks_realtime: startup.transcript_blocks_realtime,
      runtime_actions: actions.map((row) => row.action),
      client_methods: {
        observe_platform_candidates: observeCandidates?.client_method,
        observe_meeting_app: observeMeetingApp?.client_method,
        insert_annotation: insertAnnotation?.client_method,
        speaker_track: speakerTrack?.client_method,
        participant_track: participantTrack?.client_method,
      },
      required_host_steps: [
        'install_selected_surface_runtime',
        'observe_platform_candidates_before_first_annotation',
        'insert_annotation_with_captured_at_ms',
        'keep_provider_events_nonblocking',
        'keep_transcript_import_nonblocking',
      ],
      optional_host_steps: [
        speakerTrack ? 'emit_speaker_track_markers' : undefined,
        participantTrack ? 'emit_participant_track_markers' : undefined,
      ].filter(Boolean),
    });
  });
  const readyCount = rows.filter((row) => row.realtime_startup_ready === true && row.adapter_blueprint_ready === true).length;
  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    package_id: pkg.id,
    accepted: acceptance.accepted === true && readyCount === acceptance.platform_count,
    target: acceptance.target,
    base_url: pkg.base_url,
    runtime_event_endpoint: pkg.runtime_events?.endpoint,
    platform_count: acceptance.platform_count,
    surface_count: acceptance.surface_count,
    ready_count: readyCount,
    timestamp_field: firstNonEmpty(pkg.contracts?.timestamp_field, pkg.host_package?.runtime_contract?.annotation_timestamp_field),
    contracts: {
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: pkg.contracts?.provider_events_block_realtime,
      transcript_blocks_realtime: pkg.contracts?.transcript_blocks_realtime,
      startup_plan_required_before_runtime_install: pkg.contracts?.startup_plan_required_before_runtime_install,
      adapter_blueprint_required_before_host_wiring: pkg.contracts?.adapter_blueprint_required_before_host_wiring,
    },
    files_to_read_first: [
      'connector-handoff.json',
      'startup-plan-matrix.json',
      'adapter-blueprint-matrix.json',
      'runtime-event-plan-matrix.json',
      'connector-package.json',
    ],
    rows,
    acceptance,
    next_actions: acceptance.next_actions,
  });
}

export function buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(checklistOrPackage = {}, options = {}) {
  const checklist = checklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA
    ? checklistOrPackage
    : buildMeetingAppTimelineConnectorHostInstallChecklist(checklistOrPackage, options);
  const rows = checklist.rows ?? [];
  const issues = [];

  if (checklist.schema !== MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA) {
    addIssue(issues, 'invalid_schema', 'Expected a meeting app timeline connector host install checklist', {
      expected_schema: MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA,
      actual_schema: checklist.schema,
    });
  }
  if (checklist.accepted !== true) addIssue(issues, 'checklist_not_accepted', 'Host install checklist accepted flag is not true');
  if (!checklist.runtime_event_endpoint) addIssue(issues, 'missing_runtime_event_endpoint', 'Runtime event endpoint is required');
  if (checklist.timestamp_field !== 'captured_at_ms') {
    addIssue(issues, 'invalid_timestamp_contract', 'Host install checklist must use captured_at_ms', {
      timestamp_field: checklist.timestamp_field,
    });
  }
  if (checklist.contracts?.provider_events_block_realtime !== false) {
    addIssue(issues, 'provider_events_block_realtime', 'Provider events must not block realtime annotation insertion');
  }
  if (checklist.contracts?.transcript_blocks_realtime !== false) {
    addIssue(issues, 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion');
  }
  if (checklist.contracts?.startup_plan_required_before_runtime_install !== true) {
    addIssue(issues, 'startup_plan_not_required', 'Startup plan must be required before host runtime install');
  }
  if (checklist.contracts?.adapter_blueprint_required_before_host_wiring !== true) {
    addIssue(issues, 'adapter_blueprint_not_required', 'Adapter blueprint must be required before host wiring');
  }
  if ((checklist.platform_count ?? 0) <= 0) addIssue(issues, 'missing_platforms', 'Host install checklist must include at least one platform');
  if (rows.length !== checklist.platform_count) {
    addIssue(issues, 'row_count_mismatch', 'Host install checklist row count must equal platform_count', {
      row_count: rows.length,
      platform_count: checklist.platform_count,
    });
  }

  for (const row of rows) {
    const platform = normalizeKey(row.platform);
    const runtimeActions = asArray(row.runtime_actions).map((action) => normalizeMeetingPlatformRuntimeEventAction(action));
    const requiredSteps = asArray(row.required_host_steps);
    if (!platform) addIssue(issues, 'row_missing_platform', 'Checklist row is missing platform');
    if (!row.selected_surface) addIssue(issues, 'row_missing_selected_surface', 'Checklist row is missing selected_surface', { platform });
    if (!row.install_target) addIssue(issues, 'row_missing_install_target', 'Checklist row is missing install_target', { platform });
    if (row.realtime_startup_ready !== true) addIssue(issues, 'row_not_realtime_startup_ready', 'Checklist row is not realtime startup ready', { platform });
    if (row.adapter_blueprint_ready !== true) addIssue(issues, 'row_adapter_blueprint_not_ready', 'Checklist row adapter blueprint is not ready', { platform });
    if (!runtimeActions.includes('observe_platform_candidates')) {
      addIssue(issues, 'row_missing_observe_platform_candidates_action', 'Checklist row must support observe_platform_candidates', { platform });
    }
    if (!runtimeActions.includes('insert_annotation')) {
      addIssue(issues, 'row_missing_insert_annotation_action', 'Checklist row must support insert_annotation', { platform });
    }
    if (row.client_methods?.insert_annotation !== 'insertAnnotation') {
      addIssue(issues, 'row_missing_insert_annotation_client_method', 'Checklist row must expose insertAnnotation client method', {
        platform,
        client_method: row.client_methods?.insert_annotation,
      });
    }
    if (!requiredSteps.includes('insert_annotation_with_captured_at_ms')) {
      addIssue(issues, 'row_missing_captured_at_ms_step', 'Checklist row must require insert_annotation_with_captured_at_ms', { platform });
    }
    if (row.provider_events_block_realtime !== false) {
      addIssue(issues, 'row_provider_events_block_realtime', 'Provider events must not block realtime annotations for row', { platform });
    }
    if (row.transcript_blocks_realtime !== false) {
      addIssue(issues, 'row_transcript_blocks_realtime', 'Transcript import must not block realtime annotations for row', { platform });
    }
  }

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: checklist.target,
    package_id: checklist.package_id,
    platform_count: checklist.platform_count ?? 0,
    row_count: rows.length,
    ready_count: rows.filter((row) => row.realtime_startup_ready === true && row.adapter_blueprint_ready === true).length,
    timestamp_field: checklist.timestamp_field,
    runtime_event_endpoint: checklist.runtime_event_endpoint,
    checklist_accepted: checklist.accepted === true,
    issue_count: issues.length,
    issues,
    rows: rows.map((row) => compactObject({
      platform: row.platform,
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      realtime_startup_ready: row.realtime_startup_ready,
      adapter_blueprint_ready: row.adapter_blueprint_ready,
      insert_annotation_client_method: row.client_methods?.insert_annotation,
    })),
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue.code}`),
        ...(checklist.next_actions ?? []),
      ])
      : checklist.next_actions ?? [],
  });
}

export function assertMeetingAppTimelineConnectorHostInstallChecklist(checklistOrPackage = {}, options = {}) {
  const report = buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(checklistOrPackage, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector host install checklist is not accepted', {
      report,
      issues: report.issues,
    });
  }
  return checklistOrPackage;
}

export function buildMeetingAppTimelineConnectorSmokePlan(checklistOrPackage = {}, options = {}) {
  const checklist = checklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA
    ? checklistOrPackage
    : buildMeetingAppTimelineConnectorHostInstallChecklist(checklistOrPackage, options);
  const baseMs = Number.isFinite(options.baseCapturedAtMs)
    ? options.baseCapturedAtMs
    : Number.isFinite(options.base_captured_at_ms)
      ? options.base_captured_at_ms
      : 1_782_614_400_000;
  const rows = (checklist.rows ?? []).map((row, index) => {
    const platform = normalizeKey(row.platform);
    const capturedAtMs = baseMs + (index * 10_000);
    const runtimeActions = asArray(row.runtime_actions).map((action) => normalizeMeetingPlatformRuntimeEventAction(action));
    const steps = [
      {
        id: `${platform}:observe_candidates`,
        order: 1,
        action: 'observe_platform_candidates',
        client_method: row.client_methods?.observe_platform_candidates ?? 'observePlatformCandidates',
        required: true,
        input: {
          captured_at_ms: capturedAtMs,
          tabs: [{
            active: true,
            url: sampleUrlForPlatform(platform),
            title: row.display_name ?? platform,
          }],
        },
        expected: {
          creates_or_updates_realtime_axis: true,
          timestamp_field: 'captured_at_ms',
        },
      },
      {
        id: `${platform}:insert_annotation`,
        order: 2,
        action: 'insert_annotation',
        client_method: row.client_methods?.insert_annotation ?? 'insertAnnotation',
        required: true,
        input: {
          platform,
          annotation: {
            id: `${platform}-smoke-mark`,
            label: 'smoke mark',
            captured_at_ms: capturedAtMs + 1_000,
          },
        },
        expected: {
          inserted_on_current_axis: true,
          timestamp_field: 'captured_at_ms',
        },
      },
      runtimeActions.includes('speaker_track') ? {
        id: `${platform}:speaker_track`,
        order: 3,
        action: 'speaker_track',
        client_method: row.client_methods?.speaker_track ?? 'speakerTrack',
        required: false,
        input: {
          platform,
          speaker: {
            id: `${platform}-speaker-1`,
            name: 'Speaker 1',
          },
          captured_at_ms: capturedAtMs + 2_000,
        },
        expected: {
          emits_position_marker_without_transcript_text: true,
        },
      } : undefined,
      runtimeActions.includes('participant_track') ? {
        id: `${platform}:participant_track`,
        order: 4,
        action: 'participant_track',
        client_method: row.client_methods?.participant_track ?? 'participantTrack',
        required: false,
        input: {
          platform,
          participant: {
            id: `${platform}-participant-1`,
            name: 'Participant 1',
          },
          captured_at_ms: capturedAtMs + 3_000,
        },
        expected: {
          emits_position_marker_without_transcript_text: true,
        },
      } : undefined,
    ].filter(Boolean);
    return compactObject({
      platform,
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      realtime_startup_ready: row.realtime_startup_ready,
      adapter_blueprint_ready: row.adapter_blueprint_ready,
      required_action_count: steps.filter((step) => step.required).length,
      optional_action_count: steps.filter((step) => !step.required).length,
      steps,
    });
  });
  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    package_id: checklist.package_id,
    accepted: checklist.accepted === true && rows.every((row) => row.steps?.some((step) => step.action === 'observe_platform_candidates') && row.steps?.some((step) => step.action === 'insert_annotation')),
    target: checklist.target,
    runtime_event_endpoint: checklist.runtime_event_endpoint,
    platform_count: checklist.platform_count,
    row_count: rows.length,
    timestamp_field: 'captured_at_ms',
    source_checklist_schema: checklist.schema,
    rows,
    next_actions: checklist.next_actions,
  });
}

export function buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(planOrChecklistOrPackage = {}, options = {}) {
  const plan = planOrChecklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA
    ? planOrChecklistOrPackage
    : buildMeetingAppTimelineConnectorSmokePlan(planOrChecklistOrPackage, options);
  const rows = plan.rows ?? [];
  const issues = [];

  if (plan.schema !== MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA) {
    addIssue(issues, 'invalid_schema', 'Expected a meeting app timeline connector smoke plan', {
      expected_schema: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA,
      actual_schema: plan.schema,
    });
  }
  if (plan.accepted !== true) addIssue(issues, 'smoke_plan_not_accepted', 'Connector smoke plan accepted flag is not true');
  if (!plan.runtime_event_endpoint) addIssue(issues, 'missing_runtime_event_endpoint', 'Smoke plan requires a runtime event endpoint');
  if (plan.timestamp_field !== 'captured_at_ms') {
    addIssue(issues, 'invalid_timestamp_contract', 'Smoke plan must use captured_at_ms', {
      timestamp_field: plan.timestamp_field,
    });
  }
  if ((plan.platform_count ?? 0) <= 0) addIssue(issues, 'missing_platforms', 'Smoke plan must include at least one platform');
  if (rows.length !== plan.platform_count) {
    addIssue(issues, 'row_count_mismatch', 'Smoke plan row count must equal platform_count', {
      row_count: rows.length,
      platform_count: plan.platform_count,
    });
  }

  for (const row of rows) {
    const platform = normalizeKey(row.platform);
    const steps = row.steps ?? [];
    const actions = steps.map((step) => normalizeMeetingPlatformRuntimeEventAction(step.action));
    const insertStep = steps.find((step) => normalizeMeetingPlatformRuntimeEventAction(step.action) === 'insert_annotation');
    const observeStep = steps.find((step) => normalizeMeetingPlatformRuntimeEventAction(step.action) === 'observe_platform_candidates');
    if (!platform) addIssue(issues, 'row_missing_platform', 'Smoke plan row is missing platform');
    if (!row.selected_surface) addIssue(issues, 'row_missing_selected_surface', 'Smoke plan row is missing selected_surface', { platform });
    if (!row.install_target) addIssue(issues, 'row_missing_install_target', 'Smoke plan row is missing install_target', { platform });
    if (row.realtime_startup_ready !== true) addIssue(issues, 'row_not_realtime_startup_ready', 'Smoke plan row is not realtime startup ready', { platform });
    if (row.adapter_blueprint_ready !== true) addIssue(issues, 'row_adapter_blueprint_not_ready', 'Smoke plan row adapter blueprint is not ready', { platform });
    if (!actions.includes('observe_platform_candidates')) addIssue(issues, 'row_missing_observe_candidates_step', 'Smoke plan row must include observe_platform_candidates step', { platform });
    if (!actions.includes('insert_annotation')) addIssue(issues, 'row_missing_insert_annotation_step', 'Smoke plan row must include insert_annotation step', { platform });
    if (observeStep?.client_method !== 'observePlatformCandidates') {
      addIssue(issues, 'row_invalid_observe_candidates_client_method', 'observe_platform_candidates step must use observePlatformCandidates', {
        platform,
        client_method: observeStep?.client_method,
      });
    }
    if (insertStep?.client_method !== 'insertAnnotation') {
      addIssue(issues, 'row_invalid_insert_annotation_client_method', 'insert_annotation step must use insertAnnotation', {
        platform,
        client_method: insertStep?.client_method,
      });
    }
    if (insertStep?.input?.annotation?.captured_at_ms == null) {
      addIssue(issues, 'row_insert_annotation_missing_captured_at_ms', 'insert_annotation smoke step must include annotation.captured_at_ms', { platform });
    }
    if (observeStep?.input?.captured_at_ms == null) {
      addIssue(issues, 'row_observe_candidates_missing_captured_at_ms', 'observe_platform_candidates smoke step must include captured_at_ms', { platform });
    }
  }

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_ACCEPTANCE_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: plan.target,
    package_id: plan.package_id,
    platform_count: plan.platform_count ?? 0,
    row_count: rows.length,
    required_step_count: rows.reduce((count, row) => count + (row.steps ?? []).filter((step) => step.required).length, 0),
    optional_step_count: rows.reduce((count, row) => count + (row.steps ?? []).filter((step) => !step.required).length, 0),
    timestamp_field: plan.timestamp_field,
    runtime_event_endpoint: plan.runtime_event_endpoint,
    smoke_plan_accepted: plan.accepted === true,
    issue_count: issues.length,
    issues,
    rows: rows.map((row) => compactObject({
      platform: row.platform,
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      required_action_count: row.required_action_count,
      optional_action_count: row.optional_action_count,
    })),
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue.code}`),
        ...(plan.next_actions ?? []),
      ])
      : plan.next_actions ?? [],
  });
}

export function assertMeetingAppTimelineConnectorSmokePlan(planOrChecklistOrPackage = {}, options = {}) {
  const report = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(planOrChecklistOrPackage, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector smoke plan is not accepted', {
      report,
      issues: report.issues,
    });
  }
  return planOrChecklistOrPackage;
}

export function createMeetingAppTimelineConnectorRuntimeClient(pkg = {}, options = {}) {
  const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(pkg, options);
  if (acceptance.accepted !== true && options.assertPackage !== false && options.assert_package !== false) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector package is not accepted for runtime client', {
      report: acceptance,
      issues: acceptance.issues,
    });
  }
  const endpoint = firstNonEmpty(
    options.endpoint,
    options.runtimeEventEndpoint,
    options.runtime_event_endpoint,
    pkg.runtime_events?.endpoint,
  );
  if (!endpoint) {
    throw new MeetingTimelineSdkError('runtime event endpoint is required for connector runtime client');
  }
  const platforms = selectedPlatforms(pkg, options);
  const surfaces = selectedSurfaces(pkg, options);
  const actionRows = runtimeActionRows(pkg);
  const actionSet = runtimeActionSet(pkg);
  const platformActionMap = runtimePlatformActionMap(pkg);
  const client = createMeetingPlatformRuntimeEventClient({
    ...options,
    endpoint,
  });

  function assertSupported(actionInput, platformInput, supportOptions = {}) {
    const action = normalizeMeetingPlatformRuntimeEventAction(actionInput);
    const platform = platformInput == null ? undefined : normalizeKey(platformInput);
    if (supportOptions.allowUnsupportedAction === true || supportOptions.allow_unsupported_action === true) {
      return { action, platform, supported: true, bypassed: true };
    }
    if (!actionSet.has(action)) {
      throw new MeetingTimelineSdkError(`Runtime action ${action} is not present in connector package`, {
        action,
        platform,
        package_id: pkg.id,
      });
    }
    if (platform) {
      const actions = platformActionMap.get(platform);
      if (!actions?.has(action)) {
        throw new MeetingTimelineSdkError(`Runtime action ${action} is not supported for ${platform}`, {
          action,
          platform,
          package_id: pkg.id,
        });
      }
    }
    return { action, platform, supported: true };
  }

  function buildEvent(eventInput = {}, eventOptions = {}) {
    const event = buildMeetingPlatformRuntimeEvent(eventInput, {
      ...options,
      ...eventOptions,
    });
    assertSupported(event.action, event.platform, eventOptions);
    return event;
  }

  async function send(eventInput = {}, sendOptions = {}) {
    const event = eventInput?.schema === 'meeting_platform_runtime_event'
      ? eventInput
      : buildEvent(eventInput, sendOptions);
    assertSupported(event.action, event.platform, sendOptions);
    return client.send(event, sendOptions);
  }

  return {
    type: MEETING_APP_TIMELINE_CONNECTOR_RUNTIME_CLIENT_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_RUNTIME_CLIENT_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    package_id: pkg.id,
    endpoint,
    platforms,
    surfaces,
    acceptance,
    runtime_event_client: client,
    supported_action_count: actionRows.length,
    supported_actions: unique(actionRows.map((row) => row.action)),
    supported_actions_by_platform: Object.fromEntries([...platformActionMap.entries()].map(([platform, actions]) => [
      platform,
      [...actions],
    ])),
    action_rows: actionRows,
    supports(action, platform) {
      try {
        assertSupported(action, platform);
        return true;
      } catch {
        return false;
      }
    },
    assertSupported,
    buildEvent,
    send,
    observeMeetingApp(platform, snapshot = {}, observeOptions = {}) {
      assertSupported('observe_meeting_app', platform, observeOptions);
      return client.observeMeetingApp(platform, snapshot, observeOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      assertSupported('observe_platform_candidates', undefined, observeOptions);
      return client.observePlatformCandidates(input, observeOptions);
    },
    ingestProvider(platform, payload = {}, ingestOptions = {}) {
      assertSupported('provider_event', platform, ingestOptions);
      return client.ingestProvider(platform, payload, ingestOptions);
    },
    insertAnnotation(platform, annotationInput = {}, markOptions = {}) {
      assertSupported('insert_annotation', platform, markOptions);
      return client.insertAnnotation(platform, annotationInput, markOptions);
    },
    insertMark(platform, annotationInput = {}, markOptions = {}) {
      return this.insertAnnotation(platform, annotationInput, markOptions);
    },
    speakerTrack(platform, input = {}, trackOptions = {}) {
      assertSupported('speaker_track', platform, trackOptions);
      return client.speakerTrack(platform, input, trackOptions);
    },
    participantTrack(platform, input = {}, trackOptions = {}) {
      assertSupported('participant_track', platform, trackOptions);
      return client.participantTrack(platform, input, trackOptions);
    },
    timelineView(platform, input = {}, viewOptions = {}) {
      assertSupported('timeline_view', platform, viewOptions);
      return client.timelineView(platform, input, viewOptions);
    },
    adapterRoute(platform, routeOptions = {}) {
      assertSupported('adapter_route', platform, routeOptions);
      return client.adapterRoute(platform, routeOptions);
    },
    adapterRoutes(routeOptions = {}) {
      assertSupported('adapter_routes', undefined, routeOptions);
      return client.adapterRoutes(routeOptions);
    },
    adapterBlueprint(platform, blueprintOptions = {}) {
      assertSupported('adapter_blueprint', platform, blueprintOptions);
      return client.adapterBlueprint(platform, blueprintOptions);
    },
    adapterBlueprints(blueprintOptions = {}) {
      assertSupported('adapter_blueprints', undefined, blueprintOptions);
      return client.adapterBlueprints(blueprintOptions);
    },
    runManifest(manifestOptions = {}) {
      assertSupported('run_manifest', undefined, manifestOptions);
      return client.runManifest(manifestOptions);
    },
    runHandoffReadiness(readinessOptions = {}) {
      assertSupported('run_handoff_readiness', undefined, readinessOptions);
      return client.runHandoffReadiness(readinessOptions);
    },
  };
}
