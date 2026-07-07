import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformRuntimeEvent,
  createMeetingPlatformRuntimeEventClient,
  normalizeMeetingPlatformRuntimeEventAction,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformConsumerHandoff,
} from './platform-consumer-handoff.mjs';
import {
  buildMeetingPlatformFieldIntakeMatrix,
} from './platform-field-intake.mjs';
import {
  createMeetingPlatformConnectorContentScriptBridge,
} from './meeting-platform-connector.mjs';

export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA = 'meeting_app_timeline_connector_package';
export const MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_package_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_HANDOFF_SCHEMA = 'meeting_app_timeline_connector_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA = 'meeting_app_timeline_connector_host_install_checklist';
export const MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_host_install_checklist_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA = 'meeting_app_timeline_connector_adoption_index';
export const MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA = 'meeting_app_timeline_connector_field_intake_index';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA = 'meeting_app_timeline_connector_bridge_handoff';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_bridge_handoff_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_SMOKE_REPORT_SCHEMA = 'meeting_app_timeline_connector_bridge_smoke_report';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA = 'meeting_app_timeline_connector_smoke_plan';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_ACCEPTANCE_SCHEMA = 'meeting_app_timeline_connector_smoke_plan_acceptance_report';
export const MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA = 'meeting_app_timeline_connector_smoke_run_report';
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

function bridgeSmokeNode(tagName, attrs = {}, text = '') {
  return {
    tagName: String(tagName).toUpperCase(),
    attributes: attrs,
    dataset: Object.fromEntries(Object.entries(attrs)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [
        key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()),
        value,
      ])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attrs[name] ?? null;
    },
  };
}

function bridgeSmokeSelectorMatches(item, selector) {
  const text = String(selector);
  if (text === '*') return true;
  if (text === 'button') return item.tagName === 'BUTTON';
  if (text.includes('speaking')) {
    return /speaking|active speaker|正在发言|正在讲话|正在说话/i.test(item.attributes?.['aria-label'] ?? '');
  }
  const attrParts = [...text.matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (!attrParts.length) return false;
  return attrParts.every((match) => {
    const [, attrName, operator, doubleQuoted, singleQuoted, bare] = match;
    const actual = item.attributes?.[attrName];
    if (operator == null) return actual != null;
    if (actual == null) return false;
    const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
    if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
    return String(actual) === String(expected);
  });
}

function createBridgeSmokeWindow(platform, options = {}) {
  const url = firstNonEmpty(options.url, options.meeting_url, sampleUrlForPlatform(platform));
  const nodes = asArray(options.nodes).length > 0
    ? asArray(options.nodes)
    : [
      bridgeSmokeNode('button', { 'aria-label': 'Leave call' }),
      bridgeSmokeNode('button', { 'aria-label': 'Participants' }),
      bridgeSmokeNode('div', {
        'data-participant-id': `${normalizeKey(platform)}-speaker-1`,
        'aria-label': 'Speaker 1 is speaking',
      }),
    ];
  const document = {
    nodeType: 9,
    title: 'Connector bridge smoke',
    hidden: false,
    location: { href: url },
    body: { nodeType: 1 },
    documentElement: { nodeType: 1 },
    querySelectorAll(selector) {
      return nodes.filter((item) => bridgeSmokeSelectorMatches(item, selector));
    },
  };
  return {
    document,
    location: { href: url, origin: new URL(url).origin },
    navigator: { userAgent: 'Connector bridge smoke fixture' },
    addEventListener() {},
    removeEventListener() {},
  };
}

function createBridgeSmokeRecordingFetch(calls = []) {
  return async (url, init = {}) => {
    const body = init.body ? JSON.parse(init.body) : {};
    calls.push(compactObject({
      url,
      method: init.method,
      action: body.action,
      platform: body.platform,
      body,
    }));
    const payload = JSON.stringify({
      ok: true,
      accepted: true,
      action: body.action,
      platform: body.platform,
      annotation: body.annotation,
      event: body,
    });
    return new Response(payload, {
      status: 201,
      headers: { 'content-type': 'application/json' },
    });
  };
}

function summarizeBridgeSmokeCall(call = {}) {
  return compactObject({
    method: call.method,
    action: call.action,
    platform: call.platform,
    captured_at_ms: call.body?.captured_at_ms ?? call.body?.annotation?.captured_at_ms,
    annotation_id: call.body?.annotation?.id,
    url: call.url,
  });
}

function smokeStepCapturedAtMs(step = {}) {
  return firstNonEmpty(
    step.input?.annotation?.captured_at_ms,
    step.input?.captured_at_ms,
    step.input?.speaker?.captured_at_ms,
    step.input?.participant?.captured_at_ms,
  );
}

function summarizeSmokeRunResult(result = {}) {
  if (!result || typeof result !== 'object') return { result_type: typeof result };
  return compactObject({
    ok: result.ok,
    accepted: result.accepted,
    schema: result.schema,
    action: result.action,
    method: result.method,
    platform: result.platform,
    captured_at_ms: result.captured_at_ms,
    status: result.status,
  });
}

function summarizeSmokeRunCall(call = {}) {
  return compactObject({
    method: call.method,
    platform: call.platform,
    action: call.action,
    captured_at_ms: smokeStepCapturedAtMs({ input: call.payload }),
    label: call.payload?.annotation?.label ?? call.payload?.label,
    speaker_id: call.payload?.speaker?.id ?? call.payload?.speaker_id,
    participant_id: call.payload?.participant?.id ?? call.payload?.participant_id,
    candidate_count: call.payload?.tabs?.length ?? call.payload?.candidates?.length,
  });
}

function createRecordingConnectorSmokeRuntimeClient(calls = []) {
  const push = (method, action, platform, payload = {}, options = {}) => {
    const call = compactObject({
      method,
      action,
      platform,
      payload,
      options,
    });
    calls.push(call);
    return compactObject({
      ok: true,
      method,
      action,
      platform,
      captured_at_ms: smokeStepCapturedAtMs({ input: payload }),
    });
  };
  return {
    observePlatformCandidates(input = {}, options = {}) {
      return push('observePlatformCandidates', 'observe_platform_candidates', input.platform, input, options);
    },
    insertAnnotation(platform, annotation = {}, options = {}) {
      return push('insertAnnotation', 'insert_annotation', normalizeKey(platform), { platform: normalizeKey(platform), annotation }, options);
    },
    speakerTrack(platform, input = {}, options = {}) {
      return push('speakerTrack', 'speaker_track', normalizeKey(platform), { platform: normalizeKey(platform), ...input }, options);
    },
    participantTrack(platform, input = {}, options = {}) {
      return push('participantTrack', 'participant_track', normalizeKey(platform), { platform: normalizeKey(platform), ...input }, options);
    },
  };
}

async function runConnectorSmokeStep(client, row = {}, step = {}, options = {}) {
  const action = normalizeMeetingPlatformRuntimeEventAction(step.action);
  const platform = normalizeKey(firstNonEmpty(step.input?.platform, row.platform));
  const stepOptions = compactObject({
    smoke: true,
    dry_run: options.dryRun === true || options.dry_run === true,
    smoke_step_id: step.id,
    smoke_action: action,
  });
  const base = {
    id: step.id,
    order: step.order,
    action,
    client_method: step.client_method,
    required: step.required === true,
    platform,
    captured_at_ms: smokeStepCapturedAtMs(step),
  };
  try {
    let result;
    if (action === 'observe_platform_candidates') {
      result = await client.observePlatformCandidates(step.input ?? {}, stepOptions);
    } else if (action === 'insert_annotation') {
      result = await client.insertAnnotation(platform, step.input?.annotation ?? {}, stepOptions);
    } else if (action === 'speaker_track') {
      result = await client.speakerTrack(platform, {
        ...(step.input ?? {}),
        ...(step.input?.speaker ?? {}),
      }, stepOptions);
    } else if (action === 'participant_track') {
      result = await client.participantTrack(platform, {
        ...(step.input ?? {}),
        ...(step.input?.participant ?? {}),
      }, stepOptions);
    } else {
      throw new MeetingTimelineSdkError(`Unsupported connector smoke action: ${action}`, { action, step });
    }
    return compactObject({
      ...base,
      accepted: true,
      result: summarizeSmokeRunResult(result),
    });
  } catch (error) {
    return compactObject({
      ...base,
      accepted: false,
      error: {
        name: error.name,
        message: error.message,
      },
    });
  }
}

function connectorSmokeRunRowIssues(row = {}, stepResults = []) {
  const issues = [];
  const observeIndex = stepResults.findIndex((step) => step.action === 'observe_platform_candidates' && step.accepted === true);
  const insertIndex = stepResults.findIndex((step) => step.action === 'insert_annotation' && step.accepted === true);
  const insertStep = stepResults.find((step) => step.action === 'insert_annotation');
  const failedSteps = stepResults.filter((step) => step.accepted !== true);
  if (!row.selected_surface) issues.push('missing_selected_surface');
  if (!row.install_target) issues.push('missing_install_target');
  if (row.realtime_startup_ready !== true) issues.push('row_not_realtime_startup_ready');
  if (row.adapter_blueprint_ready !== true) issues.push('row_adapter_blueprint_not_ready');
  if (observeIndex < 0) issues.push('observe_candidates_not_executed');
  if (insertIndex < 0) issues.push('insert_annotation_not_executed');
  if (observeIndex < 0 || insertIndex <= observeIndex) issues.push('axis_not_observed_before_insert');
  if (insertStep?.captured_at_ms == null) issues.push('insert_annotation_missing_captured_at_ms');
  for (const failed of failedSteps) {
    issues.push(failed.required ? `required_step_failed:${failed.action}` : `declared_optional_step_failed:${failed.action}`);
  }
  return unique(issues);
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

function fieldIntakeMatrixForChecklist(checklist = {}, options = {}) {
  return buildMeetingPlatformFieldIntakeMatrix({
    ...options,
    baseUrl: firstNonEmpty(options.baseUrl, options.base_url, checklist.base_url),
    platforms: (checklist.rows ?? []).map((row) => normalizeKey(row.platform)),
  });
}

function summarizeFieldOperatorSteps(steps = []) {
  return asArray(steps).map((step) => compactObject({
    id: step.id,
    owner: step.owner,
    objective: step.objective,
    command: step.command,
    output: step.output,
    required_snapshot: step.required_snapshot,
    required_coverage: step.required_coverage,
  }));
}

export function buildMeetingAppTimelineConnectorFieldIntakeIndex(checklistOrPackage = {}, options = {}) {
  const checklist = checklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA
    ? checklistOrPackage
    : buildMeetingAppTimelineConnectorHostInstallChecklist(checklistOrPackage, options);
  const matrix = fieldIntakeMatrixForChecklist(checklist, options);
  const checklistByPlatform = new Map((checklist.rows ?? []).map((row) => [normalizeKey(row.platform), row]));
  const rows = (matrix.plans ?? []).map((plan) => {
    const platform = normalizeKey(plan.platform);
    const checklistRow = checklistByPlatform.get(platform) ?? {};
    return compactObject({
      platform,
      display_name: plan.display_name ?? checklistRow.display_name,
      status: plan.status,
      production_ready: plan.production_ready === true,
      ready_for_realtime_annotations: plan.ready_for_realtime_annotations === true,
      connector_selected_surface: checklistRow.selected_surface,
      connector_install_target: checklistRow.install_target,
      connector_runtime_preset: checklistRow.runtime_preset,
      provider_endpoint: plan.provider_endpoint,
      field_evidence_input: plan.files?.field_evidence_input,
      evidence_package: plan.files?.evidence_package,
      required_inputs: plan.required_inputs,
      forbidden_inputs: plan.forbidden_inputs,
      required_local_snapshots: plan.acceptance?.required_local_snapshots ?? [],
      required_provider_coverage: plan.acceptance?.required_provider_coverage ?? [],
      missing_env: plan.provider_connection?.security?.missing_env ?? [],
      provider_transport: plan.provider_connection?.transport,
      provider_realtime_blocking: plan.provider_connection?.realtime_annotation_policy?.provider_events_block_realtime,
      commands: {
        export_field_intake_plan: plan.commands?.export_field_intake_plan,
        build_field_evidence: plan.commands?.build_field_evidence,
        inspect_capture_gaps: plan.commands?.inspect_capture_gaps,
        validate_real_intake: plan.commands?.validate_real_intake,
        validate_live_readiness: plan.commands?.validate_live_readiness,
      },
      operator_steps: summarizeFieldOperatorSteps(plan.operator_steps),
      next_actions: plan.next_actions,
    });
  });
  const missing = rows.flatMap((row) => [
    row.field_evidence_input ? undefined : `${row.platform}:field_evidence_input`,
    row.evidence_package ? undefined : `${row.platform}:evidence_package`,
    row.commands?.build_field_evidence ? undefined : `${row.platform}:build_field_evidence_command`,
    row.commands?.validate_real_intake ? undefined : `${row.platform}:validate_real_intake_command`,
    (row.required_local_snapshots ?? []).length > 0 ? undefined : `${row.platform}:required_local_snapshots`,
    (row.required_provider_coverage ?? []).length > 0 ? undefined : `${row.platform}:required_provider_coverage`,
  ].filter(Boolean));
  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_FIELD_INTAKE_INDEX_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: missing.length === 0,
    target: checklist.target,
    package_id: checklist.package_id,
    base_url: checklist.base_url,
    platform_count: checklist.platform_count,
    row_count: rows.length,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    provider_blocked_count: matrix.provider_blocked_count,
    local_capture_needed_count: matrix.local_capture_needed_count,
    provider_capture_needed_count: matrix.provider_capture_needed_count,
    field_matrix_schema: matrix.schema,
    rows,
    issue_count: unique(missing).length,
    issues: unique(missing),
    next_actions: unique(rows.flatMap((row) => row.next_actions ?? [])),
  });
}

export function assertMeetingAppTimelineConnectorFieldIntakeIndex(checklistOrPackage = {}, options = {}) {
  const index = buildMeetingAppTimelineConnectorFieldIntakeIndex(checklistOrPackage, options);
  if (!index.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector field intake index is not accepted', {
      index,
      issues: index.issues,
    });
  }
  return index;
}

function adoptionEvidenceAccepted(options = {}, platform) {
  const key = normalizeKey(platform);
  const byPlatform = firstNonEmpty(
    options.productionEvidenceAcceptedByPlatform,
    options.production_evidence_accepted_by_platform,
    options.liveEvidenceAcceptedByPlatform,
    options.live_evidence_accepted_by_platform,
  ) ?? {};
  return byPlatform[key] === true
    || byPlatform[platform] === true
    || options.productionEvidenceAccepted === true
    || options.production_evidence_accepted === true
    || options.liveEvidenceAccepted === true
    || options.live_evidence_accepted === true;
}

function adoptionStatusFor({ realtimeReady, bridgeReady, productionEvidenceAccepted } = {}) {
  if (!realtimeReady) return 'blocked_before_realtime_pilot';
  if (!bridgeReady) return 'needs_bridge_install';
  if (!productionEvidenceAccepted) return 'pilot_ready_needs_live_evidence';
  return 'production_evidence_ready';
}

function adoptionNextActions(row = {}) {
  const actions = [];
  for (const missing of row.missing ?? []) actions.push(`fix_${missing}`);
  if (row.realtime_ready && row.bridge_ready && !row.production_evidence_accepted) {
    actions.push('capture_live_meeting_app_snapshot');
    actions.push('capture_provider_reconcile_records');
    actions.push('run_handoff_readiness_with_runtime_host_replay');
  }
  if (row.realtime_ready && row.bridge_ready && row.production_evidence_accepted) {
    actions.push('ship_platform_adapter_to_host_project');
  }
  return unique(actions);
}

export function buildMeetingAppTimelineConnectorAdoptionIndex(checklistOrPackage = {}, options = {}) {
  const checklist = checklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA
    ? checklistOrPackage
    : buildMeetingAppTimelineConnectorHostInstallChecklist(checklistOrPackage, options);
  const fieldIntakeIndex = buildMeetingAppTimelineConnectorFieldIntakeIndex(checklist, options);
  const bridgeHandoff = buildMeetingAppTimelineConnectorBridgeHandoff(checklist, options);
  const bridgeAcceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(bridgeHandoff, options);
  const smokePlan = buildMeetingAppTimelineConnectorSmokePlan(checklist, options);
  const smokeAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(smokePlan, options);
  const fieldIntakeByPlatform = new Map((fieldIntakeIndex.rows ?? []).map((row) => [normalizeKey(row.platform), row]));
  const bridgeByPlatform = new Map((bridgeHandoff.rows ?? []).map((row) => [normalizeKey(row.platform), row]));
  const smokeByPlatform = new Map((smokePlan.rows ?? []).map((row) => [normalizeKey(row.platform), row]));
  const rows = (checklist.rows ?? []).map((row) => {
    const platform = normalizeKey(row.platform);
    const actions = asArray(row.runtime_actions).map((action) => normalizeMeetingPlatformRuntimeEventAction(action));
    const fieldIntake = fieldIntakeByPlatform.get(platform) ?? {};
    const bridgeRow = bridgeByPlatform.get(platform) ?? {};
    const smokeRow = smokeByPlatform.get(platform) ?? {};
    const hasObserveCandidates = actions.includes('observe_platform_candidates');
    const hasInsertAnnotation = actions.includes('insert_annotation');
    const hasSpeakerTrack = actions.includes('speaker_track');
    const hasParticipantTrack = actions.includes('participant_track');
    const providerNonBlocking = row.provider_events_block_realtime === false;
    const transcriptNonBlocking = row.transcript_blocks_realtime === false;
    const bridgeMessages = new Set(bridgeRow.supported_message_types ?? []);
    const bridgeActions = new Set(bridgeRow.output_runtime_actions ?? []);
    const realtimeReady = row.realtime_startup_ready === true
      && row.adapter_blueprint_ready === true
      && hasObserveCandidates
      && hasInsertAnnotation
      && providerNonBlocking
      && transcriptNonBlocking
      && checklist.timestamp_field === 'captured_at_ms';
    const bridgeReady = bridgeRow.bridge_factory === 'installMeetingPlatformConnectorContentScriptBridge'
      && bridgeMessages.has('meeting_timeline.observe_candidates')
      && bridgeMessages.has('meeting_timeline.insert_mark')
      && bridgeActions.has('observe_platform_candidates')
      && bridgeActions.has('insert_annotation');
    const smokeReady = smokeRow.required_action_count >= 2
      && (smokeRow.steps ?? []).some((step) => step.action === 'observe_platform_candidates')
      && (smokeRow.steps ?? []).some((step) => step.action === 'insert_annotation');
    const productionEvidenceAccepted = adoptionEvidenceAccepted(options, platform);
    const missing = [
      row.selected_surface ? undefined : 'selected_surface',
      row.install_target ? undefined : 'install_target',
      row.realtime_startup_ready === true ? undefined : 'realtime_startup_ready',
      row.adapter_blueprint_ready === true ? undefined : 'adapter_blueprint_ready',
      hasObserveCandidates ? undefined : 'observe_platform_candidates_action',
      hasInsertAnnotation ? undefined : 'insert_annotation_action',
      providerNonBlocking ? undefined : 'provider_nonblocking_contract',
      transcriptNonBlocking ? undefined : 'transcript_nonblocking_contract',
      checklist.timestamp_field === 'captured_at_ms' ? undefined : 'captured_at_ms_contract',
      bridgeReady ? undefined : 'connector_bridge_runtime_path',
      smokeReady ? undefined : 'connector_smoke_plan',
    ].filter(Boolean);
    const adoptionRow = compactObject({
      platform,
      display_name: row.display_name,
      status: adoptionStatusFor({ realtimeReady, bridgeReady, productionEvidenceAccepted }),
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      runtime_preset: row.runtime_preset,
      p0_axis_bootstrap: {
        source: row.selected_surface === 'browser_extension'
          ? 'browser_extension_content_script_or_webview_preload'
          : row.selected_surface === 'native_detector'
            ? 'native_detector_or_desktop_accessibility'
            : row.selected_surface,
        first_action: row.observe_action ?? 'observePlatformCandidates',
        must_precede: 'insert_annotation',
        timestamp_field: checklist.timestamp_field,
      },
      p0_annotation_intake: {
        action: row.insert_action ?? 'insertAnnotation',
        required_field: 'captured_at_ms',
        provider_events_block_realtime: row.provider_events_block_realtime,
        transcript_blocks_realtime: row.transcript_blocks_realtime,
      },
      p1_tracks: {
        speaker_track_ready: hasSpeakerTrack,
        participant_track_ready: hasParticipantTrack,
        text_required: false,
      },
      p1_provider_reconcile: {
        required_for_realtime: false,
        required_for_production: platform !== 'local_detector',
        non_blocking_for_realtime: providerNonBlocking,
      },
      p2_post_meeting_backfill: {
        transcript_blocks_realtime: row.transcript_blocks_realtime,
        required_for_realtime: false,
      },
      field_intake: {
        status: fieldIntake.status,
        ready_for_realtime_annotations: fieldIntake.ready_for_realtime_annotations,
        provider_endpoint: fieldIntake.provider_endpoint,
        field_evidence_input: fieldIntake.field_evidence_input,
        evidence_package: fieldIntake.evidence_package,
        missing_env: fieldIntake.missing_env,
        required_local_snapshots: fieldIntake.required_local_snapshots,
        required_provider_coverage: fieldIntake.required_provider_coverage,
        commands: fieldIntake.commands,
      },
      runtime_actions: actions,
      client_methods: row.client_methods,
      bridge_ready: bridgeReady,
      bridge_factory: bridgeRow.bridge_factory,
      bridge_message_types: bridgeRow.supported_message_types,
      smoke_plan_ready: smokeReady,
      smoke_required_action_count: smokeRow.required_action_count,
      smoke_optional_action_count: smokeRow.optional_action_count,
      realtime_ready: realtimeReady,
      can_start_axis_before_provider: hasObserveCandidates && providerNonBlocking,
      can_insert_annotation_on_current_axis: hasInsertAnnotation && checklist.timestamp_field === 'captured_at_ms',
      production_evidence_accepted: productionEvidenceAccepted,
      production_evidence_required: productionEvidenceAccepted ? [] : [
        'live_meeting_app_snapshot',
        platform === 'local_detector' ? undefined : 'provider_reconcile_records',
        'runtime_host_replay',
      ].filter(Boolean),
      missing,
    });
    return {
      ...adoptionRow,
      next_actions: unique([
        ...adoptionNextActions(adoptionRow),
        ...(fieldIntake.next_actions ?? []),
      ]),
    };
  });
  const realtimeReadyCount = rows.filter((row) => row.realtime_ready === true).length;
  const bridgeReadyCount = rows.filter((row) => row.bridge_ready === true).length;
  const productionEvidenceReadyCount = rows.filter((row) => row.production_evidence_accepted === true).length;
  const issues = [
    ...(checklist.accepted === true ? [] : ['host_install_checklist_not_accepted']),
    ...(bridgeAcceptance.accepted === true ? [] : bridgeAcceptance.issues.map((issue) => `bridge:${issue.code}`)),
    ...(smokeAcceptance.accepted === true ? [] : smokeAcceptance.issues.map((issue) => `smoke:${issue.code}`)),
    ...rows.flatMap((row) => (row.realtime_ready && row.bridge_ready ? [] : row.missing.map((item) => `${row.platform}:${item}`))),
  ];
  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_ADOPTION_INDEX_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: checklist.target,
    package_id: checklist.package_id,
    runtime_event_endpoint: checklist.runtime_event_endpoint,
    timestamp_field: checklist.timestamp_field,
    platform_count: checklist.platform_count,
    row_count: rows.length,
    realtime_ready_count: realtimeReadyCount,
    bridge_ready_count: bridgeReadyCount,
    production_evidence_ready_count: productionEvidenceReadyCount,
    production_evidence_pending_count: Math.max(0, rows.length - productionEvidenceReadyCount),
    source_schemas: {
      host_install_checklist: checklist.schema,
      bridge_handoff: bridgeHandoff.schema,
      bridge_handoff_acceptance: bridgeAcceptance.schema,
      smoke_plan: smokePlan.schema,
      smoke_plan_acceptance: smokeAcceptance.schema,
      field_intake_index: fieldIntakeIndex.schema,
    },
    files_to_read_first: [
      'connector-adoption-index.json',
      'connector-field-intake-index.json',
      'host-install-checklist.json',
      'connector-bridge-handoff.json',
      'connector-bridge-smoke-report.json',
      'connector-smoke-plan.json',
      'startup-plan-matrix.json',
    ],
    rows,
    issue_count: unique(issues).length,
    issues: unique(issues),
    next_actions: issues.length > 0
      ? unique(rows.flatMap((row) => row.next_actions))
      : ['capture_live_evidence_before_production_rollout'],
  });
}

export function assertMeetingAppTimelineConnectorAdoptionIndex(checklistOrPackage = {}, options = {}) {
  const index = buildMeetingAppTimelineConnectorAdoptionIndex(checklistOrPackage, options);
  if (!index.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector adoption index is not accepted', {
      index,
      issues: index.issues,
    });
  }
  return index;
}

export function buildMeetingAppTimelineConnectorBridgeHandoff(pkgOrChecklist = {}, options = {}) {
  const isChecklist = pkgOrChecklist?.schema === MEETING_APP_TIMELINE_CONNECTOR_HOST_INSTALL_CHECKLIST_SCHEMA;
  const checklist = isChecklist
    ? pkgOrChecklist
    : buildMeetingAppTimelineConnectorHostInstallChecklist(pkgOrChecklist, options);
  const platforms = unique((checklist.rows ?? []).map((row) => normalizeKey(row.platform))).filter(Boolean);
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url, checklist.base_url, pkgOrChecklist.base_url);
  const consumerHandoff = buildMeetingPlatformConsumerHandoff({
    ...options,
    baseUrl,
    platforms,
  });
  const lightweight = consumerHandoff.lightweight_connector_handoff ?? {};
  const rows = (checklist.rows ?? []).map((row) => compactObject({
    platform: normalizeKey(row.platform),
    display_name: row.display_name,
    selected_surface: row.selected_surface,
    install_target: row.install_target,
    runtime_actions: row.runtime_actions,
    client_methods: row.client_methods,
    required_host_steps: row.required_host_steps,
    bridge_factory: lightweight.factories?.install_content_script_bridge,
    runtime_factory: lightweight.factories?.create_runtime,
    hub_factory: lightweight.factories?.create_hub,
    supported_message_types: lightweight.content_script_bridge?.message_types,
    output_runtime_actions: lightweight.content_script_bridge?.output_runtime_actions,
  }));
  const issues = [
    ...(checklist.accepted === true ? [] : ['host_install_checklist_not_accepted']),
    ...(lightweight.accepted === true ? [] : ['lightweight_connector_handoff_not_accepted']),
    ...((consumerHandoff.issues ?? [])
      .filter((issue) => issue.severity === 'error')
      .map((issue) => issue.code)),
  ];
  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: checklist.target,
    package_id: checklist.package_id,
    base_url: baseUrl,
    runtime_event_endpoint: checklist.runtime_event_endpoint ?? lightweight.runtime_event_endpoint,
    platform_count: platforms.length,
    row_count: rows.length,
    platforms,
    module: lightweight.module,
    factories: lightweight.factories,
    kit_methods: lightweight.kit_methods,
    message_contract: lightweight.content_script_bridge,
    host_requirements: lightweight.host_requirements,
    startup_order: [
      'create_or_load_connector_package',
      'read_connector_bridge_handoff',
      'install_platform_content_script_or_webview_preload_bridge',
      'send_meeting_timeline_observe_candidates_before_first_mark',
      'send_meeting_timeline_insert_mark_with_captured_at_ms',
      'send_speaker_and_participant_track_markers_when_available',
      'run_connector_smoke_plan_against_host_runtime_client',
    ],
    sample_messages: [
      {
        type: 'meeting_timeline.observe_candidates',
        payload: {
          captured_at_ms: 1_782_614_400_000,
          tabs: [{
            active: true,
            url: sampleUrlForPlatform(platforms[0] ?? 'google_meet'),
            title: 'Current meeting tab',
          }],
        },
      },
      {
        type: 'meeting_timeline.insert_mark',
        payload: {
          mark: {
            id: 'mark-1',
            label: 'why?',
            captured_at_ms: 1_782_614_401_000,
          },
        },
      },
      {
        type: 'meeting_timeline.sample_tracks',
        payload: {
          captured_at_ms: 1_782_614_402_000,
          speakers: [{ id: 'speaker-1', name: 'Speaker 1' }],
          participants: [{ id: 'participant-1', name: 'Participant 1' }],
        },
      },
    ],
    rows,
    issue_count: issues.length,
    issues: unique(issues),
    source_schemas: {
      host_install_checklist: checklist.schema,
      consumer_handoff: consumerHandoff.schema,
      lightweight_connector_handoff: lightweight.type,
    },
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue}`),
        ...(checklist.next_actions ?? []),
        ...(consumerHandoff.next_actions ?? []),
      ])
      : checklist.next_actions ?? consumerHandoff.next_actions ?? [],
  });
}

export function buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(handoffOrPackageOrChecklist = {}, options = {}) {
  const handoff = handoffOrPackageOrChecklist?.schema === MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA
    ? handoffOrPackageOrChecklist
    : buildMeetingAppTimelineConnectorBridgeHandoff(handoffOrPackageOrChecklist, options);
  const rows = handoff.rows ?? [];
  const issues = [];
  const messageTypes = new Set(handoff.message_contract?.message_types ?? []);
  const outputActions = new Set(handoff.message_contract?.output_runtime_actions ?? []);
  const requiredMessages = [
    'meeting_timeline.observe_candidates',
    'meeting_timeline.insert_mark',
    'meeting_timeline.sample_tracks',
  ];
  const requiredOutputActions = [
    'observe_platform_candidates',
    'insert_annotation',
    'speaker_track',
    'participant_track',
  ];

  if (handoff.schema !== MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA) {
    addIssue(issues, 'invalid_schema', 'Expected a meeting app timeline connector bridge handoff', {
      expected_schema: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA,
      actual_schema: handoff.schema,
    });
  }
  if (handoff.accepted !== true) addIssue(issues, 'bridge_handoff_not_accepted', 'Connector bridge handoff accepted flag is not true');
  if (!handoff.module) addIssue(issues, 'missing_module', 'Connector bridge handoff must declare SDK module');
  if (handoff.factories?.install_content_script_bridge !== 'installMeetingPlatformConnectorContentScriptBridge') {
    addIssue(issues, 'missing_install_content_script_bridge_factory', 'Bridge handoff must expose installMeetingPlatformConnectorContentScriptBridge', {
      actual_factory: handoff.factories?.install_content_script_bridge,
    });
  }
  if (handoff.factories?.create_hub !== 'createMeetingPlatformConnectorHub') {
    addIssue(issues, 'missing_create_hub_factory', 'Bridge handoff must expose createMeetingPlatformConnectorHub', {
      actual_factory: handoff.factories?.create_hub,
    });
  }
  if (!handoff.runtime_event_endpoint && !handoff.host_requirements?.runtime_event_endpoint) {
    addIssue(issues, 'missing_runtime_event_endpoint', 'Bridge handoff must declare runtime event endpoint');
  }
  if (handoff.host_requirements?.timestamp_field !== 'captured_at_ms') {
    addIssue(issues, 'invalid_timestamp_contract', 'Bridge handoff must require captured_at_ms', {
      timestamp_field: handoff.host_requirements?.timestamp_field,
    });
  }
  if (handoff.host_requirements?.provider_events_block_realtime !== false) {
    addIssue(issues, 'provider_events_block_realtime', 'Provider events must not block realtime annotation insertion');
  }
  for (const message of requiredMessages) {
    if (!messageTypes.has(message)) addIssue(issues, 'missing_message_type', 'Bridge handoff is missing required message type', { message_type: message });
  }
  for (const action of requiredOutputActions) {
    if (!outputActions.has(action)) addIssue(issues, 'missing_output_runtime_action', 'Bridge handoff is missing required output runtime action', { action });
  }
  if ((handoff.platform_count ?? 0) <= 0) addIssue(issues, 'missing_platforms', 'Bridge handoff must include at least one platform');
  if (rows.length !== handoff.platform_count) {
    addIssue(issues, 'row_count_mismatch', 'Bridge handoff row count must equal platform_count', {
      row_count: rows.length,
      platform_count: handoff.platform_count,
    });
  }
  for (const row of rows) {
    const platform = normalizeKey(row.platform);
    if (!platform) addIssue(issues, 'row_missing_platform', 'Bridge handoff row is missing platform');
    if (!row.selected_surface) addIssue(issues, 'row_missing_selected_surface', 'Bridge handoff row is missing selected_surface', { platform });
    if (!row.install_target) addIssue(issues, 'row_missing_install_target', 'Bridge handoff row is missing install_target', { platform });
    if (row.bridge_factory !== 'installMeetingPlatformConnectorContentScriptBridge') {
      addIssue(issues, 'row_missing_bridge_factory', 'Bridge handoff row must expose install bridge factory', {
        platform,
        bridge_factory: row.bridge_factory,
      });
    }
    if (row.client_methods?.insert_annotation !== 'insertAnnotation') {
      addIssue(issues, 'row_missing_insert_annotation_client_method', 'Bridge handoff row must expose insertAnnotation client method', {
        platform,
        client_method: row.client_methods?.insert_annotation,
      });
    }
    if (!(row.supported_message_types ?? []).includes('meeting_timeline.insert_mark')) {
      addIssue(issues, 'row_missing_insert_mark_message_type', 'Bridge handoff row must support meeting_timeline.insert_mark', { platform });
    }
  }

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_ACCEPTANCE_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: handoff.target,
    package_id: handoff.package_id,
    platform_count: handoff.platform_count ?? 0,
    row_count: rows.length,
    runtime_event_endpoint: handoff.runtime_event_endpoint ?? handoff.host_requirements?.runtime_event_endpoint,
    timestamp_field: handoff.host_requirements?.timestamp_field,
    module: handoff.module,
    install_content_script_bridge_factory: handoff.factories?.install_content_script_bridge,
    create_hub_factory: handoff.factories?.create_hub,
    required_message_types: requiredMessages,
    required_output_runtime_actions: requiredOutputActions,
    bridge_handoff_accepted: handoff.accepted === true,
    issue_count: issues.length,
    issues,
    rows: rows.map((row) => compactObject({
      platform: row.platform,
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      bridge_factory: row.bridge_factory,
      insert_annotation_client_method: row.client_methods?.insert_annotation,
    })),
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue.code}`),
        ...(handoff.next_actions ?? []),
      ])
      : handoff.next_actions ?? [],
  });
}

export function assertMeetingAppTimelineConnectorBridgeHandoff(handoffOrPackageOrChecklist = {}, options = {}) {
  const report = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(handoffOrPackageOrChecklist, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector bridge handoff is not accepted', {
      report,
      issues: report.issues,
    });
  }
  return handoffOrPackageOrChecklist;
}

export async function runMeetingAppTimelineConnectorBridgeSmoke(handoffOrPackageOrChecklist = {}, options = {}) {
  const handoff = handoffOrPackageOrChecklist?.schema === MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_HANDOFF_SCHEMA
    ? handoffOrPackageOrChecklist
    : buildMeetingAppTimelineConnectorBridgeHandoff(handoffOrPackageOrChecklist, options);
  const acceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(handoff);
  const calls = [];
  const platform = normalizeKey(firstNonEmpty(options.platform, options.platform_key, handoff.platforms?.[0], 'google_meet'));
  const baseMs = Number.isFinite(options.baseCapturedAtMs)
    ? options.baseCapturedAtMs
    : Number.isFinite(options.base_captured_at_ms)
      ? options.base_captured_at_ms
      : 1_782_614_400_000;
  const window = options.window ?? createBridgeSmokeWindow(platform, options);
  const fetchImpl = options.fetch ?? options.fetchImpl ?? createBridgeSmokeRecordingFetch(calls);
  const bridge = createMeetingPlatformConnectorContentScriptBridge({
    ...options,
    baseUrl: firstNonEmpty(options.baseUrl, options.base_url, handoff.base_url, 'https://timeline.example.com'),
    platforms: handoff.platforms,
    fetch: fetchImpl,
    window,
    now: () => baseMs,
    observeTracks: true,
    observe_tracks: true,
    trackRuntimeOptions: {
      insert: true,
      ...(options.trackRuntimeOptions ?? {}),
      ...(options.track_runtime_options ?? {}),
    },
  });
  const steps = [];
  const pushStep = (step) => {
    steps.push(compactObject(step));
    return step;
  };

  try {
    const before = calls.length;
    const result = await bridge.observePlatformCandidates({
      captured_at_ms: baseMs,
      tabs: [{
        active: true,
        url: sampleUrlForPlatform(platform),
        title: `${platform} bridge smoke`,
        in_meeting: true,
      }],
    });
    pushStep({
      id: 'observe_platform_candidates',
      mode: 'direct_method',
      accepted: result?.ok !== false,
      runtime_event_count: calls.length - before,
      result_action: result?.action,
    });
  } catch (error) {
    pushStep({
      id: 'observe_platform_candidates',
      mode: 'direct_method',
      accepted: false,
      error: { name: error.name, message: error.message },
    });
  }

  const messageSteps = [
    {
      id: 'insert_mark_message',
      message: {
        type: 'meeting_timeline.insert_mark',
        payload: {
          platform,
          mark: {
            id: `${platform}-bridge-smoke-mark`,
            label: 'bridge smoke mark',
            captured_at_ms: baseMs + 1_000,
          },
        },
      },
    },
    {
      id: 'sample_tracks_message',
      message: {
        type: 'meeting_timeline.sample_tracks',
        payload: {
          platform,
          captured_at_ms: baseMs + 2_000,
        },
      },
    },
    {
      id: 'preflight_current_window_message',
      message: {
        type: 'meeting_timeline.preflight_current_window',
        payload: {
          platform,
          options: {
            requireSpeakerTrack: true,
            observedAtMs: baseMs + 3_000,
          },
        },
      },
    },
  ];

  for (const item of messageSteps) {
    try {
      const before = calls.length;
      const result = await bridge.dispatchMessage(item.message);
      pushStep({
        id: item.id,
        mode: 'dispatch_message',
        message_type: item.message.type,
        accepted: result?.handled === true && result?.result?.ok !== false,
        handled: result?.handled === true,
        action: result?.action,
        runtime_event_count: calls.length - before,
        result_schema: result?.result?.schema,
      });
    } catch (error) {
      pushStep({
        id: item.id,
        mode: 'dispatch_message',
        message_type: item.message.type,
        accepted: false,
        error: { name: error.name, message: error.message },
      });
    }
  }

  const callActions = calls.map((call) => normalizeMeetingPlatformRuntimeEventAction(call.action));
  const observeIndex = callActions.indexOf('observe_platform_candidates');
  const insertIndex = callActions.indexOf('insert_annotation');
  const issues = [
    ...(acceptance.accepted ? [] : acceptance.issues.map((issue) => `acceptance:${issue.code}`)),
    ...(steps.every((step) => step.accepted === true) ? [] : steps.filter((step) => step.accepted !== true).map((step) => `step_failed:${step.id}`)),
    ...(callActions.includes('observe_platform_candidates') ? [] : ['missing_observe_platform_candidates_runtime_event']),
    ...(callActions.includes('insert_annotation') ? [] : ['missing_insert_annotation_runtime_event']),
    ...(observeIndex >= 0 && insertIndex > observeIndex ? [] : ['axis_not_observed_before_insert']),
  ];

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_SMOKE_REPORT_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_BRIDGE_SMOKE_REPORT_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: issues.length === 0,
    bridge_handoff_accepted: handoff.accepted === true,
    bridge_handoff_acceptance_accepted: acceptance.accepted === true,
    dry_run: fetchImpl !== options.fetch && fetchImpl !== options.fetchImpl,
    target: handoff.target,
    package_id: handoff.package_id,
    platform,
    platform_count: handoff.platform_count,
    runtime_event_endpoint: handoff.runtime_event_endpoint ?? handoff.host_requirements?.runtime_event_endpoint,
    timestamp_field: handoff.host_requirements?.timestamp_field,
    step_count: steps.length,
    accepted_step_count: steps.filter((step) => step.accepted === true).length,
    runtime_event_count: calls.length,
    runtime_event_actions: unique(callActions),
    observe_before_insert: observeIndex >= 0 && insertIndex > observeIndex,
    steps,
    calls: calls.map((call) => summarizeBridgeSmokeCall(call)),
    issue_count: issues.length,
    issues: unique(issues),
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue}`),
        ...(handoff.next_actions ?? []),
      ])
      : handoff.next_actions ?? [],
  });
}

export async function assertMeetingAppTimelineConnectorBridgeSmoke(handoffOrPackageOrChecklist = {}, options = {}) {
  const report = await runMeetingAppTimelineConnectorBridgeSmoke(handoffOrPackageOrChecklist, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector bridge smoke failed', {
      report,
      issues: report.issues,
    });
  }
  return report;
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

export async function runMeetingAppTimelineConnectorSmokePlan(planOrChecklistOrPackage = {}, options = {}) {
  const plan = planOrChecklistOrPackage?.schema === MEETING_APP_TIMELINE_CONNECTOR_SMOKE_PLAN_SCHEMA
    ? planOrChecklistOrPackage
    : buildMeetingAppTimelineConnectorSmokePlan(planOrChecklistOrPackage, options);
  const planAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(plan);
  const calls = [];
  const explicitClient = firstNonEmpty(options.client, options.runtimeClient, options.runtime_client);
  const dryRun = explicitClient
    ? options.dryRun === true || options.dry_run === true
    : options.dryRun !== false && options.dry_run !== false;
  const client = explicitClient
    ?? (dryRun
      ? createRecordingConnectorSmokeRuntimeClient(calls)
      : createMeetingAppTimelineConnectorRuntimeClient(planOrChecklistOrPackage, options));
  const rows = [];

  for (const row of plan.rows ?? []) {
    const sortedSteps = [...(row.steps ?? [])].sort((left, right) => (left.order ?? 0) - (right.order ?? 0));
    const stepResults = [];
    for (const step of sortedSteps) {
      stepResults.push(await runConnectorSmokeStep(client, row, step, { ...options, dryRun }));
    }
    const issues = connectorSmokeRunRowIssues(row, stepResults);
    const observeAcceptedIndex = stepResults.findIndex((step) => step.action === 'observe_platform_candidates' && step.accepted === true);
    const insertAcceptedIndex = stepResults.findIndex((step) => step.action === 'insert_annotation' && step.accepted === true);
    const insertStep = stepResults.find((step) => step.action === 'insert_annotation');
    rows.push(compactObject({
      platform: normalizeKey(row.platform),
      selected_surface: row.selected_surface,
      install_target: row.install_target,
      accepted: issues.length === 0,
      required_step_count: stepResults.filter((step) => step.required).length,
      optional_step_count: stepResults.filter((step) => !step.required).length,
      executed_step_count: stepResults.length,
      failed_step_count: stepResults.filter((step) => step.accepted !== true).length,
      observe_before_insert: observeAcceptedIndex >= 0 && insertAcceptedIndex > observeAcceptedIndex,
      captured_at_ms_preserved: insertStep?.captured_at_ms != null && (insertStep.result?.captured_at_ms == null || insertStep.result.captured_at_ms === insertStep.captured_at_ms),
      steps: stepResults,
      issues,
    }));
  }

  const issues = [
    ...(planAcceptance.accepted ? [] : planAcceptance.issues.map((issue) => `plan:${issue.code}`)),
    ...rows.flatMap((row) => row.issues.map((issue) => `${row.platform}:${issue}`)),
  ];

  return compactObject({
    type: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA,
    schema: MEETING_APP_TIMELINE_CONNECTOR_SMOKE_RUN_REPORT_SCHEMA,
    schema_version: MEETING_APP_TIMELINE_CONNECTOR_PACKAGE_SCHEMA_VERSION,
    accepted: planAcceptance.accepted === true && rows.every((row) => row.accepted === true),
    dry_run: dryRun,
    plan_accepted: planAcceptance.accepted === true,
    target: plan.target,
    package_id: plan.package_id,
    platform_count: plan.platform_count ?? rows.length,
    row_count: rows.length,
    timestamp_field: plan.timestamp_field,
    runtime_event_endpoint: plan.runtime_event_endpoint,
    required_step_count: rows.reduce((count, row) => count + row.required_step_count, 0),
    optional_step_count: rows.reduce((count, row) => count + row.optional_step_count, 0),
    executed_step_count: rows.reduce((count, row) => count + row.executed_step_count, 0),
    failed_step_count: rows.reduce((count, row) => count + row.failed_step_count, 0),
    call_count: calls.length,
    calls: calls.map((call) => summarizeSmokeRunCall(call)),
    issue_count: issues.length,
    issues,
    rows,
    next_actions: issues.length > 0
      ? unique([
        ...issues.map((issue) => `fix_${issue}`),
        ...(plan.next_actions ?? []),
      ])
      : plan.next_actions ?? [],
  });
}

export async function assertMeetingAppTimelineConnectorSmokeRun(planOrChecklistOrPackage = {}, options = {}) {
  const report = await runMeetingAppTimelineConnectorSmokePlan(planOrChecklistOrPackage, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app timeline connector smoke run failed', {
      report,
      issues: report.issues,
    });
  }
  return report;
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
