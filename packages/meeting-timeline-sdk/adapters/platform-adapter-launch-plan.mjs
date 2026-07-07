import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import {
  buildMeetingPlatformAdapterCandidatePreflight,
} from './platform-adapter-preflight.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA = 'meeting_platform_adapter_launch_plan';
export const MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA = 'meeting_platform_adapter_candidate_launch_plan';
export const MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA_VERSION = 1;

const SURFACE_ALIASES = Object.freeze({
  browser: 'browser_extension',
  browser_extension: 'browser_extension',
  'browser-extension': 'browser_extension',
  extension: 'browser_extension',
  webview: 'webview_preload',
  webview_preload: 'webview_preload',
  'webview-preload': 'webview_preload',
  electron: 'webview_preload',
  native: 'native_detector',
  native_detector: 'native_detector',
  'native-detector': 'native_detector',
  native_host: 'native_detector',
  'native-host': 'native_detector',
  host: 'native_detector',
  provider: 'provider_reconcile',
  provider_reconcile: 'provider_reconcile',
  'provider-reconcile': 'provider_reconcile',
});

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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function normalizeSurface(value) {
  const raw = String(value ?? '').trim();
  return SURFACE_ALIASES[raw] ?? SURFACE_ALIASES[raw.toLowerCase()] ?? raw;
}

function maybeNormalizePlatform(value) {
  if (value == null || value === '') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return undefined;
  }
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

function manifestFrom(manifestOrInput = {}, input = {}, options = {}) {
  if (manifestOrInput?.schema === 'meeting_platform_adapter_install_manifest') return manifestOrInput;
  return firstNonEmpty(
    options.installManifest,
    options.install_manifest,
    input.installManifest,
    input.install_manifest,
    manifestOrInput.installManifest,
    manifestOrInput.install_manifest,
    manifestOrInput.manifest,
  ) ?? {};
}

function rawUrl(input = {}, options = {}) {
  if (typeof input === 'string' || input instanceof URL) return String(input);
  return firstNonEmpty(
    options.url,
    options.href,
    options.meetingUrl,
    options.meeting_url,
    firstPath(input, [
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
  );
}

function explicitPlatform(input = {}, options = {}) {
  if (typeof input === 'string' || input instanceof URL) return maybeNormalizePlatform(options.platform);
  return maybeNormalizePlatform(firstNonEmpty(
    options.platform,
    options.platform_key,
    input.platform,
    input.platform_key,
    input.provider,
    input.adapter,
    input.meeting?.platform,
    input.current_meeting?.platform,
    input.currentMeeting?.platform,
    input.detected_meeting?.platform,
    input.detectedMeeting?.platform,
    input.tab?.platform,
    input.window?.platform,
  ));
}

function wildcardToRegExp(pattern) {
  const escaped = String(pattern)
    .replace(/[.+?^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*');
  return new RegExp(`^${escaped}$`, 'i');
}

function urlMatchesPattern(url, pattern) {
  if (!url || !pattern) return false;
  try {
    return wildcardToRegExp(pattern).test(String(url));
  } catch {
    return false;
  }
}

function browserScriptFor(manifest = {}, platform, url) {
  const scripts = asArray(manifest.browser_extension?.content_scripts);
  const platformScripts = scripts.filter((script) => script.platform === platform);
  return platformScripts.find((script) => asArray(script.matches).some((pattern) => urlMatchesPattern(url, pattern)))
    ?? platformScripts[0];
}

function platformRow(manifest = {}, platform) {
  return asArray(manifest.platform_registry).find((row) => row.platform === platform);
}

function adapterBlueprintRow(manifest = {}, platform) {
  return asArray(manifest.adapter_blueprints?.rows).find((row) => row.platform === platform);
}

function resolvePlatform(manifest = {}, input = {}, options = {}) {
  const url = rawUrl(input, options);
  const explicit = explicitPlatform(input, options);
  const detected = detectMeetingFromUrl({
    url,
    title: typeof input === 'object' ? firstPath(input, ['title', 'tab.title', 'window.title', 'meeting.title']) : undefined,
  });
  const browserMatch = asArray(manifest.browser_extension?.content_scripts).find((script) => (
    asArray(script.matches).some((pattern) => urlMatchesPattern(url, pattern))
  ));
  const platform = firstNonEmpty(explicit, detected?.platform, browserMatch?.platform);
  return {
    platform,
    detected_meeting: detected,
    detection_reason: explicit ? 'explicit_platform' : detected ? 'meeting_url' : browserMatch ? 'browser_match' : 'none',
    url,
  };
}

function selectedSurface(row = {}, input = {}, options = {}) {
  return normalizeSurface(firstNonEmpty(
    options.surface,
    options.preferredSurface,
    options.preferred_surface,
    input.surface,
    input.preferredSurface,
    input.preferred_surface,
    row.selected_surface,
  ));
}

function surfaceEntrypoint(manifest = {}, row = {}, surface = '', url) {
  if (surface === 'browser_extension') {
    const script = browserScriptFor(manifest, row.platform, url);
    return compactObject({
      surface,
      platform: row.platform,
      content_script: script,
      message_types: manifest.browser_extension?.message_types,
      first_message_type: 'meeting_timeline.observe_candidates',
      first_sdk_method: row.first_sdk_method ?? 'observePlatformCandidates',
      mark_insert_method: row.mark_insert_method ?? 'insertAnnotation',
    });
  }
  if (surface === 'webview_preload') {
    const registryRow = asArray(manifest.webview_preload?.rows).find((item) => item.platform === row.platform);
    return compactObject({
      surface,
      platform: row.platform,
      registry_row: registryRow,
      first_sdk_method: registryRow?.first_sdk_method ?? 'observeMeetingApp_or_observePlatformCandidates',
      mark_insert_method: registryRow?.mark_insert_method ?? 'insertAnnotation',
    });
  }
  if (surface === 'native_detector') {
    const registryRow = asArray(manifest.native_detector?.rows).find((item) => item.platform === row.platform)
      ?? asArray(manifest.native_host?.rows).find((item) => item.platform === row.platform);
    return compactObject({
      surface,
      platform: row.platform,
      registry_row: registryRow,
      first_sdk_method: registryRow?.first_sdk_method ?? 'observePlatformCandidates',
      mark_insert_method: registryRow?.mark_insert_method ?? 'insertAnnotation',
    });
  }
  if (surface === 'provider_reconcile') {
    return compactObject({
      surface,
      platform: row.platform,
      registry_row: asArray(manifest.provider_reconcile?.rows).find((item) => item.platform === row.platform),
      first_sdk_method: 'ingestProvider',
      realtime_blocking: false,
    });
  }
  return compactObject({
    surface,
    platform: row.platform,
    first_sdk_method: row.first_sdk_method,
    mark_insert_method: row.mark_insert_method,
  });
}

function runtimeActions(row = {}, surface = '') {
  const localSurface = surface !== 'provider_reconcile';
  return [
    localSurface ? {
      id: 'observe_platform_candidates',
      phase: 'local_axis',
      required: true,
      sdk_method: row.first_sdk_method ?? 'observePlatformCandidates',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    localSurface ? {
      id: 'insert_realtime_annotation',
      phase: 'realtime_mark',
      required: true,
      sdk_method: row.mark_insert_method ?? 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    row.speaker_track_method ? {
      id: 'emit_speaker_position',
      phase: 'optional_track',
      required: false,
      sdk_method: row.speaker_track_method,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    row.participant_track_method ? {
      id: 'emit_participant_position',
      phase: 'optional_track',
      required: false,
      sdk_method: row.participant_track_method,
      timestamp_field: 'captured_at_ms',
    } : undefined,
    {
      id: 'provider_reconcile_backfill',
      phase: 'post_axis_reconcile',
      required: false,
      sdk_method: 'ingestProvider',
      realtime_blocking: false,
    },
  ].filter(Boolean);
}

function readiness(manifest = {}, row = null, surface = '', entrypoint = {}, resolved = {}) {
  const issues = [
    manifest.schema === 'meeting_platform_adapter_install_manifest' ? undefined : issue('error', 'invalid_install_manifest', 'Launch plan requires an adapter install manifest.', {
      actual: manifest.schema,
    }),
    manifest.accepted === true ? undefined : issue('error', 'install_manifest_not_accepted', 'Adapter install manifest must be accepted before launch.'),
    resolved.platform ? undefined : issue('error', 'platform_not_detected', 'No supported meeting platform was detected from the current input.'),
    row?.platform ? undefined : issue('error', 'platform_not_installed', 'Detected meeting platform is not present in the install manifest.', {
      platform: resolved.platform,
    }),
    row?.ready === true ? undefined : issue('error', 'platform_not_ready', 'Installed platform row is not ready.', {
      platform: resolved.platform,
    }),
    surface ? undefined : issue('error', 'surface_not_selected', 'No runtime surface is selected for launch.', {
      platform: resolved.platform,
    }),
    surface === 'browser_extension' && !entrypoint.content_script ? issue('error', 'browser_content_script_not_registered', 'Browser extension launch requires a registered content script for this platform.', {
      platform: resolved.platform,
      url: resolved.url,
    }) : undefined,
    (surface === 'webview_preload' || surface === 'native_detector') && !entrypoint.registry_row ? issue('error', 'surface_not_registered_for_platform', 'Selected launch surface is not registered for this platform in the install manifest.', {
      platform: resolved.platform,
      surface,
    }) : undefined,
    surface === 'provider_reconcile' ? issue('error', 'provider_reconcile_not_realtime_launch_surface', 'Provider reconcile cannot start realtime annotation capture by itself.', {
      platform: resolved.platform,
    }) : undefined,
  ].filter(Boolean);
  return {
    accepted: issues.filter((item) => item.severity === 'error').length === 0,
    issue_count: issues.length,
    issues,
  };
}

function nextActions(plan = {}, ready = {}) {
  if (ready.accepted !== true) return unique([
    ...(ready.issues ?? []).map((item) => item.code),
    'fix_launch_plan_blockers',
  ]);
  return unique([
    'start_selected_surface_runtime',
    'observe_platform_candidates_before_first_mark',
    'insert_realtime_marks_with_captured_at_ms',
    'reconcile_provider_events_after_local_axis',
  ]);
}

function selectedCandidateRow(candidatePreflight = {}) {
  return asArray(candidatePreflight.rows).find((row) => row?.selected === true)
    ?? asArray(candidatePreflight.rows)[candidatePreflight.selected_candidate_index]
    ?? asArray(candidatePreflight.rows)[0]
    ?? {};
}

function candidateLaunchInput(candidatePreflight = {}, input = {}, options = {}) {
  const row = selectedCandidateRow(candidatePreflight);
  return compactObject({
    platform: firstNonEmpty(options.platform, input.platform, input.provider, row.platform, candidatePreflight.selected_platform),
    url: firstNonEmpty(options.url, options.href, input.url, input.href, row.url),
    title: firstNonEmpty(options.title, input.title, row.title),
    surface: firstNonEmpty(options.surface, options.preferredSurface, options.preferred_surface, input.surface, input.preferredSurface, input.preferred_surface),
    capturedAtMs: firstNonEmpty(options.capturedAtMs, options.captured_at_ms, input.capturedAtMs, input.captured_at_ms),
    captured_at_ms: firstNonEmpty(options.captured_at_ms, options.capturedAtMs, input.captured_at_ms, input.capturedAtMs),
    candidate_index: row.candidate_index,
    window_id: row.window_id,
    tab_id: row.tab_id,
    active: row.active,
  });
}

function candidateLaunchReadiness(candidatePreflight = {}, launchPlan = {}) {
  const issues = [
    candidatePreflight.accepted === true ? undefined : issue('error', 'candidate_preflight_not_accepted', 'No candidate tab/window has live evidence for realtime annotation launch.', {
      candidate_status: candidatePreflight.status,
      candidate_count: candidatePreflight.candidate_count,
      selected_platform: candidatePreflight.selected_platform,
    }),
    ...(launchPlan.readiness?.issues ?? []),
  ].filter(Boolean);
  return {
    accepted: issues.filter((item) => item.severity === 'error').length === 0,
    issue_count: issues.length,
    candidate_preflight_accepted: candidatePreflight.accepted === true,
    launch_plan_accepted: launchPlan.accepted === true,
    issues,
  };
}

function candidateLaunchStatus(candidatePreflight = {}, launchPlan = {}, ready = {}) {
  if (ready.accepted === true) return 'ready_for_realtime_launch';
  if (candidatePreflight.accepted !== true) return candidatePreflight.status ?? 'candidate_preflight_not_accepted';
  if (launchPlan.accepted !== true) return 'launch_plan_not_ready';
  return 'not_ready';
}

function candidateLaunchNextActions(candidatePreflight = {}, launchPlan = {}, ready = {}) {
  if (ready.accepted === true) return unique([
    'start_selected_surface_runtime',
    'open_adapter_session_from_launch_plan',
    'observe_platform_candidates_before_first_mark',
    'insert_realtime_marks_with_captured_at_ms',
  ]);
  return unique([
    ...(ready.issues ?? []).map((item) => item.code),
    ...(candidatePreflight.next_actions ?? []),
    ...(launchPlan.next_actions ?? []),
  ]);
}

export function buildMeetingPlatformAdapterLaunchPlan(manifestOrInput = {}, input = {}, options = {}) {
  const manifest = manifestFrom(manifestOrInput, input, options);
  const resolved = resolvePlatform(manifest, input, options);
  const row = platformRow(manifest, resolved.platform);
  const blueprint = adapterBlueprintRow(manifest, resolved.platform);
  const surface = selectedSurface(row, input, options);
  const entrypoint = row ? surfaceEntrypoint(manifest, row, surface, resolved.url) : {};
  const ready = readiness(manifest, row, surface, entrypoint, resolved);
  const plan = {
    type: 'meeting_platform_adapter_launch_plan',
    schema: MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA_VERSION,
    accepted: ready.accepted,
    platform: resolved.platform,
    detection_reason: resolved.detection_reason,
    selected_surface: surface,
    install_manifest_schema: manifest.schema,
    install_manifest_accepted: manifest.accepted === true,
    detected_meeting: resolved.detected_meeting,
    current_url: resolved.url,
    platform_row: row,
    adapter_blueprint: blueprint,
    surface_entrypoint: entrypoint,
    axis_contract: {
      timestamp_field: 'captured_at_ms',
      local_axis_first: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    },
    runtime_actions: row ? runtimeActions(row, surface) : [],
    mark_template: row ? {
      platform: row.platform,
      captured_at_ms: firstNonEmpty(options.capturedAtMs, options.captured_at_ms, input.capturedAtMs, input.captured_at_ms, 0),
      source: 'meeting_platform_adapter_launch_plan',
    } : undefined,
    readiness: ready,
  };
  return {
    ...plan,
    next_actions: nextActions(plan, ready),
  };
}

export function buildMeetingPlatformAdapterCandidateLaunchPlan(manifestOrInput = {}, input = {}, options = {}) {
  const manifest = manifestFrom(manifestOrInput, input, options);
  const candidatePreflight = buildMeetingPlatformAdapterCandidatePreflight(input, options);
  const selectedCandidate = selectedCandidateRow(candidatePreflight);
  const launchInput = candidateLaunchInput(candidatePreflight, input, options);
  const launchPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, launchInput, options);
  const ready = candidateLaunchReadiness(candidatePreflight, launchPlan);
  return compactObject({
    type: 'meeting_platform_adapter_candidate_launch_plan',
    schema: MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA_VERSION,
    accepted: ready.accepted,
    status: candidateLaunchStatus(candidatePreflight, launchPlan, ready),
    platform: launchPlan.platform,
    selected_platform: candidatePreflight.selected_platform,
    selected_surface: launchPlan.selected_surface,
    selected_candidate_index: candidatePreflight.selected_candidate_index,
    selected_candidate: selectedCandidate,
    install_manifest_schema: launchPlan.install_manifest_schema,
    install_manifest_accepted: launchPlan.install_manifest_accepted,
    candidate_preflight: candidatePreflight,
    launch_input: launchInput,
    launch_plan: launchPlan,
    axis_contract: launchPlan.axis_contract,
    adapter_blueprint: launchPlan.adapter_blueprint,
    runtime_actions: launchPlan.runtime_actions,
    mark_template: launchPlan.mark_template,
    readiness: ready,
    next_actions: candidateLaunchNextActions(candidatePreflight, launchPlan, ready),
  });
}

export function assertMeetingPlatformAdapterLaunchPlan(planOrInput = {}, input = {}, options = {}) {
  const plan = planOrInput.schema === MEETING_PLATFORM_ADAPTER_LAUNCH_PLAN_SCHEMA
    ? planOrInput
    : buildMeetingPlatformAdapterLaunchPlan(planOrInput, input, options);
  if (plan.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter launch plan is not ready', {
      code: 'meeting_platform_adapter_launch_plan_not_ready',
      issues: plan.readiness?.issues ?? [],
      platform: plan.platform,
      surface: plan.selected_surface,
    });
  }
  return plan;
}

export function assertMeetingPlatformAdapterCandidateLaunchPlan(planOrInput = {}, input = {}, options = {}) {
  const plan = planOrInput.schema === MEETING_PLATFORM_ADAPTER_CANDIDATE_LAUNCH_PLAN_SCHEMA
    ? planOrInput
    : buildMeetingPlatformAdapterCandidateLaunchPlan(planOrInput, input, options);
  if (plan.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter candidate launch plan is not ready', {
      code: 'meeting_platform_adapter_candidate_launch_plan_not_ready',
      status: plan.status,
      issues: plan.readiness?.issues ?? [],
      platform: plan.platform,
      surface: plan.selected_surface,
      selected_candidate: plan.selected_candidate,
    });
  }
  return plan;
}
