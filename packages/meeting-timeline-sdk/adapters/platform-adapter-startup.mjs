import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterDecision,
} from './platform-adapter-decision.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeProfile,
} from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA = 'meeting_platform_adapter_startup_plan';
export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_MATRIX_SCHEMA = 'meeting_platform_adapter_startup_plan_matrix';
export const MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA_VERSION = 1;

const DEFAULT_STARTUP_PLATFORMS = Object.freeze([
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

function selectedPlatforms(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    input.platforms,
    input.platform_keys,
    DEFAULT_STARTUP_PLATFORMS,
  )).map((platform) => String(platform).trim()).filter(Boolean));
}

function inputForPlatform(platform, input = {}, options = {}) {
  const sources = [
    input.inputs,
    input.inputByPlatform,
    input.input_by_platform,
    input.snapshots,
    input.snapshotByPlatform,
    input.snapshot_by_platform,
    options.inputs,
    options.inputByPlatform,
    options.input_by_platform,
    options.snapshots,
    options.snapshotByPlatform,
    options.snapshot_by_platform,
  ];
  const keys = unique([
    platform,
    String(platform).replaceAll('_', '-'),
    platform === 'microsoft_teams' ? 'teams' : undefined,
    platform === 'google_meet' ? 'google-meet' : undefined,
  ]);
  for (const source of sources) {
    if (!isPlainObject(source)) continue;
    for (const key of keys) {
      if (source[key] != null) return source[key];
    }
  }
  return input.input ?? input.snapshot ?? input.sample ?? options.input ?? options.snapshot ?? options.sample ?? {};
}

function inputUrl(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return String(input);
  return firstNonEmpty(
    input.url,
    input.href,
    input.meeting_url,
    input.meetingUrl,
    input.join_url,
    input.joinUrl,
    input.location?.href,
    input.window?.url,
    input.browser?.url,
    input.tab?.url,
    input.tabs?.find?.((tab) => tab?.active)?.url,
    input.tabs?.[0]?.url,
  );
}

function inputTitle(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return undefined;
  return firstNonEmpty(
    input.title,
    input.meeting_title,
    input.meetingTitle,
    input.window?.title,
    input.browser?.title,
    input.tab?.title,
    input.tabs?.find?.((tab) => tab?.active)?.title,
    input.tabs?.[0]?.title,
  );
}

function startupId(platform, surface, options = {}) {
  return String(firstNonEmpty(
    options.startupId,
    options.startup_id,
    platform && surface ? `${platform}-${surface}-adapter-startup` : 'meeting-platform-adapter-startup',
  ));
}

function contentScriptFor(bundle = {}, url) {
  const scripts = bundle.browser?.content_scripts ?? bundle.browser?.manifest?.content_scripts ?? [];
  if (!url) return scripts[0];
  return scripts.find((script) => (script.matches ?? []).some((pattern) => {
    const escaped = String(pattern)
      .replace(/[.+?^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*');
    try {
      return new RegExp(`^${escaped}$`, 'i').test(String(url));
    } catch {
      return false;
    }
  })) ?? scripts[0];
}

function installTarget(surface) {
  if (surface === 'browser_extension') return 'manifest_v3_content_script';
  if (surface === 'native_detector') return 'native_or_desktop_observer';
  if (surface === 'provider_reconcile') return 'provider_event_consumer';
  if (surface === 'webview_preload') return 'webview_or_electron_preload';
  return 'host_adapter';
}

function bridgeFor(surface, bundle = {}) {
  if (surface === 'browser_extension' || surface === 'webview_preload') {
    return {
      preferred_bridge: 'lightweight_connector_bridge',
      module: bundle.runtime?.lightweight_connector_bridge?.module,
      create_function: bundle.runtime?.lightweight_connector_bridge?.create_function,
      install_function: bundle.runtime?.lightweight_connector_bridge?.install_function,
      options: bundle.runtime?.lightweight_connector_bridge?.options,
      fallback_bridge: bundle.runtime?.content_script_bridge,
    };
  }
  if (surface === 'native_detector') {
    return {
      preferred_bridge: 'runtime_event_client_or_sdk_facade',
      module: '@ai-annotation/meeting-timeline-sdk',
      create_function: 'createMeetingAppTimelineSdk',
      first_method: 'observePlatformCandidates',
    };
  }
  if (surface === 'provider_reconcile') {
    return {
      preferred_bridge: 'provider_normalizer',
      module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      first_method: 'ingestProvider',
    };
  }
  return undefined;
}

function startupActions(surface, decision = {}, bundle = {}) {
  const localSurface = surface !== 'provider_reconcile';
  const actions = [
    {
      id: 'resolve_current_platform',
      phase: 'startup',
      required: true,
      sdk_method: 'platformAdapterDecision',
      input: ['url', 'title', 'platform_or_provider'],
    },
    surface === 'browser_extension' || surface === 'webview_preload' ? {
      id: 'install_page_bridge',
      phase: 'startup',
      required: true,
      bridge: bridgeFor(surface, bundle)?.preferred_bridge,
      install_function: bridgeFor(surface, bundle)?.install_function,
    } : undefined,
    localSurface ? {
      id: 'observe_axis',
      phase: 'axis',
      required: true,
      sdk_method: 'observePlatformCandidates',
      message_type: 'meeting_timeline.observe_candidates',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    localSurface ? {
      id: 'insert_realtime_annotation',
      phase: 'mark',
      required: true,
      sdk_method: 'insertAnnotation',
      message_type: 'meeting_timeline.insert_mark',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    decision.runtime_actions?.some((action) => action.action === 'speaker_track') ? {
      id: 'emit_speaker_position',
      phase: 'optional_track',
      required: false,
      sdk_method: 'speakerTrack',
      message_type: 'meeting_timeline.sample_tracks',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    decision.runtime_actions?.some((action) => action.action === 'participant_track') ? {
      id: 'emit_participant_position',
      phase: 'optional_track',
      required: false,
      sdk_method: 'participantTrack',
      message_type: 'meeting_timeline.sample_tracks',
      timestamp_field: 'captured_at_ms',
    } : undefined,
    {
      id: 'provider_reconcile',
      phase: surface === 'provider_reconcile' ? 'axis' : 'post_realtime_reconcile',
      required: surface === 'provider_reconcile',
      sdk_method: 'ingestProvider',
      blocks_realtime_annotation: false,
    },
  ];
  return actions.filter(Boolean);
}

function messageContract(surface, bundle = {}) {
  if (surface === 'provider_reconcile') {
    return {
      input: 'provider_webhook_or_long_connection_event',
      output_sdk_method: 'ingestProvider',
      realtime_axis_source: 'not_allowed_as_primary_realtime_axis',
    };
  }
  return {
    observe_candidates: bundle.messaging?.candidate_observation?.message_type ?? 'meeting_timeline.observe_candidates',
    insert_annotation: 'meeting_timeline.insert_mark',
    sample: 'meeting_timeline.sample',
    sample_tracks: 'meeting_timeline.sample_tracks',
    provider_event: 'meeting_timeline.provider_event',
    timestamp_field: 'captured_at_ms',
  };
}

function codeRefs(surface, bundle = {}) {
  if (surface === 'browser_extension' || surface === 'webview_preload') {
    return {
      import: "import { installMeetingPlatformConnectorContentScriptBridge } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector';",
      setup: 'installMeetingPlatformConnectorContentScriptBridge({ baseUrl, platforms: [platform], startOptions });',
      first_message: "chrome.runtime.sendMessage({ type: 'meeting_timeline.observe_candidates', tabs });",
      insert_message: "chrome.runtime.sendMessage({ type: 'meeting_timeline.insert_mark', payload: { mark } });",
    };
  }
  if (surface === 'native_detector') {
    return {
      import: "import { createMeetingAppTimelineSdk } from '@ai-annotation/meeting-timeline-sdk';",
      setup: 'const sdk = createMeetingAppTimelineSdk({ baseUrl, platforms: [platform] });',
      first_call: 'await sdk.observePlatformCandidates({ windows, captured_at_ms: Date.now() }, { remote: true });',
      insert_call: 'await sdk.insertAnnotation(platform, mark, { remote: true });',
    };
  }
  return {
    import: "import { createMeetingAppTimelineSdk } from '@ai-annotation/meeting-timeline-sdk';",
    setup: 'const sdk = createMeetingAppTimelineSdk({ baseUrl, platforms: [platform] });',
    first_call: 'await sdk.ingestProvider(platform, providerEvent, { remote: true });',
    note: 'provider events are reconcile/backfill signals and must not gate realtime annotation insertion',
  };
}

function readinessIssues(decision = {}, surface, bundle = {}) {
  const issues = [];
  if (decision.accepted !== true) {
    issues.push({
      severity: 'error',
      code: 'decision_not_accepted',
      message: 'Adapter decision is not accepted for the current input.',
    });
  }
  if (surface === 'provider_reconcile') {
    issues.push({
      severity: 'warn',
      code: 'provider_reconcile_not_realtime_surface',
      message: 'Provider reconcile can align provider events later, but it is not a primary realtime annotation surface.',
    });
  }
  if ((surface === 'browser_extension' || surface === 'webview_preload') && !bundle.runtime?.lightweight_connector_bridge?.install_function) {
    issues.push({
      severity: 'error',
      code: 'missing_content_script_bridge',
      message: 'Browser/WebView startup requires a content-script or preload bridge install function.',
    });
  }
  return issues;
}

export function buildMeetingPlatformAdapterStartupPlan(input = {}, options = {}) {
  const objectInput = typeof input === 'string' || input instanceof URL
    ? { url: String(input) }
    : (input ?? {});
  const decision = buildMeetingPlatformAdapterDecision(objectInput, options);
  if (!decision.platform) {
    return {
      type: 'meeting_platform_adapter_startup_plan',
      schema: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA_VERSION,
      accepted: false,
      realtime_startup_ready: false,
      status: 'missing_platform',
      decision,
      issues: decision.issues ?? [{
        severity: 'error',
        code: 'missing_platform',
        message: 'Startup plan requires an explicit platform or a detectable meeting URL.',
      }],
      next_actions: decision.next_actions ?? ['provide_platform_or_supported_meeting_url'],
    };
  }
  const platform = decision.platform;
  const bundle = buildMeetingPlatformRuntimeBundle(platform, {
    ...options,
    url: inputUrl(objectInput),
  });
  const profile = buildMeetingPlatformRuntimeProfile(platform, options);
  const surface = decision.selected_surface ?? 'browser_extension';
  const contentScript = contentScriptFor(bundle, inputUrl(objectInput));
  const issues = readinessIssues(decision, surface, bundle);
  const realtimeStartupReady = decision.accepted === true && surface !== 'provider_reconcile'
    && issues.every((item) => item.severity !== 'error');
  return compactObject({
    type: 'meeting_platform_adapter_startup_plan',
    schema: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA_VERSION,
    id: startupId(platform, surface, options),
    accepted: decision.accepted === true && issues.every((item) => item.severity !== 'error'),
    realtime_startup_ready: realtimeStartupReady,
    status: realtimeStartupReady ? 'ready_to_start_realtime_axis' : 'needs_startup_attention',
    platform,
    display_name: decision.display_name ?? bundle.display_name ?? profile.display_name,
    input: compactObject({
      url: inputUrl(objectInput),
      title: inputTitle(objectInput),
    }),
    selected_surface: surface,
    install_target: installTarget(surface),
    decision,
    runtime_contract: {
      annotation_timestamp_field: 'captured_at_ms',
      axis_timebase: 'absolute_unix_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      per_meeting_annotation_isolation_required: true,
    },
    browser: surface === 'browser_extension' ? {
      matches: bundle.browser?.matches ?? [],
      host_permissions: bundle.browser?.host_permissions ?? [],
      content_script: contentScript,
      manifest: bundle.browser?.manifest,
    } : undefined,
    bridge: bridgeFor(surface, bundle),
    runtime: {
      preset: bundle.runtime?.preset,
      start_options: bundle.runtime?.start_options,
      observation_loop: bundle.runtime?.observation_loop,
      observer_scheduler_ready: bundle.readiness?.observer_scheduler_ready,
      runtime_host_ready: bundle.readiness?.runtime_host_ready,
    },
    axis: {
      start_source: profile.axis?.start?.create_on,
      end_source: profile.axis?.end?.create_on,
      end_fallbacks: profile.axis?.end?.fallbacks ?? [],
      provider_reconcile_required_for_realtime: false,
    },
    actions: startupActions(surface, decision, bundle),
    message_contract: messageContract(surface, bundle),
    code_refs: codeRefs(surface, bundle),
    provider_reconcile: {
      required_for_realtime: false,
      endpoint: profile.axis?.provider_reconcile?.endpoint,
      events: {
        start: profile.axis?.start?.provider_reconcile_events ?? [],
        end: profile.axis?.end?.provider_reconcile_events ?? [],
      },
    },
    reports: options.includeReports === true || options.include_reports === true ? {
      bundle,
      profile,
    } : undefined,
    issues,
    next_actions: unique([
      ...issues.map((item) => item.code),
      ...(decision.next_actions ?? []),
      surface === 'provider_reconcile' ? 'add_local_surface_for_realtime_annotations' : 'wire_startup_plan_into_host_runtime',
      'send_captured_at_ms_with_every_annotation',
      'keep_provider_and_transcript_reconcile_nonblocking',
    ]),
  });
}

export function buildMeetingPlatformAdapterStartupPlanMatrix(input = {}, options = {}) {
  const platforms = selectedPlatforms(input, options);
  const plans = platforms.map((platform) => buildMeetingPlatformAdapterStartupPlan({
    ...inputForPlatform(platform, input, options),
    platform,
  }, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_startup_plan_matrix',
    schema: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_STARTUP_PLAN_SCHEMA_VERSION,
    platform_count: plans.length,
    accepted_count: plans.filter((plan) => plan.accepted).length,
    realtime_startup_ready_count: plans.filter((plan) => plan.realtime_startup_ready).length,
    browser_surface_count: plans.filter((plan) => plan.selected_surface === 'browser_extension').length,
    native_surface_count: plans.filter((plan) => plan.selected_surface === 'native_detector').length,
    provider_reconcile_surface_count: plans.filter((plan) => plan.selected_surface === 'provider_reconcile').length,
    platforms: plans.map((plan) => plan.platform).filter(Boolean),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      accepted: plan.accepted,
      realtime_startup_ready: plan.realtime_startup_ready,
      selected_surface: plan.selected_surface,
      install_target: plan.install_target,
      runtime_preset: plan.runtime?.preset,
      first_action: plan.actions?.[0]?.id,
      observe_action: plan.actions?.find((action) => action.id === 'observe_axis')?.sdk_method,
      insert_action: plan.actions?.find((action) => action.id === 'insert_realtime_annotation')?.sdk_method,
      provider_events_block_realtime: plan.runtime_contract?.provider_events_block_realtime,
      transcript_blocks_realtime: plan.runtime_contract?.transcript_blocks_realtime,
      issue_count: plan.issues?.length ?? 0,
      first_next_action: plan.next_actions?.[0],
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function assertMeetingPlatformAdapterStartupPlan(input = {}, options = {}) {
  const plan = buildMeetingPlatformAdapterStartupPlan(input, options);
  if (plan.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter startup plan is not accepted', {
      platform: plan.platform,
      status: plan.status,
      issues: plan.issues,
      next_actions: plan.next_actions,
      plan,
    });
  }
  return plan;
}

export function assertMeetingPlatformAdapterStartupPlanMatrix(input = {}, options = {}) {
  const matrix = buildMeetingPlatformAdapterStartupPlanMatrix(input, options);
  if (matrix.accepted_count !== matrix.platform_count) {
    throw new MeetingTimelineSdkError('Meeting platform adapter startup plan matrix is not accepted', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      failed_platforms: matrix.rows.filter((row) => row.accepted !== true).map((row) => row.platform),
      next_actions: matrix.next_actions,
      matrix,
    });
  }
  return matrix;
}
