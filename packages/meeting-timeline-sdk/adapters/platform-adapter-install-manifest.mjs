import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA,
  buildMeetingPlatformAdapterImportPlanMatrix,
} from './platform-adapter-import-plan.mjs';

export const MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA = 'meeting_platform_adapter_install_manifest';
export const MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA_VERSION = 1;

const DEFAULT_CONTENT_SCRIPT_JS = Object.freeze(['meeting-app-content-script.bundle.js']);
const DEFAULT_MESSAGE_TYPES = Object.freeze([
  'meeting_timeline.attached',
  'meeting_timeline.status',
  'meeting_timeline.observe_candidates',
  'meeting_timeline.preflight_current_window',
  'meeting_timeline.preflight_candidates',
  'meeting_timeline.insert_mark',
  'meeting_timeline.sample_tracks',
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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function isImportPlan(value = {}) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA
    || value?.type === 'meeting_platform_adapter_import_plan';
}

function isExportPackage(value = {}) {
  return value?.schema === 'meeting_platform_adapter_export_package'
    || value?.type === 'meeting_platform_adapter_export_package';
}

function planListFrom(plansOrInput = {}, input = {}, options = {}) {
  const explicit = Array.isArray(plansOrInput)
    ? plansOrInput
    : firstNonEmpty(
      plansOrInput.plans,
      plansOrInput.importPlans,
      plansOrInput.import_plans,
      plansOrInput.importPlanMatrix?.plans,
      plansOrInput.import_plan_matrix?.plans,
    );
  const explicitRows = asArray(explicit);
  if (explicitRows.length > 0) {
    if (explicitRows.every((row) => isImportPlan(row))) return explicitRows;
    if (explicitRows.every((row) => isExportPackage(row))) {
      return buildMeetingPlatformAdapterImportPlanMatrix(explicitRows, input, options).plans;
    }
  }
  const packages = Array.isArray(plansOrInput)
    ? []
    : asArray(firstNonEmpty(plansOrInput.packages, plansOrInput.exportPackages, plansOrInput.export_packages));
  if (packages.length > 0) return buildMeetingPlatformAdapterImportPlanMatrix(packages, input, options).plans;
  return [];
}

function contentScriptJs(input = {}, options = {}) {
  return unique(asArray(firstNonEmpty(
    options.contentScriptJs,
    options.content_script_js,
    input.contentScriptJs,
    input.content_script_js,
    DEFAULT_CONTENT_SCRIPT_JS,
  )));
}

function selectedBaseUrl(input = {}, options = {}) {
  return firstNonEmpty(options.baseUrl, options.base_url, input.baseUrl, input.base_url);
}

function selectedInstallTarget(input = {}, options = {}) {
  return String(firstNonEmpty(
    options.installTarget,
    options.install_target,
    input.installTarget,
    input.install_target,
    'host_project',
  ));
}

function sdkImports(plans = []) {
  const first = plans.find((plan) => plan.sdk_imports)?.sdk_imports ?? {};
  return {
    sdk_root: first.sdk_root ?? '@ai-annotation/meeting-timeline-sdk',
    platform_kit: first.platform_kit ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    adapter_import_plan: first.adapter_import_plan ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-import-plan',
    adapter_blueprint: first.adapter_blueprint ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint',
    adapter_install_manifest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-install-manifest',
    connector: first.connector ?? '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
    runtime_bundle: first.runtime_bundle ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
  };
}

function platformRow(plan = {}) {
  const entry = plan.surface_entrypoint ?? {};
  return compactObject({
    platform: plan.platform,
    display_name: plan.display_name,
    selected_surface: plan.selected_surface,
    ready: plan.accepted === true,
    timestamp_field: plan.runtime_contract?.timestamp_field ?? entry.timestamp_field,
    local_axis_first: plan.runtime_contract?.local_axis_first === true,
    provider_events_block_realtime: plan.runtime_contract?.provider_events_block_realtime === true,
    transcript_blocks_realtime: plan.runtime_contract?.transcript_blocks_realtime === true,
    first_sdk_method: entry.first_sdk_method,
    mark_insert_method: entry.mark_insert_method ?? 'insertAnnotation',
    speaker_track_method: entry.speaker_track_method,
    participant_track_method: entry.participant_track_method,
    browser_matches: entry.matches,
    provider_path: entry.path,
    provider_transport: entry.transport,
    provider_required_for_realtime: entry.required_for_realtime === true,
    adapter_blueprint_path: plan.adapter_blueprint?.path,
    adapter_blueprint_primary_surface: plan.adapter_blueprint?.primary_surface,
    adapter_blueprint_first_gate: plan.adapter_blueprint?.first_acceptance_gate,
    source_package: plan.source_package,
    first_issue: plan.readiness?.issues?.[0]?.code,
  });
}

function browserRegistry(plans = [], input = {}, options = {}) {
  const js = contentScriptJs(input, options);
  const rows = plans
    .map(platformRow)
    .filter((row) => row.selected_surface === 'browser_extension');
  const contentScripts = rows.map((row) => ({
    platform: row.platform,
    matches: row.browser_matches ?? [],
    js,
    run_at: 'document_idle',
    message_types: [...DEFAULT_MESSAGE_TYPES],
  }));
  return {
    enabled: rows.length > 0,
    platform_count: rows.length,
    host_permissions: unique(contentScripts.flatMap((script) => script.matches ?? [])),
    content_scripts: contentScripts,
    message_types: [...DEFAULT_MESSAGE_TYPES],
  };
}

function surfaceRegistry(plans = [], surface) {
  const rows = plans
    .map(platformRow)
    .filter((row) => row.selected_surface === surface);
  return {
    enabled: rows.length > 0,
    platform_count: rows.length,
    rows,
  };
}

function providerRegistry(plans = []) {
  const rows = plans.map((plan) => {
    const row = platformRow(plan);
    const provider = plan.surface_entrypoints?.provider_reconcile ?? {};
    return compactObject({
      platform: row.platform,
      provider_path: provider.path ?? row.provider_path,
      provider_transport: provider.transport ?? row.provider_transport,
      required_for_realtime: provider.required_for_realtime === true || row.provider_required_for_realtime === true,
      blocks_realtime_annotation: provider.blocks_realtime_annotation === true || provider.required_for_realtime === true || row.provider_required_for_realtime === true,
      first_sdk_method: provider.first_sdk_method ?? 'ingestProvider',
    });
  });
  return {
    enabled: rows.some((row) => row.provider_path),
    platform_count: rows.filter((row) => row.provider_path).length,
    realtime_blocking_count: rows.filter((row) => row.required_for_realtime === true || row.blocks_realtime_annotation === true).length,
    rows: rows.filter((row) => row.provider_path),
  };
}

function adapterBlueprintRegistry(plans = []) {
  const rows = plans.map((plan) => compactObject({
    platform: plan.platform,
    display_name: plan.display_name,
    available: plan.adapter_blueprint?.available === true,
    schema: plan.adapter_blueprint?.schema,
    path: plan.adapter_blueprint?.path,
    command: plan.adapter_blueprint?.command,
    primary_surface: plan.adapter_blueprint?.primary_surface,
    surface_order: plan.adapter_blueprint?.surface_order,
    timestamp_field: plan.adapter_blueprint?.timestamp_field,
    provider_blocks_realtime: plan.adapter_blueprint?.provider_blocks_realtime,
    transcript_blocks_realtime: plan.adapter_blueprint?.transcript_blocks_realtime,
    first_acceptance_gate: plan.adapter_blueprint?.first_acceptance_gate,
  }));
  return {
    enabled: rows.some((row) => row.available === true),
    platform_count: rows.filter((row) => row.available === true).length,
    rows,
  };
}

function duplicatePlatforms(plans = []) {
  const counts = new Map();
  for (const plan of plans) counts.set(plan.platform, (counts.get(plan.platform) ?? 0) + 1);
  return [...counts.entries()].filter(([, count]) => count > 1).map(([platform]) => platform);
}

function readiness(plans = [], input = {}, options = {}) {
  const allowPartialInstall = firstNonEmpty(
    options.allowPartialInstall,
    options.allow_partial_install,
    input.allowPartialInstall,
    input.allow_partial_install,
    false,
  ) === true;
  const duplicates = duplicatePlatforms(plans);
  const blockedPlans = plans.filter((plan) => plan.accepted !== true);
  const timestampIssues = plans.filter((plan) => plan.runtime_contract?.timestamp_field !== 'captured_at_ms');
  const localAxisIssues = plans.filter((plan) => plan.runtime_contract?.local_axis_first !== true);
  const providerBlocking = plans.filter((plan) => plan.runtime_contract?.provider_events_block_realtime === true);
  const transcriptBlocking = plans.filter((plan) => plan.runtime_contract?.transcript_blocks_realtime === true);
  const noPlans = plans.length === 0;
  const issues = [
    noPlans ? issue('error', 'missing_import_plans', 'Install manifest requires at least one adapter import plan.') : undefined,
    duplicates.length > 0 ? issue('error', 'duplicate_platforms', 'Install manifest contains duplicate platform plans.', {
      platforms: duplicates,
    }) : undefined,
    blockedPlans.length > 0 && !allowPartialInstall ? issue('error', 'blocked_import_plans', 'Every platform import plan must be accepted before installation.', {
      platforms: blockedPlans.map((plan) => plan.platform),
    }) : undefined,
    timestampIssues.length > 0 ? issue('error', 'invalid_timestamp_contract', 'Every installed adapter must use captured_at_ms.', {
      platforms: timestampIssues.map((plan) => plan.platform),
    }) : undefined,
    localAxisIssues.length > 0 ? issue('error', 'local_axis_first_required', 'Every installed adapter must bind a local axis before realtime annotation insertion.', {
      platforms: localAxisIssues.map((plan) => plan.platform),
    }) : undefined,
    providerBlocking.length > 0 ? issue('error', 'provider_blocks_realtime', 'Provider events must not block realtime annotations.', {
      platforms: providerBlocking.map((plan) => plan.platform),
    }) : undefined,
    transcriptBlocking.length > 0 ? issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotations.', {
      platforms: transcriptBlocking.map((plan) => plan.platform),
    }) : undefined,
  ].filter(Boolean);
  const readyPlans = plans.filter((plan) => plan.accepted === true);
  return {
    accepted: issues.filter((item) => item.severity === 'error').length === 0,
    allow_partial_install: allowPartialInstall,
    plan_count: plans.length,
    ready_plan_count: readyPlans.length,
    blocked_plan_count: blockedPlans.length,
    duplicate_platforms: duplicates,
    issue_count: issues.length,
    issues,
  };
}

function installSequence(plans = []) {
  const platforms = unique(plans.map((plan) => plan.platform));
  return [
    {
      step: 1,
      id: 'load_import_plans',
      action: 'read adapter-import-plan.json for every selected meeting platform',
      required: true,
      platform_count: platforms.length,
    },
    {
      step: 2,
      id: 'load_adapter_blueprints',
      action: 'read adapter-blueprint.json for surface order, evidence contract, and acceptance gates',
      required: false,
      platform_count: platforms.length,
    },
    {
      step: 3,
      id: 'create_sdk_facade',
      action: 'createMeetingAppTimelineSdk({ baseUrl, platforms })',
      required: true,
      sdk_method: 'createMeetingAppTimelineSdk',
      platforms,
    },
    {
      step: 4,
      id: 'register_selected_surfaces',
      action: 'install the selected local surface for each platform before any provider reconcile path',
      required: true,
      surfaces: unique(plans.map((plan) => plan.selected_surface)),
    },
    {
      step: 5,
      id: 'bind_axis_before_marks',
      action: 'start local candidate observation and bind the current meeting axis',
      required: true,
      sdk_method: 'observePlatformCandidates',
      timestamp_field: 'captured_at_ms',
    },
    {
      step: 6,
      id: 'insert_marks_in_realtime',
      action: 'insert every annotation with captured_at_ms from the local device/app clock',
      required: true,
      sdk_method: 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
    },
    {
      step: 7,
      id: 'reconcile_provider_events_later',
      action: 'connect provider events only for reconcile/backfill after the local axis exists',
      required: false,
      sdk_method: 'ingestProvider',
    },
  ];
}

function nextActions(manifest = {}, ready = {}) {
  return unique([
    ...(ready.issues ?? []).map((item) => item.code),
    manifest.accepted ? 'install_manifest_into_host_runtime' : 'fix_install_manifest_blockers',
    manifest.accepted ? 'register_platform_surface_registry' : '',
    manifest.accepted ? 'start_local_axis_observers_before_marks' : '',
  ]);
}

export function buildMeetingPlatformAdapterInstallManifest(plansOrInput = {}, input = {}, options = {}) {
  const plans = planListFrom(plansOrInput, input, options);
  const ready = readiness(plans, input, options);
  const rows = plans.map(platformRow);
  const target = String(firstNonEmpty(options.target, input.target, plans[0]?.target, 'static'));
  const manifest = {
    type: 'meeting_platform_adapter_install_manifest',
    schema: MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA_VERSION,
    target,
    install_target: selectedInstallTarget(input, options),
    base_url: selectedBaseUrl(input, options),
    accepted: ready.accepted,
    platform_count: plans.length,
    ready_platform_count: rows.filter((row) => row.ready === true).length,
    blocked_platform_count: rows.filter((row) => row.ready !== true).length,
    platforms: rows.map((row) => row.platform),
    selected_surfaces: unique(rows.map((row) => row.selected_surface)),
    sdk_imports: sdkImports(plans),
    runtime_contract: {
      timestamp_field: 'captured_at_ms',
      local_axis_first: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
    },
    platform_registry: rows,
    browser_extension: browserRegistry(plans, input, options),
    webview_preload: surfaceRegistry(plans, 'webview_preload'),
    native_detector: surfaceRegistry(plans, 'native_detector'),
    native_host: surfaceRegistry(plans, 'native_host'),
    adapter_blueprints: adapterBlueprintRegistry(plans),
    provider_reconcile: providerRegistry(plans),
    install_sequence: installSequence(plans),
    readiness: ready,
  };
  return {
    ...manifest,
    next_actions: nextActions(manifest, ready),
  };
}

export function assertMeetingPlatformAdapterInstallManifest(manifestOrInput = {}, input = {}, options = {}) {
  const manifest = manifestOrInput.schema === MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA
    ? manifestOrInput
    : buildMeetingPlatformAdapterInstallManifest(manifestOrInput, input, options);
  if (manifest.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter install manifest is not ready', {
      code: 'meeting_platform_adapter_install_manifest_not_ready',
      issues: manifest.readiness?.issues ?? [],
      blocked_platforms: manifest.platform_registry?.filter((row) => row.ready !== true).map((row) => row.platform) ?? [],
    });
  }
  return manifest;
}
