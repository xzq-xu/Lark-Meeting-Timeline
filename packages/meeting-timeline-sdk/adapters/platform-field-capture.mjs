import { compactObject } from '../index.mjs';
import {
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
  buildMeetingAppExtensionMatchPatterns,
} from './meeting-app-extension.mjs';
import { MEETING_APP_FIXTURE_PLATFORMS } from './meeting-app-fixtures.mjs';
import { buildMeetingAppLiveSnapshotCapturePlan } from './meeting-app-profile.mjs';
import {
  buildMeetingPlatformEvidencePackage,
  buildMeetingPlatformEvidencePackageSummary,
  verifyMeetingPlatformEvidencePackage,
} from './platform-evidence-package.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { buildMeetingPlatformRuntimeProfile } from './platform-runtime-profile.mjs';

export const MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA = 'meeting_platform_field_capture_plan';
export const MEETING_PLATFORM_FIELD_CAPTURE_MATRIX_SCHEMA = 'meeting_platform_field_capture_matrix';
export const MEETING_PLATFORM_FIELD_CAPTURE_MANIFEST_SCHEMA = 'meeting_platform_field_capture_manifest';
export const MEETING_PLATFORM_FIELD_CAPTURE_MANIFEST_MATRIX_SCHEMA = 'meeting_platform_field_capture_manifest_matrix';
export const MEETING_PLATFORM_FIELD_COLLECTOR_CONFIG_SCHEMA = 'meeting_platform_field_collector_config';
export const MEETING_PLATFORM_FIELD_COLLECTOR_CONFIG_MATRIX_SCHEMA = 'meeting_platform_field_collector_config_matrix';
export const MEETING_PLATFORM_FIELD_EVIDENCE_BUNDLE_SCHEMA = 'meeting_platform_field_evidence_bundle';
export const MEETING_PLATFORM_FIELD_EVIDENCE_MATRIX_SCHEMA = 'meeting_platform_field_evidence_matrix';
export const MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION = 1;

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

function matchesPlatform(value, platform) {
  if (!value) return true;
  try {
    return normalizeMeetingPlatform(value) === platform;
  } catch {
    return false;
  }
}

function selectedEvidencePackage(platform, options = {}) {
  const input = firstNonEmpty(
    options.evidencePackage,
    options.evidence_package,
    options.package,
    options.platformEvidencePackage,
    options.platform_evidence_package,
  );
  if (!input) return undefined;
  if (Array.isArray(input)) {
    return input.find((item) => matchesPlatform(item?.platform ?? item?.rollout_plan?.platform, platform));
  }
  if (input.schema === 'meeting_platform_evidence_package' || input.rollout_plan || input.evidence_summary) {
    return matchesPlatform(input.platform ?? input.rollout_plan?.platform, platform) ? input : undefined;
  }
  if (typeof input === 'object') {
    for (const [key, value] of Object.entries(input)) {
      if (matchesPlatform(key, platform)) return value;
    }
  }
  return undefined;
}

function selectedFieldEvidenceInput(platform, options = {}) {
  const input = firstNonEmpty(
    options.fieldEvidence,
    options.field_evidence,
    options.fieldCaptureEvidence,
    options.field_capture_evidence,
    options.evidenceInput,
    options.evidence_input,
    options.inputs,
    options.input,
    options.evidence,
  );
  if (!input) return {};
  if (Array.isArray(input)) {
    return input.find((item) => matchesPlatform(item?.platform ?? item?.rollout_plan?.platform, platform)) ?? {};
  }
  if (input.schema === 'meeting_platform_evidence_package' || input.providerRecords || input.provider_records || input.meetingAppRecordSet || input.meeting_app_record_set) {
    return matchesPlatform(input.platform ?? input.rollout_plan?.platform, platform) ? input : {};
  }
  if (typeof input === 'object') {
    for (const [key, value] of Object.entries(input)) {
      if (matchesPlatform(key, platform)) return value ?? {};
    }
  }
  return {};
}

function evidencePackageFromFieldInput(platform, input = {}, options = {}) {
  if (input?.schema === 'meeting_platform_evidence_package') return input;
  const nestedPackage = firstNonEmpty(
    input?.evidencePackage,
    input?.evidence_package,
    input?.platformEvidencePackage,
    input?.platform_evidence_package,
    options.evidencePackage,
    options.evidence_package,
    options.platformEvidencePackage,
    options.platform_evidence_package,
  );
  const selected = selectedEvidencePackage(platform, { evidencePackage: nestedPackage });
  if (selected) return selected;
  return buildMeetingPlatformEvidencePackage(platform, input, {
    ...options,
    includeRunbook: firstNonEmpty(
      options.includeRunbook,
      options.include_runbook,
      input?.includeRunbook,
      input?.include_runbook,
      false,
    ),
    source: firstNonEmpty(options.source, input?.source, 'meeting_platform_field_evidence_bundle'),
  });
}

function outputPaths(platform, options = {}) {
  const root = String(firstNonEmpty(options.evidenceDir, options.evidence_dir, 'data'));
  return {
    meeting_app_evidence: `${root}/meeting-app-evidence/${platform}.json`,
    provider_evidence: platform === 'local_detector'
      ? `${root}/local-detector-evidence/${platform}.json`
      : `${root}/provider-evidence/${platform}.json`,
    evidence_package: `${root}/meeting-platform-evidence-packages/${platform}.json`,
    capture_report: `${root}/meeting-platform-field-capture/${platform}.json`,
  };
}

function fieldEvidencePaths(platform, options = {}) {
  const root = String(firstNonEmpty(options.evidenceDir, options.evidence_dir, 'data'));
  return {
    field_evidence_input_dir: `${root}/meeting-platform-field-evidence`,
    field_evidence_input: `${root}/meeting-platform-field-evidence/${platform}.json`,
    field_evidence_bundle_dir: `${root}/meeting-platform-field-evidence-bundles`,
    field_evidence_bundle: `${root}/meeting-platform-field-evidence-bundles/${platform}.json`,
    evidence_package_dir: `${root}/meeting-platform-evidence-packages`,
    evidence_package: `${root}/meeting-platform-evidence-packages/${platform}.json`,
    field_evidence_report: `${root}/meeting-platform-field-evidence-report.json`,
    field_capture_report: `${root}/meeting-platform-field-capture-report.json`,
  };
}

function requiredProviderCoverage(platform, runtime = {}) {
  if (platform === 'local_detector') return ['meeting_start', 'meeting_end'];
  return [
    runtime.axis?.start?.provider_reconcile_events?.length > 0 ? 'meeting_start' : undefined,
    runtime.axis?.end?.provider_reconcile_events?.length > 0 ? 'meeting_end' : undefined,
  ].filter(Boolean);
}

function providerEventPlan(platform, runtime = {}, provider = {}) {
  const coverage = requiredProviderCoverage(platform, runtime);
  return {
    source: platform === 'local_detector' ? 'trusted_host_detector' : provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    required_coverage: coverage,
    minimum_record_count: coverage.length,
    start_events: runtime.axis?.start?.provider_reconcile_events ?? [],
    end_events: runtime.axis?.end?.provider_reconcile_events ?? [],
    participant_events: runtime.provider_events?.participant_events ?? [],
    artifact_events: runtime.provider_events?.artifact_events ?? [],
    lifecycle_events: runtime.provider_events?.lifecycle_events ?? [],
    security: {
      verifier: provider.security?.verifier,
      missing_env: provider.security?.missing_env ?? [],
      required_env: provider.security?.required_env ?? [],
    },
  };
}

function localSnapshotPlan(platform, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) return undefined;
  const plan = buildMeetingAppLiveSnapshotCapturePlan(platform, options);
  return {
    source: 'browser_extension_or_native_host_observer',
    required_snapshots: plan.required_snapshots,
    recommended_snapshots: plan.recommended_snapshots,
    minimum_record_count: plan.minimum_record_count,
    validation: plan.validation,
    runtime_config: plan.runtime_config,
    recorder: plan.recorder,
  };
}

function summaryForPackage(platform, options = {}) {
  const pkg = selectedEvidencePackage(platform, options);
  return pkg ? buildMeetingPlatformEvidencePackageSummary(pkg, options) : undefined;
}

function missingLocalItems(localPlan = undefined, summary = undefined) {
  if (!localPlan) return [];
  if (!summary) {
    return [
      'capture_live_dom_snapshots_for_local_observer',
      ...localPlan.required_snapshots.map((item) => `capture_required_snapshot:${item.id}`),
    ];
  }
  const missing = [];
  if (Number(summary.meeting_app_record_count ?? 0) < Number(localPlan.minimum_record_count ?? 0)) {
    missing.push('capture_live_dom_snapshots_for_local_observer');
  }
  for (const item of summary.local_dom_missing_required_coverage ?? []) {
    missing.push(`local_dom_missing:${item}`);
  }
  return unique(missing);
}

function missingProviderItems(providerPlan = {}, summary = undefined) {
  if (!summary) {
    return [
      'capture_real_provider_start_end_events',
      ...providerPlan.required_coverage.map((item) => `provider_missing:${item}`),
    ];
  }
  const missing = [];
  if (Number(summary.provider_record_count ?? 0) < Number(providerPlan.minimum_record_count ?? 0)) {
    missing.push('capture_real_provider_start_end_events');
  }
  for (const item of summary.provider_missing_required_coverage ?? []) {
    missing.push(`provider_missing:${item}`);
  }
  return unique(missing);
}

function fieldStatus(summary = undefined, localMissing = [], providerMissing = []) {
  if (summary?.production_ready === true) return 'production_ready';
  if (summary?.ready_for_realtime_annotations === true && providerMissing.length > 0) return 'pilot_ready_provider_pending';
  if (localMissing.length === 0 && providerMissing.length > 0) return 'needs_provider_events';
  if (localMissing.length > 0 && providerMissing.length === 0) return 'needs_local_observer_snapshots';
  return 'needs_local_and_provider_evidence';
}

function captureChecklist(platform, localPlan = undefined, providerPlan = {}, runtime = {}) {
  const items = [];
  if (localPlan) {
    items.push({
      id: 'capture_active_speaker_snapshot',
      channel: 'meeting_app_observer',
      when: 'after_joining_the_real_meeting_and_a_speaker_is_visibly_active',
      output: 'meeting_app_snapshot_record',
      required_snapshot: 'active_speaker',
    });
    if ((localPlan.required_snapshots ?? []).some((item) => item.id === 'meeting_ended')) {
      items.push({
        id: 'capture_meeting_ended_snapshot',
        channel: 'meeting_app_observer',
        when: 'after_leaving_or_ending_the_same_real_meeting',
        output: 'meeting_app_snapshot_record',
        required_snapshot: 'meeting_ended',
      });
    }
    items.push({
      id: 'insert_live_annotation_sample',
      channel: 'timeline_sdk',
      when: 'while_the_local_axis_is_active',
      output: 'annotation_mark_with_captured_at_ms',
      invariant: runtime.runtime_contract?.annotation_timestamp_field,
    });
  }
  items.push({
    id: platform === 'local_detector' ? 'capture_detector_start_end' : 'capture_provider_start_end',
    channel: platform === 'local_detector' ? 'trusted_host_detector' : 'provider_webhook',
    when: platform === 'local_detector'
      ? 'when_the_host_detector_reports_meeting_started_and_meeting_ended'
      : 'when_the_provider_delivers_real_start_and_end_events_for_the_same_meeting',
    output: 'platform_capture_records',
    required_coverage: providerPlan.required_coverage,
  });
  return items;
}

function acceptedInputContracts() {
  return [
    {
      schema: 'raw_field_evidence',
      description: 'Unwrapped platform capture input from field tools.',
      required_fields: ['platform'],
      optional_fields: ['providerRecords', 'providerSamples', 'meetingAppRecordSet', 'meetingAppRecords', 'metadata'],
    },
    {
      schema: 'platform_map',
      description: 'Object keyed by platform, where each value is raw_field_evidence.',
      example_keys: ['google_meet', 'teams', 'zoom', 'webex', 'lark'],
    },
    {
      schema: 'meeting_platform_evidence_package',
      description: 'Previously exported package from buildMeetingPlatformEvidencePackage.',
    },
    {
      schema: 'meeting_platform_field_evidence_bundle',
      description: 'Previously exported bundle from buildMeetingPlatformFieldEvidenceBundle.',
    },
    {
      schema: 'meeting_app_snapshot_record_set',
      description: 'Local DOM snapshot record set for browser-extension or native-host observers.',
    },
  ];
}

function manifestCommands(platform, paths = {}, options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const baseUrlArg = baseUrl ? ` --base-url=${baseUrl}` : '';
  return {
    capture_manifest: `buildMeetingPlatformFieldCaptureManifest('${platform}', options)`,
    build_field_evidence: `npm run meeting-platform:field-evidence -- --platforms=${platform}${baseUrlArg} --dir=${paths.field_evidence_input_dir} --package-dir=${paths.evidence_package_dir} --bundle-dir=${paths.field_evidence_bundle_dir} --report-file=${paths.field_evidence_report}`,
    inspect_capture_gaps: `npm run meeting-platform:field-capture -- --platforms=${platform}${baseUrlArg} --dir=${paths.evidence_package_dir} --report-file=${paths.field_capture_report}`,
    verify_package: `npm run meeting-platform:evidence-package -- --input=${paths.evidence_package}${baseUrlArg}`,
  };
}

function endpointMap(options = {}) {
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const entries = Object.entries(MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS).map(([key, path]) => [
    key,
    baseUrl ? new URL(path, String(baseUrl)).toString() : path,
  ]);
  return Object.fromEntries(entries);
}

function extensionCollector(platform, options = {}) {
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(platform)) return undefined;
  const patterns = buildMeetingAppExtensionMatchPatterns([platform], {
    js: firstNonEmpty(options.collectorScript, options.collector_script, options.contentScriptJs, options.content_script_js, 'meeting-app-field-collector.bundle.js'),
    runAt: firstNonEmpty(options.runAt, options.run_at, 'document_idle'),
    allFrames: firstNonEmpty(options.allFrames, options.all_frames, false),
    extraMatches: options.extraMatches ?? options.extra_matches,
    extraHostPermissions: options.extraHostPermissions ?? options.extra_host_permissions,
  });
  return {
    type: 'browser_extension_or_webview_preload',
    platforms: patterns.platforms,
    matches: patterns.matches,
    host_permissions: patterns.host_permissions,
    content_scripts: patterns.content_scripts,
    timeline_endpoints: endpointMap(options),
  };
}

export function buildMeetingPlatformFieldCapturePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const localPlan = localSnapshotPlan(key, options);
  const providerPlan = providerEventPlan(key, runtime, provider);
  const evidenceSummary = summaryForPackage(key, options);
  const localMissing = missingLocalItems(localPlan, evidenceSummary);
  const providerMissing = missingProviderItems(providerPlan, evidenceSummary);
  const status = fieldStatus(evidenceSummary, localMissing, providerMissing);
  return compactObject({
    type: 'meeting_platform_field_capture_plan',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform: key,
    display_name: runtime.display_name,
    status,
    production_ready: status === 'production_ready',
    ready_for_realtime_annotations: evidenceSummary?.ready_for_realtime_annotations === true,
    objective: 'capture_real_meeting_evidence_for_timeline_annotation_rollout',
    output_paths: outputPaths(key, options),
    runtime_contract: runtime.runtime_contract,
    local_observer: localPlan,
    provider_events: providerPlan,
    checklist: captureChecklist(key, localPlan, providerPlan, runtime),
    current_evidence: evidenceSummary,
    missing_items: unique([
      ...localMissing,
      ...providerMissing,
    ]),
    validation: {
      build_package: 'buildMeetingPlatformEvidencePackage(platform, { providerRecords, meetingAppRecordSet, env, baseUrl })',
      verify_package: 'verifyMeetingPlatformEvidencePackage(evidencePackage)',
      production_condition: 'evidencePackage.rollout_plan.production_ready === true',
      pilot_condition: 'evidencePackage.rollout_plan.ready_for_realtime_annotations === true',
    },
    next_actions: status === 'production_ready'
      ? ['ship_with_monitoring_and_keep_collecting_regression_evidence']
      : unique([
        ...localMissing,
        ...providerMissing,
        'export_meeting_platform_evidence_package',
        'verify_meeting_platform_evidence_package',
      ]),
  });
}

export function buildMeetingPlatformFieldCaptureMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldCapturePlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_field_capture_matrix',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    realtime_ready_count: plans.filter((plan) => plan.ready_for_realtime_annotations).length,
    missing_item_count: plans.reduce((total, plan) => total + (plan.missing_items?.length ?? 0), 0),
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      production_ready: plan.production_ready,
      ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
      local_required_records: plan.local_observer?.minimum_record_count ?? 0,
      provider_required_records: plan.provider_events?.minimum_record_count ?? 0,
      missing_items: plan.missing_items,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformFieldCaptureManifest(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const plan = buildMeetingPlatformFieldCapturePlan(key, options);
  const paths = fieldEvidencePaths(key, options);
  const requiredSnapshots = (plan.local_observer?.required_snapshots ?? []).map((item) => item.id);
  return compactObject({
    type: 'meeting_platform_field_capture_manifest',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform: key,
    display_name: plan.display_name,
    status: plan.status,
    production_ready: plan.production_ready,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    objective: 'portable_field_capture_contract_for_meeting_timeline_platform_adaptation',
    file_contract: {
      input_dir: paths.field_evidence_input_dir,
      bundle_dir: paths.field_evidence_bundle_dir,
      package_dir: paths.evidence_package_dir,
      files: {
        field_evidence_input: paths.field_evidence_input,
        field_evidence_bundle: paths.field_evidence_bundle,
        evidence_package: paths.evidence_package,
        field_evidence_report: paths.field_evidence_report,
        field_capture_report: paths.field_capture_report,
      },
      write_policy: {
        field_evidence_input: 'one_json_per_platform_or_platform_map',
        field_evidence_bundle: 'generated_by_meeting-platform_field-evidence',
        evidence_package: 'generated_by_meeting-platform_field-evidence_or_buildMeetingPlatformEvidencePackage',
      },
    },
    input_contract: {
      accepted_inputs: acceptedInputContracts(),
      raw_field_evidence_shape: {
        platform: key,
        providerRecords: 'array_of_platform_capture_records_from_provider_webhooks',
        providerSamples: 'array_or_platform_map_of_raw_provider_sample_events',
        meetingAppRecordSet: 'meeting_app_snapshot_record_set_from_browser_extension_or_native_host',
        meetingAppRecords: 'array_of_meeting_app_snapshot_records',
        metadata: 'source_files_or_capture_environment_metadata',
      },
    },
    runtime_contract: {
      annotation_timestamp_field: runtime.runtime_contract?.annotation_timestamp_field,
      provider_transcript_blocking: runtime.runtime_contract?.provider_transcript_blocking,
      local_axis_policy: runtime.axis,
      speaker_marker_filter: runtime.speaker_markers?.filter,
    },
    local_observer_contract: plan.local_observer ? {
      source: plan.local_observer.source,
      required_snapshots: requiredSnapshots,
      minimum_record_count: plan.local_observer.minimum_record_count,
      recorder: plan.local_observer.recorder,
      validation: plan.local_observer.validation,
    } : undefined,
    provider_contract: {
      transport: provider.transport,
      endpoint: provider.endpoint,
      status_endpoint: provider.status_endpoint,
      required_coverage: plan.provider_events?.required_coverage ?? [],
      minimum_record_count: plan.provider_events?.minimum_record_count ?? 0,
      start_events: plan.provider_events?.start_events ?? [],
      end_events: plan.provider_events?.end_events ?? [],
      participant_events: plan.provider_events?.participant_events ?? [],
      artifact_events: plan.provider_events?.artifact_events ?? [],
      lifecycle_events: plan.provider_events?.lifecycle_events ?? [],
      security: plan.provider_events?.security,
    },
    acceptance: {
      production_condition: 'bundle.production_ready === true && bundle.verification.passed === true',
      pilot_condition: 'bundle.ready_for_realtime_annotations === true',
      required_provider_coverage: plan.provider_events?.required_coverage ?? [],
      required_local_snapshots: requiredSnapshots,
      minimum_provider_records: plan.provider_events?.minimum_record_count ?? 0,
      minimum_local_records: plan.local_observer?.minimum_record_count ?? 0,
      current_missing_items: plan.missing_items ?? [],
    },
    automation: {
      commands: manifestCommands(key, paths, options),
      sdk_methods: [
        'buildMeetingPlatformFieldCaptureManifest',
        'buildMeetingPlatformFieldEvidenceBundle',
        'buildMeetingPlatformFieldCapturePlan',
        'verifyMeetingPlatformEvidencePackage',
      ],
    },
    handoff: {
      primary_entry: '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture',
      kit_methods: ['platformFieldCaptureManifest', 'platformFieldEvidenceBundle', 'platformFieldCapturePlan'],
      report_fields: ['status', 'production_ready', 'ready_for_realtime_annotations', 'missing_items'],
    },
    field_capture_plan: plan,
    next_actions: plan.next_actions ?? [],
  });
}

export function buildMeetingPlatformFieldCaptureManifestMatrix(options = {}) {
  const manifests = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldCaptureManifest(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_field_capture_manifest_matrix',
    schema: MEETING_PLATFORM_FIELD_CAPTURE_MANIFEST_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform_count: manifests.length,
    production_ready_count: manifests.filter((manifest) => manifest.production_ready).length,
    realtime_ready_count: manifests.filter((manifest) => manifest.ready_for_realtime_annotations).length,
    missing_item_count: manifests.reduce((total, manifest) => total + (manifest.acceptance?.current_missing_items?.length ?? 0), 0),
    platforms: manifests.map((manifest) => manifest.platform),
    rows: manifests.map((manifest) => ({
      platform: manifest.platform,
      display_name: manifest.display_name,
      status: manifest.status,
      production_ready: manifest.production_ready,
      ready_for_realtime_annotations: manifest.ready_for_realtime_annotations,
      input_file: manifest.file_contract?.files?.field_evidence_input,
      bundle_file: manifest.file_contract?.files?.field_evidence_bundle,
      evidence_package_file: manifest.file_contract?.files?.evidence_package,
      required_provider_coverage: manifest.acceptance?.required_provider_coverage ?? [],
      required_local_snapshots: manifest.acceptance?.required_local_snapshots ?? [],
      missing_items: manifest.acceptance?.current_missing_items ?? [],
    })),
    manifests,
    next_actions: unique(manifests.flatMap((manifest) => manifest.next_actions ?? [])),
  };
}

export function buildMeetingPlatformFieldCollectorConfig(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const manifest = buildMeetingPlatformFieldCaptureManifest(key, options);
  const extension = extensionCollector(key, options);
  return compactObject({
    type: 'meeting_platform_field_collector_config',
    schema: MEETING_PLATFORM_FIELD_COLLECTOR_CONFIG_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform: key,
    display_name: manifest.display_name,
    mode: extension ? 'hybrid_browser_observer_and_provider_events' : 'provider_or_host_detector_only',
    objective: 'runtime_configuration_for_real_meeting_field_capture_tools',
    browser_observer: extension,
    provider_observer: {
      transport: manifest.provider_contract?.transport,
      endpoint: manifest.provider_contract?.endpoint,
      status_endpoint: manifest.provider_contract?.status_endpoint,
      required_coverage: manifest.provider_contract?.required_coverage ?? [],
      minimum_record_count: manifest.provider_contract?.minimum_record_count ?? 0,
      start_events: manifest.provider_contract?.start_events ?? [],
      end_events: manifest.provider_contract?.end_events ?? [],
      participant_events: manifest.provider_contract?.participant_events ?? [],
      artifact_events: manifest.provider_contract?.artifact_events ?? [],
      lifecycle_events: manifest.provider_contract?.lifecycle_events ?? [],
      security: manifest.provider_contract?.security,
    },
    local_snapshot_collector: manifest.local_observer_contract ? {
      required_snapshots: manifest.local_observer_contract.required_snapshots ?? [],
      minimum_record_count: manifest.local_observer_contract.minimum_record_count ?? 0,
      recorder: manifest.local_observer_contract.recorder,
      validation: manifest.local_observer_contract.validation,
    } : undefined,
    timeline_ingest: {
      endpoints: endpointMap(options),
      annotation_timestamp_field: manifest.runtime_contract?.annotation_timestamp_field,
      realtime_annotation_policy: {
        use_captured_at_ms: true,
        do_not_wait_for_transcript: true,
        provider_events_are_reconciliation_only: true,
      },
    },
    storage: {
      input_contract: manifest.input_contract,
      files: manifest.file_contract?.files,
      write_policy: manifest.file_contract?.write_policy,
    },
    acceptance: manifest.acceptance,
    automation: manifest.automation,
    manifest,
    next_actions: manifest.next_actions ?? [],
  });
}

export function buildMeetingPlatformFieldCollectorConfigMatrix(options = {}) {
  const configs = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldCollectorConfig(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_field_collector_config_matrix',
    schema: MEETING_PLATFORM_FIELD_COLLECTOR_CONFIG_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform_count: configs.length,
    browser_observer_count: configs.filter((config) => config.browser_observer).length,
    production_ready_count: configs.filter((config) => config.acceptance?.current_missing_items?.length === 0 && config.manifest?.production_ready).length,
    realtime_ready_count: configs.filter((config) => config.manifest?.ready_for_realtime_annotations).length,
    platforms: configs.map((config) => config.platform),
    rows: configs.map((config) => ({
      platform: config.platform,
      display_name: config.display_name,
      mode: config.mode,
      browser_matches: config.browser_observer?.matches ?? [],
      provider_required_coverage: config.provider_observer?.required_coverage ?? [],
      local_required_snapshots: config.local_snapshot_collector?.required_snapshots ?? [],
      input_file: config.storage?.files?.field_evidence_input,
      bundle_file: config.storage?.files?.field_evidence_bundle,
      evidence_package_file: config.storage?.files?.evidence_package,
      missing_items: config.acceptance?.current_missing_items ?? [],
    })),
    configs,
    next_actions: unique(configs.flatMap((config) => config.next_actions ?? [])),
  };
}

export function buildMeetingPlatformFieldEvidenceBundle(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const evidencePackage = evidencePackageFromFieldInput(key, input, options);
  const fieldCapturePlan = buildMeetingPlatformFieldCapturePlan(key, {
    ...options,
    evidencePackage,
  });
  const verification = verifyMeetingPlatformEvidencePackage(evidencePackage, {
    ...options,
    requireProductionReady: firstNonEmpty(options.requireProductionReady, options.require_production_ready, true),
    requireCorrelation: firstNonEmpty(options.requireCorrelation, options.require_correlation, true),
  });
  const summary = buildMeetingPlatformEvidencePackageSummary(evidencePackage, options);
  return compactObject({
    type: 'meeting_platform_field_evidence_bundle',
    schema: MEETING_PLATFORM_FIELD_EVIDENCE_BUNDLE_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform: key,
    status: fieldCapturePlan.status,
    production_ready: fieldCapturePlan.production_ready,
    ready_for_realtime_annotations: fieldCapturePlan.ready_for_realtime_annotations,
    package_id: evidencePackage.id,
    evidence_package: evidencePackage,
    evidence_summary: summary,
    verification,
    field_capture_plan: fieldCapturePlan,
    handoff: {
      package_schema: 'meeting_platform_evidence_package',
      field_capture_schema: MEETING_PLATFORM_FIELD_CAPTURE_PLAN_SCHEMA,
      build_method: 'buildMeetingPlatformFieldEvidenceBundle(platform, input, options)',
      package_method: 'buildMeetingPlatformEvidencePackage(platform, input, options)',
      verify_method: 'verifyMeetingPlatformEvidencePackage(evidencePackage)',
      production_condition: 'field_capture_plan.production_ready === true && verification.passed === true',
      pilot_condition: 'field_capture_plan.ready_for_realtime_annotations === true',
    },
    next_actions: fieldCapturePlan.production_ready && verification.passed
      ? ['ship_with_monitoring_and_keep_collecting_regression_evidence']
      : unique([
        ...(fieldCapturePlan.next_actions ?? []),
        ...(verification.next_actions ?? []),
      ]),
  });
}

export function buildMeetingPlatformFieldEvidenceMatrix(options = {}) {
  const bundles = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldEvidenceBundle(
    platform,
    selectedFieldEvidenceInput(platform, options),
    {
      ...options,
      platforms: undefined,
      platform_keys: undefined,
      input: undefined,
      inputs: undefined,
      evidence: undefined,
      evidenceInput: undefined,
      evidence_input: undefined,
      fieldEvidence: undefined,
      field_evidence: undefined,
      fieldCaptureEvidence: undefined,
      field_capture_evidence: undefined,
    },
  ));
  return {
    type: 'meeting_platform_field_evidence_matrix',
    schema: MEETING_PLATFORM_FIELD_EVIDENCE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_CAPTURE_SCHEMA_VERSION,
    platform_count: bundles.length,
    production_ready_count: bundles.filter((bundle) => bundle.production_ready).length,
    realtime_ready_count: bundles.filter((bundle) => bundle.ready_for_realtime_annotations).length,
    verified_count: bundles.filter((bundle) => bundle.verification?.passed).length,
    missing_item_count: bundles.reduce((total, bundle) => total + (bundle.field_capture_plan?.missing_items?.length ?? 0), 0),
    platforms: bundles.map((bundle) => bundle.platform),
    rows: bundles.map((bundle) => ({
      platform: bundle.platform,
      status: bundle.status,
      production_ready: bundle.production_ready,
      ready_for_realtime_annotations: bundle.ready_for_realtime_annotations,
      verified: bundle.verification?.passed,
      package_id: bundle.package_id,
      provider_record_count: bundle.evidence_summary?.provider_record_count ?? 0,
      meeting_app_record_count: bundle.evidence_summary?.meeting_app_record_count ?? 0,
      missing_items: bundle.field_capture_plan?.missing_items ?? [],
    })),
    bundles,
    next_actions: unique(bundles.flatMap((bundle) => bundle.next_actions ?? [])),
  };
}
