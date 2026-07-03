import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformFieldCaptureManifest,
  buildMeetingPlatformFieldCapturePlan,
  buildMeetingPlatformFieldCollectorConfig,
} from './platform-field-capture.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRealEvidenceIntakePlan,
} from './platform-real-intake.mjs';
import {
  normalizeMeetingPlatform,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_FIELD_INTAKE_PLAN_SCHEMA = 'meeting_platform_field_intake_plan';
export const MEETING_PLATFORM_FIELD_INTAKE_MATRIX_SCHEMA = 'meeting_platform_field_intake_matrix';
export const MEETING_PLATFORM_FIELD_INTAKE_SCHEMA_VERSION = 1;

const DEFAULT_FIELD_INTAKE_PLATFORMS = Object.freeze([
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

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    DEFAULT_FIELD_INTAKE_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function evidenceDirs(options = {}) {
  const root = String(firstNonEmpty(options.evidenceDir, options.evidence_dir, 'data'));
  return {
    root,
    provider_evidence: `${root}/provider-evidence`,
    meeting_app_evidence: `${root}/meeting-app-evidence`,
    field_evidence_input: `${root}/meeting-platform-field-evidence`,
    field_evidence_bundles: `${root}/meeting-platform-field-evidence-bundles`,
    evidence_packages: `${root}/meeting-platform-evidence-packages`,
    field_manifests: `${root}/meeting-platform-field-manifests`,
    field_intake_plans: `${root}/meeting-platform-field-intake-plans`,
  };
}

function commandArgs(platform, options = {}) {
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787'));
  const dirs = evidenceDirs(options);
  return {
    platform,
    baseUrl,
    dirs,
    baseUrlArg: `--base-url=${baseUrl}`,
  };
}

function planCommands(platform, manifest = {}, options = {}) {
  const { baseUrlArg, dirs } = commandArgs(platform, options);
  const reportRoot = dirs.root;
  return {
    export_field_intake_plan: `npm run meeting-platform:field-intake -- --platforms=${platform} ${baseUrlArg} --evidence-dir=${dirs.root}`,
    export_field_manifest: `npm run meeting-platform:field-manifest -- --platforms=${platform} ${baseUrlArg} --evidence-dir=${dirs.root} --out-dir=${dirs.field_manifests}`,
    build_field_evidence: manifest.automation?.commands?.build_field_evidence
      ?? `npm run meeting-platform:field-evidence -- --platforms=${platform} ${baseUrlArg} --dir=${dirs.field_evidence_input} --package-dir=${dirs.evidence_packages} --bundle-dir=${dirs.field_evidence_bundles}`,
    inspect_capture_gaps: manifest.automation?.commands?.inspect_capture_gaps
      ?? `npm run meeting-platform:field-capture -- --platforms=${platform} ${baseUrlArg} --dir=${dirs.evidence_packages}`,
    validate_real_intake: `npm run meeting-platform:real-intake -- --platforms=${platform} ${baseUrlArg} --provider-dir=${dirs.provider_evidence} --dom-dir=${dirs.meeting_app_evidence} --package-dir=${dirs.evidence_packages} --fail-on-incomplete=true --report-file=${reportRoot}/meeting-platform-real-intake-report.json`,
    validate_live_readiness: `npm run meeting-platform:live-readiness -- --dir=${dirs.evidence_packages} --report-file=${reportRoot}/meeting-platform-live-readiness-report.json`,
  };
}

function statusFor(provider = {}, field = {}) {
  if (field.production_ready === true) return 'production_ready';
  const missingEnv = provider.security?.missing_env ?? [];
  if (missingEnv.length > 0) return 'provider_setup_missing_env';
  if (provider.readiness?.ready === false) return 'provider_setup_not_ready';
  const missing = field.missing_items ?? [];
  if (missing.some((item) => String(item).startsWith('capture_required_snapshot') || item === 'capture_live_dom_snapshots_for_local_observer')) {
    return 'needs_local_observer_capture';
  }
  if (missing.some((item) => String(item).startsWith('provider_missing') || item === 'capture_real_provider_start_end_events')) {
    return 'needs_provider_event_capture';
  }
  if (missing.length > 0) return 'needs_field_evidence';
  return 'ready_to_collect_field_evidence';
}

function operatorSteps(platform, provider = {}, manifest = {}, realIntake = {}, commands = {}) {
  const requiredSnapshots = manifest.acceptance?.required_local_snapshots ?? [];
  const requiredProviderCoverage = manifest.acceptance?.required_provider_coverage ?? [];
  return [
    {
      id: 'verify_provider_connection',
      owner: 'server',
      objective: 'provider_event_endpoint_can_receive_authenticated_events',
      command: commands.export_field_intake_plan,
      required_env: provider.security?.required_env ?? [],
      missing_env: provider.security?.missing_env ?? [],
    },
    {
      id: 'install_local_observer',
      owner: 'host_app_or_browser_extension',
      objective: 'capture_low_latency_meeting_app_state_for_the_current_axis',
      browser_matches: manifest.field_capture_plan?.local_observer?.runtime_config?.matches
        ?? manifest.local_observer_contract?.recorder?.matches
        ?? [],
      required_snapshots: requiredSnapshots,
    },
    {
      id: 'capture_active_meeting_window',
      owner: 'field_operator',
      objective: 'record_active_meeting_dom_or_native_snapshot_before_provider_reconcile',
      output: manifest.file_contract?.files?.field_evidence_input,
      required_snapshot: requiredSnapshots.includes('active_speaker') ? 'active_speaker' : requiredSnapshots[0],
    },
    {
      id: 'insert_live_annotation_sample',
      owner: 'host_app',
      objective: 'prove_a_human_mark_can_be_inserted_on_the_current_axis_without_waiting_for_transcript',
      timestamp_field: manifest.runtime_contract?.annotation_timestamp_field,
      invariant: 'captured_at_ms_must_be_written_at_capture_time',
    },
    {
      id: 'capture_provider_start_end_events',
      owner: 'provider_webhook',
      objective: 'capture_real_start_and_end_events_for_the_same_meeting',
      endpoint: realIntake.provider_endpoint,
      required_coverage: requiredProviderCoverage,
    },
    {
      id: 'capture_ended_state',
      owner: 'field_operator',
      objective: 'record_meeting_end_state_or_host_leave_state_for_axis_close_validation',
      required_snapshot: requiredSnapshots.includes('meeting_ended') ? 'meeting_ended' : undefined,
    },
    {
      id: 'build_evidence_package',
      owner: 'sdk_cli',
      objective: 'normalize_field_evidence_into_a_handoff_package',
      command: commands.build_field_evidence,
      output: manifest.file_contract?.files?.evidence_package,
    },
    {
      id: 'run_real_intake_gate',
      owner: 'sdk_cli',
      objective: 'reject_fixture_or_partial_evidence_before_handoff',
      command: commands.validate_real_intake,
      forbidden_inputs: realIntake.forbidden_inputs ?? [],
    },
  ].map((step) => compactObject(step));
}

function nextActions(provider = {}, field = {}, status = '') {
  const missingEnv = (provider.security?.missing_env ?? []).map((name) => `configure_env:${name}`);
  const fieldActions = field.next_actions ?? [];
  const statusAction = {
    provider_setup_missing_env: 'complete_provider_security_configuration',
    provider_setup_not_ready: 'fix_provider_setup_readiness',
    needs_local_observer_capture: 'capture_real_meeting_app_snapshots',
    needs_provider_event_capture: 'capture_real_provider_start_end_events',
    needs_field_evidence: 'build_and_verify_field_evidence_package',
    ready_to_collect_field_evidence: 'run_field_collection_in_real_meeting',
    production_ready: 'handoff_to_host_project_with_monitoring',
  }[status];
  return unique([
    ...missingEnv,
    ...fieldActions,
    statusAction,
  ]);
}

export function buildMeetingPlatformFieldIntakePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787'));
  const merged = {
    ...options,
    baseUrl,
  };
  const provider = buildMeetingPlatformProviderConnectionPack(key, merged);
  const fieldCapturePlan = buildMeetingPlatformFieldCapturePlan(key, merged);
  const manifest = buildMeetingPlatformFieldCaptureManifest(key, merged);
  const collector = buildMeetingPlatformFieldCollectorConfig(key, merged);
  const realIntake = buildMeetingPlatformRealEvidenceIntakePlan(key, merged);
  const status = statusFor(provider, fieldCapturePlan);
  const dirs = evidenceDirs(merged);
  const commands = planCommands(key, manifest, merged);
  return compactObject({
    type: 'meeting_platform_field_intake_plan',
    schema: MEETING_PLATFORM_FIELD_INTAKE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_INTAKE_SCHEMA_VERSION,
    platform: key,
    display_name: fieldCapturePlan.display_name ?? provider.display_name,
    status,
    production_ready: fieldCapturePlan.production_ready === true,
    ready_for_realtime_annotations: fieldCapturePlan.ready_for_realtime_annotations === true,
    objective: 'operator_ready_plan_for_real_meeting_timeline_adapter_intake',
    base_url: baseUrl,
    provider_endpoint: realIntake.provider_endpoint,
    evidence_dirs: dirs,
    files: manifest.file_contract?.files,
    required_inputs: realIntake.required_inputs,
    forbidden_inputs: realIntake.forbidden_inputs,
    provider_connection: {
      transport: provider.transport,
      endpoint: provider.endpoint,
      status_endpoint: provider.status_endpoint,
      event_mapping: provider.event_mapping,
      security: provider.security,
      readiness: provider.readiness,
      realtime_annotation_policy: provider.realtime_annotation_policy,
    },
    local_observer: collector.local_snapshot_collector ? {
      mode: collector.mode,
      browser_observer: collector.browser_observer,
      local_snapshot_collector: collector.local_snapshot_collector,
      timeline_ingest: collector.timeline_ingest,
    } : undefined,
    provider_observer: collector.provider_observer,
    acceptance: {
      production_condition: realIntake.output_contract?.production_condition,
      pilot_condition: manifest.acceptance?.pilot_condition,
      timestamp_field: manifest.runtime_contract?.annotation_timestamp_field,
      required_provider_coverage: manifest.acceptance?.required_provider_coverage ?? [],
      required_local_snapshots: manifest.acceptance?.required_local_snapshots ?? [],
      minimum_provider_records: manifest.acceptance?.minimum_provider_records ?? 0,
      minimum_local_records: manifest.acceptance?.minimum_local_records ?? 0,
      current_missing_items: fieldCapturePlan.missing_items ?? [],
    },
    commands,
    operator_steps: operatorSteps(key, provider, manifest, realIntake, commands),
    sdk_methods: [
      'buildMeetingPlatformFieldIntakePlan',
      'buildMeetingPlatformFieldCaptureManifest',
      'buildMeetingPlatformFieldEvidenceBundle',
      'buildMeetingPlatformRealEvidenceIntakeReport',
    ],
    handoff: {
      primary_entry: '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake',
      kit_methods: ['platformFieldIntakePlan', 'platformFieldIntakeMatrix'],
      package_schema: 'meeting_platform_evidence_package',
      final_gate: 'meeting_platform_real_evidence_intake',
    },
    field_capture_plan: fieldCapturePlan,
    field_capture_manifest: options.includeManifest === true || options.include_manifest === true ? manifest : undefined,
    collector_config: options.includeCollectorConfig === true || options.include_collector_config === true ? collector : undefined,
    real_intake_plan: options.includeRealIntakePlan === true || options.include_real_intake_plan === true ? realIntake : undefined,
    next_actions: nextActions(provider, fieldCapturePlan, status),
  });
}

export function buildMeetingPlatformFieldIntakeMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformFieldIntakePlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_field_intake_matrix',
    schema: MEETING_PLATFORM_FIELD_INTAKE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_FIELD_INTAKE_SCHEMA_VERSION,
    platform_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    realtime_ready_count: plans.filter((plan) => plan.ready_for_realtime_annotations).length,
    provider_blocked_count: plans.filter((plan) => String(plan.status).startsWith('provider_setup')).length,
    local_capture_needed_count: plans.filter((plan) => plan.status === 'needs_local_observer_capture').length,
    provider_capture_needed_count: plans.filter((plan) => plan.status === 'needs_provider_event_capture').length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      production_ready: plan.production_ready,
      ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
      provider_endpoint: plan.provider_endpoint,
      missing_env: plan.provider_connection?.security?.missing_env ?? [],
      required_provider_coverage: plan.acceptance?.required_provider_coverage ?? [],
      required_local_snapshots: plan.acceptance?.required_local_snapshots ?? [],
      field_evidence_input: plan.files?.field_evidence_input,
      evidence_package: plan.files?.evidence_package,
      next_actions: plan.next_actions,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}
