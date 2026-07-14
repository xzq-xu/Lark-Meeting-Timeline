import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA = 'meeting_platform_adapter_import_plan';
export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_MATRIX_SCHEMA = 'meeting_platform_adapter_import_plan_matrix';
export const MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA_VERSION = 1;

const TARGET_RANK = Object.freeze({
  static: 0,
  pilot: 1,
  production: 2,
});

const SURFACE_ALIASES = Object.freeze({
  browser: 'browser_extension',
  browser_extension: 'browser_extension',
  'browser-extension': 'browser_extension',
  extension: 'browser_extension',
  browser_extension_or_native_detector: 'browser_extension',
  browser_extension_or_desktop_observer: 'browser_extension',
  browser_extension_or_webview: 'browser_extension',
  browser_extension_or_webview_preload: 'browser_extension',
  webview: 'webview_preload',
  webview_preload: 'webview_preload',
  'webview-preload': 'webview_preload',
  electron: 'webview_preload',
  native: 'native_detector',
  native_detector: 'native_detector',
  'native-detector': 'native_detector',
  native_host: 'native_detector',
  'native-host': 'native_detector',
  desktop: 'native_detector',
  desktop_observer: 'native_detector',
  desktop_or_browser_observer: 'native_detector',
  native_or_desktop_observer: 'native_detector',
  host: 'native_detector',
  detector: 'native_detector',
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

function targetFrom(input = {}, options = {}) {
  const raw = String(firstNonEmpty(
    options.target,
    options.acceptanceTarget,
    options.acceptance_target,
    input.target,
    input.acceptanceTarget,
    input.acceptance_target,
    'static',
  ));
  return TARGET_RANK[raw] == null ? 'static' : raw;
}

function requiredForTarget(requiredFrom = 'static', target = 'static') {
  return TARGET_RANK[target] >= TARGET_RANK[requiredFrom ?? 'static'];
}

function normalizeSurface(value) {
  const key = String(value ?? '').trim();
  return SURFACE_ALIASES[key] ?? SURFACE_ALIASES[key.toLowerCase()] ?? key;
}

function surfacePreference(pkg = {}, input = {}, options = {}) {
  const explicit = firstNonEmpty(
    options.surface,
    options.preferredSurface,
    options.preferred_surface,
    input.surface,
    input.preferredSurface,
    input.preferred_surface,
  );
  if (explicit) return normalizeSurface(explicit);
  const selected = adapterSelectionSummary(pkg, input, options);
  if (selected.axis_surface) return normalizeSurface(selected.axis_surface);
  const recommended = normalizeSurface(pkg.recommended_first_surface);
  if (pkg.surface_entrypoints?.[recommended]) return recommended;
  if (pkg.surface_entrypoints?.browser_extension?.ready === true) return 'browser_extension';
  if (pkg.surface_entrypoints?.webview_preload?.ready === true) return 'webview_preload';
  if (pkg.surface_entrypoints?.native_detector?.ready === true) return 'native_detector';
  if (pkg.surface_entrypoints?.native_host?.ready === true) return 'native_detector';
  if (pkg.surface_entrypoints?.provider_reconcile?.ready === true) return 'provider_reconcile';
  return recommended || 'native_detector';
}

function surfaceEntrypoint(pkg = {}, surface = '') {
  if (surface === 'native_detector') {
    return pkg.surface_entrypoints?.native_detector ?? pkg.surface_entrypoints?.native_host;
  }
  return pkg.surface_entrypoints?.[surface];
}

function normalizedPaths(values = []) {
  return new Set(asArray(values).map((value) => String(value).replace(/\\/g, '/').replace(/^\.\//, '')));
}

function hostFileCoverage(pkg = {}, input = {}, options = {}) {
  const available = firstNonEmpty(
    options.availableFiles,
    options.available_files,
    input.availableFiles,
    input.available_files,
  );
  const target = targetFrom(input, options);
  const requiredFiles = asArray(pkg.host_files).filter((file) => requiredForTarget(file.required_from, target));
  if (available == null) {
    return {
      checked: false,
      status: 'not_checked',
      required_count: requiredFiles.length,
      present_count: 0,
      missing: [],
    };
  }
  const paths = normalizedPaths(available);
  const missing = requiredFiles.filter((file) => !paths.has(String(file.path)));
  return {
    checked: true,
    status: missing.length === 0 ? 'complete' : 'missing_required_files',
    required_count: requiredFiles.length,
    present_count: requiredFiles.length - missing.length,
    missing,
  };
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function adapterSelectionSummary(pkg = {}, input = {}, options = {}) {
  const selection = firstNonEmpty(
    options.adapterSelection,
    options.adapter_selection,
    input.adapterSelection,
    input.adapter_selection,
    pkg.adapter_selection,
    pkg.artifacts?.adapter_selection,
  );
  const ref = pkg.artifact_refs?.adapter_selection ?? {};
  const selected = selection?.selection ?? selection ?? {};
  return compactObject({
    available: Boolean(selection || ref.path),
    schema: selection?.schema ?? ref.schema ?? (ref.path ? 'meeting_platform_adapter_selection' : undefined),
    path: ref.path,
    ready: selection?.readiness?.selection_ready ?? selection?.ready,
    axis_source: selected.axis_source,
    axis_surface: selected.axis_surface,
    annotation_source: selected.annotation_source,
    timestamp_field: selected.timestamp_field,
    provider_reconcile_source: selected.provider_reconcile_source,
    provider_reconcile_required_for_production: selected.provider_reconcile_required_for_production,
    speaker_track_source: selected.speaker_track_source,
    post_meeting_artifact_source: selected.post_meeting_artifact_source,
    provider_events_block_realtime: firstNonEmpty(
      selection?.runtime_policy?.provider_events_block_realtime,
      selection?.provider_events_block_realtime,
    ),
    transcript_blocks_realtime: firstNonEmpty(
      selection?.runtime_policy?.transcript_blocks_realtime,
      selection?.transcript_blocks_realtime,
    ),
    startup_order: selection?.runtime_policy?.startup_order ?? selection?.startup_order,
    runtime_actions: selection?.runtime_policy?.runtime_actions,
    command: pkg.commands?.adapter_selection,
    next_actions: selection?.next_actions,
  });
}

function readiness(pkg = {}, surface = '', coverage = {}, input = {}, options = {}, adapterSelection = {}) {
  const target = targetFrom(input, options);
  const selected = surfaceEntrypoint(pkg, surface) ?? {};
  const allowCustomAuthoring = firstNonEmpty(
    options.allowCustomAuthoring,
    options.allow_custom_authoring,
    input.allowCustomAuthoring,
    input.allow_custom_authoring,
    false,
  ) === true;
  const schemaValid = pkg.schema === 'meeting_platform_adapter_export_package';
  const hardContractReady = pkg.timestamp_field === 'captured_at_ms'
    && pkg.local_axis_first === true
    && pkg.provider_events_block_realtime === false
    && pkg.transcript_blocks_realtime === false;
  const preflightContractReady = pkg.adapter_preflight?.live_evidence_required === true
    && pkg.adapter_preflight?.url_only_status === 'needs_live_page_evidence'
    && Boolean(pkg.adapter_preflight?.required_before);
  const rawSignalValidationReady = Boolean(
    pkg.artifact_refs?.raw_signal_validation?.path
      || pkg.artifacts?.raw_signal_validation
      || pkg.raw_signal_validation,
  ) || (allowCustomAuthoring && pkg.built_in !== true);
  const surfaceReady = selected.ready === true || (surface === 'provider_reconcile' && selected.required_for_realtime === false);
  const packageAccepted = pkg.accepted === true || (allowCustomAuthoring && pkg.built_in !== true);
  const fileCoverageReady = coverage.checked !== true || coverage.missing.length === 0;
  const adapterSelectionRequired = pkg.built_in === true || adapterSelection.available === true;
  const adapterSelectionReady = !adapterSelectionRequired || (
    adapterSelection.available === true
    && adapterSelection.ready === true
    && adapterSelection.timestamp_field === 'captured_at_ms'
    && adapterSelection.provider_events_block_realtime === false
    && adapterSelection.transcript_blocks_realtime === false
  );
  const importReady = schemaValid && hardContractReady && preflightContractReady && rawSignalValidationReady && adapterSelectionReady && surfaceReady && packageAccepted && fileCoverageReady;
  const issues = [
    schemaValid ? undefined : issue('error', 'invalid_export_package_schema', 'Adapter export package schema is invalid.', {
      actual: pkg.schema,
    }),
    pkg.timestamp_field === 'captured_at_ms' ? undefined : issue('error', 'invalid_timestamp_field', 'Realtime marks must use captured_at_ms.', {
      actual: pkg.timestamp_field,
    }),
    pkg.local_axis_first === true ? undefined : issue('error', 'local_axis_first_required', 'Realtime annotations must bind a local axis before provider reconcile.'),
    pkg.provider_events_block_realtime === false ? undefined : issue('error', 'provider_events_must_be_nonblocking', 'Provider events must not block realtime mark insertion.'),
    pkg.transcript_blocks_realtime === false ? undefined : issue('error', 'transcript_must_be_nonblocking', 'Transcript artifacts must remain post-meeting backfill.'),
    preflightContractReady ? undefined : issue('error', 'adapter_preflight_gate_required', 'Adapter import must declare a live-evidence preflight gate before realtime annotation insertion.'),
    rawSignalValidationReady ? undefined : issue('error', 'raw_signal_validation_required', 'Adapter import must include raw signal validation before realtime wiring.', {
      expected: 'raw-signal-validation.json or raw_signal_validation artifact',
    }),
    adapterSelectionReady ? undefined : issue('error', 'adapter_selection_required', 'Adapter import must include a ready adapter selection before host runtime surface wiring.', {
      expected: 'adapter-selection.json or adapter_selection export package field',
      path: adapterSelection.path,
      ready: adapterSelection.ready,
      timestamp_field: adapterSelection.timestamp_field,
    }),
    surfaceReady ? undefined : issue('error', 'selected_surface_not_ready', 'Selected import surface is not ready in the export package.', {
      surface,
    }),
    packageAccepted ? undefined : issue('error', 'export_package_not_accepted', 'Adapter export package did not pass its target acceptance gate.', {
      target: pkg.target,
      built_in: pkg.built_in,
    }),
    fileCoverageReady ? undefined : issue('error', 'missing_required_host_files', 'Required host files are not available to the consumer project.', {
      target,
      missing: coverage.missing.map((file) => file.path),
    }),
  ].filter(Boolean);
  return {
    schema_valid: schemaValid,
    hard_contract_ready: hardContractReady,
    adapter_preflight_contract_ready: preflightContractReady,
    raw_signal_validation_ready: rawSignalValidationReady,
    adapter_selection_ready: adapterSelectionReady,
    selected_surface_ready: surfaceReady,
    export_package_accepted: pkg.accepted === true,
    package_accepted_for_import: packageAccepted,
    file_coverage_ready: fileCoverageReady,
    import_ready: importReady,
    issue_count: issues.length,
    issues,
  };
}

function sdkImports(pkg = {}) {
  const imports = pkg.import_paths ?? {};
  return {
    sdk_root: imports.sdk_root ?? '@ai-annotation/meeting-timeline-sdk',
    platform_kit: imports.platform_kit ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    adapter_export_package: imports.adapter_export_package ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-export-package',
    adapter_blueprint: imports.adapter_blueprint ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint',
    adapter_selection: imports.adapter_selection ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-selection',
    adapter_preflight: imports.adapter_preflight ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight',
    adapter_import_plan: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-import-plan',
    connector: imports.connector ?? '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector',
    runtime_bundle: imports.runtime_bundle ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
    provider_connection: imports.provider_connection ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection',
    raw_signal: imports.raw_signal ?? '@ai-annotation/meeting-timeline-sdk/adapters/platform-raw-signal',
  };
}

function adapterPreflightSummary(pkg = {}) {
  const preflight = pkg.adapter_preflight ?? {};
  return compactObject({
    status: preflight.status,
    accepted: preflight.accepted === true,
    selected_surface: preflight.selected_surface,
    startup_ready: preflight.startup_ready === true,
    live_evidence_ready: preflight.live_evidence_ready === true,
    realtime_annotation_ready: preflight.realtime_annotation_ready === true,
    required_before: preflight.required_before,
    live_evidence_required: preflight.live_evidence_required === true,
    url_only_status: preflight.url_only_status,
    sdk_methods: preflight.sdk_methods,
    bridge_messages: preflight.bridge_messages,
    command: preflight.command ?? pkg.commands?.adapter_preflight,
    first_next_action: preflight.first_next_action,
  });
}

function adapterBlueprintSummary(pkg = {}, input = {}, options = {}) {
  const blueprint = firstNonEmpty(
    options.adapterBlueprint,
    options.adapter_blueprint,
    input.adapterBlueprint,
    input.adapter_blueprint,
    pkg.artifacts?.adapter_blueprint,
  );
  const ref = pkg.artifact_refs?.adapter_blueprint ?? {};
  return compactObject({
    available: Boolean(ref.path || blueprint),
    schema: blueprint?.schema ?? ref.schema ?? (ref.path ? 'meeting_platform_adapter_blueprint' : undefined),
    path: ref.path,
    command: pkg.commands?.adapter_blueprint,
    ready: blueprint?.readiness?.ready,
    primary_surface: blueprint?.primary_surface,
    surface_order: blueprint?.surface_order,
    timestamp_field: firstNonEmpty(
      blueprint?.annotation_contract?.timestamp_field,
      blueprint?.realtime_axis_contract?.timestamp_field,
    ),
    provider_blocks_realtime: firstNonEmpty(
      blueprint?.surfaces?.provider_reconcile?.blocks_realtime,
      blueprint?.runtime_contract?.provider_events_block_realtime,
    ),
    transcript_blocks_realtime: blueprint?.runtime_contract?.transcript_blocks_realtime,
    first_acceptance_gate: blueprint?.acceptance_gates?.realtime_pilot?.[0],
  });
}

function rawSignalValidationSummary(pkg = {}, input = {}, options = {}) {
  const rawSignalValidation = firstNonEmpty(
    options.rawSignalValidation,
    options.raw_signal_validation,
    input.rawSignalValidation,
    input.raw_signal_validation,
    pkg.artifacts?.raw_signal_validation,
    pkg.raw_signal_validation,
  );
  const ref = pkg.artifact_refs?.raw_signal_validation ?? {};
  const outputContract = rawSignalValidation?.output_contract ?? {};
  const sampleSummary = rawSignalValidation?.sample_summary ?? {};
  return compactObject({
    available: Boolean(ref.path || rawSignalValidation),
    schema: rawSignalValidation?.schema ?? ref.schema ?? (ref.path ? 'meeting_platform_adapter_raw_signal_validation' : undefined),
    path: ref.path,
    command: rawSignalValidation?.command ?? pkg.commands?.raw_signal_validation,
    status: rawSignalValidation?.status,
    input_contract: rawSignalValidation?.input_contract,
    runtime_actions: outputContract.runtime_actions,
    runtime_action_count: outputContract.runtime_actions?.length,
    signal_count: sampleSummary.signal_count,
    runtime_event_count: sampleSummary.runtime_event_count,
    filtered_speaker_event_count: sampleSummary.filtered_speaker_event_count,
    filter_active_speaker_samples: rawSignalValidation?.filter_policy?.filter_active_speaker_samples,
  });
}

function installSteps(pkg = {}, surface = '', coverage = {}, rawSignalValidation = {}) {
  const selected = surfaceEntrypoint(pkg, surface) ?? {};
  return [
    {
      step: 1,
      id: 'load_adapter_export_package',
      action: 'read adapter-export-package.json and host_files',
      required: true,
      files_checked: coverage.checked,
      missing_files: coverage.missing.map((file) => file.path),
    },
    {
      step: 2,
      id: 'read_adapter_selection',
      action: 'read adapter_selection and bind selected axis/surface before host runtime wiring',
      required: true,
      command: pkg.commands?.adapter_selection,
      artifact_path: pkg.artifact_refs?.adapter_selection?.path,
      sdk_method: 'platformAdapterSelection',
    },
    {
      step: 3,
      id: 'install_sdk',
      action: 'install @ai-annotation/meeting-timeline-sdk and import SDK facade',
      required: true,
      imports: sdkImports(pkg),
    },
    {
      step: 4,
      id: 'create_sdk_or_kit',
      action: 'createMeetingAppTimelineSdk({ baseUrl, platforms }) or createMeetingPlatformTimelineKit(client)',
      required: true,
      platform: pkg.platform,
    },
    {
      step: 5,
      id: 'install_selected_surface',
      action: `install ${surface} surface using export package entrypoint metadata`,
      required: true,
      surface,
      entrypoint: selected,
    },
    {
      step: 6,
      id: 'run_raw_signal_validation',
      action: 'run raw signal validation with examples, then replay host-captured snapshots, marks, and speaker samples',
      required: true,
      command: rawSignalValidation.command ?? pkg.commands?.raw_signal_validation,
      sdk_method: 'platformRawSignalBatch',
      artifact_path: rawSignalValidation.path,
      expected_runtime_actions: rawSignalValidation.runtime_actions,
    },
    {
      step: 7,
      id: 'run_adapter_preflight',
      action: 'run adapter preflight with live DOM/native-window evidence before opening a realtime annotation session',
      required: true,
      sdk_method: 'platformAdapterPreflight',
      message_types: [
        'meeting_timeline.preflight_current_window',
        'meeting_timeline.preflight_candidates',
      ],
      url_only_status: 'needs_live_page_evidence',
      timestamp_field: 'captured_at_ms',
    },
    {
      step: 8,
      id: 'bind_current_axis',
      action: 'call observePlatformCandidates before first realtime mark',
      required: true,
      sdk_method: 'observePlatformCandidates',
      timestamp_field: 'captured_at_ms',
    },
    {
      step: 9,
      id: 'insert_realtime_marks',
      action: `call insertAnnotation('${pkg.platform}', { captured_at_ms, ...mark })`,
      required: true,
      sdk_method: 'insertAnnotation',
      timestamp_field: 'captured_at_ms',
    },
    {
      step: 10,
      id: 'emit_speaker_participant_positions',
      action: 'call speakerTrack/participantTrack when local samples are available',
      required: false,
      sdk_methods: ['speakerTrack', 'participantTrack'],
    },
    {
      step: 11,
      id: 'connect_provider_reconcile',
      action: 'connect provider events only as reconcile/backfill after local axis exists',
      required: false,
      sdk_method: 'ingestProvider',
      provider_realtime_blocking: false,
    },
  ];
}

function nextActions(pkg = {}, ready = {}, coverage = {}) {
  return unique([
    ...(ready.issues ?? []).map((item) => item.code),
    ...(coverage.missing ?? []).map((file) => `provide_host_file:${file.path}`),
    ...(pkg.next_actions ?? []),
    ready.import_ready ? 'run_raw_signal_validation_with_host_captured_samples' : undefined,
    ready.import_ready ? 'run_adapter_preflight_with_live_window_evidence_before_first_insert' : undefined,
    ready.import_ready ? 'wire_selected_surface_into_host_runtime' : 'fix_adapter_import_plan_blockers',
  ]);
}

export function buildMeetingPlatformAdapterImportPlan(exportPackage = {}, input = {}, options = {}) {
  const target = targetFrom(input, options);
  const adapterSelection = adapterSelectionSummary(exportPackage, input, options);
  const surface = surfacePreference(exportPackage, input, options);
  const coverage = hostFileCoverage(exportPackage, input, { ...options, target });
  const ready = readiness(exportPackage, surface, coverage, input, { ...options, target }, adapterSelection);
  const blueprint = adapterBlueprintSummary(exportPackage, input, options);
  const preflight = adapterPreflightSummary(exportPackage);
  const rawSignalValidation = rawSignalValidationSummary(exportPackage, input, options);
  return {
    type: 'meeting_platform_adapter_import_plan',
    schema: MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA_VERSION,
    platform: exportPackage.platform,
    display_name: exportPackage.display_name,
    target,
    selected_surface: surface,
    accepted: ready.import_ready,
    import_ready: ready.import_ready,
    built_in: exportPackage.built_in === true,
    export_package_schema: exportPackage.schema,
    export_package_target: exportPackage.target,
    export_package_accepted: exportPackage.accepted === true,
    runtime_contract: {
      timestamp_field: exportPackage.timestamp_field,
      local_axis_first: exportPackage.local_axis_first === true,
      provider_events_block_realtime: exportPackage.provider_events_block_realtime === true,
      transcript_blocks_realtime: exportPackage.transcript_blocks_realtime === true,
      adapter_selection_required_before_surface_install: exportPackage.built_in === true || adapterSelection.available === true,
      adapter_selection_axis_source: adapterSelection.axis_source,
      adapter_selection_axis_surface: adapterSelection.axis_surface,
      raw_signal_validation_required_before_preflight: rawSignalValidation.available === true,
      adapter_preflight_required_before_realtime_insert: preflight.live_evidence_required === true,
      adapter_preflight_url_only_status: preflight.url_only_status,
    },
    adapter_blueprint: blueprint,
    adapter_selection: adapterSelection,
    raw_signal_validation: rawSignalValidation,
    adapter_preflight: preflight,
    sdk_imports: sdkImports(exportPackage),
    host_file_coverage: coverage,
    surface_entrypoints: exportPackage.surface_entrypoints,
    surface_entrypoint: surfaceEntrypoint(exportPackage, surface),
    install_steps: installSteps(exportPackage, surface, coverage, rawSignalValidation),
    commands: compactObject({
      export_package: exportPackage.commands?.export_package,
      acceptance_checklist: exportPackage.commands?.acceptance_checklist,
      adapter_selection: exportPackage.commands?.adapter_selection,
      raw_signal_validation: exportPackage.commands?.raw_signal_validation,
      runtime_bundle: exportPackage.commands?.runtime_bundle,
      import_plan: 'npm run meeting-platform:adapter-import-plan',
      sdk_test: exportPackage.commands?.sdk_test ?? 'npm run sdk:test',
    }),
    readiness: ready,
    source_package: compactObject({
      platform: exportPackage.platform,
      schema: exportPackage.schema,
      target: exportPackage.target,
      host_file_count: exportPackage.host_files?.length ?? 0,
      package_role: exportPackage.package_role,
      adapter_selection_path: adapterSelection.path,
      adapter_blueprint_path: blueprint.path,
      raw_signal_validation_path: rawSignalValidation.path,
    }),
    next_actions: nextActions(exportPackage, ready, coverage),
  };
}

export function assertMeetingPlatformAdapterImportPlan(planOrPackage = {}, input = {}, options = {}) {
  const plan = planOrPackage.schema === MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA
    ? planOrPackage
    : buildMeetingPlatformAdapterImportPlan(planOrPackage, input, options);
  if (plan.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter import plan is not ready', {
      code: 'meeting_platform_adapter_import_plan_not_ready',
      issues: plan.readiness?.issues ?? [],
      missing_host_files: plan.host_file_coverage?.missing?.map((file) => file.path) ?? [],
    });
  }
  return plan;
}

export function buildMeetingPlatformAdapterImportPlanMatrix(packagesOrInput = {}, input = {}, options = {}) {
  const packages = Array.isArray(packagesOrInput)
    ? packagesOrInput
    : asArray(firstNonEmpty(packagesOrInput.packages, packagesOrInput.exportPackages, packagesOrInput.export_packages));
  const matrixInput = Array.isArray(packagesOrInput) ? input : { ...packagesOrInput, ...input };
  const plans = packages.map((pkg) => buildMeetingPlatformAdapterImportPlan(pkg, matrixInput, options));
  return {
    type: 'meeting_platform_adapter_import_plan_matrix',
    schema: MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_IMPORT_PLAN_SCHEMA_VERSION,
    target: targetFrom(matrixInput, options),
    package_count: packages.length,
    accepted_count: plans.filter((plan) => plan.accepted === true).length,
    blocked_count: plans.filter((plan) => plan.accepted !== true).length,
    file_coverage_checked_count: plans.filter((plan) => plan.host_file_coverage.checked === true).length,
    missing_file_count: plans.reduce((count, plan) => count + (plan.host_file_coverage.missing?.length ?? 0), 0),
    adapter_preflight_startup_ready_count: plans.filter((plan) => plan.adapter_preflight?.startup_ready === true).length,
    adapter_preflight_realtime_ready_count: plans.filter((plan) => plan.adapter_preflight?.realtime_annotation_ready === true).length,
    adapter_selection_ready_count: plans.filter((plan) => plan.adapter_selection?.ready === true).length,
    raw_signal_validation_ready_count: plans.filter((plan) => plan.raw_signal_validation?.available === true).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      selected_surface: plan.selected_surface,
      accepted: plan.accepted,
      built_in: plan.built_in,
      export_package_accepted: plan.export_package_accepted,
      hard_contract_ready: plan.readiness.hard_contract_ready,
      adapter_selection_ready: plan.adapter_selection?.ready === true,
      adapter_selection_axis_source: plan.adapter_selection?.axis_source,
      adapter_selection_axis_surface: plan.adapter_selection?.axis_surface,
      adapter_selection_timestamp_field: plan.adapter_selection?.timestamp_field,
      raw_signal_validation_ready: plan.readiness.raw_signal_validation_ready,
      raw_signal_validation_status: plan.raw_signal_validation?.status,
      selected_surface_ready: plan.readiness.selected_surface_ready,
      file_coverage_ready: plan.readiness.file_coverage_ready,
      adapter_blueprint_available: plan.adapter_blueprint?.available === true,
      adapter_blueprint_primary_surface: plan.adapter_blueprint?.primary_surface,
      adapter_preflight_status: plan.adapter_preflight?.status,
      adapter_preflight_selected_surface: plan.adapter_preflight?.selected_surface,
      adapter_preflight_startup_ready: plan.adapter_preflight?.startup_ready === true,
      adapter_preflight_realtime_ready: plan.adapter_preflight?.realtime_annotation_ready === true,
      missing_file_count: plan.host_file_coverage.missing?.length ?? 0,
      first_issue: plan.readiness.issues?.[0]?.code,
      first_next_action: plan.next_actions?.[0],
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}
