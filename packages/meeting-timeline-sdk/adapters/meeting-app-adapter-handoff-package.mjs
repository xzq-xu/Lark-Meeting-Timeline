import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingAppAdapterManifest,
} from './meeting-app-adapter-manifest.mjs';
import {
  MEETING_APP_ADAPTER_SPEC_SCHEMA,
  buildMeetingAppAdapterSpec,
} from './meeting-app-adapter-spec.mjs';
import {
  MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA,
  buildMeetingAppAdapterRuntimeConfig,
} from './meeting-app-adapter-runtime-config.mjs';

export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA = 'meeting_app_adapter_handoff_package';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA = 'meeting_app_adapter_handoff_package_matrix';
export const MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA = 'meeting_app_adapter_verification_plan';
export const MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA = 'meeting_app_adapter_verification_report';
export const MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA = 'meeting_app_adapter_verification_report_matrix';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION = 1;

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

function packageInputs(options = {}) {
  const specs = asArray(firstNonEmpty(options.packages, options.configs, options.adapters, options.adapter_specs, options.adapterSpecs, options.specs));
  const platforms = asArray(firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys, []));
  return [...platforms, ...specs];
}

function packageOrReportInputs(options = {}) {
  const reports = asArray(firstNonEmpty(options.reports, options.verification_reports, options.verificationReports, []));
  return reports.length > 0 ? reports : packageInputs(options);
}

function cleanOptions(options = {}) {
  return {
    ...options,
    packages: undefined,
    configs: undefined,
    adapters: undefined,
    adapter_specs: undefined,
    adapterSpecs: undefined,
    specs: undefined,
    platforms: undefined,
    platform_keys: undefined,
    platformKeys: undefined,
  };
}

function maybeBuiltInManifest(spec = {}, options = {}) {
  if (spec.source !== 'built_in_manifest') return undefined;
  try {
    return buildMeetingAppAdapterManifest(spec.platform ?? spec.adapter_key, options);
  } catch {
    return undefined;
  }
}

function integrationReadme({ spec, runtimeConfig, manifest }) {
  return [
    `# ${spec.display_name || spec.adapter_key} Meeting Adapter Handoff`,
    '',
    'This package is the host-side contract for adding this meeting app to Meeting Timeline annotations.',
    '',
    '## Files',
    '- `adapter-spec.json`: static URL, permission, selector, MutationObserver and nonblocking contract.',
    '- `runtime-config.json`: options for browser extension, WebView preload or Electron content script runtime.',
    '- `extension-manifest-fragment.json`: match patterns, permissions and content script fragment.',
    '- `verification-plan.json`: live evidence and acceptance checklist for the host implementation.',
    manifest ? '- `adapter-manifest.json`: built-in adapter manifest used to derive this package.' : '',
    '',
    '## Minimal Integration',
    '1. Install `@ai-annotation/meeting-timeline-sdk` in the host project.',
    '2. Load `runtime-config.json` and pass `content_script_options` to `installMeetingAppContentScriptBridge()`.',
    '3. Pass `browser_runtime_options` to `createMeetingAppBrowserRuntime()` when observing inside a page context.',
    '4. Pass `capture_options` to `captureMeetingAppDomSnapshot()` when taking manual live snapshots.',
    '5. Keep provider events and transcript import nonblocking for realtime annotation placement.',
    '',
    '## Validation',
    '- Capture live DOM snapshots before production rollout.',
    '- Verify speaker and participant track extraction on real meetings.',
    '- Verify annotation insertions keep `captured_at_ms` and land on the current meeting axis.',
    '- Keep `verification-plan.json` with the evidence package used for handoff readiness.',
    '',
    '## Contract',
    `- adapter_key: ${spec.adapter_key}`,
    `- timestamp_field: ${runtimeConfig.contracts?.timestamp_field}`,
    `- provider_events_block_realtime: ${runtimeConfig.contracts?.provider_events_block_realtime}`,
    `- transcript_blocks_realtime: ${runtimeConfig.contracts?.transcript_blocks_realtime}`,
    '',
  ].filter(Boolean).join('\n');
}

function verificationPlan({ spec, runtimeConfig, manifest }) {
  const adapterKey = spec.adapter_key;
  const requiredEvidence = [
    {
      id: 'live_dom_snapshot',
      phase: 'live_capture',
      required: true,
      minimum_count: 1,
      accepted_artifacts: ['meeting_app_snapshot_record_set', 'meeting_app_live_evidence_package'],
      proof_fields: ['snapshot_records[].captured_at_ms', 'snapshot_records[].page.controls', 'snapshot_records[].page.participants'],
      purpose: 'Proves the selector set can see controls and participants on a real meeting page.',
    },
    {
      id: 'candidate_observation',
      phase: 'runtime_observation',
      required: true,
      minimum_count: 1,
      accepted_artifacts: ['observe_candidates_message', 'meeting_platform_runtime_event'],
      proof_fields: ['captured_at_ms', 'platform', 'url'],
      purpose: 'Proves local observation can timestamp meeting candidates without provider events.',
    },
    {
      id: 'speaker_track',
      phase: 'runtime_observation',
      required: true,
      minimum_count: 1,
      accepted_artifacts: ['speaker_track', 'meeting_platform_evidence_package'],
      proof_fields: ['speaker_segments[].captured_at_ms', 'speaker_segments[].speaker_label'],
      purpose: 'Proves speaker positions can be drawn on the meeting timeline without realtime transcript.',
    },
    {
      id: 'participant_track',
      phase: 'runtime_observation',
      required: true,
      minimum_count: 1,
      accepted_artifacts: ['participant_track', 'meeting_platform_evidence_package'],
      proof_fields: ['participant_segments[].captured_at_ms', 'participant_segments[].participant_label'],
      purpose: 'Proves visible participant positions can be tracked for the timeline.',
    },
    {
      id: 'annotation_insert_current_axis',
      phase: 'annotation_runtime',
      required: true,
      minimum_count: 1,
      accepted_artifacts: ['timeline_annotation', 'meeting_platform_runtime_replay'],
      proof_fields: ['annotations[].captured_at_ms', 'annotations[].meeting_id', 'annotations[].axis_id'],
      purpose: 'Proves handwritten marks land on the current meeting axis at capture time.',
    },
  ];
  const staticChecks = [
    {
      id: 'adapter_spec_accepted',
      phase: 'static',
      blocking: true,
      status: spec.accepted === true ? 'passed' : 'failed',
      evidence: 'adapter-spec.json',
    },
    {
      id: 'runtime_config_accepted',
      phase: 'static',
      blocking: true,
      status: runtimeConfig.accepted === true ? 'passed' : 'failed',
      evidence: 'runtime-config.json',
    },
    {
      id: 'content_script_ready',
      phase: 'static',
      blocking: true,
      status: runtimeConfig.readiness?.content_script_ready === true ? 'passed' : 'failed',
      evidence: 'extension-manifest-fragment.json',
    },
    {
      id: 'captured_at_ms_contract',
      phase: 'static',
      blocking: true,
      status: runtimeConfig.contracts?.timestamp_field === 'captured_at_ms' ? 'passed' : 'failed',
      evidence: 'runtime-config.json',
    },
    {
      id: 'provider_transcript_nonblocking',
      phase: 'static',
      blocking: true,
      status: runtimeConfig.contracts?.provider_events_block_realtime === false
        && runtimeConfig.contracts?.transcript_blocks_realtime === false
        ? 'passed'
        : 'failed',
      evidence: 'runtime-config.json',
    },
  ];
  const liveChecks = requiredEvidence.map((item) => ({
    id: `verify_${item.id}`,
    phase: item.phase,
    blocking: true,
    status: 'pending_live_evidence',
    evidence: item.accepted_artifacts,
  }));
  return {
    type: 'meeting_app_adapter_verification_plan',
    schema: MEETING_APP_ADAPTER_VERIFICATION_PLAN_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    adapter_key: adapterKey,
    display_name: spec.display_name,
    source: spec.source,
    surface: spec.surface,
    generated_from: {
      adapter_spec: spec.schema,
      runtime_config: runtimeConfig.schema,
      adapter_manifest: manifest?.schema,
    },
    commands: {
      export_handoff_package: `npm run meeting-app:adapter-handoff-package -- --platforms=${adapterKey}`,
      run_live_evidence_gate: 'npm run meeting-app:evidence-gate',
      run_handoff_readiness: `npm run meeting-platform:handoff-readiness -- --platforms=${adapterKey}`,
      run_package_smoke: 'npm run sdk:package-smoke',
    },
    acceptance_policy: {
      static_checks_must_pass: true,
      live_evidence_required_before_production: true,
      provider_events_must_not_block_realtime: true,
      transcript_import_must_not_block_realtime: true,
      timestamp_field: 'captured_at_ms',
    },
    required_evidence: requiredEvidence,
    acceptance_checks: [...staticChecks, ...liveChecks],
    next_actions: [
      'capture_live_dom_snapshot',
      'run_candidate_observation',
      'verify_speaker_and_participant_tracks',
      'verify_current_axis_annotation_insert',
      'attach_evidence_package_to_handoff_readiness',
    ],
  };
}

function signalCount(value) {
  if (value === true) return 1;
  if (value === false || value == null || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? Math.max(0, value) : 0;
  if (Array.isArray(value)) return value.length;
  if (typeof value === 'object') {
    if (typeof value.count === 'number') return Math.max(0, value.count);
    if (typeof value.record_count === 'number') return Math.max(0, value.record_count);
    if (typeof value.accepted_count === 'number') return Math.max(0, value.accepted_count);
    if (typeof value.length === 'number') return Math.max(0, value.length);
    if (Array.isArray(value.records)) return value.records.length;
    if (Array.isArray(value.segments)) return value.segments.length;
    if (Array.isArray(value.annotations)) return value.annotations.length;
  }
  return 0;
}

function evidenceSource(options = {}) {
  return firstNonEmpty(
    options.evidence,
    options.evidence_package,
    options.evidencePackage,
    options.live_evidence,
    options.liveEvidence,
    {},
  );
}

function explicitEvidence(evidence = {}, id) {
  return firstNonEmpty(
    evidence[id],
    evidence.evidence?.[id],
    evidence.evidence_by_id?.[id],
    evidence.evidenceById?.[id],
    evidence.signals?.[id],
  );
}

function countRuntimeEvents(evidence = {}, matcher = () => false) {
  return asArray(firstNonEmpty(evidence.runtime_events, evidence.runtimeEvents, evidence.events))
    .filter((event) => matcher(event)).length;
}

function countEvidence(evidence = {}, id) {
  const explicit = explicitEvidence(evidence, id);
  if (explicit != null) return signalCount(explicit);
  switch (id) {
    case 'live_dom_snapshot':
      return Math.max(
        signalCount(firstNonEmpty(evidence.live_dom_snapshot_count, evidence.liveDomSnapshotCount, evidence.snapshot_count, evidence.snapshotCount)),
        signalCount(evidence.snapshot_records ?? evidence.snapshotRecords),
        signalCount(evidence.record_set ?? evidence.recordSet),
        signalCount(evidence.meeting_app_record_set ?? evidence.meetingAppRecordSet),
        signalCount(evidence.records),
      );
    case 'candidate_observation':
      return Math.max(
        signalCount(firstNonEmpty(evidence.candidate_observation_count, evidence.candidateObservationCount)),
        signalCount(evidence.candidate_observations ?? evidence.candidateObservations),
        signalCount(evidence.observe_candidates ?? evidence.observeCandidates),
        countRuntimeEvents(evidence, (event) => String(event?.type ?? event?.message_type ?? event?.action ?? '').includes('candidate')),
      );
    case 'speaker_track':
      return Math.max(
        signalCount(firstNonEmpty(evidence.speaker_track_count, evidence.speakerTrackCount)),
        signalCount(evidence.speaker_segments ?? evidence.speakerSegments),
        signalCount(evidence.speaker_track ?? evidence.speakerTrack),
        signalCount(evidence.tracks?.speaker),
      );
    case 'participant_track':
      return Math.max(
        signalCount(firstNonEmpty(evidence.participant_track_count, evidence.participantTrackCount)),
        signalCount(evidence.participant_segments ?? evidence.participantSegments),
        signalCount(evidence.participant_track ?? evidence.participantTrack),
        signalCount(evidence.tracks?.participant),
      );
    case 'annotation_insert_current_axis':
      return Math.max(
        signalCount(firstNonEmpty(evidence.annotation_insert_current_axis_count, evidence.annotationInsertCurrentAxisCount)),
        signalCount(evidence.current_axis_annotations ?? evidence.currentAxisAnnotations),
        signalCount(evidence.annotations),
        evidence.runtime_replay_accepted === true || evidence.runtimeHostReplayAccepted === true ? 1 : 0,
      );
    default:
      return 0;
  }
}

function packageIssues({ spec, runtimeConfig, files }) {
  return [
    spec.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
      ? undefined
      : issue('error', 'missing_adapter_spec', 'Handoff package requires an adapter spec.'),
    spec.accepted === true
      ? undefined
      : issue('error', 'adapter_spec_not_accepted', 'Adapter spec must pass static acceptance.', {
        adapter_key: spec.adapter_key,
        spec_issue_count: spec.issue_count,
      }),
    runtimeConfig.schema === MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA
      ? undefined
      : issue('error', 'missing_runtime_config', 'Handoff package requires a runtime config.'),
    runtimeConfig.accepted === true
      ? undefined
      : issue('error', 'runtime_config_not_accepted', 'Runtime config must pass static acceptance.', {
        adapter_key: runtimeConfig.adapter_key,
        runtime_issue_count: runtimeConfig.issue_count,
      }),
    runtimeConfig.readiness?.content_script_ready === true
      ? undefined
      : issue('error', 'content_script_not_ready', 'Handoff package requires content script match patterns.'),
    runtimeConfig.readiness?.capture_options_ready === true
      ? undefined
      : issue('error', 'capture_options_not_ready', 'Handoff package requires control and participant selectors.'),
    runtimeConfig.readiness?.mutation_observer_ready === true
      ? undefined
      : issue('error', 'mutation_observer_not_ready', 'Handoff package requires mutation observation selectors.'),
    files.some((file) => file.path === 'runtime-config.json')
      ? undefined
      : issue('error', 'missing_runtime_config_file', 'Handoff file list must include runtime-config.json.'),
    files.some((file) => file.path === 'verification-plan.json')
      ? undefined
      : issue('error', 'missing_verification_plan_file', 'Handoff file list must include verification-plan.json.'),
  ].filter(Boolean);
}

export function buildMeetingAppAdapterHandoffPackage(specOrPlatform = {}, options = {}) {
  const packageOptions = cleanOptions(options);
  const spec = specOrPlatform?.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
    ? specOrPlatform
    : buildMeetingAppAdapterSpec(specOrPlatform, packageOptions);
  const runtimeConfig = buildMeetingAppAdapterRuntimeConfig(spec, packageOptions);
  const manifest = maybeBuiltInManifest(spec, packageOptions);
  const plan = verificationPlan({ spec, runtimeConfig, manifest });
  const files = [
    {
      path: 'adapter-spec.json',
      media_type: 'application/json',
      schema: spec.schema,
      content: spec,
    },
    manifest
      ? {
        path: 'adapter-manifest.json',
        media_type: 'application/json',
        schema: manifest.schema,
        content: manifest,
      }
      : undefined,
    {
      path: 'runtime-config.json',
      media_type: 'application/json',
      schema: runtimeConfig.schema,
      content: runtimeConfig,
    },
    {
      path: 'extension-manifest-fragment.json',
      media_type: 'application/json',
      schema: 'browser_extension_manifest_fragment',
      content: runtimeConfig.extension,
    },
    {
      path: 'verification-plan.json',
      media_type: 'application/json',
      schema: plan.schema,
      content: plan,
    },
    {
      path: 'integration-readme.md',
      media_type: 'text/markdown',
      schema: 'meeting_app_adapter_handoff_readme',
      content: integrationReadme({ spec, runtimeConfig, manifest }),
    },
  ].filter(Boolean);
  const issues = packageIssues({ spec, runtimeConfig, files });
  const blocking = issues.filter((item) => item.severity === 'error');
  const nextActions = unique([
    ...(spec.next_actions ?? []),
    ...(runtimeConfig.next_actions ?? []),
    'capture_live_dom_snapshots_before_production_rollout',
    'verify_speaker_track_on_real_meeting',
    'verify_annotation_insert_on_current_axis',
  ]);
  return {
    type: 'meeting_app_adapter_handoff_package',
    schema: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    adapter_key: spec.adapter_key,
    display_name: spec.display_name,
    source: spec.source,
    surface: spec.surface,
    artifact_count: files.length,
    file_count: files.length,
    files,
    file_paths: files.map((file) => file.path),
    adapter_spec: spec,
    adapter_manifest: manifest,
    runtime_config: runtimeConfig,
    extension_manifest_fragment: runtimeConfig.extension,
    verification_plan: plan,
    consumer_entrypoints: runtimeConfig.entrypoints,
    contracts: {
      timestamp_field: runtimeConfig.contracts?.timestamp_field,
      provider_events_block_realtime: runtimeConfig.contracts?.provider_events_block_realtime,
      transcript_blocks_realtime: runtimeConfig.contracts?.transcript_blocks_realtime,
      local_observer_may_start_axis: runtimeConfig.contracts?.local_observer_may_start_axis,
      required_signals: runtimeConfig.contracts?.required_signals ?? [],
      output_intents: runtimeConfig.contracts?.output_intents ?? [],
    },
    validation: {
      static_ready: blocking.length === 0,
      spec_accepted: spec.accepted === true,
      runtime_config_accepted: runtimeConfig.accepted === true,
      capture_ready: runtimeConfig.readiness?.capture_options_ready === true,
      mutation_ready: runtimeConfig.readiness?.mutation_observer_ready === true,
      content_script_ready: runtimeConfig.readiness?.content_script_ready === true,
      requires_live_snapshot_before_production: true,
      required_live_evidence: [
        'live_dom_snapshot',
        'candidate_observation',
        'speaker_track',
        'participant_track',
        'annotation_insert_current_axis',
      ],
    },
    implementation_steps: [
      'load_runtime_config_json_in_host_project',
      'install_content_script_bridge_with_content_script_options',
      'create_browser_runtime_with_browser_runtime_options',
      'capture_live_snapshots_with_capture_options',
      'run_handoff_readiness_with_live_snapshot_evidence',
    ],
    issue_count: issues.length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: nextActions,
  };
}

export function buildMeetingAppAdapterHandoffPackageMatrix(options = {}) {
  const packageOptions = cleanOptions(options);
  const packages = packageInputs(options).map((input) => buildMeetingAppAdapterHandoffPackage(input, packageOptions));
  return {
    type: 'meeting_app_adapter_handoff_package_matrix',
    schema: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted: packages.every((pkg) => pkg.accepted === true),
    package_count: packages.length,
    accepted_count: packages.filter((pkg) => pkg.accepted === true).length,
    custom_count: packages.filter((pkg) => pkg.source !== 'built_in_manifest').length,
    built_in_count: packages.filter((pkg) => pkg.source === 'built_in_manifest').length,
    runtime_config_ready_count: packages.filter((pkg) => pkg.validation.runtime_config_accepted).length,
    content_script_ready_count: packages.filter((pkg) => pkg.validation.content_script_ready).length,
    live_evidence_required_count: packages.filter((pkg) => pkg.validation.requires_live_snapshot_before_production).length,
    rows: packages.map((pkg) => ({
      adapter_key: pkg.adapter_key,
      display_name: pkg.display_name,
      source: pkg.source,
      accepted: pkg.accepted,
      file_count: pkg.file_count,
      runtime_config_accepted: pkg.validation.runtime_config_accepted,
      content_script_ready: pkg.validation.content_script_ready,
      capture_ready: pkg.validation.capture_ready,
      mutation_ready: pkg.validation.mutation_ready,
      blocking_count: pkg.blocking_count,
      warning_count: pkg.warning_count,
      first_next_action: pkg.next_actions?.[0],
    })),
    packages,
    next_actions: unique(packages.flatMap((pkg) => pkg.next_actions ?? [])),
  };
}

export function buildMeetingAppAdapterVerificationReport(packageOrSpec = {}, options = {}) {
  const handoffPackage = packageOrSpec?.schema === MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA
    ? packageOrSpec
    : buildMeetingAppAdapterHandoffPackage(packageOrSpec, options);
  const evidence = evidenceSource(options);
  const target = firstNonEmpty(
    options.target,
    options.acceptance_target,
    options.acceptanceTarget,
    options.requireProductionReady === false || options.require_production_ready === false ? 'pilot' : 'production',
  );
  const plan = handoffPackage.verification_plan ?? verificationPlan({
    spec: handoffPackage.adapter_spec,
    runtimeConfig: handoffPackage.runtime_config,
    manifest: handoffPackage.adapter_manifest,
  });
  const checks = plan.acceptance_checks.map((check) => {
    if (!String(check.id).startsWith('verify_')) return check;
    const evidenceId = String(check.id).replace(/^verify_/, '');
    const required = plan.required_evidence.find((item) => item.id === evidenceId) ?? {};
    const count = countEvidence(evidence, evidenceId);
    const minimum = Number(required.minimum_count ?? 1);
    return {
      ...check,
      evidence_id: evidenceId,
      observed_count: count,
      minimum_count: minimum,
      status: count >= minimum ? 'passed' : 'missing_evidence',
    };
  });
  const staticChecks = checks.filter((check) => !String(check.id).startsWith('verify_'));
  const liveChecks = checks.filter((check) => String(check.id).startsWith('verify_'));
  const staticPassed = staticChecks.every((check) => check.status === 'passed');
  const livePassed = liveChecks.every((check) => check.status === 'passed');
  const pilotReady = staticPassed && livePassed;
  const productionReady = pilotReady;
  const accepted = target === 'pilot' ? pilotReady : productionReady;
  const issues = [
    ...staticChecks
      .filter((check) => check.blocking && check.status !== 'passed')
      .map((check) => issue('error', check.id, 'Static adapter verification check failed.', {
        adapter_key: handoffPackage.adapter_key,
        check_status: check.status,
      })),
    ...liveChecks
      .filter((check) => check.blocking && check.status !== 'passed')
      .map((check) => issue('error', check.id, 'Required live evidence is missing for adapter verification.', {
        adapter_key: handoffPackage.adapter_key,
        evidence_id: check.evidence_id,
        observed_count: check.observed_count,
        minimum_count: check.minimum_count,
      })),
  ];
  return {
    type: 'meeting_app_adapter_verification_report',
    schema: MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted,
    target,
    adapter_key: handoffPackage.adapter_key,
    display_name: handoffPackage.display_name,
    static_ready: staticPassed,
    live_evidence_ready: livePassed,
    pilot_ready: pilotReady,
    production_ready: productionReady,
    check_count: checks.length,
    passed_check_count: checks.filter((check) => check.status === 'passed').length,
    missing_evidence_count: liveChecks.filter((check) => check.status !== 'passed').length,
    checks,
    evidence_summary: {
      live_dom_snapshot: countEvidence(evidence, 'live_dom_snapshot'),
      candidate_observation: countEvidence(evidence, 'candidate_observation'),
      speaker_track: countEvidence(evidence, 'speaker_track'),
      participant_track: countEvidence(evidence, 'participant_track'),
      annotation_insert_current_axis: countEvidence(evidence, 'annotation_insert_current_axis'),
    },
    handoff_package: {
      schema: handoffPackage.schema,
      accepted: handoffPackage.accepted,
      file_paths: handoffPackage.file_paths,
    },
    verification_plan: plan,
    issue_count: issues.length,
    blocking_count: issues.filter((item) => item.severity === 'error').length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: accepted
      ? ['handoff_adapter_to_host_project']
      : unique([
        ...issues.map((item) => item.code),
        ...plan.next_actions,
      ]),
  };
}

export function buildMeetingAppAdapterVerificationReportMatrix(options = {}) {
  const evidenceByAdapter = options.evidence_by_adapter ?? options.evidenceByAdapter ?? {};
  const reports = packageOrReportInputs(options).map((input) => {
    if (input?.schema === MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA) return input;
    const key = typeof input === 'string' ? input : firstNonEmpty(input.adapter_key, input.platform, input.provider, input.key);
    const evidence = firstNonEmpty(evidenceByAdapter[key], evidenceByAdapter[String(key ?? '').replace(/-/g, '_')], evidenceSource(options));
    return buildMeetingAppAdapterVerificationReport(input, {
      ...options,
      evidence,
      reports: undefined,
    });
  });
  return {
    type: 'meeting_app_adapter_verification_report_matrix',
    schema: MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted: reports.every((report) => report.accepted === true),
    target: firstNonEmpty(options.target, options.acceptance_target, options.acceptanceTarget, 'production'),
    report_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted === true).length,
    pilot_ready_count: reports.filter((report) => report.pilot_ready === true).length,
    production_ready_count: reports.filter((report) => report.production_ready === true).length,
    missing_evidence_count: reports.reduce((total, report) => total + report.missing_evidence_count, 0),
    rows: reports.map((report) => ({
      adapter_key: report.adapter_key,
      accepted: report.accepted,
      target: report.target,
      static_ready: report.static_ready,
      live_evidence_ready: report.live_evidence_ready,
      pilot_ready: report.pilot_ready,
      production_ready: report.production_ready,
      missing_evidence_count: report.missing_evidence_count,
      first_next_action: report.next_actions?.[0],
    })),
    reports,
    next_actions: unique(reports.flatMap((report) => report.next_actions ?? [])),
  };
}

export function assertMeetingAppAdapterHandoffPackage(packageOrSpec = {}, options = {}) {
  const handoffPackage = packageOrSpec?.schema === MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA
    ? packageOrSpec
    : buildMeetingAppAdapterHandoffPackage(packageOrSpec, options);
  if (handoffPackage.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter handoff package acceptance failed', {
      code: 'meeting_app_adapter_handoff_package_rejected',
      adapter_key: handoffPackage.adapter_key,
      issues: handoffPackage.issues,
      next_actions: handoffPackage.next_actions,
    });
  }
  return handoffPackage;
}

export function assertMeetingAppAdapterVerificationReport(reportOrPackage = {}, options = {}) {
  const report = reportOrPackage?.schema === MEETING_APP_ADAPTER_VERIFICATION_REPORT_SCHEMA
    ? reportOrPackage
    : buildMeetingAppAdapterVerificationReport(reportOrPackage, options);
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter verification failed', {
      code: 'meeting_app_adapter_verification_failed',
      adapter_key: report.adapter_key,
      issues: report.issues,
      next_actions: report.next_actions,
    });
  }
  return report;
}

export function assertMeetingAppAdapterVerificationReportMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_VERIFICATION_REPORT_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterVerificationReportMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter verification matrix failed', {
      code: 'meeting_app_adapter_verification_matrix_failed',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}

export function assertMeetingAppAdapterHandoffPackageMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterHandoffPackageMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter handoff package matrix acceptance failed', {
      code: 'meeting_app_adapter_handoff_package_matrix_rejected',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}
