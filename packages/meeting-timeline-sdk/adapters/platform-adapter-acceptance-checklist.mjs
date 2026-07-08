import { compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterPortfolio,
  buildMeetingPlatformAdapterPortfolioItem,
} from './platform-adapter-portfolio.mjs';
import {
  buildMeetingPlatformHandoffReadiness,
} from './platform-handoff-readiness.mjs';

export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA = 'meeting_platform_adapter_acceptance_checklist';
export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_MATRIX_SCHEMA = 'meeting_platform_adapter_acceptance_checklist_matrix';
export const MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA_VERSION = 1;

const TARGET_RANK = Object.freeze({
  static: 0,
  pilot: 1,
  production: 2,
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

function targetFrom(options = {}) {
  const raw = String(firstNonEmpty(
    options.target,
    options.acceptanceTarget,
    options.acceptance_target,
    options.requireProductionReady === true || options.require_production_ready === true ? 'production' : undefined,
    'pilot',
  ));
  return TARGET_RANK[raw] == null ? 'pilot' : raw;
}

function requiredForTarget(requiredFrom, target) {
  return TARGET_RANK[target] >= TARGET_RANK[requiredFrom];
}

function checklistItem(target, input = {}) {
  const requiredFrom = input.required_from ?? 'static';
  const required = requiredForTarget(requiredFrom, target);
  const passed = input.passed === true;
  return compactObject({
    id: input.id,
    group: input.group,
    label: input.label,
    required_from: requiredFrom,
    required,
    passed,
    status: passed ? 'pass' : required ? 'fail' : 'skip',
    blocking: required && !passed,
    evidence: input.evidence,
    expected: input.expected,
    actual: input.actual,
    next_action: input.next_action,
  });
}

function safeHandoffReadiness(portfolioItem = {}, input = {}, options = {}) {
  if (portfolioItem.built_in !== true) return undefined;
  try {
    return buildMeetingPlatformHandoffReadiness(portfolioItem.platform, input, options);
  } catch {
    return undefined;
  }
}

function evidenceCount(readiness = {}, key) {
  return Number(readiness.evidence_counts?.[key] ?? 0);
}

function itemsFor(target, portfolioItem = {}, readiness = {}) {
  const p0 = portfolioItem.p0_realtime_axis ?? {};
  const p1 = portfolioItem.p1_provider_reconcile ?? {};
  const implementation = portfolioItem.implementation ?? {};
  const missing = readiness.missing ?? {};
  return [
    checklistItem(target, {
      id: 'adapter_registered_or_authorable',
      group: 'static_contract',
      label: 'Platform adapter is registered or has an authoring plan',
      required_from: 'static',
      passed: portfolioItem.built_in === true,
      evidence: portfolioItem.adapter_status,
      expected: 'built_in_adapter_available',
      actual: portfolioItem.adapter_status,
      next_action: portfolioItem.built_in ? undefined : 'add_platform_setup_entry',
    }),
    checklistItem(target, {
      id: 'p0_local_axis_first',
      group: 'p0_realtime_axis',
      label: 'Realtime annotations use a local axis before provider reconcile',
      required_from: 'static',
      passed: p0.local_axis_first === true,
      expected: true,
      actual: p0.local_axis_first,
      next_action: 'fix_adapter_route_local_axis_first',
    }),
    checklistItem(target, {
      id: 'captured_at_ms_contract',
      group: 'p0_realtime_axis',
      label: 'Realtime annotation timestamp field is captured_at_ms',
      required_from: 'static',
      passed: p0.timestamp_field === 'captured_at_ms',
      expected: 'captured_at_ms',
      actual: p0.timestamp_field,
      next_action: 'fix_annotation_timestamp_contract',
    }),
    checklistItem(target, {
      id: 'provider_nonblocking_contract',
      group: 'p0_realtime_axis',
      label: 'Provider events never block realtime annotation insertion',
      required_from: 'static',
      passed: p0.provider_events_block_realtime === false,
      expected: false,
      actual: p0.provider_events_block_realtime,
      next_action: 'set_provider_events_block_realtime_false',
    }),
    checklistItem(target, {
      id: 'transcript_nonblocking_contract',
      group: 'p2_post_meeting_backfill',
      label: 'Transcript and recording artifacts are post-meeting backfill only',
      required_from: 'static',
      passed: p0.transcript_blocks_realtime === false,
      expected: false,
      actual: p0.transcript_blocks_realtime,
      next_action: 'set_transcript_blocks_realtime_false',
    }),
    checklistItem(target, {
      id: 'candidate_observation_contract',
      group: 'p0_realtime_axis',
      label: 'Host can observe current meeting candidates',
      required_from: 'static',
      passed: p0.candidate_observation_message_type === 'meeting_timeline.observe_candidates',
      expected: 'meeting_timeline.observe_candidates',
      actual: p0.candidate_observation_message_type,
      next_action: 'wire_observe_platform_candidates',
    }),
    checklistItem(target, {
      id: 'provider_reconcile_path_declared',
      group: 'p1_provider_reconcile',
      label: 'Provider reconcile path is declared',
      required_from: 'static',
      passed: Boolean(p1.path),
      evidence: p1.path,
      next_action: 'declare_provider_reconcile_path',
    }),
    checklistItem(target, {
      id: 'implementation_handoff_available',
      group: 'static_contract',
      label: 'Implementation handoff is available for the platform',
      required_from: 'static',
      passed: implementation.implementation_ready === true,
      expected: true,
      actual: implementation.implementation_ready,
      next_action: 'run_meeting_platform_implementation_handoff',
    }),
    checklistItem(target, {
      id: 'browser_or_native_surface_declared',
      group: 'p0_realtime_axis',
      label: 'A first local observation surface is declared',
      required_from: 'static',
      passed: Boolean(portfolioItem.recommended_first_surface),
      evidence: portfolioItem.recommended_first_surface,
      next_action: 'declare_local_observation_surface',
    }),
    checklistItem(target, {
      id: 'real_local_observer_evidence',
      group: 'pilot_evidence',
      label: 'Real meeting local observation evidence is available',
      required_from: 'pilot',
      passed: readiness.local_observer_ready === true,
      evidence: { dom_records: evidenceCount(readiness, 'dom_records') },
      next_action: 'capture_real_meeting_app_snapshots',
    }),
    checklistItem(target, {
      id: 'handoff_ready',
      group: 'pilot_evidence',
      label: 'SDK handoff readiness gate passes',
      required_from: 'pilot',
      passed: readiness.handoff_ready === true,
      evidence: readiness.status,
      next_action: readiness.next_actions?.[0] ?? 'run_meeting_platform_handoff_readiness',
    }),
    checklistItem(target, {
      id: 'provider_security_ready',
      group: 'production_evidence',
      label: 'Provider security credentials and webhook verification are configured',
      required_from: 'production',
      passed: (missing.provider_env?.length ?? 0) === 0,
      evidence: missing.provider_env ?? [],
      next_action: 'configure_provider_security_env',
    }),
    checklistItem(target, {
      id: 'provider_reconcile_evidence',
      group: 'production_evidence',
      label: 'Provider start/end reconcile evidence is available',
      required_from: 'production',
      passed: readiness.provider_reconcile_ready === true,
      evidence: { provider_records: evidenceCount(readiness, 'provider_records') },
      next_action: 'capture_real_provider_events_and_build_evidence_package',
    }),
    checklistItem(target, {
      id: 'runtime_host_replay_ready',
      group: 'production_evidence',
      label: 'Runtime host replay gate passes when required',
      required_from: 'production',
      passed: readiness.runtime_host_replay_required !== true || readiness.runtime_host_replay_accepted === true,
      evidence: {
        required: readiness.runtime_host_replay_required === true,
        accepted: readiness.runtime_host_replay_accepted,
        missing: missing.runtime_host_replay_missing ?? [],
      },
      next_action: 'run_meeting_platform_runtime_host_replay',
    }),
    checklistItem(target, {
      id: 'production_ready',
      group: 'production_evidence',
      label: 'Production evidence gate passes',
      required_from: 'production',
      passed: readiness.production_ready === true,
      evidence: readiness.status,
      next_action: readiness.next_actions?.[0] ?? 'run_meeting_platform_real_intake',
    }),
  ];
}

function summarize(items = []) {
  const requiredItems = items.filter((item) => item.required);
  const blockingItems = items.filter((item) => item.blocking);
  return {
    item_count: items.length,
    required_count: requiredItems.length,
    passed_count: items.filter((item) => item.passed).length,
    skipped_count: items.filter((item) => item.status === 'skip').length,
    blocking_count: blockingItems.length,
    failed_required_ids: blockingItems.map((item) => item.id),
  };
}

function surfaceInstallStepId(surface) {
  if (surface === 'browser_extension' || surface === 'webview_preload') {
    return 'install_browser_extension_or_webview_preload';
  }
  if (surface === 'native_detector') return 'wire_native_window_or_accessibility_detector';
  if (surface === 'provider_reconcile') return 'wire_provider_reconcile_as_backfill_only';
  return 'wire_selected_local_observer_surface';
}

function surfaceInstallExpected(surface) {
  if (surface === 'browser_extension') return 'content_script_can_observe_current_meeting_window';
  if (surface === 'webview_preload') return 'preload_bridge_can_observe_embedded_meeting_view';
  if (surface === 'native_detector') return 'native_window_or_accessibility_detector_can_observe_meeting_state';
  if (surface === 'provider_reconcile') return 'provider_events_are_backfill_only_and_do_not_block_realtime_marks';
  return 'local_observer_surface_can_emit_meeting_candidates';
}

function buildRuntimeEventContract(portfolioItem = {}) {
  const p0 = portfolioItem.p0_realtime_axis ?? {};
  const p1 = portfolioItem.p1_provider_reconcile ?? {};
  return compactObject({
    first_runtime_action: 'observe_platform_candidates',
    observe_candidates_message_type: p0.candidate_observation_message_type ?? 'meeting_timeline.observe_candidates',
    mark_runtime_action: 'insert_annotation',
    insert_mark_message_type: 'meeting_timeline.insert_mark',
    observe_before_insert_required: true,
    timestamp_field: p0.timestamp_field ?? 'captured_at_ms',
    time_origin: 'local_meeting_axis_start_ms',
    per_meeting_axis_required: true,
    provider_reconcile_required_for_realtime: p1.required_for_realtime === true,
    provider_events_block_realtime: p0.provider_events_block_realtime === true,
    transcript_blocks_realtime: p0.transcript_blocks_realtime === true,
    speaker_position_markers: {
      transcript_text_required: false,
      filter_policy: 'merge_same_speaker_short_gaps_and_drop_sub_threshold_blips',
      timeline_role: 'speaker_position_only',
    },
    required_payload_fields: ['platform', 'meeting_id', 'captured_at_ms'],
  });
}

function buildSdkEntrypoints(platform) {
  return {
    checklist: `sdk.platformAdapterAcceptanceChecklist('${platform}')`,
    matrix_row: `sdk.platformAdapterMatrixRow('${platform}')`,
    host_config: `sdk.connectorHostAdapterConfig('${platform}')`,
    bootstrap_plan: 'sdk.connectorHostAdapterBootstrapPlan(input, connectorPackage)',
    observe_candidates: 'sdk.observePlatformCandidates(input)',
    insert_annotation: `sdk.insertAnnotation('${platform}', mark)`,
    speaker_position: `sdk.speakerTrack('${platform}', sample)`,
    provider_reconcile: `sdk.providerReplayReport('${platform}', providerEvents)`,
  };
}

function buildImplementationSequence(platform, portfolioItem = {}) {
  const p0 = portfolioItem.p0_realtime_axis ?? {};
  const p1 = portfolioItem.p1_provider_reconcile ?? {};
  const p2 = portfolioItem.p2_post_meeting_backfill ?? {};
  const surface = portfolioItem.recommended_first_surface;
  const browserMatches = asArray(p0.browser_matches);
  return [
    compactObject({
      order: 1,
      id: 'resolve_adapter_checklist',
      required: true,
      owner: 'host_project',
      sdk_entrypoint: `sdk.platformAdapterAcceptanceChecklist('${platform}')`,
      output: 'platform_acceptance_and_runtime_contract',
    }),
    compactObject({
      order: 2,
      id: surfaceInstallStepId(surface),
      group: 'local_observer_install',
      required: true,
      surface,
      expected: surfaceInstallExpected(surface),
      browser_matches: browserMatches.length ? browserMatches : undefined,
    }),
    compactObject({
      order: 3,
      id: 'run_adapter_preflight_before_first_mark',
      group: 'runtime_gate',
      required: true,
      expected: 'current_meeting_candidate_is_observable_before_insert',
      evidence: 'current_window_or_candidate_snapshot',
    }),
    compactObject({
      order: 4,
      id: 'observe_platform_candidates_before_insert',
      group: 'runtime_event',
      required: true,
      runtime_action: 'observe_platform_candidates',
      message_type: p0.candidate_observation_message_type ?? 'meeting_timeline.observe_candidates',
      output: 'active_meeting_axis_candidate',
    }),
    compactObject({
      order: 5,
      id: 'insert_annotation_with_captured_at_ms',
      group: 'runtime_event',
      required: true,
      runtime_action: 'insert_annotation',
      message_type: 'meeting_timeline.insert_mark',
      timestamp_field: p0.timestamp_field ?? 'captured_at_ms',
      expected: 'mark_is_written_to_the_current_local_axis',
    }),
    compactObject({
      order: 6,
      id: 'emit_speaker_position_markers_after_filtering',
      group: 'runtime_event',
      required: true,
      runtime_action: 'speaker_track',
      transcript_text_required: false,
      filter_policy: 'merge_same_speaker_short_gaps_and_drop_sub_threshold_blips',
      expected: 'speaker_position_markers_without_transcript_text_dependency',
    }),
    compactObject({
      order: 7,
      id: 'ingest_provider_events_for_nonblocking_reconcile',
      group: 'provider_reconcile',
      required_for: 'production',
      required: Boolean(p1.path),
      provider_path: p1.path,
      transport: p1.transport,
      blocks_realtime: false,
      expected: p1.path
        ? 'provider_start_end_events_reconcile_after_local_axis_exists'
        : 'declare_provider_reconcile_path_or_keep_reconcile_unavailable',
    }),
    compactObject({
      order: 8,
      id: 'import_post_meeting_artifacts_after_meeting',
      group: 'post_meeting_backfill',
      required: false,
      source: p2.source,
      blocks_realtime: false,
      expected: 'transcript_or_recording_artifacts_never_gate_realtime_marks',
    }),
  ];
}

function buildEvidenceCollectionPlan(target, portfolioItem = {}) {
  const platform = portfolioItem.platform;
  return compactObject({
    current_target: target,
    static: [
      'adapter_registered_or_authorable',
      'captured_at_ms_contract',
      'provider_nonblocking_contract',
      'candidate_observation_contract',
      'browser_or_native_surface_declared',
    ],
    pilot: [
      'real_meeting_local_observer_snapshot',
      'candidate_observation_smoke',
      'insert_annotation_current_axis',
      'speaker_position_marker_smoke',
    ],
    production: [
      'provider_start_end_reconcile_evidence',
      'provider_security_env',
      'runtime_host_replay',
      'post_meeting_transcript_backfill_sample',
    ],
    commands: compactObject({
      static: `npm run meeting-platform:adapter-acceptance-checklist -- --platforms=${platform} --target=static`,
      pilot: `npm run meeting-platform:adapter-acceptance-checklist -- --platforms=${platform} --target=pilot`,
      production: `npm run meeting-platform:adapter-acceptance-checklist -- --platforms=${platform} --target=production`,
    }),
  });
}

export function buildMeetingPlatformAdapterAcceptanceChecklist(platform, input = {}, options = {}) {
  const target = targetFrom(options);
  const portfolioItem = buildMeetingPlatformAdapterPortfolioItem(platform, options);
  const readiness = safeHandoffReadiness(portfolioItem, input, {
    ...options,
    target,
    requireProductionReady: target === 'production' || options.requireProductionReady === true,
    require_production_ready: target === 'production' || options.require_production_ready === true,
  });
  const items = itemsFor(target, portfolioItem, readiness);
  const summary = summarize(items);
  return {
    type: 'meeting_platform_adapter_acceptance_checklist',
    schema: MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA_VERSION,
    platform: portfolioItem.platform,
    display_name: portfolioItem.display_name,
    target,
    accepted: summary.blocking_count === 0,
    adapter_status: portfolioItem.adapter_status,
    recommended_first_surface: portfolioItem.recommended_first_surface,
    runtime_event_contract: buildRuntimeEventContract(portfolioItem),
    implementation_sequence: buildImplementationSequence(portfolioItem.platform, portfolioItem),
    sdk_entrypoints: buildSdkEntrypoints(portfolioItem.platform),
    evidence_collection_plan: buildEvidenceCollectionPlan(target, portfolioItem),
    checklist: items,
    summary,
    portfolio_item: portfolioItem,
    handoff_readiness: readiness,
    commands: compactObject({
      adapter_portfolio: `npm run meeting-platform:adapter-portfolio -- --platforms=${portfolioItem.platform}`,
      handoff_readiness: portfolioItem.built_in ? `npm run meeting-platform:handoff-readiness -- --platforms=${portfolioItem.platform}` : undefined,
      real_intake: portfolioItem.built_in ? `npm run meeting-platform:real-intake -- --platforms=${portfolioItem.platform}` : undefined,
      field_capture: portfolioItem.built_in ? `npm run meeting-platform:field-capture -- --platforms=${portfolioItem.platform}` : undefined,
    }),
    next_actions: unique([
      ...items.filter((item) => item.blocking).map((item) => item.next_action),
      ...(readiness?.next_actions ?? []),
      ...(portfolioItem.next_actions ?? []),
    ]),
  };
}

export function buildMeetingPlatformAdapterAcceptanceChecklistMatrix(input = {}, options = {}) {
  const target = targetFrom({ ...input, ...options });
  const portfolio = buildMeetingPlatformAdapterPortfolio({
    ...options,
    platforms: firstNonEmpty(input.platforms, input.platform_keys, options.platforms, options.platform_keys),
  });
  const items = portfolio.platforms.map((platform) => buildMeetingPlatformAdapterAcceptanceChecklist(
    platform,
    input[platform] ?? input.platforms?.[platform] ?? input.byPlatform?.[platform] ?? input.by_platform?.[platform] ?? input,
    {
      ...options,
      target,
      platforms: undefined,
      platform_keys: undefined,
    },
  ));
  return {
    type: 'meeting_platform_adapter_acceptance_checklist_matrix',
    schema: MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_ACCEPTANCE_CHECKLIST_SCHEMA_VERSION,
    target,
    platform_count: items.length,
    accepted_count: items.filter((item) => item.accepted).length,
    blocked_count: items.filter((item) => !item.accepted).length,
    static_ready_count: items.filter((item) => item.summary.failed_required_ids.includes('adapter_registered_or_authorable') === false).length,
    pilot_ready_count: items.filter((item) => item.handoff_readiness?.pilot_ready === true).length,
    production_ready_count: items.filter((item) => item.handoff_readiness?.production_ready === true).length,
    platforms: items.map((item) => item.platform),
    rows: items.map((item) => ({
      platform: item.platform,
      display_name: item.display_name,
      target: item.target,
      accepted: item.accepted,
      adapter_status: item.adapter_status,
      recommended_first_surface: item.recommended_first_surface,
      required_count: item.summary.required_count,
      passed_count: item.summary.passed_count,
      blocking_count: item.summary.blocking_count,
      failed_required_ids: item.summary.failed_required_ids,
      handoff_status: item.handoff_readiness?.status,
      pilot_ready: item.handoff_readiness?.pilot_ready === true,
      production_ready: item.handoff_readiness?.production_ready === true,
      runtime_contract_timestamp_field: item.runtime_event_contract?.timestamp_field,
      provider_reconcile_nonblocking: item.runtime_event_contract?.provider_events_block_realtime === false,
      first_implementation_step: item.implementation_sequence?.[0]?.id,
      local_observer_install_step: item.implementation_sequence?.find((step) => step.group === 'local_observer_install')?.id,
      first_next_action: item.next_actions?.[0],
    })),
    checklists: items,
    next_actions: unique(items.flatMap((item) => item.next_actions ?? [])),
  };
}
