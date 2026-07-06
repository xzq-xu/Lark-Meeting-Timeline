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
      first_next_action: item.next_actions?.[0],
    })),
    checklists: items,
    next_actions: unique(items.flatMap((item) => item.next_actions ?? [])),
  };
}
