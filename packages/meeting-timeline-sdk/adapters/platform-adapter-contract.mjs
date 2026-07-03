import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import { buildMeetingPlatformRuntimeProfile } from './platform-runtime-profile.mjs';
import { buildMeetingPlatformFieldCollectorConfig } from './platform-field-capture.mjs';

export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA = 'meeting_platform_adapter_contract';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA = 'meeting_platform_adapter_contract_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_SCHEMA = 'meeting_platform_adapter_contract_acceptance';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_MATRIX_SCHEMA = 'meeting_platform_adapter_contract_acceptance_matrix';
export const MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION = 1;

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

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function providerEvents(runtime = {}) {
  return {
    start_events: runtime.provider_events?.start_events ?? [],
    end_events: runtime.provider_events?.end_events ?? [],
    participant_events: runtime.provider_events?.participant_events ?? [],
    artifact_events: runtime.provider_events?.artifact_events ?? [],
    lifecycle_events: runtime.provider_events?.lifecycle_events ?? [],
  };
}

function localObserverContract(collector = {}, runtime = {}) {
  const observer = collector.browser_observer;
  const snapshots = collector.local_snapshot_collector;
  if (!observer && !snapshots && !runtime.speaker_markers) return undefined;
  return compactObject({
    required_for_realtime_axis: true,
    surface: observer?.type ?? 'trusted_host_or_native_observer',
    matches: observer?.matches ?? [],
    host_permissions: observer?.host_permissions ?? [],
    content_scripts: observer?.content_scripts ?? [],
    snapshot_collector: snapshots,
    speaker_markers: runtime.speaker_markers,
  });
}

function providerObserverContract(provider = {}, runtime = {}, collector = {}) {
  const events = providerEvents(runtime);
  return compactObject({
    role: provider.provider_role,
    transport: provider.transport,
    endpoint: provider.endpoint,
    status_endpoint: provider.status_endpoint,
    required_for_realtime: false,
    required_for_production_evidence: runtime.axis?.provider_reconcile?.required_for_production_evidence,
    event_mapping: provider.event_mapping ?? [],
    events,
    security: provider.security,
    readiness: provider.readiness,
    collector_requirements: collector.provider_observer,
  });
}

function evidenceContract(collector = {}) {
  const manifest = collector.manifest ?? {};
  const acceptance = collector.acceptance ?? {};
  return compactObject({
    status: manifest.status,
    production_ready: manifest.production_ready === true,
    ready_for_realtime_annotations: manifest.ready_for_realtime_annotations === true,
    files: collector.storage?.files,
    input_contract: collector.storage?.input_contract,
    required_provider_coverage: acceptance.required_provider_coverage ?? [],
    required_local_snapshots: acceptance.required_local_snapshots ?? [],
    minimum_provider_records: acceptance.minimum_provider_records ?? 0,
    minimum_local_records: acceptance.minimum_local_records ?? 0,
    missing_items: acceptance.current_missing_items ?? [],
    production_condition: acceptance.production_condition,
    pilot_condition: acceptance.pilot_condition,
  });
}

function implementationContract(capabilities = {}, collector = {}, runtime = {}) {
  return {
    imports: compactObject({
      client: '@ai-annotation/meeting-timeline-sdk',
      kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
      contract: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract',
      events: capabilities.sdk_modules?.events,
      ingest: capabilities.sdk_modules?.ingest,
      local_observer: capabilities.sdk_modules?.local_observer,
      transcript: capabilities.sdk_modules?.transcript,
      security: capabilities.sdk_modules?.security,
    }),
    kit_methods: [
      'platformAdapterContract',
      'platformRuntimeProfile',
      'platformProviderConnectionPack',
      'platformFieldCollectorConfig',
      'platformFieldEvidenceBundle',
      'insertMark',
      'importTranscript',
    ],
    commands: collector.automation?.commands ?? {},
    order: runtime.implementation_order ?? [],
  };
}

function readinessContract(provider = {}, collector = {}, runtime = {}) {
  const evidence = evidenceContract(collector);
  return {
    status: evidence.status,
    production_ready: evidence.production_ready,
    ready_for_realtime_annotations: evidence.ready_for_realtime_annotations,
    provider_ready: provider.readiness?.ready === true,
    provider_missing_env: provider.security?.missing_env ?? [],
    missing_items: evidence.missing_items ?? [],
    risks: runtime.risks ?? [],
  };
}

function targetFromOptions(options = {}) {
  return String(firstNonEmpty(
    options.target,
    options.acceptanceTarget,
    options.acceptance_target,
    options.requireProductionReady === true || options.require_production_ready === true ? 'production' : undefined,
    options.requireRealtimeReady === true || options.require_realtime_ready === true ? 'pilot' : undefined,
    'contract',
  ));
}

function contractFromInput(contractOrPlatform, options = {}) {
  if (contractOrPlatform?.schema === MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA) return contractOrPlatform;
  return buildMeetingPlatformAdapterContract(contractOrPlatform, options);
}

function hasValue(value) {
  return value != null && value !== '';
}

function hasArrayItems(value) {
  return Array.isArray(value) && value.length > 0;
}

function includesValue(value, expected) {
  return Array.isArray(value) && value.includes(expected);
}

function buildAcceptanceIssues(contract = {}, target = 'contract') {
  const issues = [];
  if (contract.schema !== MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA) {
    issues.push(issue('error', 'invalid_schema', 'Contract schema must be meeting_platform_adapter_contract.', {
      actual_schema: contract.schema,
    }));
  }
  if (!hasValue(contract.platform)) {
    issues.push(issue('error', 'missing_platform', 'Contract must include a normalized platform key.'));
  }
  if (contract.timebase?.annotation_timestamp_field !== 'captured_at_ms') {
    issues.push(issue('error', 'invalid_annotation_timestamp_field', 'Realtime annotations must use captured_at_ms.', {
      actual: contract.timebase?.annotation_timestamp_field,
    }));
  }
  if (contract.timebase?.provider_events_block_realtime !== false) {
    issues.push(issue('error', 'provider_blocks_realtime', 'Provider events must not block realtime annotation insertion.'));
  }
  if (contract.timebase?.transcript_blocks_realtime !== false) {
    issues.push(issue('error', 'transcript_blocks_realtime', 'Post-meeting transcript import must not block realtime annotation insertion.'));
  }
  if (contract.supported_surfaces?.realtime_transcript_required !== false) {
    issues.push(issue('error', 'realtime_transcript_required', 'Meeting timeline annotation must not require realtime transcript.'));
  }
  if (!hasValue(contract.annotations?.endpoints?.insertMark)) {
    issues.push(issue('error', 'missing_insert_mark_endpoint', 'Contract must expose annotations.endpoints.insertMark.'));
  }
  if (contract.annotations?.timestamp_field !== 'captured_at_ms') {
    issues.push(issue('error', 'invalid_annotation_endpoint_timestamp_field', 'Annotation endpoint contract must name captured_at_ms.'));
  }
  if (!includesValue(contract.realtime_axis?.rules, 'insert_annotation_by_captured_at_ms_on_the_active_axis')) {
    issues.push(issue('error', 'missing_realtime_annotation_rule', 'Contract must state annotations are placed by captured_at_ms on the active axis.'));
  }
  if (!includesValue(contract.realtime_axis?.rules, 'use_provider_events_only_for_reconcile_and_backfill')) {
    issues.push(issue('error', 'missing_provider_reconcile_rule', 'Contract must state provider events are reconcile/backfill only.'));
  }
  if (!hasValue(contract.realtime_axis?.start?.create_on)) {
    issues.push(issue('error', 'missing_axis_start_policy', 'Contract must define realtime_axis.start.create_on.'));
  }
  if (!hasValue(contract.realtime_axis?.end?.create_on)) {
    issues.push(issue('error', 'missing_axis_end_policy', 'Contract must define realtime_axis.end.create_on.'));
  }
  if (contract.supported_surfaces?.browser_observer === true) {
    if (!hasArrayItems(contract.local_observer?.matches)) {
      issues.push(issue('error', 'missing_browser_matches', 'Browser-observed platforms must declare URL match patterns.'));
    }
    if (!includesValue(contract.local_observer?.snapshot_collector?.required_snapshots, 'active_speaker')) {
      issues.push(issue('error', 'missing_active_speaker_snapshot', 'Browser observer must require an active_speaker snapshot.'));
    }
    if (!includesValue(contract.local_observer?.snapshot_collector?.required_snapshots, 'meeting_ended')) {
      issues.push(issue('error', 'missing_meeting_ended_snapshot', 'Browser observer must require a meeting_ended snapshot.'));
    }
  }
  if (contract.supported_surfaces?.provider_webhook_or_event_subscription === true) {
    if (contract.provider_observer?.required_for_realtime !== false) {
      issues.push(issue('error', 'provider_required_for_realtime', 'Provider observer must be optional for realtime annotation insertion.'));
    }
    if (!hasArrayItems(contract.provider_observer?.events?.start_events)) {
      issues.push(issue('error', 'missing_provider_start_events', 'Provider observer must list meeting start reconcile events.'));
    }
    if (!hasArrayItems(contract.provider_observer?.events?.end_events)) {
      issues.push(issue('error', 'missing_provider_end_events', 'Provider observer must list meeting end reconcile events.'));
    }
    if (!hasValue(contract.provider_observer?.security?.verifier)) {
      issues.push(issue('error', 'missing_provider_security_verifier', 'Provider observer must declare a security verifier.'));
    }
  }
  if (contract.transcript?.realtime_dependency !== false) {
    issues.push(issue('error', 'transcript_realtime_dependency', 'Transcript must be post-meeting or non-blocking for realtime annotation.'));
  }
  if (!includesValue(contract.implementation?.kit_methods, 'platformAdapterContract')) {
    issues.push(issue('error', 'missing_kit_contract_method', 'Implementation handoff must mention platformAdapterContract.'));
  }
  if (!includesValue(contract.implementation?.kit_methods, 'insertMark')) {
    issues.push(issue('error', 'missing_insert_mark_method', 'Implementation handoff must mention insertMark.'));
  }
  if (target === 'pilot' && contract.readiness?.ready_for_realtime_annotations !== true) {
    issues.push(issue('error', 'pilot_not_ready', 'Pilot acceptance requires ready_for_realtime_annotations=true.', {
      status: contract.readiness?.status,
      missing_items: contract.readiness?.missing_items ?? [],
    }));
  }
  if (target === 'production' && contract.readiness?.production_ready !== true) {
    issues.push(issue('error', 'production_not_ready', 'Production acceptance requires production_ready=true.', {
      status: contract.readiness?.status,
      missing_items: contract.readiness?.missing_items ?? [],
    }));
  }
  return issues;
}

export function buildMeetingPlatformAdapterContract(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capabilities = platformCapabilityContract(key, options);
  const runtime = buildMeetingPlatformRuntimeProfile(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const collector = buildMeetingPlatformFieldCollectorConfig(key, options);
  const readiness = readinessContract(provider, collector, runtime);
  const nextActions = unique([
    ...(collector.next_actions ?? []),
    ...(provider.next_actions ?? []),
    ...(runtime.next_actions ?? []),
  ]);

  return compactObject({
    type: 'meeting_platform_adapter_contract',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? runtime.display_name,
    objective: 'portable_sdk_contract_for_adapting_a_meeting_app_to_the_annotation_timeline',
    mode: collector.mode,
    supported_surfaces: {
      local_observer_or_host_detector: true,
      browser_observer: Boolean(collector.browser_observer),
      provider_webhook_or_event_subscription: key !== 'local_detector',
      post_meeting_transcript_import: capabilities.post_meeting_transcript?.status ?? 'unknown',
      realtime_transcript_required: false,
    },
    timebase: runtime.runtime_contract,
    realtime_axis: {
      primary_source: runtime.axis?.primary_source,
      start: runtime.axis?.start,
      end: runtime.axis?.end,
      provider_reconcile: runtime.axis?.provider_reconcile,
      rules: [
        'create_or_bind_axis_from_local_observer_before_provider_event_arrives',
        'insert_annotation_by_captured_at_ms_on_the_active_axis',
        'use_provider_events_only_for_reconcile_and_backfill',
        'never_wait_for_post_meeting_transcript_before_inserting_realtime_marks',
      ],
    },
    annotations: {
      endpoints: collector.timeline_ingest?.endpoints,
      timestamp_field: runtime.runtime_contract?.annotation_timestamp_field,
      policy: runtime.annotations,
      realtime_policy: collector.timeline_ingest?.realtime_annotation_policy,
    },
    local_observer: localObserverContract(collector, runtime),
    provider_observer: providerObserverContract(provider, runtime, collector),
    transcript: runtime.transcript,
    evidence: evidenceContract(collector),
    implementation: implementationContract(capabilities, collector, runtime),
    readiness,
    next_actions: nextActions,
  });
}

export function buildMeetingPlatformAdapterContractMatrix(options = {}) {
  const contracts = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterContract(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_contract_matrix',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    platform_count: contracts.length,
    browser_observer_count: contracts.filter((contract) => contract.supported_surfaces.browser_observer).length,
    provider_observer_count: contracts.filter((contract) => contract.supported_surfaces.provider_webhook_or_event_subscription).length,
    production_ready_count: contracts.filter((contract) => contract.readiness.production_ready).length,
    realtime_ready_count: contracts.filter((contract) => contract.readiness.ready_for_realtime_annotations).length,
    platforms: contracts.map((contract) => contract.platform),
    rows: contracts.map((contract) => ({
      platform: contract.platform,
      display_name: contract.display_name,
      mode: contract.mode,
      browser_observer: contract.supported_surfaces.browser_observer,
      provider_transport: contract.provider_observer?.transport,
      provider_ready: contract.readiness.provider_ready,
      production_ready: contract.readiness.production_ready,
      ready_for_realtime_annotations: contract.readiness.ready_for_realtime_annotations,
      start_create_on: contract.realtime_axis?.start?.create_on,
      end_create_on: contract.realtime_axis?.end?.create_on,
      provider_start_event_count: contract.provider_observer?.events?.start_events?.length ?? 0,
      provider_end_event_count: contract.provider_observer?.events?.end_events?.length ?? 0,
      speaker_min_stable_ms: contract.local_observer?.speaker_markers?.filter?.min_stable_ms,
      transcript_import: contract.transcript?.availability,
      missing_item_count: contract.readiness.missing_items?.length ?? 0,
    })),
    contracts,
    next_actions: unique(contracts.flatMap((contract) => contract.next_actions ?? [])),
  };
}

export function buildMeetingPlatformAdapterContractAcceptanceReport(contractOrPlatform, options = {}) {
  const contract = contractFromInput(contractOrPlatform, options);
  const target = targetFromOptions(options);
  const issues = buildAcceptanceIssues(contract, target);
  const errorCount = issues.filter((item) => item.severity === 'error').length;
  const warningCount = issues.filter((item) => item.severity === 'warn').length;
  return {
    type: 'meeting_platform_adapter_contract_acceptance',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    platform: contract.platform,
    display_name: contract.display_name,
    target,
    accepted: errorCount === 0,
    issue_count: issues.length,
    error_count: errorCount,
    warning_count: warningCount,
    issues,
    summary: {
      mode: contract.mode,
      browser_observer: contract.supported_surfaces?.browser_observer === true,
      provider_observer: contract.supported_surfaces?.provider_webhook_or_event_subscription === true,
      provider_ready: contract.readiness?.provider_ready === true,
      production_ready: contract.readiness?.production_ready === true,
      ready_for_realtime_annotations: contract.readiness?.ready_for_realtime_annotations === true,
      insert_mark_endpoint: contract.annotations?.endpoints?.insertMark,
      provider_start_event_count: contract.provider_observer?.events?.start_events?.length ?? 0,
      provider_end_event_count: contract.provider_observer?.events?.end_events?.length ?? 0,
      missing_item_count: contract.readiness?.missing_items?.length ?? 0,
    },
    contract,
  };
}

export function buildMeetingPlatformAdapterContractAcceptanceMatrix(options = {}) {
  const reports = selectedPlatforms(options).map((platform) => buildMeetingPlatformAdapterContractAcceptanceReport(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_adapter_contract_acceptance_matrix',
    schema: MEETING_PLATFORM_ADAPTER_CONTRACT_ACCEPTANCE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_CONTRACT_SCHEMA_VERSION,
    target: targetFromOptions(options),
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    rejected_count: reports.filter((report) => !report.accepted).length,
    issue_count: reports.reduce((total, report) => total + report.issue_count, 0),
    error_count: reports.reduce((total, report) => total + report.error_count, 0),
    warning_count: reports.reduce((total, report) => total + report.warning_count, 0),
    platforms: reports.map((report) => report.platform),
    rows: reports.map((report) => ({
      platform: report.platform,
      display_name: report.display_name,
      target: report.target,
      accepted: report.accepted,
      mode: report.summary.mode,
      browser_observer: report.summary.browser_observer,
      provider_observer: report.summary.provider_observer,
      provider_ready: report.summary.provider_ready,
      production_ready: report.summary.production_ready,
      ready_for_realtime_annotations: report.summary.ready_for_realtime_annotations,
      issue_count: report.issue_count,
      error_count: report.error_count,
      first_issue: report.issues[0]?.code,
    })),
    reports,
  };
}

export function assertMeetingPlatformAdapterContract(contractOrPlatform, options = {}) {
  const report = buildMeetingPlatformAdapterContractAcceptanceReport(contractOrPlatform, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform adapter contract acceptance failed', report);
  }
  return report;
}
