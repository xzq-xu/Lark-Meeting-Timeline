import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  buildMeetingAppSnapshotRecord,
  buildMeetingAppSnapshotRecordSet,
  meetingAppSnapshotRecords,
} from './meeting-app-snapshot-recorder.mjs';
import {
  capturePlatformWebhookEvent,
} from './platform-capture.mjs';
import {
  buildMeetingPlatformEvidenceCorrelation,
} from './platform-evidence-correlation.mjs';
import {
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformRolloutPlan,
} from './platform-rollout.mjs';
import {
  buildMeetingPlatformAdapterRoute,
  summarizeMeetingPlatformAdapterRoute,
  verifyMeetingPlatformAdapterRouteReadiness,
} from './platform-adapter-route.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA = 'meeting_platform_evidence_package';
export const MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function normalizeArgs(platformOrInput, inputOrOptions = {}, maybeOptions = {}) {
  if (platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput)) {
    const input = platformOrInput;
    return {
      input,
      platform: firstNonEmpty(input.platform, input.provider, input.adapter, input.rollout_plan?.platform),
      options: inputOrOptions ?? {},
    };
  }
  return {
    input: inputOrOptions ?? {},
    platform: platformOrInput,
    options: maybeOptions ?? {},
  };
}

function matchesPlatform(value, platform) {
  if (!value) return true;
  try {
    return normalizeMeetingPlatform(value) === platform;
  } catch {
    return false;
  }
}

function platformFromRecord(record = {}) {
  return firstNonEmpty(
    record.platform,
    record.provider,
    record.adapter,
    record.snapshot?.platform,
    record.snapshot?.provider,
    record.snapshot?.capture?.profile,
  );
}

function recordsForPlatform(platform, records = []) {
  return asArray(records).filter((record) => matchesPlatform(platformFromRecord(record), platform));
}

function providerRecordsFor(platform, input = {}, options = {}) {
  return recordsForPlatform(platform, [
    ...asArray(input.providerRecords),
    ...asArray(input.provider_records),
    ...asArray(input.providerCaptureRecords),
    ...asArray(input.provider_capture_records),
    ...asArray(input.captureRecords),
    ...asArray(input.capture_records),
    ...asArray(input.provider_record_set?.records),
    ...asArray(options.providerRecords),
    ...asArray(options.provider_records),
    ...asArray(options.providerCaptureRecords),
    ...asArray(options.provider_capture_records),
    ...asArray(options.captureRecords),
    ...asArray(options.capture_records),
  ]);
}

function samplesFromCollection(platform, collection) {
  if (!collection) return [];
  if (Array.isArray(collection)) {
    return collection.filter((sample) => matchesPlatform(platformFromRecord(sample), platform));
  }
  if (typeof collection === 'object') {
    const samples = [];
    for (const [key, value] of Object.entries(collection)) {
      if (matchesPlatform(key, platform)) samples.push(...asArray(value));
    }
    return samples;
  }
  return [];
}

function providerSamplesFor(platform, input = {}, options = {}) {
  return [
    ...samplesFromCollection(platform, input.providerSamples),
    ...samplesFromCollection(platform, input.provider_samples),
    ...samplesFromCollection(platform, input.sampleEvents),
    ...samplesFromCollection(platform, input.sample_events),
    ...samplesFromCollection(platform, input.samples),
    ...samplesFromCollection(platform, options.providerSamples),
    ...samplesFromCollection(platform, options.provider_samples),
    ...samplesFromCollection(platform, options.sampleEvents),
    ...samplesFromCollection(platform, options.sample_events),
    ...samplesFromCollection(platform, options.samples),
  ];
}

function recordsFromDomCollection(platform, collection) {
  if (!collection) return [];
  if (Array.isArray(collection)) return recordsForPlatform(platform, collection);
  if (collection.records) return recordsForPlatform(platform, meetingAppSnapshotRecords(collection));
  if (typeof collection === 'object') {
    const records = [];
    for (const [key, value] of Object.entries(collection)) {
      if (matchesPlatform(key, platform)) records.push(...asArray(value));
    }
    return records;
  }
  return [];
}

function meetingAppRecordsFor(platform, input = {}, options = {}) {
  return [
    ...recordsFromDomCollection(platform, input.meetingAppRecordSet),
    ...recordsFromDomCollection(platform, input.meeting_app_record_set),
    ...recordsFromDomCollection(platform, input.recordSet),
    ...recordsFromDomCollection(platform, input.record_set),
    ...recordsFromDomCollection(platform, input.meetingAppRecords),
    ...recordsFromDomCollection(platform, input.meeting_app_records),
    ...recordsFromDomCollection(platform, input.meetingAppSnapshotRecords),
    ...recordsFromDomCollection(platform, input.meeting_app_snapshot_records),
    ...recordsFromDomCollection(platform, input.meetingAppSnapshots),
    ...recordsFromDomCollection(platform, input.meeting_app_snapshots),
    ...recordsFromDomCollection(platform, input.domSnapshots),
    ...recordsFromDomCollection(platform, input.dom_snapshots),
    ...recordsFromDomCollection(platform, input.snapshots),
    ...recordsFromDomCollection(platform, options.meetingAppRecordSet),
    ...recordsFromDomCollection(platform, options.meeting_app_record_set),
    ...recordsFromDomCollection(platform, options.meetingAppRecords),
    ...recordsFromDomCollection(platform, options.meeting_app_records),
    ...recordsFromDomCollection(platform, options.meetingAppSnapshotRecords),
    ...recordsFromDomCollection(platform, options.meeting_app_snapshot_records),
    ...recordsFromDomCollection(platform, options.meetingAppSnapshots),
    ...recordsFromDomCollection(platform, options.meeting_app_snapshots),
    ...recordsFromDomCollection(platform, options.domSnapshots),
    ...recordsFromDomCollection(platform, options.dom_snapshots),
    ...recordsFromDomCollection(platform, options.snapshots),
  ];
}

function buildMeetingAppRecordSetFor(platform, records = [], options = {}) {
  if (records.length === 0) return undefined;
  const builtRecords = records.map((record) => buildMeetingAppSnapshotRecord(record, {
    platform,
  }));
  const createdAtMs = normalizeAbsoluteMs(firstNonEmpty(
    options.createdAtMs,
    options.created_at_ms,
    options.capturedAtMs,
    options.captured_at_ms,
    Date.now(),
  ), 'meeting_platform_evidence_package_dom_created_at');
  return buildMeetingAppSnapshotRecordSet(builtRecords, {
    id: firstNonEmpty(
      options.meetingAppRecordSetId,
      options.meeting_app_record_set_id,
      `${platform}-evidence-dom-${createdAtMs}`,
    ),
    label: firstNonEmpty(options.label),
    source: firstNonEmpty(options.source, 'meeting_platform_evidence_package'),
    createdAtMs,
  });
}

function safeEnvSummary(env = {}, options = {}) {
  if (!env || typeof env !== 'object') return undefined;
  const keys = Object.keys(env)
    .filter((key) => env[key] != null && env[key] !== '')
    .sort();
  if (keys.length === 0) return undefined;
  const includeValues = options.includeEnvValues === true || options.include_env_values === true;
  return compactObject({
    configured_keys: keys,
    configured: Object.fromEntries(keys.map((key) => [key, true])),
    values: includeValues ? Object.fromEntries(keys.map((key) => [key, env[key]])) : undefined,
  });
}

function providerSamplesOption(platform, samples = []) {
  return samples.length > 0 ? { [platform]: samples } : undefined;
}

function adapterRouteReady(adapterRoute = {}) {
  return verifyMeetingPlatformAdapterRouteReadiness(adapterRoute).ready === true;
}

function buildEvidenceSummary(rolloutPlan = {}, providerRecords = [], providerSamples = [], meetingAppRecordSet = undefined, evidenceCorrelation = undefined, adapterRoute = undefined) {
  return compactObject({
    status: rolloutPlan.status,
    production_ready: rolloutPlan.production_ready,
    ready_for_realtime_annotations: rolloutPlan.ready_for_realtime_annotations,
    correlation: evidenceCorrelation ? {
      status: evidenceCorrelation.status,
      passed: evidenceCorrelation.passed,
      confidence: evidenceCorrelation.confidence,
    } : undefined,
    provider: {
      record_count: providerRecords.length,
      sample_count: providerSamples.length,
      evidence_level: rolloutPlan.provider_events?.evidence_level,
      production_ready: rolloutPlan.provider_events?.production_ready,
      missing_required_coverage: rolloutPlan.provider_events?.missing_required_coverage,
    },
    local_dom: rolloutPlan.local_observer ? {
      record_count: meetingAppRecordSet?.record_count ?? 0,
      evidence_level: rolloutPlan.local_observer.evidence_level,
      production_ready: rolloutPlan.local_observer.production_ready,
      missing_required_coverage: rolloutPlan.local_observer.missing_required_coverage,
    } : undefined,
    adapter_route: summarizeMeetingPlatformAdapterRoute(adapterRoute),
    next_actions: rolloutPlan.next_actions,
  });
}

export function buildMeetingPlatformEvidencePackage(platformOrInput, inputOrOptions = {}, maybeOptions = {}) {
  const { input, platform, options } = normalizeArgs(platformOrInput, inputOrOptions, maybeOptions);
  const key = normalizeMeetingPlatform(platform);
  const createdAtMs = normalizeAbsoluteMs(firstNonEmpty(
    options.createdAtMs,
    options.created_at_ms,
    input.createdAtMs,
    input.created_at_ms,
    Date.now(),
  ), 'meeting_platform_evidence_package_created_at');
  const env = {
    ...(input.env ?? {}),
    ...(options.env ?? {}),
  };
  const providerRecords = providerRecordsFor(key, input, options);
  const providerSamples = providerSamplesFor(key, input, options);
  const meetingAppRecords = meetingAppRecordsFor(key, input, options);
  const meetingAppRecordSet = buildMeetingAppRecordSetFor(key, meetingAppRecords, {
    ...input,
    ...options,
    createdAtMs,
  });
  const rolloutOptions = {
    ...input,
    ...options,
    env,
    providerRecords,
    providerSamples: providerSamplesOption(key, providerSamples),
    meetingAppRecordSet,
  };
  const rolloutPlan = buildMeetingPlatformRolloutPlan(key, rolloutOptions);
  const includeRunbook = firstNonEmpty(
    options.includeRunbook,
    options.include_runbook,
    input.includeRunbook,
    input.include_runbook,
    true,
  );
  const runbook = includeRunbook === false
    ? undefined
    : buildMeetingPlatformAdaptationRunbook(key, rolloutOptions);
  const evidenceCorrelation = buildMeetingPlatformEvidenceCorrelation(key, {
    providerRecords,
    meetingAppRecordSet,
  }, options);
  const adapterRoute = buildMeetingPlatformAdapterRoute(key, rolloutOptions);
  const evidenceSummary = buildEvidenceSummary(
    rolloutPlan,
    providerRecords,
    providerSamples,
    meetingAppRecordSet,
    evidenceCorrelation,
    adapterRoute,
  );
  return compactObject({
    schema: MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA,
    schema_version: MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA_VERSION,
    id: firstNonEmpty(
      options.id,
      input.id,
      `${key}-evidence-package-${createdAtMs}`,
    ),
    platform: key,
    display_name: rolloutPlan.display_name,
    createdAtMs,
    created_at_ms: createdAtMs,
    source: firstNonEmpty(options.source, input.source, 'meeting_platform_evidence_package'),
    label: firstNonEmpty(options.label, input.label),
    base_url: firstNonEmpty(options.baseUrl, options.base_url, input.baseUrl, input.base_url),
    evidence_summary: evidenceSummary,
    provider_records: providerRecords,
    provider_samples: providerSamples.length > 0 ? { [key]: providerSamples } : undefined,
    meeting_app_record_set: meetingAppRecordSet,
    evidence_correlation: evidenceCorrelation,
    rollout_plan: rolloutPlan,
    adapter_route: adapterRoute,
    runbook,
    handoff: {
      provider_evidence_input: 'providerRecords',
      meeting_app_evidence_input: 'meetingAppRecordSet',
      adapter_route_input: 'adapter_route',
      adapter_route_gate: 'buildMeetingPlatformAdapterRoute',
      final_gate: 'buildMeetingPlatformRolloutPlan',
      pilot_condition: 'ready_for_realtime_annotations === true',
      production_condition: 'production_ready === true',
      status: rolloutPlan.status,
      recommended_mode: rolloutPlan.recommended_mode,
      adapter_recommended_mode: adapterRoute.recommended_mode,
      adapter_first_route: adapterRoute.routes?.[0]?.route,
      next_actions: rolloutPlan.next_actions,
    },
    env_summary: safeEnvSummary(env, { ...input, ...options }),
    notes: firstNonEmpty(options.notes, input.notes),
    metadata: firstNonEmpty(options.metadata, input.metadata),
  });
}

export function buildMeetingPlatformEvidencePackageSummary(packageOrInput, options = {}) {
  const pkg = packageOrInput?.schema === MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA
    ? packageOrInput
    : buildMeetingPlatformEvidencePackage(packageOrInput, options);
  const plan = pkg.rollout_plan ?? {};
  return compactObject({
    type: 'meeting_platform_evidence_package_summary',
    package_id: pkg.id,
    platform: pkg.platform,
    display_name: pkg.display_name,
    status: plan.status,
    recommended_mode: plan.recommended_mode,
    production_ready: plan.production_ready,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    provider_record_count: pkg.provider_records?.length ?? 0,
    provider_sample_count: Object.values(pkg.provider_samples ?? {}).flat().length,
    meeting_app_record_count: pkg.meeting_app_record_set?.record_count ?? 0,
    correlation_status: pkg.evidence_correlation?.status,
    correlation_passed: pkg.evidence_correlation?.passed,
    correlation_confidence: pkg.evidence_correlation?.confidence,
    adapter_route_ready: adapterRouteReady(pkg.adapter_route),
    adapter_recommended_mode: pkg.adapter_route?.recommended_mode,
    adapter_first_route: pkg.adapter_route?.routes?.[0]?.route,
    provider_events_block_realtime: pkg.adapter_route?.realtime_invariants?.provider_events_block_realtime,
    transcript_blocks_realtime: pkg.adapter_route?.realtime_invariants?.transcript_blocks_realtime,
    provider_missing_required_coverage: plan.provider_events?.missing_required_coverage ?? [],
    local_dom_missing_required_coverage: plan.local_observer?.missing_required_coverage ?? [],
    next_actions: plan.next_actions ?? [],
  });
}

function normalizedStatus(plan = {}) {
  return {
    status: plan.status,
    production_ready: Boolean(plan.production_ready),
    ready_for_realtime_annotations: Boolean(plan.ready_for_realtime_annotations),
    provider_evidence_level: plan.provider_events?.evidence_level,
    provider_evidence_count: plan.provider_events?.evidence_count ?? 0,
    local_dom_evidence_level: plan.local_observer?.evidence_level,
    local_dom_evidence_count: plan.local_observer?.evidence_count ?? 0,
  };
}

function statusMatches(left = {}, right = {}) {
  const a = normalizedStatus(left);
  const b = normalizedStatus(right);
  return Object.keys(a).every((key) => a[key] === b[key]);
}

export function verifyMeetingPlatformEvidencePackage(packageOrInput, options = {}) {
  const verifiedPackage = buildMeetingPlatformEvidencePackage(packageOrInput, {
    ...options,
    includeRunbook: firstNonEmpty(options.includeRunbook, options.include_runbook, false),
  });
  const embeddedPlan = packageOrInput?.schema === MEETING_PLATFORM_EVIDENCE_PACKAGE_SCHEMA
    ? packageOrInput.rollout_plan
    : undefined;
  const verifiedPlan = verifiedPackage.rollout_plan ?? {};
  const requireProductionReady = options.requireProductionReady !== false
    && options.require_production_ready !== false;
  const requireCorrelation = options.requireCorrelation !== false
    && options.require_correlation !== false;
  const passed = requireProductionReady
    ? verifiedPlan.production_ready === true
    : verifiedPlan.ready_for_realtime_annotations === true;
  const correlationPassed = verifiedPackage.evidence_correlation?.passed !== false;
  const routeReady = adapterRouteReady(verifiedPackage.adapter_route);
  const finalPassed = passed && (!requireCorrelation || correlationPassed) && routeReady;
  const embeddedPlanMatches = embeddedPlan ? statusMatches(embeddedPlan, verifiedPlan) : undefined;
  return compactObject({
    type: 'meeting_platform_evidence_package_verification',
    package_id: verifiedPackage.id,
    platform: verifiedPackage.platform,
    passed: finalPassed,
    requirement: requireProductionReady ? 'production_ready' : 'ready_for_realtime_annotations',
    status: verifiedPlan.status,
    production_ready: Boolean(verifiedPlan.production_ready),
    ready_for_realtime_annotations: Boolean(verifiedPlan.ready_for_realtime_annotations),
    embedded_plan_matches: embeddedPlanMatches,
    embedded_status: embeddedPlan ? normalizedStatus(embeddedPlan) : undefined,
    verified_status: normalizedStatus(verifiedPlan),
    provider_record_count: verifiedPackage.provider_records?.length ?? 0,
    provider_sample_count: Object.values(verifiedPackage.provider_samples ?? {}).flat().length,
    meeting_app_record_count: verifiedPackage.meeting_app_record_set?.record_count ?? 0,
    correlation_required: requireCorrelation,
    correlation_passed: correlationPassed,
    correlation_status: verifiedPackage.evidence_correlation?.status,
    correlation_confidence: verifiedPackage.evidence_correlation?.confidence,
    evidence_correlation: verifiedPackage.evidence_correlation,
    adapter_route_ready: routeReady,
    adapter_recommended_mode: verifiedPackage.adapter_route?.recommended_mode,
    adapter_first_route: verifiedPackage.adapter_route?.routes?.[0]?.route,
    provider_events_block_realtime: verifiedPackage.adapter_route?.realtime_invariants?.provider_events_block_realtime,
    transcript_blocks_realtime: verifiedPackage.adapter_route?.realtime_invariants?.transcript_blocks_realtime,
    provider_missing_required_coverage: verifiedPlan.provider_events?.missing_required_coverage ?? [],
    local_dom_missing_required_coverage: verifiedPlan.local_observer?.missing_required_coverage ?? [],
    next_actions: verifiedPlan.next_actions ?? [],
    verified_package: options.includePackage === true || options.include_package === true
      ? verifiedPackage
      : undefined,
  });
}

export function createMeetingPlatformEvidencePackageBuilder(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const providerRecords = [];
  const providerSamples = [];
  const meetingAppRecords = [];

  function addProviderRecord(record = {}) {
    providerRecords.push(record);
    return record;
  }

  function captureProviderWebhook(input = {}, payload, captureOptions = {}) {
    const objectInput = input && typeof input === 'object' && !Array.isArray(input);
    const record = capturePlatformWebhookEvent({
      ...(objectInput ? input : {}),
      platform: key,
    }, objectInput ? payload : input, {
      ...options,
      ...captureOptions,
      platform: key,
    });
    providerRecords.push(record);
    return record;
  }

  function addProviderSample(sample = {}) {
    providerSamples.push(sample);
    return sample;
  }

  function addMeetingAppRecord(input = {}, recordOptions = {}) {
    const record = buildMeetingAppSnapshotRecord(input, {
      ...options,
      ...recordOptions,
      platform: key,
    });
    meetingAppRecords.push(record);
    return record;
  }

  function exportPackage(packageOptions = {}) {
    return buildMeetingPlatformEvidencePackage(key, {
      providerRecords,
      providerSamples,
      meetingAppRecords,
    }, {
      ...options,
      ...packageOptions,
    });
  }

  function summary(summaryOptions = {}) {
    return buildMeetingPlatformEvidencePackageSummary(exportPackage(summaryOptions));
  }

  function reset() {
    const removed = {
      provider_records: providerRecords.length,
      provider_samples: providerSamples.length,
      meeting_app_records: meetingAppRecords.length,
    };
    providerRecords.splice(0);
    providerSamples.splice(0);
    meetingAppRecords.splice(0);
    return { removed };
  }

  function getState() {
    return {
      platform: key,
      provider_record_count: providerRecords.length,
      provider_sample_count: providerSamples.length,
      meeting_app_record_count: meetingAppRecords.length,
      latest_provider_record_id: providerRecords.at(-1)?.id,
      latest_meeting_app_record_id: meetingAppRecords.at(-1)?.id,
    };
  }

  return {
    platform: key,
    addProviderRecord,
    captureProviderWebhook,
    addProviderSample,
    addMeetingAppRecord,
    captureMeetingAppSnapshot: addMeetingAppRecord,
    exportPackage,
    summary,
    reset,
    getState,
    providerRecords() {
      return providerRecords.slice();
    },
    providerSamples() {
      return providerSamples.slice();
    },
    meetingAppRecords() {
      return meetingAppRecords.slice();
    },
  };
}
