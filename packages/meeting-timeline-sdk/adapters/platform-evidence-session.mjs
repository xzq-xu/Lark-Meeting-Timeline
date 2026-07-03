import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  buildMeetingAppSnapshotRecord,
  buildMeetingAppSnapshotRecordSet,
} from './meeting-app-snapshot-recorder.mjs';
import { capturePlatformWebhookEvent } from './platform-capture.mjs';
import { buildMeetingPlatformEvidenceCorrelation } from './platform-evidence-correlation.mjs';
import {
  buildMeetingPlatformEvidencePackage,
  buildMeetingPlatformEvidencePackageSummary,
  verifyMeetingPlatformEvidencePackage,
} from './platform-evidence-package.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA = 'meeting_platform_evidence_session';
export const MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function createdAt(options = {}, fallback = Date.now()) {
  return normalizeAbsoluteMs(firstNonEmpty(
    options.createdAtMs,
    options.created_at_ms,
    options.capturedAtMs,
    options.captured_at_ms,
    fallback,
  ), 'meeting_platform_evidence_session_time');
}

function providerSamplesOption(platform, samples = []) {
  return samples.length > 0 ? { [platform]: samples } : undefined;
}

function mergeOptions(base = {}, next = {}) {
  return compactObject({
    ...base,
    ...next,
    env: {
      ...(base.env ?? {}),
      ...(next.env ?? {}),
    },
  });
}

function normalizeProviderSample(sample = {}, platform) {
  if (!sample || typeof sample !== 'object' || Array.isArray(sample)) return sample;
  return {
    ...sample,
    platform,
  };
}

export function createMeetingPlatformEvidenceSession(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const sessionCreatedAtMs = createdAt(options);
  const sessionId = firstNonEmpty(
    options.id,
    options.sessionId,
    options.session_id,
    `${key}-evidence-session-${sessionCreatedAtMs}`,
  );
  const providerRecords = [];
  const providerSamples = [];
  const meetingAppRecords = [];

  function addProviderRecord(record = {}) {
    const row = record && typeof record === 'object' && !Array.isArray(record)
      ? { ...record, platform: key }
      : record;
    providerRecords.push(row);
    return row;
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
    const row = normalizeProviderSample(sample, key);
    providerSamples.push(row);
    return row;
  }

  function addMeetingAppRecord(input = {}, recordOptions = {}) {
    const record = buildMeetingAppSnapshotRecord(input, {
      ...options,
      ...recordOptions,
      platform: key,
      source: firstNonEmpty(recordOptions.source, options.source, 'meeting_platform_evidence_session'),
    });
    meetingAppRecords.push(record);
    return record;
  }

  function meetingAppRecordSet(recordSetOptions = {}) {
    if (meetingAppRecords.length === 0) return undefined;
    const merged = mergeOptions(options, recordSetOptions);
    const setCreatedAtMs = createdAt(merged);
    return buildMeetingAppSnapshotRecordSet(meetingAppRecords, {
      ...merged,
      id: firstNonEmpty(
        merged.meetingAppRecordSetId,
        merged.meeting_app_record_set_id,
        merged.recordSetId,
        merged.record_set_id,
        `${sessionId}-dom-${setCreatedAtMs}`,
      ),
      source: firstNonEmpty(merged.source, 'meeting_platform_evidence_session'),
      createdAtMs: setCreatedAtMs,
    });
  }

  function packageInput(recordSetOptions = {}) {
    return compactObject({
      providerRecords,
      providerSamples,
      meetingAppRecordSet: meetingAppRecordSet(recordSetOptions),
    });
  }

  function exportPackage(packageOptions = {}) {
    const merged = mergeOptions(options, packageOptions);
    return buildMeetingPlatformEvidencePackage(key, packageInput(merged), {
      ...merged,
      id: firstNonEmpty(
        packageOptions.packageId,
        packageOptions.package_id,
        packageOptions.id,
        options.packageId,
        options.package_id,
        `${sessionId}-package`,
      ),
      source: firstNonEmpty(merged.source, 'meeting_platform_evidence_session'),
    });
  }

  function correlation(correlationOptions = {}) {
    const merged = mergeOptions(options, correlationOptions);
    return buildMeetingPlatformEvidenceCorrelation(key, packageInput(merged), merged);
  }

  function strategy(strategyOptions = {}) {
    const merged = mergeOptions(options, strategyOptions);
    return buildMeetingPlatformAdaptationStrategy(key, {
      ...merged,
      providerRecords,
      providerSamples: providerSamplesOption(key, providerSamples),
      meetingAppRecordSet: meetingAppRecordSet(merged),
    });
  }

  function verify(verifyOptions = {}) {
    const merged = mergeOptions(options, verifyOptions);
    return verifyMeetingPlatformEvidencePackage(exportPackage({
      ...merged,
      includeRunbook: firstNonEmpty(merged.includeRunbook, merged.include_runbook, false),
    }), merged);
  }

  function summary(summaryOptions = {}) {
    const merged = mergeOptions(options, summaryOptions);
    const pkg = exportPackage({
      ...merged,
      includeRunbook: firstNonEmpty(merged.includeRunbook, merged.include_runbook, false),
    });
    const packageSummary = buildMeetingPlatformEvidencePackageSummary(pkg);
    const verification = verifyMeetingPlatformEvidencePackage(pkg, merged);
    const rollout = pkg.rollout_plan ?? {};
    const evidenceCorrelation = pkg.evidence_correlation ?? {};
    return compactObject({
      type: 'meeting_platform_evidence_session_summary',
      schema: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA,
      schema_version: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA_VERSION,
      session_id: sessionId,
      package_id: pkg.id,
      platform: key,
      display_name: pkg.display_name,
      createdAtMs: sessionCreatedAtMs,
      created_at_ms: sessionCreatedAtMs,
      status: rollout.status,
      recommended_mode: rollout.recommended_mode,
      production_ready: Boolean(rollout.production_ready),
      ready_for_realtime_annotations: Boolean(rollout.ready_for_realtime_annotations),
      can_insert_realtime_marks: Boolean(rollout.ready_for_realtime_annotations),
      local_observer_ready: rollout.local_observer?.production_ready === true,
      provider_reconcile_ready: rollout.provider_events?.production_ready === true,
      package_ready_for_handoff: providerRecords.length > 0 || providerSamples.length > 0 || meetingAppRecords.length > 0,
      verification_passed: verification.passed,
      verification_requirement: verification.requirement,
      provider_record_count: providerRecords.length,
      provider_sample_count: providerSamples.length,
      meeting_app_record_count: meetingAppRecords.length,
      correlation_status: evidenceCorrelation.status,
      correlation_passed: evidenceCorrelation.passed,
      correlation_confidence: evidenceCorrelation.confidence,
      provider_missing_required_coverage: rollout.provider_events?.missing_required_coverage ?? [],
      local_dom_missing_required_coverage: rollout.local_observer?.missing_required_coverage ?? [],
      next_actions: uniqueList([
        ...(rollout.next_actions ?? []),
        ...(verification.next_actions ?? []),
      ]),
      handoff: {
        export_method: 'exportPackage',
        verify_method: 'verify',
        package_schema: 'meeting_platform_evidence_package',
        pilot_condition: 'ready_for_realtime_annotations === true',
        production_condition: 'production_ready === true',
      },
      evidence_package_summary: merged.includePackageSummary === true || merged.include_package_summary === true
        ? packageSummary
        : undefined,
      verification: merged.includeVerification === true || merged.include_verification === true
        ? verification
        : undefined,
    });
  }

  function getState() {
    return {
      type: 'meeting_platform_evidence_session_state',
      schema: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA,
      schema_version: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA_VERSION,
      session_id: sessionId,
      platform: key,
      createdAtMs: sessionCreatedAtMs,
      created_at_ms: sessionCreatedAtMs,
      provider_record_count: providerRecords.length,
      provider_sample_count: providerSamples.length,
      meeting_app_record_count: meetingAppRecords.length,
      latest_provider_record_id: providerRecords.at(-1)?.id,
      latest_meeting_app_record_id: meetingAppRecords.at(-1)?.id,
    };
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

  return {
    schema: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA,
    schema_version: MEETING_PLATFORM_EVIDENCE_SESSION_SCHEMA_VERSION,
    platform: key,
    id: sessionId,
    session_id: sessionId,
    addProviderRecord,
    captureProviderWebhook,
    addProviderSample,
    addMeetingAppRecord,
    captureMeetingAppSnapshot: addMeetingAppRecord,
    meetingAppRecordSet,
    correlation,
    strategy,
    exportPackage,
    verify,
    summary,
    getState,
    reset,
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
