import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  buildMeetingAppFixtureSnapshot,
} from './meeting-app-fixtures.mjs';
import {
  buildPlatformFixtureEnv,
  buildPlatformFixtureEvent,
  DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES,
} from './platform-fixtures.mjs';
import {
  createMeetingPlatformLiveAdapter,
  buildMeetingPlatformLiveAdapterReadiness,
} from './platform-live-adapter.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import {
  normalizeMeetingPlatform,
  platformEventEndpoint,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA = 'meeting_platform_adapter_sample';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_MATRIX_SCHEMA = 'meeting_platform_adapter_sample_matrix';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_PLAN_SCHEMA = 'meeting_platform_adapter_sample_plan';
export const MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA_VERSION = 1;

export const MEETING_PLATFORM_ADAPTER_SAMPLE_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

const DEFAULT_START_MS = 1_784_010_000_000;
const DEFAULT_DURATION_MS = 120_000;
const DEFAULT_ANNOTATION_OFFSET_MS = 10_000;

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
    MEETING_PLATFORM_ADAPTER_SAMPLE_PLATFORMS,
  )).map((platform) => normalizeMeetingPlatform(platform)));
}

function sampleClock(options = {}) {
  const startMs = normalizeAbsoluteMs(firstNonEmpty(
    options.startMs,
    options.start_ms,
    DEFAULT_START_MS,
  ), 'meeting_platform_adapter_sample_start_ms');
  const durationMs = Number(firstNonEmpty(
    options.durationMs,
    options.duration_ms,
    DEFAULT_DURATION_MS,
  ));
  const annotationOffsetMs = Number(firstNonEmpty(
    options.annotationOffsetMs,
    options.annotation_offset_ms,
    DEFAULT_ANNOTATION_OFFSET_MS,
  ));
  return {
    startMs,
    start_ms: startMs,
    endMs: startMs + durationMs,
    end_ms: startMs + durationMs,
    durationMs,
    duration_ms: durationMs,
    annotationMs: startMs + annotationOffsetMs,
    annotation_ms: startMs + annotationOffsetMs,
  };
}

function providerSignalTypes(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const explicit = firstNonEmpty(options.providerSignalTypes, options.provider_signal_types);
  if (explicit) return unique(asArray(explicit));
  const supported = DEFAULT_PLATFORM_FIXTURE_SIGNAL_TYPES[key] ?? [];
  const preferred = key === 'local_detector'
    ? ['meeting_start', 'speaker_activity', 'participant_join', 'participant_left', 'meeting_end']
    : ['meeting_start', 'participant_join', 'participant_left', 'meeting_end', 'transcript_ready'];
  return preferred.filter((signal) => supported.includes(signal));
}

function createRecordingTimelineClient(calls = []) {
  function push(method, input, options) {
    const row = compactObject({ method, input, options });
    calls.push(row);
    return { ok: true, method, input };
  }
  return {
    calls,
    async startMeeting(input, options) {
      return push('startMeeting', input, options);
    },
    async endMeeting(input, options) {
      return push('endMeeting', input, options);
    },
    async insertMark(input, options) {
      return push('insertMark', input, options);
    },
    async insertMarks(input, options) {
      return push('insertMarks', input, options);
    },
    async importTranscript(input, options) {
      return push('importTranscript', input, options);
    },
  };
}

function methodCounts(calls = []) {
  const counts = {};
  for (const call of calls) {
    counts[call.method] = (counts[call.method] ?? 0) + 1;
  }
  return counts;
}

function sampleAnnotation(platform, clock = {}, options = {}) {
  return {
    id: firstNonEmpty(options.annotation?.id, options.markId, options.mark_id, `${platform}-sample-mark-1`),
    kind: firstNonEmpty(options.annotation?.kind, options.kind, 'question'),
    label: firstNonEmpty(options.annotation?.label, options.label, 'why?'),
    capturedAtMs: firstNonEmpty(options.annotation?.capturedAtMs, options.annotation?.captured_at_ms, clock.annotationMs),
    target: firstNonEmpty(options.annotation?.target, options.target, 'adapter sample timeline'),
    ...(options.annotation ?? {}),
  };
}

export function buildMeetingPlatformAdapterSamplePlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787'));
  const clock = sampleClock(options);
  const signals = providerSignalTypes(key, options);
  return {
    type: 'meeting_platform_adapter_sample_plan',
    schema: MEETING_PLATFORM_ADAPTER_SAMPLE_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA_VERSION,
    platform: key,
    base_url: baseUrl,
    endpoint: platformEventEndpoint(baseUrl, key),
    clock,
    provider_signal_types: signals,
    local_observer_snapshots: [
      { state: 'active', captured_at_ms: clock.startMs },
      { state: 'prejoin', captured_at_ms: clock.endMs },
    ],
    annotation: sampleAnnotation(key, clock, options),
    flow: [
      'observe active local meeting app snapshot',
      'insert realtime annotation with capturedAtMs',
      'ingest provider events for async reconcile',
      'observe ended local meeting app snapshot',
      'export and verify evidence package',
    ],
    contract: buildMeetingPlatformAdapterContract(key, { ...options, baseUrl }),
    contract_acceptance: buildMeetingPlatformAdapterContractAcceptanceReport(key, { ...options, baseUrl }),
  };
}

export async function runMeetingPlatformAdapterSample(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const plan = buildMeetingPlatformAdapterSamplePlan(key, options);
  const calls = [];
  const client = firstNonEmpty(options.client, createRecordingTimelineClient(calls));
  const baseUrl = plan.base_url;
  const env = buildPlatformFixtureEnv({
    ...options,
    baseUrl,
    env: {
      ...(options.env ?? {}),
    },
  });
  const adapter = firstNonEmpty(options.adapter, options.liveAdapter, options.live_adapter)
    ?? createMeetingPlatformLiveAdapter(key, client, {
      ...options,
      baseUrl,
      env,
      id: firstNonEmpty(options.id, options.sessionId, options.session_id, `${key}-adapter-sample`),
      createdAtMs: plan.clock.startMs,
      requireMeetingEnd: firstNonEmpty(options.requireMeetingEnd, options.require_meeting_end, false),
      speakerOptions: {
        minStableMs: 0,
        ...(options.speakerOptions ?? {}),
        ...(options.speaker_options ?? {}),
      },
      reconcileOptions: {
        duplicateWindowMs: 60_000,
        ...(options.reconcileOptions ?? {}),
        ...(options.reconcile_options ?? {}),
      },
    });

  const activeSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options.fixtureOptions,
    ...options.fixture_options,
    state: 'active',
    observedAtMs: plan.clock.startMs,
    startMs: plan.clock.startMs,
  });
  const observeStart = await adapter.observeMeetingApp(activeSnapshot, {
    capturedAtMs: plan.clock.startMs,
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
  });

  const annotation = sampleAnnotation(key, plan.clock, options);
  const insertAnnotation = await adapter.insertAnnotation(annotation, {
    capturedAtMs: annotation.capturedAtMs,
  });

  const providerEvents = [];
  for (const signalType of plan.provider_signal_types) {
    const body = buildPlatformFixtureEvent(key, signalType, {
      ...options.providerFixtureOptions,
      ...options.provider_fixture_options,
      baseUrl,
      startMs: plan.clock.startMs,
      durationMs: plan.clock.durationMs,
    });
    const result = await adapter.ingestProvider({
      method: 'POST',
      url: plan.endpoint,
      body,
    }, undefined, {
      capturedAtMs: plan.clock.startMs,
    });
    providerEvents.push({
      signal_type: signalType,
      body,
      result,
    });
  }

  const endedSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options.fixtureOptions,
    ...options.fixture_options,
    state: 'prejoin',
    observedAtMs: plan.clock.endMs,
    startMs: plan.clock.startMs,
  });
  const observeEnd = await adapter.observeMeetingApp(endedSnapshot, {
    phase: 'ended',
    capturedAtMs: plan.clock.endMs,
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
  });

  const evidencePackage = adapter.exportPackage({
    includeRunbook: firstNonEmpty(options.includeRunbook, options.include_runbook, false),
  });
  const verification = adapter.verify({
    requireProductionReady: firstNonEmpty(options.requireProductionReady, options.require_production_ready, true),
    includePackage: firstNonEmpty(options.includePackage, options.include_package, false),
  });
  const summary = adapter.summary({
    includeVerification: true,
  });
  const readiness = buildMeetingPlatformLiveAdapterReadiness(key, {
    ...options,
    baseUrl,
    env,
    target: firstNonEmpty(options.target, 'production'),
    adapter,
    evidencePackage,
  });
  const counts = methodCounts(calls);
  const accepted = plan.contract_acceptance.accepted === true
    && readiness.passed === true
    && verification.passed === true
    && summary.can_insert_realtime_marks === true
    && counts.startMeeting > 0
    && counts.insertMark > 0;

  return compactObject({
    type: 'meeting_platform_adapter_sample',
    schema: MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA_VERSION,
    platform: key,
    accepted,
    base_url: baseUrl,
    endpoint: plan.endpoint,
    operation_count: 3 + providerEvents.length,
    provider_event_count: providerEvents.length,
    timeline_call_count: calls.length,
    timeline_method_counts: counts,
    contract_acceptance: plan.contract_acceptance,
    readiness,
    verification,
    summary,
    plan,
    operations: {
      observe_start: observeStart,
      insert_annotation: insertAnnotation,
      provider_events: providerEvents,
      observe_end: observeEnd,
    },
    evidence_package: firstNonEmpty(options.includeEvidencePackage, options.include_evidence_package, false)
      ? evidencePackage
      : undefined,
    timeline_calls: firstNonEmpty(options.includeTimelineCalls, options.include_timeline_calls, true)
      ? calls
      : undefined,
  });
}

export async function runMeetingPlatformAdapterSampleMatrix(options = {}) {
  const platforms = selectedPlatforms(options);
  const samples = [];
  for (const platform of platforms) {
    samples.push(await runMeetingPlatformAdapterSample(platform, {
      ...options,
      platforms,
    }));
  }
  return {
    type: 'meeting_platform_adapter_sample_matrix',
    schema: MEETING_PLATFORM_ADAPTER_SAMPLE_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA_VERSION,
    platform_count: samples.length,
    accepted_count: samples.filter((sample) => sample.accepted).length,
    rejected_count: samples.filter((sample) => !sample.accepted).length,
    production_ready_count: samples.filter((sample) => sample.readiness?.production_ready === true).length,
    realtime_ready_count: samples.filter((sample) => sample.readiness?.ready_for_realtime_annotations === true).length,
    platforms,
    rows: samples.map((sample) => ({
      platform: sample.platform,
      accepted: sample.accepted,
      readiness_status: sample.readiness?.status,
      production_ready: sample.readiness?.production_ready,
      ready_for_realtime_annotations: sample.readiness?.ready_for_realtime_annotations,
      provider_event_count: sample.provider_event_count,
      timeline_method_counts: sample.timeline_method_counts,
    })),
    samples,
  };
}

export function assertMeetingPlatformAdapterSample(sampleOrPlatform, options = {}) {
  const assertSample = async () => {
    const sample = sampleOrPlatform?.schema === MEETING_PLATFORM_ADAPTER_SAMPLE_SCHEMA
      ? sampleOrPlatform
      : await runMeetingPlatformAdapterSample(sampleOrPlatform, options);
    if (sample.accepted !== true) {
      throw new MeetingTimelineSdkError(`Meeting platform adapter sample failed for ${sample.platform}`, {
        platform: sample.platform,
        readiness_status: sample.readiness?.status,
        contract_accepted: sample.contract_acceptance?.accepted,
        verification_passed: sample.verification?.passed,
        timeline_method_counts: sample.timeline_method_counts,
        sample,
      });
    }
    return sample;
  };
  return assertSample();
}

export async function assertMeetingPlatformAdapterSampleMatrix(options = {}) {
  const matrix = await runMeetingPlatformAdapterSampleMatrix(options);
  if (matrix.rejected_count > 0) {
    throw new MeetingTimelineSdkError('Meeting platform adapter sample matrix failed', {
      platform_count: matrix.platform_count,
      accepted_count: matrix.accepted_count,
      rejected_count: matrix.rejected_count,
      failed_platforms: matrix.rows.filter((row) => !row.accepted).map((row) => row.platform),
      matrix,
    });
  }
  return matrix;
}
