import { compactObject } from '../index.mjs';
import { createMeetingSourceAggregator } from './meeting-source.mjs';
import { createMeetingPlatformEvidenceSession } from './platform-evidence-session.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { buildMeetingPlatformAdaptationStrategy } from './platform-strategy.mjs';

export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA = 'meeting_platform_live_adapter';
export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION = 1;
export const MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA = 'meeting_platform_live_adapter_plan';
export const MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA = 'meeting_platform_live_adapter_matrix';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function isTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.startMeeting === 'function'
    && typeof value.endMeeting === 'function'
    && typeof value.insertMark === 'function';
}

function liveSessionOptions(options = {}) {
  return compactObject({
    requireMeetingEnd: firstNonEmpty(options.requireMeetingEnd, options.require_meeting_end, false),
    ...options,
  });
}

function providerPayload(input, payload) {
  if (input && typeof input === 'object' && !Array.isArray(input)) {
    return firstNonEmpty(input.payload, input.body, input.event, input.raw, payload);
  }
  return payload;
}

function meetingAppRecordOptions(input = {}, options = {}) {
  return {
    ...options,
    phase: firstNonEmpty(
      options.phase,
      options.state,
      input.phase,
      input.state,
      input.fixture_state,
      input.snapshot?.phase,
      input.snapshot?.state,
      input.snapshot?.fixture_state,
    ),
    capturedAtMs: firstNonEmpty(
      options.capturedAtMs,
      options.captured_at_ms,
      options.observedAtMs,
      options.observed_at_ms,
      input.capturedAtMs,
      input.captured_at_ms,
      input.observedAtMs,
      input.observed_at_ms,
      input.snapshot?.observedAtMs,
      input.snapshot?.observed_at_ms,
    ),
  };
}

function resultWithEvidence(result = {}, evidenceSession, action, options = {}) {
  return compactObject({
    ...result,
    action,
    live_evidence: evidenceSession.summary(options.evidenceOptions ?? options.evidence_options ?? {}),
    evidence_state: evidenceSession.getState(),
  });
}

function operationWithEvidence(action, result, evidenceSession, options = {}) {
  return compactObject({
    action,
    result,
    live_evidence: evidenceSession.summary(options.evidenceOptions ?? options.evidence_options ?? {}),
    evidence_state: evidenceSession.getState(),
  });
}

function selectedPlatforms(options = {}) {
  return asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS));
}

export function buildMeetingPlatformLiveAdapterPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const integration = buildPlatformIntegrationPlan(key, options);
  const strategy = buildMeetingPlatformAdaptationStrategy(key, options);
  return compactObject({
    schema: MEETING_PLATFORM_LIVE_ADAPTER_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    type: 'meeting_platform_live_adapter_plan',
    platform: key,
    display_name: strategy.display_name ?? integration.display_name,
    rollout_status: strategy.rollout_status,
    production_ready: strategy.production_ready,
    ready_for_realtime_annotations: strategy.ready_for_realtime_annotations,
    pilot_ready: strategy.ready_for_realtime_annotations === true,
    recommended_mode: strategy.recommendation,
    source_priority: strategy.source_priority,
    live_adapter: {
      module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter',
      factory: 'createMeetingPlatformLiveAdapter',
      kit_method: 'platformLiveAdapter',
      realtime_methods: ['observeMeetingApp', 'ingestProvider', 'insertAnnotation'],
      evidence_methods: ['summary', 'correlation', 'exportPackage', 'verify'],
      timestamp_field: strategy.realtime_axis?.annotation_time_field ?? 'captured_at_ms',
    },
    realtime_axis: strategy.realtime_axis,
    local_observer: strategy.local_observer,
    provider_events: strategy.provider_events,
    speaker_activity: strategy.speaker_activity,
    post_meeting_transcript: strategy.post_meeting_transcript,
    handoff: {
      evidence_session_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session',
      evidence_package_module: '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package',
      package_schema: 'meeting_platform_evidence_package',
      verifier: 'verifyMeetingPlatformEvidencePackage',
    },
    endpoints: {
      provider_events: integration.provider_events?.endpoint,
      transcript_import: integration.post_meeting_transcript?.import_endpoint,
      status: integration.provider_events?.status_endpoint,
    },
    next_actions: uniqueList(strategy.next_actions ?? []),
  });
}

export function buildMeetingPlatformLiveAdapterMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformLiveAdapterPlan(platform, options));
  return {
    type: 'meeting_platform_live_adapter_matrix',
    schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform_count: plans.length,
    production_ready_count: plans.filter((plan) => plan.production_ready).length,
    pilot_ready_count: plans.filter((plan) => plan.pilot_ready).length,
    local_first_count: plans.filter((plan) => plan.realtime_axis?.primary_source === 'local_observer').length,
    non_blocking_provider_count: plans.filter((plan) => plan.realtime_axis?.provider_events_block_realtime === false).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      rollout_status: plan.rollout_status,
      recommended_mode: plan.recommended_mode,
      pilot_ready: plan.pilot_ready,
      production_ready: plan.production_ready,
      primary_axis_source: plan.realtime_axis?.primary_source,
      provider_blocks_realtime: plan.realtime_axis?.provider_events_block_realtime,
      transcript_blocks_realtime: plan.realtime_axis?.transcript_blocks_realtime,
      next_actions: plan.next_actions,
    })),
    plans,
  };
}

export function createMeetingPlatformLiveAdapter(platform, clientOrOptions = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const sourceInput = isTimelineClient(clientOrOptions)
    ? clientOrOptions
    : {
      ...(clientOrOptions ?? {}),
      ...(options.clientOptions ?? options.client_options ?? {}),
      baseUrl: firstNonEmpty(options.baseUrl, options.base_url, clientOrOptions?.baseUrl, clientOrOptions?.base_url),
    };
  const sourceOptions = {
    ...options,
    platform: key,
  };
  const source = createMeetingSourceAggregator(sourceInput, sourceOptions);
  const evidenceSession = createMeetingPlatformEvidenceSession(key, liveSessionOptions({
    ...options,
    id: firstNonEmpty(options.id, options.sessionId, options.session_id, `${key}-live-adapter`),
    source: firstNonEmpty(options.source, 'meeting_platform_live_adapter'),
  }));

  async function observeMeetingApp(input = {}, observeOptions = {}) {
    if (observeOptions.captureEvidence !== false && observeOptions.capture_evidence !== false) {
      evidenceSession.captureMeetingAppSnapshot(input, meetingAppRecordOptions(input, observeOptions));
    }
    const result = await source.observeMeetingApp(input, observeOptions);
    return resultWithEvidence(result, evidenceSession, 'observe_meeting_app', observeOptions);
  }

  async function ingestProvider(input = {}, payload, ingestOptions = {}) {
    if (ingestOptions.captureEvidence !== false && ingestOptions.capture_evidence !== false) {
      evidenceSession.captureProviderWebhook(input, payload, ingestOptions);
    }
    const result = await source.ingestProvider(key, providerPayload(input, payload), ingestOptions);
    return resultWithEvidence(result, evidenceSession, 'ingest_provider', ingestOptions);
  }

  return {
    schema: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platform: key,
    source,
    evidenceSession,
    session: evidenceSession,
    observeMeetingApp,
    observeApp: observeMeetingApp,
    async observeBrowser(input = {}, observeOptions = {}) {
      const result = await source.observeBrowser(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_browser', observeOptions);
    },
    async observeNative(input = {}, observeOptions = {}) {
      const result = await source.observeNative(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_native', observeOptions);
    },
    async observeLocal(input = {}, observeOptions = {}) {
      const result = await source.observeLocal(input, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_local', observeOptions);
    },
    async observeLocalCandidates(candidates = [], observeOptions = {}) {
      const result = await source.observeLocalCandidates(candidates, observeOptions);
      return resultWithEvidence(result, evidenceSession, 'observe_local_candidates', observeOptions);
    },
    ingestProvider,
    async ingestSignals(signals = [], ingestOptions = {}) {
      const result = await source.ingestSignals(signals, ingestOptions);
      return resultWithEvidence(result, evidenceSession, 'ingest_signals', ingestOptions);
    },
    captureMeetingAppSnapshot(input = {}, captureOptions = {}) {
      const record = evidenceSession.captureMeetingAppSnapshot(input, meetingAppRecordOptions(input, captureOptions));
      return operationWithEvidence('capture_meeting_app_snapshot', record, evidenceSession, captureOptions);
    },
    captureProviderWebhook(input = {}, payload, captureOptions = {}) {
      const record = evidenceSession.captureProviderWebhook(input, payload, captureOptions);
      return operationWithEvidence('capture_provider_webhook', record, evidenceSession, captureOptions);
    },
    async insertAnnotation(input = {}, markOptions = {}) {
      const result = await source.insertAnnotation({
        platform: key,
        ...input,
      }, markOptions);
      return operationWithEvidence('insert_annotation', result, evidenceSession, markOptions);
    },
    async insertMark(input = {}, markOptions = {}) {
      const result = await source.insertMark({
        platform: key,
        ...input,
      }, markOptions);
      return operationWithEvidence('insert_mark', result, evidenceSession, markOptions);
    },
    async insertMarks(inputs = [], markOptions = {}) {
      const rows = Array.isArray(inputs) ? inputs : [inputs];
      const result = await source.insertMarks(rows.map((input) => ({
        platform: key,
        ...input,
      })), markOptions);
      return operationWithEvidence('insert_marks', result, evidenceSession, markOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return source.importTranscript({
        platform: key,
        ...input,
      }, transcriptOptions);
    },
    startMeeting(input = {}) {
      return source.startMeeting({
        platform: key,
        ...input,
      });
    },
    endMeeting(input = {}) {
      return source.endMeeting(input);
    },
    summary(summaryOptions = {}) {
      return evidenceSession.summary(summaryOptions);
    },
    correlation(correlationOptions = {}) {
      return evidenceSession.correlation(correlationOptions);
    },
    strategy(strategyOptions = {}) {
      return evidenceSession.strategy(strategyOptions);
    },
    exportPackage(packageOptions = {}) {
      return evidenceSession.exportPackage(packageOptions);
    },
    verify(verifyOptions = {}) {
      return evidenceSession.verify(verifyOptions);
    },
    getState() {
      return {
        type: 'meeting_platform_live_adapter_state',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platform: key,
        source: source.getState(),
        evidence: evidenceSession.getState(),
      };
    },
    reset(nextState = {}) {
      return {
        source: source.reset(nextState.source ?? nextState.meeting_source ?? {}),
        evidence: evidenceSession.reset(),
      };
    },
  };
}

export function createMeetingPlatformLiveAdapterSuite(clientOrOptions = {}, options = {}) {
  const cache = new Map();
  const platforms = selectedPlatforms(options).map((platform) => normalizeMeetingPlatform(platform));

  function adapter(platform, adapterOptions = {}) {
    const key = normalizeMeetingPlatform(platform);
    const cacheKey = `${key}:${adapterOptions.id ?? adapterOptions.sessionId ?? adapterOptions.session_id ?? 'default'}`;
    if (!cache.has(cacheKey)) {
      cache.set(cacheKey, createMeetingPlatformLiveAdapter(key, clientOrOptions, {
        ...options,
        ...adapterOptions,
      }));
    }
    return cache.get(cacheKey);
  }

  return {
    schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
    platforms,
    adapter,
    liveAdapter: adapter,
    adapters(adapterOptions = {}) {
      return Object.fromEntries(platforms.map((platform) => [platform, adapter(platform, adapterOptions)]));
    },
    plan(platform, planOptions = {}) {
      return buildMeetingPlatformLiveAdapterPlan(platform, {
        ...options,
        ...planOptions,
      });
    },
    matrix(matrixOptions = {}) {
      return buildMeetingPlatformLiveAdapterMatrix({
        ...options,
        ...matrixOptions,
        platforms: matrixOptions.platforms ?? matrixOptions.platform_keys ?? platforms,
      });
    },
    summary(summaryOptions = {}) {
      const matrix = buildMeetingPlatformLiveAdapterMatrix({
        ...options,
        ...summaryOptions,
        platforms: summaryOptions.platforms ?? summaryOptions.platform_keys ?? platforms,
      });
      return compactObject({
        type: 'meeting_platform_live_adapter_suite_summary',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platform_count: matrix.platform_count,
        production_ready_count: matrix.production_ready_count,
        pilot_ready_count: matrix.pilot_ready_count,
        platforms: matrix.platforms,
        rows: matrix.rows,
      });
    },
    getState() {
      return {
        type: 'meeting_platform_live_adapter_suite_state',
        schema: MEETING_PLATFORM_LIVE_ADAPTER_MATRIX_SCHEMA,
        schema_version: MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION,
        platforms,
        adapter_count: cache.size,
        adapters: Object.fromEntries([...cache.entries()].map(([key, item]) => [key, item.getState()])),
      };
    },
    reset(nextState = {}) {
      const states = Object.fromEntries([...cache.entries()].map(([key, item]) => [key, item.reset(nextState[key] ?? {})]));
      cache.clear();
      return {
        removed_adapters: Object.keys(states).length,
        adapters: states,
      };
    },
  };
}
