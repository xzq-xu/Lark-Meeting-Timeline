import { compactObject } from '../index.mjs';
import { createMeetingSourceAggregator } from './meeting-source.mjs';
import { createMeetingPlatformEvidenceSession } from './platform-evidence-session.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA = 'meeting_platform_live_adapter';
export const MEETING_PLATFORM_LIVE_ADAPTER_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
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
