import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { applyMeetingSignals } from './core.mjs';
import { meetingPlatformEventAdapterFor } from './platform-registry.mjs';
import { createMeetingSignalReconciler } from './signal-reconciler.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function normalizeInputArgs(platformOrInput, payload, options) {
  if (typeof platformOrInput === 'object' && platformOrInput != null && !Array.isArray(platformOrInput)) {
    const input = platformOrInput;
    return {
      platform: firstNonEmpty(input.platform, input.provider, input.adapter),
      payload: firstNonEmpty(input.payload, input.body, input.event, input.raw, payload),
      options: {
        ...(input.options ?? {}),
        receivedAtMs: firstNonEmpty(
          input.receivedAtMs,
          input.received_at_ms,
          input.received_at,
          input.timestamp,
          input.ts,
          input.options?.receivedAtMs,
          input.options?.received_at_ms,
        ),
        normalizerOptions: input.normalizerOptions
          ?? input.normalizer_options
          ?? input.options?.normalizerOptions
          ?? input.options?.normalizer_options,
        applyOptions: input.applyOptions
          ?? input.apply_options
          ?? input.options?.applyOptions
          ?? input.options?.apply_options,
        ...options,
      },
    };
  }
  return { platform: platformOrInput, payload, options };
}

function normalizeContext(options = {}) {
  return {
    ...(options.normalizerOptions ?? {}),
    ...(options.normalizer_options ?? {}),
    receivedAtMs: firstNonEmpty(
      options.receivedAtMs,
      options.received_at_ms,
      options.received_at,
      options.normalizerOptions?.receivedAtMs,
      options.normalizer_options?.receivedAtMs,
    ),
  };
}

function applyContext(options = {}) {
  const {
    receivedAtMs,
    received_at_ms: receivedAtMsSnake,
    received_at: receivedAt,
    normalizerOptions,
    normalizer_options: normalizerOptionsSnake,
    applyOptions,
    apply_options: applyOptionsSnake,
    ...rest
  } = options;
  void receivedAtMs;
  void receivedAtMsSnake;
  void receivedAt;
  void normalizerOptions;
  void normalizerOptionsSnake;
  return {
    ...rest,
    ...(applyOptionsSnake ?? {}),
    ...(applyOptions ?? {}),
  };
}

function publicAdapter(adapter = {}) {
  return compactObject({
    key: adapter.key,
    source: adapter.source,
    aliases: adapter.aliases,
  });
}

function issue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function meetingKey(meeting = {}) {
  return [
    meeting.platform,
    meeting.meeting_id,
    meeting.external_meeting_id,
    meeting.meeting_url,
  ].filter(Boolean).join('|');
}

function meetingRefs(signals = []) {
  const seen = new Set();
  const meetings = [];
  for (const signal of signals) {
    if (!signal.meeting) continue;
    const key = meetingKey(signal.meeting);
    if (seen.has(key)) continue;
    seen.add(key);
    meetings.push(signal.meeting);
  }
  return meetings;
}

function subjectOf(signal = {}) {
  if (signal.type === 'participant_joined' || signal.type === 'participant_left') {
    return compactObject({
      participant_id: signal.participant_id,
      participant_name: signal.participant_name,
    });
  }
  if (signal.type === 'speaker_started' || signal.type === 'speaker_ended') {
    return compactObject({
      speaker_id: signal.speaker_id,
      speaker_name: signal.speaker_name,
    });
  }
  if (signal.type === 'artifact_ready') {
    return compactObject({
      artifact_kind: signal.artifact_kind,
      artifact_id: signal.artifact_id,
      artifact_url: signal.artifact_url,
    });
  }
  if (signal.type === 'subscription_lifecycle') {
    return compactObject({
      lifecycle_type: signal.lifecycle_type,
      subscription_id: signal.subscription_id,
      subscription_name: signal.subscription_name,
      expires_at_ms: signal.expires_at_ms,
    });
  }
  return undefined;
}

function signalPreview(signal = {}) {
  return compactObject({
    type: signal.type,
    occurred_at_ms: signal.occurred_at_ms,
    source: signal.source,
    source_event_id: signal.source_event_id,
    meeting: signal.meeting,
    platform: signal.platform,
    ...subjectOf(signal),
  });
}

function signalCoverage(signalTypes = []) {
  const set = new Set(signalTypes);
  return {
    realtime_axis: set.has('meeting_started') || set.has('meeting_ended'),
    meeting_start: set.has('meeting_started'),
    meeting_end: set.has('meeting_ended'),
    participant_track: set.has('participant_joined') || set.has('participant_left'),
    speaker_activity: set.has('speaker_started') || set.has('speaker_ended'),
    artifact_ready: set.has('artifact_ready'),
    subscription_lifecycle: set.has('subscription_lifecycle'),
  };
}

function diagnosticsIssues(normalized = {}) {
  const signals = normalized.signals ?? [];
  const issues = [];
  if (signals.length === 0) {
    issues.push(issue(
      'warning',
      'no_supported_signals',
      'Payload was accepted by the platform adapter but did not produce meeting timeline signals.',
    ));
  }
  const signalTypes = new Set(signals.map((signal) => signal.type));
  if (signals.length > 0 && [...signalTypes].every((type) => type === 'subscription_lifecycle')) {
    issues.push(issue(
      'info',
      'subscription_lifecycle_only',
      'Payload only reports subscription health and will not create or update a meeting axis.',
    ));
  }
  if (normalized.source !== 'local_detector') {
    const missingUrl = signals.find((signal) => (
      (signal.type === 'meeting_started' || signal.type === 'meeting_ended')
        && !signal.meeting?.meeting_url
    ));
    if (missingUrl) {
      issues.push(issue(
        'warning',
        'provider_event_missing_meeting_url',
        'Provider event can write a meeting axis, but URL-based reconciliation with local observers may be weaker.',
        { signal_type: missingUrl.type, meeting_id: missingUrl.meeting?.meeting_id },
      ));
    }
  }
  return issues;
}

export function normalizePlatformEvent(platformOrInput, payload, options = {}) {
  const input = normalizeInputArgs(platformOrInput, payload, options);
  const adapter = meetingPlatformEventAdapterFor(input.platform);
  if (!adapter) {
    throw new MeetingTimelineSdkError(`Unsupported meeting platform: ${String(input.platform || '(empty)')}`, {
      platform: input.platform,
    });
  }
  const signals = adapter.normalize(input.payload, normalizeContext(input.options));
  return {
    adapter,
    platform: adapter.key,
    source: adapter.source,
    signals,
  };
}

export function diagnosePlatformEvent(platformOrInput, payload, options = {}) {
  const input = normalizeInputArgs(platformOrInput, payload, options);
  try {
    const normalized = normalizePlatformEvent(platformOrInput, payload, options);
    const signalTypes = [...new Set(normalized.signals.map((signal) => signal.type))];
    const issues = diagnosticsIssues(normalized);
    return compactObject({
      ok: true,
      supported: true,
      actionable: normalized.signals.length > 0,
      platform: normalized.platform,
      source: normalized.source,
      adapter: publicAdapter(normalized.adapter),
      signal_count: normalized.signals.length,
      signal_types: signalTypes,
      coverage: signalCoverage(signalTypes),
      meetings: meetingRefs(normalized.signals),
      signals: normalized.signals.map(signalPreview),
      issues,
      raw_signals: options.includeRawSignals === true || options.include_raw_signals === true
        ? normalized.signals
        : undefined,
    });
  } catch (error) {
    const adapter = meetingPlatformEventAdapterFor(input.platform);
    return compactObject({
      ok: false,
      supported: Boolean(adapter),
      actionable: false,
      platform: adapter?.key ?? input.platform,
      source: adapter?.source,
      adapter: adapter ? publicAdapter(adapter) : undefined,
      error: error.message ?? String(error),
      details: error.details,
      issues: [issue(
        'error',
        adapter ? 'normalization_failed' : 'unsupported_platform',
        error.message ?? String(error),
        { platform: input.platform },
      )],
    });
  }
}

export async function ingestPlatformEvent(client, platformOrInput, payload, options = {}) {
  if (!client) {
    throw new MeetingTimelineSdkError('Meeting timeline client is required for ingestPlatformEvent');
  }
  const normalized = normalizePlatformEvent(platformOrInput, payload, options);
  const input = normalizeInputArgs(platformOrInput, payload, options);
  const results = await applyMeetingSignals(client, normalized.signals, applyContext(input.options));
  return {
    ...normalized,
    results,
  };
}

function reconcileContext(baseOptions = {}, ingestOptions = {}) {
  return {
    ...(baseOptions.reconcileOptions ?? {}),
    ...(baseOptions.reconcile_options ?? {}),
    ...(ingestOptions.reconcileOptions ?? {}),
    ...(ingestOptions.reconcile_options ?? {}),
  };
}

export function createReconciledPlatformEventIngestor(client, options = {}) {
  if (!client) {
    throw new MeetingTimelineSdkError('Meeting timeline client is required for createReconciledPlatformEventIngestor');
  }
  const reconciler = options.reconciler
    ?? options.signalReconciler
    ?? options.signal_reconciler
    ?? createMeetingSignalReconciler(options.reconcileOptions ?? options.reconcile_options);
  return {
    async ingest(platformOrInput, payload, ingestOptions = {}) {
      const normalized = normalizePlatformEvent(platformOrInput, payload, {
        ...options,
        ...ingestOptions,
      });
      const input = normalizeInputArgs(platformOrInput, payload, {
        ...options,
        ...ingestOptions,
      });
      const reconciliation = reconciler.reconcile(
        normalized.signals,
        reconcileContext(options, input.options),
      );
      const results = await applyMeetingSignals(client, reconciliation.signals, applyContext(input.options));
      return {
        ...normalized,
        rawSignals: normalized.signals,
        signals: reconciliation.signals,
        reconciliation,
        results,
      };
    },
    getState() {
      return reconciler.getState();
    },
    reset(nextState = {}) {
      return reconciler.reset(nextState);
    },
  };
}
