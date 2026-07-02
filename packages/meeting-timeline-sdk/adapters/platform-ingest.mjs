import { MeetingTimelineSdkError } from '../index.mjs';
import { applyMeetingSignals } from './core.mjs';
import { meetingPlatformEventAdapterFor } from './platform-registry.mjs';

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
