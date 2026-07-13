import {
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
} from './errors.mjs';
import {
  MeetingTimelineAnnotationProducer,
  createMeetingTimelineAnnotationProducer,
} from './producer.mjs';
import {
  buildMeetingAppAdapterIntegrationPackage,
  buildMeetingAppAdapterIntegrationPackageMatrix,
} from './adapters/meeting-app-adapter-integration-package.mjs';
import {
  createMeetingPlatformIntegrationRuntime,
  detectMeetingPlatformForBrowser,
} from './adapters/platform-integration-runtime.mjs';
import {
  createMeetingPlatformRuntimeEventClient,
} from './adapters/platform-runtime-event.mjs';
import {
  assertMeetingPlatformProviderReplayMatrix,
  assertMeetingPlatformProviderReplayReport,
  buildMeetingPlatformProviderReplayMatrix,
  buildMeetingPlatformProviderReplayReport,
} from './adapters/platform-ingest.mjs';
import {
  assertMeetingAppTimelineConnectorAdapterMatrix,
  assertMeetingAppTimelineHostAdapterConfig,
  assertMeetingAppTimelineHostAdapterBootstrapPlan,
  assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  assertMeetingAppTimelineHostAdapterConfigIndex,
  assertMeetingAppTimelineResolvedHostAdapterConfig,
  assertMeetingAppTimelineConnectorHostWiringGuide,
  assertMeetingAppTimelineConnectorPlatformRoadmap,
  assertMeetingAppTimelineConnectorReleaseGate,
  buildMeetingAppTimelineConnectorAdapterMatrix,
  buildMeetingAppTimelineConnectorHostWiringGuide,
  buildMeetingAppTimelineConnectorHostWiringGuideAcceptanceReport,
  buildMeetingAppTimelineHostAdapterBootstrapPlan,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport,
  buildMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineHostAdapterConfigIndex,
  buildMeetingAppTimelineConnectorPlatformRoadmap,
  buildMeetingAppTimelineConnectorReleaseGate,
  resolveMeetingAppTimelineHostAdapterConfig,
} from './adapters/meeting-app-connector-package.mjs';

export {
  MeetingTimelineAnnotationProducer,
  MeetingTimelineApiError,
  MeetingTimelineSdkError,
  createMeetingTimelineAnnotationProducer,
};

export const SDK_VERSION = '0.1.0';

export const DEFAULT_ENDPOINTS = Object.freeze({
  state: '/api/state',
  ingestInfo: '/api/annotation-ingest-info',
  meetingSessionStatus: '/api/meeting-session/status',
  meetingSessionStart: '/api/meeting-session/start',
  meetingSessionEnd: '/api/meeting-session/end',
  annotation: '/api/annotations',
  annotationBatch: '/api/annotations/batch',
  annotationStatus: '/api/annotations/status',
  transcriptImport: '/api/import/transcript',
  stream: '/api/stream',
});

export const DEFAULT_MEETING_APP_TIMELINE_SDK_PLATFORMS = Object.freeze([
  'google-meet',
  'teams',
  'zoom',
  'webex',
  'lark',
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function uniqueStrings(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

export function compactObject(value) {
  if (Array.isArray(value)) return value.map((item) => compactObject(item));
  if (!isPlainObject(value)) return value;
  return Object.fromEntries(
    Object.entries(value)
      .filter(([, item]) => item !== undefined)
      .map(([key, item]) => [key, compactObject(item)]),
  );
}

export function normalizeAbsoluteMs(value, fieldName = 'timestamp') {
  if (value == null || value === '') return undefined;
  if (value instanceof Date) {
    const ms = value.getTime();
    if (Number.isFinite(ms)) return ms;
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: Date is not finite`, { fieldName, value });
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new MeetingTimelineSdkError(`Invalid ${fieldName}: number is not finite`, { fieldName, value });
    }
    if (value > 10_000_000_000_000) return Math.round(value / 1000);
    if (value > 10_000_000_000) return Math.round(value);
    if (value > 1_000_000_000) return Math.round(value * 1000);
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: expected absolute unix time`, { fieldName, value });
  }
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return normalizeAbsoluteMs(numeric, fieldName);
  const parsed = Date.parse(String(value));
  if (Number.isFinite(parsed)) return parsed;
  throw new MeetingTimelineSdkError(`Invalid ${fieldName}: cannot parse absolute time`, { fieldName, value });
}

function maybeAbsoluteMs(...values) {
  const value = firstNonEmpty(...values);
  return value == null ? undefined : normalizeAbsoluteMs(value);
}

function normalizeCandidates(input = {}) {
  const candidates = firstNonEmpty(
    input.text_candidates,
    input.textCandidates,
    input.candidates,
    input.ocr_candidates,
    input.ocrCandidates,
    input.payload?.text_candidates,
  );
  if (candidates == null) return undefined;
  return Array.isArray(candidates) ? candidates.map((item) => String(item)) : [String(candidates)];
}

export function buildMeetingStartPayload(input = {}, defaults = {}) {
  const startTimeMs = maybeAbsoluteMs(
    input.start_time_ms,
    input.startTimeMs,
    input.started_at_ms,
    input.startedAtMs,
    input.start_time,
    input.startTime,
    input.started_at,
    input.startedAt,
  );
  return compactObject({
    platform: input.platform ?? defaults.platform ?? 'lark',
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId, input.id, input.session_id, input.sessionId),
    external_meeting_id: firstNonEmpty(
      input.external_meeting_id,
      input.externalMeetingId,
      input.meeting_no,
      input.meetingNo,
      input.lark_meeting_id,
      input.open_meeting_id,
    ),
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, input.url, input.join_url, input.joinUrl),
    minute_token: firstNonEmpty(input.minute_token, input.minuteToken),
    title: firstNonEmpty(input.title, input.topic, input.name),
    timezone: input.timezone ?? defaults.timezone,
    start_time_ms: startTimeMs,
    end_time: input.end_time ?? input.endTime,
    detector_source: input.detector_source ?? input.detectorSource ?? input.source ?? defaults.detectorSource,
    note: input.note,
    force: input.force ?? input.forceMeetingSession ?? defaults.force,
    keep_segments: input.keep_segments ?? input.keepSegments,
    suppress_auto_annotations: input.suppress_auto_annotations ?? input.suppressAutoAnnotations,
    suppress_demo_annotations: input.suppress_demo_annotations ?? input.suppressDemoAnnotations,
    observer_surface: input.observer_surface ?? input.observerSurface,
    run_id: input.run_id ?? input.runId,
    meeting_app_record: input.meeting_app_record ?? input.meetingAppRecord,
  });
}

export function buildMeetingEndPayload(input = {}, defaults = {}) {
  const endTimeMs = maybeAbsoluteMs(
    input.end_time_ms,
    input.endTimeMs,
    input.ended_at_ms,
    input.endedAtMs,
    input.end_time,
    input.endTime,
    input.ended_at,
    input.endedAt,
  );
  return compactObject({
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId),
    end_time_ms: endTimeMs,
    time_ms: input.time_ms ?? input.timeMs,
    detector_source: input.detector_source ?? input.detectorSource ?? input.source ?? defaults.detectorSource,
    observer_surface: input.observer_surface ?? input.observerSurface,
    run_id: input.run_id ?? input.runId,
    meeting_app_record: input.meeting_app_record ?? input.meetingAppRecord,
  });
}

function normalizeIsoTime(value) {
  const ms = maybeAbsoluteMs(value);
  return ms == null ? undefined : new Date(ms).toISOString();
}

function buildTranscriptMeetingPayload(input = {}, defaults = {}) {
  const startTime = firstNonEmpty(input.start_time, input.startTime, input.started_at, input.startedAt);
  const startTimeMs = firstNonEmpty(input.start_time_ms, input.startTimeMs, input.started_at_ms, input.startedAtMs);
  const endTime = firstNonEmpty(input.end_time, input.endTime, input.ended_at, input.endedAt);
  const endTimeMs = firstNonEmpty(input.end_time_ms, input.endTimeMs, input.ended_at_ms, input.endedAtMs);
  return compactObject({
    platform: input.platform ?? defaults.platform,
    meeting_id: firstNonEmpty(input.meeting_id, input.meetingId, input.id, defaults.meeting_id, defaults.meetingId),
    external_meeting_id: firstNonEmpty(input.external_meeting_id, input.externalMeetingId, input.externalId),
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, input.url, input.join_url, input.joinUrl),
    minute_token: firstNonEmpty(input.minute_token, input.minuteToken),
    title: firstNonEmpty(input.title, input.topic, input.name),
    timezone: input.timezone ?? defaults.timezone,
    start_time: startTime != null ? normalizeIsoTime(startTime) : normalizeIsoTime(startTimeMs),
    end_time: endTime != null ? normalizeIsoTime(endTime) : normalizeIsoTime(endTimeMs),
    source: input.source ?? defaults.source,
  });
}

function buildTranscriptSegment(input = {}, index = 0, defaults = {}) {
  return compactObject({
    id: firstNonEmpty(input.id, input.segment_id, input.segmentId, input.sentence_id, input.sentenceId, `seg-${index + 1}`),
    start_ms: firstNonEmpty(input.start_ms, input.startMs, input.start_time_ms, input.startTimeMs, input.offset_ms, input.offsetMs),
    end_ms: firstNonEmpty(input.end_ms, input.endMs, input.end_time_ms, input.endTimeMs),
    start_time: firstNonEmpty(input.start_time, input.startTime, input.start, input.ts),
    end_time: firstNonEmpty(input.end_time, input.endTime, input.end),
    speaker_id: firstNonEmpty(input.speaker_id, input.speakerId, input.participant_id, input.participantId, input.user_id, input.userId),
    speaker_name: firstNonEmpty(input.speaker_name, input.speakerName, input.participant_name, input.participantName, input.user_name, input.userName, input.name),
    text: firstNonEmpty(input.text, input.content, input.sentence, input.transcript),
    language: firstNonEmpty(input.language, input.language_code, input.languageCode),
    source: input.source ?? defaults.source ?? 'transcript_import',
    raw: input.raw,
  });
}

export function buildTranscriptImportPayload(input = {}, defaults = {}) {
  const meetingInput = input.meeting ?? input.meetingSession ?? input.session ?? {};
  const rows = firstNonEmpty(input.transcript, input.segments, input.entries, input.items, []);
  const transcript = Array.isArray(rows)
    ? rows.map((item, index) => buildTranscriptSegment(item, index, {
      source: input.source ?? defaults.source,
    }))
    : rows;
  return compactObject({
    meeting: buildTranscriptMeetingPayload(meetingInput, {
      platform: input.platform ?? defaults.platform,
      source: input.source ?? defaults.source,
      meeting_id: input.meeting_id ?? input.meetingId,
    }),
    transcript,
    artifact: input.artifact,
    source: input.source ?? defaults.source,
  });
}

export function buildTimelineMark(input = {}, defaults = {}) {
  const requireCapturedAt = defaults.requireCapturedAt ?? true;
  const capturedAtInput = firstNonEmpty(
    input.captured_at_ms,
    input.capturedAtMs,
    input.captured_at,
    input.capturedAt,
    input.ink_end_at_ms,
    input.inkEndAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.ts,
  );
  if (capturedAtInput == null && requireCapturedAt) {
    throw new MeetingTimelineSdkError('captured_at_ms is required for reliable timeline insertion', {
      fieldName: 'captured_at_ms',
    });
  }
  const capturedAtMs = capturedAtInput == null ? undefined : normalizeAbsoluteMs(capturedAtInput, 'captured_at_ms');
  const textCandidates = normalizeCandidates(input);
  const primaryText = firstNonEmpty(
    input.text,
    input.reading,
    textCandidates?.[0],
    input.payload?.text,
    input.payload?.reading,
  );
  const kind = String(firstNonEmpty(input.kind, input.type, input.intent, input.mark?.action, input.action, defaults.kind, 'annotation'));
  const label = String(firstNonEmpty(input.label, primaryText, input.intent, input.mark?.target_text, kind, '实时标注'));
  const meetingSession = input.meeting_session ?? input.meetingSession ?? defaults.meetingSession;
  return compactObject({
    id: firstNonEmpty(input.id, input.annotation_id, input.annotationId, input.mark_id, input.markId),
    source: input.source ?? defaults.source ?? input.device?.type ?? 'external_annotation',
    device_id: firstNonEmpty(input.device_id, input.deviceId, defaults.deviceId),
    captured_at_ms: capturedAtMs,
    time_ms: input.time_ms ?? input.timeMs,
    kind,
    label,
    text: primaryText,
    text_candidates: textCandidates,
    intent: input.intent ?? input.payload?.intent,
    mark: input.mark ?? input.payload?.mark,
    target: input.target ?? input.target_region ?? input.targetRegion ?? input.payload?.target,
    strokes: input.strokes ?? input.stroke_points ?? input.strokePoints ?? input.payload?.strokes,
    device: input.device ?? (defaults.deviceId ? { id: defaults.deviceId } : undefined),
    payload: input.payload,
    realtime: input.realtime ?? defaults.realtime,
    live: input.live ?? defaults.live,
    meeting_session: meetingSession ? buildMeetingStartPayload(meetingSession, defaults.meetingDefaults ?? defaults) : undefined,
    start_meeting_session: input.start_meeting_session ?? input.startMeetingSession,
    force_meeting_session: input.force_meeting_session ?? input.forceMeetingSession,
  });
}

function normalizeBaseUrl(baseUrl) {
  if (!baseUrl) throw new MeetingTimelineSdkError('baseUrl is required');
  return String(baseUrl).replace(/\/+$/, '');
}

function buildUrl(baseUrl, path, query = {}) {
  const url = new URL(path, `${normalizeBaseUrl(baseUrl)}/`);
  for (const [key, value] of Object.entries(query)) {
    if (value != null && value !== '') url.searchParams.set(key, String(value));
  }
  return url;
}

export class MeetingTimelineClient {
  constructor(options = {}) {
    this.baseUrl = normalizeBaseUrl(options.baseUrl);
    this.fetchImpl = options.fetch ?? options.fetchImpl ?? globalThis.fetch;
    if (typeof this.fetchImpl !== 'function') {
      throw new MeetingTimelineSdkError('fetch implementation is required');
    }
    this.headers = options.headers ?? {};
    this.authToken = options.authToken ?? options.token ?? null;
    this.source = options.source ?? 'external_annotation';
    this.deviceId = options.deviceId ?? options.device_id ?? null;
    this.detectorSource = options.detectorSource ?? options.detector_source ?? this.source;
    this.requireCapturedAt = options.requireCapturedAt ?? true;
    this.timeoutMs = options.timeoutMs ?? 30_000;
  }

  async request(path, options = {}) {
    const url = buildUrl(this.baseUrl, path, options.query);
    const headers = {
      accept: 'application/json',
      ...this.headers,
      ...(options.headers ?? {}),
    };
    const init = {
      method: options.method ?? 'GET',
      headers,
      signal: options.signal,
    };
    if (options.body !== undefined) {
      init.headers = { 'content-type': 'application/json; charset=utf-8', ...headers };
      init.body = JSON.stringify(options.body);
    }
    if (this.authToken && !init.headers.authorization) {
      init.headers.authorization = `Bearer ${this.authToken}`;
    }
    const response = await this.fetchImpl(String(url), init);
    const text = await response.text();
    let body = null;
    if (text) {
      try {
        body = JSON.parse(text);
      } catch {
        body = text;
      }
    }
    if (!response.ok) {
      throw new MeetingTimelineApiError(`Meeting timeline API request failed: ${response.status}`, {
        status: response.status,
        body,
        url: String(url),
      });
    }
    return body;
  }

  getState() {
    return this.request(DEFAULT_ENDPOINTS.state);
  }

  getIngestInfo() {
    return this.request(DEFAULT_ENDPOINTS.ingestInfo);
  }

  getMeetingSessionStatus() {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionStatus);
  }

  startMeeting(input = {}) {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionStart, {
      method: 'POST',
      body: buildMeetingStartPayload(input, {
        detectorSource: this.detectorSource,
      }),
    });
  }

  endMeeting(input = {}) {
    return this.request(DEFAULT_ENDPOINTS.meetingSessionEnd, {
      method: 'POST',
      body: buildMeetingEndPayload(input, {
        detectorSource: this.detectorSource,
      }),
    });
  }

  insertMark(input = {}, options = {}) {
    const body = options.raw === true
      ? input
      : buildTimelineMark(input, {
        source: this.source,
        deviceId: this.deviceId,
        detectorSource: this.detectorSource,
        requireCapturedAt: options.requireCapturedAt ?? this.requireCapturedAt,
        realtime: options.realtime,
        live: options.live,
        meetingSession: options.meetingSession,
      });
    return this.request(DEFAULT_ENDPOINTS.annotation, {
      method: 'POST',
      body,
    });
  }

  addAnnotation(input = {}, options = {}) {
    return this.insertMark(input, options);
  }

  insertAnnotation(input = {}, options = {}) {
    return this.insertMark(input, options);
  }

  insertMarks(inputs = [], options = {}) {
    const rows = Array.isArray(inputs) ? inputs : inputs.annotations ?? inputs.items ?? [];
    const annotations = options.raw === true
      ? rows
      : rows.map((item) => buildTimelineMark(item, {
        source: this.source,
        deviceId: this.deviceId,
        detectorSource: this.detectorSource,
        requireCapturedAt: options.requireCapturedAt ?? this.requireCapturedAt,
        realtime: options.realtime,
        live: options.live,
      }));
    return this.request(DEFAULT_ENDPOINTS.annotationBatch, {
      method: 'POST',
      body: { annotations },
    });
  }

  addAnnotations(inputs = [], options = {}) {
    return this.insertMarks(inputs, options);
  }

  getAnnotationStatus(id) {
    return this.request(DEFAULT_ENDPOINTS.annotationStatus, {
      query: { id },
    });
  }

  importTranscript(input = {}, options = {}) {
    const body = options.raw === true
      ? input
      : buildTranscriptImportPayload(input, {
        source: options.source,
        platform: options.platform,
      });
    return this.request(options.path ?? DEFAULT_ENDPOINTS.transcriptImport, {
      method: 'POST',
      body,
    });
  }

  importMeetingTranscript(input = {}, options = {}) {
    return this.importTranscript(input, options);
  }

  subscribeState(options = {}) {
    const EventSourceImpl = options.EventSource ?? options.EventSourceImpl ?? globalThis.EventSource;
    if (typeof EventSourceImpl !== 'function') {
      throw new MeetingTimelineSdkError('EventSource implementation is required for subscribeState');
    }
    const url = buildUrl(this.baseUrl, options.path ?? DEFAULT_ENDPOINTS.stream, options.query);
    const stream = new EventSourceImpl(String(url), options.eventSourceOptions);
    const onState = options.onState;
    const onMessage = (event) => {
      const data = JSON.parse(event.data);
      if (typeof onState === 'function') onState(data, event);
    };
    stream.addEventListener?.('state', onMessage);
    stream.addEventListener?.('message', onMessage);
    if (typeof options.onError === 'function') {
      stream.addEventListener?.('error', options.onError);
      stream.onerror = options.onError;
    }
    return {
      stream,
      close() {
        stream.removeEventListener?.('state', onMessage);
        stream.removeEventListener?.('message', onMessage);
        stream.close();
      },
    };
  }
}

export function createMeetingTimelineClient(options = {}) {
  return new MeetingTimelineClient(options);
}

export const createMeetingTimelineSdk = createMeetingTimelineClient;

function hasMeetingTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.startMeeting === 'function'
    && typeof value.insertMark === 'function';
}

function timelineClientFromHostOptions(options = {}) {
  const explicitClient = options.client ?? options.timelineClient ?? options.timeline_client;
  if (hasMeetingTimelineClient(explicitClient)) return explicitClient;
  if (hasMeetingTimelineClient(options)) return options;
  return createMeetingTimelineClient({
    ...options,
    ...(options.clientOptions ?? options.client_options ?? {}),
  });
}

function meetingAppTimelineSdkPlatforms(options = {}) {
  return firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    DEFAULT_MEETING_APP_TIMELINE_SDK_PLATFORMS,
  );
}

function runtimeEventClientOptions(options = {}) {
  return compactObject({
    ...options,
    ...(options.runtimeEventClientOptions ?? options.runtime_event_client_options ?? {}),
  });
}

function shouldUseRuntimeEvent(options = {}, defaults = {}) {
  const mode = String(firstNonEmpty(
    options.mode,
    options.delivery,
    options.route,
    defaults.mode,
    defaults.delivery,
    defaults.route,
    '',
  )).toLowerCase().replace(/[-\s]+/g, '_');
  return options.remote === true
    || options.runtimeEvent === true
    || options.runtime_event === true
    || mode === 'remote'
    || mode === 'runtime_event'
    || mode === 'runtime_events';
}

function isConnectorHostAdapterSource(value = {}) {
  return [
    'meeting_app_timeline_connector_package',
    'meeting_app_timeline_connector_host_install_checklist',
    'meeting_app_timeline_connector_adapter_matrix',
    'meeting_app_timeline_host_adapter_config_index',
  ].includes(value?.schema);
}

function platformAndInput(runtime, platformOrInput, inputOrOptions = {}, options = {}) {
  if (typeof platformOrInput === 'string') {
    return {
      platform: platformOrInput,
      input: inputOrOptions ?? {},
      options,
    };
  }
  const input = platformOrInput ?? {};
  const methodOptions = inputOrOptions ?? {};
  const resolution = runtime.resolvePlatform(input, methodOptions);
  return {
    platform: firstNonEmpty(methodOptions.platform, input.platform, input.provider, resolution.platform),
    input,
    options: methodOptions,
    resolution,
  };
}

function sdkPlatformOptions(runtime, options = {}) {
  return {
    ...options,
    platforms: firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys, runtime.platforms),
  };
}

function singlePlatformInput(runtime, platformOrOptions = {}, options = {}) {
  if (typeof platformOrOptions === 'string') {
    return {
      platform: platformOrOptions,
      options,
    };
  }
  const input = platformOrOptions ?? {};
  return {
    platform: firstNonEmpty(options.platform, options.provider, input.platform, input.provider, input.key, runtime.platforms?.[0]),
    options: {
      ...input,
      ...options,
    },
  };
}

function normalizeConnectorSurface(surface = 'browser_extension') {
  const key = String(surface || 'browser_extension').trim().toLowerCase().replace(/[-\s]+/g, '_');
  const aliases = {
    extension: 'browser_extension',
    browser: 'browser_extension',
    chrome_extension: 'browser_extension',
    edge_extension: 'browser_extension',
    electron: 'electron_webview',
    electron_preload: 'electron_webview',
    embedded_webview: 'webview',
    native: 'native_detector',
    local_detector: 'native_detector',
    desktop_detector: 'native_detector',
  };
  return aliases[key] ?? key;
}

function connectorSurfaces(options = {}) {
  return uniqueStrings(asArray(firstNonEmpty(
    options.surfaces,
    options.surface_keys,
    options.surfaceKeys,
    options.targetSurfaces,
    options.target_surfaces,
    options.surface,
    options.targetSurface,
    options.target_surface,
    ['browser_extension'],
  )).map((surface) => normalizeConnectorSurface(surface)));
}

function includeExtensionScaffold(options = {}, surfaces = []) {
  if (options.includeExtensionScaffold === false || options.include_extension_scaffold === false) return false;
  return surfaces.includes('browser_extension');
}

function connectorExtensionOptions(defaults = {}, options = {}) {
  return {
    ...defaults,
    ...(options.extensionOptions ?? {}),
    ...(options.extension_options ?? {}),
    platforms: firstNonEmpty(
      options.extensionOptions?.platforms,
      options.extension_options?.platforms,
      options.platforms,
      options.platform_keys,
      options.platformKeys,
      defaults.platforms,
    ),
  };
}

function connectorNextActions(...groups) {
  const visit = (group) => {
    if (Array.isArray(group)) return group.flatMap((item) => visit(item));
    if (typeof group === 'string') return [group];
    if (group?.next_actions) return visit(group.next_actions);
    return [];
  };
  return uniqueStrings(groups.flatMap((group) => visit(group)));
}

export function createMeetingAppTimelineSdk(options = {}) {
  const client = timelineClientFromHostOptions(options);
  const platforms = meetingAppTimelineSdkPlatforms(options);
  const defaults = compactObject({
    ...options,
    platforms,
  });
  const runtime = createMeetingPlatformIntegrationRuntime(client, defaults);
  const runtimeEvents = createMeetingPlatformRuntimeEventClient(runtimeEventClientOptions({
    ...defaults,
    fetch: options.runtimeEventFetch ?? options.runtime_event_fetch ?? options.fetch,
    fetchImpl: options.runtimeEventFetchImpl ?? options.runtime_event_fetch_impl ?? options.fetchImpl,
  }));
  const sdk = {
    type: 'meeting_app_timeline_sdk',
    schema: 'meeting_app_timeline_sdk',
    schema_version: 1,
    platforms: runtime.platforms,
    client,
    kit: runtime.kit,
    runtime,
    integrationRuntime: runtime,
    integration_runtime: runtime,
    runtimeEvents,
    runtime_events: runtimeEvents,
    detect(input = {}, detectOptions = {}) {
      return detectMeetingPlatformForBrowser(input, {
        ...defaults,
        ...detectOptions,
      });
    },
    resolve(input = {}, resolveOptions = {}) {
      return runtime.resolvePlatform(input, resolveOptions);
    },
    resolveCandidates(input = {}, resolveOptions = {}) {
      return runtime.resolvePlatformCandidates(input, resolveOptions);
    },
    package(platformOrOptions = {}, packageOptions = {}) {
      if (typeof platformOrOptions === 'string') {
        return buildMeetingAppAdapterIntegrationPackage(platformOrOptions, {
          ...defaults,
          ...packageOptions,
        });
      }
      return buildMeetingAppAdapterIntegrationPackage({
        ...defaults,
        ...platformOrOptions,
      }, packageOptions);
    },
    integrationPackage(platformOrOptions = {}, packageOptions = {}) {
      return sdk.package(platformOrOptions, packageOptions);
    },
    packageMatrix(packageOptions = {}) {
      return buildMeetingAppAdapterIntegrationPackageMatrix({
        ...defaults,
        ...packageOptions,
        platforms: packageOptions.platforms ?? packageOptions.platform_keys ?? runtime.platforms,
      });
    },
    integrationPackageMatrix(packageOptions = {}) {
      return sdk.packageMatrix(packageOptions);
    },
    integrationProfile(platformOrOptions = {}, profileOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, profileOptions);
      return runtime.kit.meetingAppIntegrationProfile(resolved.platform, resolved.options);
    },
    integrationMatrix(profileOptions = {}) {
      return runtime.kit.meetingAppIntegrationMatrix(sdkPlatformOptions(runtime, profileOptions));
    },
    runtimeAdapterProfile(input = {}, profileOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterProfile(input, sdkPlatformOptions(runtime, profileOptions));
    },
    adapterProfile(input = {}, profileOptions = {}) {
      return sdk.runtimeAdapterProfile(input, profileOptions);
    },
    runtimeAdapterProfileMatrix(profileOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterProfileMatrix(sdkPlatformOptions(runtime, profileOptions));
    },
    adapterProfileMatrix(profileOptions = {}) {
      return sdk.runtimeAdapterProfileMatrix(profileOptions);
    },
    observerPlan(platformOrInput = {}, planOptions = {}) {
      return runtime.kit.meetingAppRuntimeObserverPlan(platformOrInput, sdkPlatformOptions(runtime, planOptions));
    },
    runtimeObserverPlan(platformOrInput = {}, planOptions = {}) {
      return sdk.observerPlan(platformOrInput, planOptions);
    },
    observerPlanMatrix(planOptions = {}) {
      return runtime.kit.meetingAppRuntimeObserverPlanMatrix(sdkPlatformOptions(runtime, planOptions));
    },
    runtimeObserverPlanMatrix(planOptions = {}) {
      return sdk.observerPlanMatrix(planOptions);
    },
    selectAdapter(input = {}, selectionOptions = {}) {
      return runtime.kit.selectMeetingAppRuntimeAdapter(input, sdkPlatformOptions(runtime, selectionOptions));
    },
    selectRuntimeAdapter(input = {}, selectionOptions = {}) {
      return sdk.selectAdapter(input, selectionOptions);
    },
    runtimeAdapterHandoff(selectionOrInput = {}, handoffOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterHandoff(selectionOrInput, sdkPlatformOptions(runtime, handoffOptions));
    },
    adapterHandoff(selectionOrInput = {}, handoffOptions = {}) {
      return sdk.runtimeAdapterHandoff(selectionOrInput, handoffOptions);
    },
    handoff(selectionOrInput = {}, handoffOptions = {}) {
      return sdk.runtimeAdapterHandoff(selectionOrInput, handoffOptions);
    },
    runtimeAdapterHandoffMatrix(handoffOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterHandoffMatrix(sdkPlatformOptions(runtime, handoffOptions));
    },
    adapterHandoffMatrix(handoffOptions = {}) {
      return sdk.runtimeAdapterHandoffMatrix(handoffOptions);
    },
    handoffMatrix(handoffOptions = {}) {
      return sdk.runtimeAdapterHandoffMatrix(handoffOptions);
    },
    runtimeAdapterHandoffAcceptance(handoffOrInput = {}, acceptanceOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterHandoffAcceptance(
        handoffOrInput,
        sdkPlatformOptions(runtime, acceptanceOptions),
      );
    },
    handoffAcceptance(handoffOrInput = {}, acceptanceOptions = {}) {
      return sdk.runtimeAdapterHandoffAcceptance(handoffOrInput, acceptanceOptions);
    },
    runtimeAdapterHandoffMatrixAcceptance(matrixOrOptions = {}, acceptanceOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterHandoffMatrixAcceptance(
        matrixOrOptions,
        sdkPlatformOptions(runtime, acceptanceOptions),
      );
    },
    handoffMatrixAcceptance(matrixOrOptions = {}, acceptanceOptions = {}) {
      return sdk.runtimeAdapterHandoffMatrixAcceptance(matrixOrOptions, acceptanceOptions);
    },
    runtimeAdapterHostPackage(hostPackageOptions = {}) {
      return runtime.kit.meetingAppRuntimeAdapterHostPackage(sdkPlatformOptions(runtime, hostPackageOptions));
    },
    hostPackage(hostPackageOptions = {}) {
      return sdk.runtimeAdapterHostPackage(hostPackageOptions);
    },
    connectorPackage(connectorOptions = {}) {
      const surfaces = connectorSurfaces(connectorOptions);
      const merged = sdkPlatformOptions(runtime, {
        ...connectorOptions,
        surfaces,
      });
      const hostPackage = sdk.hostPackage(merged);
      const handoffMatrix = hostPackage.handoff_matrix ?? sdk.handoffMatrix(merged);
      const handoffAcceptance = hostPackage.handoff_acceptance ?? sdk.handoffMatrixAcceptance(handoffMatrix, merged);
      const observerPlanBySurface = Object.fromEntries(surfaces.map((surface) => [
        surface,
        sdk.observerPlanMatrix({
          ...merged,
          surface,
        }),
      ]));
      const schedulerConfigBySurface = Object.fromEntries(surfaces.map((surface) => [
        surface,
        runtime.kit.meetingAppObserverSchedulerConfigMatrix({
          ...merged,
          surface,
          observeTracks: firstNonEmpty(connectorOptions.observeTracks, connectorOptions.observe_tracks, true),
          observe_tracks: firstNonEmpty(connectorOptions.observe_tracks, connectorOptions.observeTracks, true),
        }),
      ]));
      const extensionScaffold = includeExtensionScaffold(connectorOptions, surfaces)
        ? runtime.kit.meetingAppExtensionScaffold(connectorExtensionOptions(merged, connectorOptions))
        : undefined;
      const extensionAcceptance = extensionScaffold
        ? runtime.kit.meetingAppExtensionAcceptance(connectorExtensionOptions(merged, connectorOptions))
        : undefined;
      const runtimeEventPlanMatrix = runtime.kit.platformRuntimeEventPlanMatrix(merged);
      const adapterBlueprintMatrix = runtime.kit.platformAdapterBlueprintMatrix(merged);
      const providerReplayMatrix = sdk.providerReplayMatrix({
        ...merged,
        target: firstNonEmpty(connectorOptions.providerReplayTarget, connectorOptions.provider_replay_target, 'pilot'),
        recordsByPlatform: firstNonEmpty(
          connectorOptions.providerRecordsByPlatform,
          connectorOptions.provider_records_by_platform,
          connectorOptions.providerEventsByPlatform,
          connectorOptions.provider_events_by_platform,
        ),
      });
      const startupPlanInput = firstNonEmpty(
        connectorOptions.startupInput,
        connectorOptions.startup_input,
        {},
      );
      const startupPlanMatrix = runtime.kit.platformAdapterStartupPlanMatrix(startupPlanInput, merged);
      const platformOnboardingMatrix = runtime.kit.platformAdapterAcceptanceChecklistMatrix({
        platforms: hostPackage.platforms,
      }, {
        ...merged,
        target: firstNonEmpty(
          connectorOptions.onboardingTarget,
          connectorOptions.onboarding_target,
          'static',
        ),
      });
      const accepted = hostPackage.accepted === true
        && handoffAcceptance.accepted === true
        && adapterBlueprintMatrix.ready_count === adapterBlueprintMatrix.platform_count
        && startupPlanMatrix.realtime_startup_ready_count === startupPlanMatrix.platform_count
        && platformOnboardingMatrix.blocked_count === 0
        && (extensionAcceptance ? extensionAcceptance.accepted === true : true);
      return compactObject({
        type: 'meeting_app_timeline_connector_package',
        schema: 'meeting_app_timeline_connector_package',
        schema_version: 1,
        id: firstNonEmpty(
          connectorOptions.packageId,
          connectorOptions.package_id,
          `${hostPackage.id ?? 'meeting-app-runtime-adapter'}-connector`,
        ),
        base_url: firstNonEmpty(connectorOptions.baseUrl, connectorOptions.base_url, defaults.baseUrl, defaults.base_url),
        accepted,
        platforms: hostPackage.platforms,
        surfaces: hostPackage.surfaces ?? surfaces,
        platform_count: hostPackage.platform_count,
        surface_count: surfaces.length,
        handoff_count: hostPackage.handoff_count,
        ready_count: hostPackage.ready_count,
        host_package: hostPackage,
        handoff_matrix: handoffMatrix,
        handoff_acceptance: handoffAcceptance,
        observer_plan_by_surface: observerPlanBySurface,
        scheduler_config_by_surface: schedulerConfigBySurface,
        extension: extensionScaffold ? {
          scaffold: extensionScaffold,
          acceptance: extensionAcceptance,
          install_plan: extensionScaffold.install_plan,
          manifest: extensionScaffold.manifest,
          bundle: extensionScaffold.bundle,
          file_count: extensionScaffold.files?.length ?? 0,
        } : undefined,
        runtime_events: {
          endpoint: runtimeEvents.endpoint,
          plan_matrix: runtimeEventPlanMatrix,
          action_count: runtimeEventPlanMatrix.rows?.length ?? 0,
        },
        provider_replay: {
          matrix: providerReplayMatrix,
          accepted: providerReplayMatrix.accepted === true,
          accepted_count: providerReplayMatrix.accepted_count,
          platform_count: providerReplayMatrix.platform_count,
          runtime_event_count: providerReplayMatrix.runtime_event_count,
          command: 'npm run meeting-platform:provider-replay',
          bin: 'meeting-platform-provider-replay',
          sdk_method: 'sdk.providerReplayMatrix(options)',
          realtime_policy: {
            provider_events_block_realtime: false,
            local_observer_remains_primary_realtime_axis: true,
          },
        },
        platform_onboarding: {
          matrix: platformOnboardingMatrix,
          accepted: platformOnboardingMatrix.blocked_count === 0,
          target: platformOnboardingMatrix.target,
          accepted_count: platformOnboardingMatrix.accepted_count,
          blocked_count: platformOnboardingMatrix.blocked_count,
          platform_count: platformOnboardingMatrix.platform_count,
          command: 'npm run meeting-platform:adapter-acceptance-checklist',
          sdk_method: 'sdk.platformAdapterAcceptanceChecklistMatrix({ platforms }, { target })',
        },
        adapter_blueprints: {
          matrix: {
            ...adapterBlueprintMatrix,
            blueprints: connectorOptions.includeDetails === true || connectorOptions.include_details === true
              ? adapterBlueprintMatrix.blueprints
              : undefined,
          },
          ready_count: adapterBlueprintMatrix.ready_count,
          platform_count: adapterBlueprintMatrix.platform_count,
          command: 'npm run meeting-platform:adapter-blueprint',
          sdk_method: 'sdk.platformAdapterBlueprint(platform)',
        },
        startup_plans: {
          matrix: {
            ...startupPlanMatrix,
            plans: connectorOptions.includeDetails === true || connectorOptions.include_details === true
              ? startupPlanMatrix.plans
              : undefined,
          },
          realtime_startup_ready_count: startupPlanMatrix.realtime_startup_ready_count,
          platform_count: startupPlanMatrix.platform_count,
          sdk_method: 'sdk.platformAdapterStartupPlan(input)',
          matrix_sdk_method: 'sdk.platformAdapterStartupPlanMatrix(input)',
        },
        entrypoints: [
          {
            id: 'select-adapter',
            method: 'sdk.selectAdapter(input)',
            output: 'meeting_app_runtime_adapter_selection',
          },
          {
            id: 'handoff',
            method: 'sdk.handoff(selectionOrInput, { surface })',
            output: 'meeting_app_runtime_adapter_handoff',
          },
          {
            id: 'adapter-blueprint',
            method: 'sdk.platformAdapterBlueprint(platform)',
            output: 'meeting_platform_adapter_blueprint',
          },
          {
            id: 'startup-plan',
            method: 'sdk.platformAdapterStartupPlan(input)',
            output: 'meeting_platform_adapter_startup_plan',
          },
          {
            id: 'observe',
            method: 'sdk.observeMeetingApp(platformOrInput, input, { remote: true })',
            output: 'runtime_event_or_local_observation_result',
          },
          {
            id: 'speaker-participant-track',
            method: 'sdk.speakerTrack(...) / sdk.participantTrack(...)',
            output: 'position_markers_without_transcript_text',
          },
          {
            id: 'ci-gate',
            method: 'sdk.handoffMatrixAcceptance(matrix)',
            output: 'meeting_app_runtime_adapter_handoff_matrix_acceptance_report',
          },
        ],
        contracts: {
          timestamp_field: 'captured_at_ms',
          provider_events_block_realtime: false,
          transcript_blocks_realtime: false,
          local_observer_first: true,
          speaker_track_text_required: false,
          participant_track_text_required: false,
          provider_replay_required_for_provider_ci: true,
          adapter_blueprint_required_before_host_wiring: true,
          startup_plan_required_before_runtime_install: true,
          production_requires_live_snapshot: true,
        },
        next_actions: connectorNextActions(
          hostPackage,
          handoffAcceptance,
          runtimeEventPlanMatrix,
          extensionAcceptance,
          Object.values(observerPlanBySurface),
          Object.values(schedulerConfigBySurface),
          [
            'wire_connector_package_entrypoints_into_host_project',
            'wire_startup_plan_into_host_runtime',
            'run_connector_package_handoff_acceptance_in_ci',
            'collect_live_dom_snapshots_before_production_rollout',
          ],
        ),
      });
    },
    runtimeConnectorPackage(connectorOptions = {}) {
      return sdk.connectorPackage(connectorOptions);
    },
    connectorReleaseGate(connectorOptions = {}, releaseGateOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineConnectorReleaseGate(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...releaseGateOptions,
      });
    },
    assertConnectorReleaseGate(connectorOptions = {}, releaseGateOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineConnectorReleaseGate(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...releaseGateOptions,
      });
    },
    connectorPlatformRoadmap(connectorOptions = {}, roadmapOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineConnectorPlatformRoadmap(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...roadmapOptions,
      });
    },
    assertConnectorPlatformRoadmap(connectorOptions = {}, roadmapOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineConnectorPlatformRoadmap(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...roadmapOptions,
      });
    },
    connectorAdapterMatrix(connectorOptions = {}, matrixOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineConnectorAdapterMatrix(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...matrixOptions,
      });
    },
    assertConnectorAdapterMatrix(connectorOptions = {}, matrixOptions = {}) {
      const pkg = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineConnectorAdapterMatrix(pkg, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...matrixOptions,
      });
    },
    connectorHostAdapterConfig(connectorOrPlatform = {}, platformOrOptions = {}, configOptions = {}) {
      const platformFirst = typeof connectorOrPlatform === 'string';
      const sourceOptions = platformFirst ? platformOrOptions : connectorOrPlatform;
      const source = platformFirst
        ? sdk.connectorPackage(sourceOptions)
        : (
          connectorOrPlatform?.schema === 'meeting_app_timeline_connector_package'
          || connectorOrPlatform?.schema === 'meeting_app_timeline_connector_host_install_checklist'
          || connectorOrPlatform?.schema === 'meeting_app_timeline_connector_adapter_matrix'
            ? connectorOrPlatform
            : sdk.connectorPackage(connectorOrPlatform)
        );
      return buildMeetingAppTimelineHostAdapterConfig(source, platformFirst ? connectorOrPlatform : platformOrOptions, {
        ...sdkPlatformOptions(runtime, sourceOptions),
        ...configOptions,
      });
    },
    assertConnectorHostAdapterConfig(connectorOrPlatform = {}, platformOrOptions = {}, configOptions = {}) {
      const platformFirst = typeof connectorOrPlatform === 'string';
      const sourceOptions = platformFirst ? platformOrOptions : connectorOrPlatform;
      const source = platformFirst
        ? sdk.connectorPackage(sourceOptions)
        : (
          connectorOrPlatform?.schema === 'meeting_app_timeline_connector_package'
          || connectorOrPlatform?.schema === 'meeting_app_timeline_connector_host_install_checklist'
          || connectorOrPlatform?.schema === 'meeting_app_timeline_connector_adapter_matrix'
            ? connectorOrPlatform
            : sdk.connectorPackage(connectorOrPlatform)
        );
      return assertMeetingAppTimelineHostAdapterConfig(source, platformFirst ? connectorOrPlatform : platformOrOptions, {
        ...sdkPlatformOptions(runtime, sourceOptions),
        ...configOptions,
      });
    },
    connectorHostAdapterConfigIndex(connectorOptions = {}, indexOptions = {}) {
      const source = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_adapter_matrix'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineHostAdapterConfigIndex(source, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...indexOptions,
      });
    },
    assertConnectorHostAdapterConfigIndex(connectorOptions = {}, indexOptions = {}) {
      const source = connectorOptions?.schema === 'meeting_app_timeline_connector_package'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_host_install_checklist'
        || connectorOptions?.schema === 'meeting_app_timeline_connector_adapter_matrix'
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineHostAdapterConfigIndex(source, {
        ...sdkPlatformOptions(runtime, connectorOptions),
        ...indexOptions,
      });
    },
    connectorHostWiringGuide(connectorOptions = {}, guideOptions = {}) {
      const source = connectorOptions?.schema === 'meeting_app_timeline_connector_host_wiring_guide'
        || isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineConnectorHostWiringGuide(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...guideOptions,
      });
    },
    connectorHostWiringGuideAcceptanceReport(guideOrConnectorOptions = {}, acceptanceOptions = {}) {
      const source = guideOrConnectorOptions?.schema === 'meeting_app_timeline_connector_host_wiring_guide'
        ? guideOrConnectorOptions
        : (
          isConnectorHostAdapterSource(guideOrConnectorOptions)
            ? guideOrConnectorOptions
            : sdk.connectorPackage(guideOrConnectorOptions)
        );
      return buildMeetingAppTimelineConnectorHostWiringGuideAcceptanceReport(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(guideOrConnectorOptions) ? {} : guideOrConnectorOptions),
        ...acceptanceOptions,
      });
    },
    assertConnectorHostWiringGuide(connectorOptions = {}, guideOptions = {}) {
      const source = connectorOptions?.schema === 'meeting_app_timeline_connector_host_wiring_guide'
        || isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineConnectorHostWiringGuide(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...guideOptions,
      });
    },
    resolveConnectorHostAdapterConfig(input = {}, connectorOptions = {}, resolveOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return resolveMeetingAppTimelineHostAdapterConfig(source, input, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...resolveOptions,
      });
    },
    assertResolvedConnectorHostAdapterConfig(input = {}, connectorOptions = {}, resolveOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineResolvedHostAdapterConfig(source, input, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...resolveOptions,
      });
    },
    connectorHostAdapterBootstrapPlan(input = {}, connectorOptions = {}, bootstrapOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineHostAdapterBootstrapPlan(source, input, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...bootstrapOptions,
      });
    },
    assertConnectorHostAdapterBootstrapPlan(input = {}, connectorOptions = {}, bootstrapOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineHostAdapterBootstrapPlan(source, input, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...bootstrapOptions,
      });
    },
    connectorHostAdapterBootstrapPlanMatrix(connectorOptions = {}, matrixOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...matrixOptions,
      });
    },
    connectorHostAdapterBootstrapPlanMatrixAcceptanceReport(connectorOptions = {}, acceptanceOptions = {}) {
      const source = connectorOptions?.schema === 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix'
        ? connectorOptions
        : (
          isConnectorHostAdapterSource(connectorOptions)
            ? connectorOptions
            : sdk.connectorPackage(connectorOptions)
        );
      return buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...acceptanceOptions,
      });
    },
    assertConnectorHostAdapterBootstrapPlanMatrix(connectorOptions = {}, matrixOptions = {}) {
      const source = isConnectorHostAdapterSource(connectorOptions)
        ? connectorOptions
        : sdk.connectorPackage(connectorOptions);
      return assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix(source, {
        ...sdkPlatformOptions(runtime, isConnectorHostAdapterSource(connectorOptions) ? {} : connectorOptions),
        ...matrixOptions,
      });
    },
    providerReplayReport(platformOrInput = {}, recordsOrOptions = undefined, replayOptions = {}) {
      return buildMeetingPlatformProviderReplayReport(
        platformOrInput,
        recordsOrOptions,
        sdkPlatformOptions(runtime, replayOptions),
      );
    },
    assertProviderReplayReport(platformOrInput = {}, recordsOrOptions = undefined, replayOptions = {}) {
      return assertMeetingPlatformProviderReplayReport(
        sdk.providerReplayReport(platformOrInput, recordsOrOptions, replayOptions),
      );
    },
    providerReplayMatrix(replayOptions = {}) {
      return buildMeetingPlatformProviderReplayMatrix(sdkPlatformOptions(runtime, replayOptions));
    },
    assertProviderReplayMatrix(replayOptions = {}) {
      return assertMeetingPlatformProviderReplayMatrix(sdk.providerReplayMatrix(replayOptions));
    },
    meetingAppAdapterCapability(platformOrOptions = {}, input = {}, capabilityOptions = {}) {
      if (platformOrOptions && typeof platformOrOptions === 'object' && !Array.isArray(platformOrOptions)) {
        return runtime.kit.meetingAppAdapterCapability(sdkPlatformOptions(runtime, {
          ...platformOrOptions,
          ...input,
          ...capabilityOptions,
        }));
      }
      return runtime.kit.meetingAppAdapterCapability(
        platformOrOptions,
        input,
        sdkPlatformOptions(runtime, capabilityOptions),
      );
    },
    adapterCapability(platformOrOptions = {}, input = {}, capabilityOptions = {}) {
      return sdk.meetingAppAdapterCapability(platformOrOptions, input, capabilityOptions);
    },
    meetingAppAdapterCapabilityMatrix(capabilityOptions = {}) {
      return runtime.kit.meetingAppAdapterCapabilityMatrix(sdkPlatformOptions(runtime, capabilityOptions));
    },
    adapterCapabilityMatrix(capabilityOptions = {}) {
      return sdk.meetingAppAdapterCapabilityMatrix(capabilityOptions);
    },
    meetingAppAdapterExecutionPlan(platformOrCapability = {}, input = {}, planOptions = {}) {
      if (platformOrCapability && typeof platformOrCapability === 'object' && !Array.isArray(platformOrCapability)) {
        return runtime.kit.meetingAppAdapterExecutionPlan(
          platformOrCapability,
          sdkPlatformOptions(runtime, {
            ...input,
            ...planOptions,
          }),
        );
      }
      return runtime.kit.meetingAppAdapterExecutionPlan(
        platformOrCapability,
        input,
        sdkPlatformOptions(runtime, planOptions),
      );
    },
    adapterExecutionPlan(platformOrCapability = {}, input = {}, planOptions = {}) {
      return sdk.meetingAppAdapterExecutionPlan(platformOrCapability, input, planOptions);
    },
    meetingAppAdapterExecutionPlanMatrix(planOptions = {}) {
      return runtime.kit.meetingAppAdapterExecutionPlanMatrix(sdkPlatformOptions(runtime, planOptions));
    },
    adapterExecutionPlanMatrix(planOptions = {}) {
      return sdk.meetingAppAdapterExecutionPlanMatrix(planOptions);
    },
    platformAdaptationPackage(platformOrOptions = {}, packageOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, packageOptions);
      return runtime.kit.platformAdaptationPackage(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adaptationPackage(platformOrOptions = {}, packageOptions = {}) {
      return sdk.platformAdaptationPackage(platformOrOptions, packageOptions);
    },
    platformAdaptationPackageMatrix(packageOptions = {}) {
      return runtime.kit.platformAdaptationPackageMatrix(sdkPlatformOptions(runtime, packageOptions));
    },
    adaptationPackageMatrix(packageOptions = {}) {
      return sdk.platformAdaptationPackageMatrix(packageOptions);
    },
    platformConsumerHandoff(handoffOptions = {}) {
      return runtime.kit.platformConsumerHandoff(sdkPlatformOptions(runtime, handoffOptions));
    },
    consumerHandoff(handoffOptions = {}) {
      return sdk.platformConsumerHandoff(handoffOptions);
    },
    assertPlatformConsumerHandoff(handoffOptions = {}) {
      return runtime.kit.assertPlatformConsumerHandoff(sdkPlatformOptions(runtime, handoffOptions));
    },
    assertConsumerHandoff(handoffOptions = {}) {
      return sdk.assertPlatformConsumerHandoff(handoffOptions);
    },
    platformImplementationHandoff(platformOrOptions = {}, handoffOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, handoffOptions);
      return runtime.kit.platformImplementationHandoff(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    implementationHandoff(platformOrOptions = {}, handoffOptions = {}) {
      return sdk.platformImplementationHandoff(platformOrOptions, handoffOptions);
    },
    platformImplementationHandoffMatrix(handoffOptions = {}) {
      return runtime.kit.platformImplementationHandoffMatrix(sdkPlatformOptions(runtime, handoffOptions));
    },
    implementationHandoffMatrix(handoffOptions = {}) {
      return sdk.platformImplementationHandoffMatrix(handoffOptions);
    },
    platformAdapterAuthoringPlan(platformOrOptions = {}, authoringOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, authoringOptions);
      return runtime.kit.platformAdapterAuthoringPlan(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterAuthoringPlan(platformOrOptions = {}, authoringOptions = {}) {
      return sdk.platformAdapterAuthoringPlan(platformOrOptions, authoringOptions);
    },
    platformAdapterAuthoringMatrix(authoringOptions = {}) {
      return runtime.kit.platformAdapterAuthoringMatrix(sdkPlatformOptions(runtime, authoringOptions));
    },
    adapterAuthoringMatrix(authoringOptions = {}) {
      return sdk.platformAdapterAuthoringMatrix(authoringOptions);
    },
    platformAdapterPortfolioItem(platformOrOptions = {}, portfolioOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, portfolioOptions);
      return runtime.kit.platformAdapterPortfolioItem(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterPortfolioItem(platformOrOptions = {}, portfolioOptions = {}) {
      return sdk.platformAdapterPortfolioItem(platformOrOptions, portfolioOptions);
    },
    platformAdapterPortfolio(portfolioOptions = {}) {
      return runtime.kit.platformAdapterPortfolio(sdkPlatformOptions(runtime, portfolioOptions));
    },
    adapterPortfolio(portfolioOptions = {}) {
      return sdk.platformAdapterPortfolio(portfolioOptions);
    },
    platformAdapterAcceptanceChecklist(platformOrOptions = {}, checklistInput = {}, checklistOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, checklistOptions);
      return runtime.kit.platformAdapterAcceptanceChecklist(
        resolved.platform,
        checklistInput,
        sdkPlatformOptions(runtime, resolved.options),
      );
    },
    adapterAcceptanceChecklist(platformOrOptions = {}, checklistInput = {}, checklistOptions = {}) {
      return sdk.platformAdapterAcceptanceChecklist(platformOrOptions, checklistInput, checklistOptions);
    },
    platformAdapterAcceptanceChecklistMatrix(checklistInput = {}, checklistOptions = {}) {
      return runtime.kit.platformAdapterAcceptanceChecklistMatrix(
        checklistInput,
        sdkPlatformOptions(runtime, checklistOptions),
      );
    },
    adapterAcceptanceChecklistMatrix(checklistInput = {}, checklistOptions = {}) {
      return sdk.platformAdapterAcceptanceChecklistMatrix(checklistInput, checklistOptions);
    },
    platformAdapterExportPackage(platformOrOptions = {}, exportInput = {}, exportOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, exportOptions);
      return runtime.kit.platformAdapterExportPackage(
        resolved.platform,
        exportInput,
        sdkPlatformOptions(runtime, resolved.options),
      );
    },
    adapterExportPackage(platformOrOptions = {}, exportInput = {}, exportOptions = {}) {
      return sdk.platformAdapterExportPackage(platformOrOptions, exportInput, exportOptions);
    },
    platformAdapterExportPackageMatrix(exportInput = {}, exportOptions = {}) {
      return runtime.kit.platformAdapterExportPackageMatrix(
        exportInput,
        sdkPlatformOptions(runtime, exportOptions),
      );
    },
    adapterExportPackageMatrix(exportInput = {}, exportOptions = {}) {
      return sdk.platformAdapterExportPackageMatrix(exportInput, exportOptions);
    },
    platformAdapterImportPlan(exportPackage = {}, importInput = {}, importOptions = {}) {
      return runtime.kit.platformAdapterImportPlan(
        exportPackage,
        importInput,
        sdkPlatformOptions(runtime, importOptions),
      );
    },
    adapterImportPlan(exportPackage = {}, importInput = {}, importOptions = {}) {
      return sdk.platformAdapterImportPlan(exportPackage, importInput, importOptions);
    },
    assertPlatformAdapterImportPlan(planOrPackage = {}, importInput = {}, importOptions = {}) {
      return runtime.kit.assertPlatformAdapterImportPlan(
        planOrPackage,
        importInput,
        sdkPlatformOptions(runtime, importOptions),
      );
    },
    assertAdapterImportPlan(planOrPackage = {}, importInput = {}, importOptions = {}) {
      return sdk.assertPlatformAdapterImportPlan(planOrPackage, importInput, importOptions);
    },
    platformAdapterImportPlanMatrix(packagesOrInput = {}, importInput = {}, importOptions = {}) {
      return runtime.kit.platformAdapterImportPlanMatrix(
        packagesOrInput,
        importInput,
        sdkPlatformOptions(runtime, importOptions),
      );
    },
    adapterImportPlanMatrix(packagesOrInput = {}, importInput = {}, importOptions = {}) {
      return sdk.platformAdapterImportPlanMatrix(packagesOrInput, importInput, importOptions);
    },
    platformAdapterInstallManifest(plansOrInput = {}, installInput = {}, installOptions = {}) {
      return runtime.kit.platformAdapterInstallManifest(
        plansOrInput,
        installInput,
        sdkPlatformOptions(runtime, installOptions),
      );
    },
    adapterInstallManifest(plansOrInput = {}, installInput = {}, installOptions = {}) {
      return sdk.platformAdapterInstallManifest(plansOrInput, installInput, installOptions);
    },
    assertPlatformAdapterInstallManifest(manifestOrInput = {}, installInput = {}, installOptions = {}) {
      return runtime.kit.assertPlatformAdapterInstallManifest(
        manifestOrInput,
        installInput,
        sdkPlatformOptions(runtime, installOptions),
      );
    },
    assertAdapterInstallManifest(manifestOrInput = {}, installInput = {}, installOptions = {}) {
      return sdk.assertPlatformAdapterInstallManifest(manifestOrInput, installInput, installOptions);
    },
    platformAdapterLaunchPlan(manifestOrInput = {}, launchInput = {}, launchOptions = {}) {
      return runtime.kit.platformAdapterLaunchPlan(
        manifestOrInput,
        launchInput,
        sdkPlatformOptions(runtime, launchOptions),
      );
    },
    adapterLaunchPlan(manifestOrInput = {}, launchInput = {}, launchOptions = {}) {
      return sdk.platformAdapterLaunchPlan(manifestOrInput, launchInput, launchOptions);
    },
    platformAdapterCandidateLaunchPlan(manifestOrInput = {}, launchInput = {}, launchOptions = {}) {
      return runtime.kit.platformAdapterCandidateLaunchPlan(
        manifestOrInput,
        launchInput,
        sdkPlatformOptions(runtime, launchOptions),
      );
    },
    adapterCandidateLaunchPlan(manifestOrInput = {}, launchInput = {}, launchOptions = {}) {
      return sdk.platformAdapterCandidateLaunchPlan(manifestOrInput, launchInput, launchOptions);
    },
    assertPlatformAdapterLaunchPlan(planOrInput = {}, launchInput = {}, launchOptions = {}) {
      return runtime.kit.assertPlatformAdapterLaunchPlan(
        planOrInput,
        launchInput,
        sdkPlatformOptions(runtime, launchOptions),
      );
    },
    assertAdapterLaunchPlan(planOrInput = {}, launchInput = {}, launchOptions = {}) {
      return sdk.assertPlatformAdapterLaunchPlan(planOrInput, launchInput, launchOptions);
    },
    assertPlatformAdapterCandidateLaunchPlan(planOrInput = {}, launchInput = {}, launchOptions = {}) {
      return runtime.kit.assertPlatformAdapterCandidateLaunchPlan(
        planOrInput,
        launchInput,
        sdkPlatformOptions(runtime, launchOptions),
      );
    },
    assertAdapterCandidateLaunchPlan(planOrInput = {}, launchInput = {}, launchOptions = {}) {
      return sdk.assertPlatformAdapterCandidateLaunchPlan(planOrInput, launchInput, launchOptions);
    },
    platformAdapterSession(launchPlanOrInput = {}, clientOrSessionOptions = {}, sessionOptions = {}) {
      return runtime.kit.platformAdapterSession(
        launchPlanOrInput,
        clientOrSessionOptions,
        sdkPlatformOptions(runtime, sessionOptions),
      );
    },
    adapterSession(launchPlanOrInput = {}, clientOrSessionOptions = {}, sessionOptions = {}) {
      return sdk.platformAdapterSession(launchPlanOrInput, clientOrSessionOptions, sessionOptions);
    },
    platformAdapterSessionHandoff(launchPlanOrInput = {}, handoffOptions = {}) {
      return runtime.kit.platformAdapterSessionHandoff(
        launchPlanOrInput,
        sdkPlatformOptions(runtime, handoffOptions),
      );
    },
    adapterSessionHandoff(launchPlanOrInput = {}, handoffOptions = {}) {
      return sdk.platformAdapterSessionHandoff(launchPlanOrInput, handoffOptions);
    },
    platformAdapterRunner(manifestOrInput = {}, clientOrRunnerOptions = {}, runnerOptions = {}) {
      return runtime.kit.platformAdapterRunner(
        manifestOrInput,
        clientOrRunnerOptions,
        sdkPlatformOptions(runtime, runnerOptions),
      );
    },
    adapterRunner(manifestOrInput = {}, clientOrRunnerOptions = {}, runnerOptions = {}) {
      return sdk.platformAdapterRunner(manifestOrInput, clientOrRunnerOptions, runnerOptions);
    },
    openPlatformAdapterSession(manifestOrInput = {}, launchInput = {}, openOptions = {}) {
      return runtime.kit.openPlatformAdapterSession(
        manifestOrInput,
        launchInput,
        sdkPlatformOptions(runtime, openOptions),
      );
    },
    openAdapterSession(manifestOrInput = {}, launchInput = {}, openOptions = {}) {
      return sdk.openPlatformAdapterSession(manifestOrInput, launchInput, openOptions);
    },
    platformAdapterRunnerHandoff(manifestOrInput = {}, handoffOptions = {}) {
      return runtime.kit.platformAdapterRunnerHandoff(
        manifestOrInput,
        sdkPlatformOptions(runtime, handoffOptions),
      );
    },
    adapterRunnerHandoff(manifestOrInput = {}, handoffOptions = {}) {
      return sdk.platformAdapterRunnerHandoff(manifestOrInput, handoffOptions);
    },
    platformAdapterMessageBridge(manifestOrRunner = {}, clientOrBridgeOptions = {}, bridgeOptions = {}) {
      return runtime.kit.platformAdapterMessageBridge(
        manifestOrRunner,
        clientOrBridgeOptions,
        sdkPlatformOptions(runtime, bridgeOptions),
      );
    },
    adapterMessageBridge(manifestOrRunner = {}, clientOrBridgeOptions = {}, bridgeOptions = {}) {
      return sdk.platformAdapterMessageBridge(manifestOrRunner, clientOrBridgeOptions, bridgeOptions);
    },
    platformAdapterMessageBridgeHandoff(manifestOrInput = {}, handoffOptions = {}) {
      return runtime.kit.platformAdapterMessageBridgeHandoff(
        manifestOrInput,
        sdkPlatformOptions(runtime, handoffOptions),
      );
    },
    adapterMessageBridgeHandoff(manifestOrInput = {}, handoffOptions = {}) {
      return sdk.platformAdapterMessageBridgeHandoff(manifestOrInput, handoffOptions);
    },
    platformAdapterSmoke(manifestOrOptions = {}, smokeOptions = {}) {
      return runtime.kit.platformAdapterSmoke(
        manifestOrOptions,
        sdkPlatformOptions(runtime, smokeOptions),
      );
    },
    adapterSmoke(manifestOrOptions = {}, smokeOptions = {}) {
      return sdk.platformAdapterSmoke(manifestOrOptions, smokeOptions);
    },
    assertPlatformAdapterSmoke(manifestOrOptions = {}, smokeOptions = {}) {
      return runtime.kit.assertPlatformAdapterSmoke(
        manifestOrOptions,
        sdkPlatformOptions(runtime, smokeOptions),
      );
    },
    assertAdapterSmoke(manifestOrOptions = {}, smokeOptions = {}) {
      return sdk.assertPlatformAdapterSmoke(manifestOrOptions, smokeOptions);
    },
    platformRuntimeBundle(platformOrOptions = {}, bundleOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, bundleOptions);
      return runtime.kit.platformRuntimeBundle(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    runtimeBundle(platformOrOptions = {}, bundleOptions = {}) {
      return sdk.platformRuntimeBundle(platformOrOptions, bundleOptions);
    },
    platformRuntimeBundleMatrix(bundleOptions = {}) {
      return runtime.kit.platformRuntimeBundleMatrix(sdkPlatformOptions(runtime, bundleOptions));
    },
    runtimeBundleMatrix(bundleOptions = {}) {
      return sdk.platformRuntimeBundleMatrix(bundleOptions);
    },
    platformAdapterRoute(platformOrOptions = {}, routeOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, routeOptions);
      return runtime.kit.platformAdapterRoute(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterRoute(platformOrOptions = {}, routeOptions = {}) {
      return sdk.platformAdapterRoute(platformOrOptions, routeOptions);
    },
    platformAdapterRouteMatrix(routeOptions = {}) {
      return runtime.kit.platformAdapterRouteMatrix(sdkPlatformOptions(runtime, routeOptions));
    },
    adapterRouteMatrix(routeOptions = {}) {
      return sdk.platformAdapterRouteMatrix(routeOptions);
    },
    platformAdapterSelection(platformOrOptions = {}, input = {}, selectionOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, selectionOptions);
      return runtime.kit.platformAdapterSelection(resolved.platform, input, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterSelection(platformOrOptions = {}, input = {}, selectionOptions = {}) {
      return sdk.platformAdapterSelection(platformOrOptions, input, selectionOptions);
    },
    platformAdapterSelectionMatrix(input = {}, selectionOptions = {}) {
      return runtime.kit.platformAdapterSelectionMatrix(input, sdkPlatformOptions(runtime, selectionOptions));
    },
    adapterSelectionMatrix(input = {}, selectionOptions = {}) {
      return sdk.platformAdapterSelectionMatrix(input, selectionOptions);
    },
    platformAdapterBlueprint(platformOrOptions = {}, blueprintOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, blueprintOptions);
      return runtime.kit.platformAdapterBlueprint(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterBlueprint(platformOrOptions = {}, blueprintOptions = {}) {
      return sdk.platformAdapterBlueprint(platformOrOptions, blueprintOptions);
    },
    platformAdapterBlueprintMatrix(blueprintOptions = {}) {
      return runtime.kit.platformAdapterBlueprintMatrix(sdkPlatformOptions(runtime, blueprintOptions));
    },
    adapterBlueprintMatrix(blueprintOptions = {}) {
      return sdk.platformAdapterBlueprintMatrix(blueprintOptions);
    },
    verifyPlatformAdapterBlueprint(blueprintOrPlatform = {}, blueprintOptions = {}) {
      return runtime.kit.verifyPlatformAdapterBlueprint(
        blueprintOrPlatform,
        sdkPlatformOptions(runtime, blueprintOptions),
      );
    },
    assertPlatformAdapterBlueprint(blueprintOrPlatform = {}, blueprintOptions = {}) {
      return runtime.kit.assertPlatformAdapterBlueprint(
        blueprintOrPlatform,
        sdkPlatformOptions(runtime, blueprintOptions),
      );
    },
    assertAdapterBlueprint(blueprintOrPlatform = {}, blueprintOptions = {}) {
      return sdk.assertPlatformAdapterBlueprint(blueprintOrPlatform, blueprintOptions);
    },
    assertPlatformAdapterBlueprintMatrix(matrixOrOptions = {}) {
      return runtime.kit.assertPlatformAdapterBlueprintMatrix(sdkPlatformOptions(runtime, matrixOrOptions));
    },
    assertAdapterBlueprintMatrix(matrixOrOptions = {}) {
      return sdk.assertPlatformAdapterBlueprintMatrix(matrixOrOptions);
    },
    platformAdapterDecision(input = {}, decisionOptions = {}) {
      return runtime.kit.platformAdapterDecision(input, sdkPlatformOptions(runtime, decisionOptions));
    },
    adapterDecision(input = {}, decisionOptions = {}) {
      return sdk.platformAdapterDecision(input, decisionOptions);
    },
    platformAdapterDecisionMatrix(input = {}, decisionOptions = {}) {
      return runtime.kit.platformAdapterDecisionMatrix(input, sdkPlatformOptions(runtime, decisionOptions));
    },
    adapterDecisionMatrix(input = {}, decisionOptions = {}) {
      return sdk.platformAdapterDecisionMatrix(input, decisionOptions);
    },
    platformHostProfileCompatibilityMatrix(input = {}, decisionOptions = {}) {
      return runtime.kit.platformHostProfileCompatibilityMatrix(input, sdkPlatformOptions(runtime, decisionOptions));
    },
    hostProfileCompatibilityMatrix(input = {}, decisionOptions = {}) {
      return sdk.platformHostProfileCompatibilityMatrix(input, decisionOptions);
    },
    platformAdapterMatrix(input = {}, matrixOptions = {}) {
      return runtime.kit.platformAdapterMatrix(input, sdkPlatformOptions(runtime, matrixOptions));
    },
    adapterMatrix(input = {}, matrixOptions = {}) {
      return sdk.platformAdapterMatrix(input, matrixOptions);
    },
    platformAdapterMatrixRow(platformOrOptions = {}, input = {}, matrixOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, matrixOptions);
      return runtime.kit.platformAdapterMatrixRow(resolved.platform, input, sdkPlatformOptions(runtime, resolved.options));
    },
    adapterMatrixRow(platformOrOptions = {}, input = {}, matrixOptions = {}) {
      return sdk.platformAdapterMatrixRow(platformOrOptions, input, matrixOptions);
    },
    assertPlatformAdapterDecision(input = {}, decisionOptions = {}) {
      return runtime.kit.assertPlatformAdapterDecision(input, sdkPlatformOptions(runtime, decisionOptions));
    },
    assertAdapterDecision(input = {}, decisionOptions = {}) {
      return sdk.assertPlatformAdapterDecision(input, decisionOptions);
    },
    assertPlatformAdapterDecisionMatrix(input = {}, decisionOptions = {}) {
      return runtime.kit.assertPlatformAdapterDecisionMatrix(input, sdkPlatformOptions(runtime, decisionOptions));
    },
    assertAdapterDecisionMatrix(input = {}, decisionOptions = {}) {
      return sdk.assertPlatformAdapterDecisionMatrix(input, decisionOptions);
    },
    platformAdapterStartupPlan(input = {}, startupOptions = {}) {
      return runtime.kit.platformAdapterStartupPlan(input, sdkPlatformOptions(runtime, startupOptions));
    },
    adapterStartupPlan(input = {}, startupOptions = {}) {
      return sdk.platformAdapterStartupPlan(input, startupOptions);
    },
    platformAdapterStartupPlanMatrix(input = {}, startupOptions = {}) {
      return runtime.kit.platformAdapterStartupPlanMatrix(input, sdkPlatformOptions(runtime, startupOptions));
    },
    adapterStartupPlanMatrix(input = {}, startupOptions = {}) {
      return sdk.platformAdapterStartupPlanMatrix(input, startupOptions);
    },
    assertPlatformAdapterStartupPlan(input = {}, startupOptions = {}) {
      return runtime.kit.assertPlatformAdapterStartupPlan(input, sdkPlatformOptions(runtime, startupOptions));
    },
    assertAdapterStartupPlan(input = {}, startupOptions = {}) {
      return sdk.assertPlatformAdapterStartupPlan(input, startupOptions);
    },
    assertPlatformAdapterStartupPlanMatrix(input = {}, startupOptions = {}) {
      return runtime.kit.assertPlatformAdapterStartupPlanMatrix(input, sdkPlatformOptions(runtime, startupOptions));
    },
    assertAdapterStartupPlanMatrix(input = {}, startupOptions = {}) {
      return sdk.assertPlatformAdapterStartupPlanMatrix(input, startupOptions);
    },
    platformAdapterRuntimeRecipe(input = {}, recipeOptions = {}) {
      return runtime.kit.platformAdapterRuntimeRecipe(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    adapterRuntimeRecipe(input = {}, recipeOptions = {}) {
      return sdk.platformAdapterRuntimeRecipe(input, recipeOptions);
    },
    platformAdapterRuntimeRecipeMatrix(input = {}, recipeOptions = {}) {
      return runtime.kit.platformAdapterRuntimeRecipeMatrix(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    adapterRuntimeRecipeMatrix(input = {}, recipeOptions = {}) {
      return sdk.platformAdapterRuntimeRecipeMatrix(input, recipeOptions);
    },
    platformAdapterRuntimeManifest(input = {}, recipeOptions = {}) {
      return runtime.kit.platformAdapterRuntimeManifest(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    adapterRuntimeManifest(input = {}, recipeOptions = {}) {
      return sdk.platformAdapterRuntimeManifest(input, recipeOptions);
    },
    platformAdapterRuntimeTarget(manifestOrInput = {}, input = {}, targetOptions = {}) {
      return runtime.kit.platformAdapterRuntimeTarget(manifestOrInput, input, sdkPlatformOptions(runtime, targetOptions));
    },
    adapterRuntimeTarget(manifestOrInput = {}, input = {}, targetOptions = {}) {
      return sdk.platformAdapterRuntimeTarget(manifestOrInput, input, targetOptions);
    },
    assertPlatformAdapterRuntimeRecipe(input = {}, recipeOptions = {}) {
      return runtime.kit.assertPlatformAdapterRuntimeRecipe(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    assertAdapterRuntimeRecipe(input = {}, recipeOptions = {}) {
      return sdk.assertPlatformAdapterRuntimeRecipe(input, recipeOptions);
    },
    assertPlatformAdapterRuntimeRecipeMatrix(input = {}, recipeOptions = {}) {
      return runtime.kit.assertPlatformAdapterRuntimeRecipeMatrix(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    assertAdapterRuntimeRecipeMatrix(input = {}, recipeOptions = {}) {
      return sdk.assertPlatformAdapterRuntimeRecipeMatrix(input, recipeOptions);
    },
    assertPlatformAdapterRuntimeManifest(input = {}, recipeOptions = {}) {
      return runtime.kit.assertPlatformAdapterRuntimeManifest(input, sdkPlatformOptions(runtime, recipeOptions));
    },
    assertAdapterRuntimeManifest(input = {}, recipeOptions = {}) {
      return sdk.assertPlatformAdapterRuntimeManifest(input, recipeOptions);
    },
    assertPlatformAdapterRuntimeTarget(manifestOrInput = {}, input = {}, targetOptions = {}) {
      return runtime.kit.assertPlatformAdapterRuntimeTarget(manifestOrInput, input, sdkPlatformOptions(runtime, targetOptions));
    },
    assertAdapterRuntimeTarget(manifestOrInput = {}, input = {}, targetOptions = {}) {
      return sdk.assertPlatformAdapterRuntimeTarget(manifestOrInput, input, targetOptions);
    },
    platformAdapterPreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.platformAdapterPreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    adapterPreflight(input = {}, preflightOptions = {}) {
      return sdk.platformAdapterPreflight(input, preflightOptions);
    },
    platformAdapterCurrentWindowPreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.platformAdapterCurrentWindowPreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    adapterCurrentWindowPreflight(input = {}, preflightOptions = {}) {
      return sdk.platformAdapterCurrentWindowPreflight(input, preflightOptions);
    },
    platformAdapterPreflightMatrix(input = {}, preflightOptions = {}) {
      return runtime.kit.platformAdapterPreflightMatrix(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    adapterPreflightMatrix(input = {}, preflightOptions = {}) {
      return sdk.platformAdapterPreflightMatrix(input, preflightOptions);
    },
    platformAdapterCandidatePreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.platformAdapterCandidatePreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    adapterCandidatePreflight(input = {}, preflightOptions = {}) {
      return sdk.platformAdapterCandidatePreflight(input, preflightOptions);
    },
    assertPlatformAdapterPreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.assertPlatformAdapterPreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    assertAdapterPreflight(input = {}, preflightOptions = {}) {
      return sdk.assertPlatformAdapterPreflight(input, preflightOptions);
    },
    assertPlatformAdapterCurrentWindowPreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.assertPlatformAdapterCurrentWindowPreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    assertAdapterCurrentWindowPreflight(input = {}, preflightOptions = {}) {
      return sdk.assertPlatformAdapterCurrentWindowPreflight(input, preflightOptions);
    },
    assertPlatformAdapterPreflightMatrix(input = {}, preflightOptions = {}) {
      return runtime.kit.assertPlatformAdapterPreflightMatrix(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    assertAdapterPreflightMatrix(input = {}, preflightOptions = {}) {
      return sdk.assertPlatformAdapterPreflightMatrix(input, preflightOptions);
    },
    assertPlatformAdapterCandidatePreflight(input = {}, preflightOptions = {}) {
      return runtime.kit.assertPlatformAdapterCandidatePreflight(input, sdkPlatformOptions(runtime, preflightOptions));
    },
    assertAdapterCandidatePreflight(input = {}, preflightOptions = {}) {
      return sdk.assertPlatformAdapterCandidatePreflight(input, preflightOptions);
    },
    platformRealtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
      const objectInput = platformOrInput && typeof platformOrInput === 'object' && !(platformOrInput instanceof URL);
      return runtime.kit.platformRealtimeAxisReadiness(
        platformOrInput,
        objectInput ? sdkPlatformOptions(runtime, inputOrOptions) : inputOrOptions,
        objectInput ? {} : sdkPlatformOptions(runtime, maybeOptions),
      );
    },
    realtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
      return sdk.platformRealtimeAxisReadiness(platformOrInput, inputOrOptions, maybeOptions);
    },
    platformRealtimeAxisReadinessMatrix(input = {}, readinessOptions = {}) {
      return runtime.kit.platformRealtimeAxisReadinessMatrix(input, sdkPlatformOptions(runtime, readinessOptions));
    },
    realtimeAxisReadinessMatrix(input = {}, readinessOptions = {}) {
      return sdk.platformRealtimeAxisReadinessMatrix(input, readinessOptions);
    },
    assertPlatformRealtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
      const objectInput = platformOrInput && typeof platformOrInput === 'object' && !(platformOrInput instanceof URL);
      return runtime.kit.assertPlatformRealtimeAxisReadiness(
        platformOrInput,
        objectInput ? sdkPlatformOptions(runtime, inputOrOptions) : inputOrOptions,
        objectInput ? {} : sdkPlatformOptions(runtime, maybeOptions),
      );
    },
    assertRealtimeAxisReadiness(platformOrInput = {}, inputOrOptions = {}, maybeOptions = {}) {
      return sdk.assertPlatformRealtimeAxisReadiness(platformOrInput, inputOrOptions, maybeOptions);
    },
    assertPlatformRealtimeAxisReadinessMatrix(input = {}, readinessOptions = {}) {
      return runtime.kit.assertPlatformRealtimeAxisReadinessMatrix(input, sdkPlatformOptions(runtime, readinessOptions));
    },
    assertRealtimeAxisReadinessMatrix(input = {}, readinessOptions = {}) {
      return sdk.assertPlatformRealtimeAxisReadinessMatrix(input, readinessOptions);
    },
    platformAdaptationStrategy(platformOrOptions = {}, strategyOptions = {}) {
      const resolved = singlePlatformInput(runtime, platformOrOptions, strategyOptions);
      return runtime.kit.platformAdaptationStrategy(resolved.platform, sdkPlatformOptions(runtime, resolved.options));
    },
    adaptationStrategy(platformOrOptions = {}, strategyOptions = {}) {
      return sdk.platformAdaptationStrategy(platformOrOptions, strategyOptions);
    },
    platformAdaptationStrategyMatrix(strategyOptions = {}) {
      return runtime.kit.platformAdaptationStrategyMatrix(sdkPlatformOptions(runtime, strategyOptions));
    },
    adaptationStrategyMatrix(strategyOptions = {}) {
      return sdk.platformAdaptationStrategyMatrix(strategyOptions);
    },
    platformConnectorHub(connectorOptions = {}) {
      return runtime.kit.platformConnectorHub(sdkPlatformOptions(runtime, connectorOptions));
    },
    connectorHub(connectorOptions = {}) {
      return sdk.platformConnectorHub(connectorOptions);
    },
    manifest(manifestOptions = {}) {
      return runtime.manifest(manifestOptions);
    },
    readiness(readinessOptions = {}) {
      return runtime.readiness(readinessOptions);
    },
    handoffReadiness(readinessOptions = {}) {
      return runtime.handoffReadiness(readinessOptions);
    },
    platformRawSignal(input = {}, signalOptions = {}) {
      return runtime.kit.platformRawSignal(input, sdkPlatformOptions(runtime, signalOptions));
    },
    rawSignal(input = {}, signalOptions = {}) {
      return sdk.platformRawSignal(input, signalOptions);
    },
    platformRuntimeEventsFromRawSignal(input = {}, signalOptions = {}) {
      return runtime.kit.platformRuntimeEventsFromRawSignal(input, sdkPlatformOptions(runtime, signalOptions));
    },
    runtimeEventsFromRawSignal(input = {}, signalOptions = {}) {
      return sdk.platformRuntimeEventsFromRawSignal(input, signalOptions);
    },
    platformRawSignalBatch(input = {}, signalOptions = {}) {
      return runtime.kit.platformRawSignalBatch(input, sdkPlatformOptions(runtime, signalOptions));
    },
    rawSignalBatch(input = {}, signalOptions = {}) {
      return sdk.platformRawSignalBatch(input, signalOptions);
    },
    platformRawSignalExampleBatch(signalOptions = {}) {
      return runtime.kit.platformRawSignalExampleBatch(sdkPlatformOptions(runtime, signalOptions));
    },
    rawSignalExampleBatch(signalOptions = {}) {
      return sdk.platformRawSignalExampleBatch(signalOptions);
    },
    observePlatformCandidates(input = {}, observeOptions = {}) {
      return shouldUseRuntimeEvent(observeOptions, defaults)
        ? runtimeEvents.observePlatformCandidates(input, observeOptions)
        : runtime.observePlatformCandidates(input, observeOptions);
    },
    observeMeetingApp(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      const resolved = platformAndInput(runtime, platformOrInput, inputOrOptions, optionsForPlatform);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.observeMeetingApp(resolved.platform, resolved.input, resolved.options)
        : runtime.observeMeetingApp(resolved.platform, resolved.input, resolved.options);
    },
    observeApp(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      return sdk.observeMeetingApp(platformOrInput, inputOrOptions, optionsForPlatform);
    },
    insertAnnotation(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      const resolved = platformAndInput(runtime, platformOrInput, inputOrOptions, optionsForPlatform);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.insertAnnotation(resolved.platform, resolved.input, resolved.options)
        : runtime.insertAnnotation(resolved.platform, resolved.input, resolved.options);
    },
    insertMark(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      return sdk.insertAnnotation(platformOrInput, inputOrOptions, optionsForPlatform);
    },
    ingestProvider(platformOrInput, inputOrPayload = {}, optionsOrPayload = {}, maybeOptions = {}) {
      if (typeof platformOrInput === 'string') {
        return shouldUseRuntimeEvent(maybeOptions, defaults)
          ? runtimeEvents.ingestProvider(platformOrInput, inputOrPayload, maybeOptions)
          : runtime.ingestProvider(platformOrInput, inputOrPayload, optionsOrPayload, maybeOptions);
      }
      const resolved = platformAndInput(runtime, platformOrInput, inputOrPayload);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.ingestProvider(resolved.platform, resolved.input, resolved.options)
        : runtime.ingestProvider(resolved.platform, resolved.input, resolved.input, resolved.options);
    },
    speakerTrack(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      const resolved = platformAndInput(runtime, platformOrInput, inputOrOptions, optionsForPlatform);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.speakerTrack(resolved.platform, resolved.input, resolved.options)
        : runtime.speakerTrack(resolved.platform, resolved.input, resolved.options);
    },
    participantTrack(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      const resolved = platformAndInput(runtime, platformOrInput, inputOrOptions, optionsForPlatform);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.participantTrack(resolved.platform, resolved.input, resolved.options)
        : runtime.participantTrack(resolved.platform, resolved.input, resolved.options);
    },
    timelineView(platformOrInput, inputOrOptions = {}, optionsForPlatform = {}) {
      const resolved = platformAndInput(runtime, platformOrInput, inputOrOptions, optionsForPlatform);
      return shouldUseRuntimeEvent(resolved.options, defaults)
        ? runtimeEvents.timelineView(resolved.platform, resolved.input, resolved.options)
        : runtime.timelineView(resolved.platform, resolved.input, resolved.options);
    },
    sendRuntimeEvent(eventInput = {}, sendOptions = {}) {
      return runtimeEvents.send(eventInput, sendOptions);
    },
    async sendRawSignal(input = {}, sendOptions = {}) {
      const rawSignal = sdk.platformRawSignal(input, sendOptions);
      const results = [];
      for (const event of rawSignal.runtime_events ?? []) {
        results.push(await runtimeEvents.send(event, sendOptions));
      }
      return {
        type: 'meeting_platform_raw_signal_send_result',
        schema: 'meeting_platform_raw_signal_send_result',
        signal: rawSignal,
        runtime_event_count: rawSignal.runtime_event_count,
        runtime_events: rawSignal.runtime_events,
        results,
      };
    },
    async sendRawSignalBatch(input = {}, sendOptions = {}) {
      const batch = sdk.platformRawSignalBatch(input, sendOptions);
      const results = [];
      for (const event of batch.runtime_events ?? []) {
        results.push(await runtimeEvents.send(event, sendOptions));
      }
      return {
        type: 'meeting_platform_raw_signal_batch_send_result',
        schema: 'meeting_platform_raw_signal_batch_send_result',
        batch,
        signal_count: batch.signal_count,
        runtime_event_count: batch.runtime_event_count,
        runtime_events: batch.runtime_events,
        results,
      };
    },
    handleRuntimeEvent(eventInput = {}, payload, eventOptions = {}) {
      return runtime.handleEvent(eventInput, payload, eventOptions);
    },
  };
  return sdk;
}

export * from './adapters/platform-kit.mjs';
export * from './adapters/meeting-app-adapter-capability.mjs';
export * from './adapters/meeting-app-adapter-integration-package.mjs';
export * from './adapters/meeting-app-connector-package.mjs';
export * from './adapters/meeting-platform-connector.mjs';
export * from './adapters/platform-integration-runtime.mjs';
export * from './adapters/platform-runtime-event.mjs';
export * from './adapters/platform-raw-signal.mjs';
export * from './adapters/platform-ingest.mjs';
export * from './adapters/platform-adaptation-package.mjs';
export * from './adapters/platform-consumer-handoff.mjs';
export * from './adapters/platform-implementation-handoff.mjs';
export * from './adapters/platform-adapter-authoring.mjs';
export * from './adapters/platform-adapter-portfolio.mjs';
export * from './adapters/platform-adapter-acceptance-checklist.mjs';
export * from './adapters/platform-adapter-export-package.mjs';
export * from './adapters/platform-adapter-import-plan.mjs';
export * from './adapters/platform-adapter-install-manifest.mjs';
export * from './adapters/platform-adapter-launch-plan.mjs';
export * from './adapters/platform-adapter-session.mjs';
export * from './adapters/platform-adapter-runner.mjs';
export * from './adapters/platform-adapter-message-bridge.mjs';
export * from './adapters/platform-adapter-smoke.mjs';
export * from './adapters/platform-runtime-profile.mjs';
export * from './adapters/platform-runtime-bundle.mjs';
export * from './adapters/platform-adapter-route.mjs';
export * from './adapters/platform-adapter-selection.mjs';
export * from './adapters/platform-adapter-blueprint.mjs';
export * from './adapters/platform-adapter-decision.mjs';
export * from './adapters/platform-adapter-matrix.mjs';
export * from './adapters/platform-adapter-startup.mjs';
export * from './adapters/platform-adapter-runtime-recipe.mjs';
export * from './adapters/platform-adapter-preflight.mjs';
export * from './adapters/platform-realtime-axis-readiness.mjs';
export * from './adapters/platform-strategy.mjs';
