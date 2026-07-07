import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { applyMeetingSignals } from './core.mjs';
import { meetingPlatformEventAdapterFor } from './platform-registry.mjs';
import { buildMeetingPlatformProviderRuntimeEvent } from './platform-runtime-event.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';
import { createMeetingSignalReconciler } from './signal-reconciler.mjs';

export const MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA = 'meeting_platform_provider_replay_report';
export const MEETING_PLATFORM_PROVIDER_REPLAY_MATRIX_SCHEMA = 'meeting_platform_provider_replay_matrix';
export const MEETING_PLATFORM_PROVIDER_REPLAY_ACCEPTANCE_SCHEMA = 'meeting_platform_provider_replay_acceptance_report';
export const MEETING_PLATFORM_PROVIDER_REPLAY_SCHEMA_VERSION = 1;

const DEFAULT_PROVIDER_REPLAY_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
  'lark',
]);

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

function iso(ms) {
  return new Date(ms).toISOString();
}

function replayBaseMs(options = {}) {
  const value = firstNonEmpty(options.baseReceivedAtMs, options.base_received_at_ms, options.receivedAtMs, options.received_at_ms);
  return Number.isFinite(Number(value)) ? Number(value) : 1_782_614_400_000;
}

function replayRecord(id, kind, payload, receivedAtMs, required = true) {
  return compactObject({
    id,
    kind,
    required,
    received_at_ms: receivedAtMs,
    payload,
  });
}

function sampleGoogleMeetProviderEvents(options = {}) {
  const baseMs = replayBaseMs(options);
  const recordId = firstNonEmpty(options.meeting_id, options.meetingId, 'google-provider-record-001');
  const recordName = `conferenceRecords/${recordId}`;
  return [
    replayRecord('google-start', 'meeting_start', {
      id: 'google-start',
      type: 'google.workspace.meet.conference.v2.started',
      time: iso(baseMs),
      data: {
        conferenceRecord: { name: recordName },
        meetingUri: 'https://meet.google.com/abc-defg-hij',
        title: 'Google Meet provider replay',
      },
    }, baseMs),
    replayRecord('google-participant-joined', 'participant_joined', {
      id: 'google-participant-joined',
      type: 'google.workspace.meet.participant.v2.joined',
      time: iso(baseMs + 10_000),
      data: {
        participantSession: {
          name: `${recordName}/participants/google-user-1/participantSessions/session-1`,
          participant: { displayName: 'Alex' },
        },
      },
    }, baseMs + 10_000, false),
    replayRecord('google-transcript', 'artifact_ready', {
      id: 'google-transcript',
      type: 'google.workspace.meet.transcript.v2.fileGenerated',
      time: iso(baseMs + 40_000),
      data: {
        transcript: {
          name: `${recordName}/transcripts/transcript-1`,
          docsDestination: { document: 'https://docs.google.com/document/d/google-transcript-1' },
        },
      },
    }, baseMs + 40_000, false),
    replayRecord('google-end', 'meeting_end', {
      id: 'google-end',
      type: 'google.workspace.meet.conference.v2.ended',
      time: iso(baseMs + 60_000),
      data: {
        conferenceRecord: { name: recordName },
        meetingUri: 'https://meet.google.com/abc-defg-hij',
        title: 'Google Meet provider replay',
      },
    }, baseMs + 60_000),
  ];
}

function sampleTeamsProviderEvents(options = {}) {
  const baseMs = replayBaseMs(options);
  const meetingId = firstNonEmpty(options.meeting_id, options.meetingId, 'teams-provider-meeting-001');
  const joinWebUrl = 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_provider_replay';
  return [
    replayRecord('teams-start', 'meeting_start', {
      id: 'teams-start',
      resource: `communications/onlineMeetings(joinWebUrl='${encodeURIComponent(joinWebUrl)}')`,
      resourceData: {
        id: meetingId,
        onlineMeetingId: meetingId,
        eventType: 'callStarted',
        eventDateTime: iso(baseMs),
        joinWebUrl,
        subject: 'Teams provider replay',
      },
    }, baseMs),
    replayRecord('teams-roster', 'participant_joined', {
      id: 'teams-roster',
      resourceData: {
        id: meetingId,
        onlineMeetingId: meetingId,
        eventType: 'rosterUpdated',
        eventDateTime: iso(baseMs + 12_000),
        joinWebUrl,
        participants: [{
          id: 'teams-user-1',
          displayName: 'Alex',
          joinDateTime: iso(baseMs + 12_000),
        }],
      },
    }, baseMs + 12_000, false),
    replayRecord('teams-transcript', 'artifact_ready', {
      id: 'teams-transcript',
      resourceData: {
        id: 'teams-transcript-1',
        onlineMeetingId: meetingId,
        eventType: 'transcriptCreated',
        eventDateTime: iso(baseMs + 45_000),
        joinWebUrl,
      },
    }, baseMs + 45_000, false),
    replayRecord('teams-end', 'meeting_end', {
      id: 'teams-end',
      resourceData: {
        id: meetingId,
        onlineMeetingId: meetingId,
        eventType: 'callEnded',
        eventDateTime: iso(baseMs + 65_000),
        joinWebUrl,
        subject: 'Teams provider replay',
      },
    }, baseMs + 65_000),
  ];
}

function sampleZoomProviderEvents(options = {}) {
  const baseMs = replayBaseMs(options);
  const uuid = firstNonEmpty(options.meeting_id, options.meetingId, 'zoom-provider-uuid-001');
  return [
    replayRecord('zoom-start', 'meeting_start', {
      event: 'meeting.started',
      event_ts: baseMs,
      payload: {
        object: {
          uuid,
          id: 987654321,
          topic: 'Zoom provider replay',
          join_url: 'https://zoom.us/j/987654321',
          start_time: iso(baseMs),
        },
      },
    }, baseMs),
    replayRecord('zoom-participant-joined', 'participant_joined', {
      event: 'meeting.participant_joined',
      event_ts: baseMs + 10_000,
      payload: {
        object: {
          uuid,
          id: 987654321,
          participant: {
            user_id: 'zoom-user-1',
            user_name: 'Alex',
            join_time: iso(baseMs + 10_000),
          },
        },
      },
    }, baseMs + 10_000, false),
    replayRecord('zoom-recording', 'artifact_ready', {
      event: 'recording.completed',
      event_ts: baseMs + 45_000,
      payload: {
        object: {
          uuid,
          id: 987654321,
          share_url: 'https://zoom.us/rec/share/provider-replay',
        },
      },
    }, baseMs + 45_000, false),
    replayRecord('zoom-end', 'meeting_end', {
      event: 'meeting.ended',
      event_ts: baseMs + 60_000,
      payload: {
        object: {
          uuid,
          id: 987654321,
          topic: 'Zoom provider replay',
          join_url: 'https://zoom.us/j/987654321',
          end_time: iso(baseMs + 60_000),
        },
      },
    }, baseMs + 60_000),
  ];
}

function sampleWebexProviderEvents(options = {}) {
  const baseMs = replayBaseMs(options);
  const meetingId = firstNonEmpty(options.meeting_id, options.meetingId, 'webex-provider-meeting-001');
  return [
    replayRecord('webex-start', 'meeting_start', {
      id: 'webex-start',
      resource: 'meetings',
      event: 'started',
      data: {
        id: meetingId,
        title: 'Webex provider replay',
        webLink: 'https://example.webex.com/meet/provider-replay',
        startTime: iso(baseMs),
      },
    }, baseMs),
    replayRecord('webex-participant-joined', 'participant_joined', {
      id: 'webex-participant-joined',
      resource: 'meetingParticipants',
      event: 'joined',
      data: {
        meetingId,
        personId: 'webex-user-1',
        displayName: 'Alex',
        joinTime: iso(baseMs + 10_000),
      },
    }, baseMs + 10_000, false),
    replayRecord('webex-transcript', 'artifact_ready', {
      id: 'webex-transcript',
      resource: 'meetingTranscripts',
      event: 'created',
      data: {
        meetingId,
        id: 'webex-transcript-1',
        txtDownloadLink: 'https://webex.example/provider-replay.vtt',
        created: iso(baseMs + 45_000),
      },
    }, baseMs + 45_000, false),
    replayRecord('webex-end', 'meeting_end', {
      id: 'webex-end',
      resource: 'meetings',
      event: 'ended',
      data: {
        id: meetingId,
        title: 'Webex provider replay',
        webLink: 'https://example.webex.com/meet/provider-replay',
        endTime: iso(baseMs + 60_000),
      },
    }, baseMs + 60_000),
  ];
}

function sampleLarkProviderEvents(options = {}) {
  const baseMs = replayBaseMs(options);
  const meetingId = firstNonEmpty(options.meeting_id, options.meetingId, 'lark-provider-meeting-001');
  const meeting = {
    id: meetingId,
    meeting_no: '123456789',
    topic: 'Lark provider replay',
    url: 'https://vc.feishu.cn/j/lark-provider-meeting-001',
    start_time: String(Math.round(baseMs / 1000)),
  };
  return [
    replayRecord('lark-start', 'meeting_start', {
      header: {
        event_id: 'lark-start',
        event_type: 'vc.meeting.all_meeting_started_v1',
        create_time: String(baseMs),
      },
      event: { meeting, minute_token: 'lark-minute-token-001' },
    }, baseMs),
    replayRecord('lark-join', 'participant_joined', {
      header: {
        event_id: 'lark-join',
        event_type: 'vc.meeting.join_meeting_v1',
        create_time: String(baseMs + 10_000),
      },
      event: {
        meeting,
        user: { open_id: 'lark-user-1', name: 'Alex' },
      },
    }, baseMs + 10_000, false),
    replayRecord('lark-minute', 'artifact_ready', {
      header: {
        event_id: 'lark-minute',
        event_type: 'vc.meeting.minute_created_v1',
        create_time: String(baseMs + 45_000),
      },
      event: {
        meeting,
        minute: {
          token: 'lark-minute-token-001',
          url: 'https://minutes.feishu.cn/minutes/provider-replay',
        },
      },
    }, baseMs + 45_000, false),
    replayRecord('lark-end', 'meeting_end', {
      header: {
        event_id: 'lark-end',
        event_type: 'vc.meeting.all_meeting_ended_v1',
        create_time: String(baseMs + 60_000),
      },
      event: {
        meeting: {
          ...meeting,
          end_time: String(Math.round((baseMs + 60_000) / 1000)),
        },
      },
    }, baseMs + 60_000),
  ];
}

export function sampleMeetingPlatformProviderEvents(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  if (key === 'google_meet') return sampleGoogleMeetProviderEvents(options);
  if (key === 'microsoft_teams') return sampleTeamsProviderEvents(options);
  if (key === 'zoom') return sampleZoomProviderEvents(options);
  if (key === 'webex') return sampleWebexProviderEvents(options);
  if (key === 'lark') return sampleLarkProviderEvents(options);
  return [];
}

function providerReplayInput(platformOrInput = {}, recordsOrOptions = undefined, options = {}) {
  if (platformOrInput && typeof platformOrInput === 'object' && !Array.isArray(platformOrInput)) {
    const input = platformOrInput;
    const records = firstNonEmpty(
      input.records,
      input.events,
      input.samples,
      input.provider_events,
      input.providerEvents,
      Array.isArray(recordsOrOptions) ? recordsOrOptions : undefined,
    );
    return {
      platform: firstNonEmpty(input.platform, input.provider, input.adapter, options.platform),
      records,
      options: {
        ...input,
        ...(Array.isArray(recordsOrOptions) ? {} : recordsOrOptions ?? {}),
        ...options,
      },
    };
  }
  return {
    platform: platformOrInput,
    records: Array.isArray(recordsOrOptions) ? recordsOrOptions : undefined,
    options: {
      ...(Array.isArray(recordsOrOptions) ? {} : recordsOrOptions ?? {}),
      ...options,
    },
  };
}

function replayPayload(record = {}) {
  if (record && typeof record === 'object' && !Array.isArray(record)) {
    const explicitPayload = firstNonEmpty(
      record.payload,
      record.body,
      record.raw,
      record.provider_event,
      record.providerEvent,
    );
    if (explicitPayload != null) return explicitPayload;
    const wrapperKeys = [
      'id',
      'event_id',
      'eventId',
      'source_event_id',
      'sourceEventId',
      'kind',
      'required',
      'received_at_ms',
      'receivedAtMs',
      'received_at',
      'timestamp',
      'ts',
    ];
    const looksLikeReplayWrapper = wrapperKeys.some((key) => Object.hasOwn(record, key));
    return looksLikeReplayWrapper && record.event != null ? record.event : record;
  }
  return record;
}

function replayRecordId(record = {}, index = 0) {
  return firstNonEmpty(record.id, record.event_id, record.eventId, record.source_event_id, record.sourceEventId, `provider-replay-${index + 1}`);
}

function replayRecordReceivedAt(record = {}, index = 0, options = {}) {
  return firstNonEmpty(
    record.received_at_ms,
    record.receivedAtMs,
    record.received_at,
    record.timestamp,
    record.ts,
    options.receivedAtMs,
    options.received_at_ms,
    replayBaseMs(options) + (index * 10_000),
  );
}

function mergeCoverage(rows = []) {
  const keys = ['realtime_axis', 'meeting_start', 'meeting_end', 'participant_track', 'speaker_activity', 'artifact_ready', 'subscription_lifecycle'];
  return Object.fromEntries(keys.map((key) => [key, rows.some((row) => row.coverage?.[key] === true)]));
}

function requiredCoverage(options = {}) {
  return unique(asArray(firstNonEmpty(
    options.requiredCoverage,
    options.required_coverage,
    options.target === 'production' ? ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'] : undefined,
    ['meeting_start', 'meeting_end'],
  )));
}

function replayRows(platform, records = [], options = {}) {
  return records.map((record, index) => {
    const payload = replayPayload(record);
    const receivedAtMs = replayRecordReceivedAt(record, index, options);
    const diagnostic = diagnosePlatformEvent(platform, payload, {
      ...options,
      receivedAtMs,
      includeRawSignals: true,
    });
    let runtimeEvent;
    let runtimeError;
    try {
      runtimeEvent = buildMeetingPlatformProviderRuntimeEvent(platform, payload, {
        ...options,
        event_id: replayRecordId(record, index),
        sent_at_ms: receivedAtMs,
      });
    } catch (error) {
      runtimeError = { name: error.name, message: error.message, details: error.details };
    }
    const signalTypes = diagnostic.signal_types ?? [];
    const actionableForAxis = Boolean(diagnostic.coverage?.meeting_start || diagnostic.coverage?.meeting_end || diagnostic.coverage?.participant_track || diagnostic.coverage?.artifact_ready);
    const accepted = diagnostic.ok === true
      && diagnostic.signal_count > 0
      && runtimeEvent?.action === 'provider_event';
    return compactObject({
      id: replayRecordId(record, index),
      kind: record.kind,
      required: record.required !== false,
      accepted,
      platform: diagnostic.platform,
      source: diagnostic.source,
      runtime_event_action: runtimeEvent?.action,
      runtime_event_schema: runtimeEvent?.schema,
      runtime_event_platform: runtimeEvent?.platform,
      sent_at_ms: runtimeEvent?.sent_at_ms,
      received_at_ms: receivedAtMs,
      signal_count: diagnostic.signal_count ?? 0,
      signal_types: signalTypes,
      coverage: diagnostic.coverage,
      actionable_for_axis: actionableForAxis,
      meetings: diagnostic.meetings,
      signals: diagnostic.signals,
      issues: diagnostic.issues,
      runtime_error: runtimeError,
    });
  });
}

export function buildMeetingPlatformProviderReplayReport(platformOrInput = {}, recordsOrOptions = undefined, options = {}) {
  const input = providerReplayInput(platformOrInput, recordsOrOptions, options);
  const platform = normalizeMeetingPlatform(input.platform);
  const adapter = meetingPlatformEventAdapterFor(platform);
  if (!adapter) {
    throw new MeetingTimelineSdkError(`Unsupported meeting platform: ${String(input.platform || '(empty)')}`, {
      platform: input.platform,
    });
  }
  const records = asArray(input.records).length > 0
    ? asArray(input.records)
    : sampleMeetingPlatformProviderEvents(platform, input.options);
  const rows = replayRows(platform, records, input.options);
  const coverage = mergeCoverage(rows);
  const required = requiredCoverage(input.options);
  const issues = [
    ...(records.length > 0 ? [] : ['missing_provider_event_records']),
    ...required.filter((key) => coverage[key] !== true).map((key) => `missing_coverage:${key}`),
    ...rows.flatMap((row) => row.accepted === true || row.required === false ? [] : [`${row.id}:not_accepted`]),
    ...rows.flatMap((row) => row.runtime_event_action === 'provider_event' ? [] : [`${row.id}:missing_provider_runtime_event`]),
  ];
  return compactObject({
    type: MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA,
    schema: MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA,
    schema_version: MEETING_PLATFORM_PROVIDER_REPLAY_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: firstNonEmpty(input.options.target, 'pilot'),
    platform,
    source: adapter.source,
    adapter: publicAdapter(adapter),
    record_count: records.length,
    accepted_record_count: rows.filter((row) => row.accepted === true).length,
    runtime_event_count: rows.filter((row) => row.runtime_event_action === 'provider_event').length,
    signal_count: rows.reduce((sum, row) => sum + (row.signal_count ?? 0), 0),
    signal_types: unique(rows.flatMap((row) => row.signal_types ?? [])),
    coverage,
    required_coverage: required,
    runtime_contract: {
      runtime_action: 'provider_event',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      local_observer_remains_primary_realtime_axis: true,
      replay_does_not_insert_annotations: true,
    },
    rows,
    issue_count: unique(issues).length,
    issues: unique(issues),
    next_actions: issues.length > 0
      ? unique(issues.map((item) => `fix_${item}`))
      : ['use_provider_events_for_reconcile_and_backfill_only'],
  });
}

export function buildMeetingPlatformProviderReplayAcceptanceReport(reportOrInput = {}, options = {}) {
  const report = reportOrInput?.schema === MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA
    ? reportOrInput
    : buildMeetingPlatformProviderReplayReport(reportOrInput, options);
  const issues = [];
  if (report.schema !== MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA) {
    issues.push(issue('error', 'invalid_schema', 'Expected a meeting platform provider replay report', {
      expected_schema: MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA,
      actual_schema: report.schema,
    }));
  }
  if (report.accepted !== true) issues.push(issue('error', 'provider_replay_not_accepted', 'Provider replay report accepted flag is not true'));
  if (report.runtime_contract?.runtime_action !== 'provider_event') {
    issues.push(issue('error', 'invalid_runtime_action', 'Provider replay must produce provider_event runtime events', {
      runtime_action: report.runtime_contract?.runtime_action,
    }));
  }
  if (report.runtime_contract?.provider_events_block_realtime !== false) {
    issues.push(issue('error', 'provider_events_block_realtime', 'Provider events must not block realtime annotations'));
  }
  for (const key of report.required_coverage ?? []) {
    if (report.coverage?.[key] !== true) {
      issues.push(issue('error', 'missing_required_coverage', 'Provider replay report is missing required coverage', {
        coverage: key,
      }));
    }
  }
  return compactObject({
    type: MEETING_PLATFORM_PROVIDER_REPLAY_ACCEPTANCE_SCHEMA,
    schema: MEETING_PLATFORM_PROVIDER_REPLAY_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_PROVIDER_REPLAY_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: report.target,
    platform: report.platform,
    record_count: report.record_count ?? 0,
    accepted_record_count: report.accepted_record_count ?? 0,
    runtime_event_count: report.runtime_event_count ?? 0,
    signal_count: report.signal_count ?? 0,
    required_coverage: report.required_coverage ?? [],
    coverage: report.coverage,
    issue_count: issues.length,
    issues,
    next_actions: issues.length > 0
      ? unique(issues.map((item) => `fix_${item.code}`))
      : report.next_actions ?? [],
  });
}

function replayRecordsForPlatform(platform, options = {}) {
  const source = firstNonEmpty(
    options.recordsByPlatform,
    options.records_by_platform,
    options.providerEventsByPlatform,
    options.provider_events_by_platform,
  );
  if (source && typeof source === 'object' && !Array.isArray(source)) {
    for (const key of unique([platform, platform.replaceAll('_', '-'), platform === 'microsoft_teams' ? 'teams' : undefined])) {
      if (source[key] != null) return source[key];
    }
  }
  return undefined;
}

function selectedReplayPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_PROVIDER_REPLAY_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

export function buildMeetingPlatformProviderReplayMatrix(options = {}) {
  const platforms = selectedReplayPlatforms(options);
  const reports = platforms.map((platform) => buildMeetingPlatformProviderReplayReport(platform, replayRecordsForPlatform(platform, options), options));
  const issues = reports.flatMap((report) => report.accepted === true ? [] : [`${report.platform}:not_accepted`]);
  return compactObject({
    type: MEETING_PLATFORM_PROVIDER_REPLAY_MATRIX_SCHEMA,
    schema: MEETING_PLATFORM_PROVIDER_REPLAY_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_PROVIDER_REPLAY_SCHEMA_VERSION,
    accepted: issues.length === 0,
    target: firstNonEmpty(options.target, 'pilot'),
    platform_count: platforms.length,
    accepted_count: reports.filter((report) => report.accepted === true).length,
    runtime_event_count: reports.reduce((sum, report) => sum + (report.runtime_event_count ?? 0), 0),
    signal_count: reports.reduce((sum, report) => sum + (report.signal_count ?? 0), 0),
    platforms,
    rows: reports.map((report) => compactObject({
      platform: report.platform,
      accepted: report.accepted,
      source: report.source,
      record_count: report.record_count,
      runtime_event_count: report.runtime_event_count,
      signal_count: report.signal_count,
      signal_types: report.signal_types,
      coverage: report.coverage,
      required_coverage: report.required_coverage,
      provider_events_block_realtime: report.runtime_contract?.provider_events_block_realtime,
      issue_count: report.issue_count,
      issues: report.issues,
    })),
    reports,
    issue_count: unique(issues).length,
    issues: unique(issues),
    next_actions: issues.length > 0
      ? unique(issues.map((item) => `fix_${item}`))
      : ['wire_provider_replay_matrix_into_adapter_ci'],
  });
}

export function assertMeetingPlatformProviderReplayReport(reportOrInput = {}, options = {}) {
  const report = reportOrInput?.schema === MEETING_PLATFORM_PROVIDER_REPLAY_REPORT_SCHEMA
    ? reportOrInput
    : buildMeetingPlatformProviderReplayReport(reportOrInput, options);
  const acceptance = buildMeetingPlatformProviderReplayAcceptanceReport(report);
  if (!acceptance.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform provider replay report is not accepted', {
      report,
      acceptance,
      issues: acceptance.issues,
    });
  }
  return report;
}

export function assertMeetingPlatformProviderReplayMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_PLATFORM_PROVIDER_REPLAY_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingPlatformProviderReplayMatrix({
      ...matrixOrOptions,
      ...options,
    });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform provider replay matrix is not accepted', {
      matrix,
      issues: matrix.issues,
    });
  }
  return matrix;
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
