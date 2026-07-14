import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
import {
  MEETING_APP_DOM_CAPTURE_SCHEMA,
  captureMeetingAppDomSnapshot,
} from './meeting-app-capture.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_APP_SNAPSHOT_RECORD_SCHEMA = 'meeting_app_snapshot_record';
export const MEETING_APP_SNAPSHOT_RECORD_SCHEMA_VERSION = 1;
export const MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA = 'meeting_app_snapshot_record_set';
export const MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function tryNormalizePlatform(value) {
  if (!value || typeof value === 'object') return undefined;
  try {
    return normalizeMeetingPlatform(value);
  } catch {
    return String(value);
  }
}

function stableText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function recordTimestamp(input = {}, snapshot = {}, options = {}) {
  return normalizeAbsoluteMs(firstNonEmpty(
    options.capturedAtMs,
    options.captured_at_ms,
    options.observedAtMs,
    options.observed_at_ms,
    input.capturedAtMs,
    input.captured_at_ms,
    input.observedAtMs,
    input.observed_at_ms,
    snapshot.observedAtMs,
    snapshot.observed_at_ms,
    Date.now(),
  ), 'meeting_app_snapshot_record_time');
}

function recordPlatform(input = {}, snapshot = {}, options = {}) {
  return tryNormalizePlatform(firstNonEmpty(
    options.platform,
    options.provider,
    input.platform,
    input.provider,
    snapshot.platform,
    snapshot.provider,
    snapshot.capture?.profile,
  ));
}

function recordPhase(input = {}, options = {}) {
  const value = firstNonEmpty(options.phase, options.state, input.phase, input.state, input.fixture_state);
  if (value == null || value === '') return undefined;
  const text = String(value).trim().toLowerCase();
  if (['active', 'joined', 'in_meeting', 'in-meeting', 'start', 'started'].includes(text)) return 'active';
  if (['ended', 'left', 'prejoin', 'pre_join', 'pre-join', 'inactive', 'end'].includes(text)) return 'ended';
  return text;
}

function detectedMeeting(input = {}, snapshot = {}, options = {}) {
  return detectMeetingFromUrl({
    url: firstNonEmpty(options.url, input.url, snapshot.url, snapshot.meeting_url),
    title: firstNonEmpty(options.title, input.title, snapshot.title),
  }) ?? {};
}

function recordId(input = {}, snapshot = {}, options = {}) {
  const explicit = firstNonEmpty(options.id, options.recordId, options.record_id, input.id, input.recordId, input.record_id);
  if (explicit) return String(explicit);
  const platform = recordPlatform(input, snapshot, options) ?? 'unknown';
  const phase = recordPhase(input, options) ?? 'sample';
  const capturedAtMs = recordTimestamp(input, snapshot, options);
  return `${platform}-${phase}-${capturedAtMs}`;
}

function isCapturedSnapshot(input = {}) {
  return input?.schema === MEETING_APP_DOM_CAPTURE_SCHEMA
    || input?.capture?.profile
    || input?.page
    || input?.dom;
}

function snapshotFromInput(input = {}, options = {}) {
  if (options.snapshot) return options.snapshot;
  if (input?.snapshot) return input.snapshot;
  if (isCapturedSnapshot(input)) return input;
  return captureMeetingAppDomSnapshot(input, {
    ...(options.captureOptions ?? {}),
    ...(options.capture_options ?? {}),
    observedAtMs: firstNonEmpty(
      options.observedAtMs,
      options.observed_at_ms,
      options.capturedAtMs,
      options.captured_at_ms,
    ),
  });
}

export function buildMeetingAppSnapshotRecord(input = {}, options = {}) {
  if (input?.schema === MEETING_APP_SNAPSHOT_RECORD_SCHEMA && !options.forceRebuild && !options.force_rebuild) {
    return {
      ...input,
      snapshot: input.snapshot,
    };
  }
  const snapshot = snapshotFromInput(input, options);
  const detected = detectedMeeting(input, snapshot, options);
  const capturedAtMs = recordTimestamp(input, snapshot, options);
  const platform = recordPlatform({
    ...input,
    platform: firstNonEmpty(input.platform, detected.platform),
  }, snapshot, options);
  const phase = recordPhase(input, options);
  return compactObject({
    schema: MEETING_APP_SNAPSHOT_RECORD_SCHEMA,
    schema_version: MEETING_APP_SNAPSHOT_RECORD_SCHEMA_VERSION,
    id: recordId(input, snapshot, options),
    label: firstNonEmpty(options.label, input.label),
    phase,
    platform,
    provider: platform,
    capturedAtMs,
    captured_at_ms: capturedAtMs,
    url: firstNonEmpty(options.url, input.url, snapshot.url, snapshot.meeting_url),
    title: firstNonEmpty(options.title, input.title, snapshot.title),
    meeting_id: firstNonEmpty(options.meetingId, options.meeting_id, input.meetingId, input.meeting_id, snapshot.meeting_id, detected.meeting_id),
    external_meeting_id: firstNonEmpty(options.externalMeetingId, options.external_meeting_id, input.externalMeetingId, input.external_meeting_id, snapshot.external_meeting_id, detected.external_meeting_id),
    source: firstNonEmpty(options.source, input.source, snapshot.source, 'meeting_app_snapshot_recorder'),
    notes: firstNonEmpty(options.notes, input.notes),
    snapshot,
  });
}

export function buildMeetingAppSnapshotRecordSet(records = [], options = {}) {
  const rows = asArray(records).map((record) => buildMeetingAppSnapshotRecord(record));
  const platforms = [...new Set(rows.map((record) => record.platform).filter(Boolean))];
  const createdAtMs = normalizeAbsoluteMs(firstNonEmpty(
    options.createdAtMs,
    options.created_at_ms,
    Date.now(),
  ), 'meeting_app_snapshot_record_set_time');
  return compactObject({
    schema: MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA,
    schema_version: MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA_VERSION,
    id: firstNonEmpty(options.id, options.recordSetId, options.record_set_id, `meeting-app-snapshots-${createdAtMs}`),
    createdAtMs,
    created_at_ms: createdAtMs,
    source: firstNonEmpty(options.source, 'meeting_app_snapshot_recorder'),
    label: firstNonEmpty(options.label),
    platform_count: platforms.length,
    platforms,
    record_count: rows.length,
    records: rows,
  });
}

export function meetingAppSnapshotRecords(input = []) {
  if (input?.schema === MEETING_APP_SNAPSHOT_RECORD_SET_SCHEMA) return asArray(input.records);
  if (input?.records) return asArray(input.records);
  return asArray(input);
}

export function meetingAppSnapshotsFromRecords(input = [], options = {}) {
  return meetingAppSnapshotRecords(input)
    .map((item) => item?.schema === MEETING_APP_SNAPSHOT_RECORD_SCHEMA ? item.snapshot : buildMeetingAppSnapshotRecord(item, options).snapshot)
    .filter(Boolean);
}

export function buildMeetingAppLaunchGateInputFromRecords(input = [], options = {}) {
  const records = meetingAppSnapshotRecords(input).map((item) => buildMeetingAppSnapshotRecord(item, options));
  const snapshots = {};
  for (const record of records) {
    const platform = record.platform ?? record.snapshot?.platform ?? record.snapshot?.capture?.profile ?? 'unknown';
    if (!snapshots[platform]) snapshots[platform] = [];
    snapshots[platform].push(record.snapshot);
  }
  return {
    snapshots,
    snapshot_records: records,
  };
}

export function createMeetingAppSnapshotRecorder(options = {}) {
  const records = [];

  function add(input = {}, recordOptions = {}) {
    const record = buildMeetingAppSnapshotRecord(input, {
      ...options,
      ...recordOptions,
    });
    records.push(record);
    return record;
  }

  function capture(input = {}, captureOptions = {}) {
    return add(input, captureOptions);
  }

  function exportRecords(exportOptions = {}) {
    return buildMeetingAppSnapshotRecordSet(records, {
      ...options,
      ...exportOptions,
    });
  }

  function gateInput(inputOptions = {}) {
    return buildMeetingAppLaunchGateInputFromRecords(records, inputOptions);
  }

  function reset() {
    const removed = records.length;
    records.splice(0);
    return { removed };
  }

  function getState() {
    const platforms = [...new Set(records.map((record) => record.platform).filter(Boolean))];
    const phases = [...new Set(records.map((record) => record.phase).filter(Boolean))];
    return {
      record_count: records.length,
      platforms,
      phases,
      latest_record_id: records.at(-1)?.id,
    };
  }

  return {
    add,
    capture,
    exportRecords,
    gateInput,
    reset,
    getState,
    records() {
      return records.slice();
    },
    findByLabel(label) {
      const text = stableText(label);
      return records.filter((record) => stableText(record.label) === text);
    },
  };
}
