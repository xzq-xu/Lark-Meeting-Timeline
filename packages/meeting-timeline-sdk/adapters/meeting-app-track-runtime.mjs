import {
  MeetingTimelineSdkError,
  compactObject,
  createMeetingTimelineClient,
} from '../index.mjs';
import { buildMeetingAppTrackPipeline } from './meeting-app-track-pipeline.mjs';

export const MEETING_APP_TRACK_RUNTIME_SCHEMA = 'meeting_app_track_runtime';
export const MEETING_APP_TRACK_RUNTIME_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function snapshotArray(input = {}) {
  if (Array.isArray(input)) return input;
  return asArray(firstNonEmpty(
    input.snapshots,
    input.samples,
    input.captures,
    input.rows,
    input.items,
    input.observations,
    input,
  ));
}

function isTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.insertMark === 'function';
}

function resolveClient(clientOrOptions, options = {}) {
  if (isTimelineClient(clientOrOptions)) return clientOrOptions;
  const clientOptions = {
    ...(clientOrOptions ?? {}),
    ...(options.clientOptions ?? options.client_options ?? {}),
  };
  if (!clientOptions.baseUrl && !clientOptions.base_url) {
    throw new MeetingTimelineSdkError('Meeting timeline client or client baseUrl is required for createMeetingAppTrackRuntime', {
      reason: 'missing_client',
    });
  }
  return createMeetingTimelineClient({
    ...clientOptions,
    baseUrl: clientOptions.baseUrl ?? clientOptions.base_url,
  });
}

function markKey(mark = {}) {
  return String(firstNonEmpty(
    mark.id,
    mark.annotation_id,
    [
      mark.intent,
      mark.kind,
      mark.source,
      mark.captured_at_ms,
      mark.label,
    ].filter((item) => item != null && item !== '').join(':'),
  ));
}

function maxSnapshots(options = {}) {
  const value = firstNonEmpty(options.maxSnapshots, options.max_snapshots, 500);
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric > 0 ? Math.floor(numeric) : 500;
}

function shouldInsert(options = {}, observeOptions = {}) {
  const value = firstNonEmpty(
    observeOptions.insert,
    observeOptions.apply,
    observeOptions.write,
    observeOptions.dryRun === true ? false : undefined,
    observeOptions.dry_run === true ? false : undefined,
    options.insert,
    options.apply,
    options.write,
    options.dryRun === true ? false : undefined,
    options.dry_run === true ? false : undefined,
    true,
  );
  return value !== false && value !== 'false';
}

function insertOptions(options = {}, observeOptions = {}) {
  return {
    ...(options.insertOptions ?? {}),
    ...(options.insert_options ?? {}),
    ...(observeOptions.insertOptions ?? {}),
    ...(observeOptions.insert_options ?? {}),
  };
}

async function insertMarks(client, marks = [], options = {}) {
  if (marks.length === 0) return [];
  if (typeof client.insertMarks === 'function' && options.forceSingle !== true && options.force_single !== true) {
    return [await client.insertMarks(marks, options)];
  }
  const results = [];
  for (const mark of marks) results.push(await client.insertMark(mark, options));
  return results;
}

export function createMeetingAppTrackRuntime(clientOrOptions, options = {}) {
  const client = resolveClient(clientOrOptions, options);
  let snapshots = snapshotArray(options.initialSnapshots ?? options.initial_snapshots ?? []);
  let seenMarkKeys = new Set(asArray(options.seenMarkKeys ?? options.seen_mark_keys).map(String));
  let observeCount = 0;
  let insertedCount = 0;
  let skippedDuplicateCount = 0;
  let lastReport = null;
  let lastInsertedMarks = [];

  async function observe(input = {}, observeOptions = {}) {
    const nextSnapshots = snapshotArray(input);
    snapshots.push(...nextSnapshots);
    const limit = maxSnapshots({ ...options, ...observeOptions });
    if (snapshots.length > limit) snapshots = snapshots.slice(-limit);
    const report = buildMeetingAppTrackPipeline(snapshots, {
      ...options,
      ...observeOptions,
    });
    const newMarks = [];
    const duplicateMarks = [];
    for (const mark of report.marks ?? []) {
      const key = markKey(mark);
      if (!key || seenMarkKeys.has(key)) {
        duplicateMarks.push(mark);
        continue;
      }
      newMarks.push(mark);
    }
    const apply = shouldInsert(options, observeOptions);
    const results = apply ? await insertMarks(client, newMarks, insertOptions(options, observeOptions)) : [];
    if (apply) {
      for (const mark of newMarks) seenMarkKeys.add(markKey(mark));
    }
    observeCount += 1;
    insertedCount += apply ? newMarks.length : 0;
    skippedDuplicateCount += duplicateMarks.length;
    lastReport = report;
    lastInsertedMarks = apply ? newMarks : [];
    return compactObject({
      type: 'meeting_app_track_runtime_observation',
      schema: MEETING_APP_TRACK_RUNTIME_SCHEMA,
      schema_version: MEETING_APP_TRACK_RUNTIME_SCHEMA_VERSION,
      inserted: apply,
      input_snapshot_count: nextSnapshots.length,
      retained_snapshot_count: snapshots.length,
      new_mark_count: newMarks.length,
      duplicate_mark_count: duplicateMarks.length,
      result_count: results.length,
      report,
      new_marks: newMarks,
      duplicate_marks: duplicateMarks,
      results,
      state: getState(),
    });
  }

  function preview(input = {}, previewOptions = {}) {
    const report = buildMeetingAppTrackPipeline([...snapshots, ...snapshotArray(input)], {
      ...options,
      ...previewOptions,
    });
    const newMarks = (report.marks ?? []).filter((mark) => !seenMarkKeys.has(markKey(mark)));
    return {
      report,
      new_marks: newMarks,
      new_mark_count: newMarks.length,
    };
  }

  function getState() {
    return {
      observe_count: observeCount,
      retained_snapshot_count: snapshots.length,
      seen_mark_count: seenMarkKeys.size,
      inserted_mark_count: insertedCount,
      skipped_duplicate_mark_count: skippedDuplicateCount,
      last_report_status: lastReport?.status,
      last_mark_count: lastReport?.mark_count ?? 0,
      last_inserted_mark_count: lastInsertedMarks.length,
      last_inserted_marks: lastInsertedMarks,
    };
  }

  function reset(nextState = {}) {
    snapshots = snapshotArray(nextState.snapshots ?? nextState.initialSnapshots ?? nextState.initial_snapshots ?? []);
    seenMarkKeys = new Set(asArray(nextState.seenMarkKeys ?? nextState.seen_mark_keys).map(String));
    observeCount = Number(nextState.observe_count ?? nextState.observeCount ?? 0);
    insertedCount = Number(nextState.inserted_mark_count ?? nextState.insertedMarkCount ?? 0);
    skippedDuplicateCount = Number(nextState.skipped_duplicate_mark_count ?? nextState.skippedDuplicateMarkCount ?? 0);
    lastReport = null;
    lastInsertedMarks = [];
    return getState();
  }

  return {
    client,
    observe,
    push: observe,
    preview,
    getState,
    getSnapshots() {
      return [...snapshots];
    },
    reset,
  };
}
