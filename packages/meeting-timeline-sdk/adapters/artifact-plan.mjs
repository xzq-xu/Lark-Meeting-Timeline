import { compactObject } from '../index.mjs';

export const PLATFORM_ARTIFACT_IMPORTERS = Object.freeze({
  google_meet: Object.freeze({
    transcript: Object.freeze({
      normalizer: 'normalizeGoogleMeetTranscriptEntries',
      fetch_strategy: 'google_meet_rest_transcript_entries',
      content_hint: 'conferenceRecords.transcripts.entries.list',
    }),
    recording: Object.freeze({
      fetch_strategy: 'google_meet_rest_recording_or_drive_export',
      content_hint: 'conferenceRecords.recordings or Drive export URI',
    }),
    smart_notes: Object.freeze({
      fetch_strategy: 'google_meet_rest_smart_notes_or_docs_export',
      content_hint: 'conferenceRecords.smartNotes or Docs export URI',
    }),
  }),
  microsoft_teams: Object.freeze({
    transcript: Object.freeze({
      normalizer: 'normalizeMicrosoftTeamsTranscript',
      fetch_strategy: 'microsoft_graph_transcript_content',
      content_hint: 'onlineMeetings transcripts content or VTT',
    }),
    recording: Object.freeze({
      fetch_strategy: 'microsoft_graph_recording_content',
      content_hint: 'onlineMeetings recordings content or metadata',
    }),
  }),
  zoom: Object.freeze({
    transcript: Object.freeze({
      normalizer: 'normalizeZoomTranscript',
      fetch_strategy: 'zoom_recording_transcript_vtt',
      content_hint: 'recording transcript download_url VTT',
    }),
    recording: Object.freeze({
      fetch_strategy: 'zoom_recording_file_download',
      content_hint: 'recording_files download_url/play_url',
    }),
  }),
  webex: Object.freeze({
    transcript: Object.freeze({
      normalizer: 'normalizeWebexTranscript',
      fetch_strategy: 'webex_meeting_transcript_download',
      content_hint: 'meetingTranscripts txtDownloadLink or API content',
    }),
    recording: Object.freeze({
      fetch_strategy: 'webex_recording_download_or_metadata',
      content_hint: 'recordings temporaryDirectDownloadLinks or webLink',
    }),
  }),
  lark: Object.freeze({
    transcript: Object.freeze({
      normalizer: null,
      fetch_strategy: 'lark_minutes_transcript_export',
      content_hint: 'minutes transcript export payload',
    }),
    recording: Object.freeze({
      fetch_strategy: 'lark_minutes_or_meeting_artifact_metadata',
      content_hint: 'minutes or meeting artifact metadata',
    }),
  }),
});

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function lower(value) {
  return String(value || '').trim().toLowerCase();
}

function artifactKind(value) {
  const kind = lower(value);
  if (kind === 'smartnote' || kind === 'smart_note' || kind === 'smart-notes') return 'smart_notes';
  return kind || 'unknown';
}

function artifactSignals(input = {}) {
  if (Array.isArray(input)) return input;
  if (Array.isArray(input.signals)) return input.signals;
  if (Array.isArray(input.rawSignals)) return input.rawSignals;
  if (Array.isArray(input.raw_signals)) return input.raw_signals;
  if (Array.isArray(input.diagnostic?.raw_signals)) return input.diagnostic.raw_signals;
  if (Array.isArray(input.diagnostic?.signals)) return input.diagnostic.signals;
  return [input];
}

function signalPlatform(signal = {}) {
  return signal.meeting?.platform ?? signal.platform;
}

function fetchHint(signal = {}, importer = {}) {
  const platform = signalPlatform(signal);
  const meeting = signal.meeting ?? {};
  if (platform === 'google_meet' && signal.artifact_kind === 'transcript') {
    return compactObject({
      strategy: importer.fetch_strategy,
      resource: signal.artifact_id
        ? `conferenceRecords/${meeting.meeting_id}/transcripts/${signal.artifact_id}/entries`
        : `conferenceRecords/${meeting.meeting_id}/transcripts/{transcript-id}/entries`,
      auth: 'google_workspace_oauth_required',
    });
  }
  if (platform === 'microsoft_teams') {
    return compactObject({
      strategy: importer.fetch_strategy,
      resource: signal.artifact_id
        ? `onlineMeetings/${meeting.meeting_id}/artifacts/${signal.artifact_id}`
        : `onlineMeetings/${meeting.meeting_id}/artifacts/{artifact-id}`,
      auth: 'microsoft_graph_oauth_required',
    });
  }
  if (platform === 'zoom') {
    return compactObject({
      strategy: importer.fetch_strategy,
      url: signal.artifact_url,
      auth: 'zoom_oauth_or_download_token_required',
    });
  }
  if (platform === 'webex') {
    return compactObject({
      strategy: importer.fetch_strategy,
      url: signal.artifact_url,
      resource: signal.artifact_id ? `meetingTranscripts/${signal.artifact_id}` : undefined,
      auth: 'webex_oauth_required',
    });
  }
  return compactObject({
    strategy: importer.fetch_strategy,
    url: signal.artifact_url,
    auth: platform ? `${platform}_oauth_required` : 'provider_oauth_required',
  });
}

function planStatus(signal = {}, kind = '', importer = null) {
  if (signal.type !== 'artifact_ready') return 'ignored';
  if (!importer) return 'unsupported_artifact';
  if (kind === 'recording') return signal.artifact_url ? 'metadata_ready' : 'metadata_only';
  if (kind === 'transcript' || kind === 'smart_notes') return 'requires_provider_fetch';
  return 'metadata_only';
}

function actionFor(kind = '') {
  if (kind === 'transcript') return 'fetch_and_import_transcript';
  if (kind === 'recording') return 'store_recording_artifact';
  if (kind === 'smart_notes') return 'fetch_and_store_smart_notes';
  return 'store_artifact_metadata';
}

function issueFor(signal = {}, status = '') {
  if (status === 'ignored') {
    return [{ severity: 'info', code: 'not_artifact_signal', message: 'Signal is not artifact_ready and does not need artifact import planning.' }];
  }
  if (status === 'unsupported_artifact') {
    return [{ severity: 'warning', code: 'unsupported_artifact_kind', message: 'No platform artifact importer is configured for this artifact kind.' }];
  }
  if (!signal.artifact_id && !signal.artifact_url) {
    return [{ severity: 'warning', code: 'artifact_identity_missing', message: 'Artifact signal lacks artifact_id and artifact_url; provider fetch may need the raw event.' }];
  }
  return [];
}

export function buildArtifactImportPlan(signal = {}, options = {}) {
  const kind = artifactKind(signal.artifact_kind ?? signal.artifactKind);
  const platform = signalPlatform(signal);
  const platformImporters = PLATFORM_ARTIFACT_IMPORTERS[platform] ?? {};
  const importer = platformImporters[kind] ?? null;
  const status = planStatus(signal, kind, importer);
  const importEndpoint = options.importEndpoint ?? options.import_endpoint ?? '/api/import/transcript';
  return compactObject({
    status,
    action: actionFor(kind),
    platform,
    meeting: signal.meeting,
    source_event_id: signal.source_event_id,
    occurred_at_ms: signal.occurred_at_ms,
    artifact_kind: kind,
    artifact_id: signal.artifact_id,
    artifact_url: signal.artifact_url,
    fetch: importer ? fetchHint({ ...signal, artifact_kind: kind }, importer) : undefined,
    transcript_import: kind === 'transcript' ? {
      endpoint: importEndpoint,
      module: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      normalizer: importer?.normalizer,
      sdk_method: 'importPlatformTranscript',
    } : undefined,
    content_hint: importer?.content_hint,
    issues: issueFor(signal, status),
  });
}

export function buildArtifactImportPlans(input = {}, options = {}) {
  return artifactSignals(input)
    .map((signal) => buildArtifactImportPlan(signal, options))
    .filter((plan) => options.includeIgnored === true || options.include_ignored === true || plan.status !== 'ignored');
}
