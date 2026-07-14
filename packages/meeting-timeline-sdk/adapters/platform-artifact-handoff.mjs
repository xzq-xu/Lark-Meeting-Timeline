import { compactObject } from '../index.mjs';
import {
  PLATFORM_ARTIFACT_IMPORTERS,
  buildArtifactImportPlans,
} from './artifact-plan.mjs';
import { buildArtifactFetchRequest } from './artifact-fetch.mjs';
import { buildMeetingPlatformProviderConnectionPack } from './platform-provider-connection.mjs';
import {
  MEETING_PLATFORM_KEYS,
  buildPlatformIntegrationPlan,
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';

export const MEETING_PLATFORM_ARTIFACT_HANDOFF_PLAN_SCHEMA = 'meeting_platform_artifact_handoff_plan';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_MATRIX_SCHEMA = 'meeting_platform_artifact_handoff_matrix';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA = 'meeting_platform_artifact_handoff';
export const MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA_VERSION = 1;

const ARTIFACT_EVENT_TYPES = Object.freeze({
  google_meet: Object.freeze({
    transcript: ['google.workspace.meet.transcript.v2.fileGenerated'],
    recording: ['google.workspace.meet.recording.v2.fileGenerated'],
    smart_notes: ['google.workspace.meet.smartNotes.v2.fileGenerated'],
  }),
  microsoft_teams: Object.freeze({
    transcript: ['callTranscript.created', 'callTranscript.updated'],
    recording: ['callRecording.created', 'callRecording.updated'],
  }),
  zoom: Object.freeze({
    transcript: ['recording.completed'],
    recording: ['recording.completed'],
  }),
  webex: Object.freeze({
    transcript: ['meetingTranscripts.created'],
    recording: ['recordings.created', 'recordings.updated'],
  }),
  lark: Object.freeze({
    transcript: ['minutes.created', 'minutes.updated', 'minutes.export_ready'],
    recording: ['minutes.recording_ready', 'vc.meeting.recording_ready'],
  }),
});

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
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, MEETING_PLATFORM_KEYS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

function importEndpoint(options = {}) {
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, '')).replace(/\/+$/, '');
  return firstNonEmpty(
    options.importEndpoint,
    options.import_endpoint,
    baseUrl ? `${baseUrl}/api/import/transcript` : '/api/import/transcript',
  );
}

function artifactKindsFor(platform, options = {}) {
  const configured = PLATFORM_ARTIFACT_IMPORTERS[platform] ?? {};
  const requested = asArray(firstNonEmpty(options.artifactKinds, options.artifact_kinds));
  if (requested.length > 0) return requested.map((kind) => normalizeArtifactKind(kind)).filter((kind) => configured[kind]);
  return Object.keys(configured);
}

function normalizeArtifactKind(value) {
  const text = String(value || '').trim().toLowerCase();
  if (text === 'smartnote' || text === 'smart_note' || text === 'smart-notes') return 'smart_notes';
  return text;
}

function tokenKeyFor(platform) {
  return {
    google_meet: 'GOOGLE_WORKSPACE_ACCESS_TOKEN',
    microsoft_teams: 'MICROSOFT_GRAPH_ACCESS_TOKEN',
    zoom: 'ZOOM_ACCESS_TOKEN',
    webex: 'WEBEX_ACCESS_TOKEN',
    lark: 'LARK_ACCESS_TOKEN',
  }[platform] ?? `${String(platform).toUpperCase()}_ACCESS_TOKEN`;
}

function artifactRows(platform, options = {}) {
  const configured = PLATFORM_ARTIFACT_IMPORTERS[platform] ?? {};
  return artifactKindsFor(platform, options).map((kind) => {
    const importer = configured[kind] ?? {};
    return compactObject({
      artifact_kind: kind,
      supported: true,
      action: kind === 'transcript'
        ? 'fetch_and_import_transcript'
        : kind === 'smart_notes'
          ? 'fetch_and_store_smart_notes'
          : 'store_recording_artifact',
      event_types: ARTIFACT_EVENT_TYPES[platform]?.[kind] ?? [],
      fetch_strategy: importer.fetch_strategy,
      normalizer: importer.normalizer,
      content_hint: importer.content_hint,
      token_env: tokenKeyFor(platform),
      post_meeting_only: true,
      realtime_blocking: false,
    });
  });
}

function artifactSignals(input = {}) {
  if (Array.isArray(input)) return input;
  return asArray(firstNonEmpty(
    input.signals,
    input.artifactSignals,
    input.artifact_signals,
    input.events,
    input.rawSignals,
    input.raw_signals,
    [],
  ));
}

function planPlatform(plan = {}) {
  try {
    return normalizeMeetingPlatform(plan.platform);
  } catch {
    return plan.platform;
  }
}

function sanitizeHeaders(headers = {}) {
  return Object.fromEntries(Object.entries(headers)
    .filter(([, value]) => value != null && value !== '')
    .map(([key, value]) => {
      if (String(key).toLowerCase() === 'authorization') return [key, '<redacted>'];
      return [key, value];
    }));
}

function sanitizedFetchRequest(plan = {}, options = {}) {
  try {
    const request = buildArtifactFetchRequest(plan, options.fetchOptions ?? options.fetch_options ?? {});
    return {
      ok: true,
      request: compactObject({
        method: request.method,
        url: request.url,
        headers: sanitizeHeaders(request.headers),
        response_type: request.response_type,
      }),
    };
  } catch (error) {
    return {
      ok: false,
      error: String(error?.message ?? error),
    };
  }
}

function rowForImportPlan(plan = {}, options = {}) {
  const fetch = plan.status === 'requires_provider_fetch' || plan.status === 'metadata_ready'
    ? sanitizedFetchRequest(plan, options)
    : undefined;
  return compactObject({
    platform: plan.platform,
    status: plan.status,
    action: plan.action,
    artifact_kind: plan.artifact_kind,
    artifact_id: plan.artifact_id,
    artifact_url: plan.artifact_url,
    occurred_at_ms: plan.occurred_at_ms,
    source_event_id: plan.source_event_id,
    normalizer: plan.transcript_import?.normalizer,
    import_endpoint: plan.transcript_import?.endpoint,
    fetch_strategy: plan.fetch?.strategy,
    fetch_request_ready: fetch?.ok === true,
    fetch_request: fetch?.request,
    fetch_error: fetch?.ok === false ? fetch.error : undefined,
    issue_codes: (plan.issues ?? []).map((issue) => issue.code),
    issues: plan.issues,
  });
}

export function buildMeetingPlatformArtifactHandoffPlan(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const capability = platformCapabilityContract(key, options);
  const integration = buildPlatformIntegrationPlan(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const artifacts = artifactRows(key, options);
  return compactObject({
    type: 'meeting_platform_artifact_handoff_plan',
    schema: MEETING_PLATFORM_ARTIFACT_HANDOFF_PLAN_SCHEMA,
    schema_version: MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA_VERSION,
    platform: key,
    display_name: capability.display_name ?? integration.display_name ?? provider.display_name,
    status: artifacts.length > 0 ? 'post_meeting_artifacts_supported' : 'post_meeting_artifacts_not_configured',
    post_meeting_only: true,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    import_endpoint: importEndpoint(options),
    artifacts,
    supported_artifact_kinds: artifacts.map((artifact) => artifact.artifact_kind),
    transcript_supported: artifacts.some((artifact) => artifact.artifact_kind === 'transcript'),
    recording_supported: artifacts.some((artifact) => artifact.artifact_kind === 'recording'),
    smart_notes_supported: artifacts.some((artifact) => artifact.artifact_kind === 'smart_notes'),
    source_event_types: unique(artifacts.flatMap((artifact) => artifact.event_types ?? [])),
    sdk_modules: {
      artifact_plan: '@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan',
      artifact_fetch: '@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      platform_kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    },
    handoff_contract: {
      realtime_annotations_do_not_wait_for_artifacts: true,
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      import_after_meeting_end: true,
      preserve_provider_source_fields: true,
    },
    official_docs: provider.official_docs,
    next_actions: [
      'capture_artifact_ready_provider_events',
      'fetch_post_meeting_artifact_content_when_available',
      'import_transcript_after_meeting_end_without_rebinding_live_annotations',
    ],
  });
}

export function buildMeetingPlatformArtifactHandoffMatrix(options = {}) {
  const plans = selectedPlatforms(options).map((platform) => buildMeetingPlatformArtifactHandoffPlan(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_artifact_handoff_matrix',
    schema: MEETING_PLATFORM_ARTIFACT_HANDOFF_MATRIX_SCHEMA,
    schema_version: MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA_VERSION,
    platform_count: plans.length,
    transcript_supported_count: plans.filter((plan) => plan.transcript_supported).length,
    recording_supported_count: plans.filter((plan) => plan.recording_supported).length,
    smart_notes_supported_count: plans.filter((plan) => plan.smart_notes_supported).length,
    realtime_blocking_count: plans.filter((plan) => plan.provider_events_block_realtime || plan.transcript_blocks_realtime).length,
    platforms: plans.map((plan) => plan.platform),
    rows: plans.map((plan) => ({
      platform: plan.platform,
      display_name: plan.display_name,
      status: plan.status,
      supported_artifact_kinds: plan.supported_artifact_kinds,
      transcript_supported: plan.transcript_supported,
      recording_supported: plan.recording_supported,
      smart_notes_supported: plan.smart_notes_supported,
      source_event_types: plan.source_event_types,
      post_meeting_only: plan.post_meeting_only,
      transcript_blocks_realtime: plan.transcript_blocks_realtime,
      provider_events_block_realtime: plan.provider_events_block_realtime,
    })),
    plans,
    next_actions: unique(plans.flatMap((plan) => plan.next_actions ?? [])),
  };
}

export function buildMeetingPlatformArtifactHandoff(platform, input = {}, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const plan = buildMeetingPlatformArtifactHandoffPlan(key, options);
  const importPlans = buildArtifactImportPlans({
    signals: artifactSignals(input),
  }, {
    importEndpoint: plan.import_endpoint,
    includeIgnored: options.includeIgnored === true || options.include_ignored === true,
  }).filter((item) => planPlatform(item) === key || options.includeForeignPlatform === true || options.include_foreign_platform === true);
  const rows = importPlans.map((item) => rowForImportPlan(item, options));
  return compactObject({
    type: 'meeting_platform_artifact_handoff',
    schema: MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA,
    schema_version: MEETING_PLATFORM_ARTIFACT_HANDOFF_SCHEMA_VERSION,
    platform: key,
    status: rows.length > 0 ? 'artifact_handoff_ready' : 'no_artifact_signals',
    post_meeting_only: true,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    input_signal_count: artifactSignals(input).length,
    import_plan_count: importPlans.length,
    fetch_request_count: rows.filter((row) => row.fetch_request_ready).length,
    transcript_import_count: rows.filter((row) => row.action === 'fetch_and_import_transcript').length,
    recording_count: rows.filter((row) => row.artifact_kind === 'recording').length,
    warning_count: rows.reduce((total, row) => total + (row.issues ?? []).filter((issue) => issue.severity === 'warning').length, 0),
    plan,
    import_plans: importPlans,
    rows,
    fetch_requests: rows.filter((row) => row.fetch_request_ready).map((row) => row.fetch_request),
    next_actions: rows.length > 0
      ? ['execute_fetch_requests_after_meeting_end', 'import_transcripts_without_rebinding_live_annotations']
      : ['wait_for_artifact_ready_provider_events'],
  });
}
