import { MeetingTimelineSdkError, compactObject } from './internal-utils.mjs';

export const LARK_MEETING_EVENT_TYPES = Object.freeze([
  'vc.meeting.all_meeting_started_v1',
  'vc.meeting.all_meeting_ended_v1',
  'vc.meeting.meeting_started_v1',
  'vc.meeting.meeting_ended_v1',
  'vc.meeting.join_meeting_v1',
  'vc.meeting.leave_meeting_v1',
]);

export const LOCAL_DETECTOR_EVENT_TYPES = Object.freeze([
  'meeting_started',
  'meeting_ended',
  'participant_joined',
  'participant_left',
  'speaker_started',
  'speaker_ended',
]);

export const GOOGLE_MEET_EVENT_TYPES = Object.freeze([
  'google.workspace.meet.conference.v2.started',
  'google.workspace.meet.conference.v2.ended',
  'google.workspace.meet.participant.v2.joined',
  'google.workspace.meet.participant.v2.left',
  'google.workspace.meet.recording.v2.fileGenerated',
  'google.workspace.meet.transcript.v2.fileGenerated',
  'google.workspace.meet.smartNote.v2.fileGenerated',
]);

export const GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES = Object.freeze([
  'google.workspace.events.subscription.v1.suspended',
  'google.workspace.events.subscription.v1.expirationReminder',
  'google.workspace.events.subscription.v1.expired',
]);

export const MICROSOFT_TEAMS_CHANGE_TYPES = Object.freeze(['created', 'updated']);

export const MICROSOFT_GRAPH_LIFECYCLE_EVENTS = Object.freeze([
  'reauthorizationRequired',
  'subscriptionRemoved',
  'missed',
]);

export const ZOOM_MEETING_EVENT_TYPES = Object.freeze([
  'meeting.started',
  'meeting.ended',
  'meeting.participant_joined',
  'meeting.participant_left',
  'recording.completed',
]);

export const WEBEX_WEBHOOK_RESOURCES = Object.freeze([
  { resource: 'meetings', events: ['started', 'ended'] },
  { resource: 'meetingParticipants', events: ['joined', 'left'] },
  { resource: 'recordings', events: ['created', 'updated'] },
  { resource: 'meetingTranscripts', events: ['created'] },
]);

export const MEETING_PLATFORM_KEYS = Object.freeze(['local_detector', 'lark', 'google_meet', 'microsoft_teams', 'zoom', 'webex']);

export const MEETING_PLATFORM_ALIASES = Object.freeze({
  'local-detector': 'local_detector',
  local_detector: 'local_detector',
  detector: 'local_detector',
  'desktop-observer': 'local_detector',
  desktop_observer: 'local_detector',
  observer: 'local_detector',
  manual: 'local_detector',
  lark: 'lark',
  feishu: 'lark',
  'fei-shu': 'lark',
  larksuite: 'lark',
  'lark-suite': 'lark',
  'google-meet': 'google_meet',
  google_meet: 'google_meet',
  meet: 'google_meet',
  'microsoft-teams': 'microsoft_teams',
  microsoft_teams: 'microsoft_teams',
  teams: 'microsoft_teams',
  zoom: 'zoom',
  webex: 'webex',
  'cisco-webex': 'webex',
  cisco_webex: 'webex',
});

const platformCapabilityContracts = Object.freeze({
  local_detector: {
    platform: 'local_detector',
    display_name: 'Local Meeting Detector',
    realtime_axis: {
      status: 'supported',
      source: 'Desktop observer, browser extension, e-ink host app, or explicit manual signal',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'primary_low_latency_path_when_official_events_are_delayed_or_unavailable',
    },
    participant_track: {
      status: 'supported_if_detector_provides_roster_changes',
      source: 'local observer participant signals',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'supported_if_detector_provides_active_speaker_changes',
      source: 'local observer active speaker signals',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'post_meeting_transcript_segments_can_backfill_speaker_positions',
    },
    post_meeting_transcript: {
      status: 'not_applicable',
      availability: 'provider_specific_post_meeting',
      detail: 'local_detector_only_establishes_the_axis; transcript_import_still_uses_the_detected_provider_or_generic_import',
    },
    recording: {
      status: 'not_applicable',
      detail: 'recording_artifacts_should_be_imported_from_the_detected_provider',
    },
    subscription_lifecycle: {
      status: 'not_applicable',
      detail: 'local_detector_has_no_remote_subscription_lifecycle',
    },
    realtime_transcript: {
      status: 'not_supported',
      detail: 'local_detector_does_not_require_realtime_transcript_for_annotation_alignment',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/local-detector',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      webhook_router: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'detector_must_provide_absolute_time_for_reliable_alignment',
      'detected_provider_identity_can_be_partial_without_url_or_meeting_id',
      'provider_webhook_events_should_later_reconcile_or_backfill_the_axis_when_available',
    ],
  },
  lark: {
    platform: 'lark',
    display_name: 'Feishu / Lark',
    realtime_axis: {
      status: 'supported',
      source: 'Feishu/Lark long-connection events or HTTP event callback',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'current_user_meeting_scan_or_local_detector_recommended_when_event_delivery_lags',
    },
    participant_track: {
      status: 'supported_best_effort',
      source: 'Feishu/Lark join_meeting_v1 and leave_meeting_v1 events',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'not_supported_by_official_realtime_events',
      source: 'local_detector_or_post_meeting_minutes',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'use_local_detector_for_realtime_speaker_markers_and_minutes_for_backfill',
    },
    post_meeting_transcript: {
      status: 'supported',
      availability: 'post_meeting',
      source: 'Feishu/Lark Minutes search/basic/transcript export APIs',
      import_endpoint: '/api/import/transcript',
      sdk_normalizer: null,
    },
    recording: {
      status: 'metadata_supported',
      availability: 'post_meeting',
      source: 'Feishu/Lark meeting artifact or minutes metadata when available',
      signal_types: ['artifact_ready'],
    },
    subscription_lifecycle: {
      status: 'not_applicable_for_long_connection',
      detail: 'local_demo_prefers_long_connection_event_delivery_and_current_user_scan_fallback',
    },
    realtime_transcript: {
      status: 'not_supported',
      detail: 'native_path_is_post_meeting_minutes_import',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/lark',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'event_delivery_can_lag_or_depend_on_app_publication_and_event_subscription_mode',
      'all_meeting_events_require_tenant_level_meeting_permission',
      'minutes_transcript_export_can_require_separate_oauth_scopes_or_approval',
    ],
  },
  google_meet: {
    platform: 'google_meet',
    display_name: 'Google Meet',
    realtime_axis: {
      status: 'supported_best_effort',
      source: 'Google Workspace Events conference started/ended',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'local_detector_recommended_for_low_latency_axis',
    },
    participant_track: {
      status: 'supported_best_effort',
      source: 'Google Workspace Events participant joined/left',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'not_supported_by_workspace_events',
      source: 'local_detector_or_post_meeting_transcript_entries',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'use_local_detector_for_realtime_speaker_markers_and_conferenceRecords_transcripts_entries_for_backfill',
    },
    post_meeting_transcript: {
      status: 'supported',
      availability: 'post_meeting',
      source: 'Google Meet REST conferenceRecords.transcripts.entries',
      import_endpoint: '/api/import/transcript',
      sdk_normalizer: 'normalizeGoogleMeetTranscriptEntries',
    },
    recording: {
      status: 'metadata_supported',
      availability: 'post_meeting',
      source: 'Google Meet REST conferenceRecords.recordings and Workspace Events recording fileGenerated',
      signal_types: ['artifact_ready'],
    },
    subscription_lifecycle: {
      status: 'supported',
      signal_types: ['subscription_lifecycle'],
      events: GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES,
    },
    realtime_transcript: {
      status: 'not_supported',
      detail: 'native_path_is_post_meeting_transcript_import',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/google-meet',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'workspace_event_delivery_can_lag_or_depend_on_subscription_scope',
      'transcript_entries_may_differ_from_google_docs_transcript',
      'low_latency_annotation_should_not_wait_for_transcript',
    ],
  },
  microsoft_teams: {
    platform: 'microsoft_teams',
    display_name: 'Microsoft Teams',
    realtime_axis: {
      status: 'supported_best_effort',
      source: 'Microsoft Graph meetingCallEvents callStarted/callEnded',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'local_detector_recommended_when_graph_subscription_is_unavailable',
    },
    participant_track: {
      status: 'supported_best_effort',
      source: 'Microsoft Graph meetingCallEvents rosterUpdated',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'not_supported_by_graph_meeting_call_events',
      source: 'local_detector_or_post_meeting_callTranscript',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'use_local_detector_for_realtime_speaker_markers_and_callTranscript_for_backfill',
    },
    post_meeting_transcript: {
      status: 'supported',
      availability: 'post_meeting',
      source: 'Microsoft Graph callTranscript content',
      import_endpoint: '/api/import/transcript',
      sdk_normalizer: 'normalizeMicrosoftTeamsTranscript',
    },
    recording: {
      status: 'metadata_supported',
      availability: 'post_meeting',
      source: 'Microsoft Graph recording notifications and recording content APIs',
      signal_types: ['artifact_ready'],
    },
    subscription_lifecycle: {
      status: 'supported',
      signal_types: ['subscription_lifecycle'],
      events: MICROSOFT_GRAPH_LIFECYCLE_EVENTS,
    },
    realtime_transcript: {
      status: 'not_supported',
      detail: 'native_path_is_post_meeting_transcript_import',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/microsoft-teams',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'graph_subscription_requires_tenant_admin_permissions',
      'subscription_renewal_is_required',
      'transcript_api_can_be_disabled_by_tenant_policy',
    ],
  },
  zoom: {
    platform: 'zoom',
    display_name: 'Zoom',
    realtime_axis: {
      status: 'supported_best_effort',
      source: 'Zoom Meeting webhooks meeting.started/meeting.ended',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'local_detector_recommended_for_desktop_or_browser_join',
    },
    participant_track: {
      status: 'supported_best_effort',
      source: 'Zoom Meeting webhooks participant_joined/participant_left',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'not_supported_by_zoom_meeting_webhooks',
      source: 'local_detector_or_post_meeting_transcript_vtt',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'use_local_detector_for_realtime_speaker_markers_and_recording_transcript_for_backfill',
    },
    post_meeting_transcript: {
      status: 'supported_when_cloud_recording_transcript_enabled',
      availability: 'post_meeting',
      source: 'Zoom recording transcript VTT downloaded from recording files',
      import_endpoint: '/api/import/transcript',
      sdk_normalizer: 'normalizeZoomTranscript',
    },
    recording: {
      status: 'supported',
      availability: 'post_meeting',
      source: 'Zoom recording.completed webhook',
      signal_types: ['artifact_ready'],
    },
    subscription_lifecycle: {
      status: 'not_applicable',
      detail: 'zoom_event_subscriptions_do_not_use_short_lived_graph_workspace_style_subscriptions',
    },
    realtime_transcript: {
      status: 'not_supported',
      detail: 'native_path_is_post_meeting_transcript_import',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/zoom',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'webhook_endpoint_must_ack_quickly',
      'cloud_recording_and_transcript_must_be_enabled',
      'download_url_access_can_require_zoom_auth_context',
    ],
  },
  webex: {
    platform: 'webex',
    display_name: 'Cisco Webex',
    realtime_axis: {
      status: 'supported_best_effort',
      source: 'Webex webhooks meetings started/ended',
      signal_types: ['meeting_started', 'meeting_ended'],
      fallback: 'local_detector_recommended_for_low_latency_axis',
    },
    participant_track: {
      status: 'supported_best_effort',
      source: 'Webex webhooks meetingParticipants joined/left',
      signal_types: ['participant_joined', 'participant_left'],
    },
    speaker_activity: {
      status: 'not_supported_by_webex_meeting_webhooks',
      source: 'local_detector_or_post_meeting_transcript',
      signal_types: ['speaker_started', 'speaker_ended'],
      fallback: 'use_local_detector_for_realtime_speaker_markers_and_meeting_transcripts_for_backfill',
    },
    post_meeting_transcript: {
      status: 'supported',
      availability: 'post_meeting',
      source: 'Webex Meeting Transcripts API VTT/text content or txtDownloadLink artifact',
      import_endpoint: '/api/import/transcript',
      sdk_normalizer: 'normalizeWebexTranscript',
    },
    recording: {
      status: 'metadata_supported',
      availability: 'post_meeting',
      source: 'Webex recordings webhook/API',
      signal_types: ['artifact_ready'],
    },
    subscription_lifecycle: {
      status: 'not_applicable',
      detail: 'webex_webhooks_do_not_use_short_lived_graph_workspace_style_subscriptions',
    },
    realtime_transcript: {
      status: 'not_supported_for_server_adapter',
      detail: 'backend_timeline_uses_post_meeting_transcript_import',
    },
    sdk_modules: {
      events: '@ai-annotation/meeting-timeline-sdk/adapters/webex',
      url_detection: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-url',
      local_observer: '@ai-annotation/meeting-timeline-sdk/adapters/local-observer',
      ingest: '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest',
      webhook_handler: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler',
      transcript: '@ai-annotation/meeting-timeline-sdk/adapters/transcript',
      setup: '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup',
      security: '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security',
    },
    limitations: [
      'webhook_payload_can_require_rest_enrichment_for_full_details',
      'transcripts_are_post_meeting_artifacts_not_realtime_axis_dependencies',
      'low_latency_annotation_should_not_wait_for_transcript',
    ],
  },
});

const platformAliases = new Map(Object.entries(MEETING_PLATFORM_ALIASES));

const PLATFORM_PERMISSION_FEATURE_ALIASES = Object.freeze({
  realtime: 'realtime_axis',
  realtime_axis: 'realtime_axis',
  axis: 'realtime_axis',
  meeting_axis: 'realtime_axis',
  meeting_lifecycle: 'realtime_axis',
  lifecycle: 'subscription_lifecycle',
  subscription: 'subscription_lifecycle',
  subscription_lifecycle: 'subscription_lifecycle',
  participant: 'participant_track',
  participants: 'participant_track',
  participant_track: 'participant_track',
  roster: 'participant_track',
  speaker: 'speaker_activity',
  speakers: 'speaker_activity',
  speaker_activity: 'speaker_activity',
  active_speaker: 'speaker_activity',
  transcript: 'post_meeting_transcript',
  transcripts: 'post_meeting_transcript',
  minutes: 'post_meeting_transcript',
  post_meeting_transcript: 'post_meeting_transcript',
  recording: 'recording',
  recordings: 'recording',
  artifact: 'artifact',
  artifacts: 'artifact',
  security: 'webhook_security',
  webhook_security: 'webhook_security',
  acceptance: 'acceptance',
  diagnostics: 'acceptance',
});

const PLATFORM_DEFAULT_PERMISSION_FEATURES = Object.freeze([
  'realtime_axis',
  'participant_track',
  'speaker_activity',
  'post_meeting_transcript',
  'recording',
  'subscription_lifecycle',
  'webhook_security',
]);

const PLATFORM_PERMISSION_REQUIREMENTS = Object.freeze({
  local_detector: {
    realtime_axis: {
      signal_types: ['meeting_started', 'meeting_ended'],
      event_types: ['meeting_started', 'meeting_ended'],
      setup: ['Host app must send absolute meeting start/end timestamps.'],
    },
    participant_track: {
      signal_types: ['participant_joined', 'participant_left'],
      event_types: ['participant_joined', 'participant_left'],
      setup: ['Optional: host observer must expose roster changes.'],
    },
    speaker_activity: {
      signal_types: ['speaker_started', 'speaker_ended'],
      event_types: ['speaker_started', 'speaker_ended'],
      setup: ['Optional: host observer must expose active-speaker changes.'],
    },
  },
  lark: {
    realtime_axis: {
      permissions: ['vc:meeting.all_meeting:readonly'],
      event_types: [
        'vc.meeting.all_meeting_started_v1',
        'vc.meeting.all_meeting_ended_v1',
        'vc.meeting.meeting_started_v1',
        'vc.meeting.meeting_ended_v1',
      ],
    },
    participant_track: {
      permissions: ['vc:meeting.all_meeting:readonly'],
      event_types: ['vc.meeting.join_meeting_v1', 'vc.meeting.leave_meeting_v1'],
    },
    post_meeting_transcript: {
      permissions: [
        'minutes:minutes.search:read',
        'minutes:minutes.basic:read',
        'minutes:minutes.transcript:export',
      ],
    },
  },
  google_meet: {
    realtime_axis: {
      scopes: ['https://www.googleapis.com/auth/meetings.space.readonly'],
      event_types: [
        'google.workspace.meet.conference.v2.started',
        'google.workspace.meet.conference.v2.ended',
      ],
    },
    participant_track: {
      scopes: ['https://www.googleapis.com/auth/meetings.space.readonly'],
      event_types: [
        'google.workspace.meet.participant.v2.joined',
        'google.workspace.meet.participant.v2.left',
      ],
    },
    post_meeting_transcript: {
      scopes: [
        'https://www.googleapis.com/auth/meetings.space.readonly',
        'https://www.googleapis.com/auth/drive.meet.readonly',
      ],
      event_types: ['google.workspace.meet.transcript.v2.fileGenerated'],
      setup: ['Use the Meet event to know an artifact exists; use Drive meet-readonly scope to fetch the transcript file.'],
    },
    recording: {
      scopes: [
        'https://www.googleapis.com/auth/meetings.space.readonly',
        'https://www.googleapis.com/auth/drive.meet.readonly',
      ],
      event_types: ['google.workspace.meet.recording.v2.fileGenerated'],
      setup: ['Use Drive meet-readonly scope for generated Meet recording artifacts.'],
    },
    subscription_lifecycle: {
      scopes: ['https://www.googleapis.com/auth/meetings.space.readonly'],
      event_types: GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES,
    },
  },
  microsoft_teams: {
    realtime_axis: {
      permissions: ['OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'],
      event_types: ['meetingCallEvents.created', 'meetingCallEvents.updated'],
    },
    participant_track: {
      permissions: ['OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'],
      event_types: ['meetingCallEvents.updated'],
      setup: ['Use rich notifications when possible to avoid follow-up reads for roster changes.'],
    },
    post_meeting_transcript: {
      permissions: ['OnlineMeetingTranscript.Read.All or OnlineMeetingTranscript.Read.Chat for resource-specific consent'],
      event_types: ['callTranscript.created'],
    },
    recording: {
      permissions: ['OnlineMeetingRecording.Read.All'],
      event_types: ['callRecording.created'],
    },
    subscription_lifecycle: {
      permissions: ['OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'],
      event_types: MICROSOFT_GRAPH_LIFECYCLE_EVENTS,
    },
  },
  zoom: {
    realtime_axis: {
      scopes: ['meeting:read:meeting or meeting:read:meeting:admin'],
      event_types: ['meeting.started', 'meeting.ended'],
    },
    participant_track: {
      scopes: ['meeting:read:participant or meeting:read:participant:admin'],
      event_types: ['meeting.participant_joined', 'meeting.participant_left'],
    },
    post_meeting_transcript: {
      scopes: ['cloud_recording:read:recording or cloud_recording:read:recording:admin'],
      event_types: ['recording.completed'],
      setup: ['Treat transcript files as post-meeting recording artifacts; do not block realtime annotation on this event.'],
    },
    recording: {
      scopes: ['cloud_recording:read:recording or cloud_recording:read:recording:admin'],
      event_types: ['recording.completed'],
    },
  },
  webex: {
    realtime_axis: {
      scopes: ['meeting:schedules_read'],
      admin_scopes: ['meeting:admin_schedule_read'],
      event_types: ['meetings.started', 'meetings.ended'],
    },
    participant_track: {
      scopes: ['meeting:participants_read'],
      admin_scopes: ['meeting:admin_participants_read'],
      event_types: ['meetingParticipants.joined', 'meetingParticipants.left'],
    },
    post_meeting_transcript: {
      scopes: ['meeting:transcripts_read'],
      admin_scopes: ['meeting:admin_transcripts_read'],
      event_types: ['meetingTranscripts.created'],
    },
    recording: {
      scopes: ['meeting:recordings_read'],
      admin_scopes: ['meeting:admin_recordings_read'],
      event_types: ['recordings.created', 'recordings.updated'],
    },
  },
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function ensureBaseUrl(baseUrl) {
  if (!baseUrl) throw new MeetingTimelineSdkError('baseUrl is required to build platform setup');
  return String(baseUrl).replace(/\/+$/, '');
}

function absoluteEndpoint(baseUrl, path) {
  return `${ensureBaseUrl(baseUrl)}${path.startsWith('/') ? path : `/${path}`}`;
}

function addSeconds(date, seconds) {
  return new Date(date.getTime() + seconds * 1000).toISOString();
}

function parseTimeMs(value) {
  if (value == null || value === '') return null;
  if (value instanceof Date) {
    const ms = value.getTime();
    return Number.isFinite(ms) ? ms : null;
  }
  const parsed = Date.parse(String(value));
  return Number.isFinite(parsed) ? parsed : null;
}

export function normalizeMeetingPlatform(platform) {
  const key = platformAliases.get(String(platform ?? '').toLowerCase().replace(/\s+/g, '-'));
  if (!key) {
    throw new MeetingTimelineSdkError(`Unsupported meeting platform: ${platform}`, {
      supportedPlatforms: MEETING_PLATFORM_KEYS,
    });
  }
  return key;
}

function normalizePlatform(platform) {
  return normalizeMeetingPlatform(platform);
}

function urlEncode(value) {
  return encodeURIComponent(String(value));
}

function envValue(env = {}, name) {
  return env[name] ?? env[name.toLowerCase()] ?? env[name.replace(/_/g, '-')];
}

function hasEnv(env = {}, name) {
  const value = envValue(env, name);
  return value != null && value !== '';
}

function endpointReadiness(endpoint) {
  if (!endpoint) return { ok: false, reason: 'endpoint_missing' };
  let parsed;
  try {
    parsed = new URL(endpoint);
  } catch {
    return { ok: false, reason: 'endpoint_invalid' };
  }
  const host = parsed.hostname.toLowerCase();
  const local = ['localhost', '127.0.0.1', '::1'].includes(host);
  if (parsed.protocol !== 'https:' && !local) {
    return { ok: false, reason: 'endpoint_must_be_https_or_localhost', protocol: parsed.protocol };
  }
  return { ok: true, reason: local ? 'local_development_endpoint' : 'public_https_endpoint' };
}

function missingEnv(env = {}, names = []) {
  return names.filter((name) => !hasEnv(env, name));
}

function presentEnv(env = {}, names = []) {
  return names.filter((name) => hasEnv(env, name));
}

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function normalizePermissionFeature(feature) {
  const key = String(feature ?? '').trim().toLowerCase().replace(/[\s-]+/g, '_');
  const normalized = PLATFORM_PERMISSION_FEATURE_ALIASES[key] ?? key;
  if (!normalized) {
    throw new MeetingTimelineSdkError(`Unsupported platform permission feature: ${feature}`);
  }
  return normalized;
}

function selectedPermissionFeatures(platform, options = {}) {
  const raw = firstNonEmpty(options.features, options.featureSet, options.feature_set);
  if (raw == null) {
    return platform === 'local_detector'
      ? ['realtime_axis', 'participant_track', 'speaker_activity']
      : [...PLATFORM_DEFAULT_PERMISSION_FEATURES];
  }
  const values = Array.isArray(raw) ? raw : String(raw).split(/[,\s]+/);
  return uniqueList(values.map((item) => normalizePermissionFeature(item)));
}

function platformFeatureRequirement(platform, manifest, feature) {
  const requirements = PLATFORM_PERMISSION_REQUIREMENTS[platform]?.[feature] ?? {};
  const capability = manifest.capabilities?.[feature];
  const requiredSecurity = feature === 'webhook_security'
    ? manifest.required_security_env ?? []
    : [];
  const optionalSecurity = feature === 'webhook_security'
    ? manifest.optional_security_env ?? []
    : [];
  return compactObject({
    feature,
    status: capability?.status ?? (Object.keys(requirements).length > 0 ? 'supported' : 'not_declared'),
    source: capability?.source,
    permissions: uniqueList(requirements.permissions ?? []),
    scopes: uniqueList(requirements.scopes ?? []),
    admin_scopes: uniqueList(requirements.admin_scopes ?? []),
    event_types: uniqueList(requirements.event_types ?? []),
    signal_types: uniqueList(requirements.signal_types ?? capability?.signal_types ?? []),
    required_security_env: uniqueList(requiredSecurity),
    optional_security_env: uniqueList(optionalSecurity),
    setup: requirements.setup,
    fallback: capability?.fallback,
    detail: capability?.detail,
  });
}

function platformReadinessChecks(platform, manifest = {}, env = {}) {
  const endpointCheck = endpointReadiness(manifest.endpoint);
  const checks = [{
    id: 'endpoint',
    ok: endpointCheck.ok,
    severity: endpointCheck.ok ? 'info' : 'error',
    detail: endpointCheck.reason,
  }];
  if (platform === 'google_meet') {
    const oidcConfigured = hasEnv(env, 'GOOGLE_PUBSUB_OIDC_AUDIENCE');
    const bearerConfigured = hasEnv(env, 'GOOGLE_PUBSUB_BEARER_TOKEN');
    checks.push({
      id: 'google_pubsub_auth',
      ok: oidcConfigured || bearerConfigured,
      severity: oidcConfigured ? 'info' : bearerConfigured ? 'warn' : 'error',
      detail: oidcConfigured
        ? 'oidc_configured'
        : bearerConfigured
          ? 'bearer_fallback_configured'
          : 'missing_google_pubsub_oidc_or_bearer',
      required_any_env: ['GOOGLE_PUBSUB_OIDC_AUDIENCE', 'GOOGLE_PUBSUB_BEARER_TOKEN'],
    });
    checks.push({
      id: 'google_pubsub_service_account_email',
      ok: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL'),
      severity: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL') ? 'info' : 'warn',
      detail: hasEnv(env, 'GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL')
        ? 'service_account_email_bound'
        : 'service_account_email_not_bound',
    });
  } else {
    const missing = missingEnv(env, manifest.required_security_env ?? []);
    checks.push({
      id: 'required_security_env',
      ok: missing.length === 0,
      severity: missing.length === 0 ? 'info' : 'error',
      detail: missing.length === 0 ? 'required_security_env_configured' : 'missing_required_security_env',
      missing_env: missing,
    });
  }
  return checks;
}

function readinessFromChecks(checks = []) {
  const blocking = checks.filter((item) => item.severity === 'error' && item.ok !== true);
  const warnings = checks.filter((item) => item.severity === 'warn' && item.ok !== true);
  return {
    ready: blocking.length === 0,
    warning_count: warnings.length,
    blocking_count: blocking.length,
    checks,
  };
}

export function buildGoogleMeetWorkspaceSubscriptionRequest(input = {}) {
  const targetResource = firstNonEmpty(input.targetResource, input.target_resource);
  const pubsubTopic = firstNonEmpty(input.pubsubTopic, input.pubsub_topic, input.topic);
  if (!targetResource) {
    throw new MeetingTimelineSdkError('targetResource is required for Google Meet subscription setup');
  }
  if (!pubsubTopic) {
    throw new MeetingTimelineSdkError('pubsubTopic is required for Google Meet subscription setup');
  }
  return compactObject({
    targetResource,
    eventTypes: input.eventTypes ?? input.event_types ?? GOOGLE_MEET_EVENT_TYPES,
    notificationEndpoint: {
      pubsubTopic,
    },
    payloadOptions: input.includeResource === false || input.include_resource === false
      ? undefined
      : { includeResource: true },
    ttl: input.ttl,
  });
}

export function buildMicrosoftTeamsMeetingCallSubscriptionRequest(input = {}) {
  const joinWebUrl = firstNonEmpty(input.joinWebUrl, input.join_web_url, input.meetingUrl, input.meeting_url);
  const notificationUrl = firstNonEmpty(input.notificationUrl, input.notification_url, input.webhookUrl, input.webhook_url);
  const clientState = firstNonEmpty(input.clientState, input.client_state);
  if (!joinWebUrl) {
    throw new MeetingTimelineSdkError('joinWebUrl is required for Microsoft Teams meetingCallEvents subscription setup');
  }
  if (!notificationUrl) {
    throw new MeetingTimelineSdkError('notificationUrl is required for Microsoft Teams subscription setup');
  }
  const now = input.now instanceof Date ? input.now : new Date(input.now ?? Date.now());
  return compactObject({
    changeType: input.changeType ?? input.change_type ?? MICROSOFT_TEAMS_CHANGE_TYPES.join(','),
    notificationUrl,
    resource: `/communications/onlineMeetings(joinWebUrl='${urlEncode(joinWebUrl)}')/meetingCallEvents`,
    expirationDateTime: input.expirationDateTime
      ?? input.expiration_date_time
      ?? addSeconds(now, Number(input.ttlSeconds ?? input.ttl_seconds ?? 2 * 24 * 60 * 60)),
    clientState,
    includeResourceData: input.includeResourceData ?? input.include_resource_data ?? false,
    encryptionCertificate: input.encryptionCertificate ?? input.encryption_certificate,
    encryptionCertificateId: input.encryptionCertificateId ?? input.encryption_certificate_id,
    lifecycleNotificationUrl: input.lifecycleNotificationUrl ?? input.lifecycle_notification_url ?? notificationUrl,
  });
}

export function buildZoomEventSubscriptionRequest(input = {}) {
  const webhookUrl = firstNonEmpty(input.eventWebhookUrl, input.event_webhook_url, input.webhookUrl, input.webhook_url);
  if (!webhookUrl) {
    throw new MeetingTimelineSdkError('webhookUrl is required for Zoom event subscription setup');
  }
  return compactObject({
    event_subscription_name: input.name ?? input.eventSubscriptionName ?? input.event_subscription_name ?? 'Meeting Timeline Events',
    event_webhook_url: webhookUrl,
    events: input.events ?? ZOOM_MEETING_EVENT_TYPES,
    subscription_scope: input.subscriptionScope ?? input.subscription_scope ?? 'account',
    account_id: input.accountId ?? input.account_id,
    user_ids: input.userIds ?? input.user_ids,
  });
}

export function buildWebexWebhookRequests(input = {}) {
  const targetUrl = firstNonEmpty(input.targetUrl, input.target_url, input.webhookUrl, input.webhook_url);
  if (!targetUrl) {
    throw new MeetingTimelineSdkError('targetUrl is required for Webex webhook setup');
  }
  const resources = input.resources ?? WEBEX_WEBHOOK_RESOURCES;
  const secret = firstNonEmpty(input.secret, input.webhookSecret, input.webhook_secret);
  const ownedBy = firstNonEmpty(input.ownedBy, input.owned_by);
  const status = input.status ?? 'active';
  return resources.flatMap((item) => (
    (item.events ?? []).map((event) => compactObject({
      name: input.name ? `${input.name} ${item.resource} ${event}` : `Meeting Timeline ${item.resource} ${event}`,
      targetUrl,
      resource: item.resource,
      event,
      filter: firstNonEmpty(item.filter, input.filter),
      secret,
      ownedBy,
      status,
    }))
  ));
}

export function buildMicrosoftGraphSubscriptionRenewalRequest(input = {}) {
  const subscriptionId = firstNonEmpty(input.subscriptionId, input.subscription_id, input.id);
  const now = input.now instanceof Date ? input.now : new Date(input.now ?? Date.now());
  const expirationDateTime = firstNonEmpty(
    input.expirationDateTime,
    input.expiration_date_time,
    input.expiresAt,
    input.expires_at,
  ) ?? addSeconds(now, Number(input.ttlSeconds ?? input.ttl_seconds ?? 2 * 24 * 60 * 60));
  return compactObject({
    method: 'PATCH',
    path: subscriptionId ? `/subscriptions/${subscriptionId}` : '/subscriptions/{subscription-id}',
    body: {
      expirationDateTime,
    },
  });
}

export function buildGoogleWorkspaceSubscriptionRenewalRequest(input = {}) {
  const subscriptionName = firstNonEmpty(input.subscriptionName, input.subscription_name, input.name);
  const ttl = firstNonEmpty(input.ttl, input.ttlSeconds != null ? `${input.ttlSeconds}s` : null, input.ttl_seconds != null ? `${input.ttl_seconds}s` : null);
  const expireTime = firstNonEmpty(input.expireTime, input.expire_time, input.expiresAt, input.expires_at);
  const updateMask = ttl ? 'ttl' : expireTime ? 'expire_time' : 'ttl';
  return compactObject({
    method: 'PATCH',
    path: subscriptionName ? `/v1beta/${subscriptionName}` : '/v1beta/{subscription-name}',
    query: {
      updateMask,
    },
    body: ttl ? { ttl } : expireTime ? { expireTime } : { ttl: '86400s' },
  });
}

export function platformEventEndpoint(baseUrl, platform) {
  const key = normalizePlatform(platform);
  const path = {
    local_detector: '/api/platform-events/local-detector',
    lark: '/api/platform-events/lark',
    google_meet: '/api/platform-events/google-meet',
    microsoft_teams: '/api/platform-events/teams',
    zoom: '/api/platform-events/zoom',
    webex: '/api/platform-events/webex',
  }[key];
  return absoluteEndpoint(baseUrl, path);
}

export function platformCapabilityContract(platform, options = {}) {
  const key = normalizePlatform(platform);
  const contract = platformCapabilityContracts[key];
  return compactObject({
    ...contract,
    sdk_modules: {
      ...contract.sdk_modules,
      timeline_bridge: '@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge',
      signal_reconciler: '@ai-annotation/meeting-timeline-sdk/adapters/signal-reconciler',
      session_discovery: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-session-discovery',
      active_speaker: '@ai-annotation/meeting-timeline-sdk/adapters/active-speaker',
      browser_meeting: '@ai-annotation/meeting-timeline-sdk/adapters/browser-meeting',
      native_meeting: '@ai-annotation/meeting-timeline-sdk/adapters/native-meeting',
      meeting_apps: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps',
      meeting_app_capture: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture',
      meeting_app_monitor: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-monitor',
      meeting_app_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime',
      meeting_app_browser_runtime: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime',
      meeting_app_content_script: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script',
      meeting_app_extension: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension',
      meeting_app_snapshot_recorder: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder',
      meeting_app_fixtures: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures',
      meeting_app_gate: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate',
      meeting_source: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-source',
      acceptance: '@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance',
      artifact_plan: '@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan',
      artifact_fetch: '@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch',
      onboarding: '@ai-annotation/meeting-timeline-sdk/adapters/platform-onboarding',
      webhook_router: '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router',
      platform_http: '@ai-annotation/meeting-timeline-sdk/adapters/platform-http',
      platform_node: '@ai-annotation/meeting-timeline-sdk/adapters/platform-node',
      platform_capture: '@ai-annotation/meeting-timeline-sdk/adapters/platform-capture',
      platform_gate: '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate',
      fixtures: '@ai-annotation/meeting-timeline-sdk/adapters/platform-fixtures',
      platform_kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
    },
    endpoints: options.baseUrl ? {
      platform_events: platformEventEndpoint(options.baseUrl, key),
      transcript_import: absoluteEndpoint(options.baseUrl, '/api/import/transcript'),
      status: `${platformEventEndpoint(options.baseUrl, key)}/status`,
    } : undefined,
  });
}

export function allPlatformCapabilityContracts(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => platformCapabilityContract(platform, options));
}

export function platformSetupManifest(platform, options = {}) {
  const key = normalizePlatform(platform);
  const endpoint = options.baseUrl ? platformEventEndpoint(options.baseUrl, key) : null;
  const statusEndpoint = options.baseUrl
    ? `${endpoint}/status`
    : null;
  const manifests = {
    local_detector: {
      platform: 'local_detector',
      display_name: 'Local Meeting Detector',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Local SDK call or local HTTP request',
      capabilities: platformCapabilityContract('local_detector', options),
      default_event_types: LOCAL_DETECTOR_EVENT_TYPES,
      required_security_env: [],
      optional_security_env: [],
      required_setup: [
        'Run the detector in a trusted host context such as a desktop observer, browser extension, e-ink companion app, or explicit manual controller.',
        'When the detector knows a real meeting is active, send meeting_started with detected_platform, meeting_id or meeting_url, and start_time_ms when available.',
        'Send meeting_ended when the local observer confirms the meeting is closed.',
        'Send speaker_started or active_speaker when the observer detects an active speaker change; send speaker_ended only when the observer has a reliable speech-end signal.',
        'Use captured_at_ms on annotations; do not wait for provider transcript or webhook events before creating the realtime axis.',
        'Let provider-specific webhooks backfill or reconcile metadata when they arrive later.',
      ],
      builders: [],
    },
    lark: {
      platform: 'lark',
      display_name: 'Feishu / Lark',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Feishu/Lark long connection or HTTP event callback',
      capabilities: platformCapabilityContract('lark', options),
      default_event_types: LARK_MEETING_EVENT_TYPES,
      required_permissions: [
        'vc:meeting.all_meeting:readonly',
        'vc:meeting.search:read for current-user scan fallback',
        'minutes:minutes.search:read / minutes:minutes.basic:read / minutes:minutes.transcript:export for post-meeting minutes',
      ],
      required_security_env: [],
      optional_security_env: ['LARK_APP_ID', 'LARK_APP_SECRET', 'LARK_ENCRYPT_KEY', 'LARK_VERIFICATION_TOKEN'],
      required_setup: [
        'Create a Feishu/Lark app and configure meeting event subscriptions.',
        'Prefer long-connection event delivery for local development; use HTTP callback only when a public HTTPS callback is available.',
        'Subscribe to all_meeting_started/all_meeting_ended as the direct meeting axis source.',
        'Subscribe to join_meeting/leave_meeting when participant track is needed.',
        'Use current-user meeting search or a local detector as a fallback when event delivery is delayed.',
        'Import Minutes transcript after the meeting rather than blocking realtime annotation on transcript availability.',
      ],
      builders: [],
    },
    google_meet: {
      platform: 'google_meet',
      display_name: 'Google Meet',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Google Workspace Events API -> Google Cloud Pub/Sub push',
      capabilities: platformCapabilityContract('google_meet', options),
      default_event_types: GOOGLE_MEET_EVENT_TYPES,
      lifecycle_event_types: GOOGLE_WORKSPACE_SUBSCRIPTION_LIFECYCLE_EVENT_TYPES,
      required_scopes: [
        'https://www.googleapis.com/auth/meetings.space.readonly',
        'https://www.googleapis.com/auth/drive.meet.readonly',
      ],
      required_security_env: ['GOOGLE_PUBSUB_OIDC_AUDIENCE'],
      optional_security_env: ['GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL', 'GOOGLE_PUBSUB_BEARER_TOKEN'],
      required_setup: [
        'Enable Google Meet REST API, Google Workspace Events API, and Pub/Sub.',
        'Create a Pub/Sub topic and push subscription targeting the endpoint.',
        'Configure Pub/Sub authenticated push OIDC audience to match GOOGLE_PUBSUB_OIDC_AUDIENCE.',
        'Create a Workspace Events subscription for a Meet space or user target resource.',
        'Monitor subscription lifecycle events for suspension, expiration reminders, and expiration.',
      ],
      builders: ['buildGoogleMeetWorkspaceSubscriptionRequest'],
    },
    microsoft_teams: {
      platform: 'microsoft_teams',
      display_name: 'Microsoft Teams',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Microsoft Graph change notifications',
      capabilities: platformCapabilityContract('microsoft_teams', options),
      default_change_types: MICROSOFT_TEAMS_CHANGE_TYPES,
      lifecycle_events: MICROSOFT_GRAPH_LIFECYCLE_EVENTS,
      resource_template: "/communications/onlineMeetings(joinWebUrl='{encodedJoinWebUrl}')/meetingCallEvents",
      required_permissions: ['OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All'],
      required_security_env: ['MICROSOFT_GRAPH_CLIENT_STATE'],
      required_setup: [
        'Create a Graph application permission grant for OnlineMeetings.Read.All or OnlineMeetings.ReadWrite.All.',
        'Create a change notification subscription for the meeting joinWebUrl.',
        'Set lifecycleNotificationUrl to the same endpoint or another endpoint that forwards lifecycle notifications here.',
        'Renew the subscription before its maximum 3 day expiration.',
        'Use rich notifications with includeResourceData for active meeting call changes when available.',
      ],
      builders: ['buildMicrosoftTeamsMeetingCallSubscriptionRequest'],
    },
    zoom: {
      platform: 'zoom',
      display_name: 'Zoom',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Zoom Meeting webhooks',
      capabilities: platformCapabilityContract('zoom', options),
      default_event_types: ZOOM_MEETING_EVENT_TYPES,
      required_security_env: ['ZOOM_WEBHOOK_SECRET_TOKEN'],
      required_scopes: [
        'meeting:read:meeting or meeting:read:meeting:admin',
        'meeting:read:participant or meeting:read:participant:admin',
        'cloud_recording:read:recording or cloud_recording:read:recording:admin',
      ],
      required_setup: [
        'Enable Event Subscriptions in the Zoom app or create an event subscription by API.',
        'Set the endpoint as the event notification URL and complete URL validation.',
        'Add scopes required by selected events before creating API-managed subscriptions.',
        'Return 2xx within 3 seconds and verify x-zm-signature on every request.',
      ],
      builders: ['buildZoomEventSubscriptionRequest'],
    },
    webex: {
      platform: 'webex',
      display_name: 'Cisco Webex',
      endpoint,
      status_endpoint: statusEndpoint,
      transport: 'Webex meeting related webhooks',
      capabilities: platformCapabilityContract('webex', options),
      webhook_resources: WEBEX_WEBHOOK_RESOURCES,
      required_security_env: ['WEBEX_WEBHOOK_SECRET'],
      required_scopes: [
        'meeting:schedules_read',
        'meeting:participants_read',
        'meeting:recordings_read',
        'meeting:transcripts_read',
      ],
      admin_scopes: [
        'meeting:admin_schedule_read',
        'meeting:admin_participants_read',
        'meeting:admin_recordings_read',
        'meeting:admin_transcripts_read',
      ],
      required_setup: [
        'Create one Webex webhook per meeting-related resource/event because firehose all does not include meetings started/ended or participant joined/left.',
        'Set targetUrl to the platform event endpoint and configure a shared secret.',
        'Grant the read scope for each selected resource before creating the webhook.',
        'Return 2xx quickly and verify X-Spark-Signature on every request.',
        'Treat transcripts and recordings as post-meeting artifact signals, then import VTT/text through /api/import/transcript.',
      ],
      builders: ['buildWebexWebhookRequests'],
    },
  };
  return compactObject(manifests[key]);
}

export function allPlatformSetupManifests(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => platformSetupManifest(platform, options));
}

export function buildPlatformPermissionPlan(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  const env = options.env ?? {};
  const features = selectedPermissionFeatures(key, options);
  const featurePlans = features.map((feature) => platformFeatureRequirement(key, manifest, feature));
  const requiredPermissions = uniqueList([
    ...(manifest.required_permissions ?? []),
    ...featurePlans.flatMap((item) => item.permissions ?? []),
  ]);
  const requiredScopes = uniqueList([
    ...(manifest.required_scopes ?? []),
    ...featurePlans.flatMap((item) => item.scopes ?? []),
  ]);
  const adminScopes = uniqueList([
    ...(manifest.admin_scopes ?? []),
    ...featurePlans.flatMap((item) => item.admin_scopes ?? []),
  ]);
  const requiredSecurityEnv = uniqueList([
    ...(manifest.required_security_env ?? []),
    ...featurePlans.flatMap((item) => item.required_security_env ?? []),
  ]);
  const optionalSecurityEnv = uniqueList([
    ...(manifest.optional_security_env ?? []),
    ...featurePlans.flatMap((item) => item.optional_security_env ?? []),
  ]);
  const readiness = evaluatePlatformSetupReadiness(key, options);
  return compactObject({
    platform: key,
    display_name: manifest.display_name,
    endpoint: manifest.endpoint,
    status_endpoint: manifest.status_endpoint,
    selected_features: features,
    feature_plans: featurePlans,
    required_permissions: requiredPermissions,
    required_scopes: requiredScopes,
    admin_scopes: adminScopes,
    required_security_env: requiredSecurityEnv,
    optional_security_env: optionalSecurityEnv,
    present_security_env: presentEnv(env, uniqueList([...requiredSecurityEnv, ...optionalSecurityEnv])),
    missing_security_env: missingEnv(env, requiredSecurityEnv),
    readiness,
    setup_steps: manifest.required_setup ?? [],
    builders: manifest.builders ?? [],
    notes: [
      'Use this plan before enabling provider webhooks so scope, event subscription, and signature verification requirements are reviewed together.',
      key === 'local_detector'
        ? 'Local detector has no remote provider permissions; its risk is timestamp quality and observer coverage.'
        : 'Provider webhooks should reconcile or backfill the low-latency local observer axis instead of blocking realtime annotation.',
    ],
  });
}

export function allPlatformPermissionPlans(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => buildPlatformPermissionPlan(platform, options));
}

function providerEventNames(manifest = {}) {
  if (Array.isArray(manifest.default_event_types)) return manifest.default_event_types;
  if (Array.isArray(manifest.default_change_types)) return manifest.default_change_types;
  if (Array.isArray(manifest.webhook_resources)) {
    return manifest.webhook_resources.flatMap((item) => (
      (item.events ?? []).map((event) => `${item.resource}.${event}`)
    ));
  }
  return [];
}

function preferredAxisStrategy(platform, options = {}) {
  const explicitMode = firstNonEmpty(options.axisMode, options.axis_mode);
  if (explicitMode) return String(explicitMode);
  return platform === 'local_detector' ? 'local_detector_primary' : 'hybrid_local_observer_first';
}

export function buildPlatformIntegrationPlan(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  const capabilities = manifest.capabilities;
  const sourcePriority = key === 'local_detector'
    ? ['local_detector']
    : ['local_observer', `${key}_provider_events`, 'post_meeting_transcript_import'];
  const transcript = capabilities.post_meeting_transcript ?? {};
  const maintenanceSubscription = options.subscription ?? options.subscriptions?.[key] ?? {};
  return compactObject({
    platform: key,
    display_name: manifest.display_name,
    recommended_mode: preferredAxisStrategy(key, options),
    source_priority: sourcePriority,
    modules: capabilities.sdk_modules,
    endpoints: capabilities.endpoints,
    realtime_axis: {
      strategy: preferredAxisStrategy(key, options),
      primary: key === 'local_detector' ? 'local_detector' : 'local_observer',
      reconcile_with_provider_events: key !== 'local_detector',
      signal_types: ['meeting_started', 'meeting_ended'],
      invariant: 'annotations_must_use_absolute_captured_at_ms_and_must_not_wait_for_transcript',
      local_observer: {
        module: capabilities.sdk_modules.local_observer,
        url_detection_module: capabilities.sdk_modules.url_detection,
        session_discovery_module: capabilities.sdk_modules.session_discovery,
        browser_meeting_module: capabilities.sdk_modules.browser_meeting,
        native_meeting_module: capabilities.sdk_modules.native_meeting,
        meeting_apps_module: capabilities.sdk_modules.meeting_apps,
        meeting_app_capture_module: capabilities.sdk_modules.meeting_app_capture,
        meeting_app_monitor_module: capabilities.sdk_modules.meeting_app_monitor,
        meeting_app_runtime_module: capabilities.sdk_modules.meeting_app_runtime,
        meeting_app_browser_runtime_module: capabilities.sdk_modules.meeting_app_browser_runtime,
        meeting_app_content_script_module: capabilities.sdk_modules.meeting_app_content_script,
        meeting_app_extension_module: capabilities.sdk_modules.meeting_app_extension,
        meeting_app_snapshot_recorder_module: capabilities.sdk_modules.meeting_app_snapshot_recorder,
        meeting_app_fixtures_module: capabilities.sdk_modules.meeting_app_fixtures,
        meeting_app_gate_module: capabilities.sdk_modules.meeting_app_gate,
        meeting_source_module: capabilities.sdk_modules.meeting_source,
        supported_platform_from_url: key !== 'local_detector',
      },
    },
    provider_events: {
      enabled: key !== 'local_detector',
      transport: manifest.transport,
      endpoint: manifest.endpoint,
      status_endpoint: manifest.status_endpoint,
      event_types: providerEventNames(manifest),
      lifecycle_event_types: manifest.lifecycle_event_types ?? manifest.lifecycle_events,
      adapter_module: capabilities.sdk_modules.events,
      ingest_module: capabilities.sdk_modules.ingest,
      webhook_handler_module: capabilities.sdk_modules.webhook_handler,
    },
    realtime_annotations: {
      strategy: 'insert_mark_with_absolute_capture_time',
      required_field: 'captured_at_ms',
      sdk_method: 'MeetingTimelineClient.insertMark',
      note: 'timeline_alignment_uses_meeting_axis_plus_absolute_annotation_time',
    },
    speaker_activity: {
      strategy: capabilities.speaker_activity.status?.startsWith('not_supported')
        ? 'local_detector_realtime_or_transcript_backfill'
        : 'provider_or_local_detector',
      active_speaker_module: capabilities.sdk_modules.active_speaker,
      ...capabilities.speaker_activity,
    },
    post_meeting_transcript: {
      strategy: transcript.status === 'not_applicable'
        ? 'provider_specific_or_generic_import_after_axis_exists'
        : 'import_after_meeting_ends',
      ...transcript,
      import_module: capabilities.sdk_modules.transcript,
    },
    setup: {
      required_permissions: manifest.required_permissions,
      required_scopes: manifest.required_scopes,
      admin_scopes: manifest.admin_scopes,
      required_security_env: manifest.required_security_env,
      optional_security_env: manifest.optional_security_env,
      required_steps: manifest.required_setup,
      builders: manifest.builders,
    },
    readiness: evaluatePlatformSetupReadiness(key, options),
    subscription_maintenance: evaluatePlatformSubscriptionMaintenance(key, maintenanceSubscription, options),
    limitations: capabilities.limitations,
  });
}

export function allPlatformIntegrationPlans(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => buildPlatformIntegrationPlan(platform, options));
}

export function buildPlatformSetup(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  if (key === 'google_meet') {
    return compactObject({
      ...manifest,
      workspace_subscription_request: options.googleMeetSubscription
        ? buildGoogleMeetWorkspaceSubscriptionRequest(options.googleMeetSubscription)
        : undefined,
    });
  }
  if (key === 'microsoft_teams') {
    return compactObject({
      ...manifest,
      graph_subscription_request: options.microsoftTeamsSubscription
        ? buildMicrosoftTeamsMeetingCallSubscriptionRequest(options.microsoftTeamsSubscription)
        : undefined,
    });
  }
  if (key === 'zoom') {
    return compactObject({
      ...manifest,
      zoom_event_subscription_request: options.zoomSubscription
        ? buildZoomEventSubscriptionRequest(options.zoomSubscription)
        : undefined,
    });
  }
  if (key === 'webex') {
    return compactObject({
      ...manifest,
      webex_webhook_requests: options.webexSubscription
        ? buildWebexWebhookRequests(options.webexSubscription)
        : undefined,
    });
  }
  return manifest;
}

export function evaluatePlatformSetupReadiness(platform, options = {}) {
  const key = normalizePlatform(platform);
  const manifest = platformSetupManifest(key, options);
  const env = options.env ?? {};
  return {
    platform: key,
    ...readinessFromChecks(platformReadinessChecks(key, manifest, env)),
  };
}

export function evaluateAllPlatformSetupReadiness(options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => (
    evaluatePlatformSetupReadiness(platform, options)
  ));
}

function subscriptionExpirationMs(subscription = {}) {
  return parseTimeMs(firstNonEmpty(
    subscription.expirationDateTime,
    subscription.expiration_date_time,
    subscription.expireTime,
    subscription.expire_time,
    subscription.expiresAt,
    subscription.expires_at,
    subscription.expiration,
  ));
}

export function evaluatePlatformSubscriptionMaintenance(platform, subscription = {}, options = {}) {
  const key = normalizePlatform(platform);
  const nowMs = parseTimeMs(options.now) ?? Date.now();
  const renewalWindowMs = Number(options.renewalWindowMs ?? options.renewal_window_ms ?? 12 * 60 * 60 * 1000);
  if (key === 'local_detector' || key === 'lark' || key === 'zoom' || key === 'webex') {
    return {
      platform: key,
      status: 'configured_manually',
      renewal_supported: false,
      renewal_due: false,
      expired: false,
      detail: key === 'local_detector'
        ? 'local_detector_has_no_remote_subscription_renewal'
        : key === 'lark'
        ? 'lark_long_connection_or_event_callback_does_not_use_short_cycle_subscription_renewal'
        : `${key}_webhook_subscriptions_do_not_require_short_cycle_renewal`,
    };
  }
  const expiresAtMs = subscriptionExpirationMs(subscription);
  if (expiresAtMs == null) {
    return {
      platform: key,
      status: 'unknown',
      renewal_supported: true,
      renewal_due: false,
      expired: false,
      detail: 'subscription_expiration_missing',
    };
  }
  const expiresInMs = expiresAtMs - nowMs;
  const expired = expiresInMs <= 0;
  const renewalDue = expired || expiresInMs <= renewalWindowMs;
  const renewAtMs = Math.max(nowMs, expiresAtMs - renewalWindowMs);
  const status = expired ? 'expired' : renewalDue ? 'renewal_due' : 'active';
  const renewalRequest = key === 'microsoft_teams'
    ? buildMicrosoftGraphSubscriptionRenewalRequest({
      subscriptionId: subscription.id ?? subscription.subscriptionId ?? subscription.subscription_id,
      now: new Date(nowMs),
      ttlSeconds: options.renewalTtlSeconds ?? options.renewal_ttl_seconds ?? 2 * 24 * 60 * 60,
    })
    : buildGoogleWorkspaceSubscriptionRenewalRequest({
      subscriptionName: subscription.name ?? subscription.subscriptionName ?? subscription.subscription_name,
      ttl: options.renewalTtl ?? options.renewal_ttl ?? '86400s',
    });
  return {
    platform: key,
    status,
    renewal_supported: true,
    renewal_due: renewalDue,
    expired,
    expires_at: new Date(expiresAtMs).toISOString(),
    expires_in_ms: expiresInMs,
    renew_at: new Date(renewAtMs).toISOString(),
    renewal_window_ms: renewalWindowMs,
    renewal_request: renewalRequest,
  };
}

export function evaluateAllPlatformSubscriptionMaintenance(subscriptions = {}, options = {}) {
  return MEETING_PLATFORM_KEYS.map((platform) => (
    evaluatePlatformSubscriptionMaintenance(platform, subscriptions[platform] ?? {}, options)
  ));
}

export const MEETING_PLATFORM_SETUP_BUILDERS = Object.freeze({
  MEETING_PLATFORM_KEYS,
  MEETING_PLATFORM_ALIASES,
  normalizeMeetingPlatform,
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildZoomEventSubscriptionRequest,
  buildWebexWebhookRequests,
  buildMicrosoftGraphSubscriptionRenewalRequest,
  buildGoogleWorkspaceSubscriptionRenewalRequest,
  platformSetupManifest,
  allPlatformSetupManifests,
  buildPlatformPermissionPlan,
  allPlatformPermissionPlans,
  buildPlatformSetup,
  evaluatePlatformSetupReadiness,
  evaluateAllPlatformSetupReadiness,
  evaluatePlatformSubscriptionMaintenance,
  evaluateAllPlatformSubscriptionMaintenance,
  platformCapabilityContract,
  allPlatformCapabilityContracts,
  buildPlatformIntegrationPlan,
  allPlatformIntegrationPlans,
});
