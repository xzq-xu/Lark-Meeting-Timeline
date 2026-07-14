import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { normalizeGoogleMeetEvent } from './google-meet.mjs';
import { normalizeLarkEvent } from './lark.mjs';
import { normalizeLocalDetectorEvent } from './local-detector.mjs';
import { normalizeMicrosoftTeamsEvent } from './microsoft-teams.mjs';
import {
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceReport,
} from './platform-adapter-contract.mjs';
import {
  buildMeetingPlatformProviderConnectionPack,
} from './platform-provider-connection.mjs';
import {
  buildMeetingPlatformRuntimeBundle,
} from './platform-runtime-bundle.mjs';
import {
  buildMeetingPlatformRuntimeEventPlan,
} from './platform-runtime-event.mjs';
import {
  buildMeetingPlatformAdapterRoute,
} from './platform-adapter-route.mjs';
import {
  buildMeetingPlatformAdapterSelection,
} from './platform-adapter-selection.mjs';
import {
  buildMeetingPlatformAdapterBlueprint,
} from './platform-adapter-blueprint.mjs';
import {
  normalizeMeetingPlatform,
  platformCapabilityContract,
} from './platform-setup.mjs';
import { normalizeWebexEvent } from './webex.mjs';
import { normalizeZoomEvent } from './zoom.mjs';

export const MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA = 'meeting_platform_registry_entry';
export const MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA = 'meeting_platform_registry_manifest';
export const MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA = 'meeting_platform_registry_acceptance';
export const MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION = 1;

const normalizers = Object.freeze({
  local_detector: normalizeLocalDetectorEvent,
  lark: normalizeLarkEvent,
  google_meet: normalizeGoogleMeetEvent,
  microsoft_teams: normalizeMicrosoftTeamsEvent,
  zoom: normalizeZoomEvent,
  webex: normalizeWebexEvent,
});

const sourceByPlatform = Object.freeze({
  local_detector: 'local_detector',
});

const DEFAULT_REGISTRY_PLATFORMS = Object.freeze([
  'local_detector',
  'lark',
  'google_meet',
  'microsoft_teams',
  'zoom',
  'webex',
]);

const REGISTRY_PLATFORM_ALIASES = Object.freeze({
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

function aliasesFor(platform) {
  return Object.entries(REGISTRY_PLATFORM_ALIASES)
    .filter(([, key]) => key === platform)
    .map(([alias]) => alias);
}

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

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function iso(ms) {
  return new Date(ms).toISOString();
}

function selectedPlatforms(options = {}) {
  return unique(asArray(firstNonEmpty(options.platforms, options.platform_keys, DEFAULT_REGISTRY_PLATFORMS))
    .map((platform) => normalizeMeetingPlatform(platform)));
}

export const MEETING_PLATFORM_EVENT_ADAPTERS = Object.freeze(
  DEFAULT_REGISTRY_PLATFORMS.map((platform) => Object.freeze({
    key: platform,
    aliases: Object.freeze(aliasesFor(platform)),
    source: sourceByPlatform[platform] ?? `${platform}_webhook`,
    normalize: normalizers[platform],
  })),
);

const adapterByAlias = new Map(
  MEETING_PLATFORM_EVENT_ADAPTERS.flatMap((adapter) => (
    adapter.aliases.map((alias) => [alias, adapter])
  )),
);

export function meetingPlatformEventAdapterFor(platform) {
  let normalized;
  try {
    normalized = normalizeMeetingPlatform(platform);
  } catch {
    return null;
  }
  return adapterByAlias.get(normalized) ?? MEETING_PLATFORM_EVENT_ADAPTERS.find((adapter) => adapter.key === normalized) ?? null;
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

function providerReplaySamples(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  if (key === 'local_detector') return [];
  const baseMs = replayBaseMs(options);
  if (key === 'google_meet') {
    const recordName = 'conferenceRecords/google-provider-record-001';
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
  if (key === 'microsoft_teams') {
    const meetingId = 'teams-provider-meeting-001';
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
  if (key === 'zoom') {
    const uuid = 'zoom-provider-uuid-001';
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
  if (key === 'webex') {
    const meetingId = 'webex-provider-meeting-001';
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
  if (key === 'lark') {
    const meeting = {
      id: 'lark-provider-meeting-001',
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
  return [];
}

function replayCoverage(signalTypes = []) {
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

function replayReportFromOptions(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const reportSources = [
    options.providerReplayReports,
    options.provider_replay_reports,
    options.providerReplayByPlatform,
    options.provider_replay_by_platform,
  ].filter(Boolean);
  for (const source of reportSources) {
    const report = source[key] ?? source[key.replaceAll('_', '-')] ?? (key === 'microsoft_teams' ? source.teams : undefined);
    if (report) return report;
  }
  const matrix = firstNonEmpty(options.providerReplayMatrix, options.provider_replay_matrix);
  const report = matrix?.reports?.find?.((item) => normalizeMeetingPlatform(item.platform) === key);
  if (report) return report;
  const row = matrix?.rows?.find?.((item) => normalizeMeetingPlatform(item.platform) === key);
  if (row) return row;
  return undefined;
}

function buildProviderReplayEvidence(platform, adapter, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const injected = replayReportFromOptions(key, options);
  if (injected) {
    return compactObject({
      schema: injected.schema,
      target: injected.target,
      accepted: injected.accepted === true,
      source: injected.source,
      record_count: injected.record_count,
      accepted_record_count: injected.accepted_record_count,
      runtime_event_count: injected.runtime_event_count,
      signal_count: injected.signal_count,
      signal_types: injected.signal_types,
      coverage: injected.coverage,
      required_coverage: injected.required_coverage,
      provider_events_block_realtime: injected.runtime_contract?.provider_events_block_realtime ?? injected.provider_events_block_realtime,
      transcript_blocks_realtime: injected.runtime_contract?.transcript_blocks_realtime ?? injected.transcript_blocks_realtime,
      issue_count: injected.issue_count,
      issues: injected.issues,
    });
  }
  if (key === 'local_detector') {
    return {
      accepted: true,
      not_applicable: true,
      source: adapter?.source,
      record_count: 0,
      accepted_record_count: 0,
      signal_count: 0,
      signal_types: [],
      coverage: replayCoverage([]),
      required_coverage: [],
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      issue_count: 0,
      issues: [],
    };
  }
  const samples = providerReplaySamples(key, options);
  const rows = samples.map((record, index) => {
    try {
      const signals = adapter?.normalize?.(record.payload, {
        ...(options.normalizerOptions ?? {}),
        ...(options.normalizer_options ?? {}),
        receivedAtMs: record.received_at_ms ?? replayBaseMs(options) + (index * 10_000),
      }) ?? [];
      const signalTypes = unique(signals.map((signal) => signal.type));
      return compactObject({
        id: record.id,
        kind: record.kind,
        required: record.required !== false,
        accepted: signalTypes.length > 0,
        signal_count: signals.length,
        signal_types: signalTypes,
        coverage: replayCoverage(signalTypes),
      });
    } catch (error) {
      return compactObject({
        id: record.id,
        kind: record.kind,
        required: record.required !== false,
        accepted: false,
        signal_count: 0,
        signal_types: [],
        coverage: replayCoverage([]),
        error: error.message ?? String(error),
      });
    }
  });
  const signalTypes = unique(rows.flatMap((row) => row.signal_types ?? []));
  const coverage = replayCoverage(signalTypes);
  const requiredCoverage = ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'];
  const issues = unique([
    ...(samples.length > 0 ? [] : ['missing_provider_replay_samples']),
    ...requiredCoverage.filter((keyName) => coverage[keyName] !== true).map((keyName) => `missing_coverage:${keyName}`),
    ...rows.flatMap((row) => row.accepted === true || row.required === false ? [] : [`${row.id}:not_accepted`]),
  ]);
  return compactObject({
    target: firstNonEmpty(options.target, 'registry'),
    accepted: issues.length === 0,
    source: adapter?.source,
    record_count: samples.length,
    accepted_record_count: rows.filter((row) => row.accepted === true).length,
    signal_count: rows.reduce((sum, row) => sum + (row.signal_count ?? 0), 0),
    signal_types: signalTypes,
    coverage,
    required_coverage: requiredCoverage,
    provider_events_block_realtime: false,
    transcript_blocks_realtime: false,
    issue_count: issues.length,
    issues,
  });
}

export function buildMeetingPlatformRegistryEntry(platform, options = {}) {
  const key = normalizeMeetingPlatform(platform);
  const adapter = meetingPlatformEventAdapterFor(key);
  const capabilities = platformCapabilityContract(key, options);
  const provider = buildMeetingPlatformProviderConnectionPack(key, options);
  const contract = buildMeetingPlatformAdapterContract(key, options);
  const contractAcceptance = buildMeetingPlatformAdapterContractAcceptanceReport(contract, options);
  const runtimeBundle = buildMeetingPlatformRuntimeBundle(key, options);
  const runtimeEventPlan = buildMeetingPlatformRuntimeEventPlan(key, options);
  const adapterRoute = buildMeetingPlatformAdapterRoute(key, options);
  const adapterSelection = buildMeetingPlatformAdapterSelection(key, options);
  const adapterBlueprint = buildMeetingPlatformAdapterBlueprint(key, options);
  const providerReplay = buildProviderReplayEvidence(key, adapter, options);
  const candidateObservation = runtimeBundle.messaging?.candidate_observation ?? runtimeBundle.browser?.candidate_observation;
  const candidateObservationReady = candidateObservation?.runtime_event_action === 'observe_platform_candidates'
    && runtimeEventPlan.supported_actions?.includes?.('observe_platform_candidates');
  const hostBasePath = firstNonEmpty(options.basePath, options.base_path, '/api/platform-events');

  return compactObject({
    type: 'meeting_platform_registry_entry',
    schema: MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    platform: key,
    display_name: capabilities.display_name ?? contract.display_name ?? runtimeBundle.display_name,
    aliases: adapter?.aliases ?? aliasesFor(key),
    event_adapter: {
      source: adapter?.source,
      normalize_available: typeof adapter?.normalize === 'function',
      module: capabilities.sdk_modules?.events,
    },
    supported_surfaces: contract.supported_surfaces,
    runtime: {
      bundle_schema: runtimeBundle.schema,
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      sdk_wiring_ready: runtimeBundle.readiness?.sdk_wiring_ready === true,
      browser_matches: runtimeBundle.browser?.matches ?? [],
      host_permissions: runtimeBundle.browser?.host_permissions ?? [],
      content_script_count: runtimeBundle.browser?.content_scripts?.length ?? 0,
      preset: runtimeBundle.runtime?.preset,
      sample_interval_ms: runtimeBundle.runtime?.start_options?.sampleIntervalMs,
      mutation_debounce_ms: runtimeBundle.runtime?.mutation_observer?.debounce_ms,
      speaker_min_stable_ms: runtimeBundle.runtime?.speaker_filter?.min_stable_ms,
      candidate_observation: compactObject({
        ready: candidateObservationReady,
        message_type: candidateObservation?.message_type,
        required_permission: candidateObservation?.required_permission,
        runtime_event_action: candidateObservation?.runtime_event_action,
        runtime_event_client_method: candidateObservation?.runtime_event_client_method,
        endpoint: '/api/meeting-platform/observe-candidates',
        producer: candidateObservation?.producer,
      }),
      runtime_event_plan_schema: runtimeEventPlan.schema,
      runtime_event_action_count: runtimeEventPlan.actions?.length ?? 0,
    },
    provider: {
      transport: provider.transport,
      endpoint: provider.endpoint,
      status_endpoint: provider.status_endpoint,
      role: provider.provider_role,
      required_for_realtime: false,
      security_verifier: provider.security?.verifier,
      missing_env: provider.security?.missing_env ?? [],
      ready: provider.readiness?.ready === true,
      start_events: contract.provider_observer?.events?.start_events ?? [],
      end_events: contract.provider_observer?.events?.end_events ?? [],
      participant_events: contract.provider_observer?.events?.participant_events ?? [],
      artifact_events: contract.provider_observer?.events?.artifact_events ?? [],
      lifecycle_events: contract.provider_observer?.events?.lifecycle_events ?? [],
    },
    provider_replay: providerReplay,
    annotations: {
      insert_endpoint: contract.annotations?.endpoints?.insertMark ?? runtimeBundle.host?.endpoints?.insertMark,
      runtime_event_endpoint: runtimeEventPlan.endpoint,
      runtime_event_plan: {
        schema: runtimeEventPlan.schema,
        client_factory: runtimeEventPlan.client_factory,
        supported_actions: runtimeEventPlan.supported_actions,
        action_count: runtimeEventPlan.actions?.length ?? 0,
        realtime_contract: runtimeEventPlan.realtime_contract,
      },
      candidate_observation_runtime_action: 'observe_platform_candidates',
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: contract.timebase?.provider_events_block_realtime === true,
      transcript_blocks_realtime: contract.timebase?.transcript_blocks_realtime === true,
      local_observer_may_start_axis: contract.timebase?.local_observer_may_start_axis === true,
    },
    adapter_route: {
      schema: adapterRoute.schema,
      recommended_mode: adapterRoute.recommended_mode,
      route_order: adapterRoute.route_order,
      route_count: adapterRoute.route_count,
      first_route: adapterRoute.routes?.[0]?.route,
      browser_match_count: adapterRoute.entrypoints?.browser_extension?.matches?.length ?? 0,
      native_detector: adapterRoute.entrypoints?.native_detector,
      provider_transport: adapterRoute.entrypoints?.provider_webhook?.transport,
      production_gate: adapterRoute.gates?.production,
      provider_events_block_realtime: adapterRoute.realtime_invariants?.provider_events_block_realtime,
      transcript_blocks_realtime: adapterRoute.realtime_invariants?.transcript_blocks_realtime,
    },
    adapter_selection: {
      schema: adapterSelection.schema,
      recommended_mode: adapterSelection.recommended_mode,
      axis_source: adapterSelection.selection?.axis_source,
      axis_surface: adapterSelection.selection?.axis_surface,
      timestamp_field: adapterSelection.selection?.timestamp_field,
      provider_reconcile_source: adapterSelection.selection?.provider_reconcile_source,
      provider_reconcile_required_for_production: adapterSelection.selection?.provider_reconcile_required_for_production,
      speaker_track_source: adapterSelection.selection?.speaker_track_source,
      post_meeting_artifact_source: adapterSelection.selection?.post_meeting_artifact_source,
      selection_ready: adapterSelection.readiness?.selection_ready === true,
      provider_events_block_realtime: adapterSelection.runtime_policy?.provider_events_block_realtime,
      transcript_blocks_realtime: adapterSelection.runtime_policy?.transcript_blocks_realtime,
      startup_order: adapterSelection.runtime_policy?.startup_order,
    },
    adapter_blueprint: {
      schema: adapterBlueprint.schema,
      ready: adapterBlueprint.readiness?.ready === true,
      recommended_mode: adapterBlueprint.recommended_mode,
      primary_surface: adapterBlueprint.primary_surface,
      surface_order: adapterBlueprint.surface_order,
      browser_recommended: adapterBlueprint.surfaces?.browser_extension?.recommended,
      native_recommended: adapterBlueprint.surfaces?.native_detector?.recommended,
      provider_blocks_realtime: adapterBlueprint.surfaces?.provider_reconcile?.blocks_realtime,
      transcript_blocks_realtime: adapterBlueprint.runtime_contract?.transcript_blocks_realtime,
      realtime_axis_timestamp_field: adapterBlueprint.realtime_axis_contract?.timestamp_field,
      first_acceptance_gate: adapterBlueprint.acceptance_gates?.realtime_pilot?.[0],
    },
    transcript: {
      availability: contract.transcript?.availability,
      source: contract.transcript?.source,
      import_endpoint: contract.transcript?.import_endpoint,
      realtime_dependency: contract.transcript?.realtime_dependency === true,
      blocks_realtime_annotation: runtimeBundle.transcript?.blocks_realtime_annotation === true,
    },
    host: {
      base_url: firstNonEmpty(options.baseUrl, options.base_url),
      endpoints: {
        platform_events: hostBasePath,
        platform_event: `${hostBasePath}/${key.replaceAll('_', '-')}`,
        annotations: '/api/annotations',
        adapter_routes: '/api/meeting-platform/adapter-routes',
        adapter_selections: '/api/meeting-platform/adapter-selections',
        adapter_blueprints: '/api/meeting-platform/adapter-blueprints',
        runtime_bundles: '/api/meeting-platform/runtime-bundles',
        runtime_event_plans: '/api/meeting-platform/runtime-event-plans',
        runtime_events: '/api/meeting-platform/runtime-events',
        platform_candidate_observation: '/api/meeting-platform/observe-candidates',
        readiness: '/api/meeting-platform/readiness',
        handoff: '/api/meeting-platform/handoff',
        adapter_contracts: '/api/meeting-platform/contracts',
      },
    },
    sdk: {
      package: '@ai-annotation/meeting-timeline-sdk',
      imports: {
        registry: '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry',
        kit: '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit',
        runtime_bundle: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle',
        runtime_event: '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event',
        adapter_route: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route',
        adapter_selection: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-selection',
        adapter_blueprint: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint',
        adapter_contract: '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract',
        provider_connection: '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection',
        events: capabilities.sdk_modules?.events,
      },
    },
    readiness: {
      contract_accepted: contractAcceptance.accepted === true,
      contract_issue_count: contractAcceptance.issue_count ?? 0,
      provider_required_for_realtime: runtimeBundle.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: runtimeBundle.readiness?.transcript_blocks_realtime === true,
      runtime_ready: runtimeBundle.readiness?.runtime_ready === true,
      candidate_observation_ready: candidateObservationReady,
      missing_items: unique([
        ...(contract.readiness?.missing_items ?? []),
        ...(runtimeBundle.readiness?.missing_items ?? []),
        ...(candidateObservationReady ? [] : ['candidate_observation_not_ready']),
      ]),
    },
    commands: {
      print_registry: `npm run meeting-platform:registry -- --platforms=${key}`,
      print_runtime_bundle: `npm run meeting-platform:runtime-bundle -- --platforms=${key}`,
      print_adapter_route: `npm run meeting-platform:adapter-route -- --platforms=${key}`,
      print_adapter_selection: `npm run meeting-platform:adapter-selection -- --platforms=${key}`,
      print_adapter_blueprint: `npm run meeting-platform:adapter-blueprint -- --platforms=${key}`,
      print_runtime_event_plan: `npm run meeting-platform:runtime-event-plan -- --platforms=${key}`,
      print_adaptation_package: `npm run meeting-platform:adaptation-package -- --platforms=${key}`,
      verify_contract: `npm run meeting-platform:adapter-contract -- --platforms=${key} --fail-on-rejected=true`,
      verify_handoff: `npm run meeting-platform:handoff-readiness -- --platforms=${key}`,
    },
    next_actions: unique([
      ...(runtimeBundle.next_actions ?? []),
      ...(contract.next_actions ?? []),
      'choose_platform_from_registry_manifest',
      'read_adapter_blueprint_before_wiring_external_host',
      'read_adapter_selection_before_wiring_external_host',
      'wire_host_runtime_bundles_endpoint_before_building_extension',
      'export_adapter_route_before_wiring_external_host',
      'export_adapter_selection_before_wiring_external_host',
      'export_runtime_event_plan_before_wiring_external_host',
      'verify_provider_replay_before_platform_release',
      'verify_candidate_observation_before_host_handoff',
    ]),
  });
}

export function buildMeetingPlatformRegistryManifest(options = {}) {
  const entries = selectedPlatforms(options).map((platform) => buildMeetingPlatformRegistryEntry(platform, {
    ...options,
    platforms: undefined,
    platform_keys: undefined,
  }));
  return {
    type: 'meeting_platform_registry_manifest',
    schema: MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    platform_count: entries.length,
    normalizer_count: entries.filter((entry) => entry.event_adapter?.normalize_available).length,
    runtime_ready_count: entries.filter((entry) => entry.readiness?.runtime_ready).length,
    contract_accepted_count: entries.filter((entry) => entry.readiness?.contract_accepted).length,
    candidate_observer_count: entries.filter((entry) => entry.readiness?.candidate_observation_ready).length,
    adapter_selection_ready_count: entries.filter((entry) => entry.adapter_selection?.selection_ready === true).length,
    adapter_blueprint_ready_count: entries.filter((entry) => entry.adapter_blueprint?.ready === true).length,
    provider_replay_accepted_count: entries.filter((entry) => entry.provider_replay?.accepted === true).length,
    provider_replay_record_count: entries.reduce((sum, entry) => sum + (entry.provider_replay?.record_count ?? 0), 0),
    provider_replay_signal_count: entries.reduce((sum, entry) => sum + (entry.provider_replay?.signal_count ?? 0), 0),
    provider_replay_blocking_count: entries.filter((entry) => entry.provider_replay?.provider_events_block_realtime === true).length,
    provider_required_for_realtime_count: entries.filter((entry) => entry.readiness?.provider_required_for_realtime).length,
    transcript_blocking_count: entries.filter((entry) => entry.readiness?.transcript_blocks_realtime).length,
    platforms: entries.map((entry) => entry.platform),
    rows: entries.map((entry) => ({
      platform: entry.platform,
      display_name: entry.display_name,
      aliases: entry.aliases,
      normalize_available: entry.event_adapter?.normalize_available === true,
      runtime_ready: entry.readiness?.runtime_ready === true,
      contract_accepted: entry.readiness?.contract_accepted === true,
      browser_match_count: entry.runtime?.browser_matches?.length ?? 0,
      candidate_observation_ready: entry.readiness?.candidate_observation_ready === true,
      candidate_observer_message_type: entry.runtime?.candidate_observation?.message_type,
      candidate_observer_permission: entry.runtime?.candidate_observation?.required_permission,
      candidate_observer_endpoint: entry.runtime?.candidate_observation?.endpoint,
      runtime_event_action_count: entry.runtime?.runtime_event_action_count ?? entry.annotations?.runtime_event_plan?.action_count ?? 0,
      adapter_route_mode: entry.adapter_route?.recommended_mode,
      adapter_first_route: entry.adapter_route?.first_route,
      adapter_route_count: entry.adapter_route?.route_count ?? 0,
      adapter_selection_ready: entry.adapter_selection?.selection_ready === true,
      adapter_selection_axis_source: entry.adapter_selection?.axis_source,
      adapter_selection_axis_surface: entry.adapter_selection?.axis_surface,
      adapter_selection_timestamp_field: entry.adapter_selection?.timestamp_field,
      adapter_selection_provider_reconcile_source: entry.adapter_selection?.provider_reconcile_source,
      adapter_selection_provider_blocks_realtime: entry.adapter_selection?.provider_events_block_realtime,
      adapter_selection_transcript_blocks_realtime: entry.adapter_selection?.transcript_blocks_realtime,
      adapter_blueprint_ready: entry.adapter_blueprint?.ready === true,
      adapter_blueprint_primary_surface: entry.adapter_blueprint?.primary_surface,
      adapter_blueprint_surface_order: entry.adapter_blueprint?.surface_order,
      adapter_blueprint_provider_blocks_realtime: entry.adapter_blueprint?.provider_blocks_realtime,
      adapter_blueprint_transcript_blocks_realtime: entry.adapter_blueprint?.transcript_blocks_realtime,
      adapter_blueprint_first_acceptance_gate: entry.adapter_blueprint?.first_acceptance_gate,
      provider_replay_accepted: entry.provider_replay?.accepted === true,
      provider_replay_signal_types: entry.provider_replay?.signal_types ?? [],
      provider_replay_required_coverage: entry.provider_replay?.required_coverage ?? [],
      provider_replay_coverage: entry.provider_replay?.coverage,
      provider_replay_provider_blocks_realtime: entry.provider_replay?.provider_events_block_realtime,
      provider_replay_issue_count: entry.provider_replay?.issue_count ?? 0,
      provider_transport: entry.provider?.transport,
      provider_ready: entry.provider?.ready === true,
      provider_required_for_realtime: entry.readiness?.provider_required_for_realtime === true,
      transcript_blocks_realtime: entry.readiness?.transcript_blocks_realtime === true,
      insert_endpoint: entry.annotations?.insert_endpoint,
      timestamp_field: entry.annotations?.timestamp_field,
    })),
    entries,
    next_actions: unique(entries.flatMap((entry) => entry.next_actions ?? [])),
  };
}

function registryManifestFrom(input = {}, options = {}) {
  if (input?.schema === MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA) return input;
  if (input?.schema === MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA) {
    return {
      type: 'meeting_platform_registry_manifest',
      schema: MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA,
      schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
      platform_count: 1,
      normalizer_count: input.event_adapter?.normalize_available === true ? 1 : 0,
      runtime_ready_count: input.readiness?.runtime_ready === true ? 1 : 0,
      contract_accepted_count: input.readiness?.contract_accepted === true ? 1 : 0,
      candidate_observer_count: input.readiness?.candidate_observation_ready === true ? 1 : 0,
      adapter_selection_ready_count: input.adapter_selection?.selection_ready === true ? 1 : 0,
      adapter_blueprint_ready_count: input.adapter_blueprint?.ready === true ? 1 : 0,
      provider_replay_accepted_count: input.provider_replay?.accepted === true ? 1 : 0,
      provider_replay_record_count: input.provider_replay?.record_count ?? 0,
      provider_replay_signal_count: input.provider_replay?.signal_count ?? 0,
      provider_replay_blocking_count: input.provider_replay?.provider_events_block_realtime === true ? 1 : 0,
      provider_required_for_realtime_count: input.readiness?.provider_required_for_realtime === true ? 1 : 0,
      transcript_blocking_count: input.readiness?.transcript_blocks_realtime === true ? 1 : 0,
      platforms: [input.platform],
      rows: [{
        platform: input.platform,
        display_name: input.display_name,
        aliases: input.aliases,
        normalize_available: input.event_adapter?.normalize_available === true,
        runtime_ready: input.readiness?.runtime_ready === true,
        contract_accepted: input.readiness?.contract_accepted === true,
        browser_match_count: input.runtime?.browser_matches?.length ?? 0,
        candidate_observation_ready: input.readiness?.candidate_observation_ready === true,
        candidate_observer_message_type: input.runtime?.candidate_observation?.message_type,
        candidate_observer_permission: input.runtime?.candidate_observation?.required_permission,
        candidate_observer_endpoint: input.runtime?.candidate_observation?.endpoint,
        runtime_event_action_count: input.runtime?.runtime_event_action_count ?? input.annotations?.runtime_event_plan?.action_count ?? 0,
        adapter_route_mode: input.adapter_route?.recommended_mode,
        adapter_first_route: input.adapter_route?.first_route,
        adapter_route_count: input.adapter_route?.route_count ?? 0,
        adapter_selection_ready: input.adapter_selection?.selection_ready === true,
        adapter_selection_axis_source: input.adapter_selection?.axis_source,
        adapter_selection_axis_surface: input.adapter_selection?.axis_surface,
        adapter_selection_timestamp_field: input.adapter_selection?.timestamp_field,
        adapter_selection_provider_reconcile_source: input.adapter_selection?.provider_reconcile_source,
        adapter_selection_provider_blocks_realtime: input.adapter_selection?.provider_events_block_realtime,
        adapter_selection_transcript_blocks_realtime: input.adapter_selection?.transcript_blocks_realtime,
        adapter_blueprint_ready: input.adapter_blueprint?.ready === true,
        adapter_blueprint_primary_surface: input.adapter_blueprint?.primary_surface,
        adapter_blueprint_surface_order: input.adapter_blueprint?.surface_order,
        adapter_blueprint_provider_blocks_realtime: input.adapter_blueprint?.provider_blocks_realtime,
        adapter_blueprint_transcript_blocks_realtime: input.adapter_blueprint?.transcript_blocks_realtime,
        adapter_blueprint_first_acceptance_gate: input.adapter_blueprint?.first_acceptance_gate,
        provider_replay_accepted: input.provider_replay?.accepted === true,
        provider_replay_signal_types: input.provider_replay?.signal_types ?? [],
        provider_replay_required_coverage: input.provider_replay?.required_coverage ?? [],
        provider_replay_coverage: input.provider_replay?.coverage,
        provider_replay_provider_blocks_realtime: input.provider_replay?.provider_events_block_realtime,
        provider_replay_issue_count: input.provider_replay?.issue_count ?? 0,
        provider_transport: input.provider?.transport,
        provider_ready: input.provider?.ready === true,
        provider_required_for_realtime: input.readiness?.provider_required_for_realtime === true,
        transcript_blocks_realtime: input.readiness?.transcript_blocks_realtime === true,
        insert_endpoint: input.annotations?.insert_endpoint,
        timestamp_field: input.annotations?.timestamp_field,
      }],
      entries: [input],
      next_actions: input.next_actions ?? [],
    };
  }
  return buildMeetingPlatformRegistryManifest({ ...input, ...options });
}

function registryAcceptanceIssues(manifest = {}, options = {}) {
  const issues = [];
  const entries = asArray(manifest.entries);
  const requireProviderReady = options.requireProviderReady === true || options.require_provider_ready === true;
  if (manifest.schema !== MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA) {
    issues.push(issue('error', 'invalid_registry_schema', 'Registry manifest schema must be meeting_platform_registry_manifest.', {
      actual_schema: manifest.schema,
    }));
  }
  if ((manifest.platform_count ?? 0) <= 0) {
    issues.push(issue('error', 'empty_registry_manifest', 'Registry manifest must include at least one platform.'));
  }
  if (manifest.normalizer_count !== manifest.platform_count) {
    issues.push(issue('error', 'missing_platform_normalizer', 'Every selected platform must have a normalizer.', {
      normalizer_count: manifest.normalizer_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.runtime_ready_count !== manifest.platform_count) {
    issues.push(issue('error', 'runtime_not_ready', 'Every selected platform must have a runtime-ready bundle.', {
      runtime_ready_count: manifest.runtime_ready_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.contract_accepted_count !== manifest.platform_count) {
    issues.push(issue('error', 'adapter_contract_not_accepted', 'Every selected platform must pass adapter contract acceptance.', {
      contract_accepted_count: manifest.contract_accepted_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.candidate_observer_count !== manifest.platform_count) {
    issues.push(issue('error', 'candidate_observer_not_ready', 'Every selected platform must expose candidate observation for host-level axis binding.', {
      candidate_observer_count: manifest.candidate_observer_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.adapter_selection_ready_count !== manifest.platform_count) {
    issues.push(issue('error', 'adapter_selection_not_ready', 'Every selected platform must expose a ready adapter selection for host wiring.', {
      adapter_selection_ready_count: manifest.adapter_selection_ready_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.adapter_blueprint_ready_count !== manifest.platform_count) {
    issues.push(issue('error', 'adapter_blueprint_not_ready', 'Every selected platform must expose a ready adapter blueprint for host wiring.', {
      adapter_blueprint_ready_count: manifest.adapter_blueprint_ready_count,
      platform_count: manifest.platform_count,
    }));
  }
  if (manifest.provider_replay_accepted_count !== manifest.platform_count) {
    issues.push(issue('error', 'provider_replay_not_accepted', 'Every selected platform must expose accepted provider replay evidence.', {
      provider_replay_accepted_count: manifest.provider_replay_accepted_count,
      platform_count: manifest.platform_count,
    }));
  }
  if ((manifest.provider_replay_blocking_count ?? 0) > 0) {
    issues.push(issue('error', 'provider_replay_blocks_realtime', 'Provider replay evidence must keep provider events non-blocking for realtime annotations.', {
      provider_replay_blocking_count: manifest.provider_replay_blocking_count,
    }));
  }
  if ((manifest.provider_required_for_realtime_count ?? 0) > 0) {
    issues.push(issue('error', 'provider_blocks_realtime', 'Provider events must not be required for realtime annotation insertion.', {
      provider_required_for_realtime_count: manifest.provider_required_for_realtime_count,
    }));
  }
  if ((manifest.transcript_blocking_count ?? 0) > 0) {
    issues.push(issue('error', 'transcript_blocks_realtime', 'Transcript import must not block realtime annotation insertion.', {
      transcript_blocking_count: manifest.transcript_blocking_count,
    }));
  }
  for (const entry of entries) {
    const platform = entry.platform;
    if (entry.event_adapter?.normalize_available !== true) {
      issues.push(issue('error', 'entry_missing_normalizer', 'Registry entry must expose a normalizer.', { platform }));
    }
    if (entry.readiness?.runtime_ready !== true) {
      issues.push(issue('error', 'entry_runtime_not_ready', 'Registry entry must have runtime_ready=true.', { platform }));
    }
    if (entry.readiness?.contract_accepted !== true) {
      issues.push(issue('error', 'entry_contract_not_accepted', 'Registry entry must have contract_accepted=true.', { platform }));
    }
    if (entry.annotations?.timestamp_field !== 'captured_at_ms') {
      issues.push(issue('error', 'entry_invalid_timestamp_field', 'Registry entry must use captured_at_ms for realtime annotations.', {
        platform,
        timestamp_field: entry.annotations?.timestamp_field,
      }));
    }
    if (!entry.annotations?.insert_endpoint) {
      issues.push(issue('error', 'entry_missing_insert_endpoint', 'Registry entry must include an annotation insert endpoint.', { platform }));
    }
    if (!entry.annotations?.runtime_event_endpoint) {
      issues.push(issue('error', 'entry_missing_runtime_event_endpoint', 'Registry entry must include the runtime event endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.platform_candidate_observation) {
      issues.push(issue('error', 'entry_missing_candidate_observation_endpoint', 'Registry entry must include the host candidate observation endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.runtime_event_plans) {
      issues.push(issue('error', 'entry_missing_runtime_event_plan_endpoint', 'Registry entry must include the host runtime event plan endpoint.', { platform }));
    }
    if (!entry.sdk?.imports?.runtime_event) {
      issues.push(issue('error', 'entry_missing_runtime_event_import', 'Registry entry must include the runtime event SDK import path.', { platform }));
    }
    if (!entry.annotations?.runtime_event_plan?.supported_actions?.includes?.('insert_annotation')) {
      issues.push(issue('error', 'entry_missing_runtime_insert_action', 'Registry entry runtime event plan must include insert_annotation.', { platform }));
    }
    if (!entry.annotations?.runtime_event_plan?.supported_actions?.includes?.('observe_platform_candidates')) {
      issues.push(issue('error', 'entry_missing_candidate_observation_action', 'Registry entry runtime event plan must include observe_platform_candidates.', { platform }));
    }
    if (entry.readiness?.candidate_observation_ready !== true) {
      issues.push(issue('error', 'entry_candidate_observation_not_ready', 'Registry entry must expose candidate observation for host-level axis binding.', { platform }));
    }
    if (entry.runtime?.candidate_observation?.message_type !== 'meeting_timeline.observe_candidates') {
      issues.push(issue('error', 'entry_invalid_candidate_observation_message_type', 'Registry entry must use meeting_timeline.observe_candidates for candidate observation.', {
        platform,
        message_type: entry.runtime?.candidate_observation?.message_type,
      }));
    }
    if (entry.runtime?.candidate_observation?.required_permission !== 'tabs') {
      issues.push(issue('error', 'entry_invalid_candidate_observation_permission', 'Registry entry must declare tabs permission for browser candidate observation.', {
        platform,
        required_permission: entry.runtime?.candidate_observation?.required_permission,
      }));
    }
    if (entry.annotations?.provider_events_block_realtime === true) {
      issues.push(issue('error', 'entry_provider_blocks_realtime', 'Registry entry must keep provider events non-blocking for realtime annotations.', { platform }));
    }
    if (entry.annotations?.transcript_blocks_realtime === true || entry.transcript?.blocks_realtime_annotation === true) {
      issues.push(issue('error', 'entry_transcript_blocks_realtime', 'Registry entry must keep transcript import non-blocking for realtime annotations.', { platform }));
    }
    if (entry.supported_surfaces?.browser_observer === true && (entry.runtime?.browser_matches?.length ?? 0) === 0) {
      issues.push(issue('error', 'entry_missing_browser_matches', 'Browser-observed platforms must include runtime browser match patterns.', { platform }));
    }
    if (!entry.host?.endpoints?.runtime_bundles) {
      issues.push(issue('error', 'entry_missing_runtime_bundle_endpoint', 'Registry entry must include the host runtime bundle endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.adapter_routes) {
      issues.push(issue('error', 'entry_missing_adapter_route_endpoint', 'Registry entry must include the host adapter route endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.adapter_selections) {
      issues.push(issue('error', 'entry_missing_adapter_selection_endpoint', 'Registry entry must include the host adapter selection endpoint.', { platform }));
    }
    if (!entry.host?.endpoints?.adapter_blueprints) {
      issues.push(issue('error', 'entry_missing_adapter_blueprint_endpoint', 'Registry entry must include the host adapter blueprint endpoint.', { platform }));
    }
    if (!entry.sdk?.imports?.runtime_bundle) {
      issues.push(issue('error', 'entry_missing_runtime_bundle_import', 'Registry entry must include the runtime bundle SDK import path.', { platform }));
    }
    if (!entry.sdk?.imports?.adapter_route) {
      issues.push(issue('error', 'entry_missing_adapter_route_import', 'Registry entry must include the adapter route SDK import path.', { platform }));
    }
    if (!entry.sdk?.imports?.adapter_selection) {
      issues.push(issue('error', 'entry_missing_adapter_selection_import', 'Registry entry must include the adapter selection SDK import path.', { platform }));
    }
    if (!entry.sdk?.imports?.adapter_blueprint) {
      issues.push(issue('error', 'entry_missing_adapter_blueprint_import', 'Registry entry must include the adapter blueprint SDK import path.', { platform }));
    }
    if (!entry.adapter_route?.recommended_mode || !entry.adapter_route?.first_route) {
      issues.push(issue('error', 'entry_missing_adapter_route', 'Registry entry must include adapter route planning for external host wiring.', { platform }));
    }
    if (entry.adapter_route?.provider_events_block_realtime === true || entry.adapter_route?.transcript_blocks_realtime === true) {
      issues.push(issue('error', 'entry_adapter_route_blocks_realtime', 'Adapter route must keep provider events and transcript non-blocking for realtime annotations.', { platform }));
    }
    if (entry.adapter_selection?.selection_ready !== true || !entry.adapter_selection?.axis_source) {
      issues.push(issue('error', 'entry_missing_adapter_selection', 'Registry entry must include a ready adapter selection with an axis source.', { platform }));
    }
    if (entry.adapter_selection?.timestamp_field !== 'captured_at_ms') {
      issues.push(issue('error', 'entry_adapter_selection_invalid_timestamp_field', 'Adapter selection must keep realtime annotation timestamps on captured_at_ms.', {
        platform,
        timestamp_field: entry.adapter_selection?.timestamp_field,
      }));
    }
    if (entry.adapter_selection?.provider_events_block_realtime === true || entry.adapter_selection?.transcript_blocks_realtime === true) {
      issues.push(issue('error', 'entry_adapter_selection_blocks_realtime', 'Adapter selection must keep provider events and transcript non-blocking for realtime annotations.', { platform }));
    }
    if (entry.adapter_blueprint?.ready !== true || !entry.adapter_blueprint?.primary_surface) {
      issues.push(issue('error', 'entry_missing_adapter_blueprint', 'Registry entry must include a ready adapter blueprint with a primary surface.', { platform }));
    }
    if (entry.adapter_blueprint?.provider_blocks_realtime === true || entry.adapter_blueprint?.transcript_blocks_realtime === true) {
      issues.push(issue('error', 'entry_adapter_blueprint_blocks_realtime', 'Adapter blueprint must keep provider events and transcript non-blocking for realtime annotations.', { platform }));
    }
    if (entry.adapter_blueprint?.realtime_axis_timestamp_field !== 'captured_at_ms') {
      issues.push(issue('error', 'entry_adapter_blueprint_invalid_timestamp_field', 'Adapter blueprint must keep realtime axis timestamps on captured_at_ms.', {
        platform,
        timestamp_field: entry.adapter_blueprint?.realtime_axis_timestamp_field,
      }));
    }
    if (entry.provider_replay?.accepted !== true) {
      issues.push(issue('error', 'entry_provider_replay_not_accepted', 'Registry entry must include accepted provider replay evidence.', {
        platform,
        issues: entry.provider_replay?.issues ?? [],
      }));
    }
    if (entry.provider_replay?.provider_events_block_realtime === true) {
      issues.push(issue('error', 'entry_provider_replay_blocks_realtime', 'Provider replay must keep provider events non-blocking for realtime annotations.', { platform }));
    }
    if (platform !== 'local_detector') {
      const requiredReplayCoverage = entry.provider_replay?.required_coverage ?? ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'];
      const missingReplayCoverage = requiredReplayCoverage.filter((key) => entry.provider_replay?.coverage?.[key] !== true);
      if (missingReplayCoverage.length > 0) {
        issues.push(issue('error', 'entry_provider_replay_missing_coverage', 'Provider replay must cover required platform lifecycle and handoff signals.', {
          platform,
          missing_coverage: missingReplayCoverage,
        }));
      }
    }
    if (requireProviderReady && entry.provider?.ready !== true) {
      issues.push(issue('error', 'entry_provider_not_ready', 'Provider setup must be ready when requireProviderReady=true.', {
        platform,
        missing_env: entry.provider?.missing_env ?? [],
      }));
    }
  }
  return issues;
}

export function buildMeetingPlatformRegistryAcceptanceReport(manifestOrOptions = {}, options = {}) {
  const manifest = registryManifestFrom(manifestOrOptions, options);
  const issues = registryAcceptanceIssues(manifest, options);
  const blocking = issues.filter((item) => item.severity === 'error');
  return {
    type: 'meeting_platform_registry_acceptance',
    schema: MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA,
    schema_version: MEETING_PLATFORM_REGISTRY_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    platform_count: manifest.platform_count ?? 0,
    normalizer_count: manifest.normalizer_count ?? 0,
    runtime_ready_count: manifest.runtime_ready_count ?? 0,
    contract_accepted_count: manifest.contract_accepted_count ?? 0,
    candidate_observer_count: manifest.candidate_observer_count ?? 0,
    adapter_blueprint_ready_count: manifest.adapter_blueprint_ready_count ?? 0,
    adapter_selection_ready_count: manifest.adapter_selection_ready_count ?? 0,
    provider_replay_accepted_count: manifest.provider_replay_accepted_count ?? 0,
    provider_replay_record_count: manifest.provider_replay_record_count ?? 0,
    provider_replay_signal_count: manifest.provider_replay_signal_count ?? 0,
    provider_replay_blocking_count: manifest.provider_replay_blocking_count ?? 0,
    provider_required_for_realtime_count: manifest.provider_required_for_realtime_count ?? 0,
    transcript_blocking_count: manifest.transcript_blocking_count ?? 0,
    blocking_count: blocking.length,
    warning_count: issues.length - blocking.length,
    issues,
    manifest,
    next_actions: unique([
      ...blocking.map((item) => item.code),
      ...(manifest.next_actions ?? []),
    ]),
  };
}

export function assertMeetingPlatformRegistryManifest(manifestOrOptions = {}, options = {}) {
  const report = buildMeetingPlatformRegistryAcceptanceReport(manifestOrOptions, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting platform registry manifest failed acceptance', {
      issues: report.issues,
      report,
    });
  }
  return report;
}
