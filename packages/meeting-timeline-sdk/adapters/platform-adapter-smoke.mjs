import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from './platform-adapter-export-package.mjs';
import {
  buildMeetingPlatformAdapterImportPlanMatrix,
} from './platform-adapter-import-plan.mjs';
import {
  MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA,
  buildMeetingPlatformAdapterInstallManifest,
} from './platform-adapter-install-manifest.mjs';
import {
  createMeetingPlatformAdapterMessageBridge,
} from './platform-adapter-message-bridge.mjs';

export const MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA = 'meeting_platform_adapter_smoke_report';
export const MEETING_PLATFORM_ADAPTER_SMOKE_ROW_SCHEMA = 'meeting_platform_adapter_smoke_row';
export const MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA_VERSION = 1;

const DEFAULT_SMOKE_PLATFORMS = Object.freeze([
  'google-meet',
  'teams',
  'zoom',
  'webex',
  'lark',
]);

const FIXTURE_URLS = Object.freeze({
  google_meet: 'https://meet.google.com/abc-defg-hij',
  microsoft_teams: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
  zoom: 'https://zoom.us/j/987654321',
  webex: 'https://example.webex.com/meet/sample',
  lark: 'https://vc.feishu.cn/j/123456789',
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

function normalizedPlatform(value = '') {
  const key = String(value || '').trim().toLowerCase().replace(/[-\s]+/g, '_');
  const aliases = {
    google: 'google_meet',
    meet: 'google_meet',
    google_meet: 'google_meet',
    microsoft: 'microsoft_teams',
    teams: 'microsoft_teams',
    microsoft_teams: 'microsoft_teams',
    zoom: 'zoom',
    webex: 'webex',
    cisco_webex: 'webex',
    lark: 'lark',
    feishu: 'lark',
    larksuite: 'lark',
  };
  return aliases[key] ?? key;
}

function selectedPlatforms(manifest = {}, options = {}) {
  const rows = asArray(manifest.platform_registry);
  return unique(asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    manifest.platforms,
    rows.map((row) => row.platform),
    DEFAULT_SMOKE_PLATFORMS,
  )).map((platform) => normalizedPlatform(platform)));
}

function isInstallManifest(value = {}) {
  return value?.schema === MEETING_PLATFORM_ADAPTER_INSTALL_MANIFEST_SCHEMA;
}

function buildDefaultInstallManifest(options = {}) {
  const platforms = asArray(firstNonEmpty(
    options.platforms,
    options.platform_keys,
    options.platformKeys,
    DEFAULT_SMOKE_PLATFORMS,
  ));
  const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
    platforms,
  }, {
    ...options,
    target: firstNonEmpty(options.target, 'static'),
    includeArtifacts: true,
  });
  const availableFiles = exportMatrix.packages.flatMap((pkg) => asArray(pkg.host_files).map((file) => file.path));
  const importMatrix = buildMeetingPlatformAdapterImportPlanMatrix(exportMatrix.packages, {
    ...options,
    availableFiles: firstNonEmpty(options.availableFiles, options.available_files, availableFiles),
  });
  return buildMeetingPlatformAdapterInstallManifest(importMatrix.plans, {
    ...options,
    target: firstNonEmpty(options.target, 'static'),
  });
}

function installManifestFrom(manifestOrOptions = {}, options = {}) {
  if (isInstallManifest(manifestOrOptions)) return manifestOrOptions;
  const explicit = firstNonEmpty(
    options.installManifest,
    options.install_manifest,
    options.manifest,
    manifestOrOptions.installManifest,
    manifestOrOptions.install_manifest,
    manifestOrOptions.manifest,
  );
  if (isInstallManifest(explicit)) return explicit;
  return buildDefaultInstallManifest({
    ...manifestOrOptions,
    ...options,
  });
}

function fixtureUrl(platform, options = {}) {
  const fixtures = {
    ...(options.fixtureUrls ?? {}),
    ...(options.fixture_urls ?? {}),
  };
  return firstNonEmpty(
    fixtures[platform],
    fixtures[normalizedPlatform(platform)],
    FIXTURE_URLS[normalizedPlatform(platform)],
  );
}

function fixtureTitle(platform) {
  const labels = {
    google_meet: 'Google Meet fixture',
    microsoft_teams: 'Microsoft Teams fixture',
    zoom: 'Zoom fixture',
    webex: 'Webex fixture',
    lark: 'Lark fixture',
  };
  return labels[normalizedPlatform(platform)] ?? `${platform} fixture`;
}

function smokeCapturedAtMs(index = 0, options = {}) {
  const base = Number(firstNonEmpty(
    options.captured_at_ms,
    options.capturedAtMs,
    options.baseCapturedAtMs,
    options.base_captured_at_ms,
    1_782_800_000_000,
  ));
  return base + (index * 1000);
}

function summarizeCall(call = {}) {
  return compactObject({
    method: call.method,
    platform: call.platform,
    captured_at_ms: call.payload?.captured_at_ms,
    label: call.payload?.label,
    speaker_id: call.payload?.speaker_id,
    participant_id: call.payload?.participant_id,
    event_type: call.payload?.event_type,
    current_url: call.payload?.current_url,
    candidate_count: call.payload?.candidates?.length,
  });
}

function createRecordingAdapterClient(calls = []) {
  const push = (method, platform, payload, options) => {
    const call = compactObject({ method, platform, payload, options });
    calls.push(call);
    return {
      ok: true,
      method,
      platform,
      captured_at_ms: payload?.captured_at_ms,
    };
  };
  return {
    async observePlatformCandidates(payload = {}, options = {}) {
      return push('observePlatformCandidates', payload.platform, payload, options);
    },
    async insertAnnotation(platform, payload, options = {}) {
      return push('insertAnnotation', platform, payload ?? {}, options);
    },
    async speakerTrack(platform, payload, options = {}) {
      return push('speakerTrack', platform, payload ?? {}, options);
    },
    async participantTrack(platform, payload, options = {}) {
      return push('participantTrack', platform, payload ?? {}, options);
    },
    async ingestProvider(platform, payload, options = {}) {
      return push('ingestProvider', platform, payload ?? {}, options);
    },
  };
}

function callIndex(calls = [], method) {
  return calls.findIndex((call) => call.method === method);
}

function rowIssues(row = {}) {
  const issues = [];
  if (!row.fixture_url) issues.push('missing_fixture_url');
  if (row.observe_action !== 'open_session') issues.push('observe_candidates_did_not_open_session');
  if (row.insert_action !== 'insert_annotation') issues.push('insert_mark_did_not_insert_annotation');
  if (row.observed_platform !== row.platform) issues.push('observed_platform_mismatch');
  if (row.inserted_platform !== row.platform) issues.push('inserted_platform_mismatch');
  if (row.observe_before_insert !== true) issues.push('axis_not_observed_before_insert');
  if (row.captured_at_ms_preserved !== true) issues.push('captured_at_ms_not_preserved');
  if (row.speaker_track_inserted !== true) issues.push('speaker_track_not_routed');
  if (row.participant_track_inserted !== true) issues.push('participant_track_not_routed');
  if (row.provider_reconcile_nonblocking !== true) issues.push('provider_reconcile_not_routed');
  if (row.adapter_blueprint_available !== true) issues.push('adapter_blueprint_not_available');
  return issues;
}

async function runPlatformSmoke(manifest = {}, platform, index = 0, options = {}) {
  const normalized = normalizedPlatform(platform);
  const url = fixtureUrl(normalized, options);
  const capturedAtMs = smokeCapturedAtMs(index, options);
  const calls = [];
  const client = createRecordingAdapterClient(calls);
  const bridge = createMeetingPlatformAdapterMessageBridge(manifest, client, {
    ...options,
    clock: () => capturedAtMs,
    bridgeId: firstNonEmpty(options.bridgeId, options.bridge_id, `adapter-smoke-${normalized}`),
  });
  try {
    const observe = await bridge.handleMessage({
      type: 'meeting_timeline.observe_candidates',
      request_id: `observe-${normalized}`,
      payload: {
        captured_at_ms: capturedAtMs,
        tabs: [
          { url: 'https://example.com/not-a-meeting', title: 'Other tab', active: false },
          { url, title: fixtureTitle(normalized), active: true, in_meeting: true },
        ],
      },
    });
    const insert = await bridge.handleMessage({
      type: 'meeting_timeline.insert_mark',
      request_id: `mark-${normalized}`,
      payload: {
        mark: {
          label: 'SDK smoke mark',
          kind: 'question',
          captured_at_ms: capturedAtMs,
        },
      },
    });
    const speaker = await bridge.handleMessage({
      type: 'meeting_timeline.speaker_track',
      request_id: `speaker-${normalized}`,
      payload: {
        speaker: {
          speaker_id: `speaker-${normalized}`,
          label: 'active speaker',
          captured_at_ms: capturedAtMs,
        },
      },
    });
    const participant = await bridge.handleMessage({
      type: 'meeting_timeline.participant_track',
      request_id: `participant-${normalized}`,
      payload: {
        participant: {
          participant_id: `participant-${normalized}`,
          status: 'present',
          captured_at_ms: capturedAtMs,
        },
      },
    });
    const provider = await bridge.handleMessage({
      type: 'meeting_timeline.provider_event',
      request_id: `provider-${normalized}`,
      payload: {
        event: {
          event_type: 'meeting_started',
          captured_at_ms: capturedAtMs,
        },
      },
    });
    const observeIndex = callIndex(calls, 'observePlatformCandidates');
    const insertIndex = callIndex(calls, 'insertAnnotation');
    const speakerIndex = callIndex(calls, 'speakerTrack');
    const participantIndex = callIndex(calls, 'participantTrack');
    const providerIndex = callIndex(calls, 'ingestProvider');
    const insertedCall = calls[insertIndex] ?? {};
    const adapterBlueprint = observe.adapter_blueprint
      ?? observe.result?.payload?.launch_plan?.adapter_blueprint
      ?? bridge.getState().runner?.current_launch_plan?.adapter_blueprint;
    const row = compactObject({
      type: 'meeting_platform_adapter_smoke_row',
      schema: MEETING_PLATFORM_ADAPTER_SMOKE_ROW_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA_VERSION,
      platform: normalized,
      fixture_url: url,
      observed_platform: observe.platform,
      inserted_platform: insert.platform,
      selected_surface: observe.selected_surface,
      adapter_blueprint_available: adapterBlueprint?.available === true,
      adapter_blueprint_primary_surface: adapterBlueprint?.primary_surface,
      adapter_blueprint_first_gate: adapterBlueprint?.first_acceptance_gate,
      observe_action: observe.action,
      insert_action: insert.action,
      speaker_action: speaker.action,
      participant_action: participant.action,
      provider_action: provider.action,
      observe_before_insert: observeIndex >= 0 && insertIndex > observeIndex,
      captured_at_ms_expected: capturedAtMs,
      captured_at_ms_preserved: insertedCall.payload?.captured_at_ms === capturedAtMs,
      speaker_track_inserted: speakerIndex > insertIndex,
      participant_track_inserted: participantIndex > speakerIndex,
      provider_reconcile_nonblocking: providerIndex > participantIndex,
      bridge_state_opened: bridge.getState().runner?.opened === true,
      call_count: calls.length,
      calls: calls.map(summarizeCall),
    });
    const issues = rowIssues(row);
    return {
      ...row,
      accepted: issues.length === 0,
      issues,
    };
  } catch (error) {
    return {
      type: 'meeting_platform_adapter_smoke_row',
      schema: MEETING_PLATFORM_ADAPTER_SMOKE_ROW_SCHEMA,
      schema_version: MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA_VERSION,
      platform: normalized,
      fixture_url: url,
      accepted: false,
      call_count: calls.length,
      calls: calls.map(summarizeCall),
      issues: ['smoke_exception'],
      error: compactObject({
        name: error?.name,
        message: error?.message,
        details: error?.details,
      }),
    };
  }
}

function nextActions(report = {}) {
  if (report.accepted) {
    return [
      'wire_adapter_message_bridge_into_host_project',
      'run_this_smoke_in_host_ci_before_live_platform_rollout',
      'replace_fixture_urls_with_live_window_candidates_for_each_platform',
    ];
  }
  const issueCodes = unique(asArray(report.rows).flatMap((row) => row.issues ?? []));
  return unique([
    ...issueCodes.map((code) => `fix_${code}`),
    'rerun_platform_adapter_smoke_before_handoff',
  ]);
}

export async function runMeetingPlatformAdapterSmoke(manifestOrOptions = {}, options = {}) {
  const smokeOptions = {
    ...manifestOrOptions,
    ...options,
  };
  const manifest = installManifestFrom(manifestOrOptions, options);
  const platforms = selectedPlatforms(manifest, {
    ...smokeOptions,
  });
  const rows = [];
  for (const [index, platform] of platforms.entries()) {
    rows.push(await runPlatformSmoke(manifest, platform, index, smokeOptions));
  }
  const acceptedRows = rows.filter((row) => row.accepted === true);
  const report = {
    type: 'meeting_platform_adapter_smoke_report',
    schema: MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA,
    schema_version: MEETING_PLATFORM_ADAPTER_SMOKE_REPORT_SCHEMA_VERSION,
    accepted: rows.length > 0 && acceptedRows.length === rows.length && manifest.accepted === true,
    install_manifest_schema: manifest.schema,
    install_manifest_accepted: manifest.accepted === true,
    platform_count: rows.length,
    accepted_count: acceptedRows.length,
    failed_count: rows.length - acceptedRows.length,
    platforms,
    runtime_contract: {
      local_axis_first: true,
      timestamp_field: 'captured_at_ms',
      provider_events_block_realtime: false,
      transcript_blocks_realtime: false,
      transcript_required_for_smoke: false,
    },
    verified_sequence: [
      'meeting_timeline.observe_candidates',
      'observePlatformCandidates',
      'meeting_timeline.insert_mark',
      'insertAnnotation',
      'meeting_timeline.speaker_track',
      'meeting_timeline.participant_track',
      'meeting_timeline.provider_event',
    ],
    rows,
  };
  return {
    ...report,
    next_actions: nextActions(report),
  };
}

export async function assertMeetingPlatformAdapterSmoke(manifestOrOptions = {}, options = {}) {
  const report = await runMeetingPlatformAdapterSmoke(manifestOrOptions, options);
  if (report.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting platform adapter smoke failed', {
      code: 'meeting_platform_adapter_smoke_failed',
      accepted_count: report.accepted_count,
      failed_count: report.failed_count,
      issues: report.rows.flatMap((row) => row.issues ?? []),
      rows: report.rows,
    });
  }
  return report;
}
