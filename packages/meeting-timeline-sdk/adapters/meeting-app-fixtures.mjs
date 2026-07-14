import { MeetingTimelineSdkError, compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { normalizeMeetingAppSnapshot, observeMeetingAppSample } from './meeting-apps.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_APP_FIXTURE_PLATFORMS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);

const DEFAULT_OBSERVED_AT_MS = 1_783_356_000_000;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function normalizeAppPlatform(platform) {
  const normalized = normalizeMeetingPlatform(platform);
  if (!MEETING_APP_FIXTURE_PLATFORMS.includes(normalized)) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app fixture platform: ${String(platform || '(empty)')}`, {
      platform,
      supported_platforms: MEETING_APP_FIXTURE_PLATFORMS,
    });
  }
  return normalized;
}

function observedAtMs(options = {}) {
  return normalizeAbsoluteMs(firstNonEmpty(
    options.observedAtMs,
    options.observed_at_ms,
    options.startMs,
    options.start_ms,
    DEFAULT_OBSERVED_AT_MS,
  ), 'meeting_app_fixture_observed_at_ms');
}

function fixtureIds(platform, options = {}) {
  return {
    title: firstNonEmpty(options.title, 'SDK fixture meeting'),
    speakerId: firstNonEmpty(options.speakerId, options.speaker_id, `${platform}-speaker-001`),
    speakerName: firstNonEmpty(options.speakerName, options.speaker_name, 'Ada Lovelace'),
    mutedId: firstNonEmpty(options.mutedId, options.muted_id, `${platform}-muted-001`),
    mutedName: firstNonEmpty(options.mutedName, options.muted_name, 'Grace Hopper'),
  };
}

function fixtureState(options = {}) {
  const value = String(firstNonEmpty(options.state, options.fixtureState, options.fixture_state, 'active')).toLowerCase();
  if (['active', 'joined', 'in_meeting', 'in-meeting'].includes(value)) return 'active';
  if (['prejoin', 'pre_join', 'pre-join', 'ended', 'left', 'inactive'].includes(value)) return 'prejoin';
  throw new MeetingTimelineSdkError(`Unsupported meeting app fixture state: ${String(value || '(empty)')}`, {
    state: value,
    supported_states: ['active', 'prejoin'],
  });
}

function controls(...items) {
  return items.map((label) => ({ label, ariaLabel: label }));
}

function participant(id, name, label, extra = {}) {
  return compactObject({
    id,
    name,
    display_name: name,
    ariaLabel: label,
    label,
    ...extra,
  });
}

function fixtureConfig(platform, options = {}) {
  const ids = fixtureIds(platform, options);
  const state = fixtureState(options);
  const byPlatform = {
    google_meet: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://meet.google.com/abc-defg-hij'),
      title: firstNonEmpty(options.title, state === 'prejoin' ? 'Ready to join - Google Meet' : 'SDK fixture - Google Meet'),
      buttons: state === 'prejoin'
        ? controls('Join now', 'Check your audio and video')
        : controls('Turn off microphone', 'Present now', 'Leave call'),
      participants: state === 'prejoin' ? [] : [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} is speaking`, {
          audioLevel: 0.78,
          dataset: { participantId: ids.speakerId },
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
    microsoft_teams: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_fixture%40thread.v2/0'),
      title: firstNonEmpty(options.title, state === 'prejoin' ? 'Join now | Microsoft Teams' : 'SDK fixture | Microsoft Teams'),
      buttons: state === 'prejoin'
        ? controls('Join now', 'Choose your audio and video settings')
        : controls('Leave', 'Show conversation', 'Raise hand', 'Share content'),
      participants: state === 'prejoin' ? [] : [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} speaking`, {
          audioLevel: 0.66,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName} muted`),
      ],
    },
    zoom: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://zoom.us/j/987654321'),
      title: firstNonEmpty(options.title, state === 'prejoin' ? 'Join Meeting - Zoom' : 'SDK fixture - Zoom Meeting'),
      buttons: state === 'prejoin'
        ? controls('Join Meeting', 'Join with Computer Audio')
        : controls('Mute Audio', 'Start Video', 'Participants', 'Leave Meeting'),
      participants: state === 'prejoin' ? [] : [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} is speaking`, {
          audioLevel: 0.73,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
    lark: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://vc.feishu.cn/j/123456789'),
      title: firstNonEmpty(options.title, state === 'prejoin' ? '加入会议 - 飞书' : 'SDK fixture - 飞书会议'),
      buttons: state === 'prejoin'
        ? controls('加入会议', '立即加入')
        : controls('挂断', 'AI 视图', '共享屏幕', '关闭麦克风'),
      participants: state === 'prejoin' ? [] : [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} 正在发言`, {
          audioLevel: 0.71,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName} 已静音`),
      ],
    },
    webex: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://example.webex.com/meet/sdk-fixture'),
      title: firstNonEmpty(options.title, state === 'prejoin' ? 'Join meeting - Webex' : 'SDK fixture - Webex'),
      buttons: state === 'prejoin'
        ? controls('Join meeting', 'Start meeting')
        : controls('Unmute', 'Chat', 'Participants', 'Leave meeting'),
      participants: state === 'prejoin' ? [] : [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName}, active speaker`, {
          audioLevel: 0.69,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
  };
  return {
    ...byPlatform[platform],
    state,
  };
}

export function buildMeetingAppFixtureSnapshot(platform, options = {}) {
  const key = normalizeAppPlatform(platform);
  const config = fixtureConfig(key, options);
  const atMs = observedAtMs(options);
  const detected = detectMeetingFromUrl({ url: config.url, title: config.title }) ?? {};
  const page = compactObject({
    url: config.url,
    title: config.title,
    documentVisible: options.documentVisible ?? options.document_visible ?? true,
    buttons: config.buttons,
    controls: config.buttons,
    tiles: config.participants,
    participants: config.participants,
  });
  return compactObject({
    id: `meeting-app-fixture-${key}`,
    source: options.source ?? 'meeting_app_fixture',
    observedAtMs: atMs,
    platform: key,
    provider: key,
    fixture_state: config.state,
    url: config.url,
    meeting_url: config.url,
    meeting_id: firstNonEmpty(options.meetingId, options.meeting_id, detected.meeting_id),
    external_meeting_id: firstNonEmpty(options.externalMeetingId, options.external_meeting_id, detected.external_meeting_id),
    title: config.title,
    browser: {
      name: firstNonEmpty(options.browserName, options.browser_name, 'Chrome'),
    },
    tab: {
      active: true,
      audible: config.state === 'active' ? true : undefined,
      url: config.url,
      title: config.title,
    },
    page,
    dom: {
      url: config.url,
      title: config.title,
      controls: config.buttons,
      participants: config.participants,
    },
  });
}

export function buildAllMeetingAppFixtureSnapshots(options = {}) {
  return Object.fromEntries(MEETING_APP_FIXTURE_PLATFORMS.map((platform) => [
    platform,
    buildMeetingAppFixtureSnapshot(platform, options),
  ]));
}

export function diagnoseMeetingAppFixture(platform, options = {}) {
  const key = normalizeAppPlatform(platform);
  const snapshot = buildMeetingAppFixtureSnapshot(key, options);
  const normalized = normalizeMeetingAppSnapshot(snapshot, options);
  const observation = observeMeetingAppSample(null, snapshot, {
    source: options.source ?? 'meeting_app_fixture',
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
    observedAtMs: snapshot.observedAtMs,
    ...options,
  });
  const signalTypes = observation.signals.map((item) => item.type);
  return {
    platform: key,
    snapshot,
    normalized,
    signal_types: signalTypes,
    signals: observation.signals,
    coverage: {
      platform_detected: normalized?.platform === key,
      meeting_id: Boolean(normalized?.meeting_id),
      in_meeting: normalized?.inMeeting === true,
      active_speaker: Boolean(normalized?.activeSpeaker?.id || normalized?.activeSpeaker?.name),
      meeting_started: signalTypes.includes('meeting_started'),
      speaker_started: signalTypes.includes('speaker_started'),
    },
    observation,
  };
}

export function diagnoseMeetingAppFixtureLifecycle(platform, options = {}) {
  const key = normalizeAppPlatform(platform);
  const startMs = observedAtMs(options);
  const endMs = normalizeAbsoluteMs(firstNonEmpty(
    options.endObservedAtMs,
    options.end_observed_at_ms,
    startMs + Number(firstNonEmpty(options.endOffsetMs, options.end_offset_ms, 1_000)),
  ), 'meeting_app_fixture_end_observed_at_ms');
  const activeSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options,
    state: 'active',
    observedAtMs: startMs,
  });
  const endedSnapshot = buildMeetingAppFixtureSnapshot(key, {
    ...options,
    state: 'prejoin',
    observedAtMs: endMs,
  });
  const active = observeMeetingAppSample(null, activeSnapshot, {
    source: options.source ?? 'meeting_app_fixture',
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
    observedAtMs: activeSnapshot.observedAtMs,
    ...options,
  });
  const ended = observeMeetingAppSample(active.state, endedSnapshot, {
    source: options.source ?? 'meeting_app_fixture',
    speakerOptions: {
      minStableMs: 0,
      ...(options.speakerOptions ?? {}),
      ...(options.speaker_options ?? {}),
    },
    observedAtMs: endedSnapshot.observedAtMs,
    ...options,
  });
  const signals = [...active.signals, ...ended.signals];
  const signalTypes = signals.map((item) => item.type);
  const normalizedEnded = normalizeMeetingAppSnapshot(endedSnapshot, options);
  return {
    platform: key,
    snapshots: {
      active: activeSnapshot,
      ended: endedSnapshot,
    },
    normalized_ended: normalizedEnded,
    signal_types: signalTypes,
    signals,
    coverage: {
      meeting_started: signalTypes.includes('meeting_started'),
      speaker_started: signalTypes.includes('speaker_started'),
      meeting_ended: signalTypes.includes('meeting_ended'),
      ended_in_meeting_false: normalizedEnded?.inMeeting === false,
    },
    observations: {
      active,
      ended,
    },
  };
}

export function buildMeetingAppFixtureAcceptanceReport(options = {}) {
  const platforms = options.platforms ?? options.platform_keys ?? MEETING_APP_FIXTURE_PLATFORMS;
  const reports = platforms.map((platform) => diagnoseMeetingAppFixture(platform, options));
  const lifecycle_reports = platforms.map((platform) => diagnoseMeetingAppFixtureLifecycle(platform, options));
  const required = [
    'platform_detected',
    'meeting_id',
    'in_meeting',
    'active_speaker',
    'meeting_started',
    'speaker_started',
    'meeting_ended',
  ];
  const coverageByPlatform = Object.fromEntries(reports.map((report, index) => {
    const lifecycle = lifecycle_reports[index];
    return [report.platform, {
      ...report.coverage,
      meeting_ended: lifecycle.coverage.meeting_ended,
      ended_in_meeting_false: lifecycle.coverage.ended_in_meeting_false,
    }];
  }));
  const missing = Object.entries(coverageByPlatform).flatMap(([platform, coverage]) => required
    .filter((key) => coverage[key] !== true)
    .map((key) => `${platform}:${key}`));
  return {
    type: 'meeting_app_fixture_acceptance_report',
    accepted: missing.length === 0,
    required,
    missing,
    platform_count: reports.length,
    accepted_count: Object.values(coverageByPlatform).filter((coverage) => required.every((key) => coverage[key] === true)).length,
    coverage_by_platform: coverageByPlatform,
    reports,
    lifecycle_reports,
  };
}
