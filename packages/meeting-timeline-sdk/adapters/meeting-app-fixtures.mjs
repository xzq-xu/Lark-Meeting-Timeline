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
  const byPlatform = {
    google_meet: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://meet.google.com/abc-defg-hij'),
      title: firstNonEmpty(options.title, 'SDK fixture - Google Meet'),
      buttons: controls('Turn off microphone', 'Present now', 'Leave call'),
      participants: [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} is speaking`, {
          audioLevel: 0.78,
          dataset: { participantId: ids.speakerId },
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
    microsoft_teams: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_fixture%40thread.v2/0'),
      title: firstNonEmpty(options.title, 'SDK fixture | Microsoft Teams'),
      buttons: controls('Leave', 'Show conversation', 'Raise hand', 'Share content'),
      participants: [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} speaking`, {
          audioLevel: 0.66,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName} muted`),
      ],
    },
    zoom: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://zoom.us/j/987654321'),
      title: firstNonEmpty(options.title, 'SDK fixture - Zoom Meeting'),
      buttons: controls('Mute Audio', 'Start Video', 'Participants', 'Leave Meeting'),
      participants: [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} is speaking`, {
          audioLevel: 0.73,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
    lark: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://vc.feishu.cn/j/123456789'),
      title: firstNonEmpty(options.title, 'SDK fixture - 飞书会议'),
      buttons: controls('挂断', 'AI 视图', '共享屏幕', '关闭麦克风'),
      participants: [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName} 正在发言`, {
          audioLevel: 0.71,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName} 已静音`),
      ],
    },
    webex: {
      url: firstNonEmpty(options.url, options.meeting_url, 'https://example.webex.com/meet/sdk-fixture'),
      title: firstNonEmpty(options.title, 'SDK fixture - Webex'),
      buttons: controls('Unmute', 'Chat', 'Participants', 'Leave meeting'),
      participants: [
        participant(ids.speakerId, ids.speakerName, `${ids.speakerName}, active speaker`, {
          audioLevel: 0.69,
        }),
        participant(ids.mutedId, ids.mutedName, `${ids.mutedName}, muted`),
      ],
    },
  };
  return byPlatform[platform];
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
      audible: true,
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

export function buildMeetingAppFixtureAcceptanceReport(options = {}) {
  const platforms = options.platforms ?? options.platform_keys ?? MEETING_APP_FIXTURE_PLATFORMS;
  const reports = platforms.map((platform) => diagnoseMeetingAppFixture(platform, options));
  const required = [
    'platform_detected',
    'meeting_id',
    'in_meeting',
    'active_speaker',
    'meeting_started',
    'speaker_started',
  ];
  const missing = reports.flatMap((report) => required
    .filter((key) => report.coverage[key] !== true)
    .map((key) => `${report.platform}:${key}`));
  return {
    type: 'meeting_app_fixture_acceptance_report',
    accepted: missing.length === 0,
    required,
    missing,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => required.every((key) => report.coverage[key] === true)).length,
    coverage_by_platform: Object.fromEntries(reports.map((report) => [report.platform, report.coverage])),
    reports,
  };
}
