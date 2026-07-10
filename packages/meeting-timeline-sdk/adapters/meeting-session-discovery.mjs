import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
import { createLocalMeetingObserver, createLocalMeetingTimelineObserver, selectMeetingSnapshot } from './local-observer.mjs';
import { detectMeetingFromUrl } from './meeting-url.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

const PLATFORM_FINGERPRINTS = Object.freeze([
  {
    platform: 'google_meet',
    hosts: ['meet.google.com'],
    process: [/google chrome/i, /\bchrome\b/i, /microsoft edge/i, /\bedge\b/i, /firefox/i, /safari/i],
    bundle: [/com\.google\.chrome/i, /com\.microsoft\.edgemac/i, /org\.mozilla\.firefox/i, /com\.apple\.safari/i],
    title: [/google meet/i, /\bmeet\b.*\bgoogle\b/i],
    meetingTitle: [/google meet/i, /meet\.google\.com/i],
  },
  {
    platform: 'microsoft_teams',
    hosts: ['teams.microsoft.com', 'teams.live.com', 'teams.cloud.microsoft'],
    process: [/microsoft teams/i, /\bteams\b/i],
    bundle: [/com\.microsoft\.teams/i, /com\.microsoft\.teams2/i],
    title: [/microsoft teams/i, /\bteams\b/i],
    meetingTitle: [/meeting/i, /\bcall\b/i, /会议/i, /通话/i],
  },
  {
    platform: 'zoom',
    hosts: ['zoom.us', 'zoomgov.com'],
    process: [/zoom/i],
    bundle: [/us\.zoom/i],
    title: [/zoom/i],
    meetingTitle: [/zoom meeting/i, /zoom webinar/i, /\bmeeting\b/i, /会议/i],
  },
  {
    platform: 'lark',
    hosts: ['vc.feishu.cn', 'vc.larksuite.com', 'larksuite.com', 'feishu.cn'],
    process: [/lark/i, /feishu/i, /飞书/i],
    bundle: [/larksuite/i, /feishu/i, /bytedance\.ee\.lark/i],
    title: [/lark/i, /feishu/i, /飞书/i],
    meetingTitle: [/meeting/i, /会议/i, /通话/i, /视频/i],
  },
  {
    platform: 'webex',
    hosts: ['webex.com'],
    process: [/webex/i, /cisco/i],
    bundle: [/webex/i, /cisco/i],
    title: [/webex/i],
    meetingTitle: [/webex meeting/i, /personal room/i, /\bmeeting\b/i, /会议/i],
  },
]);

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function firstPath(raw, paths) {
  return firstNonEmpty(...paths.map((path) => getPath(raw, path)));
}

function firstBoolean(raw, paths) {
  for (const path of paths) {
    const value = getPath(raw, path);
    if (typeof value === 'boolean') return value;
  }
  return undefined;
}

function normalizeText(value) {
  return String(value ?? '').trim();
}

function cleanId(value) {
  return normalizeText(value)
    .replace(/\s+/g, '-')
    .replace(/[^a-zA-Z0-9._:@~-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase();
}

function safeParseUrl(value) {
  if (!value) return null;
  try {
    return new URL(String(value));
  } catch {
    return null;
  }
}

function hostMatches(hostname, hosts = []) {
  const host = String(hostname || '').toLowerCase();
  return hosts.some((suffix) => host === suffix || host.endsWith(`.${suffix}`));
}

function regexMatches(patterns = [], value) {
  const text = String(value || '');
  return patterns.some((pattern) => pattern.test(text));
}

function urlInput(input = {}) {
  return firstPath(input, [
    'meeting.meeting_url',
    'meeting.meetingUrl',
    'meeting.url',
    'meeting.join_url',
    'meeting.joinUrl',
    'meeting_url',
    'meetingUrl',
    'join_url',
    'joinUrl',
    'url',
    'href',
    'window.url',
    'browser.url',
    'tab.url',
    'application.url',
    'app.url',
  ]);
}

function titleInput(input = {}) {
  return firstPath(input, [
    'meeting.title',
    'meeting.topic',
    'meeting.name',
    'title',
    'topic',
    'name',
    'window.title',
    'tab.title',
    'browser.title',
    'application.title',
    'app.title',
    'process.title',
  ]);
}

function appNameInput(input = {}) {
  return firstPath(input, [
    'application.name',
    'app.name',
    'process.name',
    'processName',
    'process_name',
    'executable',
    'bundle.name',
    'window.app_name',
    'window.appName',
    'window.application',
    'window.ownerName',
    'ownerName',
    'browser.name',
  ]);
}

function processNameInput(input = {}) {
  return firstPath(input, [
    'process.name',
    'processName',
    'process_name',
    'executable',
    'application.processName',
    'application.process_name',
    'app.processName',
    'app.process_name',
    'window.processName',
    'window.process_name',
  ]);
}

function bundleIdInput(input = {}) {
  return firstPath(input, [
    'bundle_id',
    'bundleId',
    'bundle.identifier',
    'bundle.id',
    'application.bundle_id',
    'application.bundleId',
    'app.bundle_id',
    'app.bundleId',
    'process.bundle_id',
    'process.bundleId',
    'window.bundle_id',
    'window.bundleId',
  ]);
}

function explicitPlatform(input = {}) {
  const value = firstPath(input, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'target_platform',
    'targetPlatform',
    'provider_platform',
    'providerPlatform',
    'app_platform',
    'appPlatform',
    'platform',
    'provider',
    'app',
  ]);
  if (!value || typeof value === 'object') return null;
  try {
    return normalizeMeetingPlatform(String(value));
  } catch {
    return null;
  }
}

function explicitMeetingId(input = {}) {
  return firstPath(input, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting.external_meeting_id',
    'meeting.externalMeetingId',
    'meeting.meeting_no',
    'meeting.meetingNo',
    'meeting_id',
    'meetingId',
    'external_meeting_id',
    'externalMeetingId',
    'meeting_no',
    'meetingNo',
    'session_id',
    'sessionId',
  ]);
}

function observedAtMs(input = {}, options = {}) {
  const value = firstNonEmpty(
    input.observed_at_ms,
    input.observedAtMs,
    input.detected_at_ms,
    input.detectedAtMs,
    input.timestamp_ms,
    input.timestampMs,
    input.timestamp,
    input.time,
    input.ts,
    options.observedAtMs,
    options.receivedAtMs,
  );
  return value == null ? undefined : normalizeAbsoluteMs(value, 'meeting_session_discovery_time');
}

function detectMeetingFromSpecialUrl(input = {}) {
  const rawUrl = urlInput(input);
  const title = titleInput(input);
  const parsed = safeParseUrl(rawUrl);
  if (parsed) {
    const protocol = parsed.protocol.replace(':', '').toLowerCase();
    if (protocol === 'zoommtg' || protocol === 'zoomus') {
      const meetingId = firstNonEmpty(parsed.searchParams.get('confno'), parsed.searchParams.get('meetingid'));
      if (!meetingId) return null;
      return {
        platform: 'zoom',
        meeting_id: cleanId(meetingId),
        external_meeting_id: cleanId(meetingId),
        meeting_url: String(rawUrl),
        title,
        confidence: 'high',
        discovery: { reason: 'url_scheme', confidence: 'high' },
      };
    }
    if (protocol === 'msteams' && String(rawUrl).includes('meetup-join')) {
      const httpsLike = String(rawUrl).replace(/^msteams:\/+/i, 'https://');
      const detected = detectMeetingFromUrl({ url: httpsLike, title });
      if (detected) {
        return {
          ...detected,
          meeting_url: String(rawUrl),
          discovery: { reason: 'url_scheme', confidence: detected.confidence ?? 'high' },
        };
      }
    }
  }
  const direct = detectMeetingFromUrl({ url: rawUrl, title });
  if (direct) {
    return {
      ...direct,
      discovery: { reason: 'url', confidence: direct.confidence ?? 'high' },
    };
  }
  return null;
}

export function detectMeetingApplication(input = {}) {
  const urlMeeting = detectMeetingFromSpecialUrl(input);
  if (urlMeeting) {
    return compactObject({
      platform: urlMeeting.platform,
      confidence: urlMeeting.confidence ?? 'high',
      reason: urlMeeting.discovery?.reason ?? 'url',
      meeting: urlMeeting,
    });
  }
  const platform = explicitPlatform(input);
  if (platform) {
    return compactObject({ platform, confidence: 'explicit', reason: 'explicit_platform' });
  }
  const rawUrl = urlInput(input);
  const parsed = safeParseUrl(rawUrl);
  if (parsed) {
    const byHost = PLATFORM_FINGERPRINTS.find((item) => hostMatches(parsed.hostname, item.hosts));
    if (byHost) return compactObject({ platform: byHost.platform, confidence: 'medium', reason: 'host' });
  }
  const appName = appNameInput(input);
  const processName = processNameInput(input);
  const bundleId = bundleIdInput(input);
  const title = titleInput(input);
  for (const fingerprint of PLATFORM_FINGERPRINTS) {
    if (regexMatches(fingerprint.bundle, bundleId)) {
      return compactObject({ platform: fingerprint.platform, confidence: 'medium', reason: 'bundle_id' });
    }
    if (regexMatches(fingerprint.process, processName) || regexMatches(fingerprint.process, appName)) {
      return compactObject({ platform: fingerprint.platform, confidence: 'medium', reason: 'process' });
    }
    if (regexMatches(fingerprint.title, title)) {
      return compactObject({ platform: fingerprint.platform, confidence: 'low', reason: 'title' });
    }
  }
  return null;
}

function activityFlag(input = {}) {
  return firstBoolean(input, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'meeting.in_meeting',
    'meeting.inMeeting',
    'call.active',
    'callActive',
    'call_active',
    'conference.active',
    'activeMeeting',
  ]);
}

function visibleFlag(input = {}) {
  return firstBoolean(input, ['visible', 'window.visible', 'tab.visible', 'application.visible', 'app.visible']);
}

function activeFlag(input = {}) {
  return firstBoolean(input, ['active', 'selected', 'focused', 'window.active', 'window.focused', 'tab.active', 'tab.highlighted']);
}

function isMeetingLike(input = {}, platform) {
  const activity = activityFlag(input);
  if (activity === true) return true;
  const title = titleInput(input);
  const fingerprint = PLATFORM_FINGERPRINTS.find((item) => item.platform === platform);
  if (!fingerprint) return false;
  return regexMatches(fingerprint.meetingTitle, title);
}

function cleanMeetingTitle(title, platform) {
  let text = normalizeText(title);
  if (!text) return null;
  const generic = {
    google_meet: [/^google meet$/i, /\s*[-|–—]\s*google meet$/i],
    microsoft_teams: [/^microsoft teams$/i, /\s*\|\s*microsoft teams$/i, /\s*[-|–—]\s*microsoft teams$/i],
    zoom: [/^zoom$/i, /^zoom workplace$/i, /^zoom meeting$/i, /\s*[-|–—]\s*zoom meeting$/i],
    lark: [/^lark$/i, /^feishu$/i, /^飞书$/i, /\s*[-|–—]\s*(lark|feishu|飞书)$/i],
    webex: [/^webex$/i, /^cisco webex$/i, /^webex meeting$/i, /\s*[-|–—]\s*webex meeting$/i],
  }[platform] ?? [];
  for (const pattern of generic) text = text.replace(pattern, '').trim();
  text = text.replace(/\s+/g, ' ').trim();
  if (!text || /^(meeting|call|会议|通话)$/i.test(text)) return null;
  return cleanId(text);
}

function stableNativeMeetingId(platform, input = {}) {
  const explicitId = explicitMeetingId(input);
  if (explicitId) return cleanId(explicitId);
  const titleId = cleanMeetingTitle(titleInput(input), platform);
  if (titleId) return `native-${platform}-${titleId}`;
  const windowId = firstPath(input, [
    'window.id',
    'window.window_id',
    'window.windowId',
    'tab.windowId',
    'tab.window_id',
    'windowId',
    'window_id',
    'id',
  ]);
  if (windowId && isMeetingLike(input, platform)) return `native-${platform}-${cleanId(windowId)}`;
  return null;
}

function inheritedSnapshot(raw = {}, inherited = {}) {
  return compactObject({
    ...raw,
    application: raw.application ?? inherited.application,
    app: raw.app ?? inherited.app,
    process: raw.process ?? inherited.process,
    window: raw.window ?? inherited.window,
    browser: raw.browser ?? inherited.browser,
  });
}

function withoutCollections(raw = {}) {
  const {
    windows,
    tabs,
    applications,
    apps,
    processes,
    candidates,
    snapshots,
    items,
    ...rest
  } = raw;
  return rest;
}

function flattenEnvironment(input = {}, inherited = {}) {
  if (Array.isArray(input)) return input.flatMap((item) => flattenEnvironment(item, inherited));
  if (!input || typeof input !== 'object') return [];
  const rows = [];
  const base = inheritedSnapshot(withoutCollections(input), inherited);
  const collections = [
    ['windows', input.windows],
    ['tabs', input.tabs],
    ['applications', input.applications],
    ['apps', input.apps],
    ['processes', input.processes],
    ['candidates', input.candidates],
    ['snapshots', input.snapshots],
    ['items', input.items],
  ].filter(([, value]) => Array.isArray(value));

  for (const [kind, collection] of collections) {
    for (const item of collection) {
      if (kind === 'tabs') {
        rows.push(...flattenEnvironment(item, { ...base, tab: item, window: base.window ?? withoutCollections(input) }));
      } else if (kind === 'windows') {
        rows.push(...flattenEnvironment(item, { ...base, window: withoutCollections(item) }));
      } else if (kind === 'applications' || kind === 'apps') {
        rows.push(...flattenEnvironment(item, { ...base, application: withoutCollections(item), app: withoutCollections(item) }));
      } else if (kind === 'processes') {
        rows.push(...flattenEnvironment(item, { ...base, process: withoutCollections(item) }));
      } else {
        rows.push(...flattenEnvironment(item, base));
      }
    }
  }

  if (collections.length === 0) rows.push(base);
  return rows;
}

function dedupeCandidates(candidates = []) {
  const seen = new Set();
  return candidates.filter((candidate) => {
    const key = [
      candidate.platform,
      candidate.meeting_id,
      firstPath(candidate, ['tab.id', 'window.id', 'id']) ?? '',
    ].join('|');
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function normalizeMeetingSessionCandidate(input = {}, options = {}) {
  const urlMeeting = detectMeetingFromSpecialUrl(input);
  const app = detectMeetingApplication(input);
  const platform = explicitPlatform(input) ?? urlMeeting?.platform ?? app?.platform;
  if (!platform) return null;

  const meetingLike = isMeetingLike(input, platform) || Boolean(urlMeeting) || Boolean(explicitMeetingId(input));
  if (!meetingLike && options.includeIdleApps !== true) return null;

  const meetingId = explicitMeetingId(input)
    ? cleanId(explicitMeetingId(input))
    : urlMeeting?.meeting_id
      ?? stableNativeMeetingId(platform, input);
  if (!meetingId) return null;

  const rawUrl = urlInput(input);
  const title = titleInput(input) ?? urlMeeting?.title;
  const atMs = observedAtMs(input, options);
  const activity = activityFlag(input);
  const visible = visibleFlag(input);
  const active = activeFlag(input);
  const confidence = urlMeeting?.confidence
    ?? (explicitMeetingId(input) ? 'explicit' : app?.confidence ?? 'low');

  return compactObject({
    ...input,
    url: rawUrl,
    meeting_url: urlMeeting?.meeting_url ?? rawUrl,
    title,
    platform,
    meeting_id: meetingId,
    external_meeting_id: firstPath(input, [
      'meeting.external_meeting_id',
      'meeting.externalMeetingId',
      'external_meeting_id',
      'externalMeetingId',
      'meeting_no',
      'meetingNo',
    ]) ?? urlMeeting?.external_meeting_id,
    observedAtMs: atMs,
    active,
    visible,
    in_meeting: activity ?? (urlMeeting ? undefined : meetingLike || undefined),
    meeting: {
      platform,
      meeting_id: meetingId,
      external_meeting_id: firstPath(input, [
        'meeting.external_meeting_id',
        'meeting.externalMeetingId',
        'external_meeting_id',
        'externalMeetingId',
      ]) ?? urlMeeting?.external_meeting_id,
      meeting_url: urlMeeting?.meeting_url ?? rawUrl,
      title,
      confidence,
    },
    discovery: {
      platform_reason: app?.reason ?? urlMeeting?.discovery?.reason,
      confidence,
      app_name: appNameInput(input),
      process_name: processNameInput(input),
      bundle_id: bundleIdInput(input),
    },
  });
}

export function normalizeMeetingSessionCandidates(input = {}, options = {}) {
  const candidates = flattenEnvironment(input)
    .map((item) => normalizeMeetingSessionCandidate(item, options))
    .filter(Boolean);
  return dedupeCandidates(candidates);
}

export function selectMeetingSessionCandidate(input = {}, options = {}) {
  const candidates = normalizeMeetingSessionCandidates(input, options);
  return {
    ...selectMeetingSnapshot(candidates, options),
    normalizedCandidates: candidates,
  };
}

export function createMeetingSessionDiscoveryObserver(options = {}) {
  const observer = createLocalMeetingObserver(options);
  return {
    observeEnvironment(input = {}, observeOptions = {}) {
      const candidates = normalizeMeetingSessionCandidates(input, {
        ...options,
        ...observeOptions,
      });
      return {
        ...observer.observeCandidates(candidates, observeOptions),
        normalizedCandidates: candidates,
      };
    },
    observeCandidate(input = {}, observeOptions = {}) {
      const candidate = normalizeMeetingSessionCandidate(input, {
        ...options,
        ...observeOptions,
      });
      const snapshot = candidate ?? {
        active: false,
        observedAtMs: observeOptions.observedAtMs ?? observeOptions.receivedAtMs,
      };
      return {
        ...observer.observe(snapshot, observeOptions),
        normalizedCandidate: candidate,
      };
    },
    select(input = {}, selectOptions = {}) {
      return selectMeetingSessionCandidate(input, {
        ...options,
        ...selectOptions,
      });
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}

export function createMeetingSessionTimelineDiscovery(client, options = {}) {
  const observer = createLocalMeetingTimelineObserver(client, options);
  return {
    async observeEnvironment(input = {}, observeOptions = {}) {
      const candidates = normalizeMeetingSessionCandidates(input, {
        ...options,
        ...observeOptions,
      });
      return {
        ...(await observer.observeCandidates(candidates, observeOptions)),
        normalizedCandidates: candidates,
      };
    },
    async observeCandidate(input = {}, observeOptions = {}) {
      const candidate = normalizeMeetingSessionCandidate(input, {
        ...options,
        ...observeOptions,
      });
      const snapshot = candidate ?? {
        active: false,
        observedAtMs: observeOptions.observedAtMs ?? observeOptions.receivedAtMs,
      };
      return {
        ...(await observer.observe(snapshot, observeOptions)),
        normalizedCandidate: candidate,
      };
    },
    select(input = {}, selectOptions = {}) {
      return selectMeetingSessionCandidate(input, {
        ...options,
        ...selectOptions,
      });
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}
