import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  createBrowserMeetingObserver,
  createBrowserMeetingTimelineObserver,
  observeBrowserMeetingSample,
} from './browser-meeting.mjs';
import { detectMeetingApplication } from './meeting-session-discovery.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

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

function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

function flatten(value) {
  return asArray(value).flatMap((item) => Array.isArray(item) ? flatten(item) : [item]);
}

function normalizeText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function compactText(value) {
  const text = normalizeText(value);
  return text || undefined;
}

function maybeTimeMs(input = {}, options = {}) {
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
  return value == null ? undefined : normalizeAbsoluteMs(value, 'meeting_app_observed_at_ms');
}

function tryPlatform(value) {
  if (!value || typeof value === 'object') return null;
  try {
    return normalizeMeetingPlatform(String(value));
  } catch {
    return null;
  }
}

function pathValue(input = {}, paths = []) {
  return firstPath(input, paths);
}

function arrayAtPaths(input = {}, paths = []) {
  return paths.flatMap((path) => {
    const value = getPath(input, path);
    return Array.isArray(value) ? value : [];
  });
}

function collectStrings(value, output = [], depth = 0) {
  if (value == null || depth > 3) return output;
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    const text = compactText(value);
    if (text) output.push(text);
    return output;
  }
  if (Array.isArray(value)) {
    for (const item of value.slice(0, 80)) collectStrings(item, output, depth + 1);
    return output;
  }
  if (typeof value === 'object') {
    for (const key of [
      'text',
      'label',
      'ariaLabel',
      'aria_label',
      'title',
      'name',
      'role',
      'status',
      'caption',
      'tooltip',
      'alt',
    ]) {
      collectStrings(value[key], output, depth + 1);
    }
  }
  return output;
}

function textCorpus(input = {}, preset = {}) {
  const fields = [
    'title',
    'topic',
    'name',
    'url',
    'href',
    'document_title',
    'documentTitle',
    'window.title',
    'tab.title',
    'page.title',
    'dom.title',
    'application.name',
    'app.name',
    'process.name',
    'processName',
    'bundle_id',
    'bundleId',
    'status',
    'call.status',
    'meeting.status',
    ...(preset.textPaths ?? []),
  ];
  const direct = fields.map((path) => getPath(input, path)).filter((value) => value != null);
  const collections = [
    'buttons',
    'controls',
    'labels',
    'texts',
    'ariaLabels',
    'aria_labels',
    'page.buttons',
    'page.controls',
    'page.labels',
    'page.texts',
    'page.ariaLabels',
    'dom.buttons',
    'dom.controls',
    'dom.labels',
    'dom.texts',
    'dom.ariaLabels',
    'window.buttons',
    'window.controls',
    'window.labels',
    'window.texts',
    'tab.buttons',
    'tab.controls',
    'tab.labels',
    'tab.texts',
    'accessibility.buttons',
    'accessibility.controls',
    'accessibility.labels',
    'accessibility.texts',
    ...(preset.textCollectionPaths ?? []),
  ].flatMap((path) => asArray(getPath(input, path)));
  return [...direct, ...collections].flatMap((value) => collectStrings(value)).filter(Boolean);
}

function textMatches(patterns = [], text) {
  const value = normalizeText(text);
  if (!value) return false;
  return patterns.some((pattern) => {
    if (pattern instanceof RegExp) return pattern.test(value);
    return value.toLowerCase().includes(String(pattern).toLowerCase());
  });
}

function corpusMatches(patterns = [], corpus = []) {
  return corpus.some((text) => textMatches(patterns, text));
}

const COMMON_ACTIVE_SPEAKER_PATHS = Object.freeze([
  'activeSpeaker',
  'active_speaker',
  'speaker',
  'dominantSpeaker',
  'dominant_speaker',
  'currentSpeaker',
  'current_speaker',
  'page.activeSpeaker',
  'page.active_speaker',
  'page.dominantSpeaker',
  'dom.activeSpeaker',
  'dom.active_speaker',
  'dom.dominantSpeaker',
  'tab.activeSpeaker',
  'window.activeSpeaker',
  'accessibility.activeSpeaker',
  'audio.activeSpeaker',
]);

const COMMON_PARTICIPANT_PATHS = Object.freeze([
  'participants',
  'participantTiles',
  'participant_tiles',
  'tiles',
  'videoTiles',
  'video_tiles',
  'people',
  'attendees',
  'members',
  'roster',
  'page.participants',
  'page.participantTiles',
  'page.tiles',
  'page.videoTiles',
  'dom.participants',
  'dom.participantTiles',
  'dom.tiles',
  'dom.videoTiles',
  'accessibility.participants',
  'accessibility.participantTiles',
  'accessibility.tiles',
  'window.participants',
  'window.participantTiles',
  'window.tiles',
  'window.videoTiles',
  'audio.participants',
  'audio.tiles',
]);

const NAME_PATHS = Object.freeze([
  'name',
  'display_name',
  'displayName',
  'user_name',
  'userName',
  'participant_name',
  'participantName',
  'title',
  'label',
  'ariaLabel',
  'aria_label',
  'text',
  'tooltip',
]);

const ID_PATHS = Object.freeze([
  'id',
  'speaker_id',
  'speakerId',
  'participant_id',
  'participantId',
  'user_id',
  'userId',
  'data_participant_id',
  'dataParticipantId',
  'dataset.participantId',
  'dataset.participant_id',
  'dataset.userId',
]);

const SPEAKING_PATHS = Object.freeze([
  'speaking',
  'is_speaking',
  'isSpeaking',
  'active_speaker',
  'activeSpeaker',
  'active',
  'audio.active',
  'audio.speaking',
  'audio.isSpeaking',
  'voice.active',
  'voice.speaking',
]);

const AUDIO_LEVEL_PATHS = Object.freeze([
  'audioLevel',
  'audio_level',
  'volume',
  'volumeLevel',
  'voiceActivity',
  'voice_activity',
  'speechLevel',
  'speech_level',
  'audio.level',
  'audio.volume',
  'audio.voiceActivity',
]);

const COMMON_SPEAKING_HINTS = Object.freeze([
  /\bis speaking\b/i,
  /\bspeaking\b/i,
  /\btalking\b/i,
  /\bactive speaker\b/i,
  /正在发言/,
  /正在讲话/,
  /正在说话/,
]);

const COMMON_MUTED_HINTS = Object.freeze([
  /\bmuted\b/i,
  /\bmicrophone off\b/i,
  /\bmic off\b/i,
  /已静音/,
  /麦克风已关闭/,
]);

export const MEETING_APP_PRESETS = Object.freeze({
  google_meet: {
    platform: 'google_meet',
    displayName: 'Google Meet',
    joinedHints: [
      /\bleave call\b/i,
      /\bleave meeting\b/i,
      /\bend call\b/i,
      /\byou are presenting\b/i,
      /\bstop presenting\b/i,
      /离开通话/,
      /退出通话/,
      /结束通话/,
      /正在展示/,
    ],
    preJoinHints: [
      /\bjoin now\b/i,
      /\bask to join\b/i,
      /\bready to join\b/i,
      /\bcheck your audio and video\b/i,
      /申请加入/,
      /立即加入/,
      /准备加入/,
    ],
    participantPaths: COMMON_PARTICIPANT_PATHS,
    activeSpeakerPaths: COMMON_ACTIVE_SPEAKER_PATHS,
  },
  microsoft_teams: {
    platform: 'microsoft_teams',
    displayName: 'Microsoft Teams',
    joinedHints: [
      /\bleave\b/i,
      /\bhang up\b/i,
      /\bturn camera off\b/i,
      /\bmute microphone\b/i,
      /\bmicrosoft teams meeting\b/i,
      /离开/,
      /挂断/,
      /会议聊天/,
    ],
    preJoinHints: [
      /\bjoin now\b/i,
      /\bpre-join\b/i,
      /\bchoose your audio and video settings\b/i,
      /立即加入/,
      /加入会议/,
    ],
    participantPaths: COMMON_PARTICIPANT_PATHS,
    activeSpeakerPaths: COMMON_ACTIVE_SPEAKER_PATHS,
  },
  zoom: {
    platform: 'zoom',
    displayName: 'Zoom',
    joinedHints: [
      /\bleave meeting\b/i,
      /\bend meeting\b/i,
      /\bmute audio\b/i,
      /\bstop video\b/i,
      /\bparticipants\b/i,
      /离开会议/,
      /结束会议/,
      /参与者/,
    ],
    preJoinHints: [
      /\bjoin meeting\b/i,
      /\bjoin with computer audio\b/i,
      /\bwaiting room\b/i,
      /加入会议/,
      /等候室/,
    ],
    participantPaths: COMMON_PARTICIPANT_PATHS,
    activeSpeakerPaths: COMMON_ACTIVE_SPEAKER_PATHS,
  },
  lark: {
    platform: 'lark',
    displayName: 'Lark/Feishu',
    joinedHints: [
      /\bleave\b/i,
      /\bend call\b/i,
      /\bai summary\b/i,
      /离开会议/,
      /结束通话/,
      /AI 总结/,
      /会议进展实时可视化/,
    ],
    preJoinHints: [/加入会议/, /立即加入/, /\bjoin now\b/i],
    participantPaths: COMMON_PARTICIPANT_PATHS,
    activeSpeakerPaths: COMMON_ACTIVE_SPEAKER_PATHS,
  },
  webex: {
    platform: 'webex',
    displayName: 'Cisco Webex',
    joinedHints: [
      /\bleave meeting\b/i,
      /\bend meeting\b/i,
      /\bmute microphone\b/i,
      /\bparticipants\b/i,
      /离开会议/,
      /结束会议/,
    ],
    preJoinHints: [/\bjoin meeting\b/i, /\bstart meeting\b/i, /加入会议/],
    participantPaths: COMMON_PARTICIPANT_PATHS,
    activeSpeakerPaths: COMMON_ACTIVE_SPEAKER_PATHS,
  },
});

function presetForPlatform(platform) {
  const normalized = tryPlatform(platform);
  return normalized ? MEETING_APP_PRESETS[normalized] ?? null : null;
}

export function detectMeetingAppPreset(input = {}, options = {}) {
  const explicit = tryPlatform(firstPath(input, [
    'meeting.platform',
    'meeting_platform',
    'meetingPlatform',
    'detected_platform',
    'detectedPlatform',
    'platform',
    'provider',
  ]) ?? options.platform);
  const detected = explicit
    ? { platform: explicit, confidence: 'explicit', reason: 'explicit_platform' }
    : detectMeetingApplication(input);
  const preset = presetForPlatform(detected?.platform);
  if (!preset) return null;
  return {
    ...preset,
    confidence: detected?.confidence,
    reason: detected?.reason,
    detected,
  };
}

function explicitInMeeting(input = {}) {
  return firstBoolean(input, [
    'in_meeting',
    'inMeeting',
    'meeting.active',
    'meeting.in_meeting',
    'meeting.inMeeting',
    'page.in_meeting',
    'page.inMeeting',
    'page.has_joined',
    'page.hasJoined',
    'page.call_active',
    'page.callActive',
    'dom.in_meeting',
    'dom.inMeeting',
    'dom.has_joined',
    'dom.hasJoined',
    'accessibility.in_meeting',
    'accessibility.inMeeting',
    'call.active',
    'callActive',
    'call_active',
  ]);
}

function activeFlag(input = {}) {
  return firstBoolean(input, [
    'active',
    'focused',
    'selected',
    'tab.active',
    'tab.highlighted',
    'window.focused',
    'window.active',
    'page.focused',
  ]);
}

function visibleFlag(input = {}) {
  return firstBoolean(input, [
    'visible',
    'tab.visible',
    'window.visible',
    'page.visible',
    'page.document_visible',
    'page.documentVisible',
    'dom.visible',
  ]);
}

function audibleFlag(input = {}) {
  return firstBoolean(input, [
    'audible',
    'tab.audible',
    'page.audible',
    'dom.audible',
    'audio.active',
    'audio.audible',
    'audio.hasAudio',
  ]);
}

function inferInMeeting(input = {}, preset = {}) {
  const explicit = explicitInMeeting(input);
  if (explicit != null) return explicit;
  const corpus = textCorpus(input, preset);
  if (corpusMatches(preset.preJoinHints, corpus)) return false;
  if (corpusMatches(preset.joinedHints, corpus)) return true;
  if (firstSpeaker(input, preset)) return true;
  if (audibleFlag(input) === true && visibleFlag(input) !== false) return true;
  return undefined;
}

function participantRows(input = {}, preset = {}) {
  const seen = new Set();
  return flatten(arrayAtPaths(input, [
    ...(preset.participantPaths ?? []),
    ...COMMON_PARTICIPANT_PATHS,
  ])).filter((item) => {
    if (!item || typeof item !== 'object') return false;
    if (seen.has(item)) return false;
    seen.add(item);
    return true;
  });
}

function speakerObject(input = {}, preset = {}) {
  return firstPath(input, [
    ...(preset.activeSpeakerPaths ?? []),
    ...COMMON_ACTIVE_SPEAKER_PATHS,
  ]);
}

function labelText(input = {}) {
  return firstPath(input, ['ariaLabel', 'aria_label', 'label', 'title', 'text', 'name', 'displayName']);
}

function cleanParticipantName(value) {
  let text = normalizeText(value);
  if (!text) return undefined;
  text = text
    .replace(/\b(is speaking|speaking|talking|active speaker|muted|microphone off|mic off|camera off|presenting)\b/ig, '')
    .replace(/(正在发言|正在讲话|正在说话|已静音|麦克风已关闭|摄像头已关闭|正在展示)/g, '')
    .replace(/[，,].*$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return text || undefined;
}

function firstNumber(raw = {}, paths = []) {
  for (const path of paths) {
    const value = getPath(raw, path);
    if (value == null || value === '') continue;
    const numeric = Number(value);
    if (Number.isFinite(numeric)) return numeric;
  }
  return undefined;
}

function participantSpeaking(raw = {}, preset = {}) {
  const explicit = firstBoolean(raw, [
    ...(preset.speakingPaths ?? []),
    ...SPEAKING_PATHS,
  ]);
  if (explicit != null) return explicit;
  const level = firstNumber(raw, [
    ...(preset.audioLevelPaths ?? []),
    ...AUDIO_LEVEL_PATHS,
  ]);
  if (level != null) return level >= Number(preset.audioLevelThreshold ?? 0.2);
  const texts = collectStrings(raw);
  if (corpusMatches([
    ...(preset.speakingHints ?? []),
    ...COMMON_SPEAKING_HINTS,
  ], texts)) return true;
  if (corpusMatches([
    ...(preset.mutedHints ?? []),
    ...COMMON_MUTED_HINTS,
  ], texts)) return false;
  return undefined;
}

function normalizeParticipant(raw = {}, preset = {}) {
  const id = pathValue(raw, [
    ...(preset.participantIdPaths ?? []),
    ...ID_PATHS,
  ]);
  const name = pathValue(raw, [
    ...(preset.participantNamePaths ?? []),
    ...NAME_PATHS,
  ]);
  const label = labelText(raw);
  const speaking = participantSpeaking(raw, preset);
  return compactObject({
    id: id == null ? undefined : String(id),
    name: cleanParticipantName(name) ?? cleanParticipantName(label),
    display_name: cleanParticipantName(name) ?? cleanParticipantName(label),
    speaking,
    isSpeaking: speaking,
    active_speaker: speaking === true ? true : undefined,
    raw,
  });
}

function normalizedParticipants(input = {}, preset = {}) {
  return participantRows(input, preset)
    .map((item) => normalizeParticipant(item, preset))
    .filter((item) => item.id || item.name || item.speaking != null);
}

function normalizeSpeaker(raw = {}, preset = {}) {
  if (!raw || typeof raw !== 'object') return null;
  const participant = normalizeParticipant(raw, preset);
  if (!participant.id && !participant.name) return null;
  return compactObject({
    id: participant.id,
    name: participant.name,
    display_name: participant.display_name,
    speaking: participant.speaking ?? true,
    raw,
  });
}

function firstSpeaker(input = {}, preset = {}) {
  const direct = normalizeSpeaker(speakerObject(input, preset), preset);
  if (direct) return direct;
  return normalizedParticipants(input, preset).find((participant) => participant.speaking === true) ?? null;
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
    'location.href',
    'page.url',
    'page.href',
    'dom.url',
    'window.url',
    'browser.url',
    'tab.url',
    'frame.url',
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
    'document_title',
    'documentTitle',
    'page.title',
    'dom.title',
    'window.title',
    'tab.title',
    'frame.title',
  ]);
}

function withoutCollections(raw = {}) {
  const {
    windows,
    tabs,
    pages,
    frames,
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

function appRows(input = {}, inherited = {}) {
  if (Array.isArray(input)) return input.flatMap((item) => appRows(item, inherited));
  if (!input || typeof input !== 'object') return [];
  const rows = [];
  const base = compactObject({
    ...inherited,
    ...withoutCollections(input),
    browser: input.browser ?? inherited.browser,
    application: input.application ?? input.app ?? inherited.application,
    app: input.app ?? input.application ?? inherited.app,
    process: input.process ?? inherited.process,
    window: input.window ?? inherited.window,
    tab: input.tab ?? inherited.tab,
    page: input.page ?? inherited.page,
    dom: input.dom ?? inherited.dom,
    accessibility: input.accessibility ?? input.ax ?? inherited.accessibility,
  });
  const collections = [
    ['applications', input.applications],
    ['apps', input.apps],
    ['processes', input.processes],
    ['windows', input.windows],
    ['tabs', input.tabs],
    ['pages', input.pages],
    ['frames', input.frames],
    ['candidates', input.candidates],
    ['snapshots', input.snapshots],
    ['items', input.items],
  ].filter(([, value]) => Array.isArray(value));

  for (const [kind, collection] of collections) {
    for (const item of collection) {
      if (kind === 'applications' || kind === 'apps') {
        rows.push(...appRows(item, {
          ...base,
          application: withoutCollections(item),
          app: withoutCollections(item),
          process: item.process ?? base.process,
        }));
      } else if (kind === 'processes') {
        rows.push(...appRows(item, {
          ...base,
          process: withoutCollections(item),
        }));
      } else if (kind === 'windows') {
        rows.push(...appRows(item, {
          ...base,
          window: withoutCollections(item),
        }));
      } else if (kind === 'tabs') {
        rows.push(...appRows(item, {
          ...base,
          tab: withoutCollections(item),
          window: base.window ?? withoutCollections(input),
        }));
      } else if (kind === 'pages') {
        rows.push(...appRows(item, {
          ...base,
          page: withoutCollections(item),
        }));
      } else if (kind === 'frames') {
        rows.push(...appRows(item, {
          ...base,
          frame: withoutCollections(item),
        }));
      } else {
        rows.push(...appRows(item, base));
      }
    }
  }

  if (collections.length === 0) rows.push(base);
  return rows;
}

export function normalizeMeetingAppSnapshot(input = {}, options = {}) {
  const preset = options.preset ?? detectMeetingAppPreset(input, options);
  if (!preset) return null;
  const participants = normalizedParticipants(input, preset);
  const activeSpeaker = firstSpeaker(input, preset);
  const platform = preset.platform;
  const detectedMeeting = preset.detected?.meeting ?? {};
  const meetingId = firstPath(input, [
    'meeting.meeting_id',
    'meeting.meetingId',
    'meeting.id',
    'meeting_id',
    'meetingId',
    'external_meeting_id',
    'externalMeetingId',
    'session_id',
    'sessionId',
  ]) ?? detectedMeeting.meeting_id;
  const meetingUrl = urlInput(input) ?? detectedMeeting.meeting_url;
  const title = titleInput(input) ?? detectedMeeting.title;
  return compactObject({
    ...input,
    platform,
    provider: platform,
    url: meetingUrl,
    meeting_url: meetingUrl,
    title,
    meeting_id: meetingId,
    external_meeting_id: firstPath(input, [
      'meeting.external_meeting_id',
      'meeting.externalMeetingId',
      'external_meeting_id',
      'externalMeetingId',
    ]) ?? detectedMeeting.external_meeting_id,
    meeting: meetingId ? {
      platform,
      meeting_id: meetingId,
      external_meeting_id: detectedMeeting.external_meeting_id,
      meeting_url: meetingUrl,
      title,
    } : undefined,
    observedAtMs: maybeTimeMs(input, options),
    active: activeFlag(input),
    visible: visibleFlag(input),
    audible: audibleFlag(input),
    inMeeting: inferInMeeting(input, preset),
    activeSpeaker,
    participants: participants.length ? participants : undefined,
    meeting_app: {
      platform,
      display_name: preset.displayName,
      preset_reason: preset.reason,
      preset_confidence: preset.confidence,
    },
  });
}

export function normalizeMeetingAppSnapshots(input = {}, options = {}) {
  return appRows(input)
    .map((item) => normalizeMeetingAppSnapshot(item, options))
    .filter(Boolean);
}

export function observeMeetingAppSample(state = null, input = {}, options = {}) {
  return observeBrowserMeetingSample(state, {
    candidates: normalizeMeetingAppSnapshots(input, options),
  }, {
    source: options.source ?? 'meeting_app_observer',
    ...options,
  });
}

export function createMeetingAppObserver(options = {}) {
  const observer = createBrowserMeetingObserver({
    source: options.source ?? 'meeting_app_observer',
    ...options,
  });
  return {
    observe(input = {}, observeOptions = {}) {
      return observer.observe({
        candidates: normalizeMeetingAppSnapshots(input, {
          ...options,
          ...observeOptions,
        }),
      }, observeOptions);
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}

export function createMeetingAppTimelineObserver(client, options = {}) {
  const observer = createBrowserMeetingTimelineObserver(client, {
    source: options.source ?? 'meeting_app_observer',
    ...options,
  });
  return {
    observe(input = {}, observeOptions = {}) {
      return observer.observe({
        candidates: normalizeMeetingAppSnapshots(input, {
          ...options,
          ...observeOptions,
        }),
      }, observeOptions);
    },
    getState() {
      return observer.getState();
    },
    reset(nextState = null) {
      return observer.reset(nextState);
    },
  };
}
