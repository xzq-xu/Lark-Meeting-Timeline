import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import {
  createBrowserMeetingObserver,
  createBrowserMeetingTimelineObserver,
  observeBrowserMeetingSample,
} from './browser-meeting.mjs';
import { detectMeetingApplication } from './meeting-session-discovery.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_APP_ADAPTER_FIT_SCHEMA = 'meeting_app_adapter_fit_report';
export const MEETING_APP_ADAPTER_FIT_MATRIX_SCHEMA = 'meeting_app_adapter_fit_matrix';
export const MEETING_APP_ADAPTER_FIT_SCHEMA_VERSION = 1;

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

function uniqueList(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
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
      /\bpresent now\b/i,
      /\bturn off microphone\b/i,
      /\bturn off camera\b/i,
      /\bmore options\b/i,
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
      /\bshow conversation\b/i,
      /\braise hand\b/i,
      /\bshare content\b/i,
      /离开/,
      /挂断/,
      /会议聊天/,
      /举手/,
      /共享屏幕/,
      /共享内容/,
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
      /\bunmute\b/i,
      /\bstop video\b/i,
      /\bstart video\b/i,
      /\bparticipants\b/i,
      /\breactions\b/i,
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
      /\bhang up\b/i,
      /\bshare screen\b/i,
      /离开会议/,
      /结束通话/,
      /挂断/,
      /AI 总结/,
      /AI 视图/,
      /会议进展实时可视化/,
      /共享屏幕/,
      /关闭麦克风/,
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
      /\bunmute\b/i,
      /\bstart video\b/i,
      /\bstop video\b/i,
      /\bparticipants\b/i,
      /\bchat\b/i,
      /离开会议/,
      /结束会议/,
      /与会者/,
      /聊天/,
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
  const discovered = detectMeetingApplication(input);
  const detected = explicit
    ? {
      ...discovered,
      platform: explicit,
      confidence: 'explicit',
      reason: 'explicit_platform',
    }
    : discovered;
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
    .replace(/\b(you|me)\b/ig, '')
    .replace(/(正在发言|正在讲话|正在说话|已静音|麦克风已关闭|摄像头已关闭|正在展示|我|本人)/g, '')
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

function fitIssue(severity, code, message, details = {}) {
  return compactObject({
    severity,
    code,
    message,
    ...details,
  });
}

function inputShapeHints(value, output = new Set(), depth = 0) {
  if (value == null || depth > 4) return output;
  if (Array.isArray(value)) {
    if (value.length > 0) output.add('array');
    for (const item of value.slice(0, 20)) inputShapeHints(item, output, depth + 1);
    return output;
  }
  if (typeof value !== 'object') return output;
  if (value.querySelectorAll || value.document?.querySelectorAll || value.window?.document?.querySelectorAll) output.add('live_dom');
  if (value.windows) output.add('browser_windows');
  if (value.tabs) output.add('browser_tabs');
  if (value.pages || value.frames) output.add('browser_pages');
  if (value.url || value.href || value.tab?.url || value.page?.url || value.window?.url) output.add('meeting_url_candidate');
  if (value.applications || value.apps || value.processes || value.application || value.app || value.process) output.add('native_app_snapshot');
  if (value.accessibility || value.ax) output.add('accessibility_snapshot');
  if (value.page || value.dom || value.controls || value.buttons || value.participants || value.tiles || value.texts) output.add('dom_like_snapshot');
  for (const key of ['windows', 'tabs', 'pages', 'frames', 'applications', 'apps', 'processes', 'candidates', 'snapshots', 'items']) {
    inputShapeHints(value[key], output, depth + 1);
  }
  return output;
}

function fitSurface(input = {}, rows = []) {
  const shapes = inputShapeHints(input);
  if (shapes.has('browser_tabs') || rows.some((row) => row.url?.startsWith?.('http'))) return 'browser_extension_or_webview';
  if (shapes.has('native_app_snapshot') || shapes.has('accessibility_snapshot')) return 'native_detector';
  if (shapes.has('live_dom') || shapes.has('dom_like_snapshot')) return 'embedded_webview_or_content_script';
  return 'unknown';
}

function snapshotRow(snapshot = {}, index) {
  const participants = asArray(snapshot.participants);
  const activeSpeaker = snapshot.activeSpeaker ?? snapshot.active_speaker;
  return compactObject({
    index,
    platform: snapshot.platform,
    meeting_id: snapshot.meeting_id ?? snapshot.meeting?.meeting_id,
    meeting_url: snapshot.meeting_url ?? snapshot.url ?? snapshot.meeting?.meeting_url,
    title: snapshot.title,
    in_meeting: snapshot.inMeeting,
    active: snapshot.active,
    visible: snapshot.visible,
    audible: snapshot.audible,
    participant_count: participants.length,
    active_speaker_id: activeSpeaker?.id,
    active_speaker_name: activeSpeaker?.name ?? activeSpeaker?.display_name,
    has_active_speaker: Boolean(activeSpeaker?.id || activeSpeaker?.name || activeSpeaker?.display_name),
  });
}

function fitNextActions(issues = []) {
  return uniqueList(issues.map((item) => {
    if (item.code === 'no_meeting_app_candidate') return 'send_browser_tabs_or_native_window_snapshot_to_adapter';
    if (item.code === 'target_platform_not_detected') return 'include_platform_or_platform_specific_url_in_snapshot';
    if (item.code === 'missing_meeting_identity') return 'include_meeting_url_title_or_stable_window_id';
    if (item.code === 'missing_meeting_start_candidate') return 'capture_joined_meeting_state_with_leave_or_call_controls_visible';
    if (item.code === 'missing_active_speaker') return 'capture_active_speaker_tile_or_audio_activity';
    if (item.code === 'missing_participants') return 'capture_participant_tiles_or_roster_rows';
    if (item.code === 'missing_meeting_end_candidate') return 'capture_after_leave_or_closed_window_state_for_end_detection';
    return item.code;
  }));
}

export function buildMeetingAppAdapterFitReport(input = {}, options = {}) {
  const expectedPlatform = tryPlatform(firstNonEmpty(
    options.platform,
    options.provider,
    input.platform,
    input.provider,
  ));
  const normalized = normalizeMeetingAppSnapshots(input, expectedPlatform ? { ...options, platform: expectedPlatform } : options);
  const rows = normalized.map((snapshot, index) => snapshotRow(snapshot, index));
  const detectedPlatforms = uniqueList(rows.map((row) => row.platform));
  const targetDetected = expectedPlatform ? detectedPlatforms.includes(expectedPlatform) : detectedPlatforms.length > 0;
  const participantCount = rows.reduce((count, row) => count + Number(row.participant_count ?? 0), 0);
  const activeSpeakerCount = rows.filter((row) => row.has_active_speaker).length;
  const startCandidateCount = rows.filter((row) => row.in_meeting === true).length;
  const endCandidateCount = rows.filter((row) => row.in_meeting === false).length;
  const identityCount = rows.filter((row) => row.meeting_id || row.meeting_url).length;
  const activeVisibleOrAudibleCount = rows.filter((row) => row.active === true || row.visible === true || row.audible === true).length;
  const coverage = {
    candidate_observation: rows.length > 0,
    platform_detected: detectedPlatforms.length > 0,
    target_platform_detected: targetDetected,
    meeting_identity: identityCount > 0,
    active_or_visible_or_audible: activeVisibleOrAudibleCount > 0,
    meeting_start_candidate: startCandidateCount > 0,
    meeting_end_candidate: endCandidateCount > 0,
    active_speaker: activeSpeakerCount > 0,
    participants: participantCount > 0,
    realtime_axis_candidate: targetDetected && identityCount > 0 && startCandidateCount > 0,
    speaker_track_candidate: targetDetected && identityCount > 0 && startCandidateCount > 0 && activeSpeakerCount > 0,
    participant_track_candidate: targetDetected && identityCount > 0 && startCandidateCount > 0 && participantCount > 0,
  };
  const issues = [];
  if (rows.length === 0) {
    issues.push(fitIssue('error', 'no_meeting_app_candidate', 'No supported meeting app candidate was detected from the supplied host input.'));
  }
  if (expectedPlatform && !targetDetected) {
    issues.push(fitIssue('error', 'target_platform_not_detected', 'The expected meeting app platform was not detected.', {
      expected_platform: expectedPlatform,
      detected_platforms: detectedPlatforms,
    }));
  }
  if (rows.length > 0 && identityCount === 0) {
    issues.push(fitIssue('error', 'missing_meeting_identity', 'No meeting id or meeting URL was inferred from the host input.'));
  }
  if (rows.length > 0 && startCandidateCount === 0) {
    issues.push(fitIssue('error', 'missing_meeting_start_candidate', 'No joined/in-meeting candidate was inferred from the host input.'));
  }
  if (rows.length > 0 && activeSpeakerCount === 0) {
    issues.push(fitIssue('warning', 'missing_active_speaker', 'No active speaker was inferred; speaker timeline markers will be unavailable until this signal is supplied.'));
  }
  if (rows.length > 0 && participantCount === 0) {
    issues.push(fitIssue('warning', 'missing_participants', 'No participants were inferred; participant track and speaker attribution may be weak.'));
  }
  if (rows.length > 0 && endCandidateCount === 0) {
    issues.push(fitIssue('warning', 'missing_meeting_end_candidate', 'No ended/left meeting candidate was observed; end events need a later sample or provider reconciliation.'));
  }
  const accepted = issues.every((item) => item.severity !== 'error');
  return {
    type: 'meeting_app_adapter_fit_report',
    schema: MEETING_APP_ADAPTER_FIT_SCHEMA,
    version: MEETING_APP_ADAPTER_FIT_SCHEMA_VERSION,
    platform: expectedPlatform ?? (detectedPlatforms.length === 1 ? detectedPlatforms[0] : undefined),
    expected_platform: expectedPlatform,
    detected_platforms: detectedPlatforms,
    accepted,
    ready_for_realtime_axis: coverage.realtime_axis_candidate,
    ready_for_speaker_track: coverage.speaker_track_candidate,
    ready_for_participant_track: coverage.participant_track_candidate,
    recommended_surface: fitSurface(input, rows),
    input_shapes: [...inputShapeHints(input)].sort(),
    candidate_count: rows.length,
    meeting_identity_count: identityCount,
    start_candidate_count: startCandidateCount,
    end_candidate_count: endCandidateCount,
    active_speaker_count: activeSpeakerCount,
    participant_count: participantCount,
    coverage,
    rows,
    issues,
    next_actions: fitNextActions(issues),
  };
}

function matrixInputForPlatform(platform, merged = {}, fallbackInput = {}) {
  const source = merged.inputs ?? merged.inputByPlatform ?? merged.input_by_platform ?? {};
  const dashed = platform.replaceAll('_', '-');
  const aliases = platform === 'microsoft_teams'
    ? ['teams', 'microsoft-teams']
    : platform === 'google_meet'
      ? ['google-meet', 'meet']
      : platform === 'lark'
        ? ['feishu', 'larksuite']
        : [dashed];
  for (const key of [platform, dashed, ...aliases]) {
    if (source[key] != null) return source[key];
  }
  return fallbackInput;
}

export function buildMeetingAppAdapterFitMatrix(input = {}, options = {}) {
  const objectInput = input && typeof input === 'object' && !Array.isArray(input) ? input : {};
  const merged = { ...objectInput, ...options };
  const source = merged.inputs ?? merged.inputByPlatform ?? merged.input_by_platform;
  const fallbackInput = firstNonEmpty(merged.input, merged.snapshot, merged.snapshots, input);
  let platforms = uniqueList(asArray(firstNonEmpty(
    merged.platforms,
    merged.platform_keys,
    merged.platformKeys,
    source && typeof source === 'object' ? Object.keys(source) : undefined,
  )).map((platform) => tryPlatform(platform)).filter(Boolean));
  if (platforms.length === 0) {
    const report = buildMeetingAppAdapterFitReport(fallbackInput, options);
    platforms = report.detected_platforms.length > 0 ? report.detected_platforms : [];
    const reports = platforms.length > 0
      ? platforms.map((platform) => buildMeetingAppAdapterFitReport(fallbackInput, { ...options, platform }))
      : [report];
    return fitMatrixFromReports(reports);
  }
  return fitMatrixFromReports(platforms.map((platform) => buildMeetingAppAdapterFitReport(
    matrixInputForPlatform(platform, merged, fallbackInput),
    { ...merged, platform },
  )));
}

function fitMatrixFromReports(reports = []) {
  return {
    type: 'meeting_app_adapter_fit_matrix',
    schema: MEETING_APP_ADAPTER_FIT_MATRIX_SCHEMA,
    version: MEETING_APP_ADAPTER_FIT_SCHEMA_VERSION,
    platform_count: reports.length,
    accepted_count: reports.filter((report) => report.accepted).length,
    realtime_axis_ready_count: reports.filter((report) => report.ready_for_realtime_axis).length,
    speaker_track_ready_count: reports.filter((report) => report.ready_for_speaker_track).length,
    participant_track_ready_count: reports.filter((report) => report.ready_for_participant_track).length,
    platforms: uniqueList(reports.flatMap((report) => report.platform ?? report.detected_platforms ?? [])),
    rows: reports.map((report) => ({
      platform: report.platform,
      expected_platform: report.expected_platform,
      detected_platforms: report.detected_platforms,
      accepted: report.accepted,
      ready_for_realtime_axis: report.ready_for_realtime_axis,
      ready_for_speaker_track: report.ready_for_speaker_track,
      ready_for_participant_track: report.ready_for_participant_track,
      recommended_surface: report.recommended_surface,
      candidate_count: report.candidate_count,
      meeting_identity_count: report.meeting_identity_count,
      start_candidate_count: report.start_candidate_count,
      end_candidate_count: report.end_candidate_count,
      active_speaker_count: report.active_speaker_count,
      participant_count: report.participant_count,
      issue_codes: report.issues.map((item) => item.code),
      next_actions: report.next_actions,
    })),
    reports,
    next_actions: uniqueList(reports.flatMap((report) => report.next_actions)),
  };
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
