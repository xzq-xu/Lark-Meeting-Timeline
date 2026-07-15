import { compactObject, normalizeAbsoluteMs } from './internal-utils.mjs';
import { normalizeMeetingAppSnapshot, normalizeMeetingAppSnapshots } from './meeting-apps.mjs';
import { detectMeetingApplication } from './meeting-session-discovery.mjs';
import { normalizeMeetingPlatform } from './platform-setup.mjs';

export const MEETING_APP_DOM_CAPTURE_SCHEMA = 'meeting_app_dom_capture';
export const MEETING_APP_DOM_CAPTURE_SCHEMA_VERSION = 1;

const DEFAULT_CONTROL_SELECTORS = Object.freeze([
  'button',
  '[role="button"]',
  '[title]',
]);

const DEFAULT_PARTICIPANT_SELECTORS = Object.freeze([
  '[data-participant-id]',
  '[data-participantid]',
  '[data-requested-participant-id]',
  '[data-self-name]',
  '[data-participant-name]',
  '[data-user-id]',
  '[data-userid]',
  '[data-speaker-id]',
  '[data-member-id]',
  '[data-attendee-id]',
  '[data-person-id]',
  '[data-display-name]',
  '[data-user-name]',
  '[data-tid*="participant" i]',
  '[data-testid*="participant" i]',
  '[data-testid*="attendee" i]',
  '[class*="participant" i]',
  '[class*="attendee" i]',
  '[class*="speaker" i]',
  '[aria-label*="speaking" i]',
  '[aria-label*="talking" i]',
  '[aria-label*="active speaker" i]',
  '[aria-label*="正在发言"]',
  '[aria-label*="正在讲话"]',
  '[aria-label*="正在说话"]',
]);

const DEFAULT_TEXT_SELECTORS = Object.freeze([
  '[role="status"]',
  '[role="alert"]',
  '[aria-live]',
  '[data-meeting-title]',
  '[data-topic]',
]);

const COMMON_INTERACTION_HINTS = Object.freeze({
  join: Object.freeze([
    /\bjoin now\b/i,
    /\bjoin meeting\b/i,
    /\bask to join\b/i,
    /\bready to join\b/i,
    /\bstart meeting\b/i,
    /立即加入/,
    /加入会议/,
    /申请加入/,
    /开始会议/,
  ]),
  waiting: Object.freeze([
    /\bwaiting room\b/i,
    /\bwaiting to be admitted\b/i,
    /\bwaiting for the host\b/i,
    /等候室/,
    /等待主持人/,
    /等待加入/,
  ]),
  leave: Object.freeze([
    /\bleave call\b/i,
    /\bleave meeting\b/i,
    /\bend call\b/i,
    /\bend meeting\b/i,
    /\bhang up\b/i,
    /离开通话/,
    /退出通话/,
    /离开会议/,
    /结束通话/,
    /结束会议/,
    /挂断/,
  ]),
  microphone: Object.freeze([
    /\bmicrophone\b/i,
    /\bmute\b/i,
    /\bunmute\b/i,
    /\bmic\b/i,
    /麦克风/,
    /静音/,
  ]),
  camera: Object.freeze([
    /\bcamera\b/i,
    /\bvideo\b/i,
    /\bturn camera\b/i,
    /\bstart video\b/i,
    /\bstop video\b/i,
    /摄像头/,
    /视频/,
  ]),
  screen_share: Object.freeze([
    /\bpresent now\b/i,
    /\byou are presenting\b/i,
    /\bstop presenting\b/i,
    /\bshare screen\b/i,
    /\bscreen share\b/i,
    /\bshare content\b/i,
    /正在展示/,
    /停止展示/,
    /共享屏幕/,
    /共享内容/,
  ]),
  screen_share_active: Object.freeze([
    /\byou are presenting\b/i,
    /\bstop presenting\b/i,
    /正在展示/,
    /停止展示/,
  ]),
  captions: Object.freeze([
    /\bcaption\b/i,
    /\bsubtitle\b/i,
    /\bclosed captions\b/i,
    /\bturn on captions\b/i,
    /\bturn off captions\b/i,
    /\bcc\b/i,
    /字幕/,
  ]),
  recording: Object.freeze([
    /\brecording\b/i,
    /\brecorded\b/i,
    /\brecord\b/i,
    /录制/,
    /正在录制/,
  ]),
  participants: Object.freeze([
    /\bparticipants\b/i,
    /\bpeople\b/i,
    /\battendees\b/i,
    /\bmembers\b/i,
    /与会者/,
    /参会人/,
    /参与者/,
    /成员/,
  ]),
  chat: Object.freeze([
    /\bchat\b/i,
    /\bconversation\b/i,
    /聊天/,
    /会议聊天/,
  ]),
  ai_summary: Object.freeze([
    /\bai summary\b/i,
    /\bai notes\b/i,
    /AI 总结/i,
    /AI 视图/i,
  ]),
});

const PLATFORM_INTERACTION_HINTS = Object.freeze({
  google_meet: Object.freeze({
    leave: Object.freeze([/\bleave call\b/i]),
    join: Object.freeze([/\bask to join\b/i, /\bready to join\b/i]),
    screen_share: Object.freeze([/\bpresent now\b/i]),
  }),
  microsoft_teams: Object.freeze({
    leave: Object.freeze([/\bleave\b/i, /\bhang up\b/i, /^离开$/, /^退出$/]),
    screen_share: Object.freeze([/\bshare content\b/i]),
    chat: Object.freeze([/\bshow conversation\b/i]),
  }),
  zoom: Object.freeze({
    leave: Object.freeze([/\bleave meeting\b/i, /\bend meeting\b/i, /^离开$/]),
    participants: Object.freeze([/\bparticipants\b/i, /参会者/]),
  }),
  lark: Object.freeze({
    leave: Object.freeze([/挂断/, /离开会议/]),
    screen_share: Object.freeze([/共享屏幕/]),
    ai_summary: Object.freeze([/AI 总结/i, /AI 视图/i]),
  }),
  webex: Object.freeze({
    leave: Object.freeze([/\bleave meeting\b/i, /\bend meeting\b/i]),
  }),
});

export const MEETING_APP_DOM_CAPTURE_PROFILES = Object.freeze({
  google_meet: Object.freeze({
    platform: 'google_meet',
    displayName: 'Google Meet',
    controlSelectors: Object.freeze([
      '[data-tooltip*="Leave call" i]',
      '[data-tooltip*="Present now" i]',
      '[data-tooltip*="microphone" i]',
      '[aria-label*="Leave call" i]',
      '[aria-label*="Present now" i]',
      '[aria-label*="microphone" i]',
      '[aria-label*="camera" i]',
      '[data-is-muted]',
      '[jscontroller][aria-label]',
    ]),
    participantSelectors: Object.freeze([
      '[data-participant-id]',
      '[data-requested-participant-id]',
      '[data-self-name]',
      '[data-tile-id]',
      '[data-avatar-tooltip]',
      '[data-is-speaking="true"]',
      '[data-speaking="true"]',
      '[aria-label*="is speaking" i]',
      '[aria-label*="speaking" i]',
      '[aria-label*="正在发言"]',
      '[aria-label*="正在讲话"]',
    ]),
    textSelectors: Object.freeze([
      '[data-meeting-title]',
      '[data-call-title]',
      '[aria-live]',
      '[role="status"]',
    ]),
  }),
  microsoft_teams: Object.freeze({
    platform: 'microsoft_teams',
    displayName: 'Microsoft Teams',
    controlSelectors: Object.freeze([
      '[data-tid*="call-" i]',
      '[data-tid*="meeting-" i]',
      '[data-testid*="call-" i]',
      '[aria-label*="Leave" i]',
      '[aria-label*="Hang up" i]',
      '[aria-label*="Share content" i]',
      '[aria-label*="Raise hand" i]',
      '[aria-label*="Mute" i]',
    ]),
    participantSelectors: Object.freeze([
      '[data-tid*="participant" i]',
      '[data-testid*="participant" i]',
      '[data-cid]',
      '[data-user-id]',
      '[data-member-id]',
      '[data-display-name]',
      '[aria-label*="speaking" i]',
      '[aria-label*="muted" i]',
      '[aria-label*="正在发言"]',
    ]),
    textSelectors: Object.freeze([
      '[data-tid*="meeting-title" i]',
      '[data-testid*="meeting-title" i]',
      '[aria-live]',
      '[role="status"]',
    ]),
  }),
  zoom: Object.freeze({
    platform: 'zoom',
    displayName: 'Zoom',
    controlSelectors: Object.freeze([
      '[aria-label*="Leave Meeting" i]',
      '[aria-label*="End Meeting" i]',
      '[aria-label*="Mute" i]',
      '[aria-label*="Unmute" i]',
      '[aria-label*="Participants" i]',
      '[aria-label*="Share Screen" i]',
      '[data-testid*="footer" i]',
      '.footer-button__button',
    ]),
    participantSelectors: Object.freeze([
      '[data-user-id]',
      '[data-participant-id]',
      '[data-participantid]',
      '[data-display-name]',
      '[data-active-speaker="true"]',
      '[data-speaking="true"]',
      '[aria-label*="is speaking" i]',
      '[aria-label*="speaking" i]',
      '.participants-item',
      '.video-avatar__avatar-name',
    ]),
    textSelectors: Object.freeze([
      '[data-testid*="meeting-title" i]',
      '[aria-live]',
      '[role="status"]',
    ]),
  }),
  lark: Object.freeze({
    platform: 'lark',
    displayName: 'Lark/Feishu',
    controlSelectors: Object.freeze([
      '[aria-label*="挂断"]',
      '[aria-label*="离开会议"]',
      '[aria-label*="AI 视图"]',
      '[aria-label*="AI 总结"]',
      '[aria-label*="共享屏幕"]',
      '[aria-label*="麦克风"]',
      '[data-testid*="meeting" i]',
    ]),
    participantSelectors: Object.freeze([
      '[data-user-id]',
      '[data-member-id]',
      '[data-participant-id]',
      '[data-display-name]',
      '[data-uid]',
      '[aria-label*="正在发言"]',
      '[aria-label*="正在讲话"]',
      '[aria-label*="speaking" i]',
    ]),
    textSelectors: Object.freeze([
      '[aria-live]',
      '[role="status"]',
      '[data-testid*="meeting-title" i]',
      '[data-meeting-title]',
    ]),
  }),
  webex: Object.freeze({
    platform: 'webex',
    displayName: 'Cisco Webex',
    controlSelectors: Object.freeze([
      '[aria-label*="Leave meeting" i]',
      '[aria-label*="End meeting" i]',
      '[aria-label*="Mute" i]',
      '[aria-label*="Unmute" i]',
      '[aria-label*="Participants" i]',
      '[aria-label*="Chat" i]',
      '[data-doi*="meeting" i]',
    ]),
    participantSelectors: Object.freeze([
      '[data-person-id]',
      '[data-participant-id]',
      '[data-user-id]',
      '[data-display-name]',
      '[data-active-speaker="true"]',
      '[data-speaking="true"]',
      '[aria-label*="active speaker" i]',
      '[aria-label*="speaking" i]',
    ]),
    textSelectors: Object.freeze([
      '[aria-live]',
      '[role="status"]',
      '[data-testid*="meeting-title" i]',
      '[data-meeting-title]',
    ]),
  }),
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  return Array.isArray(value) ? value : value == null ? [] : [value];
}

function normalizeText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function compactText(value) {
  const text = normalizeText(value);
  return text || undefined;
}

function uniqueStrings(values = []) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const text = compactText(value);
    if (!text || seen.has(text)) continue;
    seen.add(text);
    output.push(text);
  }
  return output;
}

function observedAtMs(input = {}, options = {}) {
  const value = firstNonEmpty(
    options.observedAtMs,
    options.observed_at_ms,
    input.observedAtMs,
    input.observed_at_ms,
    input.timestampMs,
    input.timestamp_ms,
    Date.now(),
  );
  return normalizeAbsoluteMs(value, 'meeting_app_dom_capture_time');
}

function maybeDocument(input = {}) {
  if (input?.querySelectorAll && input?.nodeType === 9) return input;
  if (input?.document?.querySelectorAll) return input.document;
  if (input?.window?.document?.querySelectorAll) return input.window.document;
  if (globalThis.document?.querySelectorAll) return globalThis.document;
  return null;
}

function maybeWindow(input = {}) {
  return input?.window ?? input?.defaultView ?? maybeDocument(input)?.defaultView ?? globalThis.window;
}

function maybeLocation(input = {}) {
  return input?.location
    ?? input?.window?.location
    ?? input?.document?.location
    ?? maybeDocument(input)?.location
    ?? globalThis.location;
}

function locationHref(location) {
  if (!location) return undefined;
  if (typeof location === 'string') return location;
  return String(location.href ?? location.toString?.() ?? '') || undefined;
}

function attr(node, name) {
  if (!node) return undefined;
  if (typeof node.getAttribute === 'function') {
    const value = node.getAttribute(name);
    if (value != null && value !== '') return value;
  }
  const direct = node[name];
  if (direct != null && direct !== '') return direct;
  const camel = name.replace(/-([a-z])/g, (_, char) => char.toUpperCase());
  if (node[camel] != null && node[camel] !== '') return node[camel];
  if (node.attributes && typeof node.attributes === 'object') {
    const value = node.attributes[name] ?? node.attributes[camel];
    if (value && typeof value === 'object' && 'value' in value) return value.value;
    if (value != null && value !== '') return value;
  }
  return undefined;
}

function dataAttr(node, ...names) {
  for (const name of names) {
    const dataKey = name.replace(/-([a-z])/g, (_, char) => char.toUpperCase());
    const value = node?.dataset?.[dataKey]
      ?? node?.dataset?.[name]
      ?? attr(node, `data-${name}`);
    if (value != null && value !== '') return value;
  }
  return undefined;
}

function nodeText(node) {
  return compactText(firstNonEmpty(
    attr(node, 'aria-label'),
    attr(node, 'title'),
    attr(node, 'data-tooltip'),
    attr(node, 'data-avatar-tooltip'),
    attr(node, 'alt'),
    node?.innerText,
    node?.textContent,
    node?.value,
  ));
}

function nodeRole(node) {
  return compactText(attr(node, 'role') ?? node?.role ?? String(node?.tagName ?? '').toLowerCase());
}

function nodeId(node) {
  return compactText(attr(node, 'id') ?? node?.id);
}

function nodeClassName(node) {
  const value = node?.className;
  if (typeof value === 'string') return compactText(value);
  return compactText(attr(node, 'class'));
}

function boolish(value) {
  if (value == null || value === '') return undefined;
  if (typeof value === 'boolean') return value;
  const text = String(value).trim().toLowerCase();
  if (['true', '1', 'yes', 'on'].includes(text)) return true;
  if (['false', '0', 'no', 'off'].includes(text)) return false;
  return undefined;
}

function labelHasSpeakingHint(value) {
  const text = normalizeText(value);
  if (!text) return undefined;
  if (/\b(is speaking|speaking|talking|active speaker)\b/i.test(text)) return true;
  if (/(正在发言|正在讲话|正在说话)/.test(text)) return true;
  if (/\b(muted|microphone off|mic off)\b/i.test(text)) return false;
  if (/(已静音|麦克风已关闭)/.test(text)) return false;
  return undefined;
}

function cleanParticipantLabel(value) {
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

function numberish(value) {
  if (value == null || value === '') return undefined;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : undefined;
}

function optionBool(options = {}, keys = [], fallback = false) {
  for (const key of keys) {
    const parsed = boolish(options[key]);
    if (parsed != null) return parsed;
  }
  return fallback;
}

function optionNumber(options = {}, keys = [], fallback) {
  for (const key of keys) {
    const numeric = numberish(options[key]);
    if (numeric != null) return numeric;
  }
  return fallback;
}

function normalizeProfilePlatform(value) {
  if (!value || typeof value === 'object') return null;
  const text = String(value).trim();
  if (!text || ['auto', 'detect', 'default'].includes(text.toLowerCase())) return null;
  try {
    return normalizeMeetingPlatform(text);
  } catch {
    return null;
  }
}

function profileDisabled(value) {
  if (value === false) return true;
  if (value == null) return false;
  const text = String(value).trim().toLowerCase();
  return ['none', 'off', 'false', 'disabled'].includes(text);
}

function detectCapturePlatform(input = {}, options = {}, url, title) {
  const explicit = firstNonEmpty(
    options.captureProfile,
    options.capture_profile,
    options.platform,
    options.provider,
    input.captureProfile,
    input.capture_profile,
    input.platform,
    input.provider,
    input.meeting_platform,
    input.meetingPlatform,
  );
  if (profileDisabled(explicit)) return null;
  const normalized = normalizeProfilePlatform(explicit);
  if (normalized) return normalized;
  const detected = detectMeetingApplication({
    ...input,
    url,
    title,
    meeting_url: firstNonEmpty(input.meeting_url, input.meetingUrl, url),
    tab: {
      ...(input.tab ?? {}),
      url: firstNonEmpty(input.tab?.url, url),
      title: firstNonEmpty(input.tab?.title, title),
    },
    page: {
      ...(input.page ?? {}),
      url: firstNonEmpty(input.page?.url, url),
      title: firstNonEmpty(input.page?.title, title),
    },
    dom: {
      ...(input.dom ?? {}),
      url: firstNonEmpty(input.dom?.url, url),
      title: firstNonEmpty(input.dom?.title, title),
    },
  });
  return detected?.platform ?? null;
}

export function meetingAppDomCaptureProfile(platformOrInput = {}, options = {}) {
  const input = typeof platformOrInput === 'string'
    ? { platform: platformOrInput }
    : (platformOrInput ?? {});
  const platform = detectCapturePlatform(input, options, options.url ?? input.url, options.title ?? input.title);
  const profile = platform ? MEETING_APP_DOM_CAPTURE_PROFILES[platform] : null;
  if (!profile) return null;
  return {
    platform: profile.platform,
    displayName: profile.displayName,
    controlSelectors: [...profile.controlSelectors],
    participantSelectors: [...profile.participantSelectors],
    textSelectors: [...profile.textSelectors],
  };
}

function queryAll(root, selector) {
  if (!root?.querySelectorAll) return [];
  try {
    return Array.from(root.querySelectorAll(selector));
  } catch {
    return [];
  }
}

function sameOriginDocumentRoots(root, options = {}) {
  const include = optionBool(options, [
    'includeSameOriginFrames',
    'include_same_origin_frames',
  ], false);
  if (!include || !root?.querySelectorAll) return [root];
  const maxDocuments = Math.max(1, optionNumber(options, [
    'maxFrameDocuments',
    'max_frame_documents',
  ], 12));
  const seen = new Set([root]);
  const roots = [root];
  const queue = [root];
  while (queue.length && roots.length < maxDocuments) {
    const current = queue.shift();
    for (const frame of queryAll(current, 'iframe, frame')) {
      let child = null;
      try {
        child = frame.contentDocument;
      } catch {
        child = null;
      }
      if (!child?.querySelectorAll || seen.has(child)) continue;
      seen.add(child);
      roots.push(child);
      queue.push(child);
      if (roots.length >= maxDocuments) break;
    }
  }
  return roots;
}

function shadowRoots(root, options = {}) {
  const include = optionBool(options, [
    'includeShadowDom',
    'include_shadow_dom',
    'deepDom',
    'deep_dom',
  ], false);
  if (!include || !root?.querySelectorAll) return [];
  const maxRoots = Math.max(0, optionNumber(options, ['maxShadowRoots', 'max_shadow_roots'], 20));
  const maxHosts = Math.max(0, optionNumber(options, ['maxShadowHosts', 'max_shadow_hosts'], 400));
  const seen = new Set([root]);
  const roots = [];
  const queue = [root];
  let scannedHosts = 0;
  while (queue.length && roots.length < maxRoots && scannedHosts < maxHosts) {
    const current = queue.shift();
    for (const host of queryAll(current, '*')) {
      scannedHosts += 1;
      const shadow = host?.shadowRoot;
      if (shadow?.querySelectorAll && !seen.has(shadow)) {
        seen.add(shadow);
        roots.push(shadow);
        queue.push(shadow);
        if (roots.length >= maxRoots) break;
      }
      if (scannedHosts >= maxHosts) break;
    }
  }
  return roots;
}

function selectorResults(root, selectors = [], limit = 80, options = {}) {
  if (!root?.querySelectorAll) return [];
  const seen = new Set();
  const results = [];
  const roots = [
    ...(options.__documentRoots ?? [root]),
    ...(options.__shadowRoots ?? shadowRoots(root, options)),
  ];
  for (const selector of selectors) {
    if (!selector || results.length >= limit) break;
    for (const current of roots) {
      for (const node of queryAll(current, selector)) {
        if (!node || seen.has(node)) continue;
        seen.add(node);
        results.push(node);
        if (results.length >= limit) break;
      }
      if (results.length >= limit) break;
    }
  }
  return results;
}

function controlSummary(node) {
  const label = nodeText(node);
  return compactObject({
    tag: String(node?.tagName ?? '').toLowerCase() || undefined,
    role: nodeRole(node),
    id: nodeId(node),
    className: nodeClassName(node),
    ariaLabel: compactText(attr(node, 'aria-label')),
    title: compactText(attr(node, 'title')),
    label,
    text: compactText(node?.innerText ?? node?.textContent),
    disabled: boolish(attr(node, 'disabled') ?? node?.disabled),
  });
}

function participantSummary(node) {
  const label = nodeText(node);
  const explicitSpeaking = boolish(firstNonEmpty(
    dataAttr(node, 'speaking', 'is-speaking', 'is-speaking-now', 'active-speaker', 'is-active-speaker', 'active'),
    attr(node, 'aria-current'),
  ));
  const level = numberish(firstNonEmpty(
    dataAttr(node, 'audio-level', 'voice-activity', 'volume'),
    attr(node, 'aria-valuenow'),
  ));
  const speaking = explicitSpeaking ?? (level == null ? undefined : level >= 0.2) ?? labelHasSpeakingHint(label);
  return compactObject({
    tag: String(node?.tagName ?? '').toLowerCase() || undefined,
    role: nodeRole(node),
    id: firstNonEmpty(
      dataAttr(
        node,
        'participant-id',
        'participantid',
        'requested-participant-id',
        'speaker-id',
        'member-id',
        'attendee-id',
        'person-id',
        'user-id',
        'userid',
        'uid',
        'cid',
        'tile-id',
      ),
      nodeId(node),
    ),
    name: firstNonEmpty(
      dataAttr(
        node,
        'participant-name',
        'self-name',
        'user-name',
        'user-display-name',
        'display-name',
        'displayname',
        'attendee-name',
        'person-name',
        'member-name',
        'avatar-tooltip',
      ),
      cleanParticipantLabel(label),
    ),
    ariaLabel: compactText(attr(node, 'aria-label')),
    label,
    title: compactText(attr(node, 'title')),
    text: compactText(node?.innerText ?? node?.textContent),
    speaking,
    isSpeaking: speaking,
    audioLevel: level,
  });
}

function textSummary(node) {
  return compactObject({
    role: nodeRole(node),
    ariaLive: compactText(attr(node, 'aria-live')),
    label: nodeText(node),
    text: compactText(node?.innerText ?? node?.textContent),
  });
}

function documentVisible(doc) {
  if (!doc) return undefined;
  if (doc.hidden != null) return !doc.hidden;
  if (doc.visibilityState) return doc.visibilityState !== 'hidden';
  return undefined;
}

function browserName(win = {}, options = {}) {
  return firstNonEmpty(
    options.browserName,
    options.browser_name,
    win?.navigator?.userAgentData?.brands?.[0]?.brand,
    win?.navigator?.userAgent,
    globalThis.navigator?.userAgent,
  );
}

function profileHintPatterns(profile, kind) {
  return [
    ...(COMMON_INTERACTION_HINTS[kind] ?? []),
    ...(PLATFORM_INTERACTION_HINTS[profile?.platform]?.[kind] ?? []),
  ];
}

function matchesAnyText(patterns = [], value) {
  const text = normalizeText(value);
  if (!text) return false;
  return patterns.some((pattern) => pattern.test(text));
}

function itemCorpus(item = {}) {
  return uniqueStrings([
    item.label,
    item.text,
    item.ariaLabel,
    item.aria_label,
    item.title,
    item.name,
    item.role,
  ]);
}

function itemMatches(item = {}, patterns = []) {
  return itemCorpus(item).some((text) => matchesAnyText(patterns, text));
}

function semanticSignal(type, source, item = {}, extra = {}) {
  return compactObject({
    type,
    source,
    confidence: extra.confidence ?? 'heuristic',
    label: item.label,
    text: item.text,
    role: item.role,
    id: item.id,
    participant_id: extra.participant_id,
    participant_name: extra.participant_name,
  });
}

function firstActiveParticipant(participants = []) {
  return participants.find((participant) => (
    participant.speaking === true
    || participant.isSpeaking === true
    || participant.active_speaker === true
    || participant.activeSpeaker === true
  ));
}

function buildSemanticSignals({ controls = [], participants = [], texts = [], profile = null } = {}) {
  const signals = [];
  const controlClassifiers = [
    ['meeting_join_available', 'control', 'join'],
    ['meeting_waiting_room', 'control', 'waiting'],
    ['meeting_leave_available', 'control', 'leave'],
    ['microphone_control', 'control', 'microphone'],
    ['camera_control', 'control', 'camera'],
    ['screen_share_control', 'control', 'screen_share'],
    ['captions_control', 'control', 'captions'],
    ['recording_control', 'control', 'recording'],
    ['participants_control', 'control', 'participants'],
    ['chat_control', 'control', 'chat'],
    ['ai_summary_control', 'control', 'ai_summary'],
  ];
  for (const control of controls) {
    for (const [type, source, kind] of controlClassifiers) {
      if (itemMatches(control, profileHintPatterns(profile, kind))) {
        signals.push(semanticSignal(type, source, control));
      }
    }
  }

  const textClassifiers = [
    ['meeting_join_available', 'status_text', 'join'],
    ['meeting_waiting_room', 'status_text', 'waiting'],
    ['meeting_leave_available', 'status_text', 'leave'],
    ['screen_share_active', 'status_text', 'screen_share_active'],
    ['captions_status', 'status_text', 'captions'],
    ['recording_indicator', 'status_text', 'recording'],
    ['ai_summary_status', 'status_text', 'ai_summary'],
  ];
  for (const text of texts) {
    for (const [type, source, kind] of textClassifiers) {
      if (itemMatches(text, profileHintPatterns(profile, kind))) {
        signals.push(semanticSignal(type, source, text));
      }
    }
  }

  for (const participant of participants) {
    if (participant.speaking === true || participant.isSpeaking === true) {
      signals.push(semanticSignal('active_speaker_candidate', 'participant_tile', participant, {
        confidence: participant.audioLevel == null ? 'heuristic' : 'audio_or_dom',
        participant_id: participant.id,
        participant_name: participant.name,
      }));
    }
  }
  if (participants.length > 0) {
    signals.push(compactObject({
      type: 'participant_roster_observed',
      source: 'participant_tiles',
      confidence: 'heuristic',
      participant_count: participants.length,
    }));
  }
  return signals;
}

function signalTypes(signals = []) {
  return new Set(signals.map((signal) => signal.type).filter(Boolean));
}

function semanticSignalTypeList(signals = []) {
  return uniqueStrings(signals.map((signal) => signal.type));
}

function buildControlSignalSummary(signals = []) {
  const types = signalTypes(signals);
  const signalTypeList = semanticSignalTypeList(signals);
  return compactObject({
    signal_types: signalTypeList.length ? signalTypeList : undefined,
    signal_count: signals.length || undefined,
    join_available: types.has('meeting_join_available') || undefined,
    waiting_room: types.has('meeting_waiting_room') || undefined,
    leave_available: types.has('meeting_leave_available') || undefined,
    microphone_available: types.has('microphone_control') || undefined,
    camera_available: types.has('camera_control') || undefined,
    screen_share_available: types.has('screen_share_control') || undefined,
    screen_share_active: types.has('screen_share_active') || undefined,
    captions_available: types.has('captions_control') || types.has('captions_status') || undefined,
    recording_observed: types.has('recording_indicator') || types.has('recording_control') || undefined,
    participants_available: types.has('participants_control') || types.has('participant_roster_observed') || undefined,
    chat_available: types.has('chat_control') || undefined,
    ai_summary_available: types.has('ai_summary_control') || types.has('ai_summary_status') || undefined,
    active_speaker_observed: types.has('active_speaker_candidate') || undefined,
    participant_roster_observed: types.has('participant_roster_observed') || undefined,
  });
}

function buildInteractionState(signals = [], participants = []) {
  const types = signalTypes(signals);
  const activeParticipant = firstActiveParticipant(participants);
  const hasLeave = types.has('meeting_leave_available');
  const hasJoin = types.has('meeting_join_available');
  const hasWaiting = types.has('meeting_waiting_room');
  const inCall = hasLeave || Boolean(activeParticipant) || types.has('screen_share_active');
  const preJoin = !inCall && (hasJoin || hasWaiting);
  return compactObject({
    in_call: inCall === true ? true : undefined,
    pre_join: preJoin === true ? true : undefined,
    can_join: hasJoin === true ? true : undefined,
    waiting_room: hasWaiting === true ? true : undefined,
    can_leave: hasLeave === true ? true : undefined,
    microphone_control_available: types.has('microphone_control') || undefined,
    camera_control_available: types.has('camera_control') || undefined,
    screen_share_available: types.has('screen_share_control') || undefined,
    screen_share_active: types.has('screen_share_active') || undefined,
    captions_available: types.has('captions_control') || types.has('captions_status') || undefined,
    recording_observed: types.has('recording_indicator') || types.has('recording_control') || undefined,
    participant_roster_observed: types.has('participant_roster_observed') || undefined,
    chat_available: types.has('chat_control') || undefined,
    ai_summary_available: types.has('ai_summary_control') || types.has('ai_summary_status') || undefined,
    active_speaker_candidate: activeParticipant ? compactObject({
      id: activeParticipant.id,
      name: activeParticipant.name,
      display_name: activeParticipant.name,
      speaking: true,
      audioLevel: activeParticipant.audioLevel,
    }) : undefined,
    participant_count: participants.length || undefined,
  });
}

export function captureMeetingAppDomSnapshot(input = {}, options = {}) {
  const doc = maybeDocument(input);
  const win = maybeWindow(input);
  const location = maybeLocation(input);
  const atMs = observedAtMs(input, options);
  const controlLimit = Number(options.maxControls ?? options.max_controls ?? 80);
  const participantLimit = Number(options.maxParticipants ?? options.max_participants ?? 80);
  const textLimit = Number(options.maxTexts ?? options.max_texts ?? 40);
  const url = firstNonEmpty(options.url, input.url, locationHref(location));
  const title = firstNonEmpty(options.title, input.title, doc?.title);
  const profile = meetingAppDomCaptureProfile({
    ...input,
    url,
    title,
  }, options);
  const capturedDocuments = sameOriginDocumentRoots(doc, options);
  const capturedShadowRoots = capturedDocuments.flatMap((root) => shadowRoots(root, options));
  const selectorOptions = {
    ...options,
    __documentRoots: capturedDocuments,
    __shadowRoots: capturedShadowRoots,
  };
  const controls = selectorResults(doc, [
    ...uniqueStrings([
      ...(options.controlSelectors ?? options.control_selectors ?? []),
      ...(profile?.controlSelectors ?? []),
      ...DEFAULT_CONTROL_SELECTORS,
    ]),
  ], controlLimit, selectorOptions).map(controlSummary).filter((item) => item.label || item.ariaLabel || item.title);
  const participants = selectorResults(doc, [
    ...uniqueStrings([
      ...(options.participantSelectors ?? options.participant_selectors ?? []),
      ...(profile?.participantSelectors ?? []),
      ...DEFAULT_PARTICIPANT_SELECTORS,
    ]),
  ], participantLimit, selectorOptions).map(participantSummary).filter((item) => item.id || item.name || item.label || item.audioLevel != null);
  const texts = selectorResults(doc, [
    ...uniqueStrings([
      ...(options.textSelectors ?? options.text_selectors ?? []),
      ...(profile?.textSelectors ?? []),
      ...DEFAULT_TEXT_SELECTORS,
    ]),
  ], textLimit, selectorOptions).map(textSummary).filter((item) => item.label || item.text);
  const semanticSignals = buildSemanticSignals({
    controls,
    participants,
    texts,
    profile,
  });
  const semanticSignalTypes = semanticSignalTypeList(semanticSignals);
  const controlSignalSummary = buildControlSignalSummary(semanticSignals);
  const interaction = buildInteractionState(semanticSignals, participants);
  const activeSpeaker = interaction.active_speaker_candidate;
  const inferredInMeeting = interaction.in_call === true ? true : interaction.pre_join === true ? false : undefined;

  return compactObject({
    schema: MEETING_APP_DOM_CAPTURE_SCHEMA,
    schema_version: MEETING_APP_DOM_CAPTURE_SCHEMA_VERSION,
    source: options.source ?? 'browser_dom_capture',
    observedAtMs: atMs,
    url,
    title,
    inMeeting: inferredInMeeting,
    activeSpeaker,
    interaction,
    semanticSignals,
    semanticSignalTypes,
    controlSignalSummary,
    browser: {
      name: browserName(win, options),
    },
    page: {
      url,
      title,
      documentVisible: documentVisible(doc),
      inMeeting: inferredInMeeting,
      activeSpeaker,
      interaction,
      semanticSignals,
      semanticSignalTypes,
      controlSignalSummary,
      buttons: controls,
      controls,
      tiles: participants,
      participants,
      texts,
    },
    dom: {
      url,
      title,
      inMeeting: inferredInMeeting,
      activeSpeaker,
      interaction,
      semanticSignals,
      semanticSignalTypes,
      controlSignalSummary,
      controls,
      participants,
      texts,
    },
    capture: {
      profile: profile?.platform,
      profile_display_name: profile?.displayName,
      frame_document_count: capturedDocuments.length > 1 ? capturedDocuments.length - 1 : undefined,
      shadow_root_count: capturedShadowRoots.length || undefined,
      control_count: controls.length,
      participant_count: participants.length,
      text_count: texts.length,
      semantic_signal_count: semanticSignals.length,
      semantic_signal_types: semanticSignalTypes,
      control_signal_summary: controlSignalSummary,
    },
  });
}

export function normalizeCapturedMeetingAppDomSnapshot(input = {}, options = {}) {
  return normalizeMeetingAppSnapshot(captureMeetingAppDomSnapshot(input, options), options);
}

export function normalizeCapturedMeetingAppDomSnapshots(inputs = [], options = {}) {
  return normalizeMeetingAppSnapshots(asArray(inputs).map((item) => captureMeetingAppDomSnapshot(item, options)), options);
}
