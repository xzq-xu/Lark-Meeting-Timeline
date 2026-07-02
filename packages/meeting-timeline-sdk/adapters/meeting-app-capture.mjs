import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
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
  const roots = [root, ...(options.__shadowRoots ?? shadowRoots(root, options))];
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
  const speaking = boolish(firstNonEmpty(
    dataAttr(node, 'speaking', 'is-speaking', 'is-speaking-now', 'active-speaker', 'is-active-speaker', 'active'),
    attr(node, 'aria-current'),
  ));
  const level = numberish(firstNonEmpty(
    dataAttr(node, 'audio-level', 'voice-activity', 'volume'),
    attr(node, 'aria-valuenow'),
  ));
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
      label,
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
  const capturedShadowRoots = shadowRoots(doc, options);
  const selectorOptions = {
    ...options,
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

  return compactObject({
    schema: MEETING_APP_DOM_CAPTURE_SCHEMA,
    schema_version: MEETING_APP_DOM_CAPTURE_SCHEMA_VERSION,
    source: options.source ?? 'browser_dom_capture',
    observedAtMs: atMs,
    url,
    title,
    browser: {
      name: browserName(win, options),
    },
    page: {
      url,
      title,
      documentVisible: documentVisible(doc),
      buttons: controls,
      controls,
      tiles: participants,
      participants,
      texts,
    },
    dom: {
      url,
      title,
      controls,
      participants,
      texts,
    },
    capture: {
      profile: profile?.platform,
      profile_display_name: profile?.displayName,
      shadow_root_count: capturedShadowRoots.length || undefined,
      control_count: controls.length,
      participant_count: participants.length,
      text_count: texts.length,
    },
  });
}

export function normalizeCapturedMeetingAppDomSnapshot(input = {}, options = {}) {
  return normalizeMeetingAppSnapshot(captureMeetingAppDomSnapshot(input, options), options);
}

export function normalizeCapturedMeetingAppDomSnapshots(inputs = [], options = {}) {
  return normalizeMeetingAppSnapshots(asArray(inputs).map((item) => captureMeetingAppDomSnapshot(item, options)), options);
}
