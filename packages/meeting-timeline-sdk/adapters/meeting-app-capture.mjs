import { compactObject, normalizeAbsoluteMs } from '../index.mjs';
import { normalizeMeetingAppSnapshot, normalizeMeetingAppSnapshots } from './meeting-apps.mjs';

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
  '[data-tid*="participant" i]',
  '[class*="participant" i]',
  '[aria-label*="speaking" i]',
  '[aria-label*="talking" i]',
  '[aria-label*="正在发言"]',
  '[aria-label*="正在讲话"]',
]);

const DEFAULT_TEXT_SELECTORS = Object.freeze([
  '[role="status"]',
  '[role="alert"]',
  '[aria-live]',
  '[data-meeting-title]',
  '[data-topic]',
]);

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

function selectorResults(root, selectors = [], limit = 80) {
  if (!root?.querySelectorAll) return [];
  const seen = new Set();
  const results = [];
  for (const selector of selectors) {
    if (!selector || results.length >= limit) break;
    let nodes = [];
    try {
      nodes = Array.from(root.querySelectorAll(selector));
    } catch {
      nodes = [];
    }
    for (const node of nodes) {
      if (!node || seen.has(node)) continue;
      seen.add(node);
      results.push(node);
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
    dataAttr(node, 'speaking', 'is-speaking', 'active-speaker'),
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
      dataAttr(node, 'participant-id', 'participantid', 'requested-participant-id', 'user-id', 'userid'),
      nodeId(node),
    ),
    name: firstNonEmpty(
      dataAttr(node, 'participant-name', 'self-name', 'user-name'),
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
  const controls = selectorResults(doc, [
    ...(options.controlSelectors ?? options.control_selectors ?? []),
    ...DEFAULT_CONTROL_SELECTORS,
  ], controlLimit).map(controlSummary).filter((item) => item.label || item.ariaLabel || item.title);
  const participants = selectorResults(doc, [
    ...(options.participantSelectors ?? options.participant_selectors ?? []),
    ...DEFAULT_PARTICIPANT_SELECTORS,
  ], participantLimit).map(participantSummary).filter((item) => item.id || item.name || item.label || item.audioLevel != null);
  const texts = selectorResults(doc, [
    ...(options.textSelectors ?? options.text_selectors ?? []),
    ...DEFAULT_TEXT_SELECTORS,
  ], textLimit).map(textSummary).filter((item) => item.label || item.text);
  const url = firstNonEmpty(options.url, input.url, locationHref(location));
  const title = firstNonEmpty(options.title, input.title, doc?.title);

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
