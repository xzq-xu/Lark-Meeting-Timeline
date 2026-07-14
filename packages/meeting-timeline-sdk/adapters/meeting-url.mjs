import { compactObject } from './internal-utils.mjs';

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

function urlInput(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return String(input);
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
    'window.url',
    'browser.url',
    'tab.url',
  ]);
}

function titleInput(input = {}) {
  if (typeof input === 'string' || input instanceof URL) return undefined;
  return firstPath(input, ['meeting.title', 'meeting.topic', 'meeting.name', 'title', 'topic', 'name', 'window.title', 'tab.title']);
}

function parseUrl(value) {
  if (!value) return null;
  try {
    return new URL(String(value));
  } catch {
    try {
      return new URL(`https://${String(value)}`);
    } catch {
      return null;
    }
  }
}

function cleanId(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/[^a-zA-Z0-9._:@~-]+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function pathSegments(parsed) {
  return parsed.pathname
    .split('/')
    .map((segment) => decodeURIComponent(segment))
    .filter(Boolean);
}

function firstSearchParam(parsed, names = []) {
  return firstNonEmpty(...names.map((name) => parsed.searchParams.get(name)));
}

function hostMatches(hostname, ...suffixes) {
  const host = String(hostname || '').toLowerCase();
  return suffixes.some((suffix) => host === suffix || host.endsWith(`.${suffix}`));
}

function protocolIs(parsed, ...protocols) {
  const protocol = String(parsed.protocol || '').replace(/:$/, '').toLowerCase();
  return protocols.includes(protocol);
}

function stableIdFromUrl(parsed, fallback) {
  const text = parsed
    ? `${parsed.hostname}${parsed.pathname}`.replace(/\/+$/, '')
    : String(fallback || '');
  return cleanId(text.replace(/^www\./i, '')) || null;
}

function googleMeet(parsed) {
  if (!hostMatches(parsed.hostname, 'meet.google.com')) return null;
  const [first] = pathSegments(parsed);
  const code = first && /^[a-z]{3}-?[a-z]{4}-?[a-z]{3}$/i.test(first)
    ? first.toLowerCase()
    : null;
  return {
    platform: 'google_meet',
    meeting_id: code ?? stableIdFromUrl(parsed),
    external_meeting_id: code ?? undefined,
    confidence: code ? 'high' : 'medium',
  };
}

function zoom(parsed) {
  if (!hostMatches(parsed.hostname, 'zoom.us', 'zoomgov.com')) return null;
  const segments = pathSegments(parsed);
  let meetingId = null;
  if (segments[0] === 'j' && segments[1]) meetingId = segments[1];
  if (segments[0] === 'wc' && segments[1]) meetingId = segments[1] === 'join' ? segments[2] : segments[1];
  if (segments[0] === 'my' && segments[1]) meetingId = `my-${segments[1]}`;
  meetingId = firstNonEmpty(
    firstSearchParam(parsed, ['confno', 'meetingid', 'meetingId', 'mn']),
    meetingId,
  );
  return {
    platform: 'zoom',
    meeting_id: cleanId(meetingId) || stableIdFromUrl(parsed),
    external_meeting_id: meetingId ? cleanId(meetingId) : undefined,
    confidence: meetingId ? 'high' : 'medium',
  };
}

function microsoftTeams(parsed) {
  if (!hostMatches(parsed.hostname, 'teams.microsoft.com', 'teams.live.com', 'teams.cloud.microsoft')) return null;
  const segments = pathSegments(parsed);
  let meetingId = null;
  if (segments[0] === 'l' && segments[1] === 'meetup-join' && segments[2]) meetingId = segments[2];
  if (segments[0] === 'meet' && segments[1]) meetingId = segments[1];
  meetingId = firstNonEmpty(
    firstSearchParam(parsed, ['meetingId', 'meetingid', 'threadId', 'threadid', 'conferenceId', 'conversationId']),
    meetingId,
  );
  return {
    platform: 'microsoft_teams',
    meeting_id: cleanId(meetingId) || stableIdFromUrl(parsed),
    external_meeting_id: meetingId ? cleanId(meetingId) : undefined,
    confidence: meetingId ? 'high' : 'medium',
  };
}

function lark(parsed) {
  if (!hostMatches(parsed.hostname, 'vc.feishu.cn', 'vc.larksuite.com', 'larksuite.com', 'feishu.cn') && !protocolIs(parsed, 'lark', 'feishu')) return null;
  const segments = pathSegments(parsed);
  let meetingId = null;
  if (segments[0] === 'j' && segments[1]) meetingId = segments[1];
  if (segments[0] === 'meeting' && segments[1]) meetingId = segments[1];
  meetingId = firstNonEmpty(
    firstSearchParam(parsed, ['meeting_id', 'meetingId', 'meeting_no', 'meetingNo', 'meetingNumber', 'conference_id']),
    meetingId,
  );
  return {
    platform: 'lark',
    meeting_id: cleanId(meetingId) || stableIdFromUrl(parsed),
    external_meeting_id: meetingId ? cleanId(meetingId) : undefined,
    confidence: meetingId ? 'high' : 'medium',
  };
}

function webex(parsed) {
  if (!hostMatches(parsed.hostname, 'webex.com') && !protocolIs(parsed, 'webex')) return null;
  const segments = pathSegments(parsed);
  let meetingId = null;
  if (segments[0] === 'meet' && segments[1]) meetingId = `meet-${segments[1]}`;
  if (segments[0] === 'join' && segments[1]) meetingId = segments[1];
  if (protocolIs(parsed, 'webex') && parsed.hostname === 'meet' && segments[0]) meetingId = `meet-${segments[0]}`;
  if (protocolIs(parsed, 'webex') && parsed.hostname === 'join' && segments[0]) meetingId = segments[0];
  meetingId = firstNonEmpty(
    firstSearchParam(parsed, ['meeting_id', 'meetingId', 'meetingNumber', 'meeting_number']),
    meetingId,
  );
  return {
    platform: 'webex',
    meeting_id: cleanId(meetingId) || stableIdFromUrl(parsed),
    external_meeting_id: meetingId ? cleanId(meetingId) : undefined,
    confidence: meetingId ? 'high' : 'medium',
  };
}

export function stableMeetingIdFromUrl(url) {
  const parsed = parseUrl(url);
  return stableIdFromUrl(parsed, url);
}

export function detectMeetingFromUrl(input = {}) {
  const rawUrl = urlInput(input);
  const parsed = parseUrl(rawUrl);
  if (!parsed) return null;
  const detected = googleMeet(parsed)
    ?? zoom(parsed)
    ?? microsoftTeams(parsed)
    ?? lark(parsed)
    ?? webex(parsed);
  if (!detected?.meeting_id) return null;
  return compactObject({
    ...detected,
    meeting_url: String(rawUrl),
    title: titleInput(input),
  });
}

export function detectMeetingPlatformFromUrl(input = {}) {
  return detectMeetingFromUrl(input)?.platform ?? null;
}
