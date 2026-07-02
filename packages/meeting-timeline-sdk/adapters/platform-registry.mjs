import { normalizeGoogleMeetEvent } from './google-meet.mjs';
import { normalizeLarkEvent } from './lark.mjs';
import { normalizeLocalDetectorEvent } from './local-detector.mjs';
import { normalizeMicrosoftTeamsEvent } from './microsoft-teams.mjs';
import {
  MEETING_PLATFORM_ALIASES,
  MEETING_PLATFORM_KEYS,
  normalizeMeetingPlatform,
} from './platform-setup.mjs';
import { normalizeWebexEvent } from './webex.mjs';
import { normalizeZoomEvent } from './zoom.mjs';

const normalizers = Object.freeze({
  local_detector: normalizeLocalDetectorEvent,
  lark: normalizeLarkEvent,
  google_meet: normalizeGoogleMeetEvent,
  microsoft_teams: normalizeMicrosoftTeamsEvent,
  zoom: normalizeZoomEvent,
  webex: normalizeWebexEvent,
});

const sourceByPlatform = Object.freeze({
  local_detector: 'local_detector',
});

function aliasesFor(platform) {
  return Object.entries(MEETING_PLATFORM_ALIASES)
    .filter(([, key]) => key === platform)
    .map(([alias]) => alias);
}

export const MEETING_PLATFORM_EVENT_ADAPTERS = Object.freeze(
  MEETING_PLATFORM_KEYS.map((platform) => Object.freeze({
    key: platform,
    aliases: Object.freeze(aliasesFor(platform)),
    source: sourceByPlatform[platform] ?? `${platform}_webhook`,
    normalize: normalizers[platform],
  })),
);

const adapterByAlias = new Map(
  MEETING_PLATFORM_EVENT_ADAPTERS.flatMap((adapter) => (
    adapter.aliases.map((alias) => [alias, adapter])
  )),
);

export function meetingPlatformEventAdapterFor(platform) {
  let normalized;
  try {
    normalized = normalizeMeetingPlatform(platform);
  } catch {
    return null;
  }
  return adapterByAlias.get(normalized) ?? MEETING_PLATFORM_EVENT_ADAPTERS.find((adapter) => adapter.key === normalized) ?? null;
}
