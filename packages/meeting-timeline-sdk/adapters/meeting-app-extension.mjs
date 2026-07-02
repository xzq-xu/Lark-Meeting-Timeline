import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

export const MEETING_APP_EXTENSION_SCHEMA = 'meeting_app_extension_profile';

export const MEETING_APP_EXTENSION_PLATFORM_KEYS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);

const PLATFORM_ALIASES = Object.freeze({
  google: 'google_meet',
  google_meet: 'google_meet',
  'google-meet': 'google_meet',
  meet: 'google_meet',
  microsoft: 'microsoft_teams',
  microsoft_teams: 'microsoft_teams',
  'microsoft-teams': 'microsoft_teams',
  teams: 'microsoft_teams',
  zoom: 'zoom',
  lark: 'lark',
  feishu: 'lark',
  larksuite: 'lark',
  larkoffice: 'lark',
  webex: 'webex',
});

export const MEETING_APP_EXTENSION_PROFILES = Object.freeze({
  google_meet: Object.freeze({
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platform: 'google_meet',
    display_name: 'Google Meet',
    matches: Object.freeze(['https://meet.google.com/*']),
    host_permissions: Object.freeze(['https://meet.google.com/*']),
  }),
  microsoft_teams: Object.freeze({
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platform: 'microsoft_teams',
    display_name: 'Microsoft Teams',
    matches: Object.freeze([
      'https://teams.microsoft.com/*',
      'https://*.teams.microsoft.com/*',
    ]),
    host_permissions: Object.freeze([
      'https://teams.microsoft.com/*',
      'https://*.teams.microsoft.com/*',
    ]),
  }),
  zoom: Object.freeze({
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platform: 'zoom',
    display_name: 'Zoom',
    matches: Object.freeze([
      'https://zoom.us/*',
      'https://*.zoom.us/*',
      'https://*.zoom.com/*',
    ]),
    host_permissions: Object.freeze([
      'https://zoom.us/*',
      'https://*.zoom.us/*',
      'https://*.zoom.com/*',
    ]),
  }),
  lark: Object.freeze({
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platform: 'lark',
    display_name: 'Lark / Feishu',
    matches: Object.freeze([
      'https://vc.feishu.cn/*',
      'https://*.feishu.cn/*',
      'https://*.larksuite.com/*',
      'https://*.larkoffice.com/*',
    ]),
    host_permissions: Object.freeze([
      'https://vc.feishu.cn/*',
      'https://*.feishu.cn/*',
      'https://*.larksuite.com/*',
      'https://*.larkoffice.com/*',
    ]),
  }),
  webex: Object.freeze({
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platform: 'webex',
    display_name: 'Webex',
    matches: Object.freeze([
      'https://*.webex.com/*',
      'https://*.webex.com.cn/*',
    ]),
    host_permissions: Object.freeze([
      'https://*.webex.com/*',
      'https://*.webex.com.cn/*',
    ]),
  }),
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  const seen = new Set();
  return values.filter((value) => {
    if (value == null || value === '') return false;
    const key = String(value);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function plainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function extensionOptions(platformsOrOptions = {}, options = {}) {
  if (typeof platformsOrOptions === 'string' || Array.isArray(platformsOrOptions)) {
    return { ...options, platforms: platformsOrOptions };
  }
  if (plainObject(platformsOrOptions) && Object.keys(options).length > 0) {
    return { ...platformsOrOptions, ...options };
  }
  return platformsOrOptions ?? {};
}

function platformList(input = {}) {
  const platforms = firstNonEmpty(
    input.platforms,
    input.platform_keys,
    input.platformKeys,
    input.platform,
    input.key,
    MEETING_APP_EXTENSION_PLATFORM_KEYS,
  );
  return unique(asArray(platforms).map((platform) => normalizeMeetingAppExtensionPlatform(platform)));
}

export function normalizeMeetingAppExtensionPlatform(platform) {
  const key = String(platform ?? '').trim().toLowerCase().replaceAll(' ', '_');
  const normalized = PLATFORM_ALIASES[key] ?? PLATFORM_ALIASES[key.replaceAll('_', '-')];
  if (!normalized || !MEETING_APP_EXTENSION_PROFILES[normalized]) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app extension platform: ${platform}`, {
      platform,
      supported_platforms: MEETING_APP_EXTENSION_PLATFORM_KEYS,
    });
  }
  return normalized;
}

export function meetingAppExtensionProfile(platformOrInput, options = {}) {
  const input = plainObject(platformOrInput)
    ? { ...platformOrInput, ...options }
    : { ...options, platform: platformOrInput };
  const platform = normalizeMeetingAppExtensionPlatform(firstNonEmpty(
    input.platform,
    input.key,
    input.name,
  ));
  const profile = MEETING_APP_EXTENSION_PROFILES[platform];
  return {
    ...profile,
    matches: [...profile.matches],
    host_permissions: [...profile.host_permissions],
  };
}

export function buildMeetingAppExtensionMatchPatterns(platformsOrOptions = {}, options = {}) {
  const input = extensionOptions(platformsOrOptions, options);
  const platforms = platformList(input);
  const profiles = platforms.map((platform) => meetingAppExtensionProfile(platform));
  const matches = unique([
    ...profiles.flatMap((profile) => profile.matches),
    ...asArray(firstNonEmpty(input.extraMatches, input.extra_matches)),
  ]);
  const hostPermissions = unique([
    ...profiles.flatMap((profile) => profile.host_permissions),
    ...asArray(firstNonEmpty(input.extraHostPermissions, input.extra_host_permissions)),
  ]);
  const js = unique(asArray(firstNonEmpty(
    input.js,
    input.contentScriptJs,
    input.content_script_js,
    'meeting-app-content-script.bundle.js',
  )));
  const contentScript = compactObject({
    matches,
    js,
    run_at: firstNonEmpty(input.runAt, input.run_at, 'document_idle'),
    all_frames: Boolean(firstNonEmpty(input.allFrames, input.all_frames, false)),
    match_about_blank: firstNonEmpty(input.matchAboutBlank, input.match_about_blank),
  });
  return compactObject({
    type: 'meeting_app_extension_match_patterns',
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platforms,
    profiles,
    matches,
    host_permissions: hostPermissions,
    content_scripts: [contentScript],
  });
}

export function buildMeetingAppContentScriptManifest(options = {}) {
  const patterns = buildMeetingAppExtensionMatchPatterns(options);
  const permissions = unique(asArray(firstNonEmpty(options.permissions, options.extension_permissions, [])));
  const manifestOverrides = options.manifestOverrides ?? options.manifest_overrides ?? {};
  const manifest = compactObject({
    manifest_version: firstNonEmpty(options.manifestVersion, options.manifest_version, 3),
    name: firstNonEmpty(options.name, 'Meeting Timeline Content Bridge'),
    version: firstNonEmpty(options.version, '0.1.0'),
    description: firstNonEmpty(
      options.description,
      'Injects the Meeting Timeline content script into supported meeting web apps.',
    ),
    permissions,
    host_permissions: patterns.host_permissions,
    content_scripts: patterns.content_scripts,
    web_accessible_resources: options.webAccessibleResources ?? options.web_accessible_resources,
    ...manifestOverrides,
  });
  if (manifestOverrides.host_permissions) {
    manifest.host_permissions = unique([
      ...patterns.host_permissions,
      ...asArray(manifestOverrides.host_permissions),
    ]);
  }
  return manifest;
}

export function buildMeetingAppExtensionInstallPlan(options = {}) {
  const patterns = buildMeetingAppExtensionMatchPatterns(options);
  const manifest = options.includeManifest === false || options.include_manifest === false
    ? undefined
    : buildMeetingAppContentScriptManifest(options);
  return compactObject({
    type: 'meeting_app_extension_install_plan',
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    platforms: patterns.platforms,
    matches: patterns.matches,
    host_permissions: patterns.host_permissions,
    content_scripts: patterns.content_scripts,
    manifest,
    content_script_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script',
    browser_runtime_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime',
    snapshot_recorder_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder',
    launch_gate_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate',
    runtime_contract: {
      message_prefixes: ['meeting_timeline', 'meeting-timeline'],
      required_signals: ['meeting_started', 'speaker_started', 'meeting_ended'],
      timestamp_field: 'captured_at_ms',
    },
    next_steps: [
      'Bundle meeting-app-content-script into the js file declared in content_scripts.',
      'Install the generated manifest as a Chrome/Edge compatible MV3 extension or map the same matches into an Electron WebView preload bridge.',
      'Record real meeting app snapshots with meeting-app-snapshot-recorder and validate them with meeting-app-gate before production rollout.',
    ],
    constraints: [
      'Default host permissions are limited to supported meeting web app domains.',
      'The content script does not require realtime transcript access; speaker and lifecycle events can be inferred locally.',
      'Provider webhooks should still be used when available for post-meeting reconciliation.',
    ],
  });
}

export default buildMeetingAppExtensionInstallPlan;
