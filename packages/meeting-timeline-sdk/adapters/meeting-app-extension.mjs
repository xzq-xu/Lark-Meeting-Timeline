import { MeetingTimelineSdkError, compactObject } from '../index.mjs';

export const MEETING_APP_EXTENSION_SCHEMA = 'meeting_app_extension_profile';

export const MEETING_APP_EXTENSION_PLATFORM_KEYS = Object.freeze([
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);

export const MEETING_APP_EXTENSION_MESSAGE_TYPES = Object.freeze({
  client_call: 'meeting_timeline.client_call',
  extension_attached: 'meeting_timeline.extension_attached',
  extension_status: 'meeting_timeline.extension_status',
  observe_candidates: 'meeting_timeline.observe_candidates',
  preflight_current_window: 'meeting_timeline.preflight_current_window',
  preflight_candidates: 'meeting_timeline.preflight_candidates',
  candidate_launch_plan: 'meeting_timeline.candidate_launch_plan',
  open_candidate_session: 'meeting_timeline.open_candidate_session',
});

export const MEETING_APP_EXTENSION_STATUS_STORAGE_KEY = 'meeting_timeline_extension_status';

export const MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS = Object.freeze({
  startMeeting: '/api/meeting-session/start',
  endMeeting: '/api/meeting-session/end',
  insertMark: '/api/annotations',
  insertMarks: '/api/annotations/batch',
  importTranscript: '/api/import/transcript',
});

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

const MESSAGE_TYPE_ALIASES = Object.freeze({
  client_call: MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  clientcall: MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  'client-call': MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  attached: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
  extension_attached: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
  extensionattached: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
  'extension-attached': MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
  status: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
  extension_status: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
  extensionstatus: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
  'extension-status': MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
  observe_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  observecandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  'observe-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  observe_platform_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  observeplatformcandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  'observe-platform-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  preflight_current_window: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  preflightcurrentwindow: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  'preflight-current-window': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  current_window_preflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  currentwindowpreflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  'current-window-preflight': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  adapter_preflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  adapterpreflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  'adapter-preflight': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  preflight_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  preflightcandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  'preflight-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  candidates_preflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  candidatespreflight: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  'candidates-preflight': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  preflight_platform_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  preflightplatformcandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  'preflight-platform-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  candidate_launch_plan: MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  candidatelaunchplan: MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  'candidate-launch-plan': MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  launch_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  launchcandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  'launch-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan,
  open_candidate_session: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
  opencandidatesession: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
  'open-candidate-session': MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
  open_candidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
  opencandidates: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
  'open-candidates': MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session,
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

function json(value) {
  return JSON.stringify(value, null, 2);
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

function endpointOriginPattern(value) {
  if (!value) return undefined;
  try {
    const url = new URL(String(value));
    return `${url.protocol}//${url.host}/*`;
  } catch {
    return undefined;
  }
}

function matchPatternHost(pattern) {
  const withoutProtocol = String(pattern).replace(/^[a-z]+:\/\//i, '');
  const host = withoutProtocol.split('/')[0];
  return host.replace(/^\*\./, '').replace(/^\*$/, '');
}

function sourceFile(path, content, role, mime = 'text/plain') {
  return {
    path,
    role,
    mime,
    content: content.endsWith('\n') ? content : `${content}\n`,
  };
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function fileByPath(scaffold = {}, path) {
  return asArray(scaffold.files).find((file) => file?.path === path);
}

function jsonFromFile(scaffold = {}, path, issues = []) {
  const file = fileByPath(scaffold, path);
  if (!file) {
    issues.push(issue('error', `missing_${path.replace(/[^a-z0-9]/gi, '_')}`, `Missing ${path}.`));
    return undefined;
  }
  try {
    return JSON.parse(file.content);
  } catch (error) {
    issues.push(issue('error', `invalid_${path.replace(/[^a-z0-9]/gi, '_')}`, `${path} is not valid JSON.`, {
      error: String(error?.message ?? error),
    }));
    return undefined;
  }
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

function optionalPlatform(platform) {
  if (platform == null || platform === '') return undefined;
  const value = String(platform).trim();
  if (!value) return undefined;
  if (value.toLowerCase() === 'unknown') return 'unknown';
  return normalizeMeetingAppExtensionPlatform(value);
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

export function normalizeMeetingAppExtensionMessageType(messageType) {
  const raw = String(messageType ?? '').trim();
  if (Object.values(MEETING_APP_EXTENSION_MESSAGE_TYPES).includes(raw)) return raw;
  const key = raw.toLowerCase().replaceAll(' ', '_');
  const normalized = MESSAGE_TYPE_ALIASES[key]
    ?? MESSAGE_TYPE_ALIASES[key.replaceAll('_', '-')]
    ?? MESSAGE_TYPE_ALIASES[key.replaceAll('_', '')];
  if (!normalized) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app extension message type: ${messageType}`, {
      message_type: messageType,
      supported_message_types: MEETING_APP_EXTENSION_MESSAGE_TYPES,
    });
  }
  return normalized;
}

export function meetingAppExtensionTimelineEndpoint(method) {
  const key = String(method ?? '').trim();
  const endpoint = MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS[key];
  if (!endpoint) {
    throw new MeetingTimelineSdkError(`Unsupported meeting app extension client call method: ${method}`, {
      method,
      supported_methods: Object.keys(MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS),
    });
  }
  return endpoint;
}

export function buildMeetingAppExtensionAttachedMessage(input = {}, options = {}) {
  const merged = plainObject(input)
    ? { ...input, ...options }
    : { ...options, platform: input };
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'extension_attached',
    )),
    platform: optionalPlatform(firstNonEmpty(merged.platform, merged.platform_key, merged.platformKey)),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    url: firstNonEmpty(merged.url, merged.href, merged.meeting_url, merged.meetingUrl),
  });
}

export function buildMeetingAppExtensionStatusMessage(input = {}, options = {}) {
  const merged = plainObject(input) ? { ...input, ...options } : options;
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'extension_status',
    )),
    request_id: firstNonEmpty(merged.request_id, merged.requestId),
    platform: optionalPlatform(firstNonEmpty(merged.platform, merged.platform_key, merged.platformKey)),
    url: firstNonEmpty(merged.url, merged.href, merged.meeting_url, merged.meetingUrl),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs),
  });
}

export function buildMeetingAppExtensionObserveCandidatesMessage(input = {}, options = {}) {
  const merged = plainObject(input) ? { ...input, ...options } : options;
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'observe_candidates',
    )),
    request_id: firstNonEmpty(merged.request_id, merged.requestId),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    query: firstNonEmpty(merged.query, merged.tabs_query, merged.tabsQuery),
    tabs: firstNonEmpty(merged.tabs, merged.browser_tabs, merged.browserTabs),
    windows: firstNonEmpty(merged.windows, merged.browser_windows, merged.browserWindows),
    applications: firstNonEmpty(merged.applications, merged.apps),
    candidates: merged.candidates,
    environment: firstNonEmpty(merged.environment, merged.meeting_environment, merged.meetingEnvironment),
    input: plainObject(merged.input) ? merged.input : undefined,
  });
}

export function buildMeetingAppExtensionCurrentWindowPreflightMessage(input = {}, options = {}) {
  const merged = plainObject(input) ? { ...input, ...options } : options;
  const preflightOptions = {
    ...(plainObject(merged.options) ? merged.options : {}),
    ...(plainObject(merged.preflightOptions) ? merged.preflightOptions : {}),
    ...(plainObject(merged.preflight_options) ? merged.preflight_options : {}),
  };
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'preflight_current_window',
    )),
    request_id: firstNonEmpty(merged.request_id, merged.requestId),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    platform: optionalPlatform(firstNonEmpty(merged.platform, merged.platform_key, merged.platformKey)),
    url: firstNonEmpty(merged.url, merged.href, merged.meeting_url, merged.meetingUrl),
    title: merged.title,
    input: plainObject(merged.input) ? merged.input : undefined,
    options: Object.keys(preflightOptions).length > 0 ? preflightOptions : undefined,
  });
}

export function buildMeetingAppExtensionPreflightCandidatesMessage(input = {}, options = {}) {
  const merged = plainObject(input) ? { ...input, ...options } : options;
  const preflightOptions = {
    ...(plainObject(merged.options) ? merged.options : {}),
    ...(plainObject(merged.preflightOptions) ? merged.preflightOptions : {}),
    ...(plainObject(merged.preflight_options) ? merged.preflight_options : {}),
  };
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'preflight_candidates',
    )),
    request_id: firstNonEmpty(merged.request_id, merged.requestId),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    query: firstNonEmpty(merged.query, merged.tabs_query, merged.tabsQuery),
    tabs: firstNonEmpty(merged.tabs, merged.browser_tabs, merged.browserTabs),
    input: plainObject(merged.input) ? merged.input : undefined,
    options: Object.keys(preflightOptions).length > 0 ? preflightOptions : undefined,
  });
}

function candidateLaunchOptions(merged = {}) {
  const launchOptions = {
    ...(plainObject(merged.options) ? merged.options : {}),
    ...(plainObject(merged.launchOptions) ? merged.launchOptions : {}),
    ...(plainObject(merged.launch_options) ? merged.launch_options : {}),
    ...(plainObject(merged.preflightOptions) ? merged.preflightOptions : {}),
    ...(plainObject(merged.preflight_options) ? merged.preflight_options : {}),
  };
  return Object.keys(launchOptions).length > 0 ? launchOptions : undefined;
}

function buildCandidateLaunchMessage(input = {}, options = {}, defaultType) {
  const merged = plainObject(input) ? { ...input, ...options } : options;
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      defaultType,
    )),
    request_id: firstNonEmpty(merged.request_id, merged.requestId),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    platform: optionalPlatform(firstNonEmpty(merged.platform, merged.platform_key, merged.platformKey)),
    url: firstNonEmpty(merged.url, merged.href, merged.meeting_url, merged.meetingUrl),
    title: merged.title,
    query: firstNonEmpty(merged.query, merged.tabs_query, merged.tabsQuery),
    tabs: firstNonEmpty(merged.tabs, merged.browser_tabs, merged.browserTabs),
    windows: firstNonEmpty(merged.windows, merged.browser_windows, merged.browserWindows),
    candidates: merged.candidates,
    input: plainObject(merged.input) ? merged.input : undefined,
    options: candidateLaunchOptions(merged),
  });
}

export function buildMeetingAppExtensionCandidateLaunchPlanMessage(input = {}, options = {}) {
  return buildCandidateLaunchMessage(input, options, 'candidate_launch_plan');
}

export function buildMeetingAppExtensionOpenCandidateSessionMessage(input = {}, options = {}) {
  return buildCandidateLaunchMessage(input, options, 'open_candidate_session');
}

export function buildMeetingAppExtensionClientCallMessage(methodOrInput, input = {}, options = {}) {
  const objectInput = plainObject(methodOrInput);
  const merged = objectInput
    ? { ...methodOrInput, ...(Object.keys(options).length > 0 ? options : input) }
    : { ...options, method: methodOrInput, input };
  const method = firstNonEmpty(merged.method, merged.action);
  meetingAppExtensionTimelineEndpoint(method);
  const payload = plainObject(merged.input) ? merged.input : {};
  return compactObject({
    type: normalizeMeetingAppExtensionMessageType(firstNonEmpty(
      merged.type,
      merged.messageType,
      merged.message_type,
      'client_call',
    )),
    method,
    platform: optionalPlatform(firstNonEmpty(merged.platform, merged.platform_key, merged.platformKey)),
    captured_at_ms: firstNonEmpty(merged.captured_at_ms, merged.capturedAtMs, Date.now()),
    url: firstNonEmpty(merged.url, merged.href, merged.meeting_url, merged.meetingUrl),
    input: payload,
  });
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
    content_script_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
    meeting_app_content_script_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script',
    platform_integration_runtime_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime',
    browser_runtime_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime',
    snapshot_recorder_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder',
    launch_gate_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate',
    runtime_contract: {
      message_prefixes: ['meeting_timeline', 'meeting-timeline'],
      message_types: { ...MEETING_APP_EXTENSION_MESSAGE_TYPES },
      local_content_script_messages: [
        MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
        'meeting_timeline.insert_mark',
        'meeting_timeline.sample_tracks',
      ],
      status_storage_key: MEETING_APP_EXTENSION_STATUS_STORAGE_KEY,
      timeline_endpoints: { ...MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS },
      live_capture_global: '__meetingTimelineLiveCapture',
      live_capture_methods: ['captureActive', 'captureEnded', 'exportRecords', 'evidencePackage', 'diagnose'],
      required_signals: ['meeting_started', 'speaker_started', 'meeting_ended'],
      timestamp_field: 'captured_at_ms',
    },
    next_steps: [
      'Bundle platform-integration-runtime content script bridge into the js file declared in content_scripts.',
      'Install the generated manifest as a Chrome/Edge compatible MV3 extension or map the same matches into an Electron WebView preload bridge.',
      'Use window.__meetingTimelineLiveCapture.captureActive(), captureEnded(), and diagnose() to collect and validate live DOM evidence for the launch gate.',
      'Record real meeting app snapshots with meeting-app-snapshot-recorder and validate them with meeting-app-gate before production rollout.',
    ],
    constraints: [
      'Default host permissions are limited to supported meeting web app domains.',
      'The content script does not require realtime transcript access; speaker and lifecycle events can be inferred locally.',
      'Provider webhooks should still be used when available for post-meeting reconciliation.',
    ],
  });
}

export function buildMeetingAppExtensionContentScriptSource(options = {}) {
  const installPlan = buildMeetingAppExtensionInstallPlan(options);
  const platformMap = Object.fromEntries(installPlan.platforms.map((platform) => {
    const profile = MEETING_APP_EXTENSION_PROFILES[platform];
    return [platform, profile.host_permissions.map((pattern) => matchPatternHost(pattern)).filter(Boolean)];
  }));
  const messagePrefix = firstNonEmpty(
    options.messagePrefix,
    options.message_prefix,
    MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  );
  return [
    "import { installMeetingPlatformIntegrationContentScriptBridge } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';",
    '',
    `const PLATFORM_HOSTS = ${json(platformMap)};`,
    `const CLIENT_CALL_TYPE = ${JSON.stringify(messagePrefix)};`,
    `const ATTACHED_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached)};`,
    '',
    'function extensionRuntime() {',
    '  return globalThis.chrome?.runtime ?? globalThis.browser?.runtime;',
    '}',
    '',
    'function inferPlatform(hostname = globalThis.location?.hostname ?? "") {',
    '  const host = String(hostname);',
    '  for (const [platform, hosts] of Object.entries(PLATFORM_HOSTS)) {',
    '    if (hosts.some((item) => host === item || host.endsWith(`.${item}`))) return platform;',
    '  }',
    '  return "unknown";',
    '}',
    '',
    'function callTimeline(method, input = {}) {',
    '  const runtime = extensionRuntime();',
    '  if (!runtime?.sendMessage) return Promise.resolve({ ok: false, reason: "missing_extension_runtime" });',
    '  const message = {',
    '    type: CLIENT_CALL_TYPE,',
    '    method,',
    '    platform: inferPlatform(),',
    '    captured_at_ms: Date.now(),',
    '    url: globalThis.location?.href,',
    '    input,',
    '  };',
    '  return new Promise((resolve) => {',
    '    runtime.sendMessage(message, (response) => {',
    '      const runtimeError = globalThis.chrome?.runtime?.lastError;',
    '      if (runtimeError) resolve({ ok: false, reason: "runtime_error", error: runtimeError.message });',
    '      else resolve(response ?? { ok: true });',
    '    });',
    '  });',
    '}',
    '',
    'const client = {',
    '  startMeeting: (input) => callTimeline("startMeeting", input),',
    '  endMeeting: (input) => callTimeline("endMeeting", input),',
    '  insertMark: (input) => callTimeline("insertMark", input),',
    '  insertMarks: (input) => callTimeline("insertMarks", input),',
    '  importTranscript: (input) => callTimeline("importTranscript", input),',
    '};',
    '',
    'const detectedPlatform = inferPlatform();',
    'const bridge = installMeetingPlatformIntegrationContentScriptBridge(client, {',
    '  platform: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  platforms: Object.keys(PLATFORM_HOSTS),',
    '  runtimePreset: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  browser_runtime_preset: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  source: "meeting_app_extension",',
    '  extensionMessaging: true,',
    '  windowMessaging: true,',
    '  startRuntime: true,',
    '  startOptions: { startRuntime: true },',
    '});',
    '',
    'globalThis.__meetingTimelineBridge = bridge;',
    'extensionRuntime()?.sendMessage?.({',
    '  type: ATTACHED_TYPE,',
    '  platform: inferPlatform(),',
    '  captured_at_ms: Date.now(),',
    '  url: globalThis.location?.href,',
    '});',
  ].join('\n');
}

export function buildMeetingAppExtensionLiveCaptureSource(options = {}) {
  const installPlan = buildMeetingAppExtensionInstallPlan(options);
  const platformMap = Object.fromEntries(installPlan.platforms.map((platform) => {
    const profile = MEETING_APP_EXTENSION_PROFILES[platform];
    return [platform, profile.host_permissions.map((pattern) => matchPatternHost(pattern)).filter(Boolean)];
  }));
  const globalName = firstNonEmpty(
    options.liveCaptureGlobal,
    options.live_capture_global,
    '__meetingTimelineLiveCapture',
  );
  return [
    "import { captureMeetingAppDomSnapshot } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture';",
    "import { createMeetingAppSnapshotRecorder } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder';",
    "import { buildMeetingAppDomAdaptationDiagnosis, buildMeetingAppLiveEvidencePackage } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';",
    '',
    `const PLATFORM_HOSTS = ${json(platformMap)};`,
    `const LIVE_CAPTURE_GLOBAL = ${JSON.stringify(globalName)};`,
    '',
    'function inferPlatform(hostname = globalThis.location?.hostname ?? "") {',
    '  const host = String(hostname);',
    '  for (const [platform, hosts] of Object.entries(PLATFORM_HOSTS)) {',
    '    if (hosts.some((item) => host === item || host.endsWith(`.${item}`))) return platform;',
    '  }',
    '  return "unknown";',
    '}',
    '',
    'function nowMs() {',
    '  return Date.now();',
    '}',
    '',
    'const recorder = createMeetingAppSnapshotRecorder({',
    '  source: "meeting_app_extension_live_capture",',
    '});',
    '',
    'function capture(labelOrOptions = {}, options = {}) {',
    '  const objectInput = labelOrOptions && typeof labelOrOptions === "object" && !Array.isArray(labelOrOptions);',
    '  const merged = objectInput ? { ...labelOrOptions, ...options } : { ...options, label: labelOrOptions };',
    '  const platform = merged.platform ?? inferPlatform();',
    '  const capturedAtMs = merged.captured_at_ms ?? merged.capturedAtMs ?? nowMs();',
    '  const snapshot = captureMeetingAppDomSnapshot({',
    '    window: globalThis,',
    '    document: globalThis.document,',
    '    location: globalThis.location,',
    '    platform,',
    '    url: globalThis.location?.href,',
    '    title: globalThis.document?.title,',
    '  }, {',
    '    ...merged,',
    '    platform,',
    '    captureProfile: merged.captureProfile ?? merged.capture_profile ?? platform,',
    '    source: merged.source ?? "meeting_app_extension_live_capture",',
    '    observedAtMs: capturedAtMs,',
    '  });',
    '  return recorder.add({',
    '    snapshot,',
    '    platform,',
    '    phase: merged.phase,',
    '    label: merged.label,',
    '    captured_at_ms: capturedAtMs,',
    '    url: globalThis.location?.href,',
    '    title: globalThis.document?.title,',
    '  });',
    '}',
    '',
    'function captureActive(options = {}) {',
    '  return capture({ label: options.label ?? "active_speaker", phase: "active", ...options });',
    '}',
    '',
    'function captureEnded(options = {}) {',
    '  return capture({ label: options.label ?? "meeting_ended", phase: "ended", ...options });',
    '}',
    '',
    'function captureRequired(options = {}) {',
    '  const records = [captureActive(options.active ?? options)];',
    '  if (options.includeEnded === true || options.include_ended === true) {',
    '    records.push(captureEnded(options.ended ?? options));',
    '  }',
    '  return records;',
    '}',
    '',
    'function exportRecords(options = {}) {',
    '  return recorder.exportRecords({',
    '    ...options,',
    '    label: options.label ?? "meeting_app_live_capture",',
    '  });',
    '}',
    '',
    'function evidencePackage(options = {}) {',
    '  const recordSet = exportRecords(options);',
    '  return buildMeetingAppLiveEvidencePackage({',
    '    ...options,',
    '    platforms: options.platforms ?? [inferPlatform()],',
    '    recordSet,',
    '  });',
    '}',
    '',
    'function diagnose(options = {}) {',
    '  const recordSet = exportRecords(options);',
    '  return buildMeetingAppDomAdaptationDiagnosis({',
    '    ...options,',
    '    platform: options.platform ?? inferPlatform(),',
    '    recordSet,',
    '  });',
    '}',
    '',
    'const api = {',
    '  capture,',
    '  captureActive,',
    '  captureEnded,',
    '  captureRequired,',
    '  exportRecords,',
    '  evidencePackage,',
    '  diagnose,',
    '  reset: recorder.reset,',
    '  getState: recorder.getState,',
    '  records: recorder.records,',
    '};',
    '',
    'globalThis[LIVE_CAPTURE_GLOBAL] = api;',
    'globalThis.dispatchEvent?.(new CustomEvent("meeting_timeline.live_capture_ready", {',
    '  detail: { platform: inferPlatform(), global: LIVE_CAPTURE_GLOBAL, captured_at_ms: nowMs() },',
    '}));',
  ].join('\n');
}

export function buildMeetingAppExtensionBackgroundSource(options = {}) {
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'https://timeline.example.com')).replace(/\/+$/, '');
  const messagePrefix = firstNonEmpty(
    options.messagePrefix,
    options.message_prefix,
    MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  );
  const endpoints = { ...MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS };
  return [
    "import { createMeetingPlatformRuntimeEventClient } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event';",
    '',
    `const BASE_URL = ${JSON.stringify(baseUrl)};`,
    `const CLIENT_CALL_TYPE = ${JSON.stringify(messagePrefix)};`,
    `const ATTACHED_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached)};`,
    `const STATUS_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status)};`,
    `const OBSERVE_CANDIDATES_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates)};`,
    `const PREFLIGHT_CURRENT_WINDOW_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window)};`,
    `const PREFLIGHT_CANDIDATES_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates)};`,
    `const STATUS_STORAGE_KEY = ${JSON.stringify(MEETING_APP_EXTENSION_STATUS_STORAGE_KEY)};`,
    `const ENDPOINTS = ${json(endpoints)};`,
    'const runtimeEvents = createMeetingPlatformRuntimeEventClient({',
    '  baseUrl: BASE_URL,',
    '  source: "meeting_app_extension_background",',
    '});',
    'let lastAttached = null;',
    '',
    'function runtimeApi() {',
    '  return globalThis.chrome?.runtime ?? globalThis.browser?.runtime;',
    '}',
    '',
    'function storageArea() {',
    '  return globalThis.chrome?.storage?.local ?? globalThis.browser?.storage?.local;',
    '}',
    '',
    'function tabsApi() {',
    '  return globalThis.chrome?.tabs ?? globalThis.browser?.tabs;',
    '}',
    '',
    'function normalizeTab(tab = {}) {',
    '  return {',
    '    id: tab.id,',
    '    tab_id: tab.id,',
    '    window_id: tab.windowId,',
    '    active: Boolean(tab.active),',
    '    audible: tab.audible,',
    '    pinned: tab.pinned,',
    '    discarded: tab.discarded,',
    '    muted: tab.mutedInfo?.muted,',
    '    status: tab.status,',
    '    title: tab.title,',
    '    url: tab.url ?? tab.pendingUrl,',
    '    fav_icon_url: tab.favIconUrl,',
    '    observed_at_ms: Date.now(),',
    '  };',
    '}',
    '',
    'function queryTabs(query = {}) {',
    '  const tabs = tabsApi();',
    '  if (!tabs?.query) return Promise.resolve([]);',
    '  return new Promise((resolve) => {',
    '    let settled = false;',
    '    const finish = (value) => {',
    '      if (settled) return;',
    '      settled = true;',
    '      resolve(Array.isArray(value) ? value : []);',
    '    };',
    '    try {',
    '      const maybePromise = tabs.query(query, finish);',
    '      if (maybePromise?.then) maybePromise.then(finish).catch(() => finish([]));',
    '    } catch {',
    '      finish([]);',
    '    }',
    '  });',
    '}',
    '',
    'async function observedTabsFor(message = {}) {',
    '  if (Array.isArray(message.tabs)) return message.tabs;',
    '  if (Array.isArray(message.input?.tabs)) return message.input.tabs;',
    '  const query = message.query ?? message.input?.query ?? {};',
    '  return (await queryTabs(query)).map(normalizeTab);',
    '}',
    '',
    'async function targetTabIdFor(message = {}, sender = {}) {',
    '  const explicit = message.tab_id ?? message.tabId ?? message.input?.tab_id ?? message.input?.tabId ?? message.payload?.tab_id ?? message.payload?.tabId;',
    '  if (explicit != null) return explicit;',
    '  if (sender?.tab?.id != null) return sender.tab.id;',
    '  if (lastAttached?.sender?.tab_id != null) return lastAttached.sender.tab_id;',
    '  const tabs = await observedTabsFor({ ...message, query: message.query ?? message.input?.query ?? { active: true, currentWindow: true } });',
    '  return tabs.find((tab) => tab.active)?.id ?? tabs[0]?.id;',
    '}',
    '',
    'function candidateObservationInput(message = {}, sender = {}, tabs = []) {',
    '  const capturedAtMs = message.input?.captured_at_ms ?? message.captured_at_ms ?? Date.now();',
    '  return {',
    '    ...(message.input ?? {}),',
    '    source: message.input?.source ?? message.source ?? "meeting_app_extension_background",',
    '    captured_at_ms: capturedAtMs,',
    '    tabs,',
    '    windows: message.input?.windows ?? message.windows,',
    '    applications: message.input?.applications ?? message.input?.apps ?? message.applications ?? message.apps,',
    '    candidates: message.input?.candidates ?? message.candidates,',
    '    environment: {',
    '      ...(message.input?.environment ?? message.environment ?? {}),',
    '      source: "meeting_app_extension_background",',
    '      observed_at_ms: capturedAtMs,',
    '      sender_tab_id: sender?.tab?.id,',
    '      sender_url: sender?.url,',
    '    },',
    '    extension_sender: {',
    '      tab_id: sender?.tab?.id,',
    '      frame_id: sender?.frameId,',
    '      url: sender?.url,',
    '    },',
    '  };',
    '}',
    '',
    'function preflightMessageForTab(message = {}, tab = {}) {',
    '  const capturedAtMs = message.captured_at_ms ?? message.capturedAtMs ?? message.input?.captured_at_ms ?? Date.now();',
    '  return {',
    '    ...message,',
    '    type: PREFLIGHT_CURRENT_WINDOW_TYPE,',
    '    source: message.source ?? "meeting_app_extension_background",',
    '    captured_at_ms: capturedAtMs,',
    '    url: message.url ?? message.input?.url ?? tab.url,',
    '    title: message.title ?? message.input?.title ?? tab.title,',
    '    input: {',
    '      ...(message.input ?? {}),',
    '      url: message.input?.url ?? message.url ?? tab.url,',
    '      title: message.input?.title ?? message.title ?? tab.title,',
    '      tab,',
    '    },',
    '    options: message.options ?? message.preflightOptions ?? message.preflight_options,',
    '  };',
    '}',
    '',
    'function forwardToTab(tabId, message = {}, options = {}) {',
    '  const tabs = tabsApi();',
    '  if (!tabs?.sendMessage) return Promise.resolve({ ok: false, reason: "missing_tabs_send_message", tab_id: tabId });',
    '  if (tabId == null) return Promise.resolve({ ok: false, reason: "missing_target_tab", tab_id: tabId });',
    '  return new Promise((resolve) => {',
    '    let settled = false;',
    '    const finish = (value) => {',
    '      if (settled) return;',
    '      settled = true;',
    '      const runtimeError = globalThis.chrome?.runtime?.lastError ?? globalThis.browser?.runtime?.lastError;',
    '      if (runtimeError) {',
    '        resolve({ ok: false, reason: "tabs_send_message_failed", tab_id: tabId, error: runtimeError.message ?? String(runtimeError) });',
    '      } else {',
    '        resolve({ ok: value?.handled !== false, tab_id: tabId, body: value });',
    '      }',
    '    };',
    '    try {',
    '      const maybePromise = tabs.sendMessage(tabId, message, options, finish);',
    '      if (maybePromise?.then) maybePromise.then(finish).catch((error) => finish({ handled: false, reason: "tabs_send_message_failed", error: String(error?.message ?? error) }));',
    '    } catch (error) {',
    '      finish({ handled: false, reason: "tabs_send_message_failed", error: String(error?.message ?? error) });',
    '    }',
    '  });',
    '}',
    '',
    'async function preflightCandidateTabs(message = {}, sender = {}) {',
    '  const tabs = await observedTabsFor(message);',
    '  const rows = [];',
    '  for (const tab of tabs) {',
    '    if (tab?.id == null) {',
    '      rows.push({ tab, ok: false, reason: "missing_tab_id" });',
    '      continue;',
    '    }',
    '    const response = await forwardToTab(tab.id, preflightMessageForTab(message, tab), message.frame_id != null || message.frameId != null ? { frameId: message.frame_id ?? message.frameId } : {});',
    '    rows.push({',
    '      tab,',
    '      ok: response.ok === true,',
    '      tab_id: response.tab_id,',
    '      reason: response.reason,',
    '      response: response.body,',
    '      preflight: response.body?.result,',
    '    });',
    '  }',
    '  return {',
    '    ok: rows.some((row) => row.ok),',
    '    type: PREFLIGHT_CANDIDATES_TYPE,',
    '    candidate_count: tabs.length,',
    '    preflight_count: rows.length,',
    '    accepted_count: rows.filter((row) => row.preflight?.accepted === true).length,',
    '    rows,',
    '  };',
    '}',
    '',
    'function setStorageValue(key, value) {',
    '  const area = storageArea();',
    '  if (!area?.set) return Promise.resolve({ stored: false, reason: "missing_storage_api" });',
    '  return new Promise((resolve) => {',
    '    let settled = false;',
    '    const finish = (payload) => {',
    '      if (settled) return;',
    '      settled = true;',
    '      resolve(payload);',
    '    };',
    '    try {',
    '      const maybePromise = area.set({ [key]: value }, () => finish({ stored: true }));',
    '      if (maybePromise?.then) {',
    '        maybePromise.then(() => finish({ stored: true })).catch((error) => finish({',
    '          stored: false,',
    '          reason: "storage_set_failed",',
    '          error: String(error?.message ?? error),',
    '        }));',
    '      }',
    '    } catch (error) {',
    '      finish({ stored: false, reason: "storage_set_failed", error: String(error?.message ?? error) });',
    '    }',
    '  });',
    '}',
    '',
    'function getStorageValue(key) {',
    '  const area = storageArea();',
    '  if (!area?.get) return Promise.resolve(undefined);',
    '  return new Promise((resolve) => {',
    '    let settled = false;',
    '    const finish = (value) => {',
    '      if (settled) return;',
    '      settled = true;',
    '      resolve(value);',
    '    };',
    '    try {',
    '      const maybePromise = area.get(key, (value) => finish(value?.[key]));',
    '      if (maybePromise?.then) maybePromise.then((value) => finish(value?.[key])).catch(() => finish(undefined));',
    '    } catch {',
    '      finish(undefined);',
    '    }',
    '  });',
    '}',
    '',
    'async function postJson(path, payload) {',
    '  const response = await fetch(`${BASE_URL}${path}`, {',
    '    method: "POST",',
    '    headers: { "content-type": "application/json" },',
    '    body: JSON.stringify(payload),',
    '  });',
    '  const text = await response.text();',
    '  let body;',
    '  try { body = text ? JSON.parse(text) : undefined; } catch { body = text; }',
    '  return { ok: response.ok, status: response.status, body };',
    '}',
    '',
    'runtimeApi()?.onMessage?.addListener?.((message, sender, sendResponse) => {',
    '  if (message?.type === ATTACHED_TYPE) {',
    '    lastAttached = {',
    '      platform: message.platform,',
    '      url: message.url,',
    '      captured_at_ms: message.captured_at_ms ?? Date.now(),',
    '      sender: {',
    '        tab_id: sender?.tab?.id,',
    '        frame_id: sender?.frameId,',
    '        url: sender?.url,',
    '      },',
    '    };',
    '    setStorageValue(STATUS_STORAGE_KEY, lastAttached).then((storage) => {',
    '      sendResponse({ ok: true, type: ATTACHED_TYPE, attached: lastAttached, storage });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === STATUS_TYPE) {',
    '    getStorageValue(STATUS_STORAGE_KEY).then((stored) => {',
    '      sendResponse({ ok: true, type: STATUS_TYPE, attached: stored ?? lastAttached });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === OBSERVE_CANDIDATES_TYPE) {',
    '    observedTabsFor(message).then((tabs) => {',
    '      const input = candidateObservationInput(message, sender, tabs);',
    '      return runtimeEvents.observePlatformCandidates(input, message.send_options ?? message.sendOptions ?? {});',
    '    }).then((body) => {',
    '      sendResponse({ ok: true, type: OBSERVE_CANDIDATES_TYPE, body });',
    '    }).catch((error) => {',
    '      sendResponse({',
    '        ok: false,',
    '        type: OBSERVE_CANDIDATES_TYPE,',
    '        reason: "observe_candidates_failed",',
    '        error: String(error?.message ?? error),',
    '      });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === PREFLIGHT_CANDIDATES_TYPE) {',
    '    preflightCandidateTabs(message, sender).then((result) => {',
    '      sendResponse(result);',
    '    }).catch((error) => {',
    '      sendResponse({',
    '        ok: false,',
    '        type: PREFLIGHT_CANDIDATES_TYPE,',
    '        reason: "preflight_candidates_failed",',
    '        error: String(error?.message ?? error),',
    '      });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === PREFLIGHT_CURRENT_WINDOW_TYPE) {',
    '    targetTabIdFor(message, sender).then((tabId) => forwardToTab(tabId, {',
    '      ...message,',
    '      type: PREFLIGHT_CURRENT_WINDOW_TYPE,',
    '      source: message.source ?? "meeting_app_extension_background",',
    '      captured_at_ms: message.captured_at_ms ?? message.capturedAtMs ?? Date.now(),',
    '    }, message.frame_id != null || message.frameId != null ? { frameId: message.frame_id ?? message.frameId } : {})).then((result) => {',
    '      sendResponse({ ok: result.ok === true, type: PREFLIGHT_CURRENT_WINDOW_TYPE, ...result });',
    '    }).catch((error) => {',
    '      sendResponse({',
    '        ok: false,',
    '        type: PREFLIGHT_CURRENT_WINDOW_TYPE,',
    '        reason: "preflight_current_window_failed",',
    '        error: String(error?.message ?? error),',
    '      });',
    '    });',
    '    return true;',
    '  }',
    '  if (!message || message.type !== CLIENT_CALL_TYPE) return false;',
    '  const path = ENDPOINTS[message.method];',
    '  if (!path) {',
    '    sendResponse({ ok: false, reason: "unsupported_method", method: message.method });',
    '    return false;',
    '  }',
    '  const payload = {',
    '    ...(message.input ?? {}),',
    '    platform: message.input?.platform ?? message.platform,',
    '    captured_at_ms: message.input?.captured_at_ms ?? message.captured_at_ms,',
    '    meeting_url: message.input?.meeting_url ?? message.url,',
    '    detector_source: message.input?.detector_source ?? "meeting_app_extension",',
    '    extension_sender: {',
    '      tab_id: sender?.tab?.id,',
    '      frame_id: sender?.frameId,',
    '      url: sender?.url,',
    '    },',
    '  };',
    '  postJson(path, payload).then(sendResponse).catch((error) => {',
    '    sendResponse({ ok: false, reason: "timeline_request_failed", error: String(error?.message ?? error) });',
    '  });',
    '  return true;',
    '});',
  ].join('\n');
}

export function buildMeetingAppExtensionPackageJson(options = {}) {
  const packageName = firstNonEmpty(
    options.packageName,
    options.package_name,
    'meeting-timeline-meeting-app-extension',
  );
  const sdkVersion = firstNonEmpty(
    options.sdkDependencyVersion,
    options.sdk_dependency_version,
    '^0.1.0',
  );
  const esbuildVersion = firstNonEmpty(
    options.esbuildVersion,
    options.esbuild_version,
    '^0.25.0',
  );
  return compactObject({
    name: packageName,
    version: firstNonEmpty(options.packageVersion, options.package_version, '0.1.0'),
    private: options.privatePackage ?? options.private_package ?? true,
    type: 'module',
    scripts: {
      build: 'node build.mjs',
      watch: 'node build.mjs --watch',
      ...(options.scripts ?? {}),
    },
    dependencies: {
      '@ai-annotation/meeting-timeline-sdk': sdkVersion,
      ...(options.dependencies ?? {}),
    },
    devDependencies: {
      esbuild: esbuildVersion,
      ...(options.devDependencies ?? options.dev_dependencies ?? {}),
    },
  });
}

export function buildMeetingAppExtensionBuildSource(options = {}) {
  const contentScriptEntry = firstNonEmpty(
    options.contentScriptEntry,
    options.content_script_entry,
    'src/content-script.entry.mjs',
  );
  const contentScriptOutput = firstNonEmpty(
    options.outputScript,
    options.output_script,
    'content-script.js',
  );
  const backgroundEntry = firstNonEmpty(
    options.backgroundEntry,
    options.background_entry,
    'src/background.entry.mjs',
  );
  const backgroundOutput = firstNonEmpty(
    options.backgroundScript,
    options.background_script,
    'background.js',
  );
  const target = asArray(firstNonEmpty(
    options.buildTarget,
    options.build_target,
    ['chrome114', 'edge114'],
  ));
  const includeLiveCapture = options.includeLiveCapture !== false && options.include_live_capture !== false;
  const liveCaptureEntry = firstNonEmpty(
    options.liveCaptureEntry,
    options.live_capture_entry,
    'src/live-capture.entry.mjs',
  );
  const liveCaptureOutput = firstNonEmpty(
    options.liveCaptureScript,
    options.live_capture_script,
    'live-capture.js',
  );
  const extraConfig = includeLiveCapture ? [
    '  {',
    `    entryPoints: [${JSON.stringify(liveCaptureEntry)}],`,
    `    outfile: ${JSON.stringify(liveCaptureOutput)},`,
    '    bundle: true,',
    '    platform: "browser",',
    '    format: "iife",',
    '    target,',
    '    sourcemap: !production,',
    '    minify: production,',
    '  },',
  ] : [];
  return [
    "import { build, context } from 'esbuild';",
    '',
    'const watch = process.argv.includes("--watch");',
    'const production = process.env.NODE_ENV === "production";',
    `const target = ${json(target)};`,
    '',
    'const configs = [',
    '  {',
    `    entryPoints: [${JSON.stringify(contentScriptEntry)}],`,
    `    outfile: ${JSON.stringify(contentScriptOutput)},`,
    '    bundle: true,',
    '    platform: "browser",',
    '    format: "iife",',
    '    target,',
    '    sourcemap: !production,',
    '    minify: production,',
    '  },',
    '  {',
    `    entryPoints: [${JSON.stringify(backgroundEntry)}],`,
    `    outfile: ${JSON.stringify(backgroundOutput)},`,
    '    bundle: true,',
    '    platform: "browser",',
    '    format: "esm",',
    '    target,',
    '    sourcemap: !production,',
    '    minify: production,',
    '  },',
    ...extraConfig,
    '];',
    '',
    'if (watch) {',
    '  const contexts = await Promise.all(configs.map((config) => context(config)));',
    '  await Promise.all(contexts.map((item) => item.watch()));',
    '  console.log("watching meeting timeline extension sources");',
    '} else {',
    '  await Promise.all(configs.map((config) => build(config)));',
    '}',
  ].join('\n');
}

export function buildMeetingAppExtensionReadme(options = {}) {
  const installPlan = buildMeetingAppExtensionInstallPlan(options);
  const scriptFile = installPlan.content_scripts[0]?.js?.[0] ?? 'content-script.js';
  const includeLiveCapture = options.includeLiveCapture !== false && options.include_live_capture !== false;
  const liveCaptureFile = firstNonEmpty(options.liveCaptureScript, options.live_capture_script, 'live-capture.js');
  const backgroundFile = firstNonEmpty(options.backgroundScript, options.background_script, 'background.js');
  const emittedFiles = includeLiveCapture
    ? `\`${scriptFile}\`, \`${liveCaptureFile}\`, and \`${backgroundFile}\``
    : `\`${scriptFile}\` and \`${backgroundFile}\``;
  const liveCaptureDocs = includeLiveCapture ? [
    '',
    'When the extension is loaded, `window.__meetingTimelineLiveCapture` exposes `captureActive()`, `captureEnded()`, `exportRecords()`, `evidencePackage()`, and `diagnose()` for live DOM validation.',
  ] : [];
  return [
    '# Meeting Timeline Browser Extension',
    '',
    'This scaffold installs the Meeting Timeline content bridge into supported meeting web apps.',
    '',
    '## Supported Pages',
    '',
    ...installPlan.matches.map((pattern) => `- \`${pattern}\``),
    '',
    '## Build',
    '',
    '```sh',
    'npm install',
    'npm run build',
    '```',
    '',
    `The build emits ${emittedFiles}, then \`manifest.json\` can be loaded as an unpacked Chrome/Edge extension.`,
    '',
    'The generated background worker forwards timeline client calls to the configured timeline service base URL.',
    '',
    'Protocol messages use `meeting_timeline.extension_attached`, `meeting_timeline.extension_status`, `meeting_timeline.observe_candidates`, `meeting_timeline.preflight_current_window`, `meeting_timeline.preflight_candidates`, and `meeting_timeline.client_call`; build custom callers with the SDK message helpers instead of hardcoded strings.',
    '',
    'Send `meeting_timeline.observe_candidates` from a popup, native host, or diagnostic page to let the background worker query browser tabs and forward an `observe_platform_candidates` runtime event to `/api/meeting-platform/runtime-events`.',
    '',
    'Send `meeting_timeline.preflight_current_window` to let the background worker forward the request into the current content script and return the local DOM preflight result before enabling realtime annotations.',
    '',
    'Send `meeting_timeline.preflight_candidates` to query candidate meeting tabs and run `preflight_current_window` inside each content script, returning a per-tab readiness summary for Google Meet, Teams, Zoom, Webex, and Lark.',
    '',
    'Before production rollout, capture real meeting app snapshots with `meeting-app-snapshot-recorder` and validate them with `meeting-app-gate`.',
    ...liveCaptureDocs,
  ].join('\n');
}

export function buildMeetingAppExtensionScaffold(options = {}) {
  const scriptFile = firstNonEmpty(options.outputScript, options.output_script, 'content-script.js');
  const includeLiveCapture = options.includeLiveCapture !== false && options.include_live_capture !== false;
  const liveCaptureFile = firstNonEmpty(options.liveCaptureScript, options.live_capture_script, 'live-capture.js');
  const backgroundFile = firstNonEmpty(options.backgroundScript, options.background_script, 'background.js');
  const contentScriptEntry = firstNonEmpty(options.contentScriptEntry, options.content_script_entry, 'src/content-script.entry.mjs');
  const liveCaptureEntry = firstNonEmpty(options.liveCaptureEntry, options.live_capture_entry, 'src/live-capture.entry.mjs');
  const backgroundEntry = firstNonEmpty(options.backgroundEntry, options.background_entry, 'src/background.entry.mjs');
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const endpointPermission = endpointOriginPattern(baseUrl);
  const packageJson = buildMeetingAppExtensionPackageJson(options);
  const buildSource = buildMeetingAppExtensionBuildSource({
    ...options,
    outputScript: scriptFile,
    backgroundScript: backgroundFile,
    contentScriptEntry,
    backgroundEntry,
  });
  const manifest = buildMeetingAppContentScriptManifest({
    ...options,
    js: includeLiveCapture ? [scriptFile, liveCaptureFile] : scriptFile,
    permissions: unique([
      ...asArray(firstNonEmpty(options.permissions, options.extension_permissions, [])),
      'storage',
      'tabs',
    ]),
    extraHostPermissions: unique([
      ...asArray(firstNonEmpty(options.extraHostPermissions, options.extra_host_permissions)),
      endpointPermission,
    ]),
    manifestOverrides: {
      ...(options.manifestOverrides ?? options.manifest_overrides ?? {}),
      background: {
        service_worker: backgroundFile,
        type: 'module',
        ...((options.manifestOverrides ?? options.manifest_overrides ?? {}).background ?? {}),
      },
    },
  });
  const installPlan = buildMeetingAppExtensionInstallPlan({
    ...options,
    js: includeLiveCapture ? [scriptFile, liveCaptureFile] : scriptFile,
    extraHostPermissions: unique([
      ...asArray(firstNonEmpty(options.extraHostPermissions, options.extra_host_permissions)),
      endpointPermission,
    ]),
  });
  const contentScriptSource = buildMeetingAppExtensionContentScriptSource(options);
  const liveCaptureSource = includeLiveCapture ? buildMeetingAppExtensionLiveCaptureSource(options) : undefined;
  const backgroundSource = buildMeetingAppExtensionBackgroundSource(options);
  const readme = buildMeetingAppExtensionReadme({
    ...options,
    js: scriptFile,
    backgroundScript: backgroundFile,
  });
  return compactObject({
    type: 'meeting_app_extension_scaffold',
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    install_plan: installPlan,
    manifest,
    files: [
      sourceFile('package.json', `${json(packageJson)}\n`, 'package_manifest', 'application/json'),
      sourceFile('build.mjs', buildSource, 'build_script', 'text/javascript'),
      sourceFile('manifest.json', `${json(manifest)}\n`, 'manifest', 'application/json'),
      sourceFile(contentScriptEntry, contentScriptSource, 'content_script_entry', 'text/javascript'),
      ...(includeLiveCapture ? [
        sourceFile(liveCaptureEntry, liveCaptureSource, 'live_capture_entry', 'text/javascript'),
      ] : []),
      sourceFile(backgroundEntry, backgroundSource, 'background_service_worker_entry', 'text/javascript'),
      sourceFile('README.md', readme, 'readme', 'text/markdown'),
    ],
    bundle: {
      content_script_input: contentScriptEntry,
      content_script_output: scriptFile,
      live_capture_input: includeLiveCapture ? liveCaptureEntry : undefined,
      live_capture_output: includeLiveCapture ? liveCaptureFile : undefined,
      background_input: backgroundEntry,
      background_output: backgroundFile,
    },
    validation: {
      uses_all_urls: manifest.host_permissions?.includes('<all_urls>') ?? false,
      platform_count: installPlan.platforms.length,
    },
  });
}

export function buildMeetingAppExtensionScaffoldAcceptanceReport(scaffoldOrOptions = {}, options = {}) {
  const scaffold = scaffoldOrOptions?.type === 'meeting_app_extension_scaffold'
    ? scaffoldOrOptions
    : buildMeetingAppExtensionScaffold(scaffoldOrOptions);
  const reportOptions = scaffoldOrOptions?.type === 'meeting_app_extension_scaffold' ? options : scaffoldOrOptions;
  const issues = [];
  const requiredPlatforms = platformList({
    ...reportOptions,
    platforms: firstNonEmpty(
      reportOptions.platforms,
      reportOptions.platform_keys,
      reportOptions.platformKeys,
      reportOptions.platform,
      scaffold.install_plan?.platforms,
    ),
  });
  const filePaths = asArray(scaffold.files).map((file) => file.path);
  const requiredFiles = [
    'package.json',
    'build.mjs',
    'manifest.json',
    scaffold.bundle?.content_script_input ?? 'src/content-script.entry.mjs',
    ...(scaffold.bundle?.live_capture_input ? [scaffold.bundle.live_capture_input] : []),
    scaffold.bundle?.background_input ?? 'src/background.entry.mjs',
    'README.md',
  ];
  for (const path of requiredFiles) {
    if (!filePaths.includes(path)) {
      issues.push(issue('error', 'missing_scaffold_file', `Missing generated scaffold file: ${path}.`, { path }));
    }
  }

  const manifest = jsonFromFile(scaffold, 'manifest.json', issues) ?? scaffold.manifest ?? {};
  const packageJson = jsonFromFile(scaffold, 'package.json', issues) ?? {};
  const buildFile = fileByPath(scaffold, 'build.mjs');
  const contentEntry = fileByPath(scaffold, scaffold.bundle?.content_script_input ?? 'src/content-script.entry.mjs');
  const liveCaptureEntry = scaffold.bundle?.live_capture_input ? fileByPath(scaffold, scaffold.bundle.live_capture_input) : null;
  const backgroundEntry = fileByPath(scaffold, scaffold.bundle?.background_input ?? 'src/background.entry.mjs');
  const readme = fileByPath(scaffold, 'README.md');

  if (manifest.manifest_version !== 3) {
    issues.push(issue('error', 'manifest_not_mv3', 'Extension manifest must use manifest_version 3.', {
      manifest_version: manifest.manifest_version,
    }));
  }
  if (asArray(manifest.host_permissions).includes('<all_urls>')) {
    issues.push(issue('error', 'overbroad_host_permission', 'Extension scaffold must not request <all_urls>.'));
  }
  if (!asArray(manifest.permissions).includes('storage')) {
    issues.push(issue('error', 'missing_storage_permission', 'Manifest must request storage permission for extension diagnostics.'));
  }
  if (!asArray(manifest.permissions).includes('tabs')) {
    issues.push(issue('error', 'missing_tabs_permission', 'Manifest must request tabs permission for background platform candidate observation.'));
  }
  if (!manifest.background?.service_worker) {
    issues.push(issue('error', 'missing_background_worker', 'Manifest is missing a background service worker.'));
  }
  if (manifest.background?.type !== 'module') {
    issues.push(issue('warning', 'background_worker_not_module', 'Background service worker should be module typed.'));
  }
  const contentScript = asArray(manifest.content_scripts)[0];
  if (!contentScript) {
    issues.push(issue('error', 'missing_content_script', 'Manifest is missing content_scripts[0].'));
  } else {
    if (asArray(contentScript.matches).length === 0) {
      issues.push(issue('error', 'missing_content_script_matches', 'Content script has no match patterns.'));
    }
    if (!asArray(contentScript.js).includes(scaffold.bundle?.content_script_output ?? 'content-script.js')) {
      issues.push(issue('error', 'content_script_output_mismatch', 'Manifest content script output does not match scaffold bundle output.', {
        expected: scaffold.bundle?.content_script_output,
        actual: contentScript.js,
      }));
    }
    if (scaffold.bundle?.live_capture_output && !asArray(contentScript.js).includes(scaffold.bundle.live_capture_output)) {
      issues.push(issue('error', 'live_capture_output_mismatch', 'Manifest content script output does not include the live capture bundle.', {
        expected: scaffold.bundle.live_capture_output,
        actual: contentScript.js,
      }));
    }
  }
  if (manifest.background?.service_worker !== scaffold.bundle?.background_output) {
    issues.push(issue('error', 'background_output_mismatch', 'Manifest background worker does not match scaffold bundle output.', {
      expected: scaffold.bundle?.background_output,
      actual: manifest.background?.service_worker,
    }));
  }

  if (!packageJson.dependencies?.['@ai-annotation/meeting-timeline-sdk']) {
    issues.push(issue('error', 'missing_sdk_dependency', 'package.json must depend on @ai-annotation/meeting-timeline-sdk.'));
  }
  if (!packageJson.devDependencies?.esbuild) {
    issues.push(issue('error', 'missing_esbuild_dependency', 'package.json must include esbuild as a devDependency.'));
  }
  if (packageJson.scripts?.build !== 'node build.mjs') {
    issues.push(issue('warning', 'unexpected_build_script', 'package.json scripts.build should run node build.mjs.', {
      actual: packageJson.scripts?.build,
    }));
  }

  const buildContent = buildFile?.content ?? '';
  if (!buildContent.includes(scaffold.bundle?.content_script_input ?? 'src/content-script.entry.mjs')) {
    issues.push(issue('error', 'build_missing_content_entry', 'build.mjs does not include the content script entry.'));
  }
  if (!buildContent.includes(scaffold.bundle?.background_input ?? 'src/background.entry.mjs')) {
    issues.push(issue('error', 'build_missing_background_entry', 'build.mjs does not include the background entry.'));
  }
  if (scaffold.bundle?.live_capture_input && !buildContent.includes(scaffold.bundle.live_capture_input)) {
    issues.push(issue('error', 'build_missing_live_capture_entry', 'build.mjs does not include the live capture entry.'));
  }
  if (!buildContent.includes('format: "iife"')) {
    issues.push(issue('error', 'build_missing_content_iife', 'build.mjs should bundle the content script as an IIFE.'));
  }
  if (!buildContent.includes('format: "esm"')) {
    issues.push(issue('error', 'build_missing_background_esm', 'build.mjs should bundle the background worker as ESM.'));
  }

  const contentEntryContent = contentEntry?.content ?? '';
  if (!contentEntryContent.includes('installMeetingPlatformIntegrationContentScriptBridge')) {
    issues.push(issue('error', 'content_entry_missing_bridge', 'Content script entry does not install the SDK bridge.'));
  }
  if (!contentEntryContent.includes('startRuntime: true')) {
    issues.push(issue('error', 'content_entry_not_starting_runtime', 'Content script entry does not start the browser runtime.'));
  }
  const liveCaptureContent = liveCaptureEntry?.content ?? '';
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('__meetingTimelineLiveCapture')) {
    issues.push(issue('error', 'live_capture_missing_global', 'Live capture entry does not expose __meetingTimelineLiveCapture.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('captureMeetingAppDomSnapshot')) {
    issues.push(issue('error', 'live_capture_missing_dom_capture', 'Live capture entry does not capture meeting app DOM snapshots.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('buildMeetingAppLiveEvidencePackage')) {
    issues.push(issue('error', 'live_capture_missing_evidence_package', 'Live capture entry does not build live evidence packages.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('buildMeetingAppDomAdaptationDiagnosis')) {
    issues.push(issue('error', 'live_capture_missing_dom_diagnosis', 'Live capture entry does not build DOM adaptation diagnostics.'));
  }
  const backgroundContent = backgroundEntry?.content ?? '';
  if (!backgroundContent.includes('meeting_timeline.extension_attached')) {
    issues.push(issue('error', 'background_missing_attached_diagnostic', 'Background entry does not handle extension_attached diagnostics.'));
  }
  if (!backgroundContent.includes('meeting_timeline.extension_status')) {
    issues.push(issue('error', 'background_missing_status_diagnostic', 'Background entry does not handle extension_status diagnostics.'));
  }
  if (!backgroundContent.includes('meeting_timeline.observe_candidates')) {
    issues.push(issue('error', 'background_missing_candidate_observer', 'Background entry does not handle observe_candidates runtime events.'));
  }
  if (!backgroundContent.includes('createMeetingPlatformRuntimeEventClient')) {
    issues.push(issue('error', 'background_missing_runtime_event_client', 'Background entry does not use the platform runtime event client.'));
  }
  if (!backgroundContent.includes('setStorageValue')) {
    issues.push(issue('error', 'background_missing_diagnostic_storage', 'Background entry does not persist extension diagnostics.'));
  }
  for (const endpoint of ['/api/meeting-session/start', '/api/meeting-session/end', '/api/annotations', '/api/annotations/batch']) {
    if (!backgroundContent.includes(endpoint)) {
      issues.push(issue('error', 'background_missing_timeline_endpoint', `Background entry is missing ${endpoint}.`, { endpoint }));
    }
  }
  if (!readme?.content?.includes('npm run build')) {
    issues.push(issue('warning', 'readme_missing_build_step', 'README should document npm run build.'));
  }

  const matches = asArray(contentScript?.matches);
  const hostPermissions = asArray(manifest.host_permissions);
  const coverageByPlatform = Object.fromEntries(requiredPlatforms.map((platform) => {
    const profile = MEETING_APP_EXTENSION_PROFILES[platform];
    const missingMatches = profile.matches.filter((pattern) => !matches.includes(pattern));
    const missingHostPermissions = profile.host_permissions.filter((pattern) => !hostPermissions.includes(pattern));
    if (missingMatches.length > 0) {
      issues.push(issue('error', 'missing_platform_match_patterns', `Missing content script match patterns for ${platform}.`, {
        platform,
        missing: missingMatches,
      }));
    }
    if (missingHostPermissions.length > 0) {
      issues.push(issue('error', 'missing_platform_host_permissions', `Missing host permissions for ${platform}.`, {
        platform,
        missing: missingHostPermissions,
      }));
    }
    return [platform, {
      matches: missingMatches.length === 0,
      host_permissions: missingHostPermissions.length === 0,
      accepted: missingMatches.length === 0 && missingHostPermissions.length === 0,
    }];
  }));
  const accepted = issues.every((item) => item.severity !== 'error');
  return compactObject({
    type: 'meeting_app_extension_scaffold_acceptance_report',
    schema: MEETING_APP_EXTENSION_SCHEMA,
    version: 1,
    accepted,
    platforms: requiredPlatforms,
    accepted_platform_count: Object.values(coverageByPlatform).filter((item) => item.accepted).length,
    platform_count: requiredPlatforms.length,
    coverage_by_platform: coverageByPlatform,
    files: filePaths,
    manifest: {
      manifest_version: manifest.manifest_version,
      permissions: asArray(manifest.permissions),
      host_permissions: hostPermissions,
      content_script_count: asArray(manifest.content_scripts).length,
      background_service_worker: manifest.background?.service_worker,
      uses_all_urls: hostPermissions.includes('<all_urls>'),
    },
    bundle: scaffold.bundle,
    issues,
  });
}

export function assertMeetingAppExtensionScaffold(scaffoldOrOptions = {}, options = {}) {
  const report = buildMeetingAppExtensionScaffoldAcceptanceReport(scaffoldOrOptions, options);
  if (!report.accepted) {
    throw new MeetingTimelineSdkError('Meeting app extension scaffold acceptance failed', report);
  }
  return report;
}

export default buildMeetingAppExtensionInstallPlan;
