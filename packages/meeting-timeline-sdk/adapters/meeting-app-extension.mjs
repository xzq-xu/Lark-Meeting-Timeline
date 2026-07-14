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
  service_status: 'meeting_timeline.service_status',
  configure: 'meeting_timeline.configure',
  capture_evidence: 'meeting_timeline.capture_evidence',
  observe_candidates: 'meeting_timeline.observe_candidates',
  preflight_current_window: 'meeting_timeline.preflight_current_window',
  preflight_candidates: 'meeting_timeline.preflight_candidates',
  runtime_target: 'meeting_timeline.runtime_target',
  open_session: 'meeting_timeline.open_session',
  candidate_launch_plan: 'meeting_timeline.candidate_launch_plan',
  open_candidate_session: 'meeting_timeline.open_candidate_session',
});

export const MEETING_APP_EXTENSION_STATUS_STORAGE_KEY = 'meeting_timeline_extension_status';
export const MEETING_APP_EXTENSION_BASE_URL_STORAGE_KEY = 'meeting_timeline_base_url';

export const MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS = Object.freeze({
  startMeeting: '/api/meeting-session/start',
  endMeeting: '/api/meeting-session/end',
  recordP0Reference: '/api/meeting-platform/p0-reference',
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
  service_status: MEETING_APP_EXTENSION_MESSAGE_TYPES.service_status,
  servicestatus: MEETING_APP_EXTENSION_MESSAGE_TYPES.service_status,
  'service-status': MEETING_APP_EXTENSION_MESSAGE_TYPES.service_status,
  configure: MEETING_APP_EXTENSION_MESSAGE_TYPES.configure,
  configuration: MEETING_APP_EXTENSION_MESSAGE_TYPES.configure,
  capture_evidence: MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence,
  captureevidence: MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence,
  'capture-evidence': MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence,
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
  runtime_target: MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  runtimetarget: MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  'runtime-target': MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  adapter_runtime_target: MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  adapterruntimetarget: MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  'adapter-runtime-target': MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
  open_session: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
  opensession: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
  'open-session': MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
  open_adapter_session: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
  openadaptersession: MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
  'open-adapter-session': MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
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
      'https://teams.live.com/*',
      'https://*.teams.live.com/*',
      'https://teams.cloud.microsoft/*',
      'https://*.teams.cloud.microsoft/*',
    ]),
    host_permissions: Object.freeze([
      'https://teams.microsoft.com/*',
      'https://*.teams.microsoft.com/*',
      'https://teams.live.com/*',
      'https://*.teams.live.com/*',
      'https://teams.cloud.microsoft/*',
      'https://*.teams.cloud.microsoft/*',
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

function runtimeSessionOptions(merged = {}) {
  const runtimeOptions = {
    ...(plainObject(merged.options) ? merged.options : {}),
    ...(plainObject(merged.runtimeOptions) ? merged.runtimeOptions : {}),
    ...(plainObject(merged.runtime_options) ? merged.runtime_options : {}),
  };
  return Object.keys(runtimeOptions).length > 0 ? runtimeOptions : undefined;
}

function buildRuntimeSessionMessage(input = {}, options = {}, defaultType) {
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
    active: merged.active,
    in_meeting: firstNonEmpty(merged.in_meeting, merged.inMeeting),
    query: firstNonEmpty(merged.query, merged.tabs_query, merged.tabsQuery),
    tabs: firstNonEmpty(merged.tabs, merged.browser_tabs, merged.browserTabs),
    windows: firstNonEmpty(merged.windows, merged.browser_windows, merged.browserWindows),
    candidates: merged.candidates,
    input: plainObject(merged.input) ? merged.input : undefined,
    options: runtimeSessionOptions(merged),
  });
}

export function buildMeetingAppExtensionRuntimeTargetMessage(input = {}, options = {}) {
  return buildRuntimeSessionMessage(input, options, 'runtime_target');
}

export function buildMeetingAppExtensionOpenSessionMessage(input = {}, options = {}) {
  return buildRuntimeSessionMessage(input, options, 'open_session');
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
    content_script_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-local-content-runtime',
    local_content_runtime_adapter: '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-local-content-runtime',
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
        MEETING_APP_EXTENSION_MESSAGE_TYPES.runtime_target,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.open_session,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
        MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence,
        'meeting_timeline.insert_mark',
        'meeting_timeline.sample_tracks',
      ],
      status_storage_key: MEETING_APP_EXTENSION_STATUS_STORAGE_KEY,
      base_url_storage_key: MEETING_APP_EXTENSION_BASE_URL_STORAGE_KEY,
      timeline_endpoints: { ...MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS },
      live_capture_global: '__meetingTimelineLiveCapture',
      live_capture_methods: ['captureActive', 'captureEnded', 'exportRecords', 'evidencePackage', 'diagnose'],
      required_signals: ['meeting_started', 'speaker_started', 'meeting_ended'],
      timestamp_field: 'captured_at_ms',
    },
    next_steps: [
      'Bundle the browser-safe meeting-app-local-content-runtime into the js file declared in content_scripts.',
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
    "import { installMeetingAppLocalContentRuntime } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-local-content-runtime';",
    "import { captureMeetingAppDomSnapshot } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture';",
    '',
    `const PLATFORM_HOSTS = ${json(platformMap)};`,
    `const CLIENT_CALL_TYPE = ${JSON.stringify(messagePrefix)};`,
    `const ATTACHED_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached)};`,
    `const CAPTURE_EVIDENCE_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence)};`,
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
    'function meetingAppEvidenceRecord(phase, input = {}) {',
    '  const platform = inferPlatform();',
    '  const capturedAtMs = input.start_time_ms ?? input.end_time_ms ?? input.captured_at_ms ?? Date.now();',
    '  const snapshot = captureMeetingAppDomSnapshot({',
    '    window: globalThis,',
    '    document: globalThis.document,',
    '    location: globalThis.location,',
    '    platform,',
    '    url: globalThis.location?.href,',
    '    title: globalThis.document?.title,',
    '  }, {',
    '    platform,',
    '    captureProfile: platform,',
    '    source: "meeting_app_extension_auto_evidence",',
    '    observedAtMs: capturedAtMs,',
    '    phase,',
    '  });',
    '  return {',
    '    schema: "meeting_app_snapshot_record",',
    '    schema_version: 1,',
    '    id: `${platform}:${phase}:${capturedAtMs}`,',
    '    platform,',
    '    phase,',
    '    label: phase === "ended" ? "meeting_ended" : "active_speaker",',
    '    captured_at_ms: capturedAtMs,',
    '    source: "meeting_app_extension_auto_evidence",',
    '    observer_surface: "browser_extension",',
    '    snapshot,',
    '  };',
    '}',
    '',
    'function withMeetingAppEvidence(input = {}, phase) {',
    '  return {',
    '    ...input,',
    '    observer_surface: "browser_extension",',
    '    suppress_auto_annotations: true,',
    '    meeting_app_record: meetingAppEvidenceRecord(phase, input),',
    '  };',
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
    'function operatorReferenceAction(event = {}) {',
    '  const target = event.target?.closest?.("button,[role=\\"button\\"]") ?? event.target;',
    '  if (!target) return null;',
    '  const label = [',
    '    target.getAttribute?.("aria-label"),',
    '    target.getAttribute?.("data-tooltip"),',
    '    target.getAttribute?.("title"),',
    '    target.textContent,',
    '  ].filter(Boolean).join(" ").replace(/\\s+/g, " ").trim();',
    '  if (/\\bjoin now\\b|\\bask to join\\b|\\bjoin meeting\\b|立即加入|加入会议|进入会议|开始会议/i.test(label)) return "join";',
    '  if (/\\bleave call\\b|\\bleave meeting\\b|\\bend call\\b|\\bend meeting\\b|\\bhang up\\b|离开通话|离开会议|结束通话|结束会议|挂断/i.test(label)) return "leave";',
    '  return null;',
    '}',
    '',
    'function bindOperatorReferenceCapture(bridge) {',
    '  const document = globalThis.document;',
    '  if (!document?.addEventListener) return { installed: false, reason: "missing_document_event_target" };',
    '  const lastCapturedAt = new Map();',
    '  const listener = (event) => {',
    '    const action = operatorReferenceAction(event);',
    '    if (!action) return;',
    '    const atMs = Date.now();',
    '    if (atMs - (lastCapturedAt.get(action) ?? 0) < 1_200) return;',
    '    const meetingId = bridge.preflight?.().meeting_id;',
    '    if (!meetingId) return;',
    '    lastCapturedAt.set(action, atMs);',
    '    callTimeline("recordP0Reference", {',
    '      action,',
    '      at_ms: atMs,',
    '      meeting_id: meetingId,',
    '      observer_surface: "browser_extension",',
    '    }).catch(() => {});',
    '  };',
    '  document.addEventListener("click", listener, true);',
    '  return { installed: true, event: "click", listener };',
    '}',
    '',
    'function handleEvidenceRequest(message = {}) {',
    '  const api = globalThis.__meetingTimelineLiveCapture;',
    '  if (!api) return { handled: false, reason: "live_capture_not_ready" };',
    '  const action = message.action ?? "diagnose";',
    '  const options = {',
    '    ...(message.options ?? {}),',
    '    captured_at_ms: message.captured_at_ms ?? message.capturedAtMs ?? Date.now(),',
    '  };',
    '  if (action === "capture_active") return { handled: true, action, result: api.captureActive(options) };',
    '  if (action === "capture_ended") return { handled: true, action, result: api.captureEnded(options) };',
    '  if (action === "export_records") return { handled: true, action, result: api.exportRecords(options) };',
    '  if (action === "evidence_package") return { handled: true, action, result: api.evidencePackage(options) };',
    '  if (action === "diagnose") return { handled: true, action, result: api.diagnose(options) };',
    '  if (action === "reset") return { handled: true, action, result: api.reset() };',
    '  return { handled: false, action, reason: "unsupported_evidence_action" };',
    '}',
    '',
    'const client = {',
    '  startMeeting: (input) => callTimeline("startMeeting", withMeetingAppEvidence(input, "active")),',
    '  endMeeting: (input) => callTimeline("endMeeting", withMeetingAppEvidence(input, "ended")),',
    '  insertMark: (input) => callTimeline("insertMark", input),',
    '  insertMarks: (input) => callTimeline("insertMarks", input),',
    '  importTranscript: (input) => callTimeline("importTranscript", input),',
    '};',
    '',
    'const detectedPlatform = inferPlatform();',
    'const bridge = installMeetingAppLocalContentRuntime(client, {',
    '  platform: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  platforms: Object.keys(PLATFORM_HOSTS),',
    '  runtimePreset: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  browser_runtime_preset: detectedPlatform === "unknown" ? undefined : detectedPlatform,',
    '  source: "meeting_app_extension",',
    '  extensionMessaging: true,',
    '  windowMessaging: true,',
    '  startRuntime: true,',
    '  applyOptions: { speakerAsAnnotation: true, participantAsAnnotation: false },',
    '  speakerOptions: { minStableMs: 650, switchStableMs: 700, endIdleMs: 1500 },',
    '  startOptions: { startRuntime: true },',
    '});',
    '',
    'globalThis.__meetingTimelineBridge = bridge;',
    'globalThis.__meetingTimelineOperatorReferenceCapture = bindOperatorReferenceCapture(bridge);',
    'extensionRuntime()?.onMessage?.addListener?.((message, _sender, sendResponse) => {',
    '  if (message?.type !== CAPTURE_EVIDENCE_TYPE) return false;',
    '  try {',
    '    sendResponse(handleEvidenceRequest(message));',
    '  } catch (error) {',
    '    sendResponse({ handled: false, action: message.action, reason: "evidence_capture_failed", error: String(error?.message ?? error) });',
    '  }',
    '  return false;',
    '});',
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
    '  const platforms = options.platforms ?? [inferPlatform()];',
    '  const phases = new Set(recordSet.records.map((record) => record.phase));',
    '  const accepted = phases.has("active") && phases.has("ended");',
    '  return {',
    '    type: "meeting_app_live_evidence_package",',
    '    schema: "meeting_app_live_evidence_package",',
    '    schema_version: 1,',
    '    id: options.packageId ?? options.package_id ?? options.id ?? `live-${nowMs()}`,',
    '    platforms,',
    '    record_set: recordSet,',
    '    record_count: recordSet.record_count,',
    '    accepted,',
    '    production_ready: accepted,',
    '  };',
    '}',
    '',
    'function diagnose(options = {}) {',
    '  const recordSet = exportRecords(options);',
    '  const platform = options.platform ?? inferPlatform();',
    '  const phases = new Set(recordSet.records.map((record) => record.phase));',
    '  const active = recordSet.records.find((record) => record.phase === "active")?.snapshot;',
    '  const controls = active?.page?.controls ?? active?.page?.buttons ?? [];',
    '  const participants = active?.page?.participants ?? active?.page?.tiles ?? [];',
    '  const accepted = phases.has("active") && phases.has("ended") && controls.length > 0;',
    '  return {',
    '    type: "meeting_app_dom_adaptation_diagnosis",',
    '    schema: "meeting_app_dom_adaptation_diagnosis",',
    '    schema_version: 1,',
    '    platform,',
    '    accepted,',
    '    production_ready: accepted,',
    '    record_count: recordSet.record_count,',
    '    selector_probe: { matched: { controls: controls.length > 0, participants: participants.length > 0, active_speaker: participants.some((row) => row.speaking || row.active_speaker) } },',
    '    observer_probe: { signal_types: [phases.has("active") ? "meeting_started" : null, participants.some((row) => row.speaking || row.active_speaker) ? "speaker_started" : null, phases.has("ended") ? "meeting_ended" : null].filter(Boolean) },',
    '    missing_required_coverage: [!phases.has("active") ? "active_speaker" : null, !phases.has("ended") ? "meeting_ended" : null].filter(Boolean),',
    '  };',
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
    `const DEFAULT_BASE_URL = ${JSON.stringify(baseUrl)};`,
    `const CLIENT_CALL_TYPE = ${JSON.stringify(messagePrefix)};`,
    `const ATTACHED_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached)};`,
    `const STATUS_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status)};`,
    `const SERVICE_STATUS_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.service_status)};`,
    `const CONFIGURE_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.configure)};`,
    `const CAPTURE_EVIDENCE_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.capture_evidence)};`,
    `const OBSERVE_CANDIDATES_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates)};`,
    `const PREFLIGHT_CURRENT_WINDOW_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window)};`,
    `const PREFLIGHT_CANDIDATES_TYPE = ${JSON.stringify(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates)};`,
    `const STATUS_STORAGE_KEY = ${JSON.stringify(MEETING_APP_EXTENSION_STATUS_STORAGE_KEY)};`,
    `const BASE_URL_STORAGE_KEY = ${JSON.stringify(MEETING_APP_EXTENSION_BASE_URL_STORAGE_KEY)};`,
    `const ENDPOINTS = ${json(endpoints)};`,
    'const RUNTIME_EVENT_ENDPOINT = "/api/meeting-platform/runtime-events";',
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
    'function normalizedBaseUrl(value) {',
    '  try {',
    '    const url = new URL(String(value));',
    '    if (url.protocol !== "http:" && url.protocol !== "https:") return null;',
    '    return url.toString().replace(/\\\/$/, "");',
    '  } catch {',
    '    return null;',
    '  }',
    '}',
    '',
    'async function configuredBaseUrl() {',
    '  return normalizedBaseUrl(await getStorageValue(BASE_URL_STORAGE_KEY)) ?? DEFAULT_BASE_URL;',
    '}',
    '',
    'async function getJson(path) {',
    '  const baseUrl = await configuredBaseUrl();',
    '  const response = await fetch(`${baseUrl}${path}`, { headers: { accept: "application/json" } });',
    '  const text = await response.text();',
    '  let body;',
    '  try { body = text ? JSON.parse(text) : undefined; } catch { body = text; }',
    '  return { ok: response.ok, status: response.status, body, base_url: baseUrl };',
    '}',
    '',
    'async function postJson(path, payload) {',
    '  const baseUrl = await configuredBaseUrl();',
    '  const response = await fetch(`${baseUrl}${path}`, {',
    '    method: "POST",',
    '    headers: { "content-type": "application/json" },',
    '    body: JSON.stringify(payload),',
    '  });',
    '  const text = await response.text();',
    '  let body;',
    '  try { body = text ? JSON.parse(text) : undefined; } catch { body = text; }',
    '  return { ok: response.ok, status: response.status, body, base_url: baseUrl };',
    '}',
    '',
    'function postRuntimeEvent(action, input = {}) {',
    '  return postJson(RUNTIME_EVENT_ENDPOINT, {',
    '    schema: "meeting_platform_runtime_event",',
    '    schema_version: 1,',
    '    action,',
    '    source: input.source ?? "meeting_app_extension_background",',
    '    sent_at_ms: Date.now(),',
    '    ...input,',
    '  });',
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
    '    Promise.all([getStorageValue(STATUS_STORAGE_KEY), configuredBaseUrl()]).then(([stored, baseUrl]) => {',
    '      sendResponse({ ok: true, type: STATUS_TYPE, attached: stored ?? lastAttached, base_url: baseUrl });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === CONFIGURE_TYPE) {',
    '    const baseUrl = normalizedBaseUrl(message.base_url ?? message.baseUrl);',
    '    if (!baseUrl) {',
    '      sendResponse({ ok: false, type: CONFIGURE_TYPE, reason: "invalid_base_url" });',
    '      return false;',
    '    }',
    '    setStorageValue(BASE_URL_STORAGE_KEY, baseUrl).then((storage) => {',
    '      sendResponse({ ok: storage.stored === true, type: CONFIGURE_TYPE, base_url: baseUrl, storage });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === SERVICE_STATUS_TYPE) {',
    '    getJson("/api/state").then((service) => {',
    '      sendResponse({ ok: service.ok === true, type: SERVICE_STATUS_TYPE, service });',
    '    }).catch(async (error) => {',
    '      sendResponse({',
    '        ok: false,',
    '        type: SERVICE_STATUS_TYPE,',
    '        base_url: await configuredBaseUrl(),',
    '        reason: "timeline_service_unreachable",',
    '        error: String(error?.message ?? error),',
    '      });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === CAPTURE_EVIDENCE_TYPE) {',
    '    targetTabIdFor(message, sender).then((tabId) => forwardToTab(tabId, {',
    '      ...message,',
    '      type: CAPTURE_EVIDENCE_TYPE,',
    '      captured_at_ms: message.captured_at_ms ?? message.capturedAtMs ?? Date.now(),',
    '    }, message.frame_id != null || message.frameId != null ? { frameId: message.frame_id ?? message.frameId } : {})).then((result) => {',
    '      sendResponse({ ok: result.ok === true, type: CAPTURE_EVIDENCE_TYPE, ...result });',
    '    }).catch((error) => {',
    '      sendResponse({',
    '        ok: false,',
    '        type: CAPTURE_EVIDENCE_TYPE,',
    '        reason: "capture_evidence_failed",',
    '        error: String(error?.message ?? error),',
    '      });',
    '    });',
    '    return true;',
    '  }',
    '  if (message?.type === OBSERVE_CANDIDATES_TYPE) {',
    '    observedTabsFor(message).then((tabs) => {',
    '      const input = candidateObservationInput(message, sender, tabs);',
    '      return postRuntimeEvent("observe_platform_candidates", input);',
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

export function buildMeetingAppExtensionPopupHtml(options = {}) {
  const title = firstNonEmpty(options.popupTitle, options.popup_title, 'Meeting Timeline Adapter');
  return [
    '<!doctype html>',
    '<html lang="en">',
    '<head>',
    '  <meta charset="utf-8">',
    '  <meta name="viewport" content="width=device-width, initial-scale=1">',
    `  <title>${title}</title>`,
    '  <link rel="stylesheet" href="popup.css">',
    '</head>',
    '<body>',
    '  <header class="header">',
    `    <h1>${title}</h1>`,
    '    <div class="service-line"><span id="service-dot" class="dot pending"></span><span id="service-status">Checking service</span></div>',
    '  </header>',
    '  <main>',
    '    <section class="summary" aria-label="Current meeting status">',
    '      <div><span class="label">Platform</span><strong id="platform">Unknown</strong></div>',
    '      <div><span class="label">Page adapter</span><strong id="adapter-status">Checking</strong></div>',
    '    </section>',
    '    <section class="configuration" aria-label="Timeline service configuration">',
    '      <label for="base-url">Timeline service</label>',
    '      <div class="input-row">',
    '        <input id="base-url" type="url" inputmode="url" spellcheck="false" autocomplete="off">',
    '        <button id="save-url" type="button" title="Save timeline service address">Save</button>',
    '      </div>',
    '    </section>',
    '    <section class="actions" aria-label="Adapter actions">',
    '      <button id="check" class="primary" type="button">Check current meeting</button>',
    '      <div class="button-grid">',
    '        <button id="capture-active" type="button">Capture active</button>',
    '        <button id="capture-ended" type="button">Capture ended</button>',
    '        <button id="test-mark" type="button">Send test mark</button>',
    '        <button id="export" type="button">Export evidence</button>',
    '      </div>',
    '    </section>',
    '    <section class="result" aria-live="polite">',
    '      <div class="result-title"><span>Latest result</span><span id="result-time"></span></div>',
    '      <pre id="result">No checks have run.</pre>',
    '    </section>',
    '  </main>',
    '  <script src="popup.js"></script>',
    '</body>',
    '</html>',
  ].join('\n');
}

export function buildMeetingAppExtensionPopupCss() {
  return [
    ':root { color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }',
    '* { box-sizing: border-box; }',
    'body { width: 390px; margin: 0; color: #17202a; background: #f7f8fa; font-size: 13px; }',
    'button, input { font: inherit; letter-spacing: 0; }',
    '.header { padding: 18px 18px 14px; color: #f8fafc; background: #20262e; }',
    'h1 { margin: 0 0 8px; font-size: 17px; line-height: 1.25; letter-spacing: 0; }',
    '.service-line { display: flex; align-items: center; gap: 7px; color: #cbd5e1; font-size: 12px; }',
    '.dot { width: 8px; height: 8px; border-radius: 50%; background: #94a3b8; flex: 0 0 auto; }',
    '.dot.ok { background: #33b875; }',
    '.dot.error { background: #e05252; }',
    'main { background: #fff; }',
    'section { padding: 14px 18px; border-bottom: 1px solid #e5e7eb; }',
    '.summary { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }',
    '.summary div { min-width: 0; }',
    '.label { display: block; margin-bottom: 3px; color: #687386; font-size: 11px; text-transform: uppercase; }',
    '.summary strong { display: block; overflow-wrap: anywhere; font-size: 13px; }',
    '.configuration label { display: block; margin-bottom: 7px; color: #4b5563; font-weight: 600; }',
    '.input-row { display: grid; grid-template-columns: minmax(0, 1fr) auto; gap: 7px; }',
    'input { min-width: 0; height: 34px; padding: 0 10px; border: 1px solid #cbd1d8; border-radius: 6px; color: #17202a; background: #fff; outline: none; }',
    'input:focus { border-color: #2672d9; box-shadow: 0 0 0 2px rgba(38, 114, 217, .14); }',
    'button { min-height: 34px; padding: 7px 11px; border: 1px solid #cbd1d8; border-radius: 6px; color: #24303d; background: #fff; cursor: pointer; }',
    'button:hover { background: #f2f4f7; }',
    'button:focus-visible { outline: 2px solid #2672d9; outline-offset: 1px; }',
    'button:disabled { cursor: wait; opacity: .55; }',
    'button.primary { width: 100%; border-color: #2166bd; color: #fff; background: #2166bd; font-weight: 650; }',
    'button.primary:hover { background: #1b57a3; }',
    '.button-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 7px; margin-top: 8px; }',
    '.result { border-bottom: 0; background: #f7f8fa; }',
    '.result-title { display: flex; justify-content: space-between; gap: 12px; margin-bottom: 7px; color: #596579; font-size: 11px; font-weight: 650; text-transform: uppercase; }',
    'pre { min-height: 76px; max-height: 190px; margin: 0; padding: 10px; overflow: auto; border: 1px solid #dde1e6; border-radius: 6px; color: #273342; background: #fff; font: 11px/1.45 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; white-space: pre-wrap; overflow-wrap: anywhere; }',
  ].join('\n');
}

export function buildMeetingAppExtensionPopupSource(options = {}) {
  const baseUrl = String(firstNonEmpty(options.baseUrl, options.base_url, 'http://localhost:8787')).replace(/\/+$/, '');
  const supportedHosts = Object.fromEntries(platformList(options).map((platform) => [
    platform,
    MEETING_APP_EXTENSION_PROFILES[platform].host_permissions.map((pattern) => matchPatternHost(pattern)).filter(Boolean),
  ]));
  return [
    `const DEFAULT_BASE_URL = ${JSON.stringify(baseUrl)};`,
    `const PLATFORM_HOSTS = ${json(supportedHosts)};`,
    `const TYPES = ${json(MEETING_APP_EXTENSION_MESSAGE_TYPES)};`,
    '',
    'const element = (id) => document.getElementById(id);',
    'const ui = {',
    '  serviceDot: element("service-dot"),',
    '  serviceStatus: element("service-status"),',
    '  platform: element("platform"),',
    '  adapterStatus: element("adapter-status"),',
    '  baseUrl: element("base-url"),',
    '  result: element("result"),',
    '  resultTime: element("result-time"),',
    '};',
    'let currentTab = null;',
    'let currentPlatform = "unknown";',
    '',
    'function runtimeSend(message) {',
    '  return new Promise((resolve) => {',
    '    chrome.runtime.sendMessage(message, (response) => {',
    '      const error = chrome.runtime.lastError;',
    '      resolve(error ? { ok: false, reason: "runtime_error", error: error.message } : (response ?? { ok: false, reason: "empty_response" }));',
    '    });',
    '  });',
    '}',
    '',
    'function activeTab() {',
    '  return new Promise((resolve) => chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => resolve(tabs?.[0] ?? null)));',
    '}',
    '',
    'function inferPlatform(urlValue = "") {',
    '  let host = "";',
    '  try { host = new URL(urlValue).hostname; } catch { return "unknown"; }',
    '  for (const [platform, hosts] of Object.entries(PLATFORM_HOSTS)) {',
    '    if (hosts.some((item) => host === item || host.endsWith(`.${item}`))) return platform;',
    '  }',
    '  return "unknown";',
    '}',
    '',
    'function displayPlatform(platform) {',
    '  return ({ google_meet: "Google Meet", microsoft_teams: "Microsoft Teams", zoom: "Zoom" })[platform] ?? "Unsupported page";',
    '}',
    '',
    'function showResult(label, value) {',
    '  ui.resultTime.textContent = new Date().toLocaleTimeString();',
    '  ui.result.textContent = `${label}\n${JSON.stringify(value, null, 2)}`;',
    '}',
    '',
    'function setBusy(busy) {',
    '  document.querySelectorAll("button").forEach((button) => { button.disabled = busy; });',
    '}',
    '',
    'async function checkService() {',
    '  const status = await runtimeSend({ type: TYPES.service_status });',
    '  ui.serviceDot.className = `dot ${status.ok ? "ok" : "error"}`;',
    '  ui.serviceStatus.textContent = status.ok ? `Connected to ${status.service?.base_url ?? ui.baseUrl.value}` : `Service unavailable at ${status.base_url ?? ui.baseUrl.value}`;',
    '  return status;',
    '}',
    '',
    'async function checkCurrentMeeting(options = {}) {',
    '  currentTab = await activeTab();',
    '  currentPlatform = inferPlatform(currentTab?.url);',
    '  ui.platform.textContent = displayPlatform(currentPlatform);',
    '  if (!currentTab || currentPlatform === "unknown") {',
    '    ui.adapterStatus.textContent = "Not available";',
    '    const result = { ok: false, reason: "open_supported_meeting_page", url: currentTab?.url };',
    '    if (options.show !== false) showResult("Current meeting check", result);',
    '    return result;',
    '  }',
    '  const result = await runtimeSend({',
    '    type: TYPES.preflight_current_window,',
    '    tab_id: currentTab.id,',
    '    platform: currentPlatform,',
    '    url: currentTab.url,',
    '    title: currentTab.title,',
    '    options: { requireSpeakerTrack: true },',
    '  });',
    '  const accepted = result.body?.result?.accepted === true;',
    '  ui.adapterStatus.textContent = accepted ? "Ready" : (result.ok ? "Needs evidence" : "Not attached");',
    '  if (options.show !== false) showResult("Current meeting check", result);',
    '  return result;',
    '}',
    '',
    'async function refresh() {',
    '  const extensionStatus = await runtimeSend({ type: TYPES.extension_status });',
    '  ui.baseUrl.value = extensionStatus.base_url ?? DEFAULT_BASE_URL;',
    '  await Promise.all([checkService(), checkCurrentMeeting({ show: false })]);',
    '}',
    '',
    'async function ensureOriginPermission(baseUrl) {',
    '  const url = new URL(baseUrl);',
    '  const origins = [`${url.protocol}//${url.host}/*`];',
    '  const allowed = await chrome.permissions.contains({ origins });',
    '  return allowed || chrome.permissions.request({ origins });',
    '}',
    '',
    'async function saveBaseUrl() {',
    '  const value = ui.baseUrl.value.trim().replace(/\\\/+$/, "");',
    '  let parsed;',
    '  try { parsed = new URL(value); } catch { showResult("Configuration", { ok: false, reason: "invalid_url" }); return; }',
    '  if (!["http:", "https:"].includes(parsed.protocol)) { showResult("Configuration", { ok: false, reason: "http_or_https_required" }); return; }',
    '  const permission = await ensureOriginPermission(value);',
    '  if (!permission) { showResult("Configuration", { ok: false, reason: "origin_permission_denied" }); return; }',
    '  const result = await runtimeSend({ type: TYPES.configure, base_url: value });',
    '  showResult("Configuration", result);',
    '  await checkService();',
    '}',
    '',
    'async function captureEvidence(action) {',
    '  const preflight = await checkCurrentMeeting({ show: false });',
    '  if (!currentTab || currentPlatform === "unknown") { showResult("Evidence capture", preflight); return; }',
    '  const result = await runtimeSend({ type: TYPES.capture_evidence, action, tab_id: currentTab.id, platform: currentPlatform });',
    '  showResult(`Evidence: ${action}`, result);',
    '  return result;',
    '}',
    '',
    'async function exportEvidence() {',
    '  const response = await captureEvidence("evidence_package");',
    '  const payload = response?.body?.result;',
    '  if (!payload) return;',
    '  const blob = new Blob([`${JSON.stringify(payload, null, 2)}\n`], { type: "application/json" });',
    '  const url = URL.createObjectURL(blob);',
    '  const link = document.createElement("a");',
    '  link.href = url;',
    '  link.download = `meeting-adapter-evidence-${currentPlatform}-${Date.now()}.json`;',
    '  link.click();',
    '  setTimeout(() => URL.revokeObjectURL(url), 1000);',
    '}',
    '',
    'async function sendTestMark() {',
    '  const preflight = await checkCurrentMeeting({ show: false });',
    '  if (!currentTab || currentPlatform === "unknown") { showResult("Test mark", preflight); return; }',
    '  const capturedAtMs = Date.now();',
    '  const meetingId = preflight.body?.result?.meeting_id;',
    '  const result = await runtimeSend({',
    '    type: TYPES.client_call,',
    '    method: "insertMark",',
    '    platform: currentPlatform,',
    '    captured_at_ms: capturedAtMs,',
    '    url: currentTab.url,',
    '    input: {',
    '      id: `adapter-test-${capturedAtMs}`,',
    '      meeting_id: meetingId,',
    '      captured_at_ms: capturedAtMs,',
    '      label: "Adapter test mark",',
    '      source: "meeting_timeline_extension_popup",',
    '      platform: currentPlatform,',
    '    },',
    '  });',
    '  showResult("Test mark", result);',
    '}',
    '',
    'async function run(buttonAction) {',
    '  setBusy(true);',
    '  try { await buttonAction(); } catch (error) { showResult("Error", { ok: false, error: String(error?.message ?? error) }); }',
    '  finally { setBusy(false); }',
    '}',
    '',
    'element("save-url").addEventListener("click", () => run(saveBaseUrl));',
    'element("check").addEventListener("click", () => run(checkCurrentMeeting));',
    'element("capture-active").addEventListener("click", () => run(() => captureEvidence("capture_active")));',
    'element("capture-ended").addEventListener("click", () => run(() => captureEvidence("capture_ended")));',
    'element("test-mark").addEventListener("click", () => run(sendTestMark));',
    'element("export").addEventListener("click", () => run(exportEvidence));',
    'run(refresh);',
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
  const popupEntry = firstNonEmpty(
    options.popupEntry,
    options.popup_entry,
    'src/popup.entry.mjs',
  );
  const popupOutput = firstNonEmpty(
    options.popupScript,
    options.popup_script,
    'popup.js',
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
    '  {',
    `    entryPoints: [${JSON.stringify(popupEntry)}],`,
    `    outfile: ${JSON.stringify(popupOutput)},`,
    '    bundle: true,',
    '    platform: "browser",',
    '    format: "iife",',
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
  const popupFile = firstNonEmpty(options.popupScript, options.popup_script, 'popup.js');
  const emittedFiles = includeLiveCapture
    ? `\`${scriptFile}\`, \`${liveCaptureFile}\`, \`${backgroundFile}\`, and \`${popupFile}\``
    : `\`${scriptFile}\`, \`${backgroundFile}\`, and \`${popupFile}\``;
  const liveCaptureDocs = includeLiveCapture ? [
    '',
    'When the extension is loaded, `window.__meetingTimelineLiveCapture` exposes `captureActive()`, `captureEnded()`, `exportRecords()`, `evidencePackage()`, and `diagnose()` for live DOM validation.',
  ] : [];
  return [
    '# Meeting Timeline Browser Extension',
    '',
    'This installable Chrome/Edge extension connects supported meeting web apps to Meeting Timeline.',
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
    'Open `chrome://extensions`, enable Developer mode, choose Load unpacked, and select this directory after the build completes.',
    '',
    'Pin the extension, open a supported meeting page, and use the popup to check the current adapter, capture active/ended evidence, send a test mark, or export the evidence package.',
    '',
    'The generated background worker forwards timeline client calls to the configured timeline service base URL. Changing the URL from the popup requests access only to the selected origin.',
    '',
    'Protocol messages use `meeting_timeline.extension_attached`, `meeting_timeline.extension_status`, `meeting_timeline.service_status`, `meeting_timeline.configure`, `meeting_timeline.capture_evidence`, `meeting_timeline.observe_candidates`, `meeting_timeline.preflight_current_window`, `meeting_timeline.preflight_candidates`, and `meeting_timeline.client_call`; build custom callers with the SDK message helpers instead of hardcoded strings.',
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
  const popupFile = firstNonEmpty(options.popupScript, options.popup_script, 'popup.js');
  const popupHtmlFile = firstNonEmpty(options.popupHtml, options.popup_html, 'popup.html');
  const popupCssFile = firstNonEmpty(options.popupCss, options.popup_css, 'popup.css');
  const contentScriptEntry = firstNonEmpty(options.contentScriptEntry, options.content_script_entry, 'src/content-script.entry.mjs');
  const liveCaptureEntry = firstNonEmpty(options.liveCaptureEntry, options.live_capture_entry, 'src/live-capture.entry.mjs');
  const backgroundEntry = firstNonEmpty(options.backgroundEntry, options.background_entry, 'src/background.entry.mjs');
  const popupEntry = firstNonEmpty(options.popupEntry, options.popup_entry, 'src/popup.entry.mjs');
  const baseUrl = firstNonEmpty(options.baseUrl, options.base_url);
  const endpointPermission = endpointOriginPattern(baseUrl);
  const packageJson = buildMeetingAppExtensionPackageJson(options);
  const buildSource = buildMeetingAppExtensionBuildSource({
    ...options,
    outputScript: scriptFile,
    backgroundScript: backgroundFile,
    contentScriptEntry,
    backgroundEntry,
    popupEntry,
    popupScript: popupFile,
  });
  const manifestOverrides = options.manifestOverrides ?? options.manifest_overrides ?? {};
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
    name: firstNonEmpty(options.name, 'Meeting Timeline Adapter'),
    description: firstNonEmpty(options.description, 'Connects Google Meet, Microsoft Teams, and Zoom web meetings to Meeting Timeline.'),
    manifestOverrides: {
      ...manifestOverrides,
      background: {
        service_worker: backgroundFile,
        type: 'module',
        ...(manifestOverrides.background ?? {}),
      },
      action: {
        default_title: 'Meeting Timeline Adapter',
        default_popup: popupHtmlFile,
        ...(manifestOverrides.action ?? {}),
      },
      optional_host_permissions: unique([
        'http://*/*',
        'https://*/*',
        ...asArray(manifestOverrides.optional_host_permissions),
      ]),
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
  const popupSource = buildMeetingAppExtensionPopupSource(options);
  const popupHtml = buildMeetingAppExtensionPopupHtml(options);
  const popupCss = buildMeetingAppExtensionPopupCss(options);
  const readme = buildMeetingAppExtensionReadme({
    ...options,
    js: scriptFile,
    backgroundScript: backgroundFile,
    popupScript: popupFile,
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
      sourceFile(popupEntry, popupSource, 'popup_entry', 'text/javascript'),
      sourceFile(popupHtmlFile, popupHtml, 'popup_html', 'text/html'),
      sourceFile(popupCssFile, popupCss, 'popup_stylesheet', 'text/css'),
      sourceFile('README.md', readme, 'readme', 'text/markdown'),
    ],
    bundle: {
      content_script_input: contentScriptEntry,
      content_script_output: scriptFile,
      live_capture_input: includeLiveCapture ? liveCaptureEntry : undefined,
      live_capture_output: includeLiveCapture ? liveCaptureFile : undefined,
      background_input: backgroundEntry,
      background_output: backgroundFile,
      popup_input: popupEntry,
      popup_output: popupFile,
      popup_html: popupHtmlFile,
      popup_css: popupCssFile,
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
    scaffold.bundle?.popup_input ?? 'src/popup.entry.mjs',
    scaffold.bundle?.popup_html ?? 'popup.html',
    scaffold.bundle?.popup_css ?? 'popup.css',
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
  const popupEntry = fileByPath(scaffold, scaffold.bundle?.popup_input ?? 'src/popup.entry.mjs');
  const popupHtml = fileByPath(scaffold, scaffold.bundle?.popup_html ?? 'popup.html');
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
  if (manifest.action?.default_popup !== (scaffold.bundle?.popup_html ?? 'popup.html')) {
    issues.push(issue('error', 'missing_extension_popup', 'Manifest action must open the generated adapter popup.', {
      expected: scaffold.bundle?.popup_html ?? 'popup.html',
      actual: manifest.action?.default_popup,
    }));
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
  if (!buildContent.includes(scaffold.bundle?.popup_input ?? 'src/popup.entry.mjs')) {
    issues.push(issue('error', 'build_missing_popup_entry', 'build.mjs does not include the popup entry.'));
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
  if (!contentEntryContent.includes('installMeetingAppLocalContentRuntime')) {
    issues.push(issue('error', 'content_entry_missing_bridge', 'Content script entry does not install the SDK bridge.'));
  }
  if (!contentEntryContent.includes('startRuntime: true')) {
    issues.push(issue('error', 'content_entry_not_starting_runtime', 'Content script entry does not start the browser runtime.'));
  }
  if (!contentEntryContent.includes('bindOperatorReferenceCapture')) {
    issues.push(issue('error', 'content_entry_missing_operator_reference_capture', 'Content script entry does not capture join/leave operator reference timestamps.'));
  }
  const liveCaptureContent = liveCaptureEntry?.content ?? '';
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('__meetingTimelineLiveCapture')) {
    issues.push(issue('error', 'live_capture_missing_global', 'Live capture entry does not expose __meetingTimelineLiveCapture.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('captureMeetingAppDomSnapshot')) {
    issues.push(issue('error', 'live_capture_missing_dom_capture', 'Live capture entry does not capture meeting app DOM snapshots.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('meeting_app_live_evidence_package')) {
    issues.push(issue('error', 'live_capture_missing_evidence_package', 'Live capture entry does not build live evidence packages.'));
  }
  if (scaffold.bundle?.live_capture_input && !liveCaptureContent.includes('meeting_app_dom_adaptation_diagnosis')) {
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
  if (!backgroundContent.includes('postRuntimeEvent')) {
    issues.push(issue('error', 'background_missing_runtime_event_client', 'Background entry does not post the platform runtime event envelope.'));
  }
  if (!backgroundContent.includes('setStorageValue')) {
    issues.push(issue('error', 'background_missing_diagnostic_storage', 'Background entry does not persist extension diagnostics.'));
  }
  if (!backgroundContent.includes('meeting_timeline.service_status')) {
    issues.push(issue('error', 'background_missing_service_status', 'Background entry does not expose timeline service diagnostics.'));
  }
  if (!backgroundContent.includes('meeting_timeline.capture_evidence')) {
    issues.push(issue('error', 'background_missing_evidence_capture', 'Background entry does not forward real-page evidence capture requests.'));
  }
  const popupContent = popupEntry?.content ?? '';
  if (!popupContent.includes('preflight_current_window') || !popupContent.includes('capture_evidence')) {
    issues.push(issue('error', 'popup_missing_adapter_controls', 'Popup entry does not expose preflight and evidence capture controls.'));
  }
  if (!popupHtml?.content?.includes('popup.js')) {
    issues.push(issue('error', 'popup_html_missing_script', 'Popup HTML does not load the generated popup script.'));
  }
  for (const endpoint of ['/api/meeting-session/start', '/api/meeting-session/end', '/api/meeting-platform/p0-reference', '/api/annotations', '/api/annotations/batch']) {
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
