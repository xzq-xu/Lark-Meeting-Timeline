import assert from 'node:assert/strict';

import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  MEETING_APP_EXTENSION_PLATFORM_KEYS,
  MEETING_APP_EXTENSION_PROFILES,
  MEETING_APP_EXTENSION_STATUS_STORAGE_KEY,
  MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS,
  assertMeetingAppExtensionScaffold,
  buildMeetingAppExtensionAttachedMessage,
  buildMeetingAppContentScriptManifest,
  buildMeetingAppExtensionBackgroundSource,
  buildMeetingAppExtensionBuildSource,
  buildMeetingAppExtensionClientCallMessage,
  buildMeetingAppExtensionContentScriptSource,
  buildMeetingAppExtensionCurrentWindowPreflightMessage,
  buildMeetingAppExtensionInstallPlan,
  buildMeetingAppExtensionLiveCaptureSource,
  buildMeetingAppExtensionMatchPatterns,
  buildMeetingAppExtensionObserveCandidatesMessage,
  buildMeetingAppExtensionPackageJson,
  buildMeetingAppExtensionPreflightCandidatesMessage,
  buildMeetingAppExtensionScaffold,
  buildMeetingAppExtensionScaffoldAcceptanceReport,
  buildMeetingAppExtensionStatusMessage,
  meetingAppExtensionTimelineEndpoint,
  meetingAppExtensionProfile,
  normalizeMeetingAppExtensionMessageType,
  normalizeMeetingAppExtensionPlatform,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-extension.mjs';
import {
  buildMeetingAppDomAdaptationDiagnosis,
  buildMeetingAppLiveEvidencePackage,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import { createMeetingAppSnapshotRecorder } from '../packages/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder.mjs';

function installGeneratedBackground(source) {
  const originalChrome = globalThis.chrome;
  const originalFetch = globalThis.fetch;
  const listeners = [];
  const storage = {};
  const fetchCalls = [];
  const tabMessages = [];
  const meetingTabs = [
    {
      id: 7,
      windowId: 1,
      active: true,
      audible: false,
      pinned: false,
      discarded: false,
      status: 'complete',
      title: 'Google Meet',
      url: 'https://meet.google.com/abc-defg-hij',
    },
    {
      id: 8,
      windowId: 1,
      active: false,
      audible: true,
      pinned: false,
      discarded: false,
      status: 'complete',
      title: 'Teams meeting',
      url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    },
    {
      id: 9,
      windowId: 1,
      active: false,
      audible: false,
      pinned: false,
      discarded: false,
      status: 'complete',
      title: 'Zoom meeting',
      url: 'https://us06web.zoom.us/wc/987654321/start',
    },
    {
      id: 10,
      windowId: 1,
      active: false,
      audible: false,
      pinned: false,
      discarded: false,
      status: 'complete',
      title: 'Lark meeting',
      url: 'https://vc.feishu.cn/j/123456',
    },
  ];
  const platformForUrl = (url = '') => {
    if (url.includes('teams.microsoft.com')) return 'microsoft_teams';
    if (url.includes('zoom.us')) return 'zoom';
    if (url.includes('feishu.cn') || url.includes('larksuite.com') || url.includes('larkoffice.com')) return 'lark';
    return 'google_meet';
  };
  globalThis.chrome = {
    runtime: {
      onMessage: {
        addListener(listener) {
          listeners.push(listener);
        },
      },
    },
    storage: {
      local: {
        set(value, callback) {
          Object.assign(storage, value);
          callback?.();
        },
        get(key, callback) {
          const keys = Array.isArray(key) ? key : [key];
          const result = Object.fromEntries(keys.map((item) => [item, storage[item]]));
          callback?.(result);
        },
      },
    },
    tabs: {
      query(_query, callback) {
        callback?.(meetingTabs);
      },
      sendMessage(tabId, message, options, callback) {
        const sendOptions = typeof options === 'function' ? {} : (options ?? {});
        const sendCallback = typeof options === 'function' ? options : callback;
        tabMessages.push({ tabId, message, options: sendOptions });
        const tab = meetingTabs.find((item) => item.id === tabId);
        sendCallback?.({
          handled: true,
          action: 'preflightCurrentWindow',
          result: {
            schema: 'meeting_platform_adapter_preflight',
            accepted: true,
            platform: message.platform ?? message.input?.platform ?? platformForUrl(message.input?.url ?? message.url ?? tab?.url),
            readiness: { realtime_annotation_ready: true },
          },
        });
      },
    },
  };
  globalThis.fetch = async (url, init) => {
    fetchCalls.push({ url, init });
    return new Response(JSON.stringify({ ok: true, stored: true }), {
      status: 201,
      headers: { 'content-type': 'application/json' },
    });
  };
  try {
    const executableSource = source.replace(
      "import { createMeetingPlatformRuntimeEventClient } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event';",
      `function createMeetingPlatformRuntimeEventClient(options = {}) {
        return {
          async observePlatformCandidates(input = {}) {
            const response = await fetch(new URL('/api/meeting-platform/runtime-events', options.baseUrl).toString(), {
              method: 'POST',
              headers: { 'content-type': 'application/json' },
              body: JSON.stringify({
                type: 'meeting_platform_runtime_event',
                schema: 'meeting_platform_runtime_event',
                schema_version: 1,
                action: 'observe_platform_candidates',
                source: options.source,
                sent_at_ms: input.captured_at_ms ?? Date.now(),
                ...input,
              }),
            });
            const body = await response.json();
            return { ok: response.ok, status: response.status, body };
          },
        };
      }`,
    );
    Function(executableSource)();
  } catch (error) {
    globalThis.chrome = originalChrome;
    globalThis.fetch = originalFetch;
    throw error;
  }
  assert.equal(listeners.length, 1);
  const listener = listeners[0];
  return {
    storage,
    fetchCalls,
    tabMessages,
    send(message, sender = {}) {
      return new Promise((resolve) => {
        const returned = listener(message, sender, resolve);
        if (returned === false) resolve({ handled: false });
      });
    },
    restore() {
      globalThis.chrome = originalChrome;
      globalThis.fetch = originalFetch;
    },
  };
}

function installGeneratedContentScript(source, options = {}) {
  const originalChrome = globalThis.chrome;
  const originalLocation = globalThis.location;
  const originalBridge = globalThis.__meetingTimelineBridge;
  const sentMessages = [];
  const bridgeInstalls = [];
  globalThis.location = {
    hostname: options.hostname ?? 'meet.google.com',
    href: options.href ?? 'https://meet.google.com/abc-defg-hij',
  };
  globalThis.chrome = {
    runtime: {
      sendMessage(message, callback) {
        sentMessages.push(message);
        callback?.({ ok: true, echoed_type: message.type, method: message.method });
      },
      lastError: undefined,
    },
  };
  const runnableSource = source.replace(/^import .+;\n\n?/, '');
  const installMeetingPlatformIntegrationContentScriptBridge = (client, bridgeOptions) => {
    const bridge = { client, options: bridgeOptions, installed: true };
    bridgeInstalls.push(bridge);
    return bridge;
  };
  try {
    Function(
      'installMeetingPlatformIntegrationContentScriptBridge',
      runnableSource,
    )(installMeetingPlatformIntegrationContentScriptBridge);
  } catch (error) {
    globalThis.chrome = originalChrome;
    globalThis.location = originalLocation;
    if (originalBridge === undefined) delete globalThis.__meetingTimelineBridge;
    else globalThis.__meetingTimelineBridge = originalBridge;
    throw error;
  }
  assert.equal(bridgeInstalls.length, 1);
  return {
    sentMessages,
    bridge: bridgeInstalls[0],
    restore() {
      globalThis.chrome = originalChrome;
      globalThis.location = originalLocation;
      if (originalBridge === undefined) delete globalThis.__meetingTimelineBridge;
      else globalThis.__meetingTimelineBridge = originalBridge;
    },
  };
}

function installGeneratedLiveCapture(source, options = {}) {
  const originalLocation = globalThis.location;
  const originalDocument = globalThis.document;
  const originalCustomEvent = globalThis.CustomEvent;
  const originalDispatchEvent = globalThis.dispatchEvent;
  const originalCapture = globalThis.__meetingTimelineLiveCapture;
  const dispatched = [];
  globalThis.location = {
    hostname: options.hostname ?? 'meet.google.com',
    href: options.href ?? 'https://meet.google.com/abc-defg-hij',
  };
  globalThis.document = {
    title: options.title ?? 'Design review - Google Meet',
  };
  globalThis.CustomEvent = class CustomEvent {
    constructor(type, init = {}) {
      this.type = type;
      this.detail = init.detail;
    }
  };
  globalThis.dispatchEvent = (event) => {
    dispatched.push(event);
    return true;
  };
  const captureMeetingAppDomSnapshot = (input = {}, captureOptions = {}) => {
    const ended = captureOptions.phase === 'ended';
    const controls = ended ? [{ label: 'Join now' }] : [{ label: 'Leave call' }];
    const participants = ended ? [] : [{ id: 'ada', name: 'Ada Lovelace', speaking: true }];
    return {
      schema: 'meeting_app_dom_capture',
      schema_version: 1,
      source: captureOptions.source,
      observedAtMs: captureOptions.observedAtMs,
      platform: captureOptions.platform,
      meeting_id: 'abc-defg-hij',
      url: input.url,
      title: input.title,
      page: {
        controls,
        buttons: controls,
        participants,
        tiles: participants,
      },
      capture: {
        profile: captureOptions.captureProfile,
        participant_count: participants.length,
      },
    };
  };
  const runnableSource = source.replace(/^(import .+\n)+\n?/, '');
  try {
    Function(
      'captureMeetingAppDomSnapshot',
      'createMeetingAppSnapshotRecorder',
      'buildMeetingAppLiveEvidencePackage',
      'buildMeetingAppDomAdaptationDiagnosis',
      runnableSource,
    )(
      captureMeetingAppDomSnapshot,
      createMeetingAppSnapshotRecorder,
      buildMeetingAppLiveEvidencePackage,
      buildMeetingAppDomAdaptationDiagnosis,
    );
  } catch (error) {
    globalThis.location = originalLocation;
    globalThis.document = originalDocument;
    globalThis.CustomEvent = originalCustomEvent;
    globalThis.dispatchEvent = originalDispatchEvent;
    if (originalCapture === undefined) delete globalThis.__meetingTimelineLiveCapture;
    else globalThis.__meetingTimelineLiveCapture = originalCapture;
    throw error;
  }
  assert.equal(typeof globalThis.__meetingTimelineLiveCapture?.captureActive, 'function');
  return {
    api: globalThis.__meetingTimelineLiveCapture,
    dispatched,
    restore() {
      globalThis.location = originalLocation;
      globalThis.document = originalDocument;
      globalThis.CustomEvent = originalCustomEvent;
      globalThis.dispatchEvent = originalDispatchEvent;
      if (originalCapture === undefined) delete globalThis.__meetingTimelineLiveCapture;
      else globalThis.__meetingTimelineLiveCapture = originalCapture;
    },
  };
}

assert.deepEqual(MEETING_APP_EXTENSION_PLATFORM_KEYS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);
assert.equal(MEETING_APP_EXTENSION_PROFILES.google_meet.matches.includes('https://meet.google.com/*'), true);
assert.equal(normalizeMeetingAppExtensionPlatform('google-meet'), 'google_meet');
assert.equal(normalizeMeetingAppExtensionPlatform('teams'), 'microsoft_teams');
assert.equal(normalizeMeetingAppExtensionPlatform('feishu'), 'lark');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call, 'meeting_timeline.client_call');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached, 'meeting_timeline.extension_attached');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status, 'meeting_timeline.extension_status');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates, 'meeting_timeline.observe_candidates');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window, 'meeting_timeline.preflight_current_window');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates, 'meeting_timeline.preflight_candidates');
assert.equal(MEETING_APP_EXTENSION_STATUS_STORAGE_KEY, 'meeting_timeline_extension_status');
assert.equal(MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS.insertMarks, '/api/annotations/batch');
assert.equal(normalizeMeetingAppExtensionMessageType('client_call'), MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call);
assert.equal(normalizeMeetingAppExtensionMessageType('attached'), MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached);
assert.equal(normalizeMeetingAppExtensionMessageType('extension-status'), MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status);
assert.equal(normalizeMeetingAppExtensionMessageType('observe-platform-candidates'), MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates);
assert.equal(normalizeMeetingAppExtensionMessageType('current-window-preflight'), MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window);
assert.equal(normalizeMeetingAppExtensionMessageType('preflight-platform-candidates'), MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates);
assert.equal(meetingAppExtensionTimelineEndpoint('insertMarks'), '/api/annotations/batch');
assert.throws(
  () => meetingAppExtensionTimelineEndpoint('deleteEverything'),
  /Unsupported meeting app extension client call method/,
);

const google = meetingAppExtensionProfile('google-meet');
assert.equal(google.platform, 'google_meet');
assert.deepEqual(google.matches, ['https://meet.google.com/*']);

const attachedMessage = buildMeetingAppExtensionAttachedMessage({
  platform: 'google-meet',
  capturedAtMs: 123,
  href: 'https://meet.google.com/abc-defg-hij',
});
assert.deepEqual(attachedMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
  platform: 'google_meet',
  captured_at_ms: 123,
  url: 'https://meet.google.com/abc-defg-hij',
});

const statusMessage = buildMeetingAppExtensionStatusMessage({
  requestId: 'status-001',
  capturedAtMs: 124,
});
assert.deepEqual(statusMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status,
  request_id: 'status-001',
  captured_at_ms: 124,
});

const currentWindowPreflightMessage = buildMeetingAppExtensionCurrentWindowPreflightMessage({
  platform: 'google-meet',
  requestId: 'preflight-001',
  capturedAtMs: 125,
  href: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
  options: { requireSpeakerTrack: true },
});
assert.deepEqual(currentWindowPreflightMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
  request_id: 'preflight-001',
  captured_at_ms: 125,
  platform: 'google_meet',
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
  options: { requireSpeakerTrack: true },
});

const preflightCandidatesMessage = buildMeetingAppExtensionPreflightCandidatesMessage({
  requestId: 'preflight-candidates-001',
  capturedAtMs: 126,
  query: { active: false },
  options: { requireSpeakerTrack: true },
});
assert.deepEqual(preflightCandidatesMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
  request_id: 'preflight-candidates-001',
  captured_at_ms: 126,
  query: { active: false },
  options: { requireSpeakerTrack: true },
});

const observeCandidatesMessage = buildMeetingAppExtensionObserveCandidatesMessage({
  requestId: 'observe-001',
  capturedAtMs: 1245,
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
});
assert.deepEqual(observeCandidatesMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
  request_id: 'observe-001',
  captured_at_ms: 1245,
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
});

const clientCallMessage = buildMeetingAppExtensionClientCallMessage('startMeeting', {
  meeting_id: 'meet-001',
}, {
  platform: 'google-meet',
  capturedAtMs: 125,
  url: 'https://meet.google.com/abc-defg-hij',
});
assert.deepEqual(clientCallMessage, {
  type: MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
  method: 'startMeeting',
  platform: 'google_meet',
  captured_at_ms: 125,
  url: 'https://meet.google.com/abc-defg-hij',
  input: {
    meeting_id: 'meet-001',
  },
});
assert.equal(buildMeetingAppExtensionClientCallMessage({
  method: 'endMeeting',
  input: { meeting_id: 'meet-001' },
}, {
  platform: 'google-meet',
  capturedAtMs: 126,
}).captured_at_ms, 126);

const patterns = buildMeetingAppExtensionMatchPatterns();
assert.equal(patterns.type, 'meeting_app_extension_match_patterns');
assert.equal(patterns.platforms.length, 5);
assert.equal(patterns.matches.includes('https://meet.google.com/*'), true);
assert.equal(patterns.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(patterns.matches.includes('https://*.zoom.com/*'), true);
assert.equal(patterns.matches.includes('https://vc.feishu.cn/*'), true);
assert.equal(patterns.matches.includes('https://*.webex.com/*'), true);
assert.equal(patterns.matches.includes('<all_urls>'), false);
assert.deepEqual(patterns.content_scripts[0].js, ['meeting-app-content-script.bundle.js']);

const customPatterns = buildMeetingAppExtensionMatchPatterns(['google_meet', 'google-meet'], {
  js: 'content.js',
  extraMatches: ['https://meet.google.com/*', 'https://meet.example.test/*'],
  extraHostPermissions: ['https://meet.example.test/*'],
});
assert.deepEqual(customPatterns.platforms, ['google_meet']);
assert.deepEqual(customPatterns.content_scripts[0].js, ['content.js']);
assert.equal(customPatterns.matches.filter((item) => item === 'https://meet.google.com/*').length, 1);
assert.equal(customPatterns.matches.includes('https://meet.example.test/*'), true);
assert.equal(customPatterns.host_permissions.includes('https://meet.example.test/*'), true);

const googleManifest = buildMeetingAppContentScriptManifest({
  platforms: ['google_meet'],
  js: ['content.js'],
  permissions: ['storage', 'storage'],
});
assert.equal(googleManifest.manifest_version, 3);
assert.deepEqual(googleManifest.permissions, ['storage']);
assert.deepEqual(googleManifest.host_permissions, ['https://meet.google.com/*']);
assert.deepEqual(googleManifest.content_scripts[0].matches, ['https://meet.google.com/*']);
assert.deepEqual(googleManifest.content_scripts[0].js, ['content.js']);
assert.equal(googleManifest.content_scripts[0].run_at, 'document_idle');

const plan = buildMeetingAppExtensionInstallPlan({
  platforms: ['google_meet', 'microsoft_teams'],
  js: ['content.js'],
});
assert.equal(plan.type, 'meeting_app_extension_install_plan');
assert.deepEqual(plan.platforms, ['google_meet', 'microsoft_teams']);
assert.equal(plan.content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(plan.meeting_app_content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
assert.equal(plan.platform_integration_runtime_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime');
assert.equal(plan.browser_runtime_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
assert.equal(plan.snapshot_recorder_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder');
assert.equal(plan.launch_gate_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate');
assert.equal(plan.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(plan.runtime_contract.live_capture_global, '__meetingTimelineLiveCapture');
assert.deepEqual(plan.runtime_contract.live_capture_methods, [
  'captureActive',
  'captureEnded',
  'exportRecords',
  'evidencePackage',
  'diagnose',
]);
assert.deepEqual(plan.runtime_contract.message_types, MEETING_APP_EXTENSION_MESSAGE_TYPES);
assert.equal(
  plan.runtime_contract.local_content_script_messages.includes(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window),
  true,
);
assert.equal(
  plan.runtime_contract.local_content_script_messages.includes(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates),
  true,
);
assert.equal(plan.runtime_contract.status_storage_key, MEETING_APP_EXTENSION_STATUS_STORAGE_KEY);
assert.deepEqual(plan.runtime_contract.timeline_endpoints, MEETING_APP_EXTENSION_TIMELINE_ENDPOINTS);
assert.deepEqual(plan.manifest.content_scripts[0].js, ['content.js']);

const contentScriptSource = buildMeetingAppExtensionContentScriptSource({
  platforms: ['google_meet', 'microsoft_teams'],
});
assert.match(contentScriptSource, /installMeetingPlatformIntegrationContentScriptBridge/);
assert.match(contentScriptSource, /platform-integration-runtime/);
assert.match(contentScriptSource, /meet\.google\.com/);
assert.match(contentScriptSource, /teams\.microsoft\.com/);
assert.match(contentScriptSource, /startRuntime: true/);
assert.match(contentScriptSource, /startOptions: \{ startRuntime: true \}/);

const contentScriptRuntime = installGeneratedContentScript(contentScriptSource);
try {
  assert.equal(contentScriptRuntime.bridge.options.platform, 'google_meet');
  assert.deepEqual(contentScriptRuntime.bridge.options.platforms, ['google_meet', 'microsoft_teams']);
  assert.equal(contentScriptRuntime.bridge.options.runtimePreset, 'google_meet');
  assert.equal(contentScriptRuntime.bridge.options.browser_runtime_preset, 'google_meet');
  assert.equal(contentScriptRuntime.bridge.options.source, 'meeting_app_extension');
  assert.equal(contentScriptRuntime.bridge.options.startRuntime, true);
  assert.deepEqual(contentScriptRuntime.bridge.options.startOptions, { startRuntime: true });
  assert.equal(globalThis.__meetingTimelineBridge, contentScriptRuntime.bridge);
  assert.equal(contentScriptRuntime.sentMessages.length, 1);
  assert.equal(contentScriptRuntime.sentMessages[0].type, MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached);
  assert.equal(contentScriptRuntime.sentMessages[0].platform, 'google_meet');

  const startResult = await contentScriptRuntime.bridge.client.startMeeting({
    meeting_id: 'meet-001',
  });
  assert.equal(startResult.ok, true);
  assert.equal(contentScriptRuntime.sentMessages.length, 2);
  assert.equal(contentScriptRuntime.sentMessages[1].type, MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call);
  assert.equal(contentScriptRuntime.sentMessages[1].method, 'startMeeting');
  assert.equal(contentScriptRuntime.sentMessages[1].platform, 'google_meet');
  assert.equal(contentScriptRuntime.sentMessages[1].input.meeting_id, 'meet-001');
} finally {
  contentScriptRuntime.restore();
}

const teamsContentScriptRuntime = installGeneratedContentScript(contentScriptSource, {
  hostname: 'teams.microsoft.com',
  href: 'https://teams.microsoft.com/l/meetup-join/demo',
});
try {
  assert.equal(teamsContentScriptRuntime.bridge.options.platform, 'microsoft_teams');
  assert.equal(teamsContentScriptRuntime.bridge.options.runtimePreset, 'microsoft_teams');
  assert.equal(teamsContentScriptRuntime.bridge.options.browser_runtime_preset, 'microsoft_teams');
  assert.equal(teamsContentScriptRuntime.sentMessages[0].platform, 'microsoft_teams');
} finally {
  teamsContentScriptRuntime.restore();
}

const liveCaptureSource = buildMeetingAppExtensionLiveCaptureSource({
  platforms: ['google_meet', 'microsoft_teams'],
});
assert.match(liveCaptureSource, /captureMeetingAppDomSnapshot/);
assert.match(liveCaptureSource, /createMeetingAppSnapshotRecorder/);
assert.match(liveCaptureSource, /buildMeetingAppLiveEvidencePackage/);
assert.match(liveCaptureSource, /buildMeetingAppDomAdaptationDiagnosis/);
assert.match(liveCaptureSource, /__meetingTimelineLiveCapture/);
assert.match(liveCaptureSource, /captureActive/);
assert.match(liveCaptureSource, /captureEnded/);
assert.match(liveCaptureSource, /diagnose/);
assert.match(liveCaptureSource, /meet\.google\.com/);
assert.match(liveCaptureSource, /teams\.microsoft\.com/);

const liveCaptureRuntime = installGeneratedLiveCapture(liveCaptureSource);
try {
  assert.equal(liveCaptureRuntime.dispatched.length, 1);
  assert.equal(liveCaptureRuntime.dispatched[0].type, 'meeting_timeline.live_capture_ready');
  assert.equal(liveCaptureRuntime.dispatched[0].detail.platform, 'google_meet');

  const activeRecord = liveCaptureRuntime.api.captureActive({
    capturedAtMs: 1_782_442_820_000,
  });
  assert.equal(activeRecord.platform, 'google_meet');
  assert.equal(activeRecord.phase, 'active');
  assert.equal(activeRecord.snapshot.meeting_id, 'abc-defg-hij');
  assert.equal(activeRecord.snapshot.capture.participant_count, 1);

  const endedRecord = liveCaptureRuntime.api.captureEnded({
    capturedAtMs: 1_782_442_830_000,
  });
  assert.equal(endedRecord.phase, 'ended');
  assert.equal(liveCaptureRuntime.api.getState().record_count, 2);

  const exported = liveCaptureRuntime.api.exportRecords({ id: 'live-capture-export' });
  assert.equal(exported.id, 'live-capture-export');
  assert.equal(exported.record_count, 2);
  assert.equal(exported.records[0].phase, 'active');

  const evidence = liveCaptureRuntime.api.evidencePackage({ packageId: 'live-capture-evidence' });
  assert.equal(evidence.id, 'live-capture-evidence');
  assert.equal(evidence.record_count, 2);
  assert.equal(evidence.platforms.includes('google_meet'), true);

  const diagnosis = liveCaptureRuntime.api.diagnose();
  assert.equal(diagnosis.type, 'meeting_app_dom_adaptation_diagnosis');
  assert.equal(diagnosis.platform, 'google_meet');
  assert.equal(diagnosis.accepted, true);
  assert.equal(diagnosis.production_ready, true);
  assert.equal(diagnosis.selector_probe.matched.controls, true);
  assert.equal(diagnosis.selector_probe.matched.participants, true);
  assert.equal(diagnosis.observer_probe.signal_types.includes('meeting_ended'), true);
} finally {
  liveCaptureRuntime.restore();
}

const backgroundSource = buildMeetingAppExtensionBackgroundSource({
  baseUrl: 'https://timeline.example.com/',
});
assert.match(backgroundSource, /const BASE_URL = "https:\/\/timeline\.example\.com";/);
assert.match(backgroundSource, /meeting_timeline\.extension_attached/);
assert.match(backgroundSource, /meeting_timeline\.extension_status/);
assert.match(backgroundSource, /meeting_timeline\.observe_candidates/);
assert.match(backgroundSource, /meeting_timeline\.preflight_current_window/);
assert.match(backgroundSource, /meeting_timeline\.preflight_candidates/);
assert.match(backgroundSource, /tabs\.sendMessage/);
assert.match(backgroundSource, /createMeetingPlatformRuntimeEventClient/);
assert.match(backgroundSource, /observePlatformCandidates/);
assert.match(backgroundSource, /STATUS_STORAGE_KEY/);
assert.match(backgroundSource, /setStorageValue/);
assert.match(backgroundSource, /\/api\/meeting-session\/start/);
assert.match(backgroundSource, /\/api\/annotations\/batch/);

const backgroundRuntime = installGeneratedBackground(backgroundSource);
try {
  const attachedResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_attached,
    platform: 'google_meet',
    captured_at_ms: 1_782_442_800_000,
    url: 'https://meet.google.com/abc-defg-hij',
  }, {
    tab: { id: 7 },
    frameId: 0,
    url: 'https://meet.google.com/abc-defg-hij',
  });
  assert.equal(attachedResponse.ok, true);
  assert.equal(attachedResponse.storage.stored, true);
  assert.equal(attachedResponse.attached.platform, 'google_meet');
  assert.equal(backgroundRuntime.storage[MEETING_APP_EXTENSION_STATUS_STORAGE_KEY].platform, 'google_meet');

  const statusResponse = await backgroundRuntime.send({ type: MEETING_APP_EXTENSION_MESSAGE_TYPES.extension_status });
  assert.equal(statusResponse.ok, true);
  assert.equal(statusResponse.attached.url, 'https://meet.google.com/abc-defg-hij');

  const clientCallResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.client_call,
    method: 'startMeeting',
    platform: 'google_meet',
    captured_at_ms: 1_782_442_810_000,
    url: 'https://meet.google.com/abc-defg-hij',
    input: {
      meeting_id: 'meet-001',
      title: 'Weekly sync',
    },
  }, {
    tab: { id: 7 },
    frameId: 0,
    url: 'https://meet.google.com/abc-defg-hij',
  });
  assert.equal(clientCallResponse.ok, true);
  assert.equal(clientCallResponse.status, 201);
  assert.equal(backgroundRuntime.fetchCalls.length, 1);
  assert.equal(backgroundRuntime.fetchCalls[0].url, 'https://timeline.example.com/api/meeting-session/start');
  const postedPayload = JSON.parse(backgroundRuntime.fetchCalls[0].init.body);
  assert.equal(postedPayload.platform, 'google_meet');
  assert.equal(postedPayload.meeting_id, 'meet-001');
  assert.equal(postedPayload.detector_source, 'meeting_app_extension');
  assert.equal(postedPayload.extension_sender.tab_id, 7);

  const observeResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.observe_candidates,
    captured_at_ms: 1_782_442_812_000,
  }, {
    tab: { id: 7 },
    frameId: 0,
    url: 'https://meet.google.com/abc-defg-hij',
  });
  assert.equal(observeResponse.ok, true);
  assert.equal(observeResponse.body.ok, true);
  assert.equal(backgroundRuntime.fetchCalls.length, 2);
  assert.equal(backgroundRuntime.fetchCalls[1].url, 'https://timeline.example.com/api/meeting-platform/runtime-events');
  const postedRuntimeEvent = JSON.parse(backgroundRuntime.fetchCalls[1].init.body);
  assert.equal(postedRuntimeEvent.action, 'observe_platform_candidates');
  assert.equal(postedRuntimeEvent.source, 'meeting_app_extension_background');
  assert.equal(postedRuntimeEvent.tabs[0].url, 'https://meet.google.com/abc-defg-hij');
  assert.equal(postedRuntimeEvent.extension_sender.tab_id, 7);

  const preflightResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
    request_id: 'preflight-001',
    options: { requireSpeakerTrack: true },
  }, {
    tab: { id: 7 },
    frameId: 0,
    url: 'https://meet.google.com/abc-defg-hij',
  });
  assert.equal(preflightResponse.ok, true);
  assert.equal(preflightResponse.tab_id, 7);
  assert.equal(preflightResponse.body.action, 'preflightCurrentWindow');
  assert.equal(preflightResponse.body.result.accepted, true);
  assert.equal(backgroundRuntime.tabMessages.length, 1);
  assert.equal(backgroundRuntime.tabMessages[0].tabId, 7);
  assert.equal(backgroundRuntime.tabMessages[0].message.type, MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window);
  assert.deepEqual(backgroundRuntime.tabMessages[0].message.options, { requireSpeakerTrack: true });
  assert.equal(backgroundRuntime.fetchCalls.length, 2);

  const larkPreflightResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window,
    tab_id: 8,
    platform: 'lark',
    input: {
      url: 'https://vc.feishu.cn/j/123456',
      title: 'Lark meeting',
    },
    options: { requireSpeakerTrack: false },
  });
  assert.equal(larkPreflightResponse.ok, true);
  assert.equal(larkPreflightResponse.tab_id, 8);
  assert.equal(larkPreflightResponse.body.result.platform, 'lark');
  assert.equal(backgroundRuntime.tabMessages.length, 2);
  assert.equal(backgroundRuntime.tabMessages[1].tabId, 8);
  assert.equal(backgroundRuntime.tabMessages[1].message.input.url, 'https://vc.feishu.cn/j/123456');
  assert.equal(backgroundRuntime.fetchCalls.length, 2);

  const batchPreflightResponse = await backgroundRuntime.send({
    type: MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates,
    request_id: 'preflight-candidates-001',
    query: {},
    options: { requireSpeakerTrack: true },
  });
  assert.equal(batchPreflightResponse.ok, true);
  assert.equal(batchPreflightResponse.type, MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates);
  assert.equal(batchPreflightResponse.candidate_count, 4);
  assert.equal(batchPreflightResponse.preflight_count, 4);
  assert.equal(batchPreflightResponse.accepted_count, 4);
  assert.deepEqual(batchPreflightResponse.rows.map((row) => row.preflight.platform), [
    'google_meet',
    'microsoft_teams',
    'zoom',
    'lark',
  ]);
  assert.equal(backgroundRuntime.tabMessages.length, 6);
  assert.deepEqual(backgroundRuntime.tabMessages.slice(2).map((item) => item.tabId), [7, 8, 9, 10]);
  assert.equal(backgroundRuntime.tabMessages[4].message.input.url, 'https://us06web.zoom.us/wc/987654321/start');
  assert.equal(backgroundRuntime.tabMessages[5].message.type, MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window);
  assert.equal(backgroundRuntime.fetchCalls.length, 2);
} finally {
  backgroundRuntime.restore();
}

const packageJson = buildMeetingAppExtensionPackageJson({
  packageName: 'demo-meeting-extension',
  sdkDependencyVersion: 'workspace:*',
});
assert.equal(packageJson.name, 'demo-meeting-extension');
assert.equal(packageJson.scripts.build, 'node build.mjs');
assert.equal(packageJson.dependencies['@ai-annotation/meeting-timeline-sdk'], 'workspace:*');
assert.equal(packageJson.devDependencies.esbuild, '^0.25.0');

const buildSource = buildMeetingAppExtensionBuildSource({
  outputScript: 'content-script.js',
  backgroundScript: 'background.js',
});
assert.match(buildSource, /from 'esbuild'/);
assert.match(buildSource, /src\/content-script\.entry\.mjs/);
assert.match(buildSource, /src\/background\.entry\.mjs/);
assert.match(buildSource, /src\/live-capture\.entry\.mjs/);
assert.match(buildSource, /live-capture\.js/);
assert.match(buildSource, /format: "iife"/);
assert.match(buildSource, /format: "esm"/);

const scaffold = buildMeetingAppExtensionScaffold({
  platforms: ['google_meet'],
  baseUrl: 'https://timeline.example.com',
  outputScript: 'content-script.js',
});
assert.equal(scaffold.type, 'meeting_app_extension_scaffold');
assert.equal(scaffold.validation.uses_all_urls, false);
assert.equal(scaffold.manifest.background.service_worker, 'background.js');
assert.equal(scaffold.manifest.background.type, 'module');
assert.equal(scaffold.manifest.permissions.includes('storage'), true);
assert.equal(scaffold.manifest.permissions.includes('tabs'), true);
assert.equal(scaffold.manifest.host_permissions.includes('https://timeline.example.com/*'), true);
assert.deepEqual(scaffold.manifest.content_scripts[0].js, ['content-script.js', 'live-capture.js']);
assert.equal(scaffold.bundle.background_input, 'src/background.entry.mjs');
assert.equal(scaffold.bundle.live_capture_input, 'src/live-capture.entry.mjs');
assert.equal(scaffold.bundle.live_capture_output, 'live-capture.js');
assert.equal(scaffold.files.find((file) => file.path === 'package.json').mime, 'application/json');
assert.match(scaffold.files.find((file) => file.path === 'package.json').content, /"esbuild"/);
assert.match(scaffold.files.find((file) => file.path === 'build.mjs').content, /content-script\.js/);
assert.match(scaffold.files.find((file) => file.path === 'build.mjs').content, /live-capture\.js/);
assert.equal(scaffold.files.find((file) => file.path === 'manifest.json').mime, 'application/json');
assert.match(scaffold.files.find((file) => file.path === 'src/content-script.entry.mjs').content, /platform-integration-runtime/);
assert.match(scaffold.files.find((file) => file.path === 'src/content-script.entry.mjs').content, /installMeetingPlatformIntegrationContentScriptBridge/);
assert.match(scaffold.files.find((file) => file.path === 'src/live-capture.entry.mjs').content, /__meetingTimelineLiveCapture/);
assert.match(scaffold.files.find((file) => file.path === 'src/live-capture.entry.mjs').content, /evidencePackage/);
assert.match(scaffold.files.find((file) => file.path === 'src/live-capture.entry.mjs').content, /diagnose/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /runtimeApi\(\)\?\.onMessage/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /extension_status/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /createMeetingPlatformRuntimeEventClient/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /observePlatformCandidates/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /meeting_timeline\.observe_candidates/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /meeting_timeline\.preflight_current_window/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /meeting_timeline\.preflight_candidates/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /npm run build/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /observe_candidates/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /preflight_current_window/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /preflight_candidates/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /diagnose\(\)/);

const scaffoldReport = buildMeetingAppExtensionScaffoldAcceptanceReport(scaffold);
assert.equal(scaffoldReport.accepted, true);
assert.deepEqual(scaffoldReport.platforms, ['google_meet']);
assert.equal(scaffoldReport.accepted_platform_count, 1);
assert.equal(scaffoldReport.manifest.uses_all_urls, false);
assert.equal(scaffoldReport.manifest.permissions.includes('storage'), true);
assert.equal(scaffoldReport.manifest.permissions.includes('tabs'), true);
assert.deepEqual(scaffoldReport.issues, []);
assert.equal(assertMeetingAppExtensionScaffold(scaffold).accepted, true);

const unsafeScaffold = buildMeetingAppExtensionScaffold({ platforms: ['google_meet'] });
const unsafeManifestFile = unsafeScaffold.files.find((file) => file.path === 'manifest.json');
const unsafeManifest = JSON.parse(unsafeManifestFile.content);
unsafeManifest.host_permissions = ['<all_urls>'];
unsafeManifestFile.content = `${JSON.stringify(unsafeManifest, null, 2)}\n`;
const unsafeReport = buildMeetingAppExtensionScaffoldAcceptanceReport(unsafeScaffold);
assert.equal(unsafeReport.accepted, false);
assert.equal(unsafeReport.issues.some((item) => item.code === 'overbroad_host_permission'), true);
assert.throws(
  () => assertMeetingAppExtensionScaffold(unsafeScaffold),
  /Meeting app extension scaffold acceptance failed/,
);

assert.throws(
  () => meetingAppExtensionProfile('unknown-meeting'),
  /Unsupported meeting app extension platform/,
);

console.log('ok meeting app extension');
