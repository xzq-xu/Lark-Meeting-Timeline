import assert from 'node:assert/strict';

import { buildMeetingAppFixtureSnapshot } from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA,
  MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA,
  assertMeetingPlatformIntegrationRuntimeManifest,
  buildMeetingPlatformIntegrationRuntimeManifest,
  createMeetingPlatformIntegrationBrowserRuntime,
  createMeetingPlatformIntegrationContentScriptBridge,
  createMeetingPlatformIntegrationRuntime,
  detectMeetingPlatformForBrowser,
  installMeetingPlatformIntegrationContentScriptBridge,
  resolveMeetingPlatformForInput,
} from '../packages/meeting-timeline-sdk/adapters/platform-integration-runtime.mjs';
import { buildPlatformFixtureEvent } from '../packages/meeting-timeline-sdk/adapters/platform-fixtures.mjs';

const baseUrl = 'https://timeline.example.com';
const browserStartMs = 1_782_614_400_000;

function node(tagName, attrs = {}, text = '') {
  return {
    tagName: tagName.toUpperCase(),
    attributes: attrs,
    dataset: Object.fromEntries(Object.entries(attrs)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [
        key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()),
        value,
      ])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attrs[name] ?? null;
    },
  };
}

function fakeGoogleMeetDocument() {
  const nodes = [
    node('button', { 'aria-label': 'Turn off microphone' }),
    node('button', { 'aria-label': 'Leave call' }),
    node('div', {
      'data-participant-id': 'ada',
      'data-display-name': 'Ada Lovelace',
      'aria-label': 'Ada Lovelace is speaking',
      'data-audio-level': '0.82',
    }),
  ];
  return {
    nodeType: 9,
    title: 'SDK runtime - Google Meet',
    hidden: false,
    location: { href: 'https://meet.google.com/abc-defg-hij' },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('data-participant-id')) return nodes.filter((item) => item.attributes['data-participant-id']);
      if (text.includes('data-display-name')) return nodes.filter((item) => item.attributes['data-display-name']);
      if (text.includes('speaking')) return nodes.filter((item) => /speaking/i.test(item.attributes['aria-label'] ?? ''));
      return [];
    },
  };
}

function fakeWindow(document) {
  const listeners = new Map();
  return {
    document,
    location: document.location,
    navigator: { userAgent: 'Chrome fixture' },
    addEventListener(eventName, handler) {
      const rows = listeners.get(eventName) ?? [];
      rows.push(handler);
      listeners.set(eventName, rows);
    },
    removeEventListener(eventName, handler) {
      listeners.set(eventName, (listeners.get(eventName) ?? []).filter((item) => item !== handler));
    },
    listenerCount(eventName) {
      return (listeners.get(eventName) ?? []).length;
    },
  };
}

function fakeExtensionRuntime() {
  const listeners = [];
  return {
    onMessage: {
      addListener(listener) {
        listeners.push(listener);
      },
      removeListener(listener) {
        const index = listeners.indexOf(listener);
        if (index >= 0) listeners.splice(index, 1);
      },
    },
    listenerCount() {
      return listeners.length;
    },
    emit(message, sender = { tab: { id: 1 } }) {
      if (listeners.length === 0) return Promise.resolve({ consumed: false });
      return new Promise((resolve) => {
        let settled = false;
        const consumed = listeners[0](message, sender, (response) => {
          settled = true;
          resolve({ consumed, response });
        });
        if (consumed !== true && !settled) {
          resolve({ consumed, response: undefined });
        }
      });
    },
  };
}

const browserDetection = detectMeetingPlatformForBrowser({
  url: 'https://teams.microsoft.com/l/meetup-join/abc',
  title: 'Teams meeting',
});
assert.equal(browserDetection.detected, true);
assert.equal(browserDetection.platform, 'microsoft_teams');
assert.equal(browserDetection.reason, 'url');

const resolvedGoogle = resolveMeetingPlatformForInput({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(resolvedGoogle.type, 'meeting_platform_resolution');
assert.equal(resolvedGoogle.detected, true);
assert.equal(resolvedGoogle.supported, true);
assert.equal(resolvedGoogle.platform, 'google_meet');
assert.equal(resolvedGoogle.strategy.primary_axis_source, 'local_observer');
assert.equal(resolvedGoogle.registry.runtime_event_endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(resolvedGoogle.runtime.provider_required_for_realtime, false);

const resolvedTeamsUnsupported = resolveMeetingPlatformForInput({
  url: 'https://teams.microsoft.com/l/meetup-join/abc',
}, {
  baseUrl,
  platforms: ['google-meet'],
});
assert.equal(resolvedTeamsUnsupported.platform, 'microsoft_teams');
assert.equal(resolvedTeamsUnsupported.supported, false);
assert.equal(resolvedTeamsUnsupported.strategy.primary_axis_source, 'local_observer');
assert.equal(resolvedTeamsUnsupported.runtime.runtime_ready, true);
assert.equal(resolvedTeamsUnsupported.next_actions.includes('enable_detected_platform_in_runtime_platforms'), true);

const manifest = buildMeetingPlatformIntegrationRuntimeManifest({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(manifest.schema, MEETING_PLATFORM_INTEGRATION_RUNTIME_MANIFEST_SCHEMA);
assert.equal(manifest.platform_count, 2);
assert.deepEqual(manifest.platforms, ['google_meet', 'zoom']);
assert.equal(manifest.host_integration_ready, true);
assert.equal(manifest.blocking_count, 0);
assert.equal(manifest.warning_count, 1);
assert.equal(manifest.registry_acceptance.accepted, true);
assert.equal(manifest.runtime_bundle_matrix.runtime_ready_count, 2);
assert.equal(manifest.runtime_bundle_matrix.provider_required_for_realtime_count, 0);
assert.equal(manifest.adaptation_strategy_matrix.strategy_count, 2);
assert.equal(manifest.adaptation_package_matrix.sdk_wiring_ready_count, 2);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').browser_match_count, 1);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').primary_axis_source, 'local_observer');
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').provider_blocks_realtime, false);
assert.equal(manifest.rows.find((row) => row.platform === 'zoom').sdk_wiring_ready, true);
assert.equal(assertMeetingPlatformIntegrationRuntimeManifest(manifest).host_integration_ready, true);

const calls = [];
const client = {
  async startMeeting(input) {
    calls.push({ method: 'startMeeting', input });
    return { ok: true, method: 'startMeeting', input };
  },
  async endMeeting(input) {
    calls.push({ method: 'endMeeting', input });
    return { ok: true, method: 'endMeeting', input };
  },
  async insertMark(input) {
    calls.push({ method: 'insertMark', input });
    return { ok: true, method: 'insertMark', input };
  },
  async insertMarks(input) {
    calls.push({ method: 'insertMarks', input });
    return { ok: true, method: 'insertMarks', input };
  },
  async importTranscript(input) {
    calls.push({ method: 'importTranscript', input });
    return { ok: true, method: 'importTranscript', input };
  },
};

const runtime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  verify: false,
});
assert.equal(runtime.schema, MEETING_PLATFORM_INTEGRATION_RUNTIME_SCHEMA);
assert.deepEqual(runtime.platforms, ['google_meet', 'zoom']);
assert.equal(runtime.manifest().host_integration_ready, true);
assert.equal(runtime.summary().host_integration_ready, true);
assert.equal(runtime.registry().acceptance.accepted, true);
assert.equal(runtime.runtimeBundle('google-meet').browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(runtime.runtimeBundles().platform_count, 2);
assert.equal(runtime.adaptationStrategyMatrix().strategy_count, 2);
const runtimeResolution = runtime.resolvePlatform({ url: 'https://meet.google.com/abc-defg-hij' });
assert.equal(runtimeResolution.supported, true);
assert.equal(runtimeResolution.strategy.provider_blocks_realtime, false);
assert.equal(runtime.adaptationPackages().sdk_wiring_ready_count, 2);
assert.equal(runtime.adapter('google-meet').platform, 'google_meet');

const googleSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_782_614_400_000,
});
const observeResult = await runtime.observeMeetingApp('google-meet', googleSnapshot);
assert.equal(observeResult.action, 'observe_meeting_app');
assert.equal(observeResult.live_evidence.meeting_app_record_count, 1);
assert.equal(calls.some((call) => call.method === 'startMeeting'), true);

const insertResult = await runtime.insertAnnotation('google-meet', {
  annotation: {
    id: 'note-runtime-1',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
});
assert.equal(insertResult.action, 'insert_annotation');
assert.equal(insertResult.result.actions.includes('insert_mark'), true);
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'note-runtime-1'), true);

const providerResult = await runtime.handleEvent({
  ...buildPlatformFixtureEvent('google-meet', 'meeting_end'),
  action: 'provider_event',
  platform: 'google-meet',
});
assert.equal(providerResult.action, 'ingest_provider');
assert.equal(providerResult.live_evidence.provider_record_count, 1);
assert.equal(providerResult.diagnostic.raw_signal_count, 1);

const routedInsert = await runtime.handleEvent({
  action: 'insert_annotation',
  platform: 'zoom',
  annotation: {
    id: 'zoom-note-1',
    label: 'follow up',
    captured_at_ms: 1_782_614_405_000,
  },
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
});
assert.equal(routedInsert.result.status, 'ready_to_insert');
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'zoom-note-1'), true);

const speakerTrack = runtime.speakerTrack('google-meet', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      speaker_name: 'Ada',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_402_000,
      speaker_name: 'Ada',
    },
  ],
});
assert.equal(speakerTrack.mark_count, 1);
assert.equal(speakerTrack.marks[0].intent, 'speaker_track');

const participantTrack = await runtime.handleEvent({
  action: 'participant_track',
  platform: 'zoom',
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_403_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
});
assert.equal(participantTrack.mark_count, 1);

const view = runtime.timelineView('zoom', {
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'zoom-note-1',
      label: 'follow up',
      captured_at_ms: 1_782_614_405_000,
    },
  ],
  speakerTrack,
});
assert.equal(view.diagnostics.marker_count, 2);
assert.equal(view.markers.some((marker) => marker.rail === 'speaker'), true);
assert.equal((await runtime.handleEvent({ action: 'runtime_bundles' })).platform_count, 2);
assert.equal((await runtime.handleEvent({ action: 'adaptation_strategy_matrix' })).strategy_count, 2);
assert.equal((await runtime.handleEvent({
  action: 'resolve_platform',
  url: 'https://meet.google.com/abc-defg-hij',
})).platform, 'google_meet');
assert.equal((await runtime.handleEvent({ action: 'manifest' })).host_integration_ready, true);
await assert.rejects(
  () => runtime.handleEvent({ action: 'insert_annotation' }),
  /platform is required/,
);

const browserDocument = fakeGoogleMeetDocument();
const browserWindow = fakeWindow(browserDocument);
const browserRuntime = createMeetingPlatformIntegrationBrowserRuntime(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  window: browserWindow,
  now: () => browserStartMs,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
});
assert.equal(browserRuntime.detect().platform, 'google_meet');
assert.equal(browserRuntime.resolvePlatform().platform, 'google_meet');
assert.equal(browserRuntime.platformFor(), 'google_meet');
const browserSample = await browserRuntime.sample();
assert.equal(browserSample.emitted, true);
assert.equal(browserSample.result.browser_detection.platform, 'google_meet');
assert.equal(calls.some((call) => call.method === 'startMeeting' && call.input.meeting_id === 'abc-defg-hij'), true);
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.payload?.speaker_name === 'Ada Lovelace'), true);
const browserInserted = await browserRuntime.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'browser-integration-note-1',
      label: 'why?',
      captured_at_ms: browserStartMs + 3_000,
    },
  },
});
assert.equal(browserInserted.handled, true);
assert.equal(browserInserted.result.action, 'insert_annotation');
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'browser-integration-note-1'), true);
assert.equal(browserRuntime.getState().browser_detection.platform, 'google_meet');

const extensionRuntime = fakeExtensionRuntime();
const contentBridge = createMeetingPlatformIntegrationContentScriptBridge(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
  window: browserWindow,
  now: () => browserStartMs,
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 0 },
  extensionRuntime,
  extensionMessaging: true,
  windowMessaging: false,
});
const contentStart = contentBridge.start({ startRuntime: false });
assert.equal(contentStart.installs[0].installed, true);
assert.equal(extensionRuntime.listenerCount(), 1);
assert.equal(contentBridge.detect().platform, 'google_meet');
await contentBridge.runtime.sample();
const extensionResponse = await extensionRuntime.emit({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'content-bridge-note-1',
      label: 'what?',
      captured_at_ms: browserStartMs + 4_000,
    },
  },
});
assert.equal(extensionResponse.consumed, true);
assert.equal(extensionResponse.response.handled, true);
assert.equal(extensionResponse.response.result.action, 'insert_annotation');
assert.equal(calls.some((call) => call.method === 'insertMark' && call.input.id === 'content-bridge-note-1'), true);
assert.equal(contentBridge.getState().browser_detection.platform, 'google_meet');
contentBridge.dispose();
assert.equal(extensionRuntime.listenerCount(), 0);

const installedBridgeRuntime = fakeExtensionRuntime();
const installedBridge = installMeetingPlatformIntegrationContentScriptBridge(client, {
  baseUrl,
  platforms: ['google-meet'],
  window: browserWindow,
  extensionRuntime: installedBridgeRuntime,
  startOptions: { startRuntime: false },
});
assert.equal(installedBridgeRuntime.listenerCount(), 1);
installedBridge.dispose();
assert.equal(installedBridgeRuntime.listenerCount(), 0);

console.log('ok meeting platform integration runtime');
