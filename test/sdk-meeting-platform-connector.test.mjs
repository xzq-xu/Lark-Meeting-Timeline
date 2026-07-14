import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_BROWSER_RUNTIME_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_CONTENT_SCRIPT_BRIDGE_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA,
  MEETING_PLATFORM_CONNECTOR_SCHEMA,
  assertMeetingPlatformConnector,
  buildDefaultMeetingPlatformConnectorHub,
  buildDefaultMeetingPlatformConnectorMatrix,
  buildMeetingPlatformConnector,
  buildMeetingPlatformConnectorAcceptanceReport,
  buildMeetingPlatformConnectorHub,
  buildMeetingPlatformConnectorMatrix,
  createMeetingPlatformConnectorBrowserRuntime,
  createMeetingPlatformConnectorContentScriptBridge,
  createMeetingPlatformConnectorHub,
  createMeetingPlatformConnectorRuntime,
  installMeetingPlatformConnectorContentScriptBridge,
  resolveMeetingPlatformConnectorInput,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-connector.mjs';
import {
  createMeetingPlatformConnectorHub as createMeetingPlatformConnectorHubFromRoot,
  createMeetingPlatformConnectorRuntime as createMeetingPlatformConnectorRuntimeFromRoot,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

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
    emit(message, sender = { tab: { id: 7 } }) {
      if (listeners.length === 0) return Promise.resolve({ consumed: false });
      return new Promise((resolve) => {
        const consumed = listeners[0](message, sender, (response) => resolve({ consumed, response }));
        if (consumed !== true) resolve({ consumed, response: undefined });
      });
    },
  };
}

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

function selectorAttrMatches(item, selector) {
  if (selector === 'button') return item.tagName === 'BUTTON';
  if (selector.startsWith('.')) {
    return String(item.attributes.class ?? '').split(/\s+/).includes(selector.slice(1));
  }
  const attrParts = [...String(selector).matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (!attrParts.length) return false;
  return attrParts.every((match) => {
    const [, attrName, operator, doubleQuoted, singleQuoted, bare] = match;
    const actual = item.attributes[attrName];
    if (operator == null) return actual != null;
    if (actual == null) return false;
    const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
    if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
    return String(actual) === String(expected);
  });
}

function queryNodes(nodes, selector) {
  const text = String(selector);
  if (text === '*') return nodes;
  const generic = nodes.filter((item) => selectorAttrMatches(item, text));
  if (generic.length) return generic;
  if (text.includes('speaking')) {
    return nodes.filter((item) => /speaking|active speaker|正在发言|正在讲话|正在说话/i.test(item.attributes['aria-label'] ?? ''));
  }
  if (text.includes('aria-live')) return nodes.filter((item) => item.attributes['aria-live']);
  if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
  return [];
}

function fakeBrowserWindow(url = 'https://meet.google.com/abc-defg-hij', nodes = []) {
  const document = {
    nodeType: 9,
    title: 'Connector browser runtime',
    hidden: false,
    location: { href: url },
    body: { nodeType: 1 },
    documentElement: { nodeType: 1 },
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
  return {
    document,
    location: { href: url, origin: new URL(url).origin },
    navigator: { userAgent: 'Chrome fixture' },
    addEventListener() {},
    removeEventListener() {},
  };
}

const google = buildMeetingPlatformConnector('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_CONNECTOR_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.display_name, 'Google Meet');
assert.equal(google.event_adapter.normalize_available, true);
assert.equal(google.provider.required_for_realtime, false);
assert.equal(google.provider.transport, 'Google Workspace Events API -> Google Cloud Pub/Sub push');
assert.equal(google.provider.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.browser_observer.enabled, true);
assert.equal(google.browser_observer.matches.includes('https://meet.google.com/*'), true);
assert.equal(google.browser_observer.candidate_observation_ready, true);
assert.equal(google.browser_observer.candidate_message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.browser_observer.required_permission, 'tabs');
assert.equal(google.runtime_events.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.runtime_events.supported_actions.includes('insert_annotation'), true);
assert.equal(google.runtime_events.supported_actions.includes('observe_platform_candidates'), true);
assert.equal(google.runtime_events.supported_actions.includes('adapter_blueprint'), true);
assert.equal(google.runtime_events.supported_actions.includes('adapter_blueprints'), true);
assert.equal(google.adapter_blueprint.ready, true);
assert.equal(google.adapter_blueprint.endpoint, '/api/meeting-platform/adapter-blueprints');
assert.equal(google.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(google.adapter_blueprint.realtime_axis_timestamp_field, 'captured_at_ms');
assert.equal(google.timeline_ingest.insert_endpoint, `${baseUrl}/api/annotations`);
assert.equal(google.timeline_ingest.timestamp_field, 'captured_at_ms');
assert.equal(google.realtime_policy.provider_events_block_realtime, false);
assert.equal(google.realtime_policy.transcript_blocks_realtime, false);
assert.equal(google.realtime_policy.primary_axis, 'local_observer_axis');
assert.equal(google.sdk.imports.connector, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector');
assert.equal(google.sdk.factories.create_runtime, 'createMeetingPlatformConnectorRuntime');
assert.equal(google.readiness.realtime_annotation_ready, true);

const googleAcceptance = buildMeetingPlatformConnectorAcceptanceReport(google);
assert.equal(googleAcceptance.schema, MEETING_PLATFORM_CONNECTOR_ACCEPTANCE_SCHEMA);
assert.equal(googleAcceptance.accepted, true);
assert.equal(assertMeetingPlatformConnector(google).accepted, true);

const matrix = buildMeetingPlatformConnectorMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(matrix.schema, MEETING_PLATFORM_CONNECTOR_MATRIX_SCHEMA);
assert.equal(matrix.platform_count, 5);
assert.equal(matrix.accepted_count, 5);
assert.equal(matrix.realtime_ready_count, 5);
assert.equal(matrix.candidate_observer_count, 5);
assert.equal(matrix.adapter_blueprint_ready_count, 5);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_ready, true);
assert.equal(matrix.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').browser_observer_enabled, true);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').runtime_action_count, 18);
assert.equal(matrix.registry_manifest.platform_count, 5);

const defaultMatrix = buildDefaultMeetingPlatformConnectorMatrix({ baseUrl });
assert.equal(defaultMatrix.platforms.includes('local_detector'), false);
assert.equal(defaultMatrix.platforms.includes('google_meet'), true);

const hub = buildMeetingPlatformConnectorHub({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(hub.schema, MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA);
assert.equal(hub.accepted, true);
assert.equal(hub.platform_count, 5);
assert.equal(hub.default_platform, 'google_meet');
assert.equal(hub.routing.url_detection_ready, true);
assert.equal(hub.routing.detected_platforms.includes('microsoft_teams'), true);
assert.equal(hub.matrix.accepted_count, 5);
assert.equal(hub.adapter_blueprint_ready_count, 5);
assert.equal(hub.readiness.adapter_blueprint_ready, true);

const defaultHub = buildDefaultMeetingPlatformConnectorHub({ baseUrl });
assert.equal(defaultHub.platforms.includes('local_detector'), false);
assert.equal(defaultHub.accepted, true);

const googleResolution = resolveMeetingPlatformConnectorInput('https://meet.google.com/abc-defg-hij', {
  platforms: ['google-meet', 'teams'],
});
assert.equal(googleResolution.schema, MEETING_PLATFORM_CONNECTOR_RESOLUTION_SCHEMA);
assert.equal(googleResolution.detected, true);
assert.equal(googleResolution.supported, true);
assert.equal(googleResolution.platform, 'google_meet');
assert.equal(googleResolution.reason, 'meeting_url');

const teamsResolution = resolveMeetingPlatformConnectorInput({
  tabs: [
    { active: false, url: 'https://example.com/' },
    { active: true, url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample', title: 'Teams meeting' },
  ],
}, {
  platforms: ['google-meet', 'teams'],
});
assert.equal(teamsResolution.platform, 'microsoft_teams');
assert.equal(teamsResolution.browser.active, true);

const broken = {
  ...google,
  runtime_events: {
    ...google.runtime_events,
    supported_actions: google.runtime_events.supported_actions.filter((action) => action !== 'speaker_track'),
  },
};
const brokenAcceptance = buildMeetingPlatformConnectorAcceptanceReport(broken);
assert.equal(brokenAcceptance.accepted, false);
assert.equal(brokenAcceptance.issues.some((item) => item.code === 'missing_runtime_action' && item.action === 'speaker_track'), true);
assert.throws(
  () => assertMeetingPlatformConnector(broken),
  /Meeting platform connector failed acceptance/,
);

const calls = [];
const fetchImpl = async (url, init) => {
  const body = JSON.parse(init.body);
  calls.push({ url, init, body });
  return new Response(JSON.stringify({ ok: true, accepted: body }), {
    status: 200,
    headers: { 'content-type': 'application/json' },
  });
};

const runtime = createMeetingPlatformConnectorRuntime(google, {
  fetch: fetchImpl,
  now: () => 1_782_614_400_000,
});
assert.equal(runtime.schema, MEETING_PLATFORM_CONNECTOR_RUNTIME_SCHEMA);
assert.equal(runtime.platform, 'google_meet');
assert.equal(runtime.endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(runtime.supports('insert_mark'), true);
assert.equal(runtime.supports('unknown-action'), false);

const builtEvent = runtime.buildEvent({
  action: 'insert_annotation',
  annotation: {
    id: 'mark-001',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  },
});
assert.equal(builtEvent.action, 'insert_annotation');
assert.equal(builtEvent.platform, 'google_meet');
assert.equal(builtEvent.sent_at_ms, 1_782_614_400_000);

const providerSignals = runtime.normalizeProviderEvent({
  id: 'google-event-001',
  type: 'google.workspace.meet.conference.v2.started',
  time: '2026-06-30T06:45:28.000Z',
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
    meetingUri: 'https://meet.google.com/abc-defg-hij',
    title: 'Weekly review',
  },
});
assert.equal(providerSignals.length, 1);
assert.equal(providerSignals[0].type, 'meeting_started');
assert.equal(providerSignals[0].meeting.platform, 'google_meet');
assert.equal(providerSignals[0].meeting.meeting_id, 'google-record-001');

const insertResult = await runtime.insertAnnotation({
  id: 'mark-002',
  label: 'follow up',
  captured_at_ms: 1_782_614_402_000,
});
assert.equal(insertResult.ok, true);
assert.equal(calls.at(-1).url, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(calls.at(-1).body.action, 'insert_annotation');
assert.equal(calls.at(-1).body.platform, 'google_meet');
assert.equal(calls.at(-1).body.annotation.label, 'follow up');

await runtime.observeMeetingApp({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Weekly review',
  observed_at_ms: 1_782_614_403_000,
  in_meeting: true,
});
assert.equal(calls.at(-1).body.action, 'observe_meeting_app');
assert.equal(calls.at(-1).body.platform, 'google_meet');

await runtime.observePlatformCandidates({
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Weekly review', active: true }],
});
assert.equal(calls.at(-1).body.action, 'observe_platform_candidates');
assert.equal(calls.at(-1).body.platform, undefined);
assert.equal(calls.at(-1).body.tabs[0].active, true);

await runtime.ingestProvider({
  type: 'google.workspace.meet.conference.v2.ended',
  data: {
    conferenceRecord: { name: 'conferenceRecords/google-record-001' },
  },
});
assert.equal(calls.at(-1).body.action, 'provider_event');
assert.equal(calls.at(-1).body.platform, 'google_meet');

await runtime.speakerTrack({
  signals: [{
    type: 'speaker_started',
    occurred_at_ms: 1_782_614_410_000,
    speaker: { display_name: 'Alex' },
  }],
});
assert.equal(calls.at(-1).body.action, 'speaker_track');
assert.equal(calls.at(-1).body.platform, 'google_meet');

await runtime.adapterBlueprint();
assert.equal(calls.at(-1).body.action, 'adapter_blueprint');
assert.equal(calls.at(-1).body.platform, 'google_meet');

await runtime.adapterBlueprints();
assert.equal(calls.at(-1).body.action, 'adapter_blueprints');
assert.deepEqual(calls.at(-1).body.platforms, ['google_meet']);

const hubRuntime = createMeetingPlatformConnectorHub({
  baseUrl,
  fetch: fetchImpl,
  now: () => 1_782_614_400_000,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(hubRuntime.schema, MEETING_PLATFORM_CONNECTOR_HUB_SCHEMA);
assert.equal(hubRuntime.platforms.length, 5);
assert.equal(hubRuntime.resolvePlatform({ url: 'https://zoom.us/j/987654321' }).platform, 'zoom');
assert.equal(hubRuntime.connectorFor({ url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample' }).platform, 'microsoft_teams');
assert.equal(hubRuntime.runtimeFor({ url: 'https://meet.google.com/abc-defg-hij' }).platform, 'google_meet');
assert.equal(hubRuntime.supports({ url: 'https://meet.google.com/abc-defg-hij' }, 'insert_annotation'), true);

await hubRuntime.adapterBlueprint({ url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample' });
assert.equal(calls.at(-1).body.action, 'adapter_blueprint');
assert.equal(calls.at(-1).body.platform, 'microsoft_teams');

await hubRuntime.adapterBlueprints();
assert.equal(calls.at(-1).body.action, 'adapter_blueprints');
assert.deepEqual(calls.at(-1).body.platforms, ['google_meet']);

await hubRuntime.insertAnnotation({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Weekly review',
}, {
  id: 'hub-mark-001',
  label: 'why?',
  captured_at_ms: 1_782_614_420_000,
});
assert.equal(calls.at(-1).body.action, 'insert_annotation');
assert.equal(calls.at(-1).body.platform, 'google_meet');
assert.equal(calls.at(-1).body.annotation.id, 'hub-mark-001');

await hubRuntime.observeMeetingApp({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
  title: 'Teams meeting',
  in_meeting: true,
});
assert.equal(calls.at(-1).body.action, 'observe_meeting_app');
assert.equal(calls.at(-1).body.platform, 'microsoft_teams');

await hubRuntime.observePlatformCandidates({
  tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Weekly review' }],
});
assert.equal(calls.at(-1).body.action, 'observe_platform_candidates');
assert.equal(calls.at(-1).body.platform, undefined);

const hubSignals = hubRuntime.normalizeProviderEvent('google-meet', {
  id: 'hub-google-event-001',
  type: 'google.workspace.meet.conference.v2.started',
  data: { conferenceRecord: { name: 'conferenceRecords/hub-google-record-001' } },
});
assert.equal(hubSignals[0].type, 'meeting_started');
assert.equal(hubSignals[0].meeting.platform, 'google_meet');

const rootHubRuntime = createMeetingPlatformConnectorHubFromRoot({
  baseUrl,
  fetch: fetchImpl,
});
assert.equal(rootHubRuntime.connectorFor('google-meet').platform, 'google_meet');

const browserWindow = fakeBrowserWindow('https://meet.google.com/abc-defg-hij', [
  node('button', { 'aria-label': 'Turn off microphone' }),
  node('button', { 'aria-label': 'Leave call' }),
  node('div', {
    'data-participant-id': 'ada',
    'aria-label': 'Ada Lovelace is speaking',
    'data-audio-level': '0.84',
  }),
  node('div', { role: 'status', 'aria-live': 'polite' }, 'You are presenting'),
]);
const browserRuntime = createMeetingPlatformConnectorBrowserRuntime({
  baseUrl,
  fetch: fetchImpl,
  window: browserWindow,
  now: () => 1_782_614_400_000,
});
assert.equal(browserRuntime.schema, MEETING_PLATFORM_CONNECTOR_BROWSER_RUNTIME_SCHEMA);
assert.equal(browserRuntime.resolvePlatform().platform, 'google_meet');
await browserRuntime.insertAnnotation({
  id: 'browser-mark-001',
  label: 'browser why?',
  captured_at_ms: 1_782_614_430_000,
});
assert.equal(calls.at(-1).body.action, 'insert_annotation');
assert.equal(calls.at(-1).body.platform, 'google_meet');
assert.equal(calls.at(-1).body.annotation.id, 'browser-mark-001');
await browserRuntime.observeMeetingApp({ in_meeting: true });
assert.equal(calls.at(-1).body.action, 'observe_meeting_app');
assert.equal(calls.at(-1).body.platform, 'google_meet');
await browserRuntime.observePlatformCandidates({
  tabs: [{ active: true, url: 'https://zoom.us/j/987654321', title: 'Zoom meeting' }],
});
assert.equal(calls.at(-1).body.action, 'observe_platform_candidates');
assert.equal(calls.at(-1).body.tabs[0].url, 'https://zoom.us/j/987654321');
const browserPreflight = browserRuntime.currentWindowPreflight({}, {
  requireSpeakerTrack: true,
  observedAtMs: 1_782_614_433_000,
});
assert.equal(browserPreflight.accepted, true);
assert.equal(browserPreflight.platform, 'google_meet');
assert.equal(browserPreflight.readiness.realtime_annotation_ready, true);
assert.equal(browserRuntime.preflightCurrentWindow({}, { requireSpeakerTrack: true }).accepted, true);

const connectorBridge = createMeetingPlatformConnectorContentScriptBridge({
  baseUrl,
  fetch: fetchImpl,
  window: browserWindow,
  now: () => 1_782_614_400_000,
});
assert.equal(connectorBridge.schema, MEETING_PLATFORM_CONNECTOR_CONTENT_SCRIPT_BRIDGE_SCHEMA);
assert.equal(connectorBridge.resolvePlatform().platform, 'google_meet');
const bridgeResponse = await connectorBridge.dispatchMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'bridge-mark-001',
      label: 'bridge why?',
      captured_at_ms: 1_782_614_431_000,
    },
  },
});
assert.equal(bridgeResponse.handled, true);
assert.equal(bridgeResponse.action, 'insertMark');
assert.equal(calls.at(-1).body.platform, 'google_meet');
assert.equal(calls.at(-1).body.annotation.id, 'bridge-mark-001');
const bridgePreflight = connectorBridge.currentWindowPreflight({}, {
  requireSpeakerTrack: true,
  observedAtMs: 1_782_614_434_000,
});
assert.equal(bridgePreflight.accepted, true);
assert.equal(bridgePreflight.platform, 'google_meet');
const bridgePreflightResponse = await connectorBridge.dispatchMessage({
  type: 'meeting_timeline.preflight_current_window',
  payload: {
    options: {
      requireSpeakerTrack: true,
      observedAtMs: 1_782_614_435_000,
    },
  },
});
assert.equal(bridgePreflightResponse.handled, true);
assert.equal(bridgePreflightResponse.action, 'preflightCurrentWindow');
assert.equal(bridgePreflightResponse.result.accepted, true);
assert.equal(bridgePreflightResponse.result.platform, 'google_meet');

const platformPreflightCases = [
  {
    platform: 'microsoft_teams',
    url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    nodes: [
      node('button', { 'aria-label': 'Leave' }),
      node('button', { 'aria-label': 'Share content' }),
      node('div', {
        'data-tid': 'participant-ada',
        'data-user-id': 'ada',
        'data-display-name': 'Ada Lovelace',
        'aria-label': 'Ada Lovelace speaking',
      }),
    ],
  },
  {
    platform: 'zoom',
    url: 'https://us06web.zoom.us/wc/987654321/start',
    nodes: [
      node('button', { 'aria-label': 'Leave Meeting' }),
      node('button', { 'aria-label': 'Participants' }),
      node('div', {
        'data-user-id': 'mira',
        'aria-label': 'Mira Patel is speaking',
      }),
    ],
  },
];
for (const item of platformPreflightCases) {
  const platformBridge = createMeetingPlatformConnectorContentScriptBridge({
    baseUrl,
    fetch: fetchImpl,
    window: fakeBrowserWindow(item.url, item.nodes),
    now: () => 1_782_614_400_000,
  });
  const response = await platformBridge.dispatchMessage({
    type: 'meeting_timeline.preflight_current_window',
    payload: { options: { requireSpeakerTrack: true } },
  });
  assert.equal(response.handled, true);
  assert.equal(response.result.platform, item.platform);
  assert.equal(response.result.accepted, true);
}

const extensionRuntime = fakeExtensionRuntime();
const installedBridge = installMeetingPlatformConnectorContentScriptBridge({
  baseUrl,
  fetch: fetchImpl,
  window: browserWindow,
  extensionRuntime,
  startOptions: { startRuntime: false },
});
assert.equal(extensionRuntime.listenerCount(), 1);
const extensionMessage = await extensionRuntime.emit({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      id: 'extension-mark-001',
      label: 'extension why?',
      captured_at_ms: 1_782_614_432_000,
    },
  },
});
assert.equal(extensionMessage.consumed, true);
assert.equal(extensionMessage.response.handled, true);
assert.equal(calls.at(-1).body.platform, 'google_meet');
assert.equal(calls.at(-1).body.annotation.id, 'extension-mark-001');
installedBridge.dispose();
assert.equal(extensionRuntime.listenerCount(), 0);

const rootRuntime = createMeetingPlatformConnectorRuntimeFromRoot('teams', {
  baseUrl,
  fetch: fetchImpl,
  now: () => 1_782_614_400_000,
});
assert.equal(rootRuntime.platform, 'microsoft_teams');
assert.equal(rootRuntime.supports('participant_track'), true);

const brokenRuntime = createMeetingPlatformConnectorRuntime(broken, {
  assertConnector: false,
  fetch: fetchImpl,
});
assert.throws(
  () => brokenRuntime.assertSupported('speaker_track'),
  /Runtime action is not present in meeting platform connector/,
);

console.log('ok meeting platform connector facade');
