import assert from 'node:assert/strict';

import {
  buildMeetingPlatformRuntimeHostConfig,
  buildMeetingPlatformRuntimeHostConfigMatrix,
  buildMeetingPlatformRuntimeHostHandoff,
  buildMeetingPlatformRuntimeHostHandoffMatrix,
  createMeetingPlatformRuntimeHost,
} from '../packages/meeting-timeline-sdk/adapters/meeting-platform-runtime-host.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const googleConfig = buildMeetingPlatformRuntimeHostConfig('google-meet', { baseUrl });
assert.equal(googleConfig.schema, 'meeting_platform_runtime_host_config');
assert.equal(googleConfig.platform, 'google_meet');
assert.equal(googleConfig.runtime_bundle.schema, 'meeting_platform_runtime_bundle');
assert.equal(googleConfig.runtime_factory, 'createMeetingAppBrowserRuntime');
assert.equal(googleConfig.observer_scheduler_config.schema, 'meeting_app_observer_scheduler_config');
assert.equal(googleConfig.driver.change_observer.enabled, true);
assert.equal(googleConfig.driver.change_observer.trigger, 'dom_mutation');
assert.equal(googleConfig.driver.keep_alive.enabled, true);
assert.equal(googleConfig.readiness.host_ready, true);
assert.equal(googleConfig.next_actions.includes('call_host.start_when_meeting_surface_is_open'), true);

const matrix = buildMeetingPlatformRuntimeHostConfigMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(matrix.schema, 'meeting_platform_runtime_host_config_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.host_ready_count, 3);
assert.equal(matrix.change_observer_count, 3);
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').runtime_factory, 'createMeetingAppBrowserRuntime');
assert.equal(matrix.runtime_bundle_matrix.runtime_host_ready_count, 3);
assert.equal(matrix.observer_scheduler_config_matrix.sdk_ready_count, 3);

const googleHandoff = buildMeetingPlatformRuntimeHostHandoff(googleConfig);
assert.equal(googleHandoff.schema, 'meeting_platform_runtime_host_handoff');
assert.equal(googleHandoff.platform, 'google_meet');
assert.equal(googleHandoff.package_entry, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host');
assert.equal(googleHandoff.acceptance.accepted, true);
assert.equal(googleHandoff.runtime.timestamp_field, 'captured_at_ms');
assert.equal(googleHandoff.runtime.provider_events_role, 'reconcile_and_backfill_only');
assert.equal(googleHandoff.runtime.transcript_role, 'post_meeting_backfill_only');
assert.equal(googleHandoff.host_hooks.some((item) => item.call === 'host.changed()'), true);
assert.equal(googleHandoff.example.includes("createMeetingPlatformRuntimeHost(timeline, 'google_meet'"), true);
const handoffMatrix = buildMeetingPlatformRuntimeHostHandoffMatrix({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
});
assert.equal(handoffMatrix.schema, 'meeting_platform_runtime_host_handoff_matrix');
assert.equal(handoffMatrix.platform_count, 3);
assert.equal(handoffMatrix.accepted_count, 3);
assert.equal(handoffMatrix.provider_blocking_count, 0);
assert.equal(handoffMatrix.transcript_blocking_count, 0);
assert.equal(handoffMatrix.rows.find((row) => row.platform === 'zoom').host_hook_count, 5);

const calls = [];
const runtime = {
  inputProvider() {
    return {
      platform: 'google_meet',
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
    };
  },
  async sample(options = {}) {
    calls.push({ kind: 'sample', options });
    return {
      ok: true,
      trigger: options.trigger,
      captured_at_ms: options.captured_at_ms,
    };
  },
  async sampleTracks(options = {}) {
    calls.push({ kind: 'tracks', options });
    return { ok: true, trigger: options.trigger };
  },
  stop() {
    calls.push({ kind: 'stop' });
    return { stopped: true };
  },
  getState() {
    return { sample_count: calls.filter((row) => row.kind === 'sample').length };
  },
};

const pendingTimeouts = [];
const pendingIntervals = [];
function fakeSetTimeout(fn, ms) {
  const handle = { fn, ms, cleared: false };
  pendingTimeouts.push(handle);
  return handle;
}
function fakeClearTimeout(handle) {
  if (handle) handle.cleared = true;
}
function fakeSetInterval(fn, ms) {
  const handle = { fn, ms, cleared: false };
  pendingIntervals.push(handle);
  return handle;
}
function fakeClearInterval(handle) {
  if (handle) handle.cleared = true;
}

let mutationCallback = null;
let observedMutationRoot = null;
let mutationDisconnected = false;
class FakeMutationObserver {
  constructor(callback) {
    mutationCallback = callback;
  }

  observe(root, options) {
    observedMutationRoot = { root, options };
  }

  disconnect() {
    mutationDisconnected = true;
  }
}

const eventListeners = [];
const eventTarget = {
  addEventListener(name, handler) {
    eventListeners.push([name, handler]);
  },
  removeEventListener(name, handler) {
    const index = eventListeners.findIndex(([rowName, rowHandler]) => rowName === name && rowHandler === handler);
    if (index >= 0) eventListeners.splice(index, 1);
  },
};

const host = createMeetingPlatformRuntimeHost(runtime, googleConfig, {
  setTimeout: fakeSetTimeout,
  clearTimeout: fakeClearTimeout,
  setInterval: fakeSetInterval,
  clearInterval: fakeClearInterval,
  MutationObserver: FakeMutationObserver,
  mutationRoot: { nodeType: 1 },
  eventTarget,
  now: () => 1_782_614_401_234,
});

assert.equal(host.config.schema, 'meeting_platform_runtime_host_config');
assert.equal(host.getState().platform, 'google_meet');

const started = host.start();
assert.equal(started.running, true);
assert.equal(started.change_observer_installed, true);
assert.equal(started.lifecycle_installed, true);
assert.equal(pendingIntervals.length, 1);
assert.equal(observedMutationRoot.options.childList, true);
assert.equal(typeof mutationCallback, 'function');

mutationCallback([{ type: 'childList', target: { nodeType: 1 } }]);
assert.equal(pendingTimeouts.length, 1);
assert.equal(pendingTimeouts[0].ms, 150);
await pendingTimeouts[0].fn();
assert.equal(calls.at(-1).kind, 'sample');
assert.equal(calls.at(-1).options.trigger, 'dom_mutation');
assert.equal(calls.at(-1).options.captured_at_ms, 1_782_614_401_234);

const firstScheduled = host.changed();
const secondScheduled = host.changed();
assert.equal(firstScheduled.scheduled, true);
assert.equal(secondScheduled.scheduled, true);
assert.equal(pendingTimeouts.at(-2).cleared, true);
await pendingTimeouts.at(-1).fn();
assert.equal(calls.at(-1).options.trigger, 'dom_mutation');

await host.candidateMissing(undefined, { immediate: true });
assert.equal(calls.at(-1).options.trigger, 'meeting_candidate_missing');

const stopped = host.stop();
assert.equal(stopped.running, false);
assert.equal(mutationDisconnected, true);
assert.equal(pendingIntervals[0].cleared, true);
assert.equal(calls.at(-1).kind, 'stop');

const client = {
  async startMeeting(input) {
    return { ok: true, input };
  },
  async endMeeting(input) {
    return { ok: true, input };
  },
  async insertMark(input) {
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformRuntimeHostConfig('google-meet').schema, 'meeting_platform_runtime_host_config');
assert.equal(kit.platformRuntimeHostConfigMatrix().platform_count, 2);
assert.equal(kit.report().platform_runtime_host_config_matrix.host_ready_count, 2);
assert.equal(kit.platformRuntimeHostHandoff('google-meet').schema, 'meeting_platform_runtime_host_handoff');
assert.equal(kit.platformRuntimeHostHandoffMatrix().accepted_count, 2);
assert.equal(kit.report().platform_runtime_host_handoff_matrix.accepted_count, 2);
const kitHost = kit.createPlatformRuntimeHost(runtime, 'google-meet', {
  setTimeout: fakeSetTimeout,
  clearTimeout: fakeClearTimeout,
  setInterval: fakeSetInterval,
  clearInterval: fakeClearInterval,
});
assert.equal(kitHost.config.platform, 'google_meet');

console.log('ok meeting platform runtime host');
