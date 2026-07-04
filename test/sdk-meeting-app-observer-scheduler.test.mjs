import assert from 'node:assert/strict';

import {
  buildMeetingAppRuntimeObserverPlan,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-profile.mjs';
import {
  MEETING_APP_OBSERVER_SCHEDULER_CONFIG_MATRIX_SCHEMA,
  MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA,
  buildMeetingAppObserverSchedulerConfig,
  buildMeetingAppObserverSchedulerConfigMatrix,
  createMeetingAppObserverScheduler,
  normalizeMeetingAppObserverTrigger,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-observer-scheduler.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const startMs = 1_783_528_800_000;
const input = {
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
};

assert.equal(normalizeMeetingAppObserverTrigger('mutation'), 'dom_mutation');
assert.equal(normalizeMeetingAppObserverTrigger('keep-alive'), 'unchanged_keep_alive');
assert.equal(normalizeMeetingAppObserverTrigger('speaker'), 'active_speaker_candidate');
assert.equal(normalizeMeetingAppObserverTrigger('candidate_missing'), 'meeting_candidate_missing');

const observerPlan = buildMeetingAppRuntimeObserverPlan(input, {
  platform: 'google-meet',
  surface: 'browser-extension',
});
const config = buildMeetingAppObserverSchedulerConfig(observerPlan);
assert.equal(config.schema, MEETING_APP_OBSERVER_SCHEDULER_CONFIG_SCHEMA);
assert.equal(config.platform, 'google_meet');
assert.equal(config.surface, 'browser_extension');
assert.equal(config.runtime_factory, 'createMeetingAppBrowserRuntime');
assert.equal(config.timestamp_field, 'captured_at_ms');
assert.equal(config.cadence.fallback_poll_interval_ms, 10000);
assert.equal(config.cadence.meeting_missing_end_grace_ms, 4000);
assert.equal(config.trigger_names.includes('dom_mutation'), true);
assert.equal(config.track_sample_policy.enabled, false);

const matrix = buildMeetingAppObserverSchedulerConfigMatrix({
  platforms: ['google-meet', 'zoom'],
});
assert.equal(matrix.schema, MEETING_APP_OBSERVER_SCHEDULER_CONFIG_MATRIX_SCHEMA);
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.sdk_ready_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').runtime_factory, 'createMeetingAppBrowserRuntime');

const calls = [];
const runtime = {
  async sample(sampleInput, sampleOptions) {
    calls.push({ method: 'sample', input: sampleInput, options: sampleOptions });
    return {
      ok: true,
      signal_count: 1,
      input: sampleInput,
      options: sampleOptions,
    };
  },
  async sampleTracks(sampleInput, sampleOptions) {
    calls.push({ method: 'sampleTracks', input: sampleInput, options: sampleOptions });
    return {
      ok: true,
      new_mark_count: 1,
      input: sampleInput,
      options: sampleOptions,
    };
  },
};

const scheduler = createMeetingAppObserverScheduler(runtime, observerPlan, {
  now: () => startMs,
  observeTracks: true,
});
const changed = await scheduler.triggerChanged(input);
assert.equal(changed.sampled, true);
assert.equal(changed.trigger, 'dom_mutation');
assert.equal(changed.options.changedObserveEveryMs, 150);
assert.equal(calls[0].method, 'sample');
assert.equal(calls[0].input.url, 'https://meet.google.com/abc-defg-hij');
assert.equal(calls[0].options.trigger, 'dom_mutation');
assert.equal(calls[1].method, 'sampleTracks');
assert.equal(scheduler.getState().sample_count, 1);
assert.equal(scheduler.getState().track_sample_count, 1);

const speaker = await scheduler.triggerSpeakerCandidate(input);
assert.equal(speaker.trigger, 'active_speaker_candidate');
assert.equal(speaker.options.force, true);
assert.equal(speaker.options.stabilityFollowup, true);
assert.equal(calls.at(-1).method, 'sampleTracks');
assert.equal(scheduler.getState().track_sample_count, 2);

const providerCalls = [];
const providerRuntime = {
  inputProvider() {
    return input;
  },
  async sample(sampleOptions) {
    providerCalls.push({ method: 'sample', options: sampleOptions });
    return { ok: true, options: sampleOptions };
  },
};
const providerScheduler = createMeetingAppObserverScheduler(providerRuntime, observerPlan, {
  now: () => startMs + 1_000,
  runtimeInputMode: 'provider',
});
await providerScheduler.triggerKeepAlive();
assert.equal(providerCalls[0].options.trigger, 'unchanged_keep_alive');
assert.equal(providerCalls[0].options.capturedAtMs, startMs + 1_000);

const scheduledTimeouts = [];
const clearedTimeouts = [];
const schedulerWithTimers = createMeetingAppObserverScheduler(runtime, observerPlan, {
  now: () => startMs + 2_000,
  setTimeout(callback, delayMs) {
    const timer = { callback, delayMs };
    scheduledTimeouts.push(timer);
    return timer;
  },
  clearTimeout(timer) {
    clearedTimeouts.push(timer);
  },
});
const scheduled = schedulerWithTimers.triggerCandidateMissing(input, { schedule: true });
assert.equal(scheduled.scheduled, true);
assert.equal(scheduled.trigger, 'meeting_candidate_missing');
assert.equal(scheduled.delay_ms, 4000);
assert.equal(schedulerWithTimers.getState().pending_timer_count, 1);
await scheduledTimeouts[0].callback();
assert.equal(schedulerWithTimers.getState().pending_timer_count, 0);
assert.equal(calls.at(-1).options.missingCandidate, true);

const intervals = [];
const clearedIntervals = [];
const startScheduler = createMeetingAppObserverScheduler(runtime, observerPlan, {
  now: () => startMs + 3_000,
  setInterval(callback, delayMs) {
    const timer = { callback, delayMs };
    intervals.push(timer);
    return timer;
  },
  clearInterval(timer) {
    clearedIntervals.push(timer);
  },
});
const started = startScheduler.start(() => input, { immediate: false });
assert.equal(started.running, true);
assert.equal(intervals[0].delayMs, 10000);
assert.equal(startScheduler.stop().running, false);
assert.equal(clearedIntervals.length, 1);

const client = {
  async startMeeting(value) { return { ok: true, value }; },
  async endMeeting(value) { return { ok: true, value }; },
  async insertMark(value) { return { ok: true, value }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.meetingAppObserverSchedulerConfig('google-meet').platform, 'google_meet');
assert.equal(kit.meetingAppObserverSchedulerConfigMatrix().platform_count, 2);
assert.equal(kit.report().meeting_app_observer_scheduler_config_matrix.sdk_ready_count, 2);

console.log('ok meeting app observer scheduler');
