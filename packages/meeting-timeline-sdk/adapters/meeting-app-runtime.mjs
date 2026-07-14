import { MeetingTimelineSdkError } from '../index.mjs';
import { createMeetingAppDomMonitor } from './meeting-app-monitor.mjs';
import { createMeetingSourceAggregator } from './meeting-source.mjs';
import { createMeetingAppTrackRuntime } from './meeting-app-track-runtime.mjs';

function isMeetingSource(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.observeMeetingApp === 'function'
    && typeof value.insertMark === 'function';
}

function sourceOptions(options = {}) {
  const explicit = {
    ...(options.sourceOptions ?? {}),
    ...(options.source_options ?? {}),
  };
  return {
    ...options,
    ...explicit,
  };
}

function monitorOptions(options = {}) {
  return {
    sampleIntervalMs: options.sampleIntervalMs,
    sample_interval_ms: options.sample_interval_ms,
    minObserveIntervalMs: options.minObserveIntervalMs,
    min_observe_interval_ms: options.min_observe_interval_ms,
    changedObserveEveryMs: options.changedObserveEveryMs,
    changed_observe_every_ms: options.changed_observe_every_ms,
    unchangedObserveEveryMs: options.unchangedObserveEveryMs,
    unchanged_observe_every_ms: options.unchanged_observe_every_ms,
    immediate: options.immediate,
    input: options.input,
    inputProvider: options.inputProvider,
    input_provider: options.input_provider,
    now: options.now,
    nowMs: options.nowMs,
    now_ms: options.now_ms,
    captureOptions: {
      ...(options.captureOptions ?? {}),
      ...(options.capture_options ?? {}),
    },
    observeOptions: {
      ...(options.observeOptions ?? {}),
      ...(options.observe_options ?? {}),
    },
    ...(options.monitorOptions ?? {}),
    ...(options.monitor_options ?? {}),
  };
}

function trackRuntimeOptions(options = {}) {
  return {
    ...options,
    ...(options.trackRuntimeOptions ?? {}),
    ...(options.track_runtime_options ?? {}),
  };
}

function trackMonitorOptions(options = {}) {
  return monitorOptions({
    ...options,
    ...(options.trackMonitorOptions ?? {}),
    ...(options.track_monitor_options ?? {}),
  });
}

function resolveSources(clientOrOptions, options = {}) {
  const explicit = options.sources ?? options.meetingSources ?? options.meeting_sources;
  if (explicit) {
    if (!isMeetingSource(explicit)) {
      throw new MeetingTimelineSdkError('meeting app runtime sources must expose observeMeetingApp() and insertMark()', {
        reason: 'invalid_meeting_sources',
      });
    }
    return explicit;
  }
  return createMeetingSourceAggregator(clientOrOptions, sourceOptions(options));
}

export function createMeetingAppTimelineRuntime(clientOrOptions, options = {}) {
  const sources = resolveSources(clientOrOptions, options);
  const monitor = options.monitor ?? createMeetingAppDomMonitor(sources, monitorOptions(options));
  const tracks = options.tracks
    ?? options.trackRuntime
    ?? options.track_runtime
    ?? createMeetingAppTrackRuntime(sources.client, trackRuntimeOptions(options));
  const trackMonitor = options.trackMonitor
    ?? options.track_monitor
    ?? createMeetingAppDomMonitor(tracks, trackMonitorOptions(options));

  return {
    client: sources.client,
    sources,
    monitor,
    tracks,
    trackMonitor,
    sample(input = {}, sampleOptions = {}) {
      return monitor.sample(input, sampleOptions);
    },
    tick(input = {}, sampleOptions = {}) {
      return monitor.tick(input, sampleOptions);
    },
    start(inputProvider = null, startOptions = {}) {
      return monitor.start(inputProvider, startOptions);
    },
    stop() {
      return monitor.stop();
    },
    sampleTracks(input = {}, sampleOptions = {}) {
      return trackMonitor.sample(input, sampleOptions);
    },
    tickTracks(input = {}, sampleOptions = {}) {
      return trackMonitor.tick(input, sampleOptions);
    },
    startTracks(inputProvider = null, startOptions = {}) {
      return trackMonitor.start(inputProvider, startOptions);
    },
    stopTracks() {
      return trackMonitor.stop();
    },
    observeMeetingAppTracks(input = {}, observeOptions = {}) {
      return tracks.observe(input, observeOptions);
    },
    previewMeetingAppTracks(input = {}, previewOptions = {}) {
      return tracks.preview(input, previewOptions);
    },
    observeMeetingApp(input = {}, observeOptions = {}) {
      return sources.observeMeetingApp(input, observeOptions);
    },
    observeApp(input = {}, observeOptions = {}) {
      return sources.observeApp(input, observeOptions);
    },
    observeBrowser(input = {}, observeOptions = {}) {
      return sources.observeBrowser(input, observeOptions);
    },
    observeNative(input = {}, observeOptions = {}) {
      return sources.observeNative(input, observeOptions);
    },
    observeLocal(snapshot = {}, observeOptions = {}) {
      return sources.observeLocal(snapshot, observeOptions);
    },
    ingestProvider(platformOrInput, payload, ingestOptions = {}) {
      return sources.ingestProvider(platformOrInput, payload, ingestOptions);
    },
    ingestSignals(signals = [], ingestOptions = {}) {
      return sources.ingestSignals(signals, ingestOptions);
    },
    insertMark(input = {}, markOptions = {}) {
      return sources.insertMark(input, markOptions);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return sources.insertAnnotation(input, markOptions);
    },
    insertMarks(inputs = [], markOptions = {}) {
      return sources.insertMarks(inputs, markOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return sources.importTranscript(input, transcriptOptions);
    },
    startMeeting(input = {}) {
      return sources.startMeeting(input);
    },
    endMeeting(input = {}) {
      return sources.endMeeting(input);
    },
    getState() {
      return {
        monitor: monitor.getState(),
        trackMonitor: trackMonitor.getState(),
        tracks: tracks.getState(),
        sources: sources.getState(),
      };
    },
    reset(nextState = {}) {
      return {
        monitor: monitor.reset(nextState.monitor ?? nextState.monitorState ?? nextState.monitor_state ?? null),
        trackMonitor: trackMonitor.reset(nextState.trackMonitor ?? nextState.track_monitor ?? nextState.trackMonitorState ?? null),
        tracks: tracks.reset(nextState.tracks ?? nextState.trackRuntime ?? nextState.track_runtime ?? {}),
        sources: sources.reset(nextState.sources ?? nextState.source ?? nextState.meetingSources ?? {}),
      };
    },
  };
}
