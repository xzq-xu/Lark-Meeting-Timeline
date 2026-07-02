import {
  MeetingTimelineSdkError,
  createMeetingTimelineClient,
} from '../index.mjs';
import { applyMeetingSignals } from './core.mjs';
import { createBrowserMeetingObserver } from './browser-meeting.mjs';
import { createLocalMeetingObserver } from './local-observer.mjs';
import { createMeetingAppObserver } from './meeting-apps.mjs';
import { createNativeMeetingObserver } from './native-meeting.mjs';
import { createReconciledPlatformEventIngestor } from './platform-ingest.mjs';
import { createMeetingSignalReconciler } from './signal-reconciler.mjs';
import { importPlatformTranscript } from './transcript.mjs';

function isTimelineClient(value) {
  return Boolean(value)
    && typeof value === 'object'
    && typeof value.startMeeting === 'function'
    && typeof value.endMeeting === 'function'
    && typeof value.insertMark === 'function';
}

function resolveClient(clientOrOptions, options = {}) {
  if (isTimelineClient(clientOrOptions)) return clientOrOptions;
  const clientOptions = {
    ...(clientOrOptions ?? {}),
    ...(options.clientOptions ?? options.client_options ?? {}),
  };
  if (!clientOptions.baseUrl && !clientOptions.base_url) {
    throw new MeetingTimelineSdkError('Meeting timeline client or client baseUrl is required for createMeetingSourceAggregator');
  }
  return createMeetingTimelineClient({
    ...clientOptions,
    baseUrl: clientOptions.baseUrl ?? clientOptions.base_url,
  });
}

function mergeApplyOptions(base = {}, next = {}) {
  return {
    ...(base.applyOptions ?? {}),
    ...(base.apply_options ?? {}),
    ...(next.applyOptions ?? {}),
    ...(next.apply_options ?? {}),
  };
}

function mergeReconcileOptions(base = {}, next = {}) {
  return {
    ...(base.reconcileOptions ?? {}),
    ...(base.reconcile_options ?? {}),
    ...(next.reconcileOptions ?? {}),
    ...(next.reconcile_options ?? {}),
  };
}

function defaultSource(kind) {
  return {
    browser: 'browser_meeting_observer',
    native: 'native_desktop_observer',
    local: 'local_meeting_observer',
  }[kind] ?? `${kind}_meeting_observer`;
}

function observerOptions(kind, options = {}) {
  const camel = `${kind}Options`;
  const snake = `${kind.replace(/[A-Z]/g, (char) => `_${char.toLowerCase()}`)}_options`;
  return {
    source: options.source ?? defaultSource(kind),
    ...(options[camel] ?? {}),
    ...(options[snake] ?? {}),
  };
}

function meetingAppOptions(options = {}) {
  return {
    source: options.source ?? 'meeting_app_observer',
    ...(options.appOptions ?? {}),
    ...(options.app_options ?? {}),
    ...(options.meetingAppOptions ?? {}),
    ...(options.meeting_app_options ?? {}),
    speakerOptions: options.meetingAppOptions?.speakerOptions
      ?? options.meeting_app_options?.speakerOptions
      ?? options.meetingAppOptions?.speaker_options
      ?? options.meeting_app_options?.speaker_options
      ?? options.appOptions?.speakerOptions
      ?? options.app_options?.speakerOptions
      ?? options.appOptions?.speaker_options
      ?? options.app_options?.speaker_options
      ?? options.speakerOptions
      ?? options.speaker_options,
  };
}

function signalPreview(signal = {}) {
  return {
    type: signal.type,
    occurred_at_ms: signal.occurred_at_ms,
    source: signal.source,
    meeting: signal.meeting,
    speaker_id: signal.speaker_id,
    speaker_name: signal.speaker_name,
    participant_id: signal.participant_id,
    participant_name: signal.participant_name,
  };
}

async function applyObservedSignals(client, reconciler, observed = {}, baseOptions = {}, nextOptions = {}) {
  const rawSignals = observed.signals ?? [];
  if (rawSignals.length === 0) {
    return {
      ...observed,
      rawSignals,
      reconciliation: {
        signals: [],
        skipped: [],
        decisions: [],
        state: reconciler.getState(),
      },
      results: [],
    };
  }
  const reconciliation = reconciler.reconcile(rawSignals, mergeReconcileOptions(baseOptions, nextOptions));
  const results = reconciliation.signals.length === 0
    ? []
    : await applyMeetingSignals(client, reconciliation.signals, mergeApplyOptions(baseOptions, nextOptions));
  return {
    ...observed,
    rawSignals,
    signals: reconciliation.signals,
    reconciliation,
    results,
  };
}

function sourceDiagnostic(kind, result = {}) {
  const rawSignals = result.rawSignals ?? result.signals ?? [];
  return {
    source: kind,
    raw_signal_count: rawSignals.length,
    applied_signal_count: result.signals?.length ?? 0,
    skipped_count: result.reconciliation?.skipped?.length ?? 0,
    result_count: result.results?.length ?? 0,
    signals: (result.signals ?? []).map(signalPreview),
    skipped: result.reconciliation?.skipped,
  };
}

export function createMeetingSourceAggregator(clientOrOptions, options = {}) {
  const client = resolveClient(clientOrOptions, options);
  const reconciler = options.reconciler
    ?? options.signalReconciler
    ?? options.signal_reconciler
    ?? createMeetingSignalReconciler(options.reconcileOptions ?? options.reconcile_options);
  const browserObserver = createBrowserMeetingObserver(observerOptions('browser', {
    ...options,
    browserOptions: {
      ...(options.browserOptions ?? {}),
      ...(options.browser_options ?? {}),
      speakerOptions: options.browserOptions?.speakerOptions
        ?? options.browser_options?.speakerOptions
        ?? options.browserOptions?.speaker_options
        ?? options.browser_options?.speaker_options
        ?? options.speakerOptions
        ?? options.speaker_options,
    },
  }));
  const nativeObserver = createNativeMeetingObserver(observerOptions('native', {
    ...options,
    nativeOptions: {
      ...(options.nativeOptions ?? {}),
      ...(options.native_options ?? {}),
      speakerOptions: options.nativeOptions?.speakerOptions
        ?? options.native_options?.speakerOptions
        ?? options.nativeOptions?.speaker_options
        ?? options.native_options?.speaker_options
        ?? options.speakerOptions
        ?? options.speaker_options,
    },
  }));
  const appObserver = createMeetingAppObserver(meetingAppOptions(options));
  const localObserver = createLocalMeetingObserver(observerOptions('local', options));
  const providerIngestor = createReconciledPlatformEventIngestor(client, {
    ...(options.providerOptions ?? {}),
    ...(options.provider_options ?? {}),
    reconciler,
    applyOptions: mergeApplyOptions(options, options.providerOptions ?? options.provider_options ?? {}),
  });
  async function observeMeetingAppInput(input = {}, observeOptions = {}) {
    const observed = appObserver.observe(input, observeOptions);
    const result = await applyObservedSignals(client, reconciler, observed, options, observeOptions);
    return {
      ...result,
      source: 'meeting_app',
      diagnostic: sourceDiagnostic('meeting_app', result),
    };
  }

  return {
    client,
    async observeBrowser(input = {}, observeOptions = {}) {
      const observed = browserObserver.observe(input, observeOptions);
      const result = await applyObservedSignals(client, reconciler, observed, options, observeOptions);
      return {
        ...result,
        source: 'browser',
        diagnostic: sourceDiagnostic('browser', result),
      };
    },
    async observeNative(input = {}, observeOptions = {}) {
      const observed = nativeObserver.observe(input, observeOptions);
      const result = await applyObservedSignals(client, reconciler, observed, options, observeOptions);
      return {
        ...result,
        source: 'native',
        diagnostic: sourceDiagnostic('native', result),
      };
    },
    async observeMeetingApp(input = {}, observeOptions = {}) {
      return observeMeetingAppInput(input, observeOptions);
    },
    async observeApp(input = {}, observeOptions = {}) {
      return observeMeetingAppInput(input, observeOptions);
    },
    async observeLocal(snapshot = {}, observeOptions = {}) {
      const observed = localObserver.observe(snapshot, observeOptions);
      const result = await applyObservedSignals(client, reconciler, observed, options, observeOptions);
      return {
        ...result,
        source: 'local',
        diagnostic: sourceDiagnostic('local', result),
      };
    },
    async observeLocalCandidates(candidates = [], observeOptions = {}) {
      const observed = localObserver.observeCandidates(candidates, observeOptions);
      const result = await applyObservedSignals(client, reconciler, observed, options, observeOptions);
      return {
        ...result,
        source: 'local_candidates',
        diagnostic: sourceDiagnostic('local_candidates', result),
      };
    },
    async ingestProvider(platformOrInput, payload, ingestOptions = {}) {
      const result = await providerIngestor.ingest(platformOrInput, payload, {
        ...options,
        ...ingestOptions,
      });
      return {
        ...result,
        source: 'provider',
        diagnostic: sourceDiagnostic('provider', result),
      };
    },
    async ingestSignals(signals = [], ingestOptions = {}) {
      const observed = { signals: Array.isArray(signals) ? signals : [signals] };
      const result = await applyObservedSignals(client, reconciler, observed, options, ingestOptions);
      return {
        ...result,
        source: 'signals',
        diagnostic: sourceDiagnostic('signals', result),
      };
    },
    insertMark(input = {}, markOptions = {}) {
      return client.insertMark(input, markOptions);
    },
    insertAnnotation(input = {}, markOptions = {}) {
      return client.insertMark(input, markOptions);
    },
    insertMarks(inputs = [], markOptions = {}) {
      if (typeof client.insertMarks !== 'function') {
        throw new MeetingTimelineSdkError('Meeting timeline client is missing insertMarks()');
      }
      return client.insertMarks(inputs, markOptions);
    },
    importTranscript(input = {}, transcriptOptions = {}) {
      return importPlatformTranscript(client, input, transcriptOptions);
    },
    startMeeting(input = {}) {
      return client.startMeeting(input);
    },
    endMeeting(input = {}) {
      return client.endMeeting(input);
    },
    getState() {
      return {
        browser: browserObserver.getState(),
        native: nativeObserver.getState(),
        meetingApp: appObserver.getState(),
        local: localObserver.getState(),
        reconciler: reconciler.getState(),
        client: typeof client.getState === 'function' ? client.getState() : undefined,
      };
    },
    reset(nextState = {}) {
      return {
        browser: browserObserver.reset(nextState.browser ?? nextState.browserState ?? null),
        native: nativeObserver.reset(nextState.native ?? nextState.nativeState ?? null),
        meetingApp: appObserver.reset(nextState.meetingApp ?? nextState.meeting_app ?? nextState.app ?? null),
        local: localObserver.reset(nextState.local ?? nextState.localState ?? null),
        reconciler: reconciler.reset(nextState.reconciler ?? nextState.signal_reconciler ?? nextState.signalReconciler ?? {}),
      };
    },
  };
}
