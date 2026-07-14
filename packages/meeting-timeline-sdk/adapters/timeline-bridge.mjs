import {
  MeetingTimelineSdkError,
  createMeetingTimelineClient,
} from '../index.mjs';
import { createLocalMeetingObserver } from './local-observer.mjs';
import { createReconciledPlatformEventIngestor } from './platform-ingest.mjs';
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
    throw new MeetingTimelineSdkError('Meeting timeline client or client baseUrl is required for createMeetingTimelineBridge');
  }
  return createMeetingTimelineClient({
    ...clientOptions,
    baseUrl: clientOptions.baseUrl ?? clientOptions.base_url,
  });
}

function localObserverOptions(options = {}) {
  return {
    source: options.source ?? 'timeline_bridge',
    ...(options.localObserverOptions ?? {}),
    ...(options.local_observer_options ?? {}),
    applyOptions: {
      ...(options.applyOptions ?? {}),
      ...(options.apply_options ?? {}),
      ...(options.localObserverOptions?.applyOptions ?? {}),
      ...(options.local_observer_options?.applyOptions ?? {}),
      ...(options.localObserverOptions?.apply_options ?? {}),
      ...(options.local_observer_options?.apply_options ?? {}),
    },
  };
}

function ingestorOptions(options = {}) {
  return {
    ...(options.ingestOptions ?? {}),
    ...(options.ingest_options ?? {}),
    applyOptions: {
      ...(options.applyOptions ?? {}),
      ...(options.apply_options ?? {}),
      ...(options.ingestOptions?.applyOptions ?? {}),
      ...(options.ingest_options?.applyOptions ?? {}),
      ...(options.ingestOptions?.apply_options ?? {}),
      ...(options.ingest_options?.apply_options ?? {}),
    },
  };
}

export function createMeetingTimelineBridge(clientOrOptions, options = {}) {
  const client = resolveClient(clientOrOptions, options);
  const observer = createLocalMeetingObserver(localObserverOptions(options));
  const ingestor = createReconciledPlatformEventIngestor(client, ingestorOptions(options));
  async function applyObservedSignals(observed, ingestOptions = {}) {
    if (!observed.signals?.length) {
      return {
        ...observed,
        results: [],
      };
    }
    const ingestion = await ingestor.ingest('local-detector', observed.signals, ingestOptions);
    return {
      ...observed,
      signals: ingestion.signals,
      rawSignals: observed.signals,
      reconciliation: ingestion.reconciliation,
      ingestion,
      results: ingestion.results,
    };
  }
  return {
    client,
    async observe(snapshot = {}, observeOptions = {}) {
      const observed = observer.observe(snapshot, observeOptions);
      return applyObservedSignals(observed, observeOptions);
    },
    async observeCandidates(candidates = [], observeOptions = {}) {
      const observed = observer.observeCandidates(candidates, observeOptions);
      return applyObservedSignals(observed, observeOptions);
    },
    ingest(platformOrInput, payload, ingestOptions = {}) {
      return ingestor.ingest(platformOrInput, payload, ingestOptions);
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
      return typeof client.getState === 'function' ? client.getState() : undefined;
    },
    getBridgeState() {
      return {
        local_observer: observer.getState(),
        signal_reconciler: ingestor.getState(),
      };
    },
    reset(nextState = {}) {
      return {
        local_observer: observer.reset(nextState.local_observer ?? nextState.localObserver ?? null),
        signal_reconciler: ingestor.reset(nextState.signal_reconciler ?? nextState.signalReconciler ?? {}),
      };
    },
  };
}
