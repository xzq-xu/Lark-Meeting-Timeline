import type { MeetingTimelineClient } from '../index.mjs';
import type {
  LocalObserverCandidateResult,
  LocalObserverOptions,
  LocalObserverResult,
  LocalObserverState,
  LocalTimelineObserverCandidateResult,
  LocalTimelineObserverOptions,
  LocalTimelineObserverResult,
} from './local-observer.mjs';

export interface MeetingApplicationDetection {
  platform: string;
  confidence?: string;
  reason?: string;
  meeting?: Record<string, unknown>;
}

export interface MeetingSessionDiscoverySnapshot {
  url?: string;
  href?: string;
  meeting_url?: string;
  meetingUrl?: string;
  join_url?: string;
  joinUrl?: string;
  title?: string;
  topic?: string;
  name?: string;
  platform?: string;
  provider?: string;
  meeting_id?: string;
  meetingId?: string;
  external_meeting_id?: string;
  externalMeetingId?: string;
  session_id?: string;
  sessionId?: string;
  active?: boolean;
  focused?: boolean;
  selected?: boolean;
  visible?: boolean;
  in_meeting?: boolean;
  inMeeting?: boolean;
  observed_at_ms?: number | string | Date;
  observedAtMs?: number | string | Date;
  timestamp_ms?: number | string | Date;
  timestampMs?: number | string | Date;
  timestamp?: number | string | Date;
  application?: Record<string, unknown>;
  app?: Record<string, unknown>;
  process?: Record<string, unknown>;
  window?: Record<string, unknown>;
  browser?: Record<string, unknown>;
  tab?: Record<string, unknown>;
  meeting?: Record<string, unknown>;
  call?: Record<string, unknown>;
  conference?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MeetingSessionEnvironmentSnapshot extends MeetingSessionDiscoverySnapshot {
  windows?: MeetingSessionDiscoverySnapshot[];
  tabs?: MeetingSessionDiscoverySnapshot[];
  applications?: MeetingSessionDiscoverySnapshot[];
  apps?: MeetingSessionDiscoverySnapshot[];
  processes?: MeetingSessionDiscoverySnapshot[];
  candidates?: MeetingSessionDiscoverySnapshot[];
  snapshots?: MeetingSessionDiscoverySnapshot[];
  items?: MeetingSessionDiscoverySnapshot[];
}

export interface MeetingSessionDiscoveryOptions extends LocalObserverOptions {
  includeIdleApps?: boolean;
}

export interface NormalizedMeetingSessionCandidate extends MeetingSessionDiscoverySnapshot {
  platform: string;
  meeting_id: string;
  meeting_url?: string;
  title?: string;
  observedAtMs?: number;
  meeting: {
    platform: string;
    meeting_id: string;
    external_meeting_id?: string;
    meeting_url?: string;
    title?: string;
    confidence?: string;
  };
  discovery?: {
    platform_reason?: string;
    confidence?: string;
    app_name?: string;
    process_name?: string;
    bundle_id?: string;
  };
}

export interface MeetingSessionSelection {
  selectedSnapshot: MeetingSessionDiscoverySnapshot | null;
  detectedMeeting: Record<string, unknown> | null;
  candidates: Record<string, unknown>[];
  normalizedCandidates: NormalizedMeetingSessionCandidate[];
}

export function detectMeetingApplication(input?: MeetingSessionDiscoverySnapshot): MeetingApplicationDetection | null;

export function normalizeMeetingSessionCandidate(
  input?: MeetingSessionDiscoverySnapshot,
  options?: MeetingSessionDiscoveryOptions,
): NormalizedMeetingSessionCandidate | null;

export function normalizeMeetingSessionCandidates(
  input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
  options?: MeetingSessionDiscoveryOptions,
): NormalizedMeetingSessionCandidate[];

export function selectMeetingSessionCandidate(
  input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
  options?: MeetingSessionDiscoveryOptions,
): MeetingSessionSelection;

export function createMeetingSessionDiscoveryObserver(options?: MeetingSessionDiscoveryOptions): {
  observeEnvironment(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
    options?: MeetingSessionDiscoveryOptions,
  ): LocalObserverCandidateResult & { normalizedCandidates: NormalizedMeetingSessionCandidate[] };
  observeCandidate(
    input?: MeetingSessionDiscoverySnapshot,
    options?: MeetingSessionDiscoveryOptions,
  ): LocalObserverResult & { normalizedCandidate: NormalizedMeetingSessionCandidate | null };
  select(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
    options?: MeetingSessionDiscoveryOptions,
  ): MeetingSessionSelection;
  getState(): LocalObserverState;
  reset(nextState?: LocalObserverState | null): LocalObserverState;
};

export function createMeetingSessionTimelineDiscovery(
  client: MeetingTimelineClient,
  options?: LocalTimelineObserverOptions & MeetingSessionDiscoveryOptions,
): {
  observeEnvironment(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
    options?: LocalTimelineObserverOptions & MeetingSessionDiscoveryOptions,
  ): Promise<LocalTimelineObserverCandidateResult & { normalizedCandidates: NormalizedMeetingSessionCandidate[] }>;
  observeCandidate(
    input?: MeetingSessionDiscoverySnapshot,
    options?: LocalTimelineObserverOptions & MeetingSessionDiscoveryOptions,
  ): Promise<LocalTimelineObserverResult & { normalizedCandidate: NormalizedMeetingSessionCandidate | null }>;
  select(
    input?: MeetingSessionEnvironmentSnapshot | MeetingSessionDiscoverySnapshot[],
    options?: MeetingSessionDiscoveryOptions,
  ): MeetingSessionSelection;
  getState(): LocalObserverState;
  reset(nextState?: LocalObserverState | null): LocalObserverState;
};
