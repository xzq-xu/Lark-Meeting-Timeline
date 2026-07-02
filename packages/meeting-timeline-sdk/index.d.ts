export const SDK_VERSION: string;

export interface MeetingStartInput {
  platform?: string;
  meeting_id?: string;
  meetingId?: string;
  external_meeting_id?: string;
  externalMeetingId?: string;
  meeting_no?: string;
  meetingNo?: string;
  meeting_url?: string;
  meetingUrl?: string;
  url?: string;
  title?: string;
  topic?: string;
  name?: string;
  timezone?: string;
  start_time_ms?: number | string | Date;
  startTimeMs?: number | string | Date;
  start_time?: number | string | Date;
  startTime?: number | string | Date;
  detector_source?: string;
  detectorSource?: string;
  force?: boolean;
  [key: string]: unknown;
}

export interface MeetingEndInput {
  meeting_id?: string;
  meetingId?: string;
  end_time_ms?: number | string | Date;
  endTimeMs?: number | string | Date;
  end_time?: number | string | Date;
  endTime?: number | string | Date;
  time_ms?: number;
  timeMs?: number;
  detector_source?: string;
  detectorSource?: string;
  [key: string]: unknown;
}

export interface TranscriptSegmentInput {
  id?: string;
  segment_id?: string;
  segmentId?: string;
  sentence_id?: string;
  sentenceId?: string;
  start_ms?: number | string;
  startMs?: number | string;
  start_time_ms?: number | string | Date;
  startTimeMs?: number | string | Date;
  start_time?: number | string | Date;
  startTime?: number | string | Date;
  start?: number | string | Date;
  end_ms?: number | string;
  endMs?: number | string;
  end_time_ms?: number | string | Date;
  endTimeMs?: number | string | Date;
  end_time?: number | string | Date;
  endTime?: number | string | Date;
  end?: number | string | Date;
  speaker_id?: string;
  speakerId?: string;
  speaker_name?: string;
  speakerName?: string;
  participant_id?: string;
  participantId?: string;
  participant_name?: string;
  participantName?: string;
  user_id?: string;
  userId?: string;
  user_name?: string;
  userName?: string;
  text?: string;
  content?: string;
  sentence?: string;
  transcript?: string;
  language?: string;
  language_code?: string;
  languageCode?: string;
  source?: string;
  raw?: unknown;
  [key: string]: unknown;
}

export interface TranscriptImportInput {
  meeting?: MeetingStartInput;
  meetingSession?: MeetingStartInput;
  session?: MeetingStartInput;
  meeting_id?: string;
  meetingId?: string;
  platform?: string;
  source?: string;
  transcript?: TranscriptSegmentInput[];
  segments?: TranscriptSegmentInput[];
  entries?: TranscriptSegmentInput[];
  items?: TranscriptSegmentInput[];
  artifact?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface TimelineMarkInput {
  id?: string;
  annotation_id?: string;
  annotationId?: string;
  source?: string;
  device_id?: string;
  deviceId?: string;
  captured_at_ms?: number | string | Date;
  capturedAtMs?: number | string | Date;
  captured_at?: number | string | Date;
  capturedAt?: number | string | Date;
  time_ms?: number;
  timeMs?: number;
  kind?: string;
  type?: string;
  label?: string;
  text?: string;
  reading?: string;
  text_candidates?: string[];
  textCandidates?: string[];
  intent?: string;
  mark?: Record<string, unknown>;
  target?: Record<string, unknown>;
  target_region?: Record<string, unknown>;
  strokes?: unknown[];
  stroke_points?: unknown[];
  device?: Record<string, unknown>;
  payload?: Record<string, unknown>;
  realtime?: boolean;
  live?: boolean;
  meeting_session?: MeetingStartInput;
  meetingSession?: MeetingStartInput;
  start_meeting_session?: boolean;
  startMeetingSession?: boolean;
  force_meeting_session?: boolean;
  forceMeetingSession?: boolean;
  [key: string]: unknown;
}

export interface MeetingTimelineClientOptions {
  baseUrl: string;
  fetch?: typeof fetch;
  fetchImpl?: typeof fetch;
  headers?: Record<string, string>;
  authToken?: string;
  token?: string;
  source?: string;
  deviceId?: string;
  device_id?: string;
  detectorSource?: string;
  detector_source?: string;
  requireCapturedAt?: boolean;
  timeoutMs?: number;
}

export class MeetingTimelineSdkError extends Error {
  details: Record<string, unknown>;
}

export class MeetingTimelineApiError extends MeetingTimelineSdkError {
  status: number | null;
  body: unknown;
}

export function compactObject<T>(value: T): T;
export function normalizeAbsoluteMs(value: number | string | Date, fieldName?: string): number;
export function buildMeetingStartPayload(input?: MeetingStartInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildMeetingEndPayload(input?: MeetingEndInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildTranscriptImportPayload(input?: TranscriptImportInput, defaults?: Record<string, unknown>): Record<string, unknown>;
export function buildTimelineMark(input?: TimelineMarkInput, defaults?: Record<string, unknown>): Record<string, unknown>;

export class MeetingTimelineClient {
  constructor(options: MeetingTimelineClientOptions);
  request(path: string, options?: Record<string, unknown>): Promise<unknown>;
  getState(): Promise<unknown>;
  getIngestInfo(): Promise<unknown>;
  getMeetingSessionStatus(): Promise<unknown>;
  startMeeting(input?: MeetingStartInput): Promise<unknown>;
  endMeeting(input?: MeetingEndInput): Promise<unknown>;
  insertMark(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  addAnnotation(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertAnnotation(input?: TimelineMarkInput, options?: Record<string, unknown>): Promise<unknown>;
  insertMarks(inputs?: TimelineMarkInput[] | { annotations?: TimelineMarkInput[]; items?: TimelineMarkInput[] }, options?: Record<string, unknown>): Promise<unknown>;
  addAnnotations(inputs?: TimelineMarkInput[] | { annotations?: TimelineMarkInput[]; items?: TimelineMarkInput[] }, options?: Record<string, unknown>): Promise<unknown>;
  getAnnotationStatus(id: string): Promise<unknown>;
  importTranscript(input?: TranscriptImportInput, options?: Record<string, unknown>): Promise<unknown>;
  importMeetingTranscript(input?: TranscriptImportInput, options?: Record<string, unknown>): Promise<unknown>;
  subscribeState(options?: Record<string, unknown>): { stream: unknown; close(): void };
}

export function createMeetingTimelineClient(options: MeetingTimelineClientOptions): MeetingTimelineClient;
export const createMeetingTimelineSdk: typeof createMeetingTimelineClient;
