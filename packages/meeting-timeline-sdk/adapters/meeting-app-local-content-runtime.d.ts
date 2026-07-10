export const MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA: 'meeting_app_local_content_runtime';
export const MEETING_APP_LOCAL_CONTENT_RUNTIME_SCHEMA_VERSION: 1;

export interface MeetingAppLocalContentRuntime {
  type: 'meeting_app_local_content_runtime';
  schema: 'meeting_app_local_content_runtime';
  schema_version: 1;
  client: Record<string, unknown>;
  options: Record<string, unknown>;
  sample(options?: Record<string, unknown>): Promise<Record<string, unknown>>;
  preflight(): Record<string, unknown>;
  handleMessage(message?: Record<string, unknown>): Promise<Record<string, unknown>>;
  start(): Record<string, unknown>;
  stop(): Record<string, unknown>;
  getState(): Record<string, unknown>;
}

export function createMeetingAppLocalContentRuntime(
  client: Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppLocalContentRuntime;

export function installMeetingAppLocalContentRuntime(
  client: Record<string, unknown>,
  options?: Record<string, unknown>,
): MeetingAppLocalContentRuntime;

export default createMeetingAppLocalContentRuntime;
