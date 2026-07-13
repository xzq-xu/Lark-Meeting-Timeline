export class MeetingTimelineSdkError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = 'MeetingTimelineSdkError';
    this.details = details;
  }
}

export class MeetingTimelineApiError extends MeetingTimelineSdkError {
  constructor(message, details = {}) {
    super(message, details);
    this.name = 'MeetingTimelineApiError';
    this.status = details.status ?? null;
    this.body = details.body ?? null;
  }
}
