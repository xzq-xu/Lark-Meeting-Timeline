export class MeetingTimelineSdkError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = 'MeetingTimelineSdkError';
    this.details = details;
  }
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date);
}

export function compactObject(value, stack = new WeakSet()) {
  if (Array.isArray(value)) {
    if (stack.has(value)) return undefined;
    stack.add(value);
    const compacted = value.map((item) => compactObject(item, stack));
    stack.delete(value);
    return compacted;
  }
  if (!isPlainObject(value)) return value;
  if (stack.has(value)) return undefined;
  stack.add(value);
  const compacted = Object.fromEntries(
    Object.entries(value)
      .filter(([, item]) => item !== undefined)
      .map(([key, item]) => [key, compactObject(item, stack)])
      .filter(([, item]) => item !== undefined),
  );
  stack.delete(value);
  return compacted;
}

export function normalizeAbsoluteMs(value, fieldName = 'timestamp') {
  if (value == null || value === '') return undefined;
  if (value instanceof Date) {
    const ms = value.getTime();
    if (Number.isFinite(ms)) return ms;
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: Date is not finite`, { fieldName, value });
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new MeetingTimelineSdkError(`Invalid ${fieldName}: number is not finite`, { fieldName, value });
    }
    if (value > 10_000_000_000_000) return Math.round(value / 1000);
    if (value > 10_000_000_000) return Math.round(value);
    if (value > 1_000_000_000) return Math.round(value * 1000);
    throw new MeetingTimelineSdkError(`Invalid ${fieldName}: expected absolute unix time`, { fieldName, value });
  }
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return normalizeAbsoluteMs(numeric, fieldName);
  const parsed = Date.parse(String(value));
  if (Number.isFinite(parsed)) return parsed;
  throw new MeetingTimelineSdkError(`Invalid ${fieldName}: cannot parse absolute time`, { fieldName, value });
}
