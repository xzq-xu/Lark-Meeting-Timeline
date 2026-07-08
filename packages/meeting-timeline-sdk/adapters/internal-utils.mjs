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
