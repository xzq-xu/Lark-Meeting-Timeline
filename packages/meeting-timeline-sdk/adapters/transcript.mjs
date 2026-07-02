import { buildTranscriptImportPayload, compactObject } from '../index.mjs';

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function getPath(raw, path) {
  const parts = path.split('.');
  let node = raw;
  for (const part of parts) node = node?.[part];
  return node;
}

function candidateArray(raw, keys) {
  if (Array.isArray(raw)) return raw;
  for (const key of keys) {
    const value = getPath(raw, key);
    if (Array.isArray(value)) return value;
  }
  return [];
}

function lastResourceId(name) {
  const parts = String(name || '').split('/').filter(Boolean);
  return parts.at(-1);
}

function resourceId(name, resourceName) {
  const text = String(name || '');
  const match = text.match(new RegExp(`${resourceName}/([^/]+)`));
  return match?.[1];
}

function stripTimedTextMarkup(text = '') {
  return String(text)
    .replace(/<v\s+([^>]+)>/gi, '$1: ')
    .replace(/<\/v>/gi, '')
    .replace(/<[^>]+>/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function speakerFromTimedText(text = '') {
  const stripped = stripTimedTextMarkup(text);
  const match = stripped.match(/^([^:：]{1,80})[:：]\s+(.+)$/);
  if (!match) return { speaker_name: undefined, text: stripped };
  return {
    speaker_name: match[1].trim(),
    text: match[2].trim(),
  };
}

function parseTimestampOffsetMs(value = '') {
  const text = String(value).trim().replace(',', '.');
  const match = text.match(/^(?:(\d{1,2}):)?(\d{2}):(\d{2})(?:\.(\d{1,3}))?$/);
  if (!match) return undefined;
  const hours = Number(match[1] ?? 0);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const millis = Number((match[4] ?? '').padEnd(3, '0') || 0);
  return ((hours * 60 + minutes) * 60 + seconds) * 1000 + millis;
}

function linesOf(raw) {
  if (raw == null) return [];
  const text = typeof raw === 'string' ? raw : String(raw);
  return text.replace(/\r/g, '').split('\n');
}

export function parseTimedTextTranscript(raw = '', options = {}) {
  const source = options.source ?? 'timed_text_transcript';
  const language = options.language ?? options.language_code;
  const rows = [];
  const lines = linesOf(raw);
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index].trim();
    const match = line.match(/^(.+?)\s+-->\s+(.+?)(?:\s+.*)?$/);
    if (!match) continue;
    const startMs = parseTimestampOffsetMs(match[1]);
    const endMs = parseTimestampOffsetMs(match[2]);
    const textLines = [];
    while (index + 1 < lines.length && lines[index + 1].trim()) {
      index += 1;
      textLines.push(lines[index]);
    }
    const speaker = speakerFromTimedText(textLines.join(' '));
    if (startMs == null || endMs == null || !speaker.text) continue;
    rows.push(compactObject({
      id: `${source}-${rows.length + 1}`,
      start_ms: startMs,
      end_ms: Math.max(startMs + 1, endMs),
      speaker_name: speaker.speaker_name,
      text: speaker.text,
      language,
      source,
      raw: {
        cue: line,
        text: textLines.join('\n'),
      },
    }));
  }
  return rows;
}

export function normalizeGoogleMeetTranscriptEntries(raw = {}, options = {}) {
  const rows = candidateArray(raw, [
    'transcriptEntries',
    'entries',
    'data.transcriptEntries',
    'data.entries',
    'items',
    'segments',
  ]);
  return rows.map((entry, index) => {
    const participant = firstNonEmpty(entry.participant?.name, entry.participant, entry.participantName, entry.speaker);
    return compactObject({
      id: firstNonEmpty(lastResourceId(entry.name), entry.id, entry.entry_id, `google-meet-entry-${index + 1}`),
      start_time: firstNonEmpty(entry.startTime, entry.start_time),
      end_time: firstNonEmpty(entry.endTime, entry.end_time),
      speaker_id: resourceId(participant, 'participants') ?? firstNonEmpty(entry.participant?.id, entry.participant_id, entry.participantId),
      speaker_name: firstNonEmpty(
        entry.participant?.displayName,
        entry.participant?.display_name,
        entry.speaker_name,
        entry.speakerName,
      ),
      text: firstNonEmpty(entry.text, entry.content),
      language: firstNonEmpty(entry.languageCode, entry.language_code, options.language),
      source: options.source ?? 'google_meet_transcript',
      raw: entry,
    });
  }).filter((item) => item.text);
}

export function normalizeMicrosoftTeamsTranscript(raw = {}, options = {}) {
  if (typeof raw === 'string') {
    return parseTimedTextTranscript(raw, {
      ...options,
      source: options.source ?? 'microsoft_teams_transcript',
    });
  }
  const content = firstNonEmpty(raw.content, raw.transcript, raw.text, raw.body);
  if (typeof content === 'string') {
    return parseTimedTextTranscript(content, {
      ...options,
      source: options.source ?? 'microsoft_teams_transcript',
    });
  }
  const rows = candidateArray(raw, [
    'segments',
    'items',
    'transcript',
    'data.segments',
    'data.items',
  ]);
  return rows.map((item, index) => compactObject({
    id: firstNonEmpty(item.id, item.segment_id, item.sentence_id, `teams-seg-${index + 1}`),
    start_ms: firstNonEmpty(item.start_ms, item.startMs, item.offset_ms, item.offsetMs),
    end_ms: firstNonEmpty(item.end_ms, item.endMs),
    start_time: firstNonEmpty(item.start_time, item.startTime, item.start),
    end_time: firstNonEmpty(item.end_time, item.endTime, item.end),
    speaker_id: firstNonEmpty(item.speaker_id, item.speakerId, item.user_id, item.userId),
    speaker_name: firstNonEmpty(item.speaker_name, item.speakerName, item.user_name, item.userName),
    text: firstNonEmpty(item.text, item.content),
    language: firstNonEmpty(item.language, item.language_code, item.languageCode, options.language),
    source: options.source ?? 'microsoft_teams_transcript',
    raw: item,
  })).filter((item) => item.text);
}

export function normalizeZoomTranscript(raw = {}, options = {}) {
  const content = typeof raw === 'string'
    ? raw
    : firstNonEmpty(raw.vtt, raw.content, raw.transcript, raw.text, raw.body);
  if (typeof content === 'string') {
    return parseTimedTextTranscript(content, {
      ...options,
      source: options.source ?? 'zoom_transcript_vtt',
    });
  }
  const rows = candidateArray(raw, ['segments', 'items', 'transcript', 'data.segments']);
  return rows.map((item, index) => compactObject({
    id: firstNonEmpty(item.id, item.segment_id, `zoom-seg-${index + 1}`),
    start_ms: firstNonEmpty(item.start_ms, item.startMs, item.offset_ms, item.offsetMs),
    end_ms: firstNonEmpty(item.end_ms, item.endMs),
    speaker_name: firstNonEmpty(item.speaker_name, item.speakerName),
    text: firstNonEmpty(item.text, item.content),
    language: firstNonEmpty(item.language, item.language_code, item.languageCode, options.language),
    source: options.source ?? 'zoom_transcript_vtt',
    raw: item,
  })).filter((item) => item.text);
}

export function normalizeWebexTranscript(raw = {}, options = {}) {
  const content = typeof raw === 'string'
    ? raw
    : firstNonEmpty(raw.vtt, raw.content, raw.transcript, raw.text, raw.body, raw.data?.content, raw.data?.text);
  if (typeof content === 'string') {
    return parseTimedTextTranscript(content, {
      ...options,
      source: options.source ?? 'webex_transcript',
    });
  }
  const rows = candidateArray(raw, ['segments', 'items', 'transcript', 'snippets', 'data.segments', 'data.items', 'data.snippets']);
  return rows.map((item, index) => compactObject({
    id: firstNonEmpty(item.id, item.snippet_id, item.snippetId, item.segment_id, `webex-seg-${index + 1}`),
    start_ms: firstNonEmpty(item.start_ms, item.startMs, item.offset_ms, item.offsetMs),
    end_ms: firstNonEmpty(item.end_ms, item.endMs),
    start_time: firstNonEmpty(item.start_time, item.startTime, item.start),
    end_time: firstNonEmpty(item.end_time, item.endTime, item.end),
    speaker_id: firstNonEmpty(item.speaker_id, item.speakerId, item.person_id, item.personId),
    speaker_name: firstNonEmpty(item.speaker_name, item.speakerName, item.display_name, item.displayName, item.personName),
    text: firstNonEmpty(item.text, item.content, item.transcript),
    language: firstNonEmpty(item.language, item.language_code, item.languageCode, options.language),
    source: options.source ?? 'webex_transcript',
    raw: item,
  })).filter((item) => item.text);
}

export function buildPlatformTranscriptImportPayload(input = {}) {
  const platform = String(input.platform ?? input.meeting?.platform ?? '').toLowerCase();
  const raw = input.raw ?? input.transcript ?? input.content ?? input;
  const transcript = input.segments ?? (
    platform.includes('google')
      ? normalizeGoogleMeetTranscriptEntries(raw, input)
      : platform.includes('teams') || platform.includes('microsoft')
        ? normalizeMicrosoftTeamsTranscript(raw, input)
        : platform.includes('zoom')
          ? normalizeZoomTranscript(raw, input)
          : platform.includes('webex')
            ? normalizeWebexTranscript(raw, input)
            : candidateArray(raw, ['segments', 'items', 'transcript'])
  );
  return buildTranscriptImportPayload({
    ...input,
    source: input.source ?? (platform ? `${platform}_transcript` : 'platform_transcript'),
    transcript,
  });
}
