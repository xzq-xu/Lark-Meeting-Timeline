import { execFile } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

import { MeetingTimelineApiError, MeetingTimelineSdkError, compactObject } from '../index.mjs';

const execFileAsync = promisify(execFile);

export const DESKTOP_MEETING_SCAN_SCHEMA = 'desktop_meeting_scan';
export const DESKTOP_MEETING_HOST_SCHEMA = 'desktop_meeting_adapter_host';
export const DESKTOP_MEETING_HOST_VERSION = 1;

export const DESKTOP_MEETING_PROFILES = Object.freeze({
  microsoft_teams: Object.freeze({
    platform: 'microsoft_teams',
    display_name: 'Microsoft Teams',
    process_names: Object.freeze(['MSTeams', 'ms-teams', 'Teams', 'Microsoft Teams']),
    application_names: Object.freeze(['Microsoft Teams', 'Microsoft Teams (work or school)', 'Teams']),
    bundle_ids: Object.freeze(['com.microsoft.teams2', 'com.microsoft.teams']),
    leave: /\bleave\b|\bhang up\b|\bend call\b|离开|挂断|结束通话/i,
    microphone: /\bmute\b|\bunmute\b|microphone|麦克风|静音/i,
    camera: /camera|video|摄像头|视频/i,
    participants: /participants|people|attendees|参与者|人员/i,
    share: /share|present|共享|演示/i,
    speaker: /(?:^|[,\s])(.+?)(?:\s+is)?\s+(?:currently\s+)?speaking(?:$|[,\s])|正在发言|当前发言人/i,
  }),
  zoom: Object.freeze({
    platform: 'zoom',
    display_name: 'Zoom',
    process_names: Object.freeze(['zoom.us', 'Zoom', 'Zoom Workplace']),
    application_names: Object.freeze(['zoom.us', 'Zoom', 'Zoom Workplace']),
    bundle_ids: Object.freeze(['us.zoom.xos']),
    leave: /\bleave\b|\bend meeting\b|\bend call\b|离开|结束会议|结束通话/i,
    microphone: /\bmute\b|\bunmute\b|microphone|audio|麦克风|静音|音频/i,
    camera: /camera|video|摄像头|视频/i,
    participants: /participants|attendees|参与者/i,
    share: /share screen|screen share|共享屏幕|共享/i,
    speaker: /(?:^|[,\s])(.+?)(?:\s+is)?\s+(?:currently\s+)?speaking(?:$|[,\s])|正在发言|当前发言人/i,
  }),
});

const PLATFORM_ALIASES = Object.freeze({
  teams: 'microsoft_teams',
  microsoft_teams: 'microsoft_teams',
  'microsoft-teams': 'microsoft_teams',
  zoom: 'zoom',
});

function asArray(value) {
  if (Array.isArray(value)) return value;
  return value == null ? [] : [value];
}

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function normalizedPlatforms(input) {
  const requested = asArray(input ?? ['microsoft_teams', 'zoom'])
    .flatMap((value) => String(value).split(','))
    .map((value) => PLATFORM_ALIASES[String(value).trim().toLowerCase()] ?? String(value).trim().toLowerCase())
    .filter(Boolean);
  const platforms = unique(requested);
  for (const platform of platforms) {
    if (!DESKTOP_MEETING_PROFILES[platform]) {
      throw new MeetingTimelineSdkError('Unsupported desktop meeting platform', {
        platform,
        supported_platforms: Object.keys(DESKTOP_MEETING_PROFILES),
      });
    }
  }
  return platforms;
}

function scannerPath(name) {
  return fileURLToPath(new URL(`./desktop-scanners/${name}`, import.meta.url));
}

function parseScannerOutput(stdout, platform) {
  const text = String(stdout ?? '').trim();
  if (!text) {
    throw new MeetingTimelineSdkError('Desktop scanner returned no data', { platform });
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new MeetingTimelineSdkError('Desktop scanner returned invalid JSON', {
      platform,
      output: text.slice(-2000),
      cause: String(error?.message ?? error),
    });
  }
}

export async function scanDesktopMeetingApps(options = {}) {
  const os = firstNonEmpty(options.os, options.platform, process.platform);
  const timeout = Number(options.timeoutMs ?? options.timeout_ms ?? 5000);
  const commandOptions = {
    timeout,
    maxBuffer: Number(options.maxBuffer ?? options.max_buffer ?? 8 * 1024 * 1024),
    windowsHide: true,
  };
  let result;
  if (os === 'darwin') {
    result = await execFileAsync('osascript', ['-l', 'JavaScript', scannerPath('macos.jxa')], commandOptions);
  } else if (os === 'win32') {
    result = await execFileAsync('powershell.exe', [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      scannerPath('windows.ps1'),
    ], commandOptions);
  } else {
    return {
      type: DESKTOP_MEETING_SCAN_SCHEMA,
      schema: DESKTOP_MEETING_SCAN_SCHEMA,
      schema_version: 1,
      os,
      supported: false,
      observed_at_ms: Date.now(),
      applications: [],
      issues: [{ code: 'unsupported_desktop_os', message: `Desktop scanner does not support ${os}.` }],
    };
  }
  return {
    type: DESKTOP_MEETING_SCAN_SCHEMA,
    schema: DESKTOP_MEETING_SCAN_SCHEMA,
    schema_version: 1,
    os,
    supported: true,
    observed_at_ms: Date.now(),
    ...parseScannerOutput(result.stdout, os),
  };
}

function comparable(value) {
  return String(value ?? '').trim().toLowerCase();
}

function appPlatform(application = {}, platforms = Object.keys(DESKTOP_MEETING_PROFILES)) {
  const values = [
    application.name,
    application.application_name,
    application.process_name,
    application.executable,
  ].map(comparable).filter(Boolean);
  const bundleId = comparable(application.bundle_id ?? application.bundleId);
  return platforms.find((platform) => {
    const profile = DESKTOP_MEETING_PROFILES[platform];
    return profile.bundle_ids.some((value) => comparable(value) === bundleId)
      || [...profile.application_names, ...profile.process_names]
        .some((name) => values.some((value) => value === comparable(name) || value.includes(comparable(name))));
  });
}

function controlText(control = {}) {
  return unique([
    control.name,
    control.title,
    control.label,
    control.description,
    control.help,
    control.value,
    control.identifier,
    control.automation_id,
  ]).join(' ').replace(/\s+/g, ' ').trim();
}

function windowControls(window = {}) {
  return asArray(window.controls ?? window.elements ?? window.accessibility?.controls);
}

function matchingControls(controls, expression) {
  return controls.filter((control) => expression.test(controlText(control)));
}

function explicitBoolean(...values) {
  return values.find((value) => typeof value === 'boolean');
}

function stableHash(value) {
  let hash = 2166136261;
  for (const char of String(value)) {
    hash ^= char.codePointAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36);
}

function speakerFromControls(controls = [], profile = {}) {
  for (const control of controls) {
    const text = controlText(control);
    if (!profile.speaker.test(text)) continue;
    const match = text.match(profile.speaker);
    const name = String(firstNonEmpty(match?.[1], control.name, control.label, control.title, '')).trim();
    const cleaned = name
      .replace(/\b(?:is\s+)?(?:currently\s+)?speaking\b.*$/i, '')
      .replace(/(?:正在发言|当前发言人).*$/i, '')
      .replace(/[,\s]+$/, '')
      .trim();
    if (cleaned && !/^(speaking|currently speaking)$/i.test(cleaned)) {
      return { name: cleaned, source_text: text, confidence: 0.72 };
    }
  }
  return null;
}

function compactControl(control = {}) {
  return compactObject({
    role: control.role ?? control.control_type,
    name: control.name ?? control.label ?? control.title,
    description: control.description,
    identifier: control.identifier ?? control.automation_id,
    value: typeof control.value === 'string' || typeof control.value === 'number' ? control.value : undefined,
    enabled: control.enabled,
    selected: control.selected,
  });
}

function normalizeCandidate(application = {}, window = {}, platform, observedAtMs) {
  const profile = DESKTOP_MEETING_PROFILES[platform];
  const controls = windowControls(window);
  const matched = {
    leave: matchingControls(controls, profile.leave),
    microphone: matchingControls(controls, profile.microphone),
    camera: matchingControls(controls, profile.camera),
    participants: matchingControls(controls, profile.participants),
    share: matchingControls(controls, profile.share),
  };
  const supportingCategories = ['microphone', 'camera', 'participants', 'share']
    .filter((key) => matched[key].length > 0);
  const explicitInMeeting = explicitBoolean(window.in_meeting, window.inMeeting, application.in_meeting, application.inMeeting);
  const inMeeting = explicitInMeeting ?? (matched.leave.length > 0 && supportingCategories.length >= 1);
  const title = firstNonEmpty(window.title, window.name, application.window_title, profile.display_name);
  const candidateKey = `${platform}:${application.pid ?? application.process_id ?? 'unknown'}:${stableHash(title)}`;
  const activeSpeaker = speakerFromControls(controls, profile);
  return compactObject({
    type: 'desktop_meeting_candidate',
    schema: 'desktop_meeting_candidate',
    schema_version: 1,
    platform,
    candidate_key: candidateKey,
    observed_at_ms: observedAtMs,
    in_meeting: Boolean(inMeeting),
    focused: explicitBoolean(window.focused, window.frontmost, application.frontmost) ?? false,
    visible: explicitBoolean(window.visible, application.visible) ?? true,
    title,
    application: {
      name: application.name,
      process_name: application.process_name,
      pid: application.pid ?? application.process_id,
      bundle_id: application.bundle_id ?? application.bundleId,
      frontmost: application.frontmost,
    },
    window: {
      title,
      role: window.role,
      subrole: window.subrole,
      focused: window.focused,
      control_count: controls.length,
    },
    matched_controls: Object.fromEntries(Object.entries(matched).map(([key, rows]) => [key, rows.slice(0, 4).map(compactControl)])),
    active_speaker: activeSpeaker,
    confidence: explicitInMeeting === true ? 0.98 : inMeeting ? Math.min(0.94, 0.72 + supportingCategories.length * 0.055) : 0.2,
  });
}

export function normalizeDesktopMeetingScan(scan = {}, options = {}) {
  const platforms = normalizedPlatforms(options.platforms ?? options.platform ?? ['microsoft_teams', 'zoom']);
  const observedAtMs = Number(scan.observed_at_ms ?? scan.observedAtMs ?? Date.now());
  const candidates = [];
  for (const application of asArray(scan.applications ?? scan.apps ?? scan.processes)) {
    const platform = appPlatform(application, platforms);
    if (!platform) continue;
    const windows = asArray(application.windows ?? application.window ?? []);
    if (windows.length === 0) {
      candidates.push(normalizeCandidate(application, {
        title: application.window_title ?? application.name,
        controls: application.controls ?? [],
        in_meeting: application.in_meeting,
      }, platform, observedAtMs));
      continue;
    }
    for (const window of windows) candidates.push(normalizeCandidate(application, window, platform, observedAtMs));
  }
  const active = candidates
    .filter((candidate) => candidate.in_meeting && candidate.visible !== false)
    .sort((left, right) => Number(right.focused) - Number(left.focused) || right.confidence - left.confidence);
  return {
    type: 'desktop_meeting_scan_result',
    schema: 'desktop_meeting_scan_result',
    schema_version: 1,
    os: scan.os ?? process.platform,
    observed_at_ms: observedAtMs,
    accessibility_trusted: scan.accessibility_trusted,
    scanner_error: scan.error,
    platforms,
    application_count: asArray(scan.applications ?? scan.apps ?? scan.processes).length,
    candidate_count: candidates.length,
    active_candidate_count: active.length,
    candidates,
    selected_candidate: active[0] ?? null,
  };
}

export function buildDesktopMeetingAdapterReadiness(scanOrResult = {}, options = {}) {
  const result = scanOrResult?.schema === 'desktop_meeting_scan_result'
    ? scanOrResult
    : normalizeDesktopMeetingScan(scanOrResult, options);
  const osSupported = ['darwin', 'win32'].includes(result.os);
  const permissionReady = result.accessibility_trusted !== false;
  const issues = [
    !osSupported ? { code: 'unsupported_desktop_os', message: `Desktop adapter does not support ${result.os}.` } : null,
    !permissionReady ? { code: 'accessibility_permission_required', message: 'Grant Accessibility permission to inspect Teams/Zoom meeting controls.' } : null,
    result.scanner_error ? { code: 'desktop_scanner_error', message: String(result.scanner_error) } : null,
  ].filter(Boolean);
  return {
    type: 'desktop_meeting_adapter_readiness',
    schema: 'desktop_meeting_adapter_readiness',
    schema_version: 1,
    ready: issues.length === 0,
    os: result.os,
    os_supported: osSupported,
    accessibility_ready: permissionReady,
    target_application_running: result.application_count > 0,
    active_meeting_detected: result.active_candidate_count > 0,
    application_count: result.application_count,
    candidate_count: result.candidate_count,
    active_candidate_count: result.active_candidate_count,
    issues,
  };
}

function normalizeBaseUrl(value) {
  try {
    const url = new URL(String(value ?? 'http://localhost:8787'));
    if (!['http:', 'https:'].includes(url.protocol)) throw new Error('unsupported protocol');
    return url.toString().replace(/\/$/, '');
  } catch (error) {
    throw new MeetingTimelineSdkError('Desktop meeting adapter requires an HTTP(S) baseUrl', {
      base_url: value,
      cause: String(error?.message ?? error),
    });
  }
}

async function responseBody(response) {
  const text = await response.text();
  if (!text) return null;
  try { return JSON.parse(text); } catch { return text; }
}

function desktopTimelineClient(options = {}) {
  const baseUrl = normalizeBaseUrl(options.baseUrl ?? options.base_url);
  const fetchImpl = options.fetch ?? options.fetchImpl ?? globalThis.fetch;
  if (typeof fetchImpl !== 'function') throw new MeetingTimelineSdkError('Desktop meeting adapter requires fetch');
  const headers = { 'content-type': 'application/json', ...(options.headers ?? {}) };
  async function post(path, payload) {
    if (options.dryRun === true || options.dry_run === true) return { ok: true, dry_run: true, path, payload };
    const response = await fetchImpl(`${baseUrl}${path}`, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
    });
    const body = await responseBody(response);
    if (!response.ok) {
      throw new MeetingTimelineApiError(`Desktop adapter request failed: ${response.status}`, {
        status: response.status,
        path,
        body,
      });
    }
    return body;
  }
  return {
    baseUrl,
    startMeeting: (payload) => post('/api/meeting-session/start', payload),
    endMeeting: (payload) => post('/api/meeting-session/end', payload),
    observeCandidates: (payload) => post('/api/meeting-platform/runtime-events', {
      schema: 'meeting_platform_runtime_event',
      schema_version: 1,
      action: 'observe_platform_candidates',
      source: 'desktop_meeting_adapter',
      sent_at_ms: Date.now(),
      ...payload,
    }),
    insertMark: (payload) => post('/api/annotations', payload),
  };
}

function meetingEvidence(candidate = {}) {
  return {
    schema: 'desktop_meeting_evidence',
    schema_version: 1,
    platform: candidate.platform,
    candidate_key: candidate.candidate_key,
    observed_at_ms: candidate.observed_at_ms,
    title: candidate.title,
    confidence: candidate.confidence,
    application: candidate.application,
    window: candidate.window,
    matched_controls: candidate.matched_controls,
  };
}

export function createDesktopMeetingAdapterHost(options = {}) {
  const platforms = normalizedPlatforms(options.platforms ?? options.platform ?? ['microsoft_teams', 'zoom']);
  const scanner = options.scanner ?? options.scan ?? (() => scanDesktopMeetingApps(options));
  if (typeof scanner !== 'function') throw new MeetingTimelineSdkError('Desktop meeting adapter scanner must be a function');
  const client = options.client ?? desktopTimelineClient(options);
  const startStableSamples = Math.max(1, Number(options.startStableSamples ?? options.start_stable_samples ?? 2));
  const endMissingSamples = Math.max(1, Number(options.endMissingSamples ?? options.end_missing_samples ?? 4));
  const speakerStableSamples = Math.max(1, Number(options.speakerStableSamples ?? options.speaker_stable_samples ?? 2));
  const intervalMs = Math.max(250, Number(options.intervalMs ?? options.interval_ms ?? 750));
  const onEvent = typeof options.onEvent === 'function' ? options.onEvent : typeof options.on_event === 'function' ? options.on_event : () => {};
  let timer = null;
  let running = false;
  let polling = false;
  let activeMeeting = null;
  let pendingCandidate = null;
  let missingSamples = 0;
  let pendingSpeaker = null;
  let currentSpeaker = null;
  let lastScan = null;
  let lastError = null;
  let eventCount = 0;

  async function emit(type, payload = {}) {
    const event = compactObject({
      type,
      schema: 'desktop_meeting_adapter_event',
      schema_version: 1,
      occurred_at_ms: payload.occurred_at_ms ?? Date.now(),
      ...payload,
    });
    eventCount += 1;
    await onEvent(event);
    return event;
  }

  async function startCandidate(candidate) {
    const startTimeMs = pendingCandidate?.first_seen_at_ms ?? candidate.observed_at_ms ?? Date.now();
    const meetingId = `${candidate.platform}-desktop-${startTimeMs}`;
    const payload = {
      platform: candidate.platform,
      meeting_id: meetingId,
      title: candidate.title,
      start_time_ms: startTimeMs,
      detector_source: 'desktop_meeting_adapter',
      observer_surface: 'native_detector',
      suppress_auto_annotations: true,
      meeting_app_record: {
        schema: 'meeting_app_snapshot_record',
        schema_version: 1,
        platform: candidate.platform,
        phase: 'active',
        captured_at_ms: startTimeMs,
        source: 'desktop_meeting_adapter',
        observer_surface: 'native_detector',
        snapshot: meetingEvidence(candidate),
      },
    };
    const response = await client.startMeeting(payload);
    activeMeeting = { meeting_id: meetingId, platform: candidate.platform, candidate_key: candidate.candidate_key, start_time_ms: startTimeMs, title: candidate.title };
    pendingCandidate = null;
    missingSamples = 0;
    return emit('meeting_started', { meeting: activeMeeting, candidate, response, occurred_at_ms: startTimeMs });
  }

  async function endActive(candidate, atMs = Date.now()) {
    if (!activeMeeting) return null;
    const meeting = activeMeeting;
    const payload = {
      meeting_id: meeting.meeting_id,
      end_time_ms: atMs,
      detector_source: 'desktop_meeting_adapter',
      observer_surface: 'native_detector',
      meeting_app_record: {
        schema: 'meeting_app_snapshot_record',
        schema_version: 1,
        platform: meeting.platform,
        phase: 'ended',
        captured_at_ms: atMs,
        source: 'desktop_meeting_adapter',
        observer_surface: 'native_detector',
        snapshot: candidate ? meetingEvidence(candidate) : { platform: meeting.platform, candidate_key: meeting.candidate_key, in_meeting: false },
      },
    };
    const response = await client.endMeeting(payload);
    activeMeeting = null;
    currentSpeaker = null;
    pendingSpeaker = null;
    missingSamples = 0;
    return emit('meeting_ended', { meeting, response, occurred_at_ms: atMs });
  }

  async function observeSpeaker(candidate) {
    const speaker = candidate?.active_speaker;
    if (!speaker?.name || !activeMeeting) {
      pendingSpeaker = null;
      return null;
    }
    if (currentSpeaker?.name === speaker.name) return null;
    if (pendingSpeaker?.name !== speaker.name) {
      pendingSpeaker = { ...speaker, count: 1, first_seen_at_ms: candidate.observed_at_ms ?? Date.now() };
      return null;
    }
    pendingSpeaker.count += 1;
    if (pendingSpeaker.count < speakerStableSamples) return null;
    const occurredAtMs = pendingSpeaker.first_seen_at_ms;
    const mark = {
      id: `speaker-${activeMeeting.meeting_id}-${stableHash(`${speaker.name}:${occurredAtMs}`)}`,
      meeting_id: activeMeeting.meeting_id,
      platform: activeMeeting.platform,
      kind: 'speaker_started',
      intent: 'speaker_track',
      speaker_name: speaker.name,
      label: `Speaker: ${speaker.name}`,
      captured_at_ms: occurredAtMs,
      source: 'desktop_meeting_adapter',
      confidence: speaker.confidence,
      payload: {
        speaker: { name: speaker.name },
        speaker_name: speaker.name,
        detector_source: 'desktop_accessibility',
      },
    };
    const response = await client.insertMark(mark);
    currentSpeaker = { name: speaker.name, occurred_at_ms: occurredAtMs };
    pendingSpeaker = null;
    return emit('speaker_started', { meeting: activeMeeting, speaker: currentSpeaker, mark, response, occurred_at_ms: occurredAtMs });
  }

  async function tick(tickOptions = {}) {
    if (polling) return { skipped: true, reason: 'scan_in_progress' };
    polling = true;
    try {
      const scan = tickOptions.scan ?? await scanner(tickOptions);
      const result = normalizeDesktopMeetingScan(scan, { ...options, ...tickOptions, platforms });
      lastScan = result;
      lastError = null;
      const candidate = result.selected_candidate;
      if (!activeMeeting) {
        if (!candidate) {
          pendingCandidate = null;
          return { scan: result, events: [] };
        }
        if (pendingCandidate?.candidate_key !== candidate.candidate_key) {
          pendingCandidate = { candidate_key: candidate.candidate_key, count: 1, first_seen_at_ms: candidate.observed_at_ms ?? Date.now() };
        } else {
          pendingCandidate.count += 1;
        }
        const events = [];
        if (pendingCandidate.count >= startStableSamples) events.push(await startCandidate(candidate));
        return { scan: result, events: events.filter(Boolean) };
      }

      if (!candidate || candidate.candidate_key !== activeMeeting.candidate_key) {
        missingSamples += 1;
        const events = [];
        if (missingSamples >= endMissingSamples) {
          events.push(await endActive(candidate, result.observed_at_ms));
          if (candidate) pendingCandidate = { candidate_key: candidate.candidate_key, count: 1, first_seen_at_ms: candidate.observed_at_ms ?? Date.now() };
        }
        return { scan: result, events: events.filter(Boolean), missing_samples: missingSamples };
      }

      missingSamples = 0;
      const speakerEvent = await observeSpeaker(candidate);
      return { scan: result, events: speakerEvent ? [speakerEvent] : [] };
    } catch (error) {
      lastError = String(error?.message ?? error);
      await emit('adapter_error', { error: lastError });
      if (tickOptions.throwOnError === true || tickOptions.throw_on_error === true) throw error;
      return { error: lastError, events: [] };
    } finally {
      polling = false;
    }
  }

  function getState() {
    return {
      type: DESKTOP_MEETING_HOST_SCHEMA,
      schema: DESKTOP_MEETING_HOST_SCHEMA,
      schema_version: DESKTOP_MEETING_HOST_VERSION,
      running,
      platforms,
      interval_ms: intervalMs,
      active_meeting: activeMeeting,
      pending_candidate: pendingCandidate,
      missing_samples: missingSamples,
      current_speaker: currentSpeaker,
      last_scan: lastScan,
      last_error: lastError,
      event_count: eventCount,
    };
  }

  async function start() {
    if (running) return getState();
    running = true;
    await tick();
    timer = setInterval(() => { tick().catch(() => {}); }, intervalMs);
    return getState();
  }

  async function stop(stopOptions = {}) {
    running = false;
    if (timer) clearInterval(timer);
    timer = null;
    if (stopOptions.endActive === true || stopOptions.end_active === true) await endActive(lastScan?.selected_candidate, Date.now());
    return getState();
  }

  return {
    type: DESKTOP_MEETING_HOST_SCHEMA,
    platforms,
    client,
    scan: scanner,
    tick,
    start,
    stop,
    getState,
    readiness(scan = lastScan ?? {}) {
      return buildDesktopMeetingAdapterReadiness(scan, { ...options, platforms });
    },
  };
}

export default createDesktopMeetingAdapterHost;
