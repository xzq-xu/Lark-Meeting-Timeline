import { appendFile, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

import {
  buildDesktopMeetingAdapterReadiness,
  createDesktopMeetingAdapterHost,
  normalizeDesktopMeetingScan,
  scanDesktopMeetingApps,
} from '../adapters/desktop-meeting-host.mjs';

function parseArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

function booleanArg(args, key, fallback = false) {
  if (!args.has(key)) return fallback;
  return !['false', '0', 'no'].includes(String(args.get(key)).toLowerCase());
}

function numberArg(args, key, fallback) {
  const value = Number(args.get(key));
  return Number.isFinite(value) ? value : fallback;
}

async function serviceStatus(baseUrl) {
  try {
    const response = await fetch(`${baseUrl}/api/state`, { headers: { accept: 'application/json' } });
    return { ok: response.ok, status: response.status, base_url: baseUrl };
  } catch (error) {
    return { ok: false, base_url: baseUrl, error: String(error?.message ?? error) };
  }
}

async function evidenceWriter(pathValue) {
  if (!pathValue) return async () => {};
  const path = resolve(String(pathValue));
  await mkdir(dirname(path), { recursive: true });
  return async (event) => appendFile(path, `${JSON.stringify(event)}\n`, 'utf8');
}

export async function runDesktopMeetingAdapterCli(argv = process.argv.slice(2), io = {}) {
  const args = parseArgs(argv);
  const stdout = io.stdout ?? process.stdout;
  const stderr = io.stderr ?? process.stderr;
  const writeJson = (value, stream = stdout) => stream.write(`${JSON.stringify(value, null, 2)}\n`);
  const baseUrl = String(args.get('base-url') ?? process.env.MEETING_TIMELINE_BASE_URL ?? 'http://localhost:8787').replace(/\/+$/, '');
  const platforms = String(args.get('platforms') ?? args.get('platform') ?? 'teams,zoom').split(',').map((value) => value.trim()).filter(Boolean);
  const once = booleanArg(args, 'once') || booleanArg(args, 'diagnose') || booleanArg(args, 'install-check');
  const jsonLines = booleanArg(args, 'json-lines', true);
  const writeEvidence = await evidenceWriter(args.get('evidence-file'));

  if (once) {
    const scan = await scanDesktopMeetingApps({
      os: args.get('os'),
      timeoutMs: numberArg(args, 'scan-timeout-ms', 5000),
    });
    const result = normalizeDesktopMeetingScan(scan, { platforms });
    const readiness = buildDesktopMeetingAdapterReadiness(result, { platforms });
    const service = await serviceStatus(baseUrl);
    const report = {
      type: 'desktop_meeting_adapter_diagnosis',
      schema: 'desktop_meeting_adapter_diagnosis',
      schema_version: 1,
      ok: readiness.ready && service.ok,
      generated_at: new Date().toISOString(),
      base_url: baseUrl,
      platforms: result.platforms,
      service,
      readiness,
      scan: result,
      next_actions: [
        !service.ok ? `Start Meeting Timeline at ${baseUrl}.` : null,
        readiness.accessibility_ready === false && result.os === 'darwin'
          ? 'Open System Settings > Privacy & Security > Accessibility and allow the terminal or packaged desktop adapter.'
          : null,
        readiness.accessibility_ready === false && result.os === 'win32'
          ? 'Run the desktop adapter in the same user session and privilege level as Teams/Zoom.'
          : null,
        !readiness.target_application_running ? 'Open Microsoft Teams or Zoom to continue live validation.' : null,
        readiness.target_application_running && !readiness.active_meeting_detected ? 'Join a meeting and rerun the diagnosis.' : null,
      ].filter(Boolean),
    };
    writeJson(report);
    return report;
  }

  const host = createDesktopMeetingAdapterHost({
    baseUrl,
    platforms,
    intervalMs: numberArg(args, 'interval-ms', 750),
    startStableSamples: numberArg(args, 'start-stable-samples', 2),
    endMissingSamples: numberArg(args, 'end-missing-samples', 4),
    speakerStableSamples: numberArg(args, 'speaker-stable-samples', 2),
    dryRun: booleanArg(args, 'dry-run'),
    async onEvent(event) {
      await writeEvidence(event);
      if (jsonLines) stdout.write(`${JSON.stringify(event)}\n`);
    },
  });
  await host.start();
  const started = {
    type: 'desktop_meeting_adapter_started',
    schema: 'desktop_meeting_adapter_started',
    schema_version: 1,
    pid: process.pid,
    base_url: baseUrl,
    platforms: host.platforms,
    interval_ms: host.getState().interval_ms,
  };
  writeJson(started, jsonLines ? stderr : stdout);

  const stop = async (signal) => {
    const state = await host.stop({ endActive: booleanArg(args, 'end-on-exit') });
    writeJson({ type: 'desktop_meeting_adapter_stopped', signal, state }, jsonLines ? stderr : stdout);
    process.exitCode = 0;
  };
  process.once('SIGINT', () => { stop('SIGINT').finally(() => process.exit()); });
  process.once('SIGTERM', () => { stop('SIGTERM').finally(() => process.exit()); });
  return { host, started };
}

export default runDesktopMeetingAdapterCli;
