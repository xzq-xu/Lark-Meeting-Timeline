#!/usr/bin/env node

import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { existsSync } from 'node:fs';
import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import net from 'node:net';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';
import { hasStrongInCallEvidence } from './three-platform-live-acceptance-core.mjs';
import {
  buildThreePlatformBrowserActionExpression,
  chooseThreePlatformBrowserAutomationAction,
} from './three-platform-live-browser-automation.mjs';
import { verifyThreePlatformLiveAcceptance } from './verify-three-platform-live-acceptance.mjs';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_URLS = Object.freeze({
  google_meet: 'https://meet.google.com/',
  microsoft_teams: 'https://teams.microsoft.com/v2/',
  zoom: 'https://zoom.us/test',
});
const STARTUP_PROBE_URLS = Object.freeze({
  google_meet: 'https://meet.google.com/abc-defg-hij',
  microsoft_teams: 'https://teams.microsoft.com/v2/',
  zoom: 'https://zoom.us/wc/join/123456789',
});
const DISPLAY_NAMES = Object.freeze({
  google_meet: 'Google Meet',
  microsoft_teams: 'Microsoft Teams',
  zoom: 'Zoom',
});

function parseArgs(argv = process.argv.slice(2)) {
  return new Map(argv.map((raw) => {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    return [key, rest.length ? rest.join('=') : 'true'];
  }));
}

function sleep(ms) {
  return new Promise((resolvePromise) => setTimeout(resolvePromise, ms));
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    signal: options.signal ?? AbortSignal.timeout(3_000),
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(body.error ?? `${options.method ?? 'GET'} ${url} failed with ${response.status}`);
  return body;
}

async function postJson(baseUrl, path, body) {
  return fetchJson(`${baseUrl}${path}`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify(body),
  });
}

async function waitForService(baseUrl, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try { return await fetchJson(`${baseUrl}/api/meeting-session/status`); } catch {}
    await sleep(250);
  }
  throw new Error(`Timeline service did not become ready at ${baseUrl}`);
}

async function run(program, args, options = {}) {
  const child = spawn(program, args, {
    cwd: options.cwd ?? repoRoot,
    env: options.env ?? process.env,
    stdio: options.stdio ?? 'inherit',
  });
  const [code] = await once(child, 'exit');
  if (code !== 0) throw new Error(`${program} ${args.join(' ')} exited with ${code}`);
}

async function prepareFakeAudio(args, outputDir, platform) {
  const explicit = args.get('fake-audio-file');
  if (explicit) {
    const file = resolve(String(explicit));
    if (!existsSync(file)) throw new Error(`Fake audio file does not exist: ${file}`);
    return { file, generated: false };
  }
  if (args.get('synthetic-audio') !== 'true' && args.get('speaker-peer') !== 'true') return null;
  if (process.platform !== 'darwin') {
    throw new Error('Automatic synthetic speech currently requires macOS. Pass --fake-audio-file=/absolute/path/to/mono.wav on this OS.');
  }

  const aiffFile = join(outputDir, `${platform}-acceptance-speaker.aiff`);
  const wavFile = join(outputDir, `${platform}-acceptance-speaker.wav`);
  const text = String(args.get('synthetic-audio-text') ?? [
    'Meeting timeline adapter acceptance.',
    'Why is this annotation important?',
    'This sentence creates a stable active speaker segment.',
  ].join(' '));
  const voice = String(args.get('synthetic-audio-voice') ?? 'Samantha');
  await rm(aiffFile, { force: true });
  await rm(wavFile, { force: true });
  try {
    await run('/usr/bin/say', ['-v', voice, '-r', '175', '-o', aiffFile, text], { stdio: 'ignore' });
    await run('/usr/bin/afconvert', [
      aiffFile,
      '-o', wavFile,
      '-f', 'WAVE',
      '-d', 'LEI16@44100',
      '-c', '1',
    ], { stdio: 'ignore' });
  } finally {
    await rm(aiffFile, { force: true });
  }
  if (!existsSync(wavFile)) throw new Error(`Synthetic audio generation did not create ${wavFile}`);
  return { file: wavFile, generated: true, voice, text };
}

async function ensureService(baseUrl, options = {}) {
  try {
    await waitForService(baseUrl, 1_000);
    return { child: null, owned: false };
  } catch {}
  if (options.startServer === false) throw new Error(`Timeline service is not running at ${baseUrl}`);
  const parsed = new URL(baseUrl);
  if (!['localhost', '127.0.0.1'].includes(parsed.hostname)) {
    throw new Error('Automatic server startup is only supported for localhost services');
  }
  const child = spawn(process.execPath, ['src/server.mjs'], {
    cwd: repoRoot,
    env: {
      ...process.env,
      PORT: parsed.port || '8787',
      ...(options.dataDir ? { TIMELINE_DATA_DIR: options.dataDir } : {}),
      ...(options.evidenceDir ? { MEETING_PLATFORM_FIELD_EVIDENCE_DIR: options.evidenceDir } : {}),
      REAL_DEMO_AUTO_ARM: '0',
      REAL_DEMO_AUTO_ANNOTATION: '0',
      REAL_DEMO_DEVICE_SIMULATOR: '0',
      REAL_DEMO_DEVICE_STREAM: '0',
    },
    stdio: ['ignore', 'ignore', 'inherit'],
  });
  await waitForService(baseUrl);
  return { child, owned: true };
}

async function nestedChromeCandidates(root, depth = 0) {
  if (depth > 5) return [];
  let entries = [];
  try { entries = await readdir(root, { withFileTypes: true }); } catch { return []; }
  const rows = [];
  for (const entry of entries) {
    const path = join(root, entry.name);
    if (entry.isDirectory()) rows.push(...await nestedChromeCandidates(path, depth + 1));
    else if (entry.isFile() && entry.name === 'Google Chrome for Testing') rows.push(path);
  }
  return rows;
}

async function chromeExecutable(explicit) {
  const preferred = [
    explicit,
    process.env.CHROME_PATH,
    '/Applications/Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing',
  ].filter(Boolean);
  for (const path of preferred) if (existsSync(path)) return path;
  const home = process.env.HOME;
  if (home) {
    const cached = await nestedChromeCandidates(join(home, 'Library/Caches/ms-playwright'));
    if (cached.length) return cached.sort().at(-1);
  }
  const fallback = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
  if (existsSync(fallback)) return fallback;
  throw new Error('Chrome for Testing was not found. Pass --chrome-path=/absolute/path/to/chrome');
}

async function freePort() {
  const server = net.createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = server.address().port;
  server.close();
  await once(server, 'close');
  return port;
}

async function waitForDebugPort(port, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try { return await fetchJson(`http://127.0.0.1:${port}/json/version`); } catch {}
    await sleep(250);
  }
  throw new Error(`Chromium remote debugging did not become ready on port ${port}`);
}

async function waitForExtensionWorker(port, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const targets = await fetchJson(`http://127.0.0.1:${port}/json/list`).catch(() => []);
    const worker = targets.find((target) => target.type === 'service_worker' && target.url?.endsWith('/background.js'));
    if (worker) return worker;
    await sleep(300);
  }
  throw new Error('Meeting Timeline Adapter service worker was not detected in Chromium');
}

async function withCdpClient(target, callback) {
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  let requestId = 0;
  const pending = new Map();
  const opened = new Promise((resolvePromise, reject) => {
    socket.addEventListener('open', resolvePromise, { once: true });
    socket.addEventListener('error', reject, { once: true });
  });
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (!message.id || !pending.has(message.id)) return;
    const entry = pending.get(message.id);
    pending.delete(message.id);
    if (message.error) entry.reject(new Error(message.error.message ?? 'CDP request failed'));
    else entry.resolve(message);
  });
  async function call(method, params = {}) {
    await opened;
    return new Promise((resolvePromise, reject) => {
      const id = ++requestId;
      pending.set(id, { resolve: resolvePromise, reject });
      socket.send(JSON.stringify({ id, method, params }));
    });
  }
  try {
    return await callback(call);
  } finally {
    socket.close();
  }
}

async function cdpEvaluate(target, expression, options = {}) {
  return withCdpClient(target, async (call) => {
    await call('Runtime.enable');
    let contextId;
    if (options.frameId) {
      const world = await call('Page.createIsolatedWorld', {
        frameId: options.frameId,
        worldName: 'meeting-timeline-acceptance',
        grantUniveralAccess: true,
      });
      contextId = world.result?.executionContextId;
    }
    const response = await call('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
      ...(contextId ? { contextId } : {}),
    });
    if (response.result?.exceptionDetails) {
      throw new Error(response.result.exceptionDetails.text ?? 'CDP evaluation failed');
    }
    return response.result?.result?.value;
  });
}

function flattenFrameTree(tree, rows = []) {
  if (!tree?.frame) return rows;
  rows.push({
    id: tree.frame.id,
    parent_id: tree.frame.parentId,
    url: tree.frame.url,
    name: tree.frame.name,
  });
  for (const child of tree.childFrames ?? []) flattenFrameTree(child, rows);
  return rows;
}

async function cdpPageFrames(target) {
  return withCdpClient(target, async (call) => {
    const response = await call('Page.getFrameTree');
    return flattenFrameTree(response.result?.frameTree);
  });
}

async function cdpBringToFront(target) {
  return withCdpClient(target, async (call) => {
    await call('Page.bringToFront');
  });
}

const BROWSER_PAGE_SNAPSHOT_EXPRESSION = [
  '(() => {',
  '  const isVisible = (node) => {',
  '    const rect = node.getBoundingClientRect();',
  '    const style = getComputedStyle(node);',
  "    return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none';",
  '  };',
  "  const controls = Array.from(document.querySelectorAll('button, a, input')).filter(isVisible).slice(0, 400).map((node) => ({",
  '    tag: node.tagName.toLowerCase(),',
  '    id: node.id || undefined,',
  "    text: (node.innerText || '').replace(/\\s+/g, ' ').trim().slice(0, 160) || undefined,",
  "    aria: node.getAttribute('aria-label') || undefined,",
  "    title: node.getAttribute('title') || undefined,",
  "    placeholder: node.getAttribute('placeholder') || undefined,",
  "    type: node.getAttribute('type') || undefined,",
  '    value: node.value || undefined,',
  '    value_present: Boolean(node.value),',
  '    href: node.href || undefined,',
  '    disabled: Boolean(node.disabled),',
  '    visible: true,',
  "    in_dialog: Boolean(node.closest('[role=\"dialog\"], dialog')),",
  '  }));',
  '  return { title: document.title, url: location.href, controls };',
  '})()',
].join('\n');

async function browserPageSnapshots(port) {
  const targets = await fetchJson('http://127.0.0.1:' + port + '/json/list').catch(() => []);
  const pages = [];
  for (const target of targets.filter((item) => item.type === 'page')) {
    const frames = await cdpPageFrames(target).catch(() => [{ id: target.id, url: target.url }]);
    for (const frame of frames) {
      const snapshot = await cdpEvaluate(target, BROWSER_PAGE_SNAPSHOT_EXPRESSION, {
        frameId: frame.id,
      }).catch(() => null);
      if (!snapshot) continue;
      pages.push({
        ...snapshot,
        target_id: target.id,
        frame_id: frame.id,
        parent_frame_id: frame.parent_id,
        top_level_url: target.url,
      });
    }
  }
  return pages;
}

async function performBrowserAutomationAction(port, action) {
  const targets = await fetchJson(`http://127.0.0.1:${port}/json/list`).catch(() => []);
  const target = targets.find((item) => item.type === 'page' && item.id === action.target_id);
  if (!target) return { ok: false, reason: 'page_target_missing', phase: action.phase };
  await cdpBringToFront(target);
  return cdpEvaluate(target, buildThreePlatformBrowserActionExpression(action), { frameId: action.frame_id });
}

async function waitForExtensionAttachment(port, platform, sinceMs, timeoutMs = 20_000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const targets = await fetchJson(`http://127.0.0.1:${port}/json/list`).catch(() => []);
    const worker = targets.find((target) => target.type === 'service_worker' && target.url?.endsWith('/background.js'));
    if (worker) {
      const attached = await cdpEvaluate(worker, `new Promise((resolve) => chrome.storage.local.get("meeting_timeline_extension_status", (value) => resolve(value.meeting_timeline_extension_status)))`).catch(() => null);
      if (
        attached?.platform === platform
        && Number(attached.captured_at_ms ?? 0) >= sinceMs - 10_000
      ) return attached;
    }
    await sleep(300);
  }
  throw new Error(`Meeting Timeline content script did not attach to a ${platform} page`);
}

async function configureExtensionBaseUrl(worker, baseUrl) {
  const expression = [
    'new Promise((resolve) => {',
    `  const baseUrl = ${JSON.stringify(baseUrl)};`,
    '  chrome.storage.local.set({ meeting_timeline_base_url: baseUrl }, () => resolve({',
    '    ok: !chrome.runtime.lastError,',
    '    base_url: baseUrl,',
    '    error: chrome.runtime.lastError?.message,',
    '  }));',
    '})',
  ].join('\n');
  const configured = await cdpEvaluate(worker, expression);
  if (configured?.ok !== true || configured.base_url !== baseUrl) {
    throw new Error(`Meeting Timeline extension base URL configuration failed: ${JSON.stringify(configured)}`);
  }
  return configured;
}

async function jsonFiles(root) {
  let entries = [];
  try { entries = await readdir(root, { withFileTypes: true }); } catch { return []; }
  const files = [];
  for (const entry of entries) {
    const path = resolve(root, entry.name);
    if (entry.isDirectory()) files.push(...await jsonFiles(path));
    else if (entry.isFile() && entry.name.endsWith('.json')) files.push(path);
  }
  return files;
}

async function evidenceForMeeting(evidenceDir, platform, meetingId, sinceMs) {
  for (const file of await jsonFiles(evidenceDir)) {
    try {
      const input = JSON.parse(await readFile(file, 'utf8'));
      if (normalizeMeetingPlatform(input.platform) !== platform) continue;
      if (String(input.meeting_id ?? input.meeting?.meeting_id) !== String(meetingId)) continue;
      const records = (input.meetingAppRecords ?? []).filter((record) => Number(record.captured_at_ms ?? 0) >= sinceMs - 10_000);
      if (!records.length) continue;
      return { file, input, records };
    } catch {}
  }
  return null;
}

function activeRecord(evidence = {}) {
  return evidence.records?.find((record) => (
    String(record.phase).toLowerCase() === 'active'
    && hasStrongInCallEvidence(record.snapshot)
  )) ?? null;
}

async function endStaleAxis(baseUrl, status, enabled) {
  const meeting = status.current_meeting;
  if (!meeting?.start_time || meeting.end_time) return;
  if (!enabled) {
    throw new Error(`An active timeline already exists for ${meeting.platform}:${meeting.meeting_id}. End it first or pass --end-stale=true.`);
  }
  await postJson(baseUrl, '/api/meeting-session/end', {
    meeting_id: meeting.meeting_id,
    end_time_ms: Date.now(),
    detector_source: 'three_platform_live_acceptance_stale_cleanup',
  });
}

async function terminateChild(child, timeoutMs = 2_000) {
  if (!child || child.exitCode != null || child.signalCode != null) return;
  child.kill('SIGTERM');
  await Promise.race([
    once(child, 'exit').catch(() => {}),
    sleep(timeoutMs),
  ]);
  if (child.exitCode == null && child.signalCode == null) {
    child.kill('SIGKILL');
    await once(child, 'exit').catch(() => {});
  }
}

async function main() {
  const args = parseArgs();
  const platform = normalizeMeetingPlatform(args.get('platform') ?? 'google-meet');
  if (!Object.hasOwn(DEFAULT_URLS, platform)) throw new Error(`Unsupported live acceptance platform: ${platform}`);
  const baseUrl = String(args.get('base-url') ?? 'http://localhost:8787').replace(/\/+$/, '');
  const outputDir = resolve(String(args.get('out-dir') ?? 'data/three-platform-live-acceptance'));
  const releaseDir = resolve(String(args.get('release-dir') ?? 'data/three-platform-adapters'));
  const extensionDir = resolve(String(args.get('extension-dir') ?? join(releaseDir, 'browser-extension')));
  const evidenceDir = resolve(String(args.get('evidence-dir') ?? 'data/meeting-platform-field-evidence'));
  const profileDir = resolve(String(args.get('profile-dir') ?? join(outputDir, 'browser-profiles', platform)));
  const meetingUrl = String(args.get('meeting-url') ?? DEFAULT_URLS[platform]);
  const startupProbeUrl = String(args.get('probe-url') ?? STARTUP_PROBE_URLS[platform]);
  const timeoutMs = Math.max(60_000, Number(args.get('timeout-ms') ?? 30 * 60_000));
  const autoJoin = args.get('auto-join') === 'true';
  const autoJoinAudio = args.get('auto-join-audio') !== 'false';
  const autoLeave = args.get('auto-leave') === 'true';
  const speakerPeer = args.get('speaker-peer') === 'true';
  const requireSpeaker = args.has('require-speaker')
    ? args.get('require-speaker') === 'true'
    : speakerPeer || args.get('synthetic-audio') === 'true';
  const automationPreview = args.get('preview-browser-automation') === 'true';
  const allowExternalActions = args.get('allow-external-actions') === 'true';
  const isolatedState = args.get('isolated-state') === 'true' || args.has('timeline-data-dir');
  if ((autoJoin || autoLeave) && !allowExternalActions) {
    throw new Error('Browser meeting automation requires --allow-external-actions=true because joining or leaving changes external meeting state.');
  }
  if (speakerPeer && !autoJoin) {
    throw new Error('The synthetic speaker peer requires --auto-join=true.');
  }
  if (speakerPeer && platform === 'zoom' && /\/test\/?(?:[?#]|$)/i.test(meetingUrl)) {
    throw new Error('The synthetic speaker peer requires a shared Zoom meeting URL, not zoom.us/test.');
  }
  const launchedAtMs = Date.now();
  const serviceDataDir = isolatedState
    ? resolve(String(args.get('timeline-data-dir') ?? join(outputDir, 'service-state', `${platform}-${launchedAtMs}`)))
    : null;
  const speakerPeerProfileDir = resolve(String(
    args.get('speaker-peer-profile-dir') ?? `${profileDir}-speaker-peer`,
  ));
  if (args.get('fresh-profile') === 'true') {
    await Promise.all([
      rm(profileDir, { recursive: true, force: true }),
      rm(speakerPeerProfileDir, { recursive: true, force: true }),
    ]);
  }
  await Promise.all([
    mkdir(outputDir, { recursive: true }),
    mkdir(profileDir, { recursive: true }),
    ...(serviceDataDir ? [mkdir(serviceDataDir, { recursive: true })] : []),
    ...(speakerPeer ? [mkdir(speakerPeerProfileDir, { recursive: true })] : []),
  ]);
  const fakeAudio = await prepareFakeAudio(args, outputDir, platform);
  if (speakerPeer && !fakeAudio) throw new Error('The synthetic speaker peer requires a fake audio source.');

  if (args.get('skip-build') !== 'true') {
    await run(process.execPath, [
      'scripts/export-three-platform-adapters.mjs',
      `--out-dir=${releaseDir}`,
      `--base-url=${baseUrl}`,
      '--offline=true',
    ]);
  }
  const manifest = JSON.parse(await readFile(join(extensionDir, 'manifest.json'), 'utf8'));
  if (!manifest.content_scripts?.length) throw new Error('The generated browser extension has no content script');

  const service = await ensureService(baseUrl, {
    startServer: args.get('start-server') !== 'false',
    dataDir: serviceDataDir,
    evidenceDir,
  });
  let browser = null;
  let speakerPeerBrowser = null;
  let finishPromise = null;
  const finish = () => {
    if (finishPromise) return finishPromise;
    finishPromise = (async () => {
      await Promise.all([
        terminateChild(browser),
        terminateChild(speakerPeerBrowser),
      ]);
      if (service.owned) await terminateChild(service.child);
    })();
    return finishPromise;
  };
  process.once('SIGINT', () => { finish().finally(() => process.exit(130)); });
  process.once('SIGTERM', () => { finish().finally(() => process.exit(143)); });

  try {
    const initialStatus = await waitForService(baseUrl);
    await endStaleAxis(baseUrl, initialStatus, args.get('end-stale') === 'true');
    const chromePath = await chromeExecutable(args.get('chrome-path'));
    const debugPort = args.get('debug-port') ? Number(args.get('debug-port')) : await freePort();
    const speakerPeerDebugPort = speakerPeer
      ? (args.get('speaker-peer-debug-port') ? Number(args.get('speaker-peer-debug-port')) : await freePort())
      : null;
    const chromeArgs = [
      `--user-data-dir=${profileDir}`,
      `--remote-debugging-port=${debugPort}`,
      `--disable-extensions-except=${extensionDir}`,
      `--load-extension=${extensionDir}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-background-timer-throttling',
      '--disable-renderer-backgrounding',
      '--autoplay-policy=no-user-gesture-required',
      '--use-fake-ui-for-media-stream',
      ...(fakeAudio && !speakerPeer ? [
        '--use-fake-device-for-media-stream',
        `--use-file-for-fake-audio-capture=${fakeAudio.file}`,
      ] : []),
      '--window-size=1440,1000',
      baseUrl,
      startupProbeUrl,
      meetingUrl,
    ];
    browser = spawn(chromePath, chromeArgs, { stdio: ['ignore', 'ignore', 'ignore'] });
    await waitForDebugPort(debugPort);
    if (speakerPeer) {
      const speakerPeerChromeArgs = [
        `--user-data-dir=${speakerPeerProfileDir}`,
        `--remote-debugging-port=${speakerPeerDebugPort}`,
        '--disable-extensions',
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-background-timer-throttling',
        '--disable-renderer-backgrounding',
        '--autoplay-policy=no-user-gesture-required',
        '--use-fake-ui-for-media-stream',
        '--use-fake-device-for-media-stream',
        `--use-file-for-fake-audio-capture=${fakeAudio.file}`,
        '--window-size=1200,900',
        meetingUrl,
      ];
      speakerPeerBrowser = spawn(chromePath, speakerPeerChromeArgs, { stdio: ['ignore', 'ignore', 'ignore'] });
      await waitForDebugPort(speakerPeerDebugPort);
    }
    const worker = await waitForExtensionWorker(debugPort);
    const extensionConfiguration = await configureExtensionBaseUrl(worker, baseUrl);
    const attached = await waitForExtensionAttachment(debugPort, platform, launchedAtMs);
    console.log(`LIVE_ACCEPTANCE_READY platform=${platform} browser=${chromePath}`);
    console.log(speakerPeer
      ? `The SDK observer and synthetic speaker peer will join the shared ${DISPLAY_NAMES[platform]} meeting automatically.`
      : requireSpeaker && fakeAudio
        ? `Open or join a real ${DISPLAY_NAMES[platform]} meeting in the launched browser, keep the microphone unmuted for the synthetic speaker stimulus, then end or leave the meeting.`
        : requireSpeaker
          ? `Open or join a real ${DISPLAY_NAMES[platform]} meeting in the launched browser, speak for at least two seconds, then end or leave the meeting.`
          : `Open or join a real ${DISPLAY_NAMES[platform]} meeting in the launched browser; after the acceptance annotation appears, end or leave the meeting.`);
    console.log(`The acceptance annotation will be inserted automatically after real in-call evidence appears. Extension worker=${worker.url} attached=${attached.url}`);
    if (fakeAudio) console.log(`Synthetic speaker stimulus is enabled with ${fakeAudio.file}${speakerPeer ? ' on the peer browser' : ''}`);
    if (args.get('startup-only') === 'true') {
      const previewPages = automationPreview ? await browserPageSnapshots(debugPort) : [];
      const previewAction = automationPreview
        ? chooseThreePlatformBrowserAutomationAction({ platform, pages: previewPages, autoJoin: true, meetingUrl })
        : null;
      const startupReport = {
        type: 'three_platform_live_acceptance_startup',
        schema: 'three_platform_live_acceptance_startup',
        schema_version: 1,
        generated_at: new Date().toISOString(),
        accepted: true,
        platform,
        base_url: baseUrl,
        meeting_url: meetingUrl,
        probe_url: startupProbeUrl,
        extension_dir: extensionDir,
        browser_profile_dir: profileDir,
        service_data_dir: serviceDataDir,
        debug_port: debugPort,
        extension_worker_url: worker.url,
        extension_configuration: extensionConfiguration,
        extension_attachment: attached,
        fake_audio_file: fakeAudio?.file ?? null,
        synthetic_audio_generated: fakeAudio?.generated === true,
        fake_audio_target: fakeAudio ? (speakerPeer ? 'speaker_peer' : 'observer') : null,
        acceptance_profile: requireSpeaker ? 'core_and_speaker' : 'core',
        speaker_acceptance_required: requireSpeaker,
        speaker_peer: {
          enabled: speakerPeer,
          browser_profile_dir: speakerPeer ? speakerPeerProfileDir : null,
          debug_port: speakerPeerDebugPort,
        },
        browser_automation: {
          auto_join: autoJoin,
          auto_join_audio: autoJoinAudio,
          auto_leave: autoLeave,
          allow_external_actions: allowExternalActions,
          preview_only: automationPreview,
          preview_action: previewAction,
          pages: previewPages.map((page) => ({
            target_id: page.target_id,
            frame_id: page.frame_id,
            title: page.title,
            url: page.url,
            visible_control_count: page.controls?.length ?? 0,
          })),
        },
      };
      const startupReportFile = join(outputDir, `${platform}-startup.json`);
      await writeFile(startupReportFile, `${JSON.stringify(startupReport, null, 2)}\n`, 'utf8');
      console.log(`LIVE_ACCEPTANCE_STARTUP_COMPLETE report=${startupReportFile}`);
      return;
    }

    const deadline = launchedAtMs + timeoutMs;
    let meeting = null;
    let evidence = null;
    let annotation = null;
    let speakerSeen = false;
    let completionSeenAtMs = null;
    const completedAutomationActions = new Set();
    const browserAutomationEvents = [];
    const speakerPeerCompletedActions = new Set();
    const speakerPeerAutomationEvents = [];
    while (Date.now() < deadline) {
      const status = await fetchJson(`${baseUrl}/api/meeting-session/status`).catch(() => null);
      const current = status?.current_meeting;
      let currentPlatform = null;
      try { currentPlatform = normalizeMeetingPlatform(current?.platform); } catch {}
      const currentStartMs = Date.parse(current?.start_time ?? '');
      if (
        currentPlatform === platform
        && !current?.end_time
        && Number.isFinite(currentStartMs)
        && currentStartMs >= launchedAtMs - 10_000
      ) {
        meeting = current;
        evidence = await evidenceForMeeting(evidenceDir, platform, current.meeting_id, launchedAtMs);
        if (evidence && activeRecord(evidence) && !annotation) {
          const capturedAtMs = Date.now();
          const id = `three-platform-live-acceptance-${platform}-${capturedAtMs}`;
          annotation = await postJson(baseUrl, '/api/annotations', {
            id,
            meeting_id: current.meeting_id,
            source: 'three_platform_live_acceptance',
            captured_at_ms: capturedAtMs,
            kind: 'handwriting_trigger',
            intent: 'question',
            label: 'why?',
            payload: { acceptance: 'three_platform_live', platform, meeting_id: current.meeting_id },
          });
          if (annotation.ack?.accepted !== true) throw new Error(`Acceptance annotation was rejected: ${JSON.stringify(annotation.ack ?? annotation)}`);
          if (!requireSpeaker) completionSeenAtMs = Date.now();
          console.log(`REAL_MEETING_DETECTED meeting=${current.meeting_id} annotation=${id}`);
        }
        if (evidence) {
          const nextSpeakerSeen = (evidence.input.speaker_markers ?? []).some((row) => (
            row.kind === 'speaker_started' || row.intent === 'speaker_track'
          ));
          if (nextSpeakerSeen && !speakerSeen) console.log(`SPEAKER_EVIDENCE_DETECTED meeting=${current.meeting_id}`);
          if (requireSpeaker && nextSpeakerSeen && !completionSeenAtMs) completionSeenAtMs = Date.now();
          speakerSeen = nextSpeakerSeen;
        }
      }

      if (autoJoin || autoLeave) {
        const pages = await browserPageSnapshots(debugPort);
        const action = chooseThreePlatformBrowserAutomationAction({
          platform,
          pages,
          meetingUrl,
          autoJoin,
          autoJoinAudio,
          autoLeave,
          activeMeeting: Boolean(meeting && current?.meeting_id === meeting.meeting_id && !current.end_time),
          completionSeen: requireSpeaker ? speakerSeen : Boolean(annotation),
          leaveReady: completionSeenAtMs != null && Date.now() - completionSeenAtMs >= 2_500,
          displayName: speakerPeer ? 'Timeline Adapter Observer' : 'Timeline Adapter Acceptance',
          completedActionKeys: [...completedAutomationActions],
        });
        if (action) {
          const actionAtMs = Date.now();
          const result = await performBrowserAutomationAction(debugPort, action).catch((error) => ({
            ok: false,
            reason: String(error?.message ?? error),
            phase: action.phase,
          }));
          browserAutomationEvents.push({
            captured_at_ms: actionAtMs,
            action,
            result,
          });
          if (result?.ok === true) {
            completedAutomationActions.add(action.key);
            if (action.global_key) completedAutomationActions.add(action.global_key);
            console.log(`BROWSER_AUTOMATION_ACTION phase=${action.phase} target=${action.target_id}`);
          }
        }
      }

      if (speakerPeer) {
        const peerPages = await browserPageSnapshots(speakerPeerDebugPort);
        const peerAction = chooseThreePlatformBrowserAutomationAction({
          platform,
          pages: peerPages,
          meetingUrl,
          autoJoin: true,
          autoJoinAudio: true,
          autoLeave: false,
          activeMeeting: Boolean(meeting && current?.meeting_id === meeting.meeting_id && !current.end_time),
          displayName: 'Timeline Synthetic Speaker',
          completedActionKeys: [...speakerPeerCompletedActions],
        });
        if (peerAction) {
          const actionAtMs = Date.now();
          const result = await performBrowserAutomationAction(speakerPeerDebugPort, peerAction).catch((error) => ({
            ok: false,
            reason: String(error?.message ?? error),
            phase: peerAction.phase,
          }));
          speakerPeerAutomationEvents.push({
            captured_at_ms: actionAtMs,
            action: peerAction,
            result,
          });
          if (result?.ok === true) {
            speakerPeerCompletedActions.add(peerAction.key);
            if (peerAction.global_key) speakerPeerCompletedActions.add(peerAction.global_key);
            console.log(`SPEAKER_PEER_AUTOMATION_ACTION phase=${peerAction.phase} target=${peerAction.target_id}`);
          }
        }
      }

      if (meeting && current?.meeting_id === meeting.meeting_id && current.end_time) {
        const report = await verifyThreePlatformLiveAcceptance({
          inputDir: evidenceDir,
          platforms: [platform],
          sinceMs: launchedAtMs,
          requireSpeaker,
        });
        const reportFile = join(outputDir, `${platform}-${launchedAtMs}.json`);
        await writeFile(reportFile, `${JSON.stringify({
          ...report,
          launched_at_ms: launchedAtMs,
          meeting_url: meetingUrl,
          browser_profile_dir: profileDir,
          service_data_dir: serviceDataDir,
          extension_dir: extensionDir,
          extension_configuration: extensionConfiguration,
          fake_audio_file: fakeAudio?.file ?? null,
          synthetic_audio_generated: fakeAudio?.generated === true,
          fake_audio_target: fakeAudio ? (speakerPeer ? 'speaker_peer' : 'observer') : null,
          acceptance_profile: requireSpeaker ? 'core_and_speaker' : 'core',
          speaker_acceptance_required: requireSpeaker,
          speaker_peer: {
            enabled: speakerPeer,
            browser_profile_dir: speakerPeer ? speakerPeerProfileDir : null,
            debug_port: speakerPeerDebugPort,
            events: speakerPeerAutomationEvents,
          },
          browser_automation: {
            auto_join: autoJoin,
            auto_join_audio: autoJoinAudio,
            auto_leave: autoLeave,
            allow_external_actions: allowExternalActions,
            events: browserAutomationEvents,
          },
        }, null, 2)}\n`, 'utf8');
        console.log(`LIVE_ACCEPTANCE_COMPLETE accepted=${report.accepted ? 'yes' : 'no'} report=${reportFile}`);
        if (!report.accepted) process.exitCode = 2;
        return;
      }
      await sleep(500);
    }
    const timeoutReportFile = join(outputDir, `${platform}-${launchedAtMs}-timeout.json`);
    await writeFile(timeoutReportFile, `${JSON.stringify({
      type: 'three_platform_live_acceptance_timeout',
      schema: 'three_platform_live_acceptance_timeout',
      schema_version: 1,
      generated_at: new Date().toISOString(),
      platform,
      launched_at_ms: launchedAtMs,
      meeting_url: meetingUrl,
      meeting,
      speaker_seen: speakerSeen,
      acceptance_profile: requireSpeaker ? 'core_and_speaker' : 'core',
      speaker_acceptance_required: requireSpeaker,
      annotation_inserted: Boolean(annotation),
      fake_audio_target: fakeAudio ? (speakerPeer ? 'speaker_peer' : 'observer') : null,
      speaker_peer: {
        enabled: speakerPeer,
        browser_profile_dir: speakerPeer ? speakerPeerProfileDir : null,
        debug_port: speakerPeerDebugPort,
        events: speakerPeerAutomationEvents,
        pages: speakerPeer ? await browserPageSnapshots(speakerPeerDebugPort).catch(() => []) : [],
      },
      browser_automation: {
        auto_join: autoJoin,
        auto_join_audio: autoJoinAudio,
        auto_leave: autoLeave,
        allow_external_actions: allowExternalActions,
        events: browserAutomationEvents,
      },
      pages: await browserPageSnapshots(debugPort).catch(() => []),
    }, null, 2)}\n`, 'utf8');
    throw new Error(`Live acceptance timed out after ${timeoutMs} ms`);
  } finally {
    await finish();
  }
}

main().catch((error) => {
  console.error(`LIVE_ACCEPTANCE_FAILED ${String(error?.message ?? error)}`);
  process.exitCode = 2;
});
