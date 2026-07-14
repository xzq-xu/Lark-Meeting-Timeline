#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';

const args = new Map(process.argv.slice(2).map((raw) => {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  return [key, rest.length ? rest.join('=') : 'true'];
}));
const debugUrl = String(args.get('debug-url') ?? 'http://127.0.0.1:9222').replace(/\/+$/, '');
const reportFile = args.get('report-file') ? resolve(String(args.get('report-file'))) : null;
const timeoutMs = Math.max(1000, Number(args.get('timeout-ms') ?? 12_000));

const probes = [
  { platform: 'google_meet', url: 'https://meet.google.com/abc-defg-hij' },
  { platform: 'microsoft_teams', url: 'https://teams.microsoft.com/v2/' },
  { platform: 'zoom', url: 'https://zoom.us/wc/join/123456789' },
];

async function targets() {
  const response = await fetch(`${debugUrl}/json/list`);
  if (!response.ok) throw new Error(`Cannot read Chromium targets: ${response.status}`);
  return response.json();
}

async function newTarget(url) {
  const response = await fetch(`${debugUrl}/json/new?${encodeURIComponent(url)}`, { method: 'PUT' });
  if (!response.ok) throw new Error(`Cannot create Chromium target: ${response.status}`);
  return response.json();
}

async function closeTarget(id) {
  await fetch(`${debugUrl}/json/close/${id}`).catch(() => {});
}

function cdp(target) {
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  const listeners = new Map();
  const opened = new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true });
    socket.addEventListener('error', reject, { once: true });
  });
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (message.id && pending.has(message.id)) {
      pending.get(message.id)(message);
      pending.delete(message.id);
      return;
    }
    for (const listener of listeners.get(message.method) ?? []) listener(message.params ?? {});
  });
  return {
    async call(method, params = {}) {
      await opened;
      return new Promise((resolve) => {
        const requestId = ++id;
        pending.set(requestId, resolve);
        socket.send(JSON.stringify({ id: requestId, method, params }));
      });
    },
    on(method, listener) {
      listeners.set(method, [...(listeners.get(method) ?? []), listener]);
    },
    async ready() { await opened; },
    close() { socket.close(); },
  };
}

async function probePlatform(probe) {
  const target = await newTarget('about:blank');
  const client = cdp(target);
  const evidence = [];
  let settle;
  const found = new Promise((resolve) => { settle = resolve; });
  client.on('Runtime.executionContextCreated', async ({ context }) => {
    if (context?.name !== 'Meeting Timeline Adapter') return;
    const evaluated = await client.call('Runtime.evaluate', {
      contextId: context.id,
      expression: `({href:location.href,bridge:Boolean(globalThis.__meetingTimelineBridge),liveCapture:Boolean(globalThis.__meetingTimelineLiveCapture),platform:globalThis.__meetingTimelineBridge?.options?.platform})`,
      returnByValue: true,
    });
    const row = {
      context_name: context.name,
      context_origin: context.origin,
      frame_id: context.auxData?.frameId,
      value: evaluated.result?.result?.value,
    };
    evidence.push(row);
    if (row.value?.platform === probe.platform && row.value?.bridge && row.value?.liveCapture) settle(row);
  });
  await client.call('Runtime.enable');
  await client.call('Page.enable');
  const startedAt = Date.now();
  await client.call('Page.navigate', { url: probe.url });
  const winner = await Promise.race([
    found,
    new Promise((resolve) => setTimeout(() => resolve(null), timeoutMs)),
  ]);
  const current = await client.call('Runtime.evaluate', {
    expression: `({href:location.href,title:document.title})`,
    returnByValue: true,
  }).catch(() => null);
  client.close();
  await closeTarget(target.id);
  return {
    platform: probe.platform,
    requested_url: probe.url,
    accepted: Boolean(winner),
    duration_ms: Date.now() - startedAt,
    evidence,
    final_page: current?.result?.result?.value,
    real_meeting_accepted: false,
  };
}

async function inspectPopup(extensionId) {
  const target = await newTarget(`chrome-extension://${extensionId}/popup.html`);
  const client = cdp(target);
  await client.call('Runtime.enable');
  await new Promise((resolve) => setTimeout(resolve, 750));
  const result = await client.call('Runtime.evaluate', {
    expression: `({title:document.title,service:document.querySelector('#service-status')?.textContent,platform:document.querySelector('#platform')?.textContent,adapter:document.querySelector('#adapter-status')?.textContent,script:document.querySelector('script')?.src,width:document.body.scrollWidth,height:document.body.scrollHeight})`,
    returnByValue: true,
  });
  client.close();
  await closeTarget(target.id);
  const value = result.result?.result?.value;
  return {
    accepted: value?.title === 'Meeting Timeline Adapter' && value?.script?.endsWith('/popup.js'),
    value,
  };
}

async function main() {
  const initialTargets = await targets();
  const initialWorker = initialTargets.find((target) => target.type === 'service_worker' && target.url.endsWith('/background.js'));
  const extensionPage = initialTargets.find((target) => target.url?.startsWith('chrome-extension://') && target.title === 'Meeting Timeline Adapter');
  const extensionId = (initialWorker?.url ?? extensionPage?.url)?.match(/^chrome-extension:\/\/([^/]+)/)?.[1];
  if (!extensionId) throw new Error('Meeting Timeline Adapter is not loaded in the target Chromium instance');
  const popup = await inspectPopup(extensionId);
  const activeTargets = await targets();
  const worker = activeTargets.find((target) => target.type === 'service_worker' && target.url === `chrome-extension://${extensionId}/background.js`);
  const rows = [];
  for (const probe of probes) rows.push(await probePlatform(probe));
  const accepted = Boolean(worker) && popup.accepted && rows.every((row) => row.accepted);
  const report = {
    type: 'three_platform_browser_extension_verification',
    schema: 'three_platform_browser_extension_verification',
    schema_version: 1,
    generated_at: new Date().toISOString(),
    accepted,
    install_start_verified: accepted,
    production_ready: false,
    debug_url: debugUrl,
    extension_id: extensionId,
    worker: { accepted: Boolean(worker), url: worker?.url ?? `chrome-extension://${extensionId}/background.js` },
    popup,
    platform_count: rows.length,
    accepted_platform_count: rows.filter((row) => row.accepted).length,
    rows,
    remaining_gate: 'Join real meetings and capture meeting_started, speaker_started, realtime annotation, and meeting_ended evidence.',
  };
  if (reportFile) {
    await mkdir(dirname(reportFile), { recursive: true });
    await writeFile(reportFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  }
  console.log(JSON.stringify(report, null, 2));
  if (!accepted) process.exitCode = 2;
}

main().catch((error) => {
  console.error(JSON.stringify({
    type: 'three_platform_browser_extension_verification',
    accepted: false,
    debug_url: debugUrl,
    error: String(error?.message ?? error),
  }, null, 2));
  process.exitCode = 2;
});
