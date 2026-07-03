#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  assertMeetingAppExtensionScaffold,
  buildMeetingAppExtensionScaffold,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-extension.mjs';

const args = new Map();
for (const raw of process.argv.slice(2)) {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  args.set(key, rest.length ? rest.join('=') : 'true');
}

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = resolve(String(
  args.get('out-dir')
    || args.get('output-dir')
    || args.get('dir')
    || process.env.MEETING_APP_EXTENSION_OUT_DIR
    || 'data/meeting-app-extension',
));
const baseUrl = String(
  args.get('base-url')
    || args.get('url')
    || process.env.MEETING_TIMELINE_BASE_URL
    || process.env.REAL_DEMO_BASE_URL
    || 'http://localhost:8787',
).replace(/\/$/, '');
const platforms = String(
  args.get('platforms')
    || args.get('platform')
    || process.env.MEETING_APP_EXTENSION_PLATFORMS
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()).filter(Boolean);
const jsonOutput = args.get('json') === 'true';
const reportFile = String(args.get('report-file') || args.get('write-report') || '');

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function localSdkDependencySpecifier() {
  const sdkDir = join(repoRoot, 'packages/meeting-timeline-sdk');
  const rel = relative(outDir, sdkDir).replaceAll('\\', '/');
  return `file:${rel.startsWith('.') ? rel : `./${rel}`}`;
}

function compactAcceptance(report = {}) {
  return {
    accepted: Boolean(report.accepted),
    issue_count: Array.isArray(report.issues) ? report.issues.length : 0,
    platforms: report.platforms ?? [],
    uses_all_urls: Boolean(report.manifest?.uses_all_urls),
  };
}

async function writeScaffoldFiles(scaffold) {
  const written = [];
  for (const file of scaffold.files ?? []) {
    const target = resolve(outDir, file.path);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, file.content, 'utf8');
    written.push({
      path: file.path,
      role: file.role,
      mime: file.mime,
      bytes: Buffer.byteLength(file.content, 'utf8'),
    });
  }
  return written;
}

function buildSummary(scaffold, acceptance, written) {
  return {
    type: 'meeting_app_extension_export',
    ok: Boolean(acceptance.accepted),
    out_dir: outDir,
    base_url: baseUrl,
    platforms: scaffold.install_plan?.platforms ?? [],
    files: written,
    load_unpacked_dir: outDir,
    manifest: {
      matches: scaffold.manifest?.content_scripts?.[0]?.matches ?? [],
      content_scripts: scaffold.manifest?.content_scripts?.[0]?.js ?? [],
      host_permissions: scaffold.manifest?.host_permissions ?? [],
      permissions: scaffold.manifest?.permissions ?? [],
      background: scaffold.manifest?.background ?? null,
    },
    bundle: scaffold.bundle,
    acceptance: compactAcceptance(acceptance),
    next_steps: [
      `cd ${outDir} && npm install && npm run build`,
      `Open chrome://extensions or edge://extensions, enable Developer mode, Load unpacked, then select ${outDir}.`,
      'Open a real meeting page and run window.__meetingTimelineLiveCapture.captureActive(), captureEnded(), evidencePackage(), or diagnose() in DevTools.',
    ],
  };
}

try {
  const scaffold = buildMeetingAppExtensionScaffold({
    platforms,
    baseUrl,
    packageName: String(args.get('package-name') || 'meeting-timeline-real-page-extension'),
    sdkDependencyVersion: String(args.get('sdk-dependency') || args.get('sdk-dependency-version') || localSdkDependencySpecifier()),
    outputScript: String(args.get('content-script') || 'content-script.js'),
    liveCaptureScript: String(args.get('live-capture-script') || 'live-capture.js'),
    backgroundScript: String(args.get('background-script') || 'background.js'),
  });
  const acceptance = assertMeetingAppExtensionScaffold(scaffold);
  const written = await writeScaffoldFiles(scaffold);
  const summary = buildSummary(scaffold, acceptance, written);

  if (reportFile) {
    const target = resolve(reportFile);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, `${JSON.stringify(summary, null, 2)}\n`, 'utf8');
  }

  if (jsonOutput) {
    console.log(JSON.stringify(summary, null, 2));
  } else {
    console.log(`meeting_app_extension_export | ok=${boolLabel(summary.ok)} | dir=${summary.out_dir}`);
    console.log(`platforms=${summary.platforms.join(',')}`);
    console.log(`files=${summary.files.map((file) => file.path).join(',')}`);
    console.log(`next=${summary.next_steps[0]}`);
    console.log('load_unpacked=chrome://extensions or edge://extensions');
    console.log('page_probe=window.__meetingTimelineLiveCapture.diagnose()');
  }
} catch (error) {
  const payload = {
    type: 'meeting_app_extension_export',
    ok: false,
    out_dir: outDir,
    base_url: baseUrl,
    error: String(error?.message ?? error),
  };
  if (jsonOutput) {
    console.error(JSON.stringify(payload, null, 2));
  } else {
    console.error(`meeting_app_extension_export | ok=no | error=${payload.error}`);
  }
  process.exitCode = 2;
}
