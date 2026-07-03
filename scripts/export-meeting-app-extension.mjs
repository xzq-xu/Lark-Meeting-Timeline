#!/usr/bin/env node

import { execFile } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import {
  assertMeetingAppExtensionScaffold,
  buildMeetingAppExtensionScaffold,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-extension.mjs';

const execFileAsync = promisify(execFile);

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
const buildEnabled = args.get('build') === 'true';
const installEnabled = args.has('install') ? args.get('install') === 'true' : buildEnabled;
const packageManager = String(args.get('package-manager') || process.env.MEETING_APP_EXTENSION_PACKAGE_MANAGER || 'npm');

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function tail(value, max = 1600) {
  const text = String(value ?? '').trim();
  return text.length > max ? text.slice(text.length - max) : text;
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

async function runCommand(command, commandArgs, options = {}) {
  const startedAt = Date.now();
  try {
    const result = await execFileAsync(command, commandArgs, {
      cwd: options.cwd,
      env: process.env,
      maxBuffer: 20 * 1024 * 1024,
    });
    return {
      ok: true,
      command: [command, ...commandArgs].join(' '),
      duration_ms: Date.now() - startedAt,
      stdout_tail: tail(result.stdout),
      stderr_tail: tail(result.stderr),
    };
  } catch (error) {
    return {
      ok: false,
      command: [command, ...commandArgs].join(' '),
      duration_ms: Date.now() - startedAt,
      exit_code: error?.code ?? 1,
      stdout_tail: tail(error?.stdout),
      stderr_tail: tail(error?.stderr),
      error: String(error?.message ?? error),
    };
  }
}

async function validateBundleOutputs(scaffold) {
  const outputs = [
    scaffold.bundle?.content_script_output,
    scaffold.bundle?.live_capture_output,
    scaffold.bundle?.background_output,
  ].filter(Boolean);
  const files = [];
  const missing = [];
  for (const path of outputs) {
    try {
      const content = await readFile(resolve(outDir, path), 'utf8');
      files.push({
        path,
        bytes: Buffer.byteLength(content, 'utf8'),
        contains_live_capture_global: content.includes('__meetingTimelineLiveCapture'),
        contains_client_call_protocol: content.includes('meeting_timeline.client_call'),
      });
    } catch {
      missing.push(path);
    }
  }
  return {
    ok: missing.length === 0 && files.every((file) => file.bytes > 0),
    required_outputs: outputs,
    files,
    missing,
  };
}

async function runBuildGate(scaffold) {
  const install = installEnabled
    ? await runCommand(packageManager, ['install'], { cwd: outDir })
    : { ok: true, skipped: true, command: `${packageManager} install` };
  const build = buildEnabled && install.ok
    ? await runCommand(packageManager, ['run', 'build'], { cwd: outDir })
    : {
      ok: !buildEnabled,
      skipped: true,
      command: `${packageManager} run build`,
      reason: buildEnabled ? 'install_failed' : 'build_not_requested',
    };
  const validation = buildEnabled && build.ok
    ? await validateBundleOutputs(scaffold)
    : {
      ok: !buildEnabled,
      skipped: true,
      reason: buildEnabled ? 'build_failed' : 'build_not_requested',
    };
  return {
    requested: buildEnabled || installEnabled,
    package_manager: packageManager,
    install,
    build,
    validation,
    ok: Boolean(install.ok && build.ok && validation.ok),
  };
}

function buildSummary(scaffold, acceptance, written, buildGate) {
  const ok = Boolean(acceptance.accepted && (!buildGate.requested || buildGate.ok));
  return {
    type: 'meeting_app_extension_export',
    ok,
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
    build_gate: buildGate,
    next_steps: [
      buildEnabled ? `Load unpacked from ${outDir}` : `cd ${outDir} && npm install && npm run build`,
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
  const buildGate = await runBuildGate(scaffold);
  const summary = buildSummary(scaffold, acceptance, written, buildGate);

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
    console.log(`build_gate=${boolLabel(summary.build_gate.ok)} requested=${boolLabel(summary.build_gate.requested)}`);
    console.log(`next=${summary.next_steps[0]}`);
    console.log('load_unpacked=chrome://extensions or edge://extensions');
    console.log('page_probe=window.__meetingTimelineLiveCapture.diagnose()');
  }
  if (!summary.ok) process.exitCode = 2;
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
