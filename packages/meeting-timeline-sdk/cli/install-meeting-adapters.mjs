import { existsSync } from 'node:fs';
import { cp, mkdir, readFile, rm } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const bundledExtensionDir = resolve(packageRoot, 'runtime', 'browser-extension');
const bundledReleaseStatusFile = resolve(packageRoot, 'runtime', 'three-platform-release-status.json');

function parseArgs(argv = process.argv.slice(2)) {
  return new Map(argv.map((raw) => {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    return [key, rest.length ? rest.join('=') : 'true'];
  }));
}

export async function exportBundledMeetingAdapters(options = {}) {
  if (!existsSync(resolve(bundledExtensionDir, 'manifest.json'))) {
    throw new Error('This SDK package does not contain the prebuilt meeting adapter runtime. Install the three-platform release tarball.');
  }
  const outDir = resolve(String(options.outDir ?? 'meeting-timeline-browser-extension'));
  if (existsSync(outDir)) {
    if (options.force !== true) throw new Error(`Output directory already exists: ${outDir}. Pass --force=true to replace it.`);
    await rm(outDir, { recursive: true, force: true });
  }
  await mkdir(dirname(outDir), { recursive: true });
  await cp(bundledExtensionDir, outDir, { recursive: true });
  const manifest = JSON.parse(await readFile(resolve(outDir, 'manifest.json'), 'utf8'));
  const releaseStatus = await readBundledMeetingAdapterReleaseStatus();
  return {
    type: 'meeting_timeline_bundled_adapters',
    schema: 'meeting_timeline_bundled_adapters',
    schema_version: 1,
    ok: true,
    out_dir: outDir,
    platforms: ['google_meet', 'microsoft_teams', 'zoom'],
    manifest_version: manifest.manifest_version,
    matches: manifest.content_scripts?.flatMap((item) => item.matches ?? []) ?? [],
    release_status: releaseStatus,
    next_step: 'Load this directory as an unpacked extension in Chrome or Edge.',
  };
}

export async function readBundledMeetingAdapterReleaseStatus() {
  if (!existsSync(bundledReleaseStatusFile)) {
    return {
      type: 'three_platform_adapter_release_status',
      schema: 'three_platform_adapter_release_status',
      schema_version: 1,
      ok: false,
      production_ready: false,
      speaker_ready: false,
      full_production_ready: false,
      status_file: bundledReleaseStatusFile,
      error: 'This SDK package does not contain a bundled three-platform release status.',
    };
  }
  const status = JSON.parse(await readFile(bundledReleaseStatusFile, 'utf8'));
  return {
    ...status,
    ok: status.schema === 'three_platform_adapter_release_status',
    status_file: bundledReleaseStatusFile,
  };
}

function formatBundledMeetingAdapterReleaseStatus(report = {}) {
  const lines = [
    `meeting_timeline_adapter_release_status | ok=${report.ok ? 'yes' : 'no'} | production_ready=${report.production_ready ? 'yes' : 'no'} | speaker_ready=${report.speaker_ready ? 'yes' : 'no'} | full_production_ready=${report.full_production_ready ? 'yes' : 'no'}`,
  ];
  for (const row of report.adapters ?? []) {
    lines.push(`${row.platform}: installable=${row.installable ? 'yes' : 'no'} startable=${row.startable ? 'yes' : 'no'} core=${row.production_ready ? 'yes' : 'no'} speaker=${row.speaker_real_meeting_accepted ? 'yes' : 'no'} meeting=${row.meeting_id ?? '-'}`);
  }
  if (report.remaining_gate) lines.push(`remaining_gate=${report.remaining_gate}`);
  if (report.speaker_remaining_gate) lines.push(`speaker_remaining_gate=${report.speaker_remaining_gate}`);
  if (report.error) lines.push(`error=${report.error}`);
  return lines.join('\n');
}

export async function runInstallMeetingAdaptersCli(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (args.get('status') === 'true') {
    const report = await readBundledMeetingAdapterReleaseStatus();
    console.log(args.get('json') === 'true'
      ? JSON.stringify(report, null, 2)
      : formatBundledMeetingAdapterReleaseStatus(report));
    if (args.get('fail-on-incomplete') === 'true' && report.production_ready !== true) process.exitCode = 2;
    if (args.get('require-speaker') === 'true' && report.full_production_ready !== true) process.exitCode = 2;
    return report;
  }
  if (args.get('print-path') === 'true') {
    const report = {
      type: 'meeting_timeline_bundled_adapters_path',
      ok: existsSync(resolve(bundledExtensionDir, 'manifest.json')),
      path: bundledExtensionDir,
    };
    console.log(args.get('json') === 'true' ? JSON.stringify(report, null, 2) : report.path);
    if (!report.ok) process.exitCode = 2;
    return report;
  }
  try {
    const report = await exportBundledMeetingAdapters({
      outDir: args.get('out-dir'),
      force: args.get('force') === 'true',
    });
    console.log(args.get('json') === 'true'
      ? JSON.stringify(report, null, 2)
      : `meeting_timeline_bundled_adapters | ok=yes | dir=${report.out_dir}`);
    return report;
  } catch (error) {
    const report = { type: 'meeting_timeline_bundled_adapters', ok: false, error: String(error?.message ?? error) };
    console.error(args.get('json') === 'true'
      ? JSON.stringify(report, null, 2)
      : `meeting_timeline_bundled_adapters | ok=no | error=${report.error}`);
    process.exitCode = 2;
    return report;
  }
}
