import { existsSync } from 'node:fs';
import { cp, mkdir, readFile, rm } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const bundledExtensionDir = resolve(packageRoot, 'runtime', 'browser-extension');

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
  return {
    type: 'meeting_timeline_bundled_adapters',
    schema: 'meeting_timeline_bundled_adapters',
    schema_version: 1,
    ok: true,
    out_dir: outDir,
    platforms: ['google_meet', 'microsoft_teams', 'zoom'],
    manifest_version: manifest.manifest_version,
    matches: manifest.content_scripts?.flatMap((item) => item.matches ?? []) ?? [],
    next_step: 'Load this directory as an unpacked extension in Chrome or Edge.',
  };
}

export async function runInstallMeetingAdaptersCli(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
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
