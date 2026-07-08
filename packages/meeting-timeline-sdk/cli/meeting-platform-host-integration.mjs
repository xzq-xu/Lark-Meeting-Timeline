import { mkdir, writeFile } from 'node:fs/promises';
import { basename, dirname, isAbsolute, relative, resolve } from 'node:path';
import {
  buildMeetingPlatformHostIntegrationScaffold,
  buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport,
} from '../adapters/platform-host-integration.mjs';

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function listArg(value, fallback = []) {
  if (value == null || value === '') return fallback;
  return unique(String(value).split(',').map((item) => item.trim()).filter(Boolean));
}

function boolArg(value, fallback = false) {
  if (value == null || value === '') return fallback;
  const text = String(value).trim().toLowerCase();
  if (['1', 'true', 'yes', 'on'].includes(text)) return true;
  if (['0', 'false', 'no', 'off'].includes(text)) return false;
  return fallback;
}

function stripFileContents(files = []) {
  return files.map(({ content, ...file }) => file);
}

function compactPlan(plan = {}) {
  return {
    type: plan.type,
    schema: plan.schema,
    schema_version: plan.schema_version,
    base_url: plan.base_url,
    base_path: plan.base_path,
    platforms: plan.platforms,
    runtime_contract: plan.runtime_contract,
    endpoints: plan.endpoints,
    commands: plan.commands,
    adapter_runtime_contract: plan.adapter_runtime_contract,
    adapter_startup_plan_matrix: plan.adapter_startup_plan_matrix
      ? {
        schema: plan.adapter_startup_plan_matrix.schema,
        platform_count: plan.adapter_startup_plan_matrix.platform_count,
        accepted_count: plan.adapter_startup_plan_matrix.accepted_count,
        realtime_startup_ready_count: plan.adapter_startup_plan_matrix.realtime_startup_ready_count,
        provider_reconcile_surface_count: plan.adapter_startup_plan_matrix.provider_reconcile_surface_count,
        rows: plan.adapter_startup_plan_matrix.rows,
      }
      : undefined,
    candidate_observation_contract: plan.candidate_observation_contract,
    meeting_track_contract: plan.meeting_track_contract,
    platform_conformance: plan.platform_conformance_report
      ? {
        accepted: plan.platform_conformance_report.accepted,
        platform_count: plan.platform_conformance_report.platform_count,
        accepted_count: plan.platform_conformance_report.accepted_count,
        blocking_count: plan.platform_conformance_report.blocking_count,
      }
      : undefined,
    next_actions: plan.next_actions,
  };
}

function stripScaffold(scaffold = {}) {
  return {
    ...scaffold,
    plan: compactPlan(scaffold.plan),
    files: stripFileContents(scaffold.files),
  };
}

function compactAcceptance(acceptance = {}) {
  return {
    type: acceptance.type,
    schema: acceptance.schema,
    schema_version: acceptance.schema_version,
    accepted: acceptance.accepted,
    platform_count: acceptance.platform_count,
    file_count: acceptance.file_count,
    candidate_observation_ready: acceptance.candidate_observation_ready,
    candidate_observer_count: acceptance.candidate_observer_count,
    candidate_observer_missing_count: acceptance.candidate_observer_missing_count,
    platform_conformance_ready: acceptance.platform_conformance_ready,
    platform_conformance_accepted_count: acceptance.platform_conformance_accepted_count,
    platform_conformance_blocking_count: acceptance.platform_conformance_blocking_count,
    observer_plan_ready: acceptance.observer_plan_ready,
    observer_plan_ready_count: acceptance.observer_plan_ready_count,
    observer_plan_preflight_accepted_count: acceptance.observer_plan_preflight_accepted_count,
    adapter_blueprint_ready: acceptance.adapter_blueprint_ready,
    adapter_blueprint_ready_count: acceptance.adapter_blueprint_ready_count,
    adapter_startup_ready: acceptance.adapter_startup_ready,
    adapter_startup_ready_count: acceptance.adapter_startup_ready_count,
    adapter_runtime_ready: acceptance.adapter_runtime_ready,
    adapter_runtime_ready_count: acceptance.adapter_runtime_ready_count,
    meeting_track_ready: acceptance.meeting_track_ready,
    speaker_track_ready_count: acceptance.speaker_track_ready_count,
    participant_track_ready_count: acceptance.participant_track_ready_count,
    required_files: acceptance.required_files,
    blocking_count: acceptance.blocking_count,
    warning_count: acceptance.warning_count,
    issues: acceptance.issues,
  };
}

export function parseMeetingPlatformHostIntegrationCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformHostIntegrationCliOptionsFromArgs(args = new Map()) {
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    basePath: String(args.get('base-path') || args.get('basePath') || '/api/platform-events'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    packageName: String(args.get('package-name') || args.get('packageName') || 'meeting-platform-timeline-host'),
    jsonOutput: boolArg(args.get('json'), false),
    includeScaffold: boolArg(args.get('include-scaffold'), false),
    writeScaffold: boolArg(args.get('write-scaffold'), true),
    failOnRejected: boolArg(args.get('fail-on-rejected'), false) || boolArg(args.get('fail-on-failed'), false),
    requiredPlatforms: listArg(
      args.get('required-platforms') || args.get('platforms'),
      ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    ),
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function writeContent(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, value, 'utf8');
}

function assertInsideRoot(root, file) {
  const delta = relative(root, file);
  if (delta === '' || (!delta.startsWith('..') && !isAbsolute(delta))) return;
  throw new Error(`Refusing to write outside output directory: ${file}`);
}

async function writeScaffoldFiles(outDir, scaffold = {}, options = {}) {
  if (!outDir || options.writeScaffold === false) return [];
  const root = resolve(outDir);
  const writtenFiles = [];
  for (const file of scaffold.files ?? []) {
    const destination = resolve(root, file.path);
    assertInsideRoot(root, destination);
    await writeContent(destination, file.content ?? '');
    writtenFiles.push(destination);
  }
  return writtenFiles;
}

function generatedFiles(scaffold = {}, outDir = '') {
  const root = outDir ? resolve(outDir) : '';
  return (scaffold.files ?? []).map((file) => ({
    path: file.path,
    role: file.role,
    mime: file.mime,
    file: root ? resolve(root, file.path) : undefined,
  }));
}

export async function buildMeetingPlatformHostIntegrationCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    basePath = '/api/platform-events',
    outDir = '',
    reportFile = '',
    packageName = 'meeting-platform-timeline-host',
    includeScaffold = false,
    writeScaffold = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const scaffold = buildMeetingPlatformHostIntegrationScaffold({
    ...options,
    baseUrl,
    basePath,
    packageName,
    platforms: requiredPlatforms,
  });
  const acceptance = buildMeetingPlatformHostIntegrationScaffoldAcceptanceReport(scaffold);
  const writtenFiles = await writeScaffoldFiles(outDir, scaffold, { writeScaffold });
  const report = {
    type: 'meeting_platform_host_integration_scaffold_report',
    ok: acceptance.accepted === true,
    accepted: acceptance.accepted === true,
    requirement: 'reusable_meeting_timeline_host_scaffold_for_google_meet_teams_zoom_webex_lark',
    base_url: baseUrl,
    base_path: basePath,
    package_name: packageName,
    out_dir: outDir || undefined,
    write_scaffold: writeScaffold,
    platform_count: scaffold.platforms?.length ?? 0,
    file_count: scaffold.files?.length ?? 0,
    written_file_count: writtenFiles.length,
    adapter_runtime_ready: acceptance.adapter_runtime_ready === true,
    adapter_runtime_ready_count: acceptance.adapter_runtime_ready_count,
    candidate_observation_ready: acceptance.candidate_observation_ready === true,
    candidate_observer_count: acceptance.candidate_observer_count,
    meeting_track_ready: acceptance.meeting_track_ready === true,
    speaker_track_ready_count: acceptance.speaker_track_ready_count,
    participant_track_ready_count: acceptance.participant_track_ready_count,
    platform_conformance_ready: acceptance.platform_conformance_ready === true,
    platform_conformance_accepted_count: acceptance.platform_conformance_accepted_count,
    observer_plan_ready: acceptance.observer_plan_ready === true,
    adapter_blueprint_ready: acceptance.adapter_blueprint_ready === true,
    adapter_startup_ready: acceptance.adapter_startup_ready === true,
    adapter_startup_ready_count: acceptance.adapter_startup_ready_count,
    blocking_count: acceptance.blocking_count,
    warning_count: acceptance.warning_count,
    required_platforms: requiredPlatforms,
    platforms: scaffold.platforms,
    generated_files: generatedFiles(scaffold, outDir),
    required_files: acceptance.required_files,
    written_files: writtenFiles,
    acceptance: compactAcceptance(acceptance),
    scaffold: includeScaffold ? scaffold : stripScaffold(scaffold),
    next_actions: scaffold.plan?.next_actions ?? [],
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformHostIntegrationCliReport(report = {}) {
  const lines = [
    `meeting_platform_host_integration_scaffold_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | files=${report.file_count} | written=${report.written_file_count} | adapters=${boolLabel(report.adapter_runtime_ready)} | startup=${boolLabel(report.adapter_startup_ready)} | candidates=${boolLabel(report.candidate_observation_ready)} | tracks=${boolLabel(report.meeting_track_ready)} | conformance=${boolLabel(report.platform_conformance_ready)} | blocking=${report.blocking_count}`,
  ];
  for (const platform of report.platforms ?? []) {
    const adapterFile = report.generated_files?.find((file) => file.path === `src/platform-adapters/${platform}.mjs`);
    lines.push(`${platform}: adapter=${adapterFile?.path ?? 'n/a'} captured_at_ms=yes`);
  }
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.slice(0, 8).join(',')}`);
  return lines.join('\n');
}

export async function runMeetingPlatformHostIntegrationCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformHostIntegrationCliOptionsFromArgs(
    parseMeetingPlatformHostIntegrationCliArgs(argv),
  );
  const report = await buildMeetingPlatformHostIntegrationCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformHostIntegrationCliReport(report));
  }
  if (!report.ok && options.failOnRejected) process.exitCode = 2;
  return report;
}
