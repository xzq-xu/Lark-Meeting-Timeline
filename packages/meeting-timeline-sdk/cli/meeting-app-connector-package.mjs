import { mkdir, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingAppTimelineConnectorAdoptionIndex,
  buildMeetingAppTimelineConnectorBridgeHandoff,
  buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineConnectorHostInstallChecklist,
  buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport,
  buildMeetingAppTimelineConnectorSmokePlan,
  buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport,
  createMeetingAppTimelineSdk,
  runMeetingAppTimelineConnectorBridgeSmoke,
  runMeetingAppTimelineConnectorSmokePlan,
} from '../index.mjs';

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

function stripConnectorPackage(pkg = {}) {
  return {
    ...pkg,
    extension: pkg.extension
      ? {
        ...pkg.extension,
        scaffold: pkg.extension.scaffold
          ? {
            ...pkg.extension.scaffold,
            files: stripFileContents(pkg.extension.scaffold.files),
          }
          : undefined,
      }
      : undefined,
  };
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function writeContent(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, typeof value === 'string' ? value : `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function surfaceFileName(surface) {
  return String(surface || 'surface').replace(/[^a-z0-9_-]+/gi, '-');
}

function markdownTable(rows = [], columns = []) {
  const header = `| ${columns.map((column) => column.label).join(' | ')} |`;
  const separator = `| ${columns.map(() => '---').join(' | ')} |`;
  const body = rows.map((row) => `| ${columns.map((column) => String(column.value(row) ?? '')).join(' | ')} |`);
  return [header, separator, ...body].join('\n');
}

function connectorQuickstartMarkdown(pkg = {}) {
  const startupRows = pkg.startup_plans?.matrix?.rows ?? [];
  const actionRows = pkg.runtime_events?.plan_matrix?.rows ?? [];
  const startupTable = markdownTable(startupRows, [
    { label: 'Platform', value: (row) => row.platform },
    { label: 'Surface', value: (row) => row.selected_surface },
    { label: 'Install target', value: (row) => row.install_target },
    { label: 'Runtime preset', value: (row) => row.runtime_preset },
    { label: 'Ready', value: (row) => row.realtime_startup_ready === true ? 'yes' : 'no' },
  ]);
  const actionSummary = markdownTable(actionRows.slice(0, 12), [
    { label: 'Platform', value: (row) => row.platform },
    { label: 'Action', value: (row) => row.action },
    { label: 'Client method', value: (row) => row.client_method },
    { label: 'Producer', value: (row) => row.producer },
  ]);
  return [
    '# Meeting App Timeline Connector Quickstart',
    '',
    `Package: \`${pkg.id ?? 'meeting-app-timeline-connector'}\``,
    `Runtime event endpoint: \`${pkg.runtime_events?.endpoint ?? '/api/meeting-platform/runtime-events'}\``,
    `Platforms: ${(pkg.platforms ?? []).map((platform) => `\`${platform}\``).join(', ')}`,
    `Surfaces: ${(pkg.surfaces ?? []).map((surface) => `\`${surface}\``).join(', ')}`,
    '',
    '## Read These Files First',
    '',
    '- `connector-handoff.json`: compact handoff summary for another host project.',
    '- `connector-adoption-index.json`: per-platform P0/P1/P2 adoption status, install target, realtime readiness, bridge readiness, and production evidence gaps.',
    '- `connector-bridge-handoff.json`: lightweight connector bridge handoff for browser extension, Electron WebView preload, mobile WebView, or native helper integration.',
    '- `connector-bridge-handoff-acceptance.json`: standalone gate for the lightweight connector bridge handoff.',
    '- `connector-bridge-smoke-report.json`: dry-run bridge message dispatch report for observe candidates, insert mark, tracks, and preflight.',
    '- `host-install-checklist.json`: machine-readable per-platform install checklist for host CI or setup UI.',
    '- `host-install-checklist-acceptance.json`: standalone acceptance report for the host install checklist.',
    '- `connector-smoke-plan.json`: executable per-platform smoke order: observe candidates, insert annotation, then optional speaker/participant markers.',
    '- `connector-smoke-plan-acceptance.json`: standalone gate for the connector smoke plan.',
    '- `connector-smoke-run-report.json`: dry-run execution report proving the generated smoke plan maps to SDK runtime client methods.',
    '- `startup-plan-matrix.json`: selected runtime surface, install target, bridge, and startup actions per platform.',
    '- `adapter-blueprint-matrix.json`: platform adapter blueprint, acceptance gates, and surface order.',
    '- `runtime-event-plan-matrix.json`: supported runtime actions and client methods.',
    '- `connector-package.json`: full package without embedded extension file contents.',
    '- `extension/`: generated MV3 extension scaffold when browser extension surface is enabled.',
    '',
    '## Runtime Contract',
    '',
    '- Realtime annotations must carry `captured_at_ms` from the device or ink end time.',
    '- Provider events and transcript/artifact import must not block realtime annotation insertion.',
    '- Use `startup-plan-matrix.json` to choose the host runtime surface before installing bridges.',
    '- Use `connector-handoff.json` for CI and handoff dashboards instead of parsing every raw file.',
    '',
    '## Startup Matrix',
    '',
    startupTable,
    '',
    '## Minimal Runtime Event Wiring',
    '',
    '```js',
    "import connectorPackage from './connector-package.json' assert { type: 'json' };",
    "import { createMeetingAppTimelineConnectorRuntimeClient } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-connector-package';",
    '',
    'const client = createMeetingAppTimelineConnectorRuntimeClient(connectorPackage, { fetch });',
    '',
    'await client.observePlatformCandidates({',
    '  tabs,',
    '  captured_at_ms: Date.now(),',
    '});',
    '',
    "await client.insertAnnotation('google-meet', {",
    "  id: 'mark-1',",
    "  label: 'why?',",
    '  captured_at_ms: Date.now(),',
    '});',
    '```',
    '',
    '## Runtime Actions',
    '',
    actionSummary,
    actionRows.length > 12 ? `\nShowing 12 of ${actionRows.length} actions. See \`runtime-event-plan-matrix.json\` for the full matrix.` : '',
    '',
  ].join('\n');
}

export function parseMeetingAppConnectorPackageCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingAppConnectorPackageCliOptionsFromArgs(args = new Map()) {
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  const surfaces = listArg(
    args.get('surfaces') || args.get('surface'),
    ['browser-extension', 'native-detector'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: args.get('json') === 'true',
    includePackage: boolArg(args.get('include-package'), false),
    includeExtensionScaffold: boolArg(args.get('include-extension-scaffold'), true),
    observeTracks: boolArg(args.get('observe-tracks') || args.get('observeTracks'), true),
    failOnRejected: args.get('fail-on-rejected') === 'true' || args.get('fail-on-failed') === 'true',
    requiredPlatforms,
    surfaces,
  };
}

async function writeConnectorPackageFiles(outDir, pkg = {}) {
  if (!outDir) return [];
  const writtenFiles = [];
  const root = resolve(outDir);
  const write = async (relative, value, content = false) => {
    const file = resolve(root, relative);
    if (content) await writeContent(file, value);
    else await writeJson(file, value);
    writtenFiles.push(file);
  };

  await write('connector-package.json', stripConnectorPackage(pkg));
  await write('host-package.json', pkg.host_package);
  await write('handoff-matrix.json', pkg.handoff_matrix);
  await write('handoff-acceptance.json', pkg.handoff_acceptance);
  await write('runtime-event-plan-matrix.json', pkg.runtime_events?.plan_matrix);
  await write('adapter-blueprint-matrix.json', pkg.adapter_blueprints?.matrix);
  await write('startup-plan-matrix.json', pkg.startup_plans?.matrix);
  await write('connector-handoff.json', buildMeetingAppTimelineConnectorHandoff(pkg));
  await write('connector-adoption-index.json', buildMeetingAppTimelineConnectorAdoptionIndex(pkg));
  await write('connector-bridge-handoff.json', buildMeetingAppTimelineConnectorBridgeHandoff(pkg));
  await write('connector-bridge-handoff-acceptance.json', buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(pkg));
  await write('connector-bridge-smoke-report.json', await runMeetingAppTimelineConnectorBridgeSmoke(pkg));
  await write('host-install-checklist.json', buildMeetingAppTimelineConnectorHostInstallChecklist(pkg));
  await write('host-install-checklist-acceptance.json', buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(pkg));
  await write('connector-smoke-plan.json', buildMeetingAppTimelineConnectorSmokePlan(pkg));
  await write('connector-smoke-plan-acceptance.json', buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(pkg));
  await write('connector-smoke-run-report.json', await runMeetingAppTimelineConnectorSmokePlan(pkg));
  await write('connector-quickstart.md', connectorQuickstartMarkdown(pkg), true);

  for (const [surface, plan] of Object.entries(pkg.observer_plan_by_surface ?? {})) {
    await write(`observer-plan-${surfaceFileName(surface)}.json`, plan);
  }
  for (const [surface, config] of Object.entries(pkg.scheduler_config_by_surface ?? {})) {
    await write(`scheduler-config-${surfaceFileName(surface)}.json`, config);
  }

  for (const file of pkg.extension?.scaffold?.files ?? []) {
    const target = resolve(root, 'extension', file.path);
    await writeContent(target, file.content);
    writtenFiles.push(target);
  }

  return writtenFiles;
}

export async function buildMeetingAppConnectorPackageCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    reportFile = '',
    includePackage = false,
    includeExtensionScaffold = true,
    observeTracks = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
    surfaces = ['browser-extension', 'native-detector'],
  } = options;

  const sdk = createMeetingAppTimelineSdk({
    baseUrl,
    platforms: requiredPlatforms,
  });
  const pkg = sdk.connectorPackage({
    platforms: requiredPlatforms,
    surfaces,
    includeExtensionScaffold,
    observeTracks,
  });
  const connectorHandoff = buildMeetingAppTimelineConnectorHandoff(pkg);
  const connectorAdoptionIndex = buildMeetingAppTimelineConnectorAdoptionIndex(pkg);
  const connectorBridgeHandoff = buildMeetingAppTimelineConnectorBridgeHandoff(pkg);
  const connectorBridgeHandoffAcceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(connectorBridgeHandoff);
  const connectorBridgeSmokeReport = await runMeetingAppTimelineConnectorBridgeSmoke(connectorBridgeHandoff);
  const hostInstallChecklist = buildMeetingAppTimelineConnectorHostInstallChecklist(pkg);
  const hostInstallChecklistAcceptance = buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(hostInstallChecklist);
  const smokePlan = buildMeetingAppTimelineConnectorSmokePlan(hostInstallChecklist);
  const smokePlanAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(smokePlan);
  const smokeRunReport = await runMeetingAppTimelineConnectorSmokePlan(smokePlan);
  const writtenFiles = await writeConnectorPackageFiles(outDir, pkg);

  const report = {
    type: 'meeting_app_timeline_connector_package_report',
    ok: pkg.accepted === true,
    base_url: baseUrl,
    out_dir: outDir || undefined,
    platform_count: pkg.platform_count,
    surface_count: pkg.surface_count,
    handoff_count: pkg.handoff_count,
    ready_count: pkg.ready_count,
    accepted: pkg.accepted === true,
    required_platforms: requiredPlatforms,
    surfaces: pkg.surfaces,
    extension_scaffold: Boolean(pkg.extension?.scaffold),
    extension_accepted: pkg.extension?.acceptance?.accepted,
    runtime_event_action_count: pkg.runtime_events?.action_count ?? 0,
    adapter_blueprint_ready_count: pkg.adapter_blueprints?.ready_count ?? 0,
    startup_plan_ready_count: pkg.startup_plans?.realtime_startup_ready_count ?? 0,
    observer_surface_count: Object.keys(pkg.observer_plan_by_surface ?? {}).length,
    scheduler_surface_count: Object.keys(pkg.scheduler_config_by_surface ?? {}).length,
    written_files: writtenFiles,
    rows: pkg.handoff_matrix?.rows?.map((row) => ({
      platform: row.platform,
      surface: row.surface,
      ready_to_start: row.ready_to_start,
      realtime_annotation_ready: row.realtime_annotation_ready,
      speaker_track_ready: row.speaker_track_ready,
      participant_track_ready: row.participant_track_ready,
      install_target: row.install_target,
      start_mode: row.start_mode,
    })) ?? [],
    handoff: connectorHandoff,
    adoption_index: connectorAdoptionIndex,
    bridge_handoff: connectorBridgeHandoff,
    bridge_handoff_acceptance: connectorBridgeHandoffAcceptance,
    bridge_smoke_report: connectorBridgeSmokeReport,
    host_install_checklist: hostInstallChecklist,
    host_install_checklist_acceptance: hostInstallChecklistAcceptance,
    smoke_plan: smokePlan,
    smoke_plan_acceptance: smokePlanAcceptance,
    smoke_run_report: smokeRunReport,
    package: includePackage ? pkg : stripConnectorPackage(pkg),
    next_actions: pkg.next_actions,
  };

  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingAppConnectorPackageCliReport(report = {}) {
  const lines = [
    `meeting_app_timeline_connector_package_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | surfaces=${report.surface_count} | handoffs=${report.handoff_count} | ready=${report.ready_count} | blueprint_ready=${report.adapter_blueprint_ready_count} | startup_ready=${report.startup_plan_ready_count} | extension=${boolLabel(report.extension_scaffold)} | extension_accepted=${boolLabel(report.extension_accepted)} | runtime_actions=${report.runtime_event_action_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}/${row.surface}: ready=${boolLabel(row.ready_to_start)} realtime=${boolLabel(row.realtime_annotation_ready)} speaker=${boolLabel(row.speaker_track_ready)} participant=${boolLabel(row.participant_track_ready)} install=${row.install_target ?? 'n/a'} start=${row.start_mode ?? 'n/a'}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  return lines.join('\n');
}

export async function runMeetingAppConnectorPackageCli(argv = process.argv.slice(2), io = console) {
  const options = meetingAppConnectorPackageCliOptionsFromArgs(
    parseMeetingAppConnectorPackageCliArgs(argv),
  );
  const report = await buildMeetingAppConnectorPackageCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingAppConnectorPackageCliReport(report));
  }
  if (!report.ok && options.failOnRejected) process.exitCode = 2;
  return report;
}
