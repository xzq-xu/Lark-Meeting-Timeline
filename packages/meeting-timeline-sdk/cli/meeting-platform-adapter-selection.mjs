import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterSelection,
  buildMeetingPlatformAdapterSelectionMatrix,
} from '../adapters/platform-adapter-selection.mjs';

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

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
}

function compactSelection(selection = {}) {
  if (!selection || typeof selection !== 'object') return selection;
  return compactUndefined({
    type: selection.type,
    schema: selection.schema,
    schema_version: selection.schema_version,
    platform: selection.platform,
    display_name: selection.display_name,
    recommended_mode: selection.recommended_mode,
    objective: selection.objective,
    selection: selection.selection,
    runtime_policy: selection.runtime_policy,
    sources: selection.sources,
    current_evidence: selection.current_evidence,
    readiness: selection.readiness,
    gates: selection.gates,
    risk_profile: selection.risk_profile,
    references: selection.references,
    next_actions: selection.next_actions,
  });
}

function compactMatrix(matrix = {}, options = {}) {
  const includeSelections = options.includeSelections === true;
  return compactUndefined({
    ...matrix,
    selections: includeSelections
      ? (matrix.selections ?? []).map((selection) => compactSelection(selection))
      : undefined,
  });
}

export function parseMeetingPlatformAdapterSelectionCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterSelectionCliOptionsFromArgs(args = new Map()) {
  const explicitPlatform = String(args.get('platform') || args.get('platform-key') || '');
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    explicitPlatform ? [explicitPlatform] : ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    outDir: String(args.get('out-dir') || args.get('outDir') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    meetingAppRecordSetFile: String(args.get('meeting-app-record-set-file') || args.get('meetingAppRecordSetFile') || ''),
    providerRecordsFile: String(args.get('provider-records-file') || args.get('providerRecordsFile') || ''),
    localDetectorRecordsFile: String(args.get('local-detector-records-file') || args.get('localDetectorRecordsFile') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    includeSelections: boolArg(args.get('include-selections'), false),
    includeStrategy: boolArg(args.get('include-strategy'), false),
    writeSelections: boolArg(args.get('write-selections'), true),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    platform: explicitPlatform,
    requiredPlatforms,
  };
}

async function readJsonMaybe(file) {
  if (!file) return undefined;
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function selectionFile(outDir, platform) {
  return outDir ? resolve(outDir, `${platform}.json`) : undefined;
}

async function selectionInput(options = {}) {
  const fromFile = await readJsonMaybe(options.inputFile) ?? {};
  return compactUndefined({
    ...fromFile,
    meetingAppRecordSet: await readJsonMaybe(options.meetingAppRecordSetFile) ?? fromFile.meetingAppRecordSet ?? fromFile.meeting_app_record_set,
    providerRecords: await readJsonMaybe(options.providerRecordsFile) ?? fromFile.providerRecords ?? fromFile.provider_records,
    localDetectorRecords: await readJsonMaybe(options.localDetectorRecordsFile) ?? fromFile.localDetectorRecords ?? fromFile.local_detector_records,
  });
}

async function writeSelectionFiles(outDir, matrix = {}, options = {}) {
  if (!outDir || options.writeSelections === false) return [];
  const writtenFiles = [];
  for (const selection of matrix.selections ?? []) {
    const file = selectionFile(outDir, selection.platform);
    await writeJson(file, options.includeSelections ? selection : compactSelection(selection));
    writtenFiles.push(file);
  }
  return writtenFiles;
}

export async function buildMeetingPlatformAdapterSelectionCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    outFile = '',
    reportFile = '',
    includeSelections = false,
    includeStrategy = false,
    writeSelections = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const input = await selectionInput(options);
  const matrix = buildMeetingPlatformAdapterSelectionMatrix(input, {
    ...options,
    baseUrl,
    platforms: requiredPlatforms,
    includeStrategy,
  });
  const singleSelection = options.platform || options.inputFile
    ? buildMeetingPlatformAdapterSelection(options.platform || requiredPlatforms[0], input, {
      ...options,
      baseUrl,
      includeStrategy,
    })
    : undefined;
  const writtenFiles = await writeSelectionFiles(outDir, matrix, { writeSelections, includeSelections });
  if (outFile) {
    await writeJson(resolve(outFile), singleSelection
      ? (includeSelections ? singleSelection : compactSelection(singleSelection))
      : compactMatrix(matrix, { includeSelections }));
  }
  const rows = matrix.rows.map((row) => ({
    ...row,
    selection_file: selectionFile(outDir, row.platform),
  }));
  const blockingRows = rows.filter((row) => (
    row.provider_events_block_realtime === true || row.transcript_blocks_realtime === true
  ));
  const report = {
    type: 'meeting_platform_adapter_selection_report',
    ok: matrix.platform_count > 0 && matrix.selection_ready_count === matrix.platform_count && blockingRows.length === 0,
    requirement: 'selected_platform_adapter_sources_keep_local_axis_first_and_provider_nonblocking',
    base_url: baseUrl,
    out_dir: outDir || undefined,
    out_file: outFile || undefined,
    write_selections: writeSelections,
    platform_count: matrix.platform_count,
    selection_ready_count: matrix.selection_ready_count,
    pilot_evidence_ready_count: matrix.pilot_evidence_ready_count,
    production_evidence_ready_count: matrix.production_evidence_ready_count,
    local_axis_selected_count: matrix.local_axis_selected_count,
    provider_reconcile_count: matrix.provider_reconcile_count,
    post_meeting_artifact_count: matrix.post_meeting_artifact_count,
    blocking_count: blockingRows.length,
    blocking_platforms: blockingRows.map((row) => row.platform),
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    single_selection: singleSelection ? compactSelection(singleSelection) : undefined,
    matrix: compactMatrix(matrix, { includeSelections }),
    next_actions: unique((matrix.selections ?? []).flatMap((selection) => selection.next_actions ?? [])),
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterSelectionCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_selection_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | ready=${report.selection_ready_count} | pilot=${report.pilot_evidence_ready_count} | production=${report.production_evidence_ready_count} | local_axis=${report.local_axis_selected_count} | provider_reconcile=${report.provider_reconcile_count} | artifacts=${report.post_meeting_artifact_count} | blocking=${report.blocking_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: axis=${row.axis_source} surface=${row.axis_surface ?? 'n/a'} timestamp=${row.timestamp_field} provider=${row.provider_reconcile_source ?? 'none'} pilot=${boolLabel(row.pilot_evidence_ready)} production=${boolLabel(row.production_evidence_ready)} selection=${row.selection_file ?? 'n/a'}`);
  }
  if (report.single_selection?.platform) {
    lines.push(`single_selection=${report.single_selection.platform}:${report.single_selection.selection?.axis_source}:${report.single_selection.selection?.axis_surface}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterSelectionCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterSelectionCliOptionsFromArgs(
    parseMeetingPlatformAdapterSelectionCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterSelectionCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterSelectionCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
