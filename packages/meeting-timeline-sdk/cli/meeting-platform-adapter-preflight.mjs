import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterCandidatePreflight,
  buildMeetingPlatformAdapterCurrentWindowPreflight,
  buildMeetingPlatformAdapterPreflight,
  buildMeetingPlatformAdapterPreflightMatrix,
} from '../adapters/platform-adapter-preflight.mjs';

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

function hasCandidateInput(input = {}) {
  return Boolean(input.windows || input.tabs || input.candidates || input.candidate);
}

function hasPlatformInput(input = {}) {
  return Boolean(input.platform || input.provider || input.url || input.href || input.meeting_url || input.meetingUrl);
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function compactPreflightResult(result = {}, options = {}) {
  if (!result || typeof result !== 'object') return result;
  if (result.schema === 'meeting_platform_adapter_preflight_matrix') {
    return {
      ...result,
      preflights: options.includePreflights === true ? result.preflights : undefined,
    };
  }
  if (result.schema === 'meeting_platform_adapter_candidate_preflight') {
    return {
      ...result,
      preflights: options.includePreflights === true ? result.preflights : undefined,
    };
  }
  return result;
}

function expandGroupedEvidenceToPlatformInputs(input = {}) {
  const groupedSnapshots = isPlainObject(input.snapshots) ? input.snapshots : undefined;
  const groupedDomSnapshots = isPlainObject(input.domSnapshots) ? input.domSnapshots : undefined;
  const groupedDomSnapshotsSnake = isPlainObject(input.dom_snapshots) ? input.dom_snapshots : undefined;
  if (!groupedSnapshots && !groupedDomSnapshots && !groupedDomSnapshotsSnake) return input;
  const inputByPlatform = { ...(input.inputByPlatform ?? input.input_by_platform ?? {}) };
  const mergeGrouped = (source, keyName) => {
    if (!source) return;
    for (const [platform, value] of Object.entries(source)) {
      inputByPlatform[platform] = {
        ...(inputByPlatform[platform] ?? {}),
        [keyName]: value,
      };
    }
  };
  mergeGrouped(groupedSnapshots, 'snapshots');
  mergeGrouped(groupedDomSnapshots, 'domSnapshots');
  mergeGrouped(groupedDomSnapshotsSnake, 'dom_snapshots');
  return compactUndefined({
    ...input,
    snapshots: groupedSnapshots ? undefined : input.snapshots,
    domSnapshots: groupedDomSnapshots ? undefined : input.domSnapshots,
    dom_snapshots: groupedDomSnapshotsSnake ? undefined : input.dom_snapshots,
    inputByPlatform,
  });
}

export function parseMeetingPlatformAdapterPreflightCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterPreflightCliOptionsFromArgs(args = new Map()) {
  const explicitPlatform = String(args.get('platform') || args.get('platform-key') || '');
  const requiredPlatforms = listArg(
    args.get('required-platforms') || args.get('platforms'),
    explicitPlatform ? [explicitPlatform] : ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  );
  return {
    baseUrl: String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787'),
    inputFile: String(args.get('input-file') || args.get('inputFile') || ''),
    outFile: String(args.get('out-file') || args.get('outFile') || ''),
    reportFile: String(args.get('report-file') || args.get('write-report') || ''),
    jsonOutput: boolArg(args.get('json'), false),
    mode: String(args.get('mode') || args.get('preflight-mode') || 'auto'),
    failOnRejected: boolArg(args.get('fail-on-rejected'), false) || boolArg(args.get('fail-on-failed'), false),
    includePreflights: boolArg(args.get('include-preflights'), false),
    includeCapturedSnapshot: boolArg(args.get('include-captured-snapshot'), false),
    requireSpeakerTrack: boolArg(args.get('require-speaker-track'), false)
      || boolArg(args.get('requireSpeakerTrack'), false),
    requireCompleteLifecycle: boolArg(args.get('require-complete-lifecycle'), false)
      || boolArg(args.get('requireCompleteLifecycle'), false),
    requireMeetingEnd: boolArg(args.get('require-meeting-end'), false)
      || boolArg(args.get('requireMeetingEnd'), false),
    url: String(args.get('url') || args.get('href') || args.get('meeting-url') || ''),
    title: String(args.get('title') || ''),
    platform: explicitPlatform,
    requiredPlatforms,
  };
}

async function readJson(file) {
  return JSON.parse(await readFile(resolve(file), 'utf8'));
}

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function preflightInput(options = {}) {
  const fromFile = options.inputFile ? await readJson(options.inputFile) : {};
  return compactUndefined({
    ...(options.input && typeof options.input === 'object' ? options.input : {}),
    ...fromFile,
    url: options.url || fromFile.url,
    title: options.title || fromFile.title,
    platform: options.platform || fromFile.platform || undefined,
  });
}

function resolvePreflightMode(input = {}, options = {}) {
  const mode = String(options.mode || 'auto').replaceAll('_', '-');
  if (mode !== 'auto') return mode;
  if (hasCandidateInput(input)) return 'candidates';
  if (!hasPlatformInput(input) && (options.requiredPlatforms?.length ?? 0) > 1) return 'matrix';
  return 'single';
}

function buildPreflightResult(input = {}, options = {}) {
  const preflightOptions = compactUndefined({
    baseUrl: options.baseUrl,
    base_url: options.base_url,
    requireSpeakerTrack: options.requireSpeakerTrack,
    requireCompleteLifecycle: options.requireCompleteLifecycle,
    requireMeetingEnd: options.requireMeetingEnd,
    platforms: options.requiredPlatforms,
    includeCapturedSnapshot: options.includeCapturedSnapshot,
    platform: options.platform || undefined,
  });
  const mode = resolvePreflightMode(input, options);
  if (mode === 'matrix') {
    const matrixInput = expandGroupedEvidenceToPlatformInputs(input);
    return {
      mode,
      result: buildMeetingPlatformAdapterPreflightMatrix(matrixInput, preflightOptions),
    };
  }
  if (mode === 'candidate' || mode === 'candidates') {
    return {
      mode: 'candidates',
      result: buildMeetingPlatformAdapterCandidatePreflight(input, preflightOptions),
    };
  }
  if (mode === 'current-window' || mode === 'current-window-preflight') {
    return {
      mode: 'current-window',
      result: buildMeetingPlatformAdapterCurrentWindowPreflight(input, preflightOptions),
    };
  }
  return {
    mode: 'single',
    result: buildMeetingPlatformAdapterPreflight(input, {
      ...preflightOptions,
      platforms: undefined,
      platform_keys: undefined,
    }),
  };
}

function reportAccepted(result = {}, mode = '') {
  if (mode === 'matrix') return result.accepted_count === result.platform_count;
  return result.accepted === true;
}

export async function buildMeetingPlatformAdapterPreflightCliReport(options = {}) {
  const input = await preflightInput(options);
  const { mode, result } = buildPreflightResult(input, options);
  const compactResult = compactPreflightResult(result, options);
  const report = compactUndefined({
    type: 'meeting_platform_adapter_preflight_report',
    ok: reportAccepted(result, mode),
    mode,
    base_url: options.baseUrl ?? 'http://localhost:8787',
    required_platforms: options.requiredPlatforms,
    platform: result.platform,
    selected_platform: result.selected_platform,
    selection_strategy: result.selection_strategy,
    selected_candidate_index: result.selected_candidate_index,
    selected_candidate_score: result.selected_candidate_score,
    selected_candidate_reason: result.rows?.find?.((row) => row.selected === true)?.selection_reason,
    status: result.status,
    accepted: result.accepted,
    platform_count: result.platform_count,
    candidate_count: result.candidate_count,
    accepted_count: result.accepted_count,
    realtime_ready_count: result.realtime_ready_count,
    live_evidence_ready_count: result.live_evidence_ready_count,
    meeting_start_ready_count: result.meeting_start_ready_count,
    meeting_end_ready_count: result.meeting_end_ready_count,
    speaker_track_ready_count: result.speaker_track_ready_count,
    selected_surface: result.summary?.selected_surface ?? result.startup?.selected_surface,
    issue_codes: result.issues?.map((issue) => issue.code),
    rows: result.rows,
    result: compactResult,
    next_actions: result.next_actions ?? [],
  });
  if (options.outFile) await writeJson(resolve(options.outFile), compactResult);
  if (options.reportFile) await writeJson(resolve(options.reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterPreflightCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_preflight_report | ok=${boolLabel(report.ok)} | mode=${report.mode} | status=${report.status ?? 'n/a'} | platforms=${report.platform_count ?? 'n/a'} | candidates=${report.candidate_count ?? 'n/a'} | accepted=${report.accepted_count ?? boolLabel(report.accepted)} | realtime=${report.realtime_ready_count ?? 'n/a'} | live=${report.live_evidence_ready_count ?? 'n/a'} | surface=${report.selected_surface ?? 'n/a'}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform ?? row.selected_platform ?? 'candidate'}: accepted=${boolLabel(row.accepted)} realtime=${boolLabel(row.realtime_annotation_ready)} startup=${boolLabel(row.startup_ready)} live=${boolLabel(row.live_evidence_ready)} status=${row.status ?? 'n/a'} surface=${row.selected_surface ?? 'n/a'} rank=${row.selection_rank ?? 'n/a'} score=${row.selection_score ?? 'n/a'} reason=${row.selection_reason ?? 'n/a'}`);
  }
  if (report.issue_codes?.length > 0) lines.push(`issues=${report.issue_codes.join(',')}`);
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterPreflightCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterPreflightCliOptionsFromArgs(
    parseMeetingPlatformAdapterPreflightCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterPreflightCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterPreflightCliReport(report));
  }
  if (!report.ok && options.failOnRejected) process.exitCode = 2;
  return report;
}
