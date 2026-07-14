#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-contract.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const evidenceDir = String(args.get('evidence-dir') || args.get('evidenceDir') || 'data');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || args.get('contract-dir') || args.get('contractDir') || 'data/meeting-platform-adapter-contracts'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writeContracts = args.get('write-contracts') !== 'false';
const failOnIncomplete = args.get('fail-on-incomplete') === 'true' || args.get('fail-on-missing') === 'true';
const failOnRejected = args.get('fail-on-rejected') === 'true' || args.get('fail-on-acceptance') === 'true';
const acceptanceTarget = String(args.get('acceptance-target') || args.get('target') || 'contract');
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function contractFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function summarizeContract(contract = {}) {
  return {
    platform: contract.platform,
    display_name: contract.display_name,
    mode: contract.mode,
    browser_observer: contract.supported_surfaces?.browser_observer === true,
    candidate_observation_ready: contract.candidate_observation?.runtime_event_action === 'observe_platform_candidates',
    candidate_observer_message_type: contract.candidate_observation?.message_type,
    candidate_observer_permission: contract.candidate_observation?.required_permission,
    candidate_observer_endpoint: contract.candidate_observation?.endpoint,
    provider_observer: contract.supported_surfaces?.provider_webhook_or_event_subscription === true,
    provider_transport: contract.provider_observer?.transport,
    provider_ready: contract.readiness?.provider_ready === true,
    production_ready: contract.readiness?.production_ready === true,
    ready_for_realtime_annotations: contract.readiness?.ready_for_realtime_annotations === true,
    start_create_on: contract.realtime_axis?.start?.create_on,
    end_create_on: contract.realtime_axis?.end?.create_on,
    insert_mark_endpoint: contract.annotations?.endpoints?.insertMark,
    browser_matches: contract.local_observer?.matches ?? [],
    provider_start_events: contract.provider_observer?.events?.start_events ?? [],
    provider_end_events: contract.provider_observer?.events?.end_events ?? [],
    speaker_min_stable_ms: contract.local_observer?.speaker_markers?.filter?.min_stable_ms,
    transcript_realtime_dependency: contract.transcript?.realtime_dependency,
    missing_items: contract.readiness?.missing_items ?? [],
    contract_file: contractFile(contract.platform),
  };
}

async function buildReport() {
  const buildOptions = {
    baseUrl,
    env: process.env,
    evidenceDir,
    platforms: requiredPlatforms,
  };
  const matrix = buildMeetingPlatformAdapterContractMatrix(buildOptions);
  const acceptance = buildMeetingPlatformAdapterContractAcceptanceMatrix({
    ...buildOptions,
    target: acceptanceTarget,
  });
  const writtenFiles = [];
  if (writeContracts) {
    for (const contract of matrix.contracts) {
      const file = contractFile(contract.platform);
      await writeJson(file, contract);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.contracts.map((contract) => summarizeContract(contract));
  const missingItemCount = rows.reduce((total, row) => total + row.missing_items.length, 0);
  const ok = matrix.platform_count > 0
    && (!failOnIncomplete || missingItemCount === 0)
    && (!failOnRejected || acceptance.rejected_count === 0);
  return {
    type: 'meeting_platform_adapter_contract_report',
    ok,
    requirement: failOnRejected
      ? `adapter_contract_acceptance:${acceptanceTarget}`
      : failOnIncomplete ? 'no_missing_adapter_evidence_items' : 'adapter_contract_exported',
    base_url: baseUrl,
    evidence_dir: evidenceDir,
    acceptance_target: acceptanceTarget,
    out_dir: writeContracts ? outDir : undefined,
    write_contracts: writeContracts,
    platform_count: matrix.platform_count,
    browser_observer_count: matrix.browser_observer_count,
    candidate_observer_count: matrix.candidate_observer_count,
    provider_observer_count: matrix.provider_observer_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    missing_item_count: missingItemCount,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    acceptance,
    matrix,
    next_actions: matrix.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_adapter_contract_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | browser_observers=${report.browser_observer_count} | candidate_observers=${report.candidate_observer_count} | providers=${report.provider_observer_count} | accepted=${report.acceptance.accepted_count}/${report.acceptance.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | missing_items=${report.missing_item_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      const acceptanceRow = report.acceptance.rows.find((item) => item.platform === row.platform);
      console.log(`${row.platform}: mode=${row.mode} accepted=${boolLabel(acceptanceRow?.accepted)} browser=${boolLabel(row.browser_observer)} candidates=${boolLabel(row.candidate_observation_ready)} provider=${boolLabel(row.provider_observer)} provider_ready=${boolLabel(row.provider_ready)} missing=${row.missing_items.length} contract=${row.contract_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && (failOnIncomplete || failOnRejected)) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
