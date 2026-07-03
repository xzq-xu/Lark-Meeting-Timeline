#!/usr/bin/env node

import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformFieldIntakeMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-field-intake.mjs';
import {
  boolLabel,
  parseCliArgs,
  unique,
} from './meeting-app-evidence-utils.mjs';

const args = parseCliArgs();
const baseUrl = String(args.get('base-url') || args.get('baseUrl') || 'http://localhost:8787');
const evidenceDir = String(args.get('evidence-dir') || args.get('evidenceDir') || 'data');
const outDir = resolve(String(args.get('out-dir') || args.get('outDir') || args.get('plan-dir') || args.get('planDir') || 'data/meeting-platform-field-intake-plans'));
const reportFile = String(args.get('report-file') || args.get('write-report') || '');
const jsonOutput = args.get('json') === 'true';
const writePlans = args.get('write-plans') !== 'false';
const failOnBlocked = args.get('fail-on-blocked') === 'true';
const includePlans = args.get('include-plans') === 'true';
const requiredPlatforms = unique(String(
  args.get('required-platforms')
    || args.get('platforms')
    || 'google-meet,teams,zoom,webex,lark',
).split(',').map((item) => item.trim()));

async function writeJson(file, value) {
  await mkdir(dirname(file), { recursive: true });
  await writeFile(file, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function planFile(platform) {
  return resolve(outDir, `${platform}.json`);
}

function summarizePlan(plan = {}) {
  return {
    platform: plan.platform,
    display_name: plan.display_name,
    status: plan.status,
    production_ready: plan.production_ready,
    ready_for_realtime_annotations: plan.ready_for_realtime_annotations,
    provider_endpoint: plan.provider_endpoint,
    plan_file: planFile(plan.platform),
    missing_env: plan.provider_connection?.security?.missing_env ?? [],
    required_provider_coverage: plan.acceptance?.required_provider_coverage ?? [],
    required_local_snapshots: plan.acceptance?.required_local_snapshots ?? [],
    field_evidence_input: plan.files?.field_evidence_input,
    evidence_package: plan.files?.evidence_package,
    validate_real_intake_command: plan.commands?.validate_real_intake,
    next_actions: plan.next_actions ?? [],
  };
}

async function buildReport() {
  const matrix = buildMeetingPlatformFieldIntakeMatrix({
    baseUrl,
    env: process.env,
    evidenceDir,
    platforms: requiredPlatforms,
  });
  const writtenFiles = [];
  if (writePlans) {
    for (const plan of matrix.plans) {
      const file = planFile(plan.platform);
      await writeJson(file, plan);
      writtenFiles.push(file);
    }
  }
  const rows = matrix.plans.map((plan) => summarizePlan(plan));
  const blockedCount = rows.filter((row) => String(row.status).startsWith('provider_setup')).length;
  const ok = matrix.platform_count > 0 && (!failOnBlocked || blockedCount === 0);
  return {
    type: 'meeting_platform_field_intake_report',
    ok,
    requirement: failOnBlocked ? 'no_provider_setup_blockers' : 'field_intake_plan_exported',
    base_url: baseUrl,
    evidence_dir: evidenceDir,
    out_dir: writePlans ? outDir : undefined,
    write_plans: writePlans,
    platform_count: matrix.platform_count,
    production_ready_count: matrix.production_ready_count,
    realtime_ready_count: matrix.realtime_ready_count,
    provider_blocked_count: matrix.provider_blocked_count,
    local_capture_needed_count: matrix.local_capture_needed_count,
    provider_capture_needed_count: matrix.provider_capture_needed_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows,
    matrix: includePlans ? matrix : undefined,
    next_actions: matrix.next_actions ?? [],
  };
}

try {
  const report = await buildReport();
  if (reportFile) await writeJson(resolve(reportFile), report);
  if (jsonOutput) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`meeting_platform_field_intake_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | production_ready=${report.production_ready_count} | realtime_ready=${report.realtime_ready_count} | provider_blocked=${report.provider_blocked_count} | written=${report.written_files.length}`);
    for (const row of report.rows) {
      console.log(`${row.platform}: status=${row.status} production_ready=${boolLabel(row.production_ready)} realtime_ready=${boolLabel(row.ready_for_realtime_annotations)} missing_env=${row.missing_env.length} plan=${row.plan_file}`);
    }
    if (report.next_actions.length > 0) console.log(`next_actions=${report.next_actions.join(',')}`);
  }
  if (!report.ok && failOnBlocked) process.exitCode = 2;
} catch (error) {
  console.error(error?.stack ?? error);
  process.exitCode = 1;
}
