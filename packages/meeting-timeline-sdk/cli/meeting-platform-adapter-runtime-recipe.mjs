import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { basename, dirname, resolve } from 'node:path';
import {
  buildMeetingPlatformAdapterRuntimeRecipe,
  buildMeetingPlatformAdapterRuntimeRecipeMatrix,
} from '../adapters/platform-adapter-runtime-recipe.mjs';

function boolLabel(value) {
  return value ? 'yes' : 'no';
}

function compactUndefined(value = {}) {
  return Object.fromEntries(Object.entries(value).filter(([, item]) => item !== undefined));
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

function compactExamples(examples = {}) {
  return compactUndefined({
    schema: examples.schema,
    signal_count: examples.signal_count,
    runtime_event_count: examples.runtime_event_count,
    filtered_speaker_event_count: examples.filtered_speaker_event_count,
    runtime_actions: unique(examples.runtime_events?.map((event) => event.action) ?? examples.runtime_actions ?? []),
  });
}

function compactRecipe(recipe = {}) {
  if (!recipe || typeof recipe !== 'object') return recipe;
  return compactUndefined({
    type: recipe.type,
    schema: recipe.schema,
    schema_version: recipe.schema_version,
    accepted: recipe.accepted,
    runtime_ready: recipe.runtime_ready,
    status: recipe.status,
    platform: recipe.platform,
    display_name: recipe.display_name,
    selected_surface: recipe.selected_surface,
    host_profile: recipe.host_profile,
    install_target: recipe.install_target,
    runtime_contract: recipe.runtime_contract,
    host_wiring: recipe.host_wiring,
    sequence: recipe.sequence,
    raw_signal_examples: compactExamples(recipe.raw_signal_examples),
    readiness: recipe.readiness,
    issues: recipe.issues,
    next_actions: recipe.next_actions,
  });
}

function compactMatrix(matrix = {}, options = {}) {
  const includeRecipes = options.includeRecipes === true;
  return compactUndefined({
    ...matrix,
    recipes: includeRecipes ? (matrix.recipes ?? []).map((recipe) => compactRecipe(recipe)) : undefined,
  });
}

export function parseMeetingPlatformAdapterRuntimeRecipeCliArgs(argv = process.argv.slice(2)) {
  const args = new Map();
  for (const raw of argv) {
    const [key, ...rest] = raw.replace(/^--/, '').split('=');
    args.set(key, rest.length ? rest.join('=') : 'true');
  }
  return args;
}

export function meetingPlatformAdapterRuntimeRecipeCliOptionsFromArgs(args = new Map()) {
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
    jsonOutput: boolArg(args.get('json'), false),
    includeRecipes: boolArg(args.get('include-recipes'), false),
    includeExamples: boolArg(args.get('include-examples'), false),
    writeRecipes: boolArg(args.get('write-recipes'), true),
    failOnBlocked: boolArg(args.get('fail-on-blocked'), false) || boolArg(args.get('fail-on-failed'), false),
    url: String(args.get('url') || args.get('href') || args.get('meeting-url') || ''),
    title: String(args.get('title') || ''),
    platform: explicitPlatform,
    surface: String(args.get('surface') || args.get('preferred-surface') || ''),
    hostProfile: String(args.get('host-profile') || args.get('hostProfile') || ''),
    filterActiveSpeakerSamples: boolArg(args.get('filter-active-speaker-samples'), true),
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

function recipeFile(outDir, platform) {
  return outDir ? resolve(outDir, `${platform}.json`) : undefined;
}

async function writeRuntimeRecipes(outDir, matrix = {}, options = {}) {
  if (!outDir || options.writeRecipes === false) return [];
  const writtenFiles = [];
  for (const recipe of matrix.recipes ?? []) {
    const file = recipeFile(outDir, recipe.platform);
    await writeJson(file, compactRecipe(recipe));
    writtenFiles.push(file);
  }
  return writtenFiles;
}

async function recipeInput(options = {}) {
  const fromFile = options.inputFile ? await readJson(options.inputFile) : {};
  return compactUndefined({
    ...fromFile,
    url: options.url || fromFile.url,
    title: options.title || fromFile.title,
    platform: options.platform || fromFile.platform,
    surface: options.surface || fromFile.surface,
    hostProfile: options.hostProfile || fromFile.hostProfile || fromFile.host_profile,
  });
}

export async function buildMeetingPlatformAdapterRuntimeRecipeCliReport(options = {}) {
  const {
    baseUrl = 'http://localhost:8787',
    outDir = '',
    outFile = '',
    reportFile = '',
    includeRecipes = false,
    includeExamples = false,
    writeRecipes = true,
    requiredPlatforms = ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  } = options;
  const input = await recipeInput(options);
  const matrix = buildMeetingPlatformAdapterRuntimeRecipeMatrix(input, {
    ...options,
    baseUrl,
    platforms: requiredPlatforms,
    includeExamples,
  });
  const singleRecipe = options.platform || options.url || options.inputFile
    ? buildMeetingPlatformAdapterRuntimeRecipe(input, {
      ...options,
      baseUrl,
      includeExamples,
    })
    : undefined;
  const writtenFiles = await writeRuntimeRecipes(outDir, matrix, { writeRecipes });
  if (outFile) await writeJson(resolve(outFile), singleRecipe ? compactRecipe(singleRecipe) : compactMatrix(matrix, { includeRecipes }));
  const report = {
    type: 'meeting_platform_adapter_runtime_recipe_report',
    ok: matrix.platform_count > 0 && matrix.runtime_ready_count === matrix.platform_count,
    base_url: baseUrl,
    out_dir: outDir || undefined,
    out_file: outFile || undefined,
    write_recipes: writeRecipes,
    include_examples: includeExamples,
    platform_count: matrix.platform_count,
    accepted_count: matrix.accepted_count,
    runtime_ready_count: matrix.runtime_ready_count,
    browser_surface_count: matrix.browser_surface_count,
    native_surface_count: matrix.native_surface_count,
    provider_reconcile_surface_count: matrix.provider_reconcile_surface_count,
    raw_signal_runtime_event_count: matrix.raw_signal_runtime_event_count,
    required_platforms: requiredPlatforms,
    written_files: writtenFiles,
    rows: matrix.rows.map((row) => ({
      ...row,
      recipe_file: recipeFile(outDir, row.platform),
    })),
    single_recipe: singleRecipe ? compactRecipe(singleRecipe) : undefined,
    matrix: compactMatrix(matrix, { includeRecipes }),
    next_actions: matrix.next_actions,
  };
  if (reportFile) await writeJson(resolve(reportFile), report);
  return report;
}

export function formatMeetingPlatformAdapterRuntimeRecipeCliReport(report = {}) {
  const lines = [
    `meeting_platform_adapter_runtime_recipe_report | ok=${boolLabel(report.ok)} | platforms=${report.platform_count} | accepted=${report.accepted_count} | runtime_ready=${report.runtime_ready_count} | browser=${report.browser_surface_count} | native=${report.native_surface_count} | provider_only=${report.provider_reconcile_surface_count} | runtime_events=${report.raw_signal_runtime_event_count} | written=${report.written_files?.length ?? 0}`,
  ];
  for (const row of report.rows ?? []) {
    lines.push(`${row.platform}: accepted=${boolLabel(row.accepted)} runtime=${boolLabel(row.runtime_ready)} surface=${row.selected_surface ?? 'n/a'} bridge=${row.bridge_kind ?? 'n/a'} first=${row.first_required_method ?? 'n/a'} insert=${row.insert_method ?? 'n/a'} speaker=${boolLabel(row.speaker_track_sample_ready)} recipe=${row.recipe_file ?? 'n/a'}`);
  }
  if (report.single_recipe?.platform) {
    lines.push(`single_recipe=${report.single_recipe.platform}:${report.single_recipe.selected_surface}:${report.single_recipe.status}`);
  }
  if (report.next_actions?.length > 0) lines.push(`next_actions=${report.next_actions.join(',')}`);
  if (report.out_dir) lines.push(`out_dir=${basename(report.out_dir)}`);
  if (report.out_file) lines.push(`out_file=${basename(report.out_file)}`);
  return lines.join('\n');
}

export async function runMeetingPlatformAdapterRuntimeRecipeCli(argv = process.argv.slice(2), io = console) {
  const options = meetingPlatformAdapterRuntimeRecipeCliOptionsFromArgs(
    parseMeetingPlatformAdapterRuntimeRecipeCliArgs(argv),
  );
  const report = await buildMeetingPlatformAdapterRuntimeRecipeCliReport(options);
  if (options.jsonOutput) {
    io.log(JSON.stringify(report, null, 2));
  } else {
    io.log(formatMeetingPlatformAdapterRuntimeRecipeCliReport(report));
  }
  if (!report.ok && options.failOnBlocked) process.exitCode = 2;
  return report;
}
