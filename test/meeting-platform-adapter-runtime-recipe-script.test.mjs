import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-adapter-runtime-recipe-script-'));
const outDir = join(tmpDir, 'runtime-recipes');
const reportFile = join(tmpDir, 'adapter-runtime-recipe-report.json');
const outFile = join(tmpDir, 'google-runtime-recipe.json');
const baseUrl = 'https://timeline.example.com';

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-recipe.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom',
  `--out-dir=${outDir}`,
  `--report-file=${reportFile}`,
  '--json=true',
  '--include-recipes=true',
], {
  cwd: repoRoot,
});

const report = JSON.parse(stdout);
assert.equal(report.type, 'meeting_platform_adapter_runtime_recipe_report');
assert.equal(report.ok, true);
assert.equal(report.platform_count, 3);
assert.equal(report.accepted_count, 3);
assert.equal(report.runtime_ready_count, 3);
assert.equal(report.browser_surface_count, 1);
assert.equal(report.native_surface_count, 2);
assert.equal(report.provider_reconcile_surface_count, 0);
assert.equal(report.raw_signal_runtime_event_count, 12);
assert.equal(report.written_files.length, 3);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').bridge_kind, 'browser_content_script');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').first_required_method, 'observePlatformCandidates');
assert.equal(report.rows.find((row) => row.platform === 'google_meet').insert_method, 'insertAnnotation');
assert.equal(report.rows.find((row) => row.platform === 'zoom').bridge_kind, 'native_detector_runtime_event_client');
assert.equal(report.rows.find((row) => row.platform === 'zoom').speaker_track_sample_ready, true);
assert.equal(report.matrix.recipes.length, 3);
assert.equal(report.matrix.recipes.some((recipe) => recipe.startup_plan), false);
assert.equal(report.matrix.recipes.some((recipe) => recipe.raw_signal_examples?.runtime_events), false);

const writtenReport = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(writtenReport.rows.length, 3);
assert.equal(writtenReport.rows.find((row) => row.platform === 'microsoft_teams').selected_surface, 'native_detector');

const googleRecipe = JSON.parse(await readFile(join(outDir, 'google_meet.json'), 'utf8'));
assert.equal(googleRecipe.schema, 'meeting_platform_adapter_runtime_recipe');
assert.equal(googleRecipe.platform, 'google_meet');
assert.equal(googleRecipe.selected_surface, 'browser_extension');
assert.equal(googleRecipe.host_wiring.bridge_kind, 'browser_content_script');
assert.equal(googleRecipe.host_wiring.insert_method, 'insertAnnotation');
assert.equal(googleRecipe.sequence.find((step) => step.action === 'insert_realtime_annotation').timestamp_field, 'captured_at_ms');
assert.equal(googleRecipe.raw_signal_examples.runtime_event_count, 4);

const zoomRecipe = JSON.parse(await readFile(join(outDir, 'zoom.json'), 'utf8'));
assert.equal(zoomRecipe.selected_surface, 'native_detector');
assert.equal(zoomRecipe.host_wiring.bridge_kind, 'native_detector_runtime_event_client');

const { stdout: singleStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-recipe.mjs',
  `--base-url=${baseUrl}`,
  '--platform=google-meet',
  '--url=https://meet.google.com/abc-defg-hij',
  `--out-file=${outFile}`,
  '--json=true',
  '--write-recipes=false',
], {
  cwd: repoRoot,
});
const singleReport = JSON.parse(singleStdout);
assert.equal(singleReport.platform_count, 1);
assert.equal(singleReport.single_recipe.platform, 'google_meet');
assert.equal(singleReport.single_recipe.runtime_ready, true);
assert.equal(singleReport.single_recipe.host_wiring.bridge_kind, 'browser_content_script');
assert.equal(singleReport.written_files.length, 0);
const singleRecipeFile = JSON.parse(await readFile(outFile, 'utf8'));
assert.equal(singleRecipeFile.platform, 'google_meet');
assert.equal(singleRecipeFile.host_wiring.first_required_method, 'observePlatformCandidates');

const { stdout: textStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-adapter-runtime-recipe.mjs',
  '--platforms=webex',
  '--write-recipes=false',
], {
  cwd: repoRoot,
});
assert.match(textStdout, /meeting_platform_adapter_runtime_recipe_report/);
assert.match(textStdout, /platforms=1/);
assert.match(textStdout, /webex: accepted=yes/);

const { stdout: binStdout } = await execFileAsync(process.execPath, [
  'packages/meeting-timeline-sdk/bin/meeting-platform-adapter-runtime-recipe.mjs',
  '--platforms=google-meet,zoom',
  '--json=true',
  '--write-recipes=false',
], {
  cwd: repoRoot,
});
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_platform_adapter_runtime_recipe_report');
assert.equal(binReport.ok, true);
assert.equal(binReport.platform_count, 2);
assert.equal(binReport.rows.find((row) => row.platform === 'google_meet').bridge_kind, 'browser_content_script');
assert.equal(binReport.rows.find((row) => row.platform === 'zoom').bridge_kind, 'native_detector_runtime_event_client');
assert.equal(binReport.written_files.length, 0);

console.log('ok meeting platform adapter runtime recipe script');
