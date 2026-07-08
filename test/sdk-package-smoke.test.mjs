import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const sdkDir = new URL('../packages/meeting-timeline-sdk', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-timeline-sdk-package-'));
const packDir = join(tmpDir, 'pack');
const consumerDir = join(tmpDir, 'consumer');
await mkdir(packDir, { recursive: true });
await mkdir(consumerDir, { recursive: true });

const { stdout: packStdout } = await execFileAsync('npm', [
  'pack',
  '--json',
  '--pack-destination',
  packDir,
], {
  cwd: sdkDir,
});
const packInfo = JSON.parse(packStdout)[0];
assert.equal(packInfo.name, '@ai-annotation/meeting-timeline-sdk');
const packedFiles = packInfo.files.map((item) => item.path).sort();
assert.equal(packedFiles.includes('index.mjs'), true);
assert.equal(packedFiles.includes('index.d.ts'), true);
assert.equal(packedFiles.includes('bin/meeting-app-adapter-integration-package.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-app-connector-package.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-consumer-handoff.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-blueprint.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-export-package.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-import-plan.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-install-manifest.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-launch-plan.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-portfolio.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-preflight.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-smoke.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-adapter-startup.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-host-integration.mjs'), true);
assert.equal(packedFiles.includes('bin/meeting-platform-provider-replay.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-app-adapter-integration-package.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-app-connector-package.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-consumer-handoff.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-blueprint.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-export-package.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-import-plan.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-install-manifest.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-launch-plan.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-portfolio.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-preflight.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-smoke.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-adapter-startup.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-host-integration.mjs'), true);
assert.equal(packedFiles.includes('cli/meeting-platform-provider-replay.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-kit.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-kit.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-capability.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-capability.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-integration-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-integration-runtime.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-event.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-event.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-registry.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-registry.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-ingest.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-ingest.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-conformance.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-conformance.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-consumer-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-consumer-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-implementation-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-implementation-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-authoring.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-authoring.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-portfolio.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-portfolio.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-acceptance-checklist.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-acceptance-checklist.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-export-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-export-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-import-plan.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-import-plan.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-install-manifest.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-install-manifest.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-launch-plan.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-launch-plan.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-session.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-session.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-runner.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-runner.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-message-bridge.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-message-bridge.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-smoke.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-smoke.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-profile.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-profile.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-route.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-route.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-blueprint.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-blueprint.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-decision.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-decision.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-startup.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-startup.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-preflight.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-preflight.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-host-integration.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-host-integration.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-provider-connection.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-provider-connection.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-subscription-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-subscription-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-speaker-track.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-speaker-track.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-participant-track.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-participant-track.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-timeline-view.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-timeline-view.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-annotation-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-annotation-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-clock-sync.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-clock-sync.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-session-binding.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-session-binding.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-realtime-annotation.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-realtime-annotation.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-artifact-handoff.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-artifact-handoff.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adaptation-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adaptation-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-bundle.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-runtime-bundle.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-contract.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-contract.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-sample.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-adapter-sample.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-real-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-real-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-field-intake.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-field-intake.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-handoff-readiness.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-handoff-readiness.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-fixture-tracks.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-fixture-tracks.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-manifest.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-manifest.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-spec.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-spec.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-runtime-config.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-runtime-config.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-handoff-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-handoff-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-capability.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-capability.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-integration-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-adapter-integration-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-connector-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-connector-package.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-connector.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-connector.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-runtime.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-extension.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-extension.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-profile.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-profile.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-observer-scheduler.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-observer-scheduler.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host-verifier.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-platform-runtime-host-verifier.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-pipeline.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-pipeline.d.ts'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-runtime.mjs'), true);
assert.equal(packedFiles.includes('adapters/meeting-app-track-runtime.d.ts'), true);
assert.equal(packedFiles.includes('README.md'), true);
assert.equal(packedFiles.some((item) => item.startsWith('test/')), false);
assert.equal(packedFiles.some((item) => item.startsWith('scripts/')), false);

const connectorPackageDts = await readFile(
  new URL('../packages/meeting-timeline-sdk/adapters/meeting-app-connector-package.d.ts', import.meta.url),
  'utf8',
);
assert.match(connectorPackageDts, /interface MeetingAppTimelineConnectorAdapterPreflightContract/);
assert.match(connectorPackageDts, /interface MeetingAppTimelineLocalObserverContract/);
assert.match(connectorPackageDts, /interface MeetingAppTimelineLocalObserverRuntimeWiring/);
assert.match(connectorPackageDts, /rows: MeetingAppTimelineConnectorHostInstallChecklistRow\[\]/);
assert.match(connectorPackageDts, /rows: MeetingAppTimelineConnectorAdapterMatrixRow\[\]/);
assert.match(connectorPackageDts, /local_observer_contract\?: MeetingAppTimelineLocalObserverContract/);
assert.match(connectorPackageDts, /local_observer_runtime_wiring\?: MeetingAppTimelineLocalObserverRuntimeWiring/);

await writeFile(join(consumerDir, 'package.json'), JSON.stringify({
  type: 'module',
  private: true,
  dependencies: {
    '@ai-annotation/meeting-timeline-sdk': sdkDir.pathname,
  },
}, null, 2), 'utf8');

await execFileAsync('npm', [
  'install',
  '--ignore-scripts',
  '--no-audit',
  '--no-fund',
], {
  cwd: consumerDir,
});

const { stdout: binStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-app-adapter-integration-package'),
  [
    '--platforms=google-meet',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const binReport = JSON.parse(binStdout);
assert.equal(binReport.type, 'meeting_app_adapter_integration_package_report');
assert.equal(binReport.platform_count, 1);
assert.equal(binReport.required_platforms[0], 'google-meet');
assert.equal(binReport.runtime_ready_count, 1);

const { stdout: connectorBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-app-connector-package'),
  [
    '--platforms=google-meet',
    '--surfaces=browser-extension',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const connectorBinReport = JSON.parse(connectorBinStdout);
assert.equal(connectorBinReport.type, 'meeting_app_timeline_connector_package_report');
assert.equal(connectorBinReport.platform_count, 1);
assert.deepEqual(connectorBinReport.surfaces, ['browser_extension']);
assert.equal(connectorBinReport.package.schema, 'meeting_app_timeline_connector_package');
assert.equal(connectorBinReport.package.provider_replay.accepted, true);

const { stdout: consumerHandoffBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-consumer-handoff'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const consumerHandoffBinReport = JSON.parse(consumerHandoffBinStdout);
assert.equal(consumerHandoffBinReport.schema, 'meeting_platform_consumer_handoff');
assert.equal(consumerHandoffBinReport.accepted, true);
assert.equal(consumerHandoffBinReport.platform_count, 2);
assert.equal(consumerHandoffBinReport.adapter_startup_ready_count, 2);
assert.equal(consumerHandoffBinReport.rows.find((row) => row.platform === 'google_meet').adapter_startup_selected_surface, 'browser_extension');
assert.equal(consumerHandoffBinReport.rows.find((row) => row.platform === 'zoom').adapter_startup_selected_surface, 'native_detector');
assert.equal(consumerHandoffBinReport.entrypoints.commands.adapter_startup.includes('meeting-platform:adapter-startup'), true);

const { stdout: adapterExportBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-export-package'),
  [
    '--platforms=google-meet',
    '--target=static',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterExportBinReport = JSON.parse(adapterExportBinStdout);
assert.equal(adapterExportBinReport.type, 'meeting_platform_adapter_export_package_report');
assert.equal(adapterExportBinReport.platform_count, 1);
assert.equal(adapterExportBinReport.export_ready_count, 1);

const { stdout: adapterImportBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-import-plan'),
  [
    '--package-file=missing-adapter-export-package.json',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterImportBinReport = JSON.parse(adapterImportBinStdout);
assert.equal(adapterImportBinReport.type, 'meeting_platform_adapter_import_plan_report');
assert.equal(adapterImportBinReport.package_count, 0);

const { stdout: adapterInstallBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-install-manifest'),
  [
    '--plan-file=missing-adapter-import-plan.json',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterInstallBinReport = JSON.parse(adapterInstallBinStdout);
assert.equal(adapterInstallBinReport.type, 'meeting_platform_adapter_install_manifest_report');
assert.equal(adapterInstallBinReport.plan_count, 0);
assert.equal(adapterInstallBinReport.ok, false);

const { stdout: adapterLaunchBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-launch-plan'),
  [
    '--manifest-file=missing-adapter-install-manifest.json',
    '--platform=google-meet',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterLaunchBinReport = JSON.parse(adapterLaunchBinStdout);
assert.equal(adapterLaunchBinReport.type, 'meeting_platform_adapter_launch_plan_report');
assert.equal(adapterLaunchBinReport.ok, false);

const { stdout: adapterPortfolioBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-portfolio'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
    '--write-items=false',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterPortfolioBinReport = JSON.parse(adapterPortfolioBinStdout);
assert.equal(adapterPortfolioBinReport.type, 'meeting_platform_adapter_portfolio_report');
assert.equal(adapterPortfolioBinReport.ok, true);
assert.equal(adapterPortfolioBinReport.platform_count, 2);
assert.equal(adapterPortfolioBinReport.built_in_count, 2);
assert.equal(adapterPortfolioBinReport.rows.find((row) => row.platform === 'google_meet').recommended_first_surface, 'browser_extension');
assert.equal(adapterPortfolioBinReport.rows.find((row) => row.platform === 'zoom').recommended_first_surface, 'native_detector');
assert.equal(adapterPortfolioBinReport.written_files.length, 0);

const { stdout: adapterStartupBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-startup'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
    '--write-plans=false',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterStartupBinReport = JSON.parse(adapterStartupBinStdout);
assert.equal(adapterStartupBinReport.type, 'meeting_platform_adapter_startup_plan_report');
assert.equal(adapterStartupBinReport.ok, true);
assert.equal(adapterStartupBinReport.platform_count, 2);
assert.equal(adapterStartupBinReport.realtime_startup_ready_count, 2);
assert.equal(adapterStartupBinReport.rows.find((row) => row.platform === 'google_meet').selected_surface, 'browser_extension');
assert.equal(adapterStartupBinReport.rows.find((row) => row.platform === 'zoom').selected_surface, 'native_detector');
assert.equal(adapterStartupBinReport.written_files.length, 0);

const { stdout: adapterPreflightBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-preflight'),
  [
    '--mode=single',
    '--platform=google-meet',
    '--url=https://meet.google.com/abc-defg-hij',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterPreflightBinReport = JSON.parse(adapterPreflightBinStdout);
assert.equal(adapterPreflightBinReport.type, 'meeting_platform_adapter_preflight_report');
assert.equal(adapterPreflightBinReport.ok, false);
assert.equal(adapterPreflightBinReport.mode, 'single');
assert.equal(adapterPreflightBinReport.platform, 'google_meet');
assert.equal(adapterPreflightBinReport.status, 'needs_live_page_evidence');
assert.equal(adapterPreflightBinReport.selected_surface, 'browser_extension');

const { stdout: adapterSmokeBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-adapter-smoke'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const adapterSmokeBinReport = JSON.parse(adapterSmokeBinStdout);
assert.equal(adapterSmokeBinReport.type, 'meeting_platform_adapter_smoke_cli_report');
assert.equal(adapterSmokeBinReport.ok, true);
assert.equal(adapterSmokeBinReport.accepted_count, 2);

const { stdout: hostIntegrationBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-host-integration'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
    '--write-scaffold=false',
  ],
  {
    cwd: consumerDir,
  },
);
const hostIntegrationBinReport = JSON.parse(hostIntegrationBinStdout);
assert.equal(hostIntegrationBinReport.type, 'meeting_platform_host_integration_scaffold_report');
assert.equal(hostIntegrationBinReport.ok, true);
assert.equal(hostIntegrationBinReport.platform_count, 2);
assert.equal(hostIntegrationBinReport.adapter_runtime_ready, true);
assert.equal(hostIntegrationBinReport.adapter_startup_ready, true);
assert.equal(hostIntegrationBinReport.adapter_startup_ready_count, 2);
assert.equal(hostIntegrationBinReport.adapter_preflight_available, true);
assert.equal(hostIntegrationBinReport.adapter_preflight_platform_count, 2);
assert.equal(hostIntegrationBinReport.generated_files.some((file) => file.path === 'src/platform-adapters/google_meet.mjs'), true);
assert.equal(hostIntegrationBinReport.generated_files.some((file) => file.path === 'scripts/print-adapter-startup.mjs'), true);
assert.equal(hostIntegrationBinReport.generated_files.some((file) => file.path === 'scripts/print-adapter-preflight.mjs'), true);
assert.equal(hostIntegrationBinReport.written_file_count, 0);

const { stdout: providerReplayBinStdout } = await execFileAsync(
  join(consumerDir, 'node_modules', '.bin', 'meeting-platform-provider-replay'),
  [
    '--platforms=google-meet,zoom',
    '--json=true',
  ],
  {
    cwd: consumerDir,
  },
);
const providerReplayBinReport = JSON.parse(providerReplayBinStdout);
assert.equal(providerReplayBinReport.type, 'meeting_platform_provider_replay_cli_report');
assert.equal(providerReplayBinReport.ok, true);
assert.equal(providerReplayBinReport.accepted_count, 2);
assert.equal(providerReplayBinReport.rows.every((row) => row.provider_events_block_realtime === false), true);

await writeFile(join(consumerDir, 'smoke.mjs'), `
import assert from 'node:assert/strict';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA as ROOT_MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  SDK_VERSION,
  buildMeetingAppAdapterCapabilityMatrix as buildMeetingAppAdapterCapabilityMatrixFromRoot,
  buildMeetingAppAdapterExecutionPlanMatrix as buildMeetingAppAdapterExecutionPlanMatrixFromRoot,
  buildMeetingAppAdapterIntegrationPackageMatrix as buildMeetingAppAdapterIntegrationPackageMatrixFromRoot,
  buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport as buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReportFromRoot,
  buildMeetingAppTimelineConnectorBridgeHandoff as buildMeetingAppTimelineConnectorBridgeHandoffFromRoot,
  buildMeetingAppTimelineConnectorAdoptionIndex as buildMeetingAppTimelineConnectorAdoptionIndexFromRoot,
  buildMeetingAppTimelineConnectorAdapterMatrix as buildMeetingAppTimelineConnectorAdapterMatrixFromRoot,
  buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport as buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReportFromRoot,
  buildMeetingAppTimelineConnectorFieldIntakeIndex as buildMeetingAppTimelineConnectorFieldIntakeIndexFromRoot,
  buildMeetingAppTimelineHostAdapterBootstrapPlan as buildMeetingAppTimelineHostAdapterBootstrapPlanFromRoot,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix as buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixFromRoot,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport as buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReportFromRoot,
  buildMeetingAppTimelineHostAdapterConfig as buildMeetingAppTimelineHostAdapterConfigFromRoot,
  buildMeetingAppTimelineHostAdapterConfigIndex as buildMeetingAppTimelineHostAdapterConfigIndexFromRoot,
  resolveMeetingAppTimelineHostAdapterConfig as resolveMeetingAppTimelineHostAdapterConfigFromRoot,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport as buildMeetingAppTimelineConnectorPackageAcceptanceReportFromRoot,
  buildMeetingAppTimelineConnectorPlatformRoadmap as buildMeetingAppTimelineConnectorPlatformRoadmapFromRoot,
  buildMeetingAppTimelineConnectorReleaseGate as buildMeetingAppTimelineConnectorReleaseGateFromRoot,
  buildMeetingAppTimelineConnectorSmokePlan as buildMeetingAppTimelineConnectorSmokePlanFromRoot,
  buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport as buildMeetingAppTimelineConnectorSmokePlanAcceptanceReportFromRoot,
  buildMeetingPlatformProviderReplayMatrix as buildMeetingPlatformProviderReplayMatrixFromRoot,
  buildMeetingPlatformProviderReplayReport as buildMeetingPlatformProviderReplayReportFromRoot,
  sampleMeetingPlatformProviderEvents as sampleMeetingPlatformProviderEventsFromRoot,
  runMeetingAppTimelineConnectorBridgeSmoke as runMeetingAppTimelineConnectorBridgeSmokeFromRoot,
  runMeetingAppTimelineConnectorSmokePlan as runMeetingAppTimelineConnectorSmokePlanFromRoot,
  buildMeetingPlatformAdaptationPackage as buildMeetingPlatformAdaptationPackageFromRoot,
  buildMeetingPlatformAdaptationStrategy as buildMeetingPlatformAdaptationStrategyFromRoot,
  buildMeetingPlatformAdapterAuthoringPlan as buildMeetingPlatformAdapterAuthoringPlanFromRoot,
  buildMeetingPlatformAdapterAcceptanceChecklist as buildMeetingPlatformAdapterAcceptanceChecklistFromRoot,
  buildMeetingPlatformAdapterExportPackage as buildMeetingPlatformAdapterExportPackageFromRoot,
  buildMeetingPlatformAdapterImportPlan as buildMeetingPlatformAdapterImportPlanFromRoot,
  buildMeetingPlatformAdapterInstallManifest as buildMeetingPlatformAdapterInstallManifestFromRoot,
  buildMeetingPlatformAdapterCandidateLaunchPlan as buildMeetingPlatformAdapterCandidateLaunchPlanFromRoot,
  buildMeetingPlatformAdapterLaunchPlan as buildMeetingPlatformAdapterLaunchPlanFromRoot,
  buildMeetingPlatformAdapterPortfolio as buildMeetingPlatformAdapterPortfolioFromRoot,
  buildMeetingPlatformAdapterDecision as buildMeetingPlatformAdapterDecisionFromRoot,
  buildMeetingPlatformAdapterDecisionMatrix as buildMeetingPlatformAdapterDecisionMatrixFromRoot,
  buildMeetingPlatformAdapterStartupPlan as buildMeetingPlatformAdapterStartupPlanFromRoot,
  buildMeetingPlatformAdapterStartupPlanMatrix as buildMeetingPlatformAdapterStartupPlanMatrixFromRoot,
  buildMeetingPlatformAdapterCandidatePreflight as buildMeetingPlatformAdapterCandidatePreflightFromRoot,
  buildMeetingPlatformAdapterCurrentWindowPreflight as buildMeetingPlatformAdapterCurrentWindowPreflightFromRoot,
  buildMeetingPlatformAdapterPreflight as buildMeetingPlatformAdapterPreflightFromRoot,
  buildMeetingPlatformAdapterPreflightMatrix as buildMeetingPlatformAdapterPreflightMatrixFromRoot,
  buildMeetingPlatformAdapterRoute as buildMeetingPlatformAdapterRouteFromRoot,
  buildMeetingPlatformAdapterMessageBridgeHandoff as buildMeetingPlatformAdapterMessageBridgeHandoffFromRoot,
  buildMeetingPlatformAdapterRunnerHandoff as buildMeetingPlatformAdapterRunnerHandoffFromRoot,
  buildMeetingPlatformAdapterSessionHandoff as buildMeetingPlatformAdapterSessionHandoffFromRoot,
  assertMeetingPlatformAdapterSmoke as assertMeetingPlatformAdapterSmokeFromRoot,
  buildMeetingPlatformRuntimeProfile as buildMeetingPlatformRuntimeProfileFromRoot,
  buildMeetingPlatformConsumerHandoff as buildMeetingPlatformConsumerHandoffFromRoot,
  buildMeetingPlatformConnector as buildMeetingPlatformConnectorFromRoot,
  buildMeetingPlatformConnectorHub as buildMeetingPlatformConnectorHubFromRoot,
  buildMeetingPlatformImplementationHandoff as buildMeetingPlatformImplementationHandoffFromRoot,
  buildMeetingPlatformIntegrationRuntimeManifest as buildMeetingPlatformIntegrationRuntimeManifestFromRoot,
  buildMeetingPlatformRuntimeBundle as buildMeetingPlatformRuntimeBundleFromRoot,
  createMeetingAppTimelineSdk,
  createMeetingAppTimelineConnectorRuntimeClient as createMeetingAppTimelineConnectorRuntimeClientFromRoot,
  createMeetingPlatformConnectorBrowserRuntime as createMeetingPlatformConnectorBrowserRuntimeFromRoot,
  createMeetingPlatformConnectorContentScriptBridge as createMeetingPlatformConnectorContentScriptBridgeFromRoot,
  createMeetingPlatformConnectorHub as createMeetingPlatformConnectorHubFromRoot,
  createMeetingPlatformConnectorRuntime as createMeetingPlatformConnectorRuntimeFromRoot,
  createMeetingPlatformAdapterSession as createMeetingPlatformAdapterSessionFromRoot,
  createMeetingPlatformAdapterRunner as createMeetingPlatformAdapterRunnerFromRoot,
  createMeetingPlatformAdapterMessageBridge as createMeetingPlatformAdapterMessageBridgeFromRoot,
  openMeetingPlatformAdapterSession as openMeetingPlatformAdapterSessionFromRoot,
  runMeetingPlatformAdapterSmoke as runMeetingPlatformAdapterSmokeFromRoot,
  createMeetingTimelineClient,
  createMeetingPlatformTimelineKit as createMeetingPlatformTimelineKitFromRoot,
  detectMeetingPlatformForBrowser as detectMeetingPlatformForBrowserFromRoot,
} from '@ai-annotation/meeting-timeline-sdk';
import {
  createMeetingPlatformTimelineKit,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit';
import {
  assertMeetingPlatformIntegrationRuntimeManifest,
  buildMeetingPlatformIntegrationRuntimeManifest,
  createMeetingPlatformIntegrationBrowserRuntime,
  createMeetingPlatformIntegrationContentScriptBridge,
	  createMeetingPlatformIntegrationRuntime,
	  detectMeetingPlatformForBrowser,
	  installMeetingPlatformIntegrationContentScriptBridge,
	  resolveMeetingPlatformCandidates,
	  resolveMeetingPlatformForInput,
	  runMeetingPlatformIntegrationRuntimeManifest,
		} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';
import {
  assertMeetingPlatformProviderReplayMatrix,
  assertMeetingPlatformProviderReplayReport,
  buildMeetingPlatformProviderReplayMatrix,
  buildMeetingPlatformProviderReplayReport,
  sampleMeetingPlatformProviderEvents,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';
import {
  MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA,
  buildMeetingPlatformAnnotationRuntimeEvent,
  buildMeetingPlatformCandidateObservationRuntimeEvent,
  buildMeetingPlatformRunHandoffReadinessRuntimeEvent,
  buildMeetingPlatformRunManifestRuntimeEvent,
  buildMeetingPlatformRuntimeEventPlan,
  buildMeetingPlatformRuntimeEventPlanMatrix,
  createMeetingPlatformRuntimeEventClient,
  meetingPlatformRuntimeEventEndpoint,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event';
import {
  assertMeetingPlatformRegistryManifest,
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryEntry,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry';
import {
  assertMeetingPlatformConformanceReport,
  buildMeetingPlatformConformanceReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-conformance';
import {
  assertMeetingPlatformConsumerHandoff,
  buildMeetingPlatformConsumerHandoff,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff';
import {
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformRolloutPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';
import {
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy';
import {
  buildMeetingPlatformAdapterRoute,
  buildMeetingPlatformAdapterRouteMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route';
import {
  buildMeetingPlatformAdapterDecision,
  buildMeetingPlatformAdapterDecisionMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-decision';
import {
  buildMeetingPlatformAdapterStartupPlan,
  buildMeetingPlatformAdapterStartupPlanMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-startup';
import {
  buildMeetingPlatformAdapterCandidatePreflight,
  buildMeetingPlatformAdapterCurrentWindowPreflight,
  buildMeetingPlatformAdapterPreflight,
  buildMeetingPlatformAdapterPreflightMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight';
import {
  buildMeetingAppFixtureSnapshot,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures';
import {
  buildMeetingPlatformEvidenceCorrelation,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-correlation';
import {
  createMeetingPlatformEvidenceSession,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session';
import {
  assertMeetingPlatformLiveAdapterReadiness,
  assertMeetingPlatformLiveAdapterReadinessMatrix,
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterHandoffBundle,
  buildMeetingPlatformLiveAdapterMatrix,
  buildMeetingPlatformLiveAdapterPlan,
  buildMeetingPlatformLiveAdapterReadiness,
  buildMeetingPlatformLiveAdapterReadinessMatrix,
  createMeetingPlatformLiveAdapter,
  createMeetingPlatformLiveAdapterSuite,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter';
import {
  buildMeetingPlatformEvidencePackage,
  verifyMeetingPlatformEvidencePackage,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package';
import {
  buildMeetingPlatformHostIntegrationPlan,
  buildMeetingPlatformHostIntegrationScaffold,
  assertMeetingPlatformHostIntegrationScaffold,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-host-integration';
import {
  buildMeetingPlatformProviderConnectionMatrix,
  buildMeetingPlatformProviderConnectionPack,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection';
import {
  buildMeetingPlatformSubscriptionHandoff,
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-subscription-handoff';
import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
  buildMeetingPlatformSpeakerTrackPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track';
import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
  buildMeetingPlatformParticipantTrackPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track';
import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
  buildMeetingPlatformTimelineViewPlan,
  zoomMeetingPlatformTimelineViewport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view';
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
  buildMeetingPlatformAnnotationIntakePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake';
import {
  buildMeetingPlatformClockSyncMatrix,
  buildMeetingPlatformClockSyncPlan,
  buildMeetingPlatformClockSyncReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync';
import {
  buildMeetingPlatformSessionBinding,
  buildMeetingPlatformSessionBindingMatrix,
  buildMeetingPlatformSessionBindingPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding';
import {
  buildMeetingPlatformRealtimeAnnotation,
  buildMeetingPlatformRealtimeAnnotationMatrix,
  buildMeetingPlatformRealtimeAnnotationPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation';
import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
  buildMeetingPlatformArtifactHandoffPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-artifact-handoff';
import {
  buildMeetingPlatformAdaptationPackage,
  buildMeetingPlatformAdaptationPackageMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package';
import {
  buildMeetingPlatformRuntimeProfile,
  buildMeetingPlatformRuntimeProfileMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-profile';
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle';
import {
  buildMeetingPlatformAdapterAuthoringMatrix,
  buildMeetingPlatformAdapterAuthoringPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-authoring';
import {
  buildMeetingPlatformAdapterPortfolio,
  buildMeetingPlatformAdapterPortfolioItem,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-portfolio';
import {
  buildMeetingPlatformAdapterAcceptanceChecklist,
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist';
import {
  buildMeetingPlatformAdapterExportPackage,
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-export-package';
import {
  assertMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-import-plan';
import {
  assertMeetingPlatformAdapterInstallManifest,
  buildMeetingPlatformAdapterInstallManifest,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-install-manifest';
import {
  assertMeetingPlatformAdapterLaunchPlan,
  buildMeetingPlatformAdapterCandidateLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-launch-plan';
import {
  buildMeetingPlatformAdapterSessionHandoff,
  createMeetingPlatformAdapterSession,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-session';
import {
  buildMeetingPlatformAdapterRunnerHandoff,
  createMeetingPlatformAdapterRunner,
  openMeetingPlatformAdapterSession,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-runner';
import {
  buildMeetingPlatformAdapterMessageBridgeHandoff,
  createMeetingPlatformAdapterMessageBridge,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-message-bridge';
import {
  assertMeetingPlatformAdapterSmoke,
  runMeetingPlatformAdapterSmoke,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-smoke';
import {
  assertMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractAcceptanceReport,
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract';
import {
  buildMeetingPlatformAdapterSamplePlan,
  runMeetingPlatformAdapterSample,
  runMeetingPlatformAdapterSampleMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-sample';
import {
  buildMeetingPlatformRealEvidenceIntakePlan,
  buildMeetingPlatformRealEvidenceIntakeReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-real-intake';
import {
  buildMeetingPlatformFieldIntakeMatrix,
  buildMeetingPlatformFieldIntakePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake';
import {
  buildMeetingPlatformHandoffReadiness,
  buildMeetingPlatformHandoffReadinessMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness';
import {
  buildMeetingPlatformFieldCaptureManifest,
  buildMeetingPlatformFieldCollectorConfig,
  buildMeetingPlatformFieldEvidenceBundle,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';
import {
  normalizeGoogleMeetEvent,
} from '@ai-annotation/meeting-timeline-sdk/adapters/google-meet';
import {
  buildMeetingAppLaunchGate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate';
import {
  buildMeetingAppAdapterFitMatrix,
  buildMeetingAppAdapterFitReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps';
import {
  buildMeetingAppAdapterCapabilityMatrix,
  buildMeetingAppAdapterCapabilityReport,
  buildMeetingAppAdapterExecutionPlan,
  buildMeetingAppAdapterExecutionPlanMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-capability';
import {
  assertMeetingAppAdapterIntegrationPackage,
  assertMeetingAppAdapterIntegrationPackageMatrix,
  buildMeetingAppAdapterIntegrationPackage,
  buildMeetingAppAdapterIntegrationPackageMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package';
import {
  assertMeetingAppTimelineConnectorHostInstallChecklist,
  assertMeetingAppTimelineConnectorPackage,
  assertMeetingAppTimelineConnectorAdoptionIndex,
  assertMeetingAppTimelineConnectorAdapterMatrix,
  assertMeetingAppTimelineConnectorFieldIntakeIndex,
  assertMeetingAppTimelineHostAdapterBootstrapPlan,
  assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  assertMeetingAppTimelineHostAdapterConfig,
  assertMeetingAppTimelineHostAdapterConfigIndex,
  assertMeetingAppTimelineResolvedHostAdapterConfig,
  assertMeetingAppTimelineConnectorBridgeSmoke,
  assertMeetingAppTimelineConnectorPlatformRoadmap,
  assertMeetingAppTimelineConnectorReleaseGate,
  assertMeetingAppTimelineConnectorSmokePlan,
  assertMeetingAppTimelineConnectorSmokeRun,
  buildMeetingAppTimelineConnectorBridgeHandoff,
  buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport,
  buildMeetingAppTimelineConnectorAdoptionIndex,
  buildMeetingAppTimelineConnectorAdapterMatrix,
  buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport,
  buildMeetingAppTimelineConnectorFieldIntakeIndex,
  buildMeetingAppTimelineHostAdapterBootstrapPlan,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport,
  buildMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineHostAdapterConfigIndex,
  resolveMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineConnectorHostInstallChecklist,
  buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport,
  buildMeetingAppTimelineConnectorPlatformRoadmap,
  buildMeetingAppTimelineConnectorReleaseGate,
  buildMeetingAppTimelineConnectorSmokePlan,
  buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport,
  createMeetingAppTimelineConnectorRuntimeClient,
  runMeetingAppTimelineConnectorBridgeSmoke,
  runMeetingAppTimelineConnectorSmokePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-connector-package';
import {
  buildMeetingPlatformConnector,
  buildMeetingPlatformConnectorAcceptanceReport,
  buildMeetingPlatformConnectorHub,
  buildMeetingPlatformConnectorMatrix,
  createMeetingPlatformConnectorBrowserRuntime,
  createMeetingPlatformConnectorContentScriptBridge,
  createMeetingPlatformConnectorHub,
  createMeetingPlatformConnectorRuntime,
  resolveMeetingPlatformConnectorInput,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector';
import {
  assertMeetingAppRuntimeAdapterHandoff,
  buildMeetingAppDomAdaptationDiagnosisMatrix,
  buildMeetingAppRuntimeAdapterHandoffAcceptanceReport,
  buildMeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport,
  buildMeetingAppRuntimeAdapterHandoff,
  buildMeetingAppRuntimeAdapterHandoffMatrix,
  buildMeetingAppRuntimeAdapterHostPackage,
  buildMeetingAppRuntimeAdapterProfileMatrix,
  buildMeetingAppRuntimeObserverPlan,
  buildMeetingAppRuntimeObserverPlanMatrix,
  resolveMeetingAppRuntimeAdapterProfile,
  selectMeetingAppRuntimeAdapter,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';
import {
  assertMeetingAppAdapterManifest,
  assertMeetingAppAdapterManifestMatrix,
  buildMeetingAppAdapterManifest,
  buildMeetingAppAdapterManifestMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest';
import {
  assertMeetingAppAdapterSpec,
  assertMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpec,
  buildMeetingAppAdapterSpecMatrix,
  buildMeetingAppAdapterSpecTemplate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec';
import {
  assertMeetingAppAdapterRuntimeConfig,
  assertMeetingAppAdapterRuntimeConfigMatrix,
  buildMeetingAppAdapterRuntimeConfig,
  buildMeetingAppAdapterRuntimeConfigMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config';
import {
  assertMeetingAppAdapterHandoffPackage,
  assertMeetingAppAdapterHandoffPackageMatrix,
  assertMeetingAppAdapterVerificationReport,
  assertMeetingAppAdapterVerificationReportMatrix,
  buildMeetingAppAdapterHandoffPackage,
  buildMeetingAppAdapterHandoffPackageMatrix,
  buildMeetingAppAdapterVerificationReport,
  buildMeetingAppAdapterVerificationReportMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package';
import {
  MEETING_APP_EXTENSION_MESSAGE_TYPES,
  buildMeetingAppExtensionBackgroundSource,
  buildMeetingAppExtensionCandidateLaunchPlanMessage,
  buildMeetingAppExtensionCurrentWindowPreflightMessage,
  buildMeetingAppExtensionOpenCandidateSessionMessage,
  buildMeetingAppExtensionPreflightCandidatesMessage,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension';
import {
  buildMeetingAppFixtureTrackReadinessReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks';
import {
  buildMeetingAppTrackPipeline,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-pipeline';
import {
  createMeetingAppTrackRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime';
import {
  createMeetingAppTimelineRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime';
import {
  buildMeetingAppObserverSchedulerConfig,
  buildMeetingAppObserverSchedulerConfigMatrix,
  createMeetingAppObserverScheduler,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-observer-scheduler';
import {
  buildMeetingPlatformRuntimeHostConfig,
  buildMeetingPlatformRuntimeHostConfigMatrix,
  buildMeetingPlatformRuntimeHostHandoff,
  buildMeetingPlatformRuntimeHostHandoffMatrix,
  createMeetingPlatformRuntimeHost,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host';
import {
  createMeetingPlatformRuntimeHostFixtureEnvironment,
  createMeetingPlatformRuntimeHostVerificationClient,
  runMeetingPlatformRuntimeHostReplay,
  runMeetingPlatformRuntimeHostReplayMatrix,
  runMeetingPlatformRuntimeHostVerification,
  runMeetingPlatformRuntimeHostVerificationMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier';

assert.equal(SDK_VERSION, '0.1.0');
const client = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true }), {
    headers: { 'content-type': 'application/json' },
  }),
});
assert.equal(typeof client.startMeeting, 'function');
assert.equal(ROOT_MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA, 'meeting_platform_runtime_event');
assert.equal(createMeetingPlatformTimelineKitFromRoot(client, {
  baseUrl: 'http://localhost:8787',
  verify: false,
}).platformRegistryManifest({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(detectMeetingPlatformForBrowserFromRoot({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(buildMeetingPlatformIntegrationRuntimeManifestFromRoot({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(buildMeetingAppAdapterIntegrationPackageMatrixFromRoot({
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(buildMeetingAppAdapterCapabilityMatrixFromRoot({
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(buildMeetingAppAdapterExecutionPlanMatrixFromRoot({
  platforms: ['google-meet'],
}).accepted_count, 1);
assert.equal(buildMeetingAppAdapterCapabilityMatrix({
  platforms: ['google-meet'],
}).provider_axis_declared_count, 1);
assert.equal(buildMeetingAppAdapterExecutionPlanMatrix({
  platforms: ['google-meet'],
}).platform_count, 1);
const rootMeetingAppSdk = createMeetingAppTimelineSdk({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
  platforms: ['google-meet'],
});
const rootGoogleActiveSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_783_356_000_000,
});
function rootSmokeNode(tagName, attrs = {}, text = '') {
  return {
    tagName: tagName.toUpperCase(),
    attributes: attrs,
    dataset: Object.fromEntries(Object.entries(attrs)
      .filter(([key]) => key.startsWith('data-'))
      .map(([key, value]) => [
        key.slice(5).replace(/-([a-z])/g, (_, char) => char.toUpperCase()),
        value,
      ])),
    innerText: text,
    textContent: text,
    getAttribute(name) {
      return attrs[name] ?? null;
    },
  };
}
function rootSmokeDocument() {
  const nodes = [
    rootSmokeNode('button', { 'aria-label': 'Leave call' }),
    rootSmokeNode('div', {
      'data-participant-id': 'ada',
      'aria-label': 'Ada Lovelace is speaking',
    }),
  ];
  return {
    nodeType: 9,
    title: 'Package smoke - Google Meet',
    hidden: false,
    location: { href: 'https://meet.google.com/abc-defg-hij' },
    querySelectorAll(selector) {
      const text = String(selector);
      if (text === 'button') return nodes.filter((item) => item.tagName === 'BUTTON');
      if (text.includes('aria-label') && text.includes('Leave call')) {
        return nodes.filter((item) => String(item.attributes['aria-label'] ?? '').toLowerCase().includes('leave call'));
      }
      if (text.includes('aria-label') && text.includes('speaking')) {
        return nodes.filter((item) => String(item.attributes['aria-label'] ?? '').toLowerCase().includes('speaking'));
      }
      if (text.includes('data-participant-id')) {
        return nodes.filter((item) => item.attributes['data-participant-id']);
      }
      return [];
    },
  };
}
assert.equal(rootMeetingAppSdk.schema, 'meeting_app_timeline_sdk');
assert.equal(rootMeetingAppSdk.hostPackage({ surfaces: ['browser-extension'] }).schema, 'meeting_app_runtime_adapter_host_package');
const rootConnectorPackage = rootMeetingAppSdk.connectorPackage({ surfaces: ['browser-extension'] });
assert.equal(rootConnectorPackage.schema, 'meeting_app_timeline_connector_package');
assert.equal(rootConnectorPackage.provider_replay.matrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(rootConnectorPackage.provider_replay.accepted, true);
assert.equal(buildMeetingAppTimelineConnectorPackageAcceptanceReportFromRoot(rootConnectorPackage).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorPackageAcceptanceReport(rootConnectorPackage).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorHandoff(rootConnectorPackage).schema, 'meeting_app_timeline_connector_handoff');
assert.equal(buildMeetingAppTimelineConnectorBridgeHandoff(rootConnectorPackage).schema, 'meeting_app_timeline_connector_bridge_handoff');
assert.equal(buildMeetingAppTimelineConnectorBridgeHandoffFromRoot(rootConnectorPackage).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(rootConnectorPackage).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReportFromRoot(rootConnectorPackage).schema, 'meeting_app_timeline_connector_bridge_handoff_acceptance_report');
assert.equal(buildMeetingAppTimelineConnectorAdoptionIndex(rootConnectorPackage).schema, 'meeting_app_timeline_connector_adoption_index');
assert.equal(buildMeetingAppTimelineConnectorAdoptionIndexFromRoot(rootConnectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorAdoptionIndex(rootConnectorPackage).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorFieldIntakeIndex(rootConnectorPackage).schema, 'meeting_app_timeline_connector_field_intake_index');
assert.equal(buildMeetingAppTimelineConnectorFieldIntakeIndexFromRoot(rootConnectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorFieldIntakeIndex(rootConnectorPackage).accepted, true);
const rootConnectorReleaseGate = buildMeetingAppTimelineConnectorReleaseGate(rootConnectorPackage);
assert.equal(rootConnectorReleaseGate.schema, 'meeting_app_timeline_connector_release_gate');
assert.equal(buildMeetingAppTimelineConnectorReleaseGateFromRoot(rootConnectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorReleaseGate(rootConnectorPackage).accepted, true);
const rootConnectorHostInstallChecklist = buildMeetingAppTimelineConnectorHostInstallChecklist(rootConnectorPackage);
assert.equal(rootConnectorHostInstallChecklist.schema, 'meeting_app_timeline_connector_host_install_checklist');
assert.equal(buildMeetingAppTimelineConnectorHostInstallChecklistAcceptanceReport(rootConnectorHostInstallChecklist).accepted, true);
const rootConnectorConsumerHandoff = {
  schema: 'meeting_platform_consumer_handoff',
  accepted: true,
  adaptation_roadmap: {
    schema: 'meeting_platform_adaptation_roadmap',
    rows: rootConnectorHostInstallChecklist.rows.map((row) => ({
      platform: row.platform,
      display_name: row.display_name,
      priority_tier: 'pilot_ready',
      recommended_first_surface: row.selected_surface,
      surface_order: [row.selected_surface],
      next_phase: 'wire_host_runtime',
    })),
  },
  surface_coverage_matrix: {
    schema: 'meeting_platform_surface_coverage_matrix',
    rows: rootConnectorHostInstallChecklist.rows.map((row) => ({
      platform: row.platform,
      surface_order: [row.selected_surface],
      browser_extension: { ready: row.selected_surface === 'browser_extension' },
      webview_preload: { ready: row.selected_surface === 'webview_preload' },
      native_detector: { ready: row.selected_surface === 'native_detector' },
      provider_reconcile: {
        ready: true,
        provider_path: 'nonblocking_reconcile',
        permission_risk: 'low',
      },
      post_meeting_backfill: { supported: true },
      lightweight_connector: { ready: true },
      speaker_track: { ready: true },
      participant_track: { ready: true },
    })),
  },
};
const rootConnectorPlatformRoadmap = buildMeetingAppTimelineConnectorPlatformRoadmap(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  consumerHandoff: rootConnectorConsumerHandoff,
});
assert.equal(rootConnectorPlatformRoadmap.schema, 'meeting_app_timeline_connector_platform_roadmap');
assert.equal(buildMeetingAppTimelineConnectorPlatformRoadmapFromRoot(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  consumerHandoff: rootConnectorConsumerHandoff,
}).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorPlatformRoadmap(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  consumerHandoff: rootConnectorConsumerHandoff,
}).accepted, true);
const rootConnectorFieldIntakeIndex = buildMeetingAppTimelineConnectorFieldIntakeIndex(rootConnectorHostInstallChecklist);
const rootConnectorBridgeHandoff = buildMeetingAppTimelineConnectorBridgeHandoff(rootConnectorHostInstallChecklist);
const rootConnectorSmokePlan = buildMeetingAppTimelineConnectorSmokePlan(rootConnectorHostInstallChecklist);
const rootConnectorAdapterMatrix = buildMeetingAppTimelineConnectorAdapterMatrix(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  platformRoadmap: rootConnectorPlatformRoadmap,
  consumerHandoff: rootConnectorConsumerHandoff,
  fieldIntakeIndex: rootConnectorFieldIntakeIndex,
  bridgeHandoff: rootConnectorBridgeHandoff,
  smokePlan: rootConnectorSmokePlan,
});
assert.equal(rootConnectorAdapterMatrix.schema, 'meeting_app_timeline_connector_adapter_matrix');
assert.equal(buildMeetingAppTimelineConnectorAdapterMatrixFromRoot(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  platformRoadmap: rootConnectorPlatformRoadmap,
  consumerHandoff: rootConnectorConsumerHandoff,
  fieldIntakeIndex: rootConnectorFieldIntakeIndex,
  bridgeHandoff: rootConnectorBridgeHandoff,
  smokePlan: rootConnectorSmokePlan,
}).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport(rootConnectorAdapterMatrix).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReportFromRoot(rootConnectorAdapterMatrix).schema, 'meeting_app_timeline_connector_adapter_matrix_acceptance_report');
assert.equal(assertMeetingAppTimelineConnectorAdapterMatrix(rootConnectorAdapterMatrix).accepted, true);
const rootConnectorHostAdapterConfig = buildMeetingAppTimelineHostAdapterConfig(rootConnectorAdapterMatrix, 'google-meet');
assert.equal(rootConnectorHostAdapterConfig.schema, 'meeting_app_timeline_host_adapter_config');
assert.equal(buildMeetingAppTimelineHostAdapterConfigFromRoot(rootConnectorAdapterMatrix, 'google-meet').accepted, true);
const rootConnectorHostAdapterConfigIndex = buildMeetingAppTimelineHostAdapterConfigIndex(rootConnectorAdapterMatrix);
assert.equal(rootConnectorHostAdapterConfigIndex.schema, 'meeting_app_timeline_host_adapter_config_index');
assert.equal(buildMeetingAppTimelineHostAdapterConfigIndexFromRoot(rootConnectorAdapterMatrix).accepted, true);
assert.equal(assertMeetingAppTimelineHostAdapterConfig(rootConnectorAdapterMatrix, 'google-meet').accepted, true);
assert.equal(assertMeetingAppTimelineHostAdapterConfigIndex(rootConnectorHostAdapterConfigIndex).accepted, true);
const rootConnectorResolvedHostConfig = resolveMeetingAppTimelineHostAdapterConfig(rootConnectorHostAdapterConfigIndex, 'https://meet.google.com/abc-defg-hij');
assert.equal(rootConnectorResolvedHostConfig.schema, 'meeting_app_timeline_host_adapter_config_resolution');
assert.equal(resolveMeetingAppTimelineHostAdapterConfigFromRoot(rootConnectorHostAdapterConfigIndex, 'https://meet.google.com/abc-defg-hij').accepted, true);
assert.equal(assertMeetingAppTimelineResolvedHostAdapterConfig(rootConnectorHostAdapterConfigIndex, 'https://meet.google.com/abc-defg-hij').platform, 'google_meet');
const rootConnectorHostAdapterBootstrapPlan = buildMeetingAppTimelineHostAdapterBootstrapPlan(rootConnectorResolvedHostConfig);
assert.equal(rootConnectorHostAdapterBootstrapPlan.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan');
assert.equal(buildMeetingAppTimelineHostAdapterBootstrapPlanFromRoot(rootConnectorResolvedHostConfig).accepted, true);
assert.equal(assertMeetingAppTimelineHostAdapterBootstrapPlan(rootConnectorResolvedHostConfig).platform, 'google_meet');
const rootConnectorHostAdapterBootstrapPlanMatrix = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix(rootConnectorHostAdapterConfigIndex);
assert.equal(rootConnectorHostAdapterBootstrapPlanMatrix.schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix');
assert.equal(buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixFromRoot(rootConnectorHostAdapterConfigIndex).accepted, true);
assert.equal(buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport(rootConnectorHostAdapterBootstrapPlanMatrix).accepted, true);
assert.equal(buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReportFromRoot(rootConnectorHostAdapterBootstrapPlanMatrix).schema, 'meeting_app_timeline_host_adapter_bootstrap_plan_matrix_acceptance_report');
assert.equal(assertMeetingAppTimelineHostAdapterBootstrapPlanMatrix(rootConnectorHostAdapterBootstrapPlanMatrix).platform_count, 1);
assert.equal(rootMeetingAppSdk.connectorReleaseGate(rootConnectorHostInstallChecklist).schema, 'meeting_app_timeline_connector_release_gate');
assert.equal(rootMeetingAppSdk.assertConnectorReleaseGate(rootConnectorHostInstallChecklist).accepted, true);
assert.equal(rootMeetingAppSdk.connectorPlatformRoadmap(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  consumerHandoff: rootConnectorConsumerHandoff,
}).schema, 'meeting_app_timeline_connector_platform_roadmap');
assert.equal(rootMeetingAppSdk.assertConnectorPlatformRoadmap(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  consumerHandoff: rootConnectorConsumerHandoff,
}).accepted, true);
assert.equal(rootMeetingAppSdk.connectorAdapterMatrix(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  platformRoadmap: rootConnectorPlatformRoadmap,
  consumerHandoff: rootConnectorConsumerHandoff,
  fieldIntakeIndex: rootConnectorFieldIntakeIndex,
  bridgeHandoff: rootConnectorBridgeHandoff,
  smokePlan: rootConnectorSmokePlan,
}).schema, 'meeting_app_timeline_connector_adapter_matrix');
assert.equal(rootMeetingAppSdk.connectorAdapterMatrix(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  platformRoadmap: rootConnectorPlatformRoadmap,
  consumerHandoff: rootConnectorConsumerHandoff,
  fieldIntakeIndex: rootConnectorFieldIntakeIndex,
  bridgeHandoff: rootConnectorBridgeHandoff,
  smokePlan: rootConnectorSmokePlan,
}).rows[0].provider_replay.accepted, true);
assert.equal(rootMeetingAppSdk.assertConnectorAdapterMatrix(rootConnectorHostInstallChecklist, {
  releaseGate: rootConnectorReleaseGate,
  platformRoadmap: rootConnectorPlatformRoadmap,
  consumerHandoff: rootConnectorConsumerHandoff,
  fieldIntakeIndex: rootConnectorFieldIntakeIndex,
  bridgeHandoff: rootConnectorBridgeHandoff,
  smokePlan: rootConnectorSmokePlan,
}).accepted, true);
assert.equal(rootMeetingAppSdk.connectorHostAdapterConfig(rootConnectorAdapterMatrix, 'google-meet').platform, 'google_meet');
assert.equal(rootMeetingAppSdk.assertConnectorHostAdapterConfig(rootConnectorAdapterMatrix, 'google-meet').accepted, true);
assert.equal(rootMeetingAppSdk.connectorHostAdapterConfigIndex(rootConnectorAdapterMatrix).accepted, true);
assert.equal(rootMeetingAppSdk.assertConnectorHostAdapterConfigIndex(rootConnectorHostAdapterConfigIndex).row_count, 1);
assert.equal(rootMeetingAppSdk.resolveConnectorHostAdapterConfig('https://meet.google.com/abc-defg-hij', rootConnectorHostAdapterConfigIndex).accepted, true);
assert.equal(rootMeetingAppSdk.assertResolvedConnectorHostAdapterConfig({ tab: { url: 'https://meet.google.com/abc-defg-hij', active: true } }, rootConnectorHostAdapterConfigIndex).platform, 'google_meet');
assert.equal(rootMeetingAppSdk.connectorHostAdapterBootstrapPlan('https://meet.google.com/abc-defg-hij', rootConnectorHostAdapterConfigIndex).accepted, true);
assert.equal(rootMeetingAppSdk.assertConnectorHostAdapterBootstrapPlan('https://meet.google.com/abc-defg-hij', rootConnectorHostAdapterConfigIndex).platform, 'google_meet');
assert.equal(rootMeetingAppSdk.connectorHostAdapterBootstrapPlanMatrix(rootConnectorHostAdapterConfigIndex).accepted, true);
assert.equal(rootMeetingAppSdk.connectorHostAdapterBootstrapPlanMatrixAcceptanceReport(rootConnectorHostAdapterBootstrapPlanMatrix).accepted, true);
assert.equal(rootMeetingAppSdk.assertConnectorHostAdapterBootstrapPlanMatrix(rootConnectorHostAdapterBootstrapPlanMatrix).platform_count, 1);
const rootProviderReplay = buildMeetingPlatformProviderReplayReport('google-meet');
assert.equal(rootProviderReplay.schema, 'meeting_platform_provider_replay_report');
assert.equal(rootProviderReplay.accepted, true);
assert.equal(rootProviderReplay.runtime_contract.provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformProviderReplayReportFromRoot('zoom').accepted, true);
assert.equal(sampleMeetingPlatformProviderEvents('webex').length, 4);
assert.equal(sampleMeetingPlatformProviderEventsFromRoot('google-meet').length, 4);
assert.equal(assertMeetingPlatformProviderReplayReport(rootProviderReplay).accepted, true);
const rootProviderReplayMatrix = buildMeetingPlatformProviderReplayMatrix({
  platforms: ['google-meet', 'zoom'],
});
assert.equal(rootProviderReplayMatrix.schema, 'meeting_platform_provider_replay_matrix');
assert.equal(rootProviderReplayMatrix.accepted, true);
assert.equal(rootProviderReplayMatrix.accepted_count, 2);
assert.equal(buildMeetingPlatformProviderReplayMatrixFromRoot({
  platforms: ['google-meet', 'zoom'],
}).accepted_count, 2);
assert.equal(assertMeetingPlatformProviderReplayMatrix({
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(rootMeetingAppSdk.providerReplayReport('google-meet').schema, 'meeting_platform_provider_replay_report');
assert.equal(rootMeetingAppSdk.assertProviderReplayReport('google-meet').accepted, true);
assert.equal(rootMeetingAppSdk.providerReplayMatrix({
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(rootMeetingAppSdk.assertProviderReplayMatrix({
  platforms: ['google-meet'],
}).accepted, true);
assert.equal((await runMeetingAppTimelineConnectorBridgeSmoke(rootConnectorPackage)).schema, 'meeting_app_timeline_connector_bridge_smoke_report');
assert.equal((await runMeetingAppTimelineConnectorBridgeSmokeFromRoot(rootConnectorPackage)).accepted, true);
assert.equal((await assertMeetingAppTimelineConnectorBridgeSmoke(rootConnectorPackage)).accepted, true);
assert.equal(rootConnectorSmokePlan.schema, 'meeting_app_timeline_connector_smoke_plan');
assert.equal(rootConnectorSmokePlan.accepted, true);
assert.equal(rootConnectorSmokePlan.rows[0].steps.some((step) => step.action === 'observe_platform_candidates'), true);
assert.equal(rootConnectorSmokePlan.rows[0].steps.some((step) => step.action === 'insert_annotation'), true);
assert.equal(buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(rootConnectorSmokePlan).accepted, true);
assert.equal(buildMeetingAppTimelineConnectorSmokePlanFromRoot(rootConnectorPackage).schema, 'meeting_app_timeline_connector_smoke_plan');
assert.equal(buildMeetingAppTimelineConnectorSmokePlanAcceptanceReportFromRoot(rootConnectorPackage).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorSmokePlan(rootConnectorSmokePlan), rootConnectorSmokePlan);
assert.equal((await runMeetingAppTimelineConnectorSmokePlan(rootConnectorSmokePlan)).schema, 'meeting_app_timeline_connector_smoke_run_report');
assert.equal((await runMeetingAppTimelineConnectorSmokePlanFromRoot(rootConnectorPackage)).accepted, true);
assert.equal((await assertMeetingAppTimelineConnectorSmokeRun(rootConnectorSmokePlan)).accepted, true);
assert.equal(assertMeetingAppTimelineConnectorHostInstallChecklist(rootConnectorPackage), rootConnectorPackage);
assert.equal(assertMeetingAppTimelineConnectorPackage(rootConnectorPackage), rootConnectorPackage);
assert.equal(rootMeetingAppSdk.meetingAppAdapterCapability('google-meet').schema, 'meeting_app_adapter_capability_report');
assert.equal(rootMeetingAppSdk.adapterCapability('google-meet').platform, 'google_meet');
assert.equal(rootMeetingAppSdk.meetingAppAdapterCapabilityMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.adapterCapabilityMatrix().provider_axis_declared_count, 1);
assert.equal(rootMeetingAppSdk.meetingAppAdapterExecutionPlan('google-meet').schema, 'meeting_app_adapter_execution_plan');
assert.equal(rootMeetingAppSdk.adapterExecutionPlan('google-meet').realtime_ready, true);
assert.equal(rootMeetingAppSdk.meetingAppAdapterExecutionPlanMatrix().accepted_count, 1);
assert.equal(rootMeetingAppSdk.adapterExecutionPlanMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.platformAdaptationPackage('google-meet').adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
assert.equal(rootMeetingAppSdk.platformAdaptationPackage('google-meet').adapter_surfaces.primary, 'browser_extension');
assert.equal(rootMeetingAppSdk.adaptationPackageMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.platformConsumerHandoff().schema, 'meeting_platform_consumer_handoff');
assert.equal(rootMeetingAppSdk.consumerHandoff().accepted, true);
assert.equal(rootMeetingAppSdk.consumerHandoff({ platforms: ['zoom'] }).adaptation_roadmap.rows[0].recommended_first_surface, 'native_detector');
assert.equal(rootMeetingAppSdk.platformImplementationHandoff('google-meet').schema, 'meeting_platform_implementation_handoff');
assert.equal(rootMeetingAppSdk.implementationHandoff('google-meet').provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(rootMeetingAppSdk.platformImplementationHandoffMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.implementationHandoffMatrix().implementation_ready_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterAuthoringPlan('google-meet').built_in, true);
assert.equal(rootMeetingAppSdk.adapterAuthoringMatrix({ platforms: ['google-meet', 'Acme Rooms'] }).external_authoring_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterPortfolioItem('google-meet').p1_provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(rootMeetingAppSdk.adapterPortfolioItem('zoom').recommended_first_surface, 'native_detector');
assert.equal(rootMeetingAppSdk.adapterPortfolio({ platforms: ['google-meet', 'Acme Rooms'] }).external_authoring_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterAcceptanceChecklist('google-meet', {}, { target: 'static' }).accepted, true);
assert.equal(rootMeetingAppSdk.adapterAcceptanceChecklistMatrix({ platforms: ['google-meet'] }, { target: 'static' }).accepted_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterExportPackage('google-meet', {}, { target: 'static' }).export_ready, true);
assert.equal(rootMeetingAppSdk.adapterExportPackageMatrix({ platforms: ['google-meet'] }, { target: 'static' }).export_ready_count, 1);
const rootGoogleExportPackage = rootMeetingAppSdk.platformAdapterExportPackage('google-meet', {}, { target: 'static' });
const rootGoogleExportFiles = rootGoogleExportPackage.host_files.map((file) => file.path);
assert.equal(rootMeetingAppSdk.platformAdapterImportPlan(rootGoogleExportPackage, { availableFiles: rootGoogleExportFiles }).accepted, true);
assert.equal(rootMeetingAppSdk.adapterImportPlanMatrix([rootGoogleExportPackage], { availableFiles: rootGoogleExportFiles }).accepted_count, 1);
const rootGoogleImportPlan = rootMeetingAppSdk.platformAdapterImportPlan(rootGoogleExportPackage, { availableFiles: rootGoogleExportFiles });
assert.equal(rootMeetingAppSdk.platformAdapterInstallManifest([rootGoogleImportPlan]).accepted, true);
assert.equal(rootMeetingAppSdk.adapterInstallManifest([rootGoogleImportPlan]).browser_extension.platform_count, 1);
assert.equal(rootMeetingAppSdk.assertPlatformAdapterInstallManifest(rootMeetingAppSdk.adapterInstallManifest([rootGoogleImportPlan])).accepted, true);
const rootInstallManifest = rootMeetingAppSdk.adapterInstallManifest([rootGoogleImportPlan]);
assert.equal(rootMeetingAppSdk.platformAdapterLaunchPlan(rootInstallManifest, { url: 'https://meet.google.com/abc-defg-hij' }).accepted, true);
assert.equal(rootMeetingAppSdk.adapterLaunchPlan(rootInstallManifest, { platform: 'google-meet' }).platform, 'google_meet');
assert.equal(rootMeetingAppSdk.adapterCandidateLaunchPlan(rootInstallManifest, {
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: [rootGoogleActiveSnapshot],
  }],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(rootMeetingAppSdk.assertPlatformAdapterLaunchPlan(rootMeetingAppSdk.adapterLaunchPlan(rootInstallManifest, { platform: 'google-meet' })).accepted, true);
const rootGoogleLaunchPlan = rootMeetingAppSdk.platformAdapterLaunchPlan(rootInstallManifest, { url: 'https://meet.google.com/abc-defg-hij' });
const rootAdapterSessionCalls = [];
const rootAdapterSessionClient = {
  async observePlatformCandidates(payload) {
    rootAdapterSessionCalls.push(['observePlatformCandidates', payload]);
    return { ok: true };
  },
  async insertAnnotation(platform, payload) {
    rootAdapterSessionCalls.push(['insertAnnotation', platform, payload]);
    return { ok: true };
  },
};
const rootAdapterSession = rootMeetingAppSdk.platformAdapterSession(rootGoogleLaunchPlan, rootAdapterSessionClient, { clock: () => 777 });
assert.equal(rootAdapterSession.schema, 'meeting_platform_adapter_session');
assert.equal((await rootAdapterSession.observeAxis()).action, 'observe_axis');
assert.equal((await rootAdapterSession.insertAnnotation({ label: 'smoke mark' })).payload.captured_at_ms, 777);
assert.equal(rootAdapterSessionCalls[1][1], 'google_meet');
assert.equal(rootMeetingAppSdk.adapterSessionHandoff(rootGoogleLaunchPlan).schema, 'meeting_platform_adapter_session_handoff');
const rootAdapterRunner = rootMeetingAppSdk.platformAdapterRunner(rootInstallManifest, rootAdapterSessionClient, { clock: () => 778 });
assert.equal(rootAdapterRunner.schema, 'meeting_platform_adapter_runner');
assert.equal((await rootAdapterRunner.open({ url: 'https://meet.google.com/abc-defg-hij' })).schema, 'meeting_platform_adapter_open_session_event');
assert.equal((await rootAdapterRunner.insertAnnotation({ label: 'runner smoke mark' })).payload.captured_at_ms, 778);
assert.equal((await rootMeetingAppSdk.openPlatformAdapterSession(rootInstallManifest, {
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  client: rootAdapterSessionClient,
  clock: () => 779,
})).payload.observe_event.captured_at_ms, 779);
assert.equal(rootMeetingAppSdk.adapterRunnerHandoff(rootInstallManifest).schema, 'meeting_platform_adapter_runner_handoff');
const rootAdapterMessageBridge = rootMeetingAppSdk.platformAdapterMessageBridge(rootInstallManifest, rootAdapterSessionClient, {
  clock: () => 780,
});
assert.equal(rootAdapterMessageBridge.schema, 'meeting_platform_adapter_message_bridge');
assert.equal((await rootAdapterMessageBridge.handleMessage({
  type: 'meeting_timeline.open_session',
  payload: { url: 'https://meet.google.com/abc-defg-hij' },
})).result.payload.observe_event.captured_at_ms, 780);
assert.equal((await rootAdapterMessageBridge.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: { mark: { label: 'bridge smoke mark' } },
})).action, 'insert_annotation');
assert.equal(rootMeetingAppSdk.adapterMessageBridgeHandoff(rootInstallManifest).schema, 'meeting_platform_adapter_message_bridge_handoff');
const rootAdapterSmoke = await rootMeetingAppSdk.adapterSmoke(rootInstallManifest, {
  platforms: ['google-meet'],
  captured_at_ms: 1_782_900_600_000,
});
assert.equal(rootAdapterSmoke.accepted, true);
assert.equal(rootAdapterSmoke.rows[0].platform, 'google_meet');
assert.equal((await rootMeetingAppSdk.assertAdapterSmoke(rootInstallManifest, {
  platforms: ['google-meet'],
})).accepted, true);
assert.equal((await runMeetingPlatformAdapterSmokeFromRoot(rootInstallManifest, {
  platforms: ['google-meet'],
})).accepted, true);
assert.equal((await assertMeetingPlatformAdapterSmokeFromRoot(rootInstallManifest, {
  platforms: ['google-meet'],
})).schema, 'meeting_platform_adapter_smoke_report');
assert.equal((await runMeetingPlatformAdapterSmoke(rootInstallManifest, {
  platforms: ['google-meet'],
})).accepted, true);
assert.equal((await assertMeetingPlatformAdapterSmoke(rootInstallManifest, {
  platforms: ['google-meet'],
})).accepted, true);
assert.equal(rootMeetingAppSdk.platformRuntimeBundle('google-meet').runtime.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(rootMeetingAppSdk.runtimeBundleMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterRoute('google-meet').platform, 'google_meet');
assert.equal(rootMeetingAppSdk.platformAdapterRoute('google-meet').adapter_surfaces.primary, 'browser_extension');
assert.equal(rootMeetingAppSdk.adapterRouteMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterBlueprint('google-meet').platform, 'google_meet');
assert.equal(rootMeetingAppSdk.platformAdapterBlueprint('google-meet').surfaces.provider_reconcile.blocks_realtime, false);
assert.equal(rootMeetingAppSdk.adapterBlueprintMatrix().platform_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
}).selected_surface, 'browser_extension');
assert.equal(rootMeetingAppSdk.adapterDecision({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(rootMeetingAppSdk.platformAdapterDecisionMatrix({}, {
  platforms: ['google-meet'],
}).accepted_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
}).realtime_startup_ready, true);
assert.equal(rootMeetingAppSdk.adapterStartupPlan({
  url: 'https://zoom.us/j/987654321',
}).platform, 'zoom');
assert.equal(rootMeetingAppSdk.platformAdapterStartupPlanMatrix({}, {
  platforms: ['google-meet'],
}).realtime_startup_ready_count, 1);
assert.equal(rootMeetingAppSdk.platformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: [rootGoogleActiveSnapshot],
}).accepted, true);
assert.equal(rootMeetingAppSdk.platformAdapterCurrentWindowPreflight({
  document: rootSmokeDocument(),
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(rootMeetingAppSdk.adapterCurrentWindowPreflight({
  document: rootSmokeDocument(),
}, {
  requireSpeakerTrack: true,
}).readiness.realtime_annotation_ready, true);
assert.equal(rootMeetingAppSdk.adapterPreflight({
  platform: 'google-meet',
  snapshots: [rootGoogleActiveSnapshot],
}).readiness.realtime_annotation_ready, true);
assert.equal(rootMeetingAppSdk.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [rootGoogleActiveSnapshot],
  },
}).realtime_ready_count, 1);
assert.equal(rootMeetingAppSdk.adapterCandidatePreflight({
  candidates: [{
    url: 'https://zoom.us/j/987654321',
    title: 'Zoom Meeting',
  }, {
    document: rootSmokeDocument(),
  }],
}, {
  requireSpeakerTrack: true,
}).selected_platform, 'google_meet');
assert.equal(rootMeetingAppSdk.platformAdaptationStrategy('google-meet').adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
assert.equal(rootMeetingAppSdk.adaptationStrategyMatrix().provider_reconcile_required_count, 1);
assert.equal(rootMeetingAppSdk.connectorHub().accepted, true);
assert.equal(buildMeetingPlatformAdaptationPackageFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
assert.equal(buildMeetingPlatformAdaptationPackageFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).evidence_thresholds.production.provider_records_required, true);
assert.equal(buildMeetingPlatformConsumerHandoffFromRoot({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(buildMeetingPlatformImplementationHandoffFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).recommended_first_surface, 'browser_extension');
assert.equal(buildMeetingPlatformRuntimeProfileFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).adapter_surfaces.primary, 'browser_extension');
assert.equal(buildMeetingPlatformRuntimeBundleFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_runtime_bundle');
assert.equal(buildMeetingPlatformAdapterRouteFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).adapter_surfaces.primary, 'browser_extension');
assert.equal(buildMeetingPlatformAdapterDecisionFromRoot({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
}).selected_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformAdapterDecisionMatrixFromRoot({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).accepted_count, 1);
assert.equal(buildMeetingPlatformAdapterStartupPlanFromRoot({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
}).message_contract.insert_annotation, 'meeting_timeline.insert_mark');
assert.equal(buildMeetingPlatformAdapterStartupPlanMatrixFromRoot({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).realtime_startup_ready_count, 1);
assert.equal(buildMeetingPlatformAdapterPreflightFromRoot({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: [rootGoogleActiveSnapshot],
}, {
  baseUrl: 'http://localhost:8787',
}).status, 'ready_for_realtime_annotations');
assert.equal(buildMeetingPlatformAdapterCurrentWindowPreflightFromRoot({
  document: rootSmokeDocument(),
}, {
  baseUrl: 'http://localhost:8787',
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterCandidatePreflightFromRoot({
  tabs: [{
    url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    title: 'Microsoft Teams',
  }, {
    document: rootSmokeDocument(),
  }],
}, {
  baseUrl: 'http://localhost:8787',
  requireSpeakerTrack: true,
}).accepted_count, 1);
assert.equal(buildMeetingPlatformAdapterPreflightMatrixFromRoot({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [rootGoogleActiveSnapshot],
  },
}).accepted_count, 1);
assert.equal(buildMeetingPlatformAdaptationStrategyFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).adaptation_playbook.integration_path.path, 'google_workspace_events_pubsub');
const rootPlatformConnector = buildMeetingPlatformConnectorFromRoot('google-meet', { baseUrl: 'http://localhost:8787' });
assert.equal(rootPlatformConnector.schema, 'meeting_platform_connector');
assert.equal(rootPlatformConnector.readiness.realtime_annotation_ready, true);
assert.equal(buildMeetingPlatformConnector('google-meet', { baseUrl: 'http://localhost:8787' }).platform, 'google_meet');
assert.equal(buildMeetingPlatformConnectorAcceptanceReport(rootPlatformConnector).accepted, true);
assert.equal(buildMeetingPlatformConnectorMatrix({ platforms: ['google-meet', 'teams'] }).accepted_count, 2);
assert.equal(buildMeetingPlatformConnectorHubFromRoot({ platforms: ['google-meet', 'teams'] }).schema, 'meeting_platform_connector_hub');
assert.equal(buildMeetingPlatformConnectorHub({ platforms: ['google-meet', 'teams'] }).accepted, true);
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_current_window, 'meeting_timeline.preflight_current_window');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.preflight_candidates, 'meeting_timeline.preflight_candidates');
assert.equal(buildMeetingAppExtensionCurrentWindowPreflightMessage({
  platform: 'google-meet',
  capturedAtMs: 123,
}).type, 'meeting_timeline.preflight_current_window');
assert.equal(buildMeetingAppExtensionPreflightCandidatesMessage({
  capturedAtMs: 124,
}).type, 'meeting_timeline.preflight_candidates');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.candidate_launch_plan, 'meeting_timeline.candidate_launch_plan');
assert.equal(MEETING_APP_EXTENSION_MESSAGE_TYPES.open_candidate_session, 'meeting_timeline.open_candidate_session');
assert.equal(buildMeetingAppExtensionCandidateLaunchPlanMessage({
  requestId: 'candidate-launch-001',
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', active: true }],
}).type, 'meeting_timeline.candidate_launch_plan');
assert.equal(buildMeetingAppExtensionOpenCandidateSessionMessage({
  requestId: 'open-candidate-001',
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', active: true }],
}).type, 'meeting_timeline.open_candidate_session');
assert.match(buildMeetingAppExtensionBackgroundSource({
  baseUrl: 'http://localhost:8787',
}), /meeting_timeline\.preflight_current_window/);
assert.match(buildMeetingAppExtensionBackgroundSource({
  baseUrl: 'http://localhost:8787',
}), /meeting_timeline\.preflight_candidates/);
assert.match(buildMeetingAppExtensionBackgroundSource({
  baseUrl: 'http://localhost:8787',
}), /tabs\.sendMessage/);
assert.equal(resolveMeetingPlatformConnectorInput('https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample', {
  platforms: ['google-meet', 'teams'],
}).platform, 'microsoft_teams');
assert.equal(createMeetingPlatformConnectorRuntime(rootPlatformConnector, {
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).supports('insert_annotation'), true);
assert.equal(createMeetingPlatformConnectorRuntimeFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).platform, 'google_meet');
assert.equal(createMeetingPlatformConnectorHub({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).connectorFor({ url: 'https://zoom.us/j/987654321' }).platform, 'zoom');
assert.equal(createMeetingPlatformConnectorHubFromRoot({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).resolvePlatform({ url: 'https://meet.google.com/abc-defg-hij' }).platform, 'google_meet');
assert.equal(createMeetingPlatformConnectorBrowserRuntime({
  platform: 'google-meet',
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).schema, 'meeting_platform_connector_browser_runtime');
assert.equal(createMeetingPlatformConnectorBrowserRuntimeFromRoot({
  platform: 'google-meet',
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).resolvePlatform({ platform: 'google-meet' }).platform, 'google_meet');
assert.equal(createMeetingPlatformConnectorContentScriptBridge({
  platform: 'google-meet',
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).schema, 'meeting_platform_connector_content_script_bridge');
assert.equal(createMeetingPlatformConnectorContentScriptBridgeFromRoot({
  platform: 'google-meet',
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).resolvePlatform({ platform: 'google-meet' }).platform, 'google_meet');
const connectorRuntimeClientFromRoot = createMeetingAppTimelineConnectorRuntimeClientFromRoot(rootConnectorPackage, {
  fetch: async () => new Response(JSON.stringify({ ok: true })),
});
assert.equal(connectorRuntimeClientFromRoot.schema, 'meeting_app_timeline_connector_runtime_client');
assert.equal(connectorRuntimeClientFromRoot.supports('insert_annotation', 'google-meet'), true);
assert.equal(createMeetingAppTimelineConnectorRuntimeClient(rootConnectorPackage, {
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).supports('observe_platform_candidates'), true);

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'http://localhost:8787',
  verify: false,
});
assert.equal(kit.platformRolloutPlan('google-meet').platform, 'google_meet');
assert.equal(kit.platformAdaptationRunbook('zoom').platform, 'zoom');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_rollout.type, 'meeting_platform_rollout_summary');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_registry_manifest.platform_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_connector_matrix.platform_count, 1);
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_connector_hub.schema, 'meeting_platform_connector_hub');
assert.equal(kit.platformConnector('google-meet').schema, 'meeting_platform_connector');
assert.equal(kit.platformConnectorMatrix({ platforms: ['google-meet', 'teams'] }).accepted_count, 2);
assert.equal(kit.platformConnectorHub({ platforms: ['google-meet', 'teams'] }).accepted, true);
assert.equal(kit.resolvePlatformConnector({ url: 'https://meet.google.com/abc-defg-hij' }).platform, 'google_meet');
assert.equal(kit.createPlatformConnectorRuntime('google-meet', {
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).supports('insert_annotation'), true);
assert.equal(kit.createPlatformConnectorHub({
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).connectorFor({ url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample' }).platform, 'microsoft_teams');
assert.equal(kit.createPlatformConnectorBrowserRuntime({
  platform: 'google-meet',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).schema, 'meeting_platform_connector_browser_runtime');
assert.equal(kit.createPlatformConnectorContentScriptBridge({
  platform: 'google-meet',
  fetch: async () => new Response(JSON.stringify({ ok: true })),
}).schema, 'meeting_platform_connector_content_script_bridge');
assert.equal(kit.meetingAppRuntimeAdapterProfile('https://meet.google.com/abc-defg-hij').platform, 'google_meet');
assert.equal(kit.meetingAppRuntimeAdapterProfileMatrix({ platforms: ['google-meet', 'teams'] }).runtime_ready_count, 2);
assert.equal(kit.meetingAppAdapterFit({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).ready_for_speaker_track, true);
assert.equal(kit.meetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).observer_runtime.factory, 'createMeetingAppBrowserRuntime');
assert.equal(kit.meetingAppObserverSchedulerConfig('google-meet').schema, 'meeting_app_observer_scheduler_config');
assert.equal(kit.meetingAppObserverSchedulerConfigMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.selectMeetingAppRuntimeAdapter('https://meet.google.com/abc-defg-hij').launch.runtime_options.runtimePreset, 'google_meet');
assert.equal(kit.meetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').readiness.ready_to_start, true);
assert.equal(kit.meetingAppRuntimeAdapterHandoffMatrix({ platforms: ['google-meet', 'teams'], surfaces: ['browser-extension'] }).handoff_count, 2);
assert.equal(kit.meetingAppRuntimeAdapterHostPackage({ platforms: ['google-meet'], surfaces: ['browser-extension'] }).accepted, true);
assert.equal(kit.meetingAppRuntimeAdapterHandoffAcceptance('https://meet.google.com/abc-defg-hij').accepted, true);
const integrationRuntimeManifest = buildMeetingPlatformIntegrationRuntimeManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
});
assert.equal(integrationRuntimeManifest.host_integration_ready, true);
assert.equal(integrationRuntimeManifest.platform_conformance_report.accepted, true);
assert.equal(assertMeetingPlatformIntegrationRuntimeManifest(integrationRuntimeManifest).platform_count, 2);
const integrationRuntimeRunManifest = await runMeetingPlatformIntegrationRuntimeManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
});
assert.equal(integrationRuntimeRunManifest.host_integration_ready, true);
assert.equal(integrationRuntimeRunManifest.handoff_readiness_matrix.runtime_host_replay_ready_count, 0);
const integrationRuntime = createMeetingPlatformIntegrationRuntime(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
});
assert.equal(integrationRuntime.manifest().host_integration_ready, true);
assert.equal(integrationRuntime.conformance().accepted, true);
assert.equal(integrationRuntime.assertConformance().accepted_count, 1);
assert.equal(integrationRuntime.manifest().adapter_route_matrix.platform_count, 1);
assert.equal((await integrationRuntime.runManifest()).host_integration_ready, true);
assert.equal(integrationRuntime.runtimeBundles().platform_count, 1);
assert.equal(integrationRuntime.adapterRoutes().rows[0].first_route, 'local_observer_axis');
assert.equal(integrationRuntime.adaptationStrategyMatrix().strategy_count, 1);
assert.equal(integrationRuntime.resolvePlatform({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(detectMeetingPlatformForBrowser({
  url: 'https://meet.google.com/abc-defg-hij',
}).platform, 'google_meet');
assert.equal(resolveMeetingPlatformForInput({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).supported, true);
assert.equal(resolveMeetingPlatformCandidates({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).selected_resolution.platform, 'google_meet');
assert.equal((await integrationRuntime.observePlatformCandidates({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  observedAtMs: 1_782_614_400_000,
})).signals[0].type, 'meeting_started');
assert.equal(createMeetingPlatformIntegrationBrowserRuntime(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  url: 'https://meet.google.com/abc-defg-hij',
}).detect().platform, 'google_meet');
assert.equal(createMeetingPlatformIntegrationContentScriptBridge(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  url: 'https://meet.google.com/abc-defg-hij',
  extensionMessaging: false,
}).detect().platform, 'google_meet');
assert.equal(typeof installMeetingPlatformIntegrationContentScriptBridge, 'function');
assert.equal(integrationRuntime.timelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [{
    id: 'runtime-smoke-note',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  }],
}).diagnostics.marker_count, 1);
const runtimeEvent = buildMeetingPlatformAnnotationRuntimeEvent('google-meet', {
  annotation: {
    id: 'package-runtime-event-note-1',
    label: 'why?',
    captured_at_ms: 1_782_614_401_000,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
}, {
  now: () => 1_782_614_402_000,
});
assert.equal(runtimeEvent.schema, MEETING_PLATFORM_RUNTIME_EVENT_SCHEMA);
assert.equal(runtimeEvent.action, 'insert_annotation');
assert.equal(buildMeetingPlatformCandidateObservationRuntimeEvent({
  windows: [{
    tabs: [{ active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' }],
  }],
}, {
  now: () => 1_782_614_402_000,
}).action, 'observe_platform_candidates');
assert.equal(buildMeetingPlatformRunManifestRuntimeEvent({
  platforms: ['google-meet'],
  requireHandoffReady: true,
}, {
  now: () => 1_782_614_402_000,
}).action, 'run_manifest');
assert.equal(buildMeetingPlatformRunHandoffReadinessRuntimeEvent({
  platforms: ['zoom'],
}, {
  now: () => 1_782_614_402_000,
}).action, 'run_handoff_readiness');
assert.equal(meetingPlatformRuntimeEventEndpoint({
  baseUrl: 'http://localhost:8787',
}), 'http://localhost:8787/api/meeting-platform/runtime-events');
const smokeRuntimeEventClient = createMeetingPlatformRuntimeEventClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response('{}'),
});
assert.equal(typeof smokeRuntimeEventClient.send, 'function');
assert.equal(typeof smokeRuntimeEventClient.runManifest, 'function');
assert.equal(typeof smokeRuntimeEventClient.runHandoffReadiness, 'function');
assert.equal(kit.platformRegistryEntry('google-meet').annotations.timestamp_field, 'captured_at_ms');
assert.equal(kit.platformRegistryManifest({ platforms: ['zoom'] }).rows[0].platform, 'zoom');
assert.equal(meetingPlatformEventAdapterFor('teams').key, 'microsoft_teams');
assert.equal(buildMeetingPlatformRegistryEntry('google-meet', {
  baseUrl: 'http://localhost:8787',
}).runtime.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(buildMeetingPlatformRegistryManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).provider_required_for_realtime_count, 0);
assert.equal(buildMeetingPlatformRegistryAcceptanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted, true);
assert.equal(assertMeetingPlatformRegistryManifest({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted, true);
assert.equal(buildMeetingPlatformConformanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(assertMeetingPlatformConformanceReport({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).rows[0].adapter_first_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformConsumerHandoff({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).consumer_ready_count, 1);
assert.equal(assertMeetingPlatformConsumerHandoff({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).rows[0].adapter_first_route, 'local_observer_axis');
assert.equal(kit.platformConsumerHandoff({ platforms: ['zoom'] }).entrypoints.kit_methods.includes('platformConsumerHandoff'), true);
assert.equal(kit.platformConsumerHandoff({ platforms: ['zoom'] }).adaptation_roadmap.rows[0].recommended_first_surface, 'native_detector');
assert.equal(kit.platformImplementationHandoff('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformImplementationHandoffMatrix({ platforms: ['zoom'] }).implementation_ready_count, 1);
assert.equal(kit.platformAdapterAuthoringPlan('zoom').provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformAdapterAuthoringMatrix({ platforms: ['zoom'] }).built_in_count, 1);
assert.equal(kit.platformAdapterPortfolioItem('zoom').p1_provider_reconcile.path, 'zoom_meeting_webhooks');
assert.equal(kit.platformAdapterPortfolioItem('zoom').recommended_first_surface, 'native_detector');
assert.equal(kit.platformAdapterPortfolio({ platforms: ['zoom'] }).pilot_ready_count, 1);
assert.equal(kit.platformAdapterAcceptanceChecklist('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(kit.platformAdapterAcceptanceChecklistMatrix({ platforms: ['zoom'] }, { target: 'static' }).accepted_count, 1);
assert.equal(kit.platformAdapterExportPackage('zoom', {}, { target: 'static' }).host_files.some((file) => file.source === 'runtime_bundle'), true);
assert.equal(kit.platformAdapterExportPackageMatrix({ platforms: ['zoom'] }, { target: 'static' }).export_ready_count, 1);
const kitZoomExportPackage = kit.platformAdapterExportPackage('zoom', {}, { target: 'static' });
const kitZoomExportFiles = kitZoomExportPackage.host_files.map((file) => file.path);
assert.equal(kit.platformAdapterImportPlan(kitZoomExportPackage, { availableFiles: kitZoomExportFiles }).selected_surface, 'native_detector');
assert.equal(kit.platformAdapterImportPlanMatrix([kitZoomExportPackage], { availableFiles: kitZoomExportFiles }).accepted_count, 1);
assert.equal(kit.platformLiveAdapterHandoff('zoom').sdk.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(kit.platformLiveAdapterHandoffBundle({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformHostIntegrationPlan({
  platforms: ['zoom'],
}).runtime_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(kit.platformHostIntegrationPlan({
  platforms: ['zoom'],
}).runtime_bundle_matrix.platform_count, 1);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'src/meeting-platform-host.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'src/platform-adapters/zoom.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-runtime-bundles.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-strategy.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-integration-runtime.mjs'), true);
assert.equal(kit.platformHostIntegrationScaffold({
  platforms: ['zoom'],
}).files.some((file) => file.path === 'scripts/print-consumer-handoff.mjs'), true);
assert.equal(kit.platformProviderConnectionPack('zoom').security.verifier, 'verifyZoomWebhookEvent');
assert.equal(kit.platformProviderConnectionMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSubscriptionHandoff('zoom').schema, 'meeting_platform_subscription_handoff');
assert.equal(kit.platformSubscriptionHandoffMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSpeakerTrackPlan('zoom').schema, 'meeting_platform_speaker_track_plan');
assert.equal(kit.platformSpeakerTrackMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSpeakerTrack('zoom', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      speaker_name: 'Ada',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_402_000,
      speaker_name: 'Ada',
    },
  ],
}).mark_count, 1);
assert.equal(kit.platformParticipantTrackPlan('zoom').schema, 'meeting_platform_participant_track_plan');
assert.equal(kit.platformParticipantTrackMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformParticipantTrack('zoom', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).mark_count, 1);
assert.equal(buildMeetingAppFixtureTrackReadinessReport({
  platforms: ['google-meet'],
}).accepted, true);
assert.equal(buildMeetingAppTrackPipeline([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).speaker_track.mark_count, 1);
assert.equal(kit.meetingAppTrackPipeline([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
], {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).marks[0].intent, 'speaker_track');
const trackRuntime = createMeetingAppTrackRuntime(client, {
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
});
assert.equal((await trackRuntime.observe([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal((await kit.meetingAppTrackRuntime({
  speakerTrackOptions: {
    minStableMs: 250,
    minSegmentMs: 0,
    closeOpenSegmentsAtMs: 1783356001000,
  },
}).observe([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal((await createMeetingAppTimelineRuntime(client, {
  trackRuntimeOptions: {
    speakerTrackOptions: {
      minStableMs: 250,
      minSegmentMs: 0,
      closeOpenSegmentsAtMs: 1783356001000,
    },
  },
}).observeMeetingAppTracks([
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000000,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
  {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    observedAtMs: 1783356000400,
    activeSpeaker: { id: 'ada', name: 'Ada', speaking: true },
    participants: [{ id: 'ada', name: 'Ada', speaking: true }],
  },
])).new_mark_count, 1);
assert.equal(resolveMeetingAppRuntimeAdapterProfile('https://meet.google.com/abc-defg-hij').platform, 'google_meet');
assert.equal(buildMeetingAppAdapterFitReport({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).accepted, true);
assert.equal(buildMeetingAppAdapterFitMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_adapter_fit_matrix');
assert.equal(buildMeetingAppAdapterCapabilityReport('google-meet', {
  input: {
    url: 'https://meet.google.com/abc-defg-hij',
    page: {
      controls: [{ label: 'Leave call' }],
      participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
    },
  },
}).pilot_ready, true);
assert.equal(buildMeetingAppAdapterCapabilityMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_adapter_capability_matrix');
const packageExecutionPlan = buildMeetingAppAdapterExecutionPlan('google-meet', {
  input: {
    url: 'https://meet.google.com/abc-defg-hij',
    page: {
      controls: [{ label: 'Leave call' }],
      participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
    },
  },
});
assert.equal(packageExecutionPlan.schema, 'meeting_app_adapter_execution_plan');
assert.equal(packageExecutionPlan.realtime_ready, true);
assert.equal(buildMeetingAppAdapterExecutionPlanMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_adapter_execution_plan_matrix');
const packageIntegration = buildMeetingAppAdapterIntegrationPackage('google-meet', {
  input: {
    url: 'https://meet.google.com/abc-defg-hij',
    page: {
      controls: [{ label: 'Leave call' }],
      participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
    },
  },
});
assert.equal(packageIntegration.schema, 'meeting_app_adapter_integration_package');
assert.equal(packageIntegration.realtime_ready, true);
assert.equal(packageIntegration.entrypoints.adapter_integration_package, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package');
assert.equal(packageIntegration.entrypoints.platform_runtime_bundle, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle');
assert.equal(packageIntegration.runtime_delivery.adapter_route.first_route, 'local_observer_axis');
assert.equal(packageIntegration.runtime_delivery.host.annotation_timestamp_field, 'captured_at_ms');
assert.equal(assertMeetingAppAdapterIntegrationPackage(packageIntegration).platform, 'google_meet');
const packageIntegrationMatrix = buildMeetingAppAdapterIntegrationPackageMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
});
assert.equal(packageIntegrationMatrix.schema, 'meeting_app_adapter_integration_package_matrix');
assert.equal(assertMeetingAppAdapterIntegrationPackageMatrix(packageIntegrationMatrix).accepted, true);
assert.equal(buildMeetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, { platform: 'google-meet' }).schema, 'meeting_app_runtime_observer_plan');
assert.equal(buildMeetingAppRuntimeObserverPlanMatrix({
  platforms: ['google-meet'],
  inputs: {
    google_meet: {
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
  },
}).schema, 'meeting_app_runtime_observer_plan_matrix');
assert.equal(buildMeetingAppAdapterManifest('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterManifest('google-meet').contracts.timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingAppAdapterManifestMatrix({ platforms: ['google-meet', 'zoom'] }).accepted_count, 2);
assert.equal(assertMeetingAppAdapterManifest('zoom').runtime.observe_mutations, true);
assert.equal(assertMeetingAppAdapterManifestMatrix({ platforms: ['google-meet'] }).platform_count, 1);
assert.equal(kit.meetingAppAdapterManifest('google-meet').capture.selector_counts.participant > 0, true);
assert.equal(kit.meetingAppAdapterManifestMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_manifest_matrix');
assert.equal(buildMeetingAppAdapterSpec('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterSpecTemplate({ adapter_key: 'whereby' }).adapter_key, 'whereby');
assert.equal(buildMeetingAppAdapterSpecMatrix({
  platforms: ['google-meet'],
  adapters: [buildMeetingAppAdapterSpecTemplate({ adapter_key: 'whereby', matches: ['https://whereby.com/*'] })],
}).accepted_count, 2);
assert.equal(assertMeetingAppAdapterSpec('zoom').adapter_key, 'zoom');
assert.equal(assertMeetingAppAdapterSpecMatrix({ platforms: ['google-meet'] }).spec_count, 1);
assert.equal(kit.meetingAppAdapterSpec('google-meet').source, 'built_in_manifest');
assert.equal(kit.meetingAppAdapterSpecMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_spec_matrix');
assert.equal(buildMeetingAppAdapterRuntimeConfig('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterRuntimeConfig('google-meet').browser_runtime_options.runtimePreset, false);
assert.equal(buildMeetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet', 'zoom'] }).accepted_count, 2);
assert.equal(assertMeetingAppAdapterRuntimeConfig('zoom').contracts.timestamp_field, 'captured_at_ms');
assert.equal(assertMeetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet'] }).config_count, 1);
assert.equal(kit.meetingAppAdapterRuntimeConfig('google-meet').capture_options.participantSelectors.length > 0, true);
assert.equal(kit.meetingAppAdapterRuntimeConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_runtime_config_matrix');
assert.equal(buildMeetingAppAdapterHandoffPackage('google-meet').accepted, true);
assert.equal(buildMeetingAppAdapterHandoffPackage('google-meet').file_paths.includes('runtime-config.json'), true);
assert.equal(buildMeetingAppAdapterHandoffPackage('google-meet').file_paths.includes('verification-plan.json'), true);
assert.equal(buildMeetingAppAdapterHandoffPackageMatrix({ platforms: ['google-meet', 'zoom'] }).accepted_count, 2);
assert.equal(assertMeetingAppAdapterHandoffPackage('zoom').contracts.timestamp_field, 'captured_at_ms');
assert.equal(assertMeetingAppAdapterHandoffPackageMatrix({ platforms: ['google-meet'] }).package_count, 1);
assert.equal(kit.meetingAppAdapterHandoffPackage('google-meet').validation.content_script_ready, true);
assert.equal(kit.meetingAppAdapterHandoffPackageMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_adapter_handoff_package_matrix');
const appAdapterEvidence = {
  live_dom_snapshot_count: 1,
  candidate_observation_count: 1,
  speaker_segments: [{ captured_at_ms: 1_782_614_401_000, speaker_label: 'Ada' }],
  participant_segments: [{ captured_at_ms: 1_782_614_401_000, participant_label: 'Ada' }],
  annotations: [{ captured_at_ms: 1_782_614_402_000, meeting_id: 'gm', axis_id: 'axis-1' }],
};
assert.equal(buildMeetingAppAdapterVerificationReport('google-meet', { evidence: appAdapterEvidence }).accepted, true);
assert.equal(buildMeetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: appAdapterEvidence },
}).production_ready_count, 1);
assert.equal(assertMeetingAppAdapterVerificationReport('zoom', { evidence: appAdapterEvidence }).production_ready, true);
assert.equal(assertMeetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: appAdapterEvidence },
}).accepted_count, 1);
assert.equal(kit.meetingAppAdapterVerificationReport('google-meet', { evidence: appAdapterEvidence }).pilot_ready, true);
assert.equal(kit.meetingAppAdapterVerificationReportMatrix({
  platforms: ['google-meet'],
  evidenceByAdapter: { google_meet: appAdapterEvidence },
}).schema, 'meeting_app_adapter_verification_report_matrix');
assert.equal(buildMeetingAppObserverSchedulerConfig('google-meet').schema, 'meeting_app_observer_scheduler_config');
assert.equal(buildMeetingAppObserverSchedulerConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_app_observer_scheduler_config_matrix');
assert.equal(typeof createMeetingAppObserverScheduler({
  async sample() {
    return { ok: true };
  },
}, 'google-meet').handleTrigger, 'function');
assert.equal(buildMeetingPlatformRuntimeHostConfig('google-meet').schema, 'meeting_platform_runtime_host_config');
assert.equal(buildMeetingPlatformRuntimeHostConfigMatrix({ platforms: ['google-meet'] }).schema, 'meeting_platform_runtime_host_config_matrix');
assert.equal(buildMeetingPlatformRuntimeHostHandoff('google-meet').schema, 'meeting_platform_runtime_host_handoff');
assert.equal(buildMeetingPlatformRuntimeHostHandoffMatrix({ platforms: ['google-meet'] }).schema, 'meeting_platform_runtime_host_handoff_matrix');
assert.equal(createMeetingPlatformRuntimeHostFixtureEnvironment('google-meet').platform, 'google_meet');
assert.equal(createMeetingPlatformRuntimeHostVerificationClient().getState().call_count, 0);
assert.equal((await runMeetingPlatformRuntimeHostVerification('google-meet')).accepted, true);
assert.equal((await runMeetingPlatformRuntimeHostVerificationMatrix({ platforms: ['google-meet', 'zoom'] })).accepted_count, 2);
const replayInput = {
  records: [
    {
      platform: 'google_meet',
      phase: 'active',
      captured_at_ms: 1_783_356_000_000,
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Leave call' }],
        participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
      },
    },
    {
      platform: 'google_meet',
      phase: 'ended',
      captured_at_ms: 1_783_356_002_000,
      url: 'https://meet.google.com/abc-defg-hij',
      page: {
        controls: [{ label: 'Join now' }],
      },
    },
  ],
};
assert.equal((await runMeetingPlatformRuntimeHostReplay('google-meet', replayInput)).accepted, true);
assert.equal((await runMeetingPlatformRuntimeHostReplayMatrix({ platforms: ['google-meet'], input: replayInput })).accepted_count, 1);
assert.equal(typeof createMeetingPlatformRuntimeHost({
  async sample() {
    return { ok: true };
  },
}, 'google-meet', {
  setInterval() {
    return {};
  },
  clearInterval() {},
}).changed, 'function');
assert.equal(resolveMeetingAppRuntimeAdapterProfile('https://example.com/not-a-meeting').detected, false);
assert.equal(selectMeetingAppRuntimeAdapter('https://meet.google.com/abc-defg-hij').selected, true);
assert.equal(buildMeetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').surface, 'browser_extension');
assert.equal(buildMeetingAppRuntimeAdapterHandoffMatrix({ platforms: ['google-meet'], surfaces: ['native-detector'] }).ready_count, 1);
assert.equal(buildMeetingAppRuntimeAdapterHostPackage({ platforms: ['google-meet'], surfaces: ['browser-extension'] }).schema, 'meeting_app_runtime_adapter_host_package');
assert.equal(buildMeetingAppRuntimeAdapterHandoffAcceptanceReport('https://meet.google.com/abc-defg-hij').accepted, true);
assert.equal(buildMeetingAppRuntimeAdapterHandoffMatrixAcceptanceReport({ platforms: ['google-meet'], surfaces: ['native-detector'] }).accepted, true);
assert.equal(assertMeetingAppRuntimeAdapterHandoff('https://meet.google.com/abc-defg-hij').accepted, true);
assert.equal(buildMeetingAppRuntimeAdapterProfileMatrix({ platforms: ['google-meet', 'teams'] }).platform_count, 2);
assert.equal(kit.platformTimelineViewPlan('zoom').schema, 'meeting_platform_timeline_view_plan');
assert.equal(kit.platformTimelineViewMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformTimelineView('zoom', {
  meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'note-1',
      label: 'why?',
      captured_at_ms: 1_782_614_401_000,
    },
  ],
}).diagnostics.marker_count, 1);
assert.equal(kit.zoomPlatformTimelineViewport({
  start_ms: 0,
  duration_ms: 600_000,
  full_duration_ms: 600_000,
}, 2).duration_ms, 300_000);
assert.equal(kit.platformAnnotationIntakePlan('zoom').schema, 'meeting_platform_annotation_intake_plan');
assert.equal(kit.platformAnnotationIntakeMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformAnnotationIntake('zoom', {
  current_meeting: {
    platform: 'zoom',
    meeting_id: '987654321',
    start_time_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-2',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'ready_to_insert_current_axis');
assert.equal(kit.platformClockSyncPlan('zoom').schema, 'meeting_platform_clock_sync_plan');
assert.equal(kit.platformClockSyncMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformClockSync('zoom', {
  clock_sync: {
    offset_ms: 10,
    rtt_ms: 40,
  },
}).status, 'clock_sync_ready');
assert.equal(kit.platformSessionBindingPlan('zoom').schema, 'meeting_platform_session_binding_plan');
assert.equal(kit.platformSessionBindingMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformSessionBinding('zoom', {
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: 1_782_614_400_000,
  },
}).status, 'open_axis_from_local_observer');
assert.equal(kit.platformRealtimeAnnotationPlan('zoom').schema, 'meeting_platform_realtime_annotation_plan');
assert.equal(kit.platformRealtimeAnnotationMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformRealtimeAnnotation('zoom', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://zoom.us/j/987654321',
    observed_at_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-realtime-smoke',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'start_axis_then_insert');
assert.equal(kit.platformArtifactHandoffPlan('zoom').schema, 'meeting_platform_artifact_handoff_plan');
assert.equal(kit.platformArtifactHandoffMatrix({
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(kit.platformArtifactHandoff('zoom', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'zoom', meeting_id: '987654321' },
      occurred_at_ms: 1_782_614_400_000,
      artifact_kind: 'transcript',
      artifact_url: 'https://zoom.us/transcript.vtt',
    },
  ],
}).fetch_request_count, 1);
assert.equal(kit.platformAdaptationPackage('zoom').schema, 'meeting_platform_adaptation_package');
assert.equal(kit.platformAdaptationPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformAdaptationPackage('google-meet', {
  baseUrl: 'http://localhost:8787',
}).extension.matches.includes('https://meet.google.com/*'), true);
assert.equal(buildMeetingPlatformAdaptationPackageMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].sdk_wiring_ready, true);
assert.equal(kit.platformRuntimeBundle('google-meet').schema, 'meeting_platform_runtime_bundle');
assert.equal(kit.platformRuntimeBundleMatrix({
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(kit.platformRuntimeHostConfig('google-meet').schema, 'meeting_platform_runtime_host_config');
assert.equal(kit.platformRuntimeHostConfigMatrix({
  platforms: ['google-meet'],
}).host_ready_count, 1);
assert.equal(kit.platformRuntimeHostHandoff('google-meet').acceptance.accepted, true);
assert.equal(kit.platformRuntimeHostHandoffMatrix({
  platforms: ['google-meet'],
}).accepted_count, 1);
assert.equal((await kit.replayPlatformRuntimeHost('google-meet', replayInput)).accepted, true);
assert.equal((await kit.replayPlatformRuntimeHostMatrix({
  platforms: ['google-meet'],
  input: replayInput,
})).accepted_count, 1);
assert.equal(kit.platformRuntimeEventPlan('google-meet').schema, 'meeting_platform_runtime_event_plan');
assert.equal(kit.platformRuntimeEventPlan('google-meet').realtime_contract.provider_events_required_for_realtime, false);
assert.equal(kit.platformRuntimeEventPlanMatrix({
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformRuntimeEventPlan('zoom', {
  baseUrl: 'http://localhost:8787',
}).examples.insert_annotation.action, 'insert_annotation');
assert.equal(buildMeetingPlatformRuntimeEventPlanMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).transcript_realtime_dependency_count, 0);
const smokeGoogleRuntimeProfile = buildMeetingPlatformRuntimeProfile('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(smokeGoogleRuntimeProfile.adapter_surfaces.primary, 'browser_extension');
assert.equal(smokeGoogleRuntimeProfile.evidence_thresholds.pilot.provider_records_required, false);
assert.equal(smokeGoogleRuntimeProfile.launch_requirements.provider_events_block_launch, false);
assert.equal(buildMeetingPlatformRuntimeProfileMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
}).rows.find((row) => row.platform === 'zoom').primary_surface, 'native_detector');
const smokeGoogleRuntimeBundle = buildMeetingPlatformRuntimeBundle('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(smokeGoogleRuntimeBundle.browser.matches.includes('https://meet.google.com/*'), true);
assert.equal(smokeGoogleRuntimeBundle.adapter_route.routes[0].route, 'local_observer_axis');
assert.equal(smokeGoogleRuntimeBundle.runtime.content_script_bridge.install_function, 'installMeetingPlatformIntegrationContentScriptBridge');
assert.equal(smokeGoogleRuntimeBundle.runtime.lightweight_connector_bridge.install_function, 'installMeetingPlatformConnectorContentScriptBridge');
assert.equal(smokeGoogleRuntimeBundle.messaging.bridge_message_types.includes('meeting_timeline.candidate_launch_plan'), true);
assert.equal(smokeGoogleRuntimeBundle.messaging.bridge_message_types.includes('meeting_timeline.open_candidate_session'), true);
assert.equal(smokeGoogleRuntimeBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.sample_tracks'), true);
assert.equal(smokeGoogleRuntimeBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.preflight_current_window'), true);
assert.equal(smokeGoogleRuntimeBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.candidate_launch_plan'), false);
assert.equal(smokeGoogleRuntimeBundle.messaging.lightweight_connector_message_types.includes('meeting_timeline.open_candidate_session'), false);
assert.equal(smokeGoogleRuntimeBundle.messaging.background_message_types.includes('meeting_timeline.preflight_candidates'), true);
assert.equal(smokeGoogleRuntimeBundle.messaging.background_message_types.includes('meeting_timeline.candidate_launch_plan'), false);
assert.equal(smokeGoogleRuntimeBundle.messaging.background_message_types.includes('meeting_timeline.open_candidate_session'), false);
assert.equal(smokeGoogleRuntimeBundle.messaging.runtime_event.plan.realtime_contract.transcript_required_for_realtime, false);
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.preflight_current_window.type, 'meeting_timeline.preflight_current_window');
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.preflight_candidates.type, 'meeting_timeline.preflight_candidates');
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.candidate_launch_plan.type, 'meeting_timeline.candidate_launch_plan');
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.open_candidate_session.type, 'meeting_timeline.open_candidate_session');
assert.equal(smokeGoogleRuntimeBundle.messaging.examples.content_script_insert_annotation.type, 'meeting_timeline.insert_mark');
assert.equal(buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].adapter_first_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].provider_required_for_realtime, false);
assert.equal(buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).lightweight_connector_ready_count, 1);

const rollout = buildMeetingPlatformRolloutPlan('teams', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(rollout.platform, 'microsoft_teams');

const runbook = buildMeetingPlatformAdaptationRunbook('webex', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(runbook.type, 'meeting_platform_adaptation_runbook');
assert.equal(runbook.steps.some((item) => item.id === 'validate_rollout'), true);

const strategy = buildMeetingPlatformAdaptationStrategy('google-meet', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(strategy.platform, 'google_meet');
assert.equal(strategy.realtime_axis.provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformAdaptationStrategyMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].platform, 'zoom');
assert.equal(buildMeetingPlatformEvidenceCorrelation('zoom', {}).status, 'single_source');
const session = createMeetingPlatformEvidenceSession('zoom', {
  baseUrl: 'http://localhost:8787',
});
assert.equal(session.summary().platform, 'zoom');
assert.equal(session.summary().package_ready_for_handoff, false);
assert.equal(kit.platformEvidenceSession('google-meet').platform, 'google_meet');
const liveAdapter = createMeetingPlatformLiveAdapter('google-meet', client, {
  baseUrl: 'http://localhost:8787',
});
assert.equal(liveAdapter.summary().platform, 'google_meet');
assert.equal(buildMeetingPlatformLiveAdapterPlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).live_adapter.kit_method, 'platformLiveAdapter');
assert.equal(buildMeetingPlatformLiveAdapterMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].platform, 'zoom');
assert.equal(buildMeetingPlatformLiveAdapterReadiness('zoom', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_live_adapter_readiness');
assert.equal(buildMeetingPlatformLiveAdapterReadinessMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(buildMeetingPlatformLiveAdapterHandoff('zoom', {
  baseUrl: 'http://localhost:8787',
}).host_contract.annotation_timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingPlatformLiveAdapterHandoffBundle({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_live_readiness, 'npm run meeting-platform:live-readiness');
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadiness('zoom', {
    baseUrl: 'http://localhost:8787',
  }),
  /Meeting platform live adapter readiness failed/,
);
assert.throws(
  () => assertMeetingPlatformLiveAdapterReadinessMatrix({
    baseUrl: 'http://localhost:8787',
    platforms: ['zoom'],
  }),
  /Meeting platform live adapter readiness matrix failed/,
);
assert.equal(createMeetingPlatformLiveAdapterSuite(client, {
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).summary().platform_count, 1);
assert.equal(kit.platformLiveAdapter('zoom').platform, 'zoom');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platforms[0], 'zoom');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).adapter_route_matrix.rows[0].first_route, 'local_observer_axis');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.platform_conformance, '/api/meeting-platform/conformance');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.runtime_bundles, '/api/meeting-platform/runtime-bundles');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.extension_plan, '/api/meeting-platform/extension-plan');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.integration_runtime, '/api/meeting-platform/integration-runtime');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.integration_runtime_run_manifest, '/api/meeting-platform/integration-runtime/run-manifest');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.handoff_readiness, '/api/meeting-platform/handoff-readiness');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).endpoints.runtime_events, '/api/meeting-platform/runtime-events');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_platform_conformance, 'npm run meeting-platform:conformance');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_integration_runtime_manifest, 'npm run meeting-platform:integration-runtime-run-manifest');
assert.equal(buildMeetingPlatformHostIntegrationPlan({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).commands.validate_handoff_readiness, 'npm run meeting-platform:handoff-readiness');
const smokeHostScaffold = buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
});
assert.equal(assertMeetingPlatformHostIntegrationScaffold(buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
})).accepted, true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/run-integration-runtime-manifest.mjs'), true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/verify-conformance.mjs'), true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/run-handoff-readiness.mjs'), true);
assert.equal(smokeHostScaffold.files.some((file) => file.path === 'scripts/print-platform-adapters.mjs'), true);
assert.equal(smokeHostScaffold.files.find((file) => file.path === 'src/platform-adapters/zoom.mjs').content.includes('captured_at_ms'), true);
assert.equal(smokeHostScaffold.files.find((file) => file.path === 'README.md').content.includes('runIntegrationRuntimeManifest'), true);
assert.equal(buildMeetingPlatformProviderConnectionPack('zoom', {
  baseUrl: 'http://localhost:8787',
}).realtime_annotation_policy.annotation_timestamp_field, 'captured_at_ms');
assert.equal(buildMeetingPlatformProviderConnectionMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).packs[0].official_docs.some((doc) => doc.url.includes('zoom.us')), true);
assert.equal(buildMeetingPlatformSubscriptionHandoff('zoom', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_subscription_handoff');
assert.equal(buildMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).platform_count, 1);
assert.equal(buildMeetingPlatformSpeakerTrackPlan('google-meet').output_contract.transcript_required, false);
assert.equal(buildMeetingPlatformSpeakerTrackMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformSpeakerTrack('google-meet', {
  signals: [
    {
      type: 'speaker_started',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      speaker_name: 'Ada',
    },
    {
      type: 'speaker_ended',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_402_000,
      speaker_name: 'Ada',
    },
  ],
}).marks[0].intent, 'speaker_track');
assert.equal(buildMeetingPlatformParticipantTrackPlan('google-meet').transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformParticipantTrackMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformParticipantTrack('google-meet', {
  signals: [
    {
      type: 'participant_joined',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      participant_id: 'ada',
      participant_name: 'Ada',
    },
  ],
}).marks[0].intent, 'participant_track');
assert.equal(buildMeetingPlatformTimelineViewPlan('google-meet').output_contract.renderer_agnostic, true);
assert.equal(buildMeetingPlatformTimelineViewMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformTimelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotations: [
    {
      id: 'note-1',
      label: 'why?',
      captured_at_ms: 1_782_614_401_000,
    },
  ],
}).markers[0].x_ratio, 0.001667);
assert.equal(zoomMeetingPlatformTimelineViewport({
  start_ms: 0,
  duration_ms: 600_000,
  full_duration_ms: 600_000,
}, 2).duration_ms, 300_000);
assert.equal(buildMeetingPlatformAnnotationIntakePlan('google-meet').realtime_policy.transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformAnnotationIntakeMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-intake-1',
    label: 'why?',
    captured_at_ms: 1_782_614_403_000,
  },
}).normalized_time_ms, 3_000);
assert.equal(buildMeetingPlatformClockSyncPlan('google-meet').required, true);
assert.equal(buildMeetingPlatformClockSyncMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformClockSyncReport('google-meet', {
  samples: [
    {
      client_send_at_ms: 1_782_614_400_000,
      server_time_ms: 1_782_614_400_060,
      client_receive_at_ms: 1_782_614_400_100,
    },
  ],
}).recommended_offset_ms, 10);
assert.equal(buildMeetingPlatformSessionBindingPlan('google-meet').realtime_policy.provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformSessionBindingMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformSessionBinding('google-meet', {
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_400_000,
  },
}).status, 'open_axis_from_local_observer');
assert.equal(buildMeetingPlatformRealtimeAnnotationPlan('google-meet').pipeline.includes('clock_sync'), true);
assert.equal(buildMeetingPlatformRealtimeAnnotationMatrix({
  platforms: ['google-meet'],
}).provider_blocking_count, 0);
assert.equal(buildMeetingPlatformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 0,
    rtt_ms: 40,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: 1_782_614_400_000,
  },
  annotation: {
    id: 'note-realtime-direct',
    label: 'why?',
    captured_at_ms: 1_782_614_402_000,
  },
}).status, 'start_axis_then_insert');
assert.equal(buildMeetingPlatformArtifactHandoffPlan('google-meet').transcript_blocks_realtime, false);
assert.equal(buildMeetingPlatformArtifactHandoffMatrix({
  platforms: ['google-meet'],
}).realtime_blocking_count, 0);
assert.equal(buildMeetingPlatformArtifactHandoff('google-meet', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: 1_782_614_400_000,
      artifact_kind: 'transcript',
      artifact_id: 'transcript-1',
    },
  ],
}).transcript_import_count, 1);
assert.equal(buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'http://localhost:8787',
}).supported_surfaces.browser_observer, true);
assert.equal(buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'http://localhost:8787',
}).adapter_surfaces.primary, 'browser_extension');
assert.equal(buildMeetingPlatformAdapterAuthoringPlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).browser_surface.matches.includes('https://meet.google.com/*'), true);
assert.equal(buildMeetingPlatformAdapterAuthoringMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'Acme Rooms'],
}).external_authoring_count, 1);
assert.equal(buildMeetingPlatformAdapterAuthoringPlanFromRoot('google-meet', {
  baseUrl: 'http://localhost:8787',
}).provider_reconcile.path, 'google_workspace_events_pubsub');
assert.equal(buildMeetingPlatformAdapterPortfolioItem('google-meet', {
  baseUrl: 'http://localhost:8787',
}).p1_provider_reconcile.official_doc_count, 3);
assert.equal(buildMeetingPlatformAdapterPortfolio({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'Acme Rooms'],
}).external_authoring_count, 1);
assert.equal(buildMeetingPlatformAdapterPortfolioFromRoot({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).pilot_ready_count, 1);
assert.equal(buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterAcceptanceChecklistMatrix({
  platforms: ['google-meet'],
}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).accepted_count, 1);
assert.equal(buildMeetingPlatformAdapterAcceptanceChecklistFromRoot('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).artifact_refs.runtime_bundle.path, 'google_meet/runtime-bundle.json');
assert.equal(buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet'],
}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).export_ready_count, 1);
assert.equal(buildMeetingPlatformAdapterExportPackageFromRoot('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
}).surface_entrypoints.provider_reconcile.blocks_realtime_annotation, false);
const directGoogleExportPackage = buildMeetingPlatformAdapterExportPackage('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  target: 'static',
});
const directGoogleExportFiles = directGoogleExportPackage.host_files.map((file) => file.path);
assert.equal(buildMeetingPlatformAdapterImportPlan(directGoogleExportPackage, {
  availableFiles: directGoogleExportFiles,
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterImportPlanMatrix([directGoogleExportPackage], {
  availableFiles: directGoogleExportFiles,
}).accepted_count, 1);
assert.equal(buildMeetingPlatformAdapterImportPlanFromRoot(directGoogleExportPackage, {
  availableFiles: directGoogleExportFiles,
}).runtime_contract.local_axis_first, true);
assert.equal(assertMeetingPlatformAdapterImportPlan(directGoogleExportPackage, {
  availableFiles: directGoogleExportFiles,
}).accepted, true);
const directGoogleImportPlan = buildMeetingPlatformAdapterImportPlan(directGoogleExportPackage, {
  availableFiles: directGoogleExportFiles,
});
assert.equal(buildMeetingPlatformAdapterInstallManifest([directGoogleImportPlan]).accepted, true);
assert.equal(buildMeetingPlatformAdapterInstallManifestFromRoot([directGoogleImportPlan]).browser_extension.platform_count, 1);
assert.equal(assertMeetingPlatformAdapterInstallManifest(buildMeetingPlatformAdapterInstallManifest([directGoogleImportPlan])).accepted, true);
const directInstallManifest = buildMeetingPlatformAdapterInstallManifest([directGoogleImportPlan]);
assert.equal(buildMeetingPlatformAdapterLaunchPlan(directInstallManifest, {
  url: 'https://meet.google.com/abc-defg-hij',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterLaunchPlanFromRoot(directInstallManifest, {
  platform: 'google-meet',
}).runtime_actions[0].id, 'observe_platform_candidates');
assert.equal(buildMeetingPlatformAdapterCandidateLaunchPlanFromRoot(directInstallManifest, {
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: [rootGoogleActiveSnapshot],
  }],
}, {
  requireSpeakerTrack: true,
}).status, 'ready_for_realtime_launch');
assert.equal(buildMeetingPlatformAdapterCandidateLaunchPlan(directInstallManifest, {
  tabs: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: [rootGoogleActiveSnapshot],
  }],
}, {
  requireSpeakerTrack: true,
}).launch_plan.platform, 'google_meet');
assert.equal(assertMeetingPlatformAdapterLaunchPlan(buildMeetingPlatformAdapterLaunchPlan(directInstallManifest, {
  platform: 'google-meet',
})).accepted, true);
const directGoogleLaunchPlan = buildMeetingPlatformAdapterLaunchPlan(directInstallManifest, {
  platform: 'google-meet',
});
assert.equal(buildMeetingPlatformAdapterSessionHandoff(directGoogleLaunchPlan).schema, 'meeting_platform_adapter_session_handoff');
assert.equal(buildMeetingPlatformAdapterSessionHandoffFromRoot(directGoogleLaunchPlan).session_factory, 'createMeetingPlatformAdapterSession');
assert.equal(createMeetingPlatformAdapterSession(directGoogleLaunchPlan, {
  async observePlatformCandidates() {
    return { ok: true };
  },
  async insertAnnotation() {
    return { ok: true };
  },
}).schema, 'meeting_platform_adapter_session');
assert.equal(createMeetingPlatformAdapterSessionFromRoot(directGoogleLaunchPlan, {
  async observePlatformCandidates() {
    return { ok: true };
  },
  async insertAnnotation() {
    return { ok: true };
  },
}).platform, 'google_meet');
const directRunnerClient = {
  async observePlatformCandidates() {
    return { ok: true };
  },
  async insertAnnotation() {
    return { ok: true };
  },
};
assert.equal(buildMeetingPlatformAdapterRunnerHandoff(directInstallManifest).schema, 'meeting_platform_adapter_runner_handoff');
assert.equal(buildMeetingPlatformAdapterRunnerHandoffFromRoot(directInstallManifest).runner_factory, 'createMeetingPlatformAdapterRunner');
const directRunner = createMeetingPlatformAdapterRunner(directInstallManifest, directRunnerClient, {
  clock: () => 901,
});
assert.equal(directRunner.schema, 'meeting_platform_adapter_runner');
assert.equal((await directRunner.open({ url: 'https://meet.google.com/abc-defg-hij' })).payload.observe_event.captured_at_ms, 901);
assert.equal(createMeetingPlatformAdapterRunnerFromRoot(directInstallManifest, directRunnerClient).schema, 'meeting_platform_adapter_runner');
assert.equal((await openMeetingPlatformAdapterSession(directInstallManifest, directRunnerClient, {
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  clock: () => 902,
})).payload.observe_event.captured_at_ms, 902);
assert.equal((await openMeetingPlatformAdapterSessionFromRoot(directInstallManifest, directRunnerClient, {
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  clock: () => 903,
})).payload.observe_event.captured_at_ms, 903);
assert.equal(buildMeetingPlatformAdapterMessageBridgeHandoff(directInstallManifest).schema, 'meeting_platform_adapter_message_bridge_handoff');
assert.equal(buildMeetingPlatformAdapterMessageBridgeHandoffFromRoot(directInstallManifest).bridge_factory, 'createMeetingPlatformAdapterMessageBridge');
const directMessageBridge = createMeetingPlatformAdapterMessageBridge(directInstallManifest, directRunnerClient, {
  clock: () => 904,
});
assert.equal((await directMessageBridge.handleMessage({
  type: 'meeting_timeline.open_session',
  payload: { url: 'https://meet.google.com/abc-defg-hij' },
})).result.payload.observe_event.captured_at_ms, 904);
assert.equal(createMeetingPlatformAdapterMessageBridgeFromRoot(directInstallManifest, directRunnerClient).schema, 'meeting_platform_adapter_message_bridge');
assert.equal(buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'http://localhost:8787',
}).annotations.endpoints.runtimeEvents, 'http://localhost:8787/api/meeting-platform/runtime-events');
assert.equal(buildMeetingPlatformAdapterContractMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).contracts[0].adapter_surfaces.primary, 'native_detector');
assert.equal(buildMeetingPlatformAdapterContractAcceptanceReport('google-meet', {
  baseUrl: 'http://localhost:8787',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterContractAcceptanceMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).accepted_count, 1);
assert.equal(assertMeetingPlatformAdapterContract('zoom', {
  baseUrl: 'http://localhost:8787',
}).accepted, true);
assert.equal(buildMeetingPlatformAdapterSamplePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).platform, 'google_meet');
assert.equal((await runMeetingPlatformAdapterSample('zoom', {
  baseUrl: 'http://localhost:8787',
})).accepted, true);
assert.equal((await runMeetingPlatformAdapterSampleMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['webex'],
})).accepted_count, 1);
assert.equal(buildMeetingPlatformRealEvidenceIntakePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).provider_endpoint, 'http://localhost:8787/api/platform-events/google-meet');
assert.equal(buildMeetingPlatformRealEvidenceIntakeReport('google-meet', {}, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).accepted, false);
assert.equal(buildMeetingPlatformFieldIntakePlan('google-meet', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_field_intake_plan');
assert.equal(buildMeetingPlatformFieldIntakeMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'zoom'],
}).platform_count, 2);
assert.equal(buildMeetingPlatformHandoffReadiness('google-meet', {
  baseUrl: 'http://localhost:8787',
}).schema, 'meeting_platform_handoff_readiness');
assert.equal(buildMeetingPlatformHandoffReadinessMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
}).platform_count, 1);

const evidencePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords: [],
  meetingAppRecords: [],
}, {
  baseUrl: 'http://localhost:8787',
  includeRunbook: false,
});
assert.equal(evidencePackage.schema, 'meeting_platform_evidence_package');
assert.equal(evidencePackage.platform, 'google_meet');
assert.equal(verifyMeetingPlatformEvidencePackage(evidencePackage, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).type, 'meeting_platform_evidence_package_verification');
assert.equal(buildMeetingPlatformFieldEvidenceBundle('google-meet', evidencePackage, {
  baseUrl: 'http://localhost:8787',
  requireProductionReady: false,
}).schema, 'meeting_platform_field_evidence_bundle');
assert.equal(buildMeetingPlatformFieldCaptureManifest('google-meet', {
  baseUrl: 'http://localhost:8787',
  evidencePackage,
}).schema, 'meeting_platform_field_capture_manifest');
assert.equal(buildMeetingPlatformFieldCollectorConfig('google-meet', {
  baseUrl: 'http://localhost:8787',
}).browser_observer.matches.includes('https://meet.google.com/*'), true);
assert.equal(kit.platformFieldCaptureManifest('google-meet', {
  evidencePackage,
}).file_contract.files.evidence_package.endsWith('/google_meet.json'), true);
assert.equal(kit.platformAdapterContract('google-meet').schema, 'meeting_platform_adapter_contract');
assert.equal(kit.platformAdapterContractMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformAdapterContractAcceptance('google-meet').accepted, true);
assert.equal(kit.platformAdapterContractAcceptanceMatrix({ platforms: ['zoom'] }).accepted_count, 1);
assert.equal(buildMeetingPlatformAdapterRoute('google-meet', {
  baseUrl: 'http://localhost:8787',
}).routes[0].route, 'local_observer_axis');
assert.equal(buildMeetingPlatformAdapterRouteMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).platform_count, 1);
assert.equal(buildMeetingPlatformAdapterDecision({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
}).platform, 'google_meet');
assert.equal(buildMeetingPlatformAdapterDecisionMatrix({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].provider_events_block_realtime, false);
assert.equal(buildMeetingPlatformAdapterStartupPlan({
  url: 'https://meet.google.com/abc-defg-hij',
}, {
  baseUrl: 'http://localhost:8787',
}).selected_surface, 'browser_extension');
assert.equal(buildMeetingPlatformAdapterStartupPlanMatrix({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['zoom'],
}).rows[0].insert_action, 'insertAnnotation');
assert.equal(buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: [rootGoogleActiveSnapshot],
}, {
  baseUrl: 'http://localhost:8787',
}).readiness.realtime_annotation_ready, true);
assert.equal(buildMeetingPlatformAdapterCurrentWindowPreflight({
  document: rootSmokeDocument(),
}, {
  baseUrl: 'http://localhost:8787',
  requireSpeakerTrack: true,
}).capture.profile, 'google_meet');
assert.equal(buildMeetingPlatformAdapterCandidatePreflight({
  windows: [{
    id: 'smoke-browser',
    tabs: [{
      id: 1,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Google Meet',
    }, {
      id: 2,
      document: rootSmokeDocument(),
    }],
  }],
}, {
  baseUrl: 'http://localhost:8787',
  requireSpeakerTrack: true,
}).selected_platform, 'google_meet');
assert.equal(buildMeetingPlatformAdapterPreflightMatrix({}, {
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [rootGoogleActiveSnapshot],
  },
}).live_evidence_ready_count, 1);
assert.equal(kit.platformAdapterRoute('google-meet').realtime_invariants.provider_events_block_realtime, false);
assert.equal(kit.platformAdapterRoute('google-meet').adapter_surfaces.primary, 'browser_extension');
assert.equal(kit.platformAdapterRouteMatrix({ platforms: ['zoom'] }).rows[0].first_route, 'local_observer_axis');
assert.equal(kit.platformAdapterDecision({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
}).platform, 'microsoft_teams');
assert.equal(kit.platformAdapterDecisionMatrix({}, {
  platforms: ['webex'],
}).accepted_count, 1);
assert.equal(kit.platformAdapterStartupPlan({
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
}).platform, 'microsoft_teams');
assert.equal(kit.platformAdapterStartupPlanMatrix({}, {
  platforms: ['webex'],
}).realtime_startup_ready_count, 1);
assert.equal(kit.platformAdapterPreflight({
  platform: 'google-meet',
  snapshots: [rootGoogleActiveSnapshot],
}).accepted, true);
assert.equal(kit.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [rootGoogleActiveSnapshot],
  },
}).meeting_start_ready_count, 1);
assert.equal(kit.platformAdapterCandidatePreflight({
  candidates: [{
    document: rootSmokeDocument(),
  }],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(kit.platformRealEvidenceIntakePlan('google-meet').schema, 'meeting_platform_real_evidence_intake_plan');
assert.equal(kit.platformRealEvidenceIntake('google-meet', {}, {
  requireProductionReady: false,
}).accepted, false);
assert.equal(kit.platformFieldCollectorConfig('google-meet').schema, 'meeting_platform_field_collector_config');
assert.equal(kit.platformFieldEvidenceBundle('google-meet', evidencePackage, {
  requireProductionReady: false,
}).field_capture_plan.schema, 'meeting_platform_field_capture_plan');
assert.equal(kit.platformFieldIntakePlan('google-meet').schema, 'meeting_platform_field_intake_plan');
assert.equal(kit.platformFieldIntakeMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(kit.platformHandoffReadiness('google-meet').schema, 'meeting_platform_handoff_readiness');
assert.equal(kit.platformHandoffReadinessMatrix({ platforms: ['zoom'] }).platform_count, 1);
assert.equal(buildMeetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet'],
}).schema, 'meeting_app_dom_adaptation_diagnosis_matrix');
assert.equal(kit.meetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet'],
}).platform_count, 1);

const gate = buildMeetingAppLaunchGate('google-meet', {
  allowFixtureEvidence: true,
  requireProductionReady: false,
});
assert.equal(gate.platform, 'google_meet');

const signals = normalizeGoogleMeetEvent({
  type: 'google.workspace.meet.conference.v2.started',
  subject: '//meet.googleapis.com/conferenceRecords/demo',
  time: '2026-06-26T02:00:00.000Z',
  data: {
    conferenceRecord: {
      name: 'conferenceRecords/demo',
      space: 'spaces/demo',
      startTime: '2026-06-26T02:00:00.000Z',
    },
  },
});
assert.equal(signals.some((signal) => signal.type === 'meeting_started'), true);

console.log('ok consumer package imports');
process.exit(0);
`, 'utf8');

const { stdout: smokeStdout } = await execFileAsync(process.execPath, ['smoke.mjs'], {
  cwd: consumerDir,
});
assert.match(smokeStdout, /ok consumer package imports/);

const installedPackage = JSON.parse(await readFile(
  join(consumerDir, 'node_modules/@ai-annotation/meeting-timeline-sdk/package.json'),
  'utf8',
));
assert.equal(installedPackage.version, '0.1.0');

console.log('ok meeting timeline SDK package smoke');
