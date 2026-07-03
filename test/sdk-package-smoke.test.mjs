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
assert.equal(packedFiles.includes('adapters/platform-kit.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-kit.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-rollout.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-strategy.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-correlation.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-session.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-live-adapter.d.ts'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.mjs'), true);
assert.equal(packedFiles.includes('adapters/platform-evidence-package.d.ts'), true);
assert.equal(packedFiles.includes('README.md'), true);
assert.equal(packedFiles.some((item) => item.startsWith('test/')), false);
assert.equal(packedFiles.some((item) => item.startsWith('scripts/')), false);

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

await writeFile(join(consumerDir, 'smoke.mjs'), `
import assert from 'node:assert/strict';
import {
  SDK_VERSION,
  createMeetingTimelineClient,
} from '@ai-annotation/meeting-timeline-sdk';
import {
  createMeetingPlatformTimelineKit,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit';
import {
  buildMeetingPlatformAdaptationRunbook,
  buildMeetingPlatformRolloutPlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';
import {
  buildMeetingPlatformAdaptationStrategy,
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy';
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
  normalizeGoogleMeetEvent,
} from '@ai-annotation/meeting-timeline-sdk/adapters/google-meet';
import {
  buildMeetingAppLaunchGate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate';

assert.equal(SDK_VERSION, '0.1.0');
const client = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787',
  fetch: async () => new Response(JSON.stringify({ ok: true }), {
    headers: { 'content-type': 'application/json' },
  }),
});
assert.equal(typeof client.startMeeting, 'function');

const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'http://localhost:8787',
  verify: false,
});
assert.equal(kit.platformRolloutPlan('google-meet').platform, 'google_meet');
assert.equal(kit.platformAdaptationRunbook('zoom').platform, 'zoom');
assert.equal(kit.report({ platforms: ['google-meet'] }).platform_rollout.type, 'meeting_platform_rollout_summary');
assert.equal(kit.platformLiveAdapterHandoff('zoom').sdk.factory, 'createMeetingPlatformLiveAdapter');
assert.equal(kit.platformLiveAdapterHandoffBundle({
  platforms: ['zoom'],
}).platform_count, 1);

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
