import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-import-plan.mjs';
import {
  buildMeetingPlatformAdapterInstallManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-install-manifest.mjs';
import {
  assertMeetingPlatformAdapterSmoke,
  runMeetingPlatformAdapterSmoke,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-smoke.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const directReport = await runMeetingPlatformAdapterSmoke({
  baseUrl,
  platforms: ['google-meet', 'zoom', 'teams'],
  captured_at_ms: 1_782_900_000_000,
});
assert.equal(directReport.schema, 'meeting_platform_adapter_smoke_report');
assert.equal(directReport.accepted, true);
assert.equal(directReport.platform_count, 3);
assert.equal(directReport.accepted_count, 3);
assert.deepEqual(directReport.platforms, ['google_meet', 'zoom', 'microsoft_teams']);
assert.equal(directReport.runtime_contract.timestamp_field, 'captured_at_ms');
assert.equal(directReport.runtime_contract.transcript_required_for_smoke, false);
assert.equal(directReport.verified_sequence.includes('meeting_timeline.insert_mark'), true);
assert.equal(directReport.rows.every((row) => row.observe_before_insert === true), true);
assert.equal(directReport.rows.every((row) => row.captured_at_ms_preserved === true), true);
assert.equal(directReport.rows.every((row) => row.speaker_track_inserted === true), true);
assert.equal(directReport.rows.every((row) => row.participant_track_inserted === true), true);
assert.equal(directReport.rows.every((row) => row.provider_reconcile_nonblocking === true), true);

const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'webex', 'lark'],
}, {
  baseUrl,
  target: 'static',
});
const availableFiles = exportMatrix.packages.flatMap((pkg) => pkg.host_files.map((file) => file.path));
const importMatrix = buildMeetingPlatformAdapterImportPlanMatrix(exportMatrix.packages, {
  availableFiles,
});
const manifest = buildMeetingPlatformAdapterInstallManifest(importMatrix.plans, {
  baseUrl,
});
assert.equal(manifest.accepted, true);

const manifestReport = await runMeetingPlatformAdapterSmoke(manifest, {
  captured_at_ms: 1_782_900_100_000,
});
assert.equal(manifestReport.accepted, true);
assert.deepEqual(manifestReport.platforms, ['google_meet', 'webex', 'lark']);
assert.equal(manifestReport.rows.find((row) => row.platform === 'lark').fixture_url, 'https://vc.feishu.cn/j/123456789');

const asserted = await assertMeetingPlatformAdapterSmoke(manifest, {
  captured_at_ms: 1_782_900_200_000,
});
assert.equal(asserted.accepted, true);

const kitClient = {
  async startMeeting(input) {
    return { ok: true, input };
  },
  async endMeeting(input) {
    return { ok: true, input };
  },
  async insertMark(input) {
    return { ok: true, input };
  },
};
const kit = createMeetingPlatformTimelineKit(kitClient, {
  baseUrl,
});
const kitReport = await kit.platformAdapterSmoke(manifest, {
  captured_at_ms: 1_782_900_300_000,
});
assert.equal(kitReport.accepted, true);
assert.equal((await kit.assertPlatformAdapterSmoke(manifest)).accepted, true);

const rootSdk = createMeetingAppTimelineSdk({
  baseUrl,
  fetch: async () => new Response(JSON.stringify({ ok: true })),
  platforms: ['google-meet'],
});
const rootReport = await rootSdk.adapterSmoke({
  platforms: ['google-meet'],
  captured_at_ms: 1_782_900_400_000,
});
assert.equal(rootReport.accepted, true);
assert.equal(rootReport.rows[0].platform, 'google_meet');
assert.equal((await rootSdk.assertAdapterSmoke({
  platforms: ['google-meet'],
  captured_at_ms: 1_782_900_500_000,
})).rows[0].platform, 'google_meet');

const badManifest = {
  ...manifest,
  accepted: false,
};
await assert.rejects(
  () => assertMeetingPlatformAdapterSmoke(badManifest, { platforms: ['google-meet'] }),
  /Meeting platform adapter smoke failed/,
);

console.log('ok meeting platform adapter smoke');
