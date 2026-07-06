import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterExportPackageMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-export-package.mjs';
import {
  buildMeetingPlatformAdapterImportPlan,
  buildMeetingPlatformAdapterImportPlanMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-import-plan.mjs';
import {
  buildMeetingPlatformAdapterInstallManifest,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-install-manifest.mjs';
import {
  assertMeetingPlatformAdapterLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-launch-plan.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';
const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'zoom'],
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

const googlePlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Design review',
  capturedAtMs: 1_782_614_400_000,
});
assert.equal(googlePlan.schema, 'meeting_platform_adapter_launch_plan');
assert.equal(googlePlan.accepted, true);
assert.equal(googlePlan.platform, 'google_meet');
assert.equal(googlePlan.detection_reason, 'meeting_url');
assert.equal(googlePlan.selected_surface, 'browser_extension');
assert.equal(googlePlan.detected_meeting.meeting_id, 'abc-defg-hij');
assert.equal(googlePlan.surface_entrypoint.content_script.matches.includes('https://meet.google.com/*'), true);
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'observe_platform_candidates').sdk_method, 'observePlatformCandidates');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'insert_realtime_annotation').sdk_method, 'insertAnnotation');
assert.equal(googlePlan.mark_template.platform, 'google_meet');
assert.equal(googlePlan.mark_template.captured_at_ms, 1_782_614_400_000);
assert.equal(assertMeetingPlatformAdapterLaunchPlan(googlePlan), googlePlan);

const zoomPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  platform: 'zoom',
  url: 'https://example.com/not-a-meeting',
});
assert.equal(zoomPlan.accepted, true);
assert.equal(zoomPlan.platform, 'zoom');
assert.equal(zoomPlan.detection_reason, 'explicit_platform');

const googleExportPackage = exportMatrix.packages.find((pkg) => pkg.platform === 'google_meet');
const googleWebviewImportPlan = buildMeetingPlatformAdapterImportPlan(googleExportPackage, {
  availableFiles: googleExportPackage.host_files.map((file) => file.path),
  surface: 'webview-preload',
});
const webviewManifest = buildMeetingPlatformAdapterInstallManifest([googleWebviewImportPlan], {
  baseUrl,
});
const webviewPlan = buildMeetingPlatformAdapterLaunchPlan(webviewManifest, {
  platform: 'google-meet',
});
assert.equal(webviewPlan.accepted, true);
assert.equal(webviewPlan.selected_surface, 'webview_preload');
assert.equal(webviewPlan.surface_entrypoint.first_sdk_method, 'observeMeetingApp_or_observePlatformCandidates');

const unregisteredSurfacePlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  platform: 'google-meet',
  surface: 'webview-preload',
});
assert.equal(unregisteredSurfacePlan.accepted, false);
assert.equal(unregisteredSurfacePlan.readiness.issues.some((issue) => issue.code === 'surface_not_registered_for_platform'), true);

const providerPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  platform: 'google-meet',
  surface: 'provider-reconcile',
});
assert.equal(providerPlan.accepted, false);
assert.equal(providerPlan.readiness.issues.some((issue) => issue.code === 'provider_reconcile_not_realtime_launch_surface'), true);

const missingPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc',
});
assert.equal(missingPlan.accepted, false);
assert.equal(missingPlan.platform, 'microsoft_teams');
assert.equal(missingPlan.readiness.issues.some((issue) => issue.code === 'platform_not_installed'), true);

const client = {
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
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformAdapterLaunchPlan(manifest, { url: 'https://meet.google.com/abc-defg-hij' }).accepted, true);
assert.equal(kit.assertPlatformAdapterLaunchPlan(googlePlan), googlePlan);
assert.equal(kit.report().platform_adapter_launch_plan.accepted, false);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterLaunchPlan(manifest, { url: 'https://meet.google.com/abc-defg-hij' }).accepted, true);
assert.equal(sdk.adapterLaunchPlan(manifest, { platform: 'zoom' }).platform, 'zoom');
assert.equal(sdk.assertPlatformAdapterLaunchPlan(googlePlan), googlePlan);

console.log('ok meeting platform adapter launch plan');
