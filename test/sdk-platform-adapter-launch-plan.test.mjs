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
  assertMeetingPlatformAdapterCandidateLaunchPlan,
  assertMeetingPlatformAdapterLaunchPlan,
  buildMeetingPlatformAdapterCandidateLaunchPlan,
  buildMeetingPlatformAdapterLaunchPlan,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-launch-plan.mjs';
import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
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

const googleActiveSnapshot = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_782_614_400_000,
});
const candidateLaunchPlan = buildMeetingPlatformAdapterCandidateLaunchPlan(manifest, {
  windows: [{
    id: 'chrome-main',
    tabs: [{
      id: 1,
      url: 'https://zoom.us/j/987654321',
      title: 'Zoom Meeting',
    }, {
      id: 2,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Design review - Google Meet',
      snapshots: [googleActiveSnapshot],
    }],
  }],
}, {
  requireSpeakerTrack: true,
  capturedAtMs: 1_782_614_400_000,
});
assert.equal(candidateLaunchPlan.schema, 'meeting_platform_adapter_candidate_launch_plan');
assert.equal(candidateLaunchPlan.accepted, true);
assert.equal(candidateLaunchPlan.status, 'ready_for_realtime_launch');
assert.equal(candidateLaunchPlan.platform, 'google_meet');
assert.equal(candidateLaunchPlan.selected_candidate.tab_id, 2);
assert.equal(candidateLaunchPlan.candidate_preflight.accepted_count, 1);
assert.equal(candidateLaunchPlan.launch_plan.accepted, true);
assert.equal(candidateLaunchPlan.launch_plan.platform, 'google_meet');
assert.equal(candidateLaunchPlan.runtime_actions[0].id, 'observe_platform_candidates');
assert.equal(assertMeetingPlatformAdapterCandidateLaunchPlan(candidateLaunchPlan), candidateLaunchPlan);

const urlOnlyCandidateLaunchPlan = buildMeetingPlatformAdapterCandidateLaunchPlan(manifest, {
  tabs: [{
    id: 3,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
  }],
}, {
  requireSpeakerTrack: true,
});
assert.equal(urlOnlyCandidateLaunchPlan.accepted, false);
assert.equal(urlOnlyCandidateLaunchPlan.status, 'needs_live_page_evidence');
assert.equal(urlOnlyCandidateLaunchPlan.launch_plan.accepted, true);
assert.equal(urlOnlyCandidateLaunchPlan.readiness.issues.some((item) => item.code === 'candidate_preflight_not_accepted'), true);
assert.throws(
  () => assertMeetingPlatformAdapterCandidateLaunchPlan(urlOnlyCandidateLaunchPlan),
  /candidate launch plan is not ready/,
);

const zoomPlan = buildMeetingPlatformAdapterLaunchPlan(manifest, {
  platform: 'zoom',
  url: 'https://example.com/not-a-meeting',
});
assert.equal(zoomPlan.accepted, true);
assert.equal(zoomPlan.platform, 'zoom');
assert.equal(zoomPlan.detection_reason, 'explicit_platform');
assert.equal(zoomPlan.selected_surface, 'native_detector');
assert.equal(zoomPlan.surface_entrypoint.registry_row.platform, 'zoom');

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
assert.equal(kit.platformAdapterCandidateLaunchPlan(manifest, {
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: [googleActiveSnapshot],
  }],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(kit.assertPlatformAdapterLaunchPlan(googlePlan), googlePlan);
assert.equal(kit.assertPlatformAdapterCandidateLaunchPlan(candidateLaunchPlan), candidateLaunchPlan);
assert.equal(kit.report().platform_adapter_launch_plan.accepted, false);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterLaunchPlan(manifest, { url: 'https://meet.google.com/abc-defg-hij' }).accepted, true);
assert.equal(sdk.adapterLaunchPlan(manifest, { platform: 'zoom' }).platform, 'zoom');
assert.equal(sdk.adapterCandidateLaunchPlan(manifest, {
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: [googleActiveSnapshot],
  }],
}, {
  requireSpeakerTrack: true,
}).platform, 'google_meet');
assert.equal(sdk.assertPlatformAdapterLaunchPlan(googlePlan), googlePlan);
assert.equal(sdk.assertAdapterCandidateLaunchPlan(candidateLaunchPlan).accepted, true);

console.log('ok meeting platform adapter launch plan');
