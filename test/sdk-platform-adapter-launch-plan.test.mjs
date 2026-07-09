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

function node(tagName, attrs = {}, text = '') {
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

function selectorAttrMatches(item, selector) {
  if (selector === 'button') return item.tagName === 'BUTTON';
  const attrParts = [...String(selector).matchAll(/\[([a-zA-Z0-9_-]+)([*]?=)?(?:"([^"]*)"|'([^']*)'|([^\]\s]+))?(?:\s+i)?\]/g)];
  if (!attrParts.length) return false;
  return attrParts.every((match) => {
    const [, attrName, operator, doubleQuoted, singleQuoted, bare] = match;
    const actual = item.attributes[attrName];
    if (operator == null) return actual != null;
    if (actual == null) return false;
    const expected = doubleQuoted ?? singleQuoted ?? bare ?? '';
    if (operator === '*=') return String(actual).toLowerCase().includes(String(expected).toLowerCase());
    return String(actual) === String(expected);
  });
}

function queryNodes(nodes, selector) {
  const text = String(selector);
  if (text === '*') return nodes;
  const generic = nodes.filter((item) => selectorAttrMatches(item, text));
  if (generic.length) return generic;
  if (text.includes('speaking')) {
    return nodes.filter((item) => /speaking|active speaker|正在发言|正在讲话|正在说话/i.test(item.attributes['aria-label'] ?? ''));
  }
  if (text.includes('aria-live')) return nodes.filter((item) => item.attributes['aria-live']);
  if (text.includes('role="status"')) return nodes.filter((item) => item.attributes.role === 'status');
  return [];
}

function fakeDocument({ url, title, nodes = [] }) {
  return {
    nodeType: 9,
    title,
    hidden: false,
    location: { href: url },
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
}

const exportMatrix = buildMeetingPlatformAdapterExportPackageMatrix({
  platforms: ['google-meet', 'zoom'],
}, {
  baseUrl,
  target: 'static',
  includeArtifacts: true,
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
assert.equal(googlePlan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(googlePlan.adapter_blueprint.first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(googlePlan.raw_signal_validation.status, 'ready');
assert.equal(googlePlan.adapter_selection.ready, true);
assert.equal(googlePlan.adapter_selection.axis_source, 'local_observer_axis');
assert.equal(googlePlan.adapter_selection.axis_surface, 'browser_extension');
assert.equal(googlePlan.adapter_selection.timestamp_field, 'captured_at_ms');
assert.equal(googlePlan.axis_contract.adapter_selection_required_before_surface_install, true);
assert.equal(googlePlan.axis_contract.adapter_selection_axis_source, 'local_observer_axis');
assert.equal(googlePlan.axis_contract.adapter_selection_axis_surface, 'browser_extension');
assert.equal(googlePlan.axis_contract.raw_signal_validation_required_before_preflight, true);
assert.equal(googlePlan.detected_meeting.meeting_id, 'abc-defg-hij');
assert.equal(googlePlan.surface_entrypoint.content_script.matches.includes('https://meet.google.com/*'), true);
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'read_adapter_selection').sdk_method, 'platformAdapterSelection');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'read_adapter_selection').axis_surface, 'browser_extension');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'validate_raw_signal').sdk_method, 'platformRawSignalBatch');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'validate_raw_signal').artifact_path, 'google_meet/raw-signal-validation.json');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'observe_platform_candidates').sdk_method, 'observePlatformCandidates');
assert.equal(googlePlan.runtime_actions.find((action) => action.id === 'insert_realtime_annotation').sdk_method, 'insertAnnotation');
assert.equal(googlePlan.mark_template.platform, 'google_meet');
assert.equal(googlePlan.mark_template.axis_source, 'local_observer_axis');
assert.equal(googlePlan.mark_template.axis_surface, 'browser_extension');
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
assert.equal(candidateLaunchPlan.selection_strategy, 'score_accepted_live_current_window_then_active_candidate');
assert.equal(candidateLaunchPlan.selected_candidate_rank, 1);
assert.equal(candidateLaunchPlan.selected_candidate_score > 0, true);
assert.equal(candidateLaunchPlan.selected_candidate_reason, 'accepted_live_candidate');
assert.equal(candidateLaunchPlan.candidate_preflight.accepted_count, 1);
assert.equal(candidateLaunchPlan.selected_evidence.accepted, true);
assert.equal(candidateLaunchPlan.selected_evidence.live_evidence_ready, true);
assert.equal(candidateLaunchPlan.launch_plan.accepted, true);
assert.equal(candidateLaunchPlan.launch_plan.platform, 'google_meet');
assert.equal(candidateLaunchPlan.adapter_selection.axis_surface, 'browser_extension');
assert.equal(candidateLaunchPlan.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(candidateLaunchPlan.raw_signal_validation.status, 'ready');
assert.equal(candidateLaunchPlan.runtime_actions.find((action) => action.id === 'read_adapter_selection').axis_surface, 'browser_extension');
assert.equal(candidateLaunchPlan.runtime_actions.find((action) => action.id === 'validate_raw_signal').sdk_method, 'platformRawSignalBatch');
assert.equal(candidateLaunchPlan.runtime_actions.find((action) => action.id === 'observe_platform_candidates').id, 'observe_platform_candidates');
assert.equal(candidateLaunchPlan.launch_input.preflight_evidence.selected_candidate, undefined);
assert.equal(candidateLaunchPlan.mark_template.preflight_evidence.accepted, true);
assert.equal(assertMeetingPlatformAdapterCandidateLaunchPlan(candidateLaunchPlan), candidateLaunchPlan);

const currentWindowCandidateLaunchPlan = buildMeetingPlatformAdapterCandidateLaunchPlan(manifest, {
  candidates: [{
    active: true,
    document: fakeDocument({
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Current Google Meet',
      nodes: [
        node('button', { 'aria-label': 'Leave call' }),
        node('button', { 'aria-label': 'Turn off microphone' }),
        node('div', {
          'data-participant-id': 'ada',
          'aria-label': 'Ada Lovelace is speaking',
        }),
      ],
    }),
  }],
}, {
  requireSpeakerTrack: true,
  capturedAtMs: 1_782_614_400_000,
});
assert.equal(currentWindowCandidateLaunchPlan.accepted, true);
assert.equal(currentWindowCandidateLaunchPlan.selected_candidate_reason, 'accepted_current_window_live_candidate');
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.current_window_captured, true);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.interaction.in_call, true);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.interaction_can_leave, true);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.semantic_signal_types.includes('meeting_leave_available'), true);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.control_signal_summary.leave_available, true);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.control_signal_gap_summary.gap_count, 0);
assert.equal(currentWindowCandidateLaunchPlan.selected_evidence.active_speaker_candidate.name, 'Ada Lovelace');
assert.equal(currentWindowCandidateLaunchPlan.launch_input.preflight_evidence.active_speaker_candidate.name, 'Ada Lovelace');
assert.equal(currentWindowCandidateLaunchPlan.launch_input.control_signal_summary.active_speaker_observed, true);
assert.equal(currentWindowCandidateLaunchPlan.launch_input.control_signal_gap_summary.gap_count, 0);
assert.equal(currentWindowCandidateLaunchPlan.mark_template.preflight_evidence.interaction_can_leave, true);

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
assert.equal(zoomPlan.adapter_selection.axis_surface, 'native_detector');
assert.equal(zoomPlan.adapter_blueprint.primary_surface, 'native_detector');
assert.equal(zoomPlan.surface_entrypoint.registry_row.platform, 'zoom');

const nativeCandidateLaunchPlan = buildMeetingPlatformAdapterCandidateLaunchPlan(manifest, {
  windows: [{
    id: 'zoom-native-main',
    platform: 'zoom',
    window: {
      title: 'Zoom Meeting',
      active: true,
      visible: true,
      inMeeting: true,
      meeting_id: '987654321',
    },
    process: { name: 'Zoom Workplace' },
    audio: { call_active: true },
    observedAtMs: 1_782_614_400_000,
  }],
}, {
  capturedAtMs: 1_782_614_400_000,
});
assert.equal(nativeCandidateLaunchPlan.accepted, true);
assert.equal(nativeCandidateLaunchPlan.platform, 'zoom');
assert.equal(nativeCandidateLaunchPlan.selected_surface, 'native_detector');
assert.equal(nativeCandidateLaunchPlan.selected_candidate.window_id, 'zoom-native-main');
assert.equal(nativeCandidateLaunchPlan.selected_candidate_rank, 1);
assert.equal(nativeCandidateLaunchPlan.selected_candidate_reason, 'accepted_live_candidate');
assert.equal(nativeCandidateLaunchPlan.launch_plan.accepted, true);
assert.equal(nativeCandidateLaunchPlan.launch_plan.surface_entrypoint.registry_row.platform, 'zoom');

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
