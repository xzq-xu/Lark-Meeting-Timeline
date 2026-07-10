import assert from 'node:assert/strict';

import {
  buildMeetingPlatformAdapterAcceptanceChecklist,
  buildMeetingPlatformAdapterAcceptanceChecklistMatrix,
  buildMeetingPlatformPilotMeasurementContract,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-acceptance-checklist.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';
import {
  createMeetingAppTimelineSdk,
} from '../packages/meeting-timeline-sdk/index.mjs';

const baseUrl = 'https://timeline.example.com';

const googlePilotContract = buildMeetingPlatformPilotMeasurementContract('google-meet', { baseUrl });
assert.equal(googlePilotContract.platform, 'google_meet');
assert.equal(googlePilotContract.required_consecutive_real_meetings, 3);
assert.equal(googlePilotContract.thresholds.annotation_visible_latency_ms.lte, 300);
assert.equal(googlePilotContract.observer_surface, 'browser_extension');

const zoomBrowserPilotContract = buildMeetingPlatformPilotMeasurementContract('zoom', {
  baseUrl,
  observerSurface: 'browser_extension',
});
assert.equal(zoomBrowserPilotContract.observer_surface, 'browser_extension');
assert.equal(zoomBrowserPilotContract.thresholds.meeting_start_detection_latency_ms.lte, 1_500);
assert.equal(zoomBrowserPilotContract.thresholds.meeting_end_detection_latency_ms.lte, 6_500);

const googleNativePilotContract = buildMeetingPlatformPilotMeasurementContract('google-meet', {
  baseUrl,
  observerSurface: 'native_detector',
});
assert.equal(googleNativePilotContract.observer_surface, 'native_detector');
assert.equal(googleNativePilotContract.thresholds.meeting_start_detection_latency_ms.lte, 2_500);
assert.equal(googleNativePilotContract.thresholds.meeting_end_detection_latency_ms.lte, 7_500);

const staticGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'static',
});
assert.equal(staticGoogle.schema, 'meeting_platform_adapter_acceptance_checklist');
assert.equal(staticGoogle.platform, 'google_meet');
assert.equal(staticGoogle.target, 'static');
assert.equal(staticGoogle.accepted, true);
assert.equal(staticGoogle.summary.blocking_count, 0);
assert.equal(staticGoogle.checklist.find((item) => item.id === 'captured_at_ms_contract').passed, true);
assert.equal(staticGoogle.checklist.find((item) => item.id === 'real_local_observer_evidence').status, 'skip');
assert.equal(staticGoogle.runtime_event_contract.timestamp_field, 'captured_at_ms');
assert.equal(staticGoogle.runtime_event_contract.observe_before_insert_required, true);
assert.equal(staticGoogle.runtime_event_contract.provider_events_block_realtime, false);
assert.equal(staticGoogle.runtime_event_contract.speaker_position_markers.transcript_text_required, false);
assert.equal(staticGoogle.implementation_sequence[0].id, 'resolve_adapter_checklist');
assert.equal(
  staticGoogle.implementation_sequence.some((step) => step.id === 'install_browser_extension_or_webview_preload'),
  true,
);
assert.equal(
  staticGoogle.implementation_sequence.some((step) => step.id === 'emit_speaker_position_markers_after_filtering'),
  true,
);
assert.equal(staticGoogle.sdk_entrypoints.observe_candidates, 'sdk.observePlatformCandidates(input)');
assert.equal(staticGoogle.evidence_collection_plan.pilot.includes('insert_annotation_current_axis'), true);
assert.equal(staticGoogle.pilot_measurement_contract.required_consecutive_real_meetings, 3);
assert.equal(staticGoogle.pilot_measurement_contract.annotations_per_meeting, 5);
assert.equal(staticGoogle.pilot_measurement_contract.provider_events_required, false);
assert.equal(staticGoogle.pilot_measurement_contract.thresholds.meeting_start_detection_latency_ms.lte, 1_500);
assert.equal(staticGoogle.pilot_measurement_contract.thresholds.annotation_visible_latency_ms.lte, 300);
assert.equal(staticGoogle.pilot_measurement_contract.thresholds.meeting_end_detection_latency_ms.lte, 6_500);
assert.equal(staticGoogle.next_actions.includes('capture_real_meeting_app_snapshots'), true);

const pilotGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'pilot',
});
assert.equal(pilotGoogle.accepted, false);
assert.equal(pilotGoogle.summary.failed_required_ids.includes('real_local_observer_evidence'), true);
assert.equal(pilotGoogle.summary.failed_required_ids.includes('handoff_ready'), true);
assert.equal(pilotGoogle.checklist.find((item) => item.id === 'provider_reconcile_evidence').status, 'skip');

const productionGoogle = buildMeetingPlatformAdapterAcceptanceChecklist('google-meet', {}, {
  baseUrl,
  target: 'production',
});
assert.equal(productionGoogle.accepted, false);
assert.equal(productionGoogle.summary.failed_required_ids.includes('production_ready'), true);
assert.equal(productionGoogle.summary.failed_required_ids.includes('provider_reconcile_evidence'), true);

const custom = buildMeetingPlatformAdapterAcceptanceChecklist('Acme Rooms', {}, {
  baseUrl,
  target: 'static',
  displayName: 'Acme Rooms',
  browserMatches: ['https://meet.acme.example/*'],
  providerPath: 'acme_rooms_webhooks',
});
assert.equal(custom.platform, 'acme_rooms');
assert.equal(custom.accepted, false);
assert.equal(custom.summary.failed_required_ids.includes('adapter_registered_or_authorable'), true);
assert.equal(custom.handoff_readiness, undefined);
assert.equal(custom.next_actions.includes('add_platform_setup_entry'), true);

const matrix = buildMeetingPlatformAdapterAcceptanceChecklistMatrix({
  platforms: ['google-meet', 'Acme Rooms'],
}, {
  baseUrl,
  target: 'static',
});
assert.equal(matrix.schema, 'meeting_platform_adapter_acceptance_checklist_matrix');
assert.equal(matrix.target, 'static');
assert.equal(matrix.platform_count, 2);
assert.equal(matrix.accepted_count, 1);
assert.equal(matrix.blocked_count, 1);
const googleRow = matrix.rows.find((row) => row.platform === 'google_meet');
assert.equal(googleRow.accepted, true);
assert.equal(googleRow.runtime_contract_timestamp_field, 'captured_at_ms');
assert.equal(googleRow.provider_reconcile_nonblocking, true);
assert.equal(googleRow.pilot_required_real_meetings, 3);
assert.equal(googleRow.annotation_visible_p95_lte_ms, 300);
assert.equal(googleRow.meeting_start_max_lte_ms, 1_500);
assert.equal(googleRow.meeting_end_max_lte_ms, 6_500);
assert.equal(googleRow.first_implementation_step, 'resolve_adapter_checklist');
assert.equal(googleRow.local_observer_install_step, 'install_browser_extension_or_webview_preload');
assert.equal(matrix.rows.find((row) => row.platform === 'acme_rooms').failed_required_ids.includes('adapter_registered_or_authorable'), true);

const zoomStatic = buildMeetingPlatformAdapterAcceptanceChecklist('zoom', {}, {
  baseUrl,
  target: 'static',
});
assert.equal(
  zoomStatic.implementation_sequence.some((step) => step.id === 'wire_native_window_or_accessibility_detector'),
  true,
);
assert.equal(zoomStatic.pilot_measurement_contract.thresholds.meeting_start_detection_latency_ms.lte, 2_500);
assert.equal(zoomStatic.pilot_measurement_contract.thresholds.meeting_end_detection_latency_ms.lte, 7_500);
assert.equal(zoomStatic.runtime_event_contract.transcript_blocks_realtime, false);

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
assert.equal(kit.platformAdapterAcceptanceChecklist('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(kit.platformAdapterAcceptanceChecklistMatrix({}, { target: 'static' }).accepted_count, 2);
assert.equal(kit.report().platform_adapter_acceptance_checklist_matrix.platform_count, 2);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterAcceptanceChecklist('google-meet', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.adapterAcceptanceChecklist('zoom', {}, { target: 'static' }).accepted, true);
assert.equal(sdk.platformAdapterAcceptanceChecklistMatrix({}, { target: 'static' }).accepted_count, 2);
assert.equal(sdk.adapterAcceptanceChecklistMatrix({ platforms: ['google-meet', 'Acme Rooms'] }, { target: 'static' }).blocked_count, 1);

console.log('ok meeting platform adapter acceptance checklist');
