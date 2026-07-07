import assert from 'node:assert/strict';

import {
  MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA,
  MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA,
  MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA,
  assertMeetingPlatformRegistryManifest,
  buildMeetingPlatformRegistryAcceptanceReport,
  buildMeetingPlatformRegistryEntry,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from '../packages/meeting-timeline-sdk/adapters/platform-registry.mjs';
import {
  createMeetingPlatformTimelineKit,
} from '../packages/meeting-timeline-sdk/adapters/platform-kit.mjs';

const baseUrl = 'https://timeline.example.com';

const googleAdapter = meetingPlatformEventAdapterFor('google-meet');
assert.equal(googleAdapter.key, 'google_meet');
assert.equal(googleAdapter.source, 'google_meet_webhook');
assert.equal(typeof googleAdapter.normalize, 'function');

const google = buildMeetingPlatformRegistryEntry('google-meet', { baseUrl });
assert.equal(google.schema, MEETING_PLATFORM_REGISTRY_ENTRY_SCHEMA);
assert.equal(google.platform, 'google_meet');
assert.equal(google.aliases.includes('google-meet'), true);
assert.equal(google.event_adapter.normalize_available, true);
assert.equal(google.event_adapter.module, '@ai-annotation/meeting-timeline-sdk/adapters/google-meet');
assert.equal(google.runtime.runtime_ready, true);
assert.equal(google.runtime.browser_matches.includes('https://meet.google.com/*'), true);
assert.equal(google.runtime.sample_interval_ms, 10_000);
assert.equal(google.runtime.candidate_observation.ready, true);
assert.equal(google.runtime.candidate_observation.message_type, 'meeting_timeline.observe_candidates');
assert.equal(google.runtime.candidate_observation.required_permission, 'tabs');
assert.equal(google.runtime.candidate_observation.endpoint, '/api/meeting-platform/observe-candidates');
assert.equal(google.runtime.runtime_event_plan_schema, 'meeting_platform_runtime_event_plan');
assert.equal(google.runtime.runtime_event_action_count, 16);
assert.equal(google.provider.required_for_realtime, false);
assert.equal(google.provider.security_verifier, 'verifyGooglePubSubOidcJwt');
assert.equal(google.provider.start_events.includes('google.workspace.meet.conference.v2.started'), true);
assert.equal(google.annotations.insert_endpoint, `${baseUrl}/api/annotations`);
assert.equal(google.annotations.runtime_event_endpoint, `${baseUrl}/api/meeting-platform/runtime-events`);
assert.equal(google.annotations.runtime_event_plan.supported_actions.includes('insert_annotation'), true);
assert.equal(google.annotations.runtime_event_plan.supported_actions.includes('observe_platform_candidates'), true);
assert.equal(google.annotations.runtime_event_plan.supported_actions.includes('run_manifest'), true);
assert.equal(google.annotations.runtime_event_plan.supported_actions.includes('run_handoff_readiness'), true);
assert.equal(google.annotations.runtime_event_plan.realtime_contract.provider_events_required_for_realtime, false);
assert.equal(google.annotations.candidate_observation_runtime_action, 'observe_platform_candidates');
assert.equal(google.annotations.timestamp_field, 'captured_at_ms');
assert.equal(google.annotations.provider_events_block_realtime, false);
assert.equal(google.annotations.transcript_blocks_realtime, false);
assert.equal(google.adapter_route.schema, 'meeting_platform_adapter_route');
assert.equal(google.adapter_route.recommended_mode, 'local_observer_first_provider_reconcile');
assert.equal(google.adapter_route.first_route, 'local_observer_axis');
assert.equal(google.adapter_route.provider_events_block_realtime, false);
assert.equal(google.adapter_route.transcript_blocks_realtime, false);
assert.equal(google.adapter_blueprint.schema, 'meeting_platform_adapter_blueprint');
assert.equal(google.adapter_blueprint.ready, true);
assert.equal(google.adapter_blueprint.primary_surface, 'browser_extension');
assert.equal(google.adapter_blueprint.provider_blocks_realtime, false);
assert.equal(google.adapter_blueprint.realtime_axis_timestamp_field, 'captured_at_ms');
assert.equal(google.transcript.blocks_realtime_annotation, false);
assert.equal(google.host.endpoints.adapter_routes, '/api/meeting-platform/adapter-routes');
assert.equal(google.host.endpoints.adapter_blueprints, '/api/meeting-platform/adapter-blueprints');
assert.equal(google.host.endpoints.runtime_bundles, '/api/meeting-platform/runtime-bundles');
assert.equal(google.host.endpoints.runtime_event_plans, '/api/meeting-platform/runtime-event-plans');
assert.equal(google.host.endpoints.runtime_events, '/api/meeting-platform/runtime-events');
assert.equal(google.host.endpoints.platform_candidate_observation, '/api/meeting-platform/observe-candidates');
assert.equal(google.sdk.imports.runtime_bundle, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle');
assert.equal(google.sdk.imports.runtime_event, '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event');
assert.equal(google.sdk.imports.adapter_route, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route');
assert.equal(google.sdk.imports.adapter_blueprint, '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint');
assert.equal(google.readiness.contract_accepted, true);
assert.equal(google.readiness.candidate_observation_ready, true);
assert.equal(google.readiness.provider_required_for_realtime, false);
assert.equal(google.readiness.transcript_blocks_realtime, false);
assert.equal(google.commands.print_registry.includes('meeting-platform:registry'), true);
assert.equal(google.commands.print_adapter_route.includes('meeting-platform:adapter-route'), true);
assert.equal(google.commands.print_adapter_blueprint.includes('meeting-platform:adapter-blueprint'), true);
assert.equal(google.commands.print_runtime_event_plan.includes('meeting-platform:runtime-event-plan'), true);

const local = buildMeetingPlatformRegistryEntry('local-detector', { baseUrl });
assert.equal(local.platform, 'local_detector');
assert.equal(local.runtime.browser_matches.length, 0);
assert.equal(local.event_adapter.source, 'local_detector');
assert.equal(local.readiness.runtime_ready, true);

const manifest = buildMeetingPlatformRegistryManifest({
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
assert.equal(manifest.schema, MEETING_PLATFORM_REGISTRY_MANIFEST_SCHEMA);
assert.equal(manifest.platform_count, 5);
assert.equal(manifest.normalizer_count, 5);
assert.equal(manifest.runtime_ready_count, 5);
assert.equal(manifest.contract_accepted_count, 5);
assert.equal(manifest.candidate_observer_count, 5);
assert.equal(manifest.adapter_blueprint_ready_count, 5);
assert.equal(manifest.provider_required_for_realtime_count, 0);
assert.equal(manifest.transcript_blocking_count, 0);
assert.equal(manifest.rows.find((row) => row.platform === 'microsoft_teams').provider_transport, 'Microsoft Graph change notifications');
assert.equal(manifest.rows.find((row) => row.platform === 'microsoft_teams').candidate_observation_ready, true);
assert.equal(manifest.rows.find((row) => row.platform === 'microsoft_teams').candidate_observer_message_type, 'meeting_timeline.observe_candidates');
assert.equal(manifest.rows.find((row) => row.platform === 'microsoft_teams').candidate_observer_permission, 'tabs');
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').runtime_event_action_count, 16);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_route_mode, 'local_observer_first_provider_reconcile');
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_first_route, 'local_observer_axis');
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_ready, true);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_primary_surface, 'browser_extension');
assert.deepEqual(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_surface_order, ['browser_extension', 'desktop_observer', 'provider_reconcile']);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_provider_blocks_realtime, false);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_transcript_blocks_realtime, false);
assert.equal(manifest.rows.find((row) => row.platform === 'google_meet').adapter_blueprint_first_acceptance_gate, 'local_candidate_preflight_accepts_active_meeting');
assert.equal(manifest.rows.find((row) => row.platform === 'zoom').browser_match_count, 3);
assert.equal(manifest.next_actions.includes('export_runtime_event_plan_before_wiring_external_host'), true);
assert.equal(manifest.next_actions.includes('export_adapter_route_before_wiring_external_host'), true);
assert.equal(manifest.next_actions.includes('read_adapter_blueprint_before_wiring_external_host'), true);

const acceptance = buildMeetingPlatformRegistryAcceptanceReport(manifest);
assert.equal(acceptance.schema, MEETING_PLATFORM_REGISTRY_ACCEPTANCE_SCHEMA);
assert.equal(acceptance.accepted, true);
assert.equal(acceptance.blocking_count, 0);
assert.equal(acceptance.candidate_observer_count, 5);
assert.equal(acceptance.adapter_blueprint_ready_count, 5);
assert.equal(assertMeetingPlatformRegistryManifest(manifest).accepted, true);

const brokenManifest = {
  ...manifest,
  normalizer_count: 4,
  entries: manifest.entries.map((entry) => entry.platform === 'zoom'
    ? {
      ...entry,
      annotations: {
        ...entry.annotations,
        timestamp_field: 'server_received_at_ms',
      },
    }
    : entry),
};
const brokenAcceptance = buildMeetingPlatformRegistryAcceptanceReport(brokenManifest);
assert.equal(brokenAcceptance.accepted, false);
assert.equal(brokenAcceptance.issues.some((item) => item.code === 'missing_platform_normalizer'), true);
assert.equal(brokenAcceptance.issues.some((item) => item.code === 'entry_invalid_timestamp_field'), true);

const missingRuntimeEventPlan = {
  ...manifest,
  entries: manifest.entries.map((entry) => entry.platform === 'google_meet'
    ? {
      ...entry,
      annotations: {
        ...entry.annotations,
        runtime_event_plan: {
          ...entry.annotations.runtime_event_plan,
          supported_actions: entry.annotations.runtime_event_plan.supported_actions.filter((action) => action !== 'insert_annotation'),
        },
      },
      sdk: {
        ...entry.sdk,
        imports: {
          ...entry.sdk.imports,
          runtime_event: undefined,
        },
      },
      host: {
        ...entry.host,
        endpoints: {
          ...entry.host.endpoints,
          runtime_event_plans: undefined,
        },
      },
    }
    : entry),
};
const missingRuntimeEventPlanAcceptance = buildMeetingPlatformRegistryAcceptanceReport(missingRuntimeEventPlan);
assert.equal(missingRuntimeEventPlanAcceptance.accepted, false);
assert.equal(missingRuntimeEventPlanAcceptance.issues.some((item) => item.code === 'entry_missing_runtime_event_import'), true);
assert.equal(missingRuntimeEventPlanAcceptance.issues.some((item) => item.code === 'entry_missing_runtime_event_plan_endpoint'), true);
assert.equal(missingRuntimeEventPlanAcceptance.issues.some((item) => item.code === 'entry_missing_runtime_insert_action'), true);

const missingAdapterRoute = {
  ...manifest,
  entries: manifest.entries.map((entry) => entry.platform === 'google_meet'
    ? {
      ...entry,
      adapter_route: undefined,
      sdk: {
        ...entry.sdk,
        imports: {
          ...entry.sdk.imports,
          adapter_route: undefined,
        },
      },
      host: {
        ...entry.host,
        endpoints: {
          ...entry.host.endpoints,
          adapter_routes: undefined,
        },
      },
    }
    : entry),
};
const missingAdapterRouteAcceptance = buildMeetingPlatformRegistryAcceptanceReport(missingAdapterRoute);
assert.equal(missingAdapterRouteAcceptance.accepted, false);
assert.equal(missingAdapterRouteAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_route'), true);
assert.equal(missingAdapterRouteAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_route_import'), true);
assert.equal(missingAdapterRouteAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_route_endpoint'), true);

const missingAdapterBlueprint = {
  ...manifest,
  adapter_blueprint_ready_count: 4,
  entries: manifest.entries.map((entry) => entry.platform === 'google_meet'
    ? {
      ...entry,
      adapter_blueprint: undefined,
      sdk: {
        ...entry.sdk,
        imports: {
          ...entry.sdk.imports,
          adapter_blueprint: undefined,
        },
      },
      host: {
        ...entry.host,
        endpoints: {
          ...entry.host.endpoints,
          adapter_blueprints: undefined,
        },
      },
    }
    : entry),
};
const missingAdapterBlueprintAcceptance = buildMeetingPlatformRegistryAcceptanceReport(missingAdapterBlueprint);
assert.equal(missingAdapterBlueprintAcceptance.accepted, false);
assert.equal(missingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'adapter_blueprint_not_ready'), true);
assert.equal(missingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_blueprint'), true);
assert.equal(missingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_blueprint_import'), true);
assert.equal(missingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'entry_missing_adapter_blueprint_endpoint'), true);
assert.equal(missingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'entry_adapter_blueprint_invalid_timestamp_field'), true);

const blockingAdapterBlueprint = {
  ...manifest,
  entries: manifest.entries.map((entry) => entry.platform === 'zoom'
    ? {
      ...entry,
      adapter_blueprint: {
        ...entry.adapter_blueprint,
        provider_blocks_realtime: true,
        transcript_blocks_realtime: true,
      },
    }
    : entry),
};
const blockingAdapterBlueprintAcceptance = buildMeetingPlatformRegistryAcceptanceReport(blockingAdapterBlueprint);
assert.equal(blockingAdapterBlueprintAcceptance.accepted, false);
assert.equal(blockingAdapterBlueprintAcceptance.issues.some((item) => item.code === 'entry_adapter_blueprint_blocks_realtime'), true);

const missingCandidateObservation = {
  ...manifest,
  candidate_observer_count: 4,
  entries: manifest.entries.map((entry) => entry.platform === 'google_meet'
    ? {
      ...entry,
      runtime: {
        ...entry.runtime,
        candidate_observation: {
          ...entry.runtime.candidate_observation,
          ready: false,
          message_type: 'wrong.message',
          required_permission: undefined,
        },
      },
      readiness: {
        ...entry.readiness,
        candidate_observation_ready: false,
      },
      annotations: {
        ...entry.annotations,
        runtime_event_plan: {
          ...entry.annotations.runtime_event_plan,
          supported_actions: entry.annotations.runtime_event_plan.supported_actions.filter((action) => action !== 'observe_platform_candidates'),
        },
      },
      host: {
        ...entry.host,
        endpoints: {
          ...entry.host.endpoints,
          platform_candidate_observation: undefined,
        },
      },
    }
    : entry),
};
const missingCandidateObservationAcceptance = buildMeetingPlatformRegistryAcceptanceReport(missingCandidateObservation);
assert.equal(missingCandidateObservationAcceptance.accepted, false);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'candidate_observer_not_ready'), true);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'entry_candidate_observation_not_ready'), true);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'entry_missing_candidate_observation_action'), true);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'entry_missing_candidate_observation_endpoint'), true);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'entry_invalid_candidate_observation_message_type'), true);
assert.equal(missingCandidateObservationAcceptance.issues.some((item) => item.code === 'entry_invalid_candidate_observation_permission'), true);

assert.throws(
  () => assertMeetingPlatformRegistryManifest(brokenManifest),
  /Meeting platform registry manifest failed acceptance/,
);

const client = {
  async startMeeting(input) { return { ok: true, input }; },
  async endMeeting(input) { return { ok: true, input }; },
  async insertMark(input) { return { ok: true, input }; },
};
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(kit.platformRegistryEntry('zoom').platform, 'zoom');
assert.equal(kit.platformRegistryManifest().platform_count, 2);
assert.equal(kit.platformRegistryAcceptance({ platforms: ['zoom'] }).accepted, true);
assert.equal(kit.assertPlatformRegistryManifest({ platforms: ['zoom'] }).accepted, true);
assert.equal(kit.report().platform_registry_manifest.provider_required_for_realtime_count, 0);
assert.equal(kit.report().platform_registry_manifest.candidate_observer_count, 2);
assert.equal(kit.report().platform_registry_manifest.adapter_blueprint_ready_count, 2);

console.log('ok meeting platform registry manifest');
