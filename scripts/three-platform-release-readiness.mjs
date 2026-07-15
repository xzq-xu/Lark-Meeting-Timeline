import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

export const THREE_PLATFORM_RELEASE_CAPABILITIES = Object.freeze([
  Object.freeze({
    platform: 'google_meet',
    installable: true,
    startable: true,
    web_ready: true,
    desktop_ready: false,
  }),
  Object.freeze({
    platform: 'microsoft_teams',
    installable: true,
    startable: true,
    web_ready: true,
    desktop_ready: true,
  }),
  Object.freeze({
    platform: 'zoom',
    installable: true,
    startable: true,
    web_ready: true,
    desktop_ready: true,
  }),
]);

export function buildThreePlatformReleaseReadiness(liveAcceptance = {}) {
  const liveRows = new Map((liveAcceptance.rows ?? []).map((row) => {
    try { return [normalizeMeetingPlatform(row.platform), row]; } catch { return [String(row.platform ?? ''), row]; }
  }));
  const adapters = THREE_PLATFORM_RELEASE_CAPABILITIES.map((capability) => {
    const evidence = liveRows.get(capability.platform);
    const realMeetingAccepted = evidence?.core_accepted === true
      || (evidence?.accepted === true && evidence?.production_ready === true);
    const speakerAccepted = evidence?.speaker_accepted === true
      || evidence?.full_production_ready === true;
    return {
      ...capability,
      real_meeting_accepted: realMeetingAccepted,
      speaker_real_meeting_accepted: speakerAccepted,
      production_ready: realMeetingAccepted,
      full_production_ready: realMeetingAccepted && speakerAccepted,
      meeting_id: evidence?.meeting_id ?? null,
      evidence_file: evidence?.file ?? null,
      failed_check_ids: evidence?.failed_core_check_ids ?? evidence?.failed_check_ids ?? ['missing_real_meeting_evidence'],
      failed_speaker_check_ids: evidence?.failed_speaker_check_ids ?? (speakerAccepted ? [] : ['missing_real_meeting_evidence']),
    };
  });
  const productionReady = liveAcceptance.production_ready === true
    && adapters.every((row) => row.production_ready);
  const fullProductionReady = productionReady
    && liveAcceptance.full_production_ready === true
    && adapters.every((row) => row.full_production_ready);
  return {
    adapters,
    production_ready: productionReady,
    speaker_ready: adapters.length > 0 && adapters.every((row) => row.speaker_real_meeting_accepted),
    full_production_ready: fullProductionReady,
    remaining_platforms: adapters.filter((row) => !row.production_ready).map((row) => row.platform),
    remaining_speaker_platforms: adapters.filter((row) => !row.speaker_real_meeting_accepted).map((row) => row.platform),
    remaining_gate: productionReady
      ? null
      : 'Capture real meeting_started, realtime annotation, and meeting_ended evidence for each remaining platform.',
    speaker_remaining_gate: fullProductionReady
      ? null
      : 'Capture a stable remote speaker marker and its visible latency for each remaining platform; this supplementary gate does not block the core timeline release.',
  };
}
