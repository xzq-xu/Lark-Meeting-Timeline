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
    const realMeetingAccepted = evidence?.accepted === true && evidence?.production_ready === true;
    return {
      ...capability,
      real_meeting_accepted: realMeetingAccepted,
      production_ready: realMeetingAccepted,
      meeting_id: evidence?.meeting_id ?? null,
      evidence_file: evidence?.file ?? null,
      failed_check_ids: evidence?.failed_check_ids ?? ['missing_real_meeting_evidence'],
    };
  });
  const productionReady = liveAcceptance.accepted === true
    && liveAcceptance.production_ready === true
    && adapters.every((row) => row.production_ready);
  return {
    adapters,
    production_ready: productionReady,
    remaining_platforms: adapters.filter((row) => !row.production_ready).map((row) => row.platform),
    remaining_gate: productionReady
      ? null
      : 'Capture real meeting_started, speaker_started, realtime annotation, and meeting_ended evidence for each remaining platform.',
  };
}
