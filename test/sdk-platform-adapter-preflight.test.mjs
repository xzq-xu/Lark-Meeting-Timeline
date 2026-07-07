import assert from 'node:assert/strict';

import {
  buildMeetingAppFixtureSnapshot,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-fixtures.mjs';
import {
  assertMeetingPlatformAdapterCandidatePreflight,
  assertMeetingPlatformAdapterCurrentWindowPreflight,
  assertMeetingPlatformAdapterPreflight,
  assertMeetingPlatformAdapterPreflightMatrix,
  buildMeetingPlatformAdapterCandidatePreflight,
  buildMeetingPlatformAdapterCurrentWindowPreflight,
  buildMeetingPlatformAdapterPreflight,
  buildMeetingPlatformAdapterPreflightMatrix,
} from '../packages/meeting-timeline-sdk/adapters/platform-adapter-preflight.mjs';
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
  if (selector.startsWith('.')) {
    return String(item.attributes.class ?? '').split(/\s+/).includes(selector.slice(1));
  }
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

function fakeDocument({ url, title, nodes = [], hidden = false }) {
  return {
    nodeType: 9,
    title,
    hidden,
    location: { href: url },
    querySelectorAll(selector) {
      return queryNodes(nodes, selector);
    },
  };
}

const googleActive = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'active',
  observedAtMs: 1_783_356_000_000,
});
const googleEnded = buildMeetingAppFixtureSnapshot('google-meet', {
  state: 'prejoin',
  observedAtMs: 1_783_356_600_000,
});

const urlOnly = buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
}, {
  baseUrl,
});
assert.equal(urlOnly.schema, 'meeting_platform_adapter_preflight');
assert.equal(urlOnly.platform, 'google_meet');
assert.equal(urlOnly.startup.realtime_startup_ready, true);
assert.equal(urlOnly.readiness.static_startup_ready, true);
assert.equal(urlOnly.readiness.live_evidence_ready, false);
assert.equal(urlOnly.accepted, false);
assert.equal(urlOnly.status, 'needs_live_page_evidence');
assert.equal(urlOnly.issues.some((issue) => issue.code === 'missing_live_page_evidence'), true);
assert.equal(urlOnly.next_actions.includes('collect_live_dom_snapshots_from_current_meeting_window'), true);

const googleLive = buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  title: 'Google Meet',
  snapshots: [googleActive],
}, {
  baseUrl,
  requireSpeakerTrack: true,
});
assert.equal(googleLive.accepted, true);
assert.equal(googleLive.status, 'ready_for_realtime_annotations');
assert.equal(googleLive.readiness.realtime_annotation_ready, true);
assert.equal(googleLive.readiness.speaker_track_ready, true);
assert.equal(googleLive.readiness.meeting_start_ready, true);
assert.equal(googleLive.readiness.meeting_end_ready, false);
assert.equal(googleLive.readiness.production_lifecycle_ready, false);
assert.equal(googleLive.summary.active_speaker_matched, true);
assert.equal(googleLive.startup.runtime_contract.transcript_blocks_realtime, false);
assert.equal(googleLive.next_actions.includes('open_adapter_session_and_insert_marks_with_captured_at_ms'), true);
assert.equal(assertMeetingPlatformAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}, {
  baseUrl,
  requireSpeakerTrack: true,
}).accepted, true);

const googleCurrentWindow = buildMeetingPlatformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Turn off microphone' }),
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
        'data-audio-level': '0.84',
      }),
      node('div', {
        'data-participant-id': 'grace',
        'aria-label': 'Grace Hopper, muted',
      }),
      node('div', { role: 'status', 'aria-live': 'polite' }, 'You are presenting'),
    ],
  }),
}, {
  baseUrl,
  observedAtMs: 1_783_356_020_000,
  requireSpeakerTrack: true,
});
assert.equal(googleCurrentWindow.accepted, true);
assert.equal(googleCurrentWindow.platform, 'google_meet');
assert.equal(googleCurrentWindow.current_window.capture_profile, 'google_meet');
assert.equal(googleCurrentWindow.capture.participant_count, 2);
assert.equal(googleCurrentWindow.readiness.realtime_annotation_ready, true);
assert.equal(googleCurrentWindow.captured_snapshot, undefined);
assert.equal(assertMeetingPlatformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Design review - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
      }),
    ],
  }),
}, {
  baseUrl,
  requireSpeakerTrack: true,
}).accepted, true);

const googleLifecycle = buildMeetingPlatformAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive, googleEnded],
}, {
  baseUrl,
  requireSpeakerTrack: true,
  requireCompleteLifecycle: true,
});
assert.equal(googleLifecycle.accepted, true);
assert.equal(googleLifecycle.status, 'ready_for_realtime_annotations_and_lifecycle');
assert.equal(googleLifecycle.readiness.meeting_end_ready, true);
assert.equal(googleLifecycle.readiness.production_lifecycle_ready, true);

const teamsActive = buildMeetingAppFixtureSnapshot('teams', {
  state: 'active',
  observedAtMs: 1_783_356_010_000,
});
const matrix = buildMeetingPlatformAdapterPreflightMatrix({}, {
  baseUrl,
  platforms: ['google-meet', 'teams', 'zoom'],
  snapshots: {
    'google-meet': [googleActive],
    teams: [teamsActive],
  },
});
assert.equal(matrix.schema, 'meeting_platform_adapter_preflight_matrix');
assert.equal(matrix.platform_count, 3);
assert.equal(matrix.accepted_count, 2);
assert.equal(matrix.realtime_ready_count, 2);
assert.equal(matrix.live_evidence_ready_count, 2);
assert.equal(matrix.rows.find((row) => row.platform === 'zoom').status, 'needs_live_page_evidence');
assert.equal(matrix.rows.find((row) => row.platform === 'microsoft_teams').realtime_annotation_ready, true);
const teamsCurrentWindow = buildMeetingPlatformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
    title: 'Weekly Sync | Microsoft Teams',
    nodes: [
      node('button', { 'aria-label': 'Leave' }),
      node('button', { 'aria-label': 'Share content' }),
      node('div', {
        'data-tid': 'participant-ada',
        'data-user-id': 'ada',
        'data-display-name': 'Ada Lovelace',
        'aria-label': 'Ada Lovelace speaking',
      }),
    ],
  }),
}, {
  baseUrl,
  requireSpeakerTrack: true,
});
assert.equal(teamsCurrentWindow.platform, 'microsoft_teams');
assert.equal(teamsCurrentWindow.accepted, true);
assert.equal(teamsCurrentWindow.capture.profile, 'microsoft_teams');

const zoomCurrentWindow = buildMeetingPlatformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://us06web.zoom.us/wc/987654321/start',
    title: 'Zoom Meeting',
    nodes: [
      node('button', { 'aria-label': 'Leave Meeting' }),
      node('button', { 'aria-label': 'Participants' }),
      node('div', {
        'data-user-id': 'mira',
        'aria-label': 'Mira Patel is speaking',
      }),
    ],
  }),
}, {
  baseUrl,
  requireSpeakerTrack: true,
  includeCapturedSnapshot: true,
});
assert.equal(zoomCurrentWindow.platform, 'zoom');
assert.equal(zoomCurrentWindow.accepted, true);
assert.equal(zoomCurrentWindow.captured_snapshot.capture.profile, 'zoom');
assert.equal(zoomCurrentWindow.capture.participant_count, 1);

const zoomNativePreflight = buildMeetingPlatformAdapterPreflight({
  platform: 'zoom',
  window: {
    id: 'zoom-native-main',
    title: 'Zoom Meeting',
    active: true,
    visible: true,
    inMeeting: true,
    meeting_id: '987654321',
  },
  process: { name: 'Zoom Workplace' },
  audio: { call_active: true },
  observedAtMs: 1_782_614_400_000,
}, {
  baseUrl,
});
assert.equal(zoomNativePreflight.accepted, true);
assert.equal(zoomNativePreflight.summary.evidence_kind, 'native_window');
assert.equal(zoomNativePreflight.summary.selected_surface, 'native_detector');
assert.equal(zoomNativePreflight.summary.native_selected_process, 'Zoom Workplace');
assert.equal(zoomNativePreflight.native_diagnosis.selected_meeting.meeting_id, '987654321');
assert.equal(zoomNativePreflight.dom_diagnosis, undefined);
assert.equal(zoomNativePreflight.readiness.meeting_start_ready, true);

const candidatePreflight = buildMeetingPlatformAdapterCandidatePreflight({
  windows: [{
    id: 'chrome-main',
    tabs: [
      {
        id: 7,
        url: 'https://meet.google.com/abc-defg-hij',
        title: 'Design review - Google Meet',
        active: false,
      },
      {
        id: 8,
        url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
        title: 'Weekly Sync | Microsoft Teams',
        active: false,
      },
      {
        id: 9,
        url: 'https://us06web.zoom.us/wc/987654321/start',
        title: 'Zoom Meeting',
        active: true,
        document: fakeDocument({
          url: 'https://us06web.zoom.us/wc/987654321/start',
          title: 'Zoom Meeting',
          nodes: [
            node('button', { 'aria-label': 'Leave Meeting' }),
            node('button', { 'aria-label': 'Participants' }),
            node('div', {
              'data-user-id': 'mira',
              'aria-label': 'Mira Patel is speaking',
            }),
          ],
        }),
      },
    ],
  }],
}, {
  baseUrl,
  requireSpeakerTrack: true,
});
assert.equal(candidatePreflight.schema, 'meeting_platform_adapter_candidate_preflight');
assert.equal(candidatePreflight.accepted, true);
assert.equal(candidatePreflight.status, 'ready_for_realtime_annotations');
assert.equal(candidatePreflight.candidate_count, 3);
assert.equal(candidatePreflight.supported_candidate_count, 3);
assert.equal(candidatePreflight.accepted_count, 1);
assert.equal(candidatePreflight.selected_candidate_index, 2);
assert.equal(candidatePreflight.selected_platform, 'zoom');
assert.deepEqual(candidatePreflight.platforms, ['google_meet', 'microsoft_teams', 'zoom']);
assert.equal(candidatePreflight.rows[0].status, 'needs_live_page_evidence');
assert.equal(candidatePreflight.rows[0].tab_id, 7);
assert.equal(candidatePreflight.rows[0].window_id, 'chrome-main');
assert.equal(candidatePreflight.rows[2].current_window_captured, true);
assert.equal(candidatePreflight.rows[2].selected, true);

const nativeCandidatePreflight = buildMeetingPlatformAdapterCandidatePreflight({
  platform: 'zoom',
  process: { name: 'Zoom Workplace' },
  audio: { call_active: true },
  windows: [{
    id: 'zoom-native-main',
    window: {
      title: 'Zoom Meeting',
      active: true,
      visible: true,
      inMeeting: true,
      meeting_id: '987654321',
    },
    observedAtMs: 1_782_614_400_000,
  }],
}, {
  baseUrl,
});
assert.equal(nativeCandidatePreflight.accepted, true);
assert.equal(nativeCandidatePreflight.candidate_count, 1);
assert.equal(nativeCandidatePreflight.selected_platform, 'zoom');
assert.equal(nativeCandidatePreflight.rows[0].selected_surface, 'native_detector');
assert.equal(nativeCandidatePreflight.rows[0].window_id, 'zoom-native-main');
assert.equal(nativeCandidatePreflight.rows[0].current_window_captured, false);
assert.equal(nativeCandidatePreflight.rows[0].meeting_start_ready, true);

assert.equal(assertMeetingPlatformAdapterCandidatePreflight({
  candidates: [{
    document: fakeDocument({
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Assert candidate - Google Meet',
      nodes: [
        node('button', { 'aria-label': 'Leave call' }),
        node('div', {
          'data-participant-id': 'ada',
          'aria-label': 'Ada Lovelace is speaking',
        }),
      ],
    }),
  }],
}, {
  baseUrl,
  requireSpeakerTrack: true,
}).selected_platform, 'google_meet');
assert.throws(
  () => assertMeetingPlatformAdapterCandidatePreflight({
    tabs: [{
      id: 10,
      url: 'https://meet.google.com/no-live-evidence',
      title: 'Google Meet',
    }],
  }, {
    baseUrl,
    requireSpeakerTrack: true,
  }),
  /candidate preflight has no accepted realtime candidate/,
);

const inputSnapshotMatrix = buildMeetingPlatformAdapterPreflightMatrix({
  snapshots: {
    'google-meet': [googleActive],
  },
}, {
  baseUrl,
  platforms: ['google-meet'],
});
assert.equal(inputSnapshotMatrix.accepted_count, 1);
assert.equal(inputSnapshotMatrix.rows[0].live_evidence_ready, true);
assert.throws(
  () => assertMeetingPlatformAdapterPreflightMatrix({}, {
    baseUrl,
    platforms: ['google-meet', 'zoom'],
    snapshots: {
      'google-meet': [googleActive],
    },
  }),
  /Meeting platform adapter preflight matrix is not accepted/,
);
assert.equal(assertMeetingPlatformAdapterPreflightMatrix({}, {
  baseUrl,
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [googleActive, googleEnded],
  },
  requireCompleteLifecycle: true,
}).accepted_count, 1);

const kit = createMeetingPlatformTimelineKit({}, {
  baseUrl,
});
assert.equal(kit.platformAdapterPreflight({
  platform: 'zoom',
}).status, 'needs_live_page_evidence');
assert.equal(kit.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet'],
  snapshots: {
    'google-meet': [googleActive],
  },
}).realtime_ready_count, 1);
assert.equal(kit.platformAdapterCandidatePreflight({
  tabs: [{
    document: fakeDocument({
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Kit candidate - Google Meet',
      nodes: [
        node('button', { 'aria-label': 'Leave call' }),
        node('div', {
          'data-participant-id': 'ada',
          'aria-label': 'Ada Lovelace is speaking',
        }),
      ],
    }),
  }],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(kit.platformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Kit preflight - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
      }),
    ],
  }),
}, {
  requireSpeakerTrack: true,
}).accepted, true);

const sdk = createMeetingAppTimelineSdk({
  baseUrl,
  platforms: ['google-meet', 'zoom'],
});
assert.equal(sdk.platformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: [googleActive],
}).accepted, true);
assert.equal(sdk.adapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}).readiness.realtime_annotation_ready, true);
assert.equal(sdk.platformAdapterPreflightMatrix({}, {
  platforms: ['google-meet', 'zoom'],
  snapshots: {
    'google-meet': [googleActive],
  },
}).accepted_count, 1);
assert.equal(sdk.adapterCandidatePreflight({
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Google Meet',
  }, {
    document: fakeDocument({
      url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_sample',
      title: 'SDK candidate - Teams',
      nodes: [
        node('button', { 'aria-label': 'Leave' }),
        node('div', {
          'data-user-id': 'ada',
          'aria-label': 'Ada Lovelace speaking',
        }),
      ],
    }),
  }],
}, {
  requireSpeakerTrack: true,
}).selected_platform, 'microsoft_teams');
assert.equal(sdk.assertAdapterCandidatePreflight({
  candidates: [{
    document: fakeDocument({
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'SDK assert candidate - Google Meet',
      nodes: [
        node('button', { 'aria-label': 'Leave call' }),
        node('div', {
          'data-participant-id': 'ada',
          'aria-label': 'Ada Lovelace is speaking',
        }),
      ],
    }),
  }],
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(sdk.assertAdapterPreflight({
  platform: 'google-meet',
  snapshots: [googleActive],
}).accepted, true);
assert.equal(sdk.platformAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'SDK preflight - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
      }),
    ],
  }),
}, {
  requireSpeakerTrack: true,
}).readiness.realtime_annotation_ready, true);
assert.equal(sdk.adapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'SDK alias preflight - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
      }),
    ],
  }),
}, {
  requireSpeakerTrack: true,
}).accepted, true);
assert.equal(sdk.assertAdapterCurrentWindowPreflight({
  document: fakeDocument({
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'SDK assert preflight - Google Meet',
    nodes: [
      node('button', { 'aria-label': 'Leave call' }),
      node('div', {
        'data-participant-id': 'ada',
        'aria-label': 'Ada Lovelace is speaking',
      }),
    ],
  }),
}, {
  requireSpeakerTrack: true,
}).accepted, true);

console.log('ok meeting platform adapter preflight');
