import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
  buildThreePlatformBrowserActionExpression,
  chooseThreePlatformBrowserAutomationAction,
} from '../scripts/three-platform-live-browser-automation.mjs';

function page(targetId, url, controls, title = 'Meeting') {
  return { target_id: targetId, url, title, controls };
}

function control(overrides = {}) {
  return { tag: 'button', visible: true, disabled: false, ...overrides };
}

const zoomStart = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  pages: [page('zoom-test', 'https://zoom.us/test', [control({ id: 'scheduleMtg', text: '开始新会议' })])],
});
assert.equal(zoomStart.phase, 'zoom_test_start');
assert.deepEqual(zoomStart.selector, { id: 'scheduleMtg' });

const zoomCurrentExperiment = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  pages: [page('zoom-test-current', 'https://zoom.us/test', [
    control({ tag: 'a', text: '加入' }),
    control({ tag: 'a', id: 'btnJoinTest', text: '加入' }),
  ])],
});
assert.equal(zoomCurrentExperiment.phase, 'zoom_test_start');
assert.deepEqual(zoomCurrentExperiment.selector, { id: 'btnJoinTest' });

const duplicateZoomStart = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  completedActionKeys: ['phase:zoom_test_start'],
  pages: [
    page('zoom-test-old', 'https://zoom.us/test', [control({ id: 'btnJoinTest', text: '加入' })]),
    page('zoom-test-new', 'https://zoom.us/test', [control({ id: 'btnJoinTest', text: '加入' })]),
  ],
});
assert.equal(duplicateZoomStart, null);

const zoomName = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  pages: [{ ...page('zoom-prejoin', 'https://zoom.us/wc/123/join', [
    control({ tag: 'input', id: 'input-for-name', type: 'text', value_present: false }),
    control({ text: '加入' }),
  ]), frame_id: 'zoom-webclient-frame' }],
});
assert.equal(zoomName.phase, 'fill_display_name');
assert.equal(zoomName.value, 'Timeline Adapter Acceptance');
assert.equal(zoomName.frame_id, 'zoom-webclient-frame');
assert.equal(zoomName.global_key, 'phase:fill_display_name');

const zoomSpeakerPeerName = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  displayName: 'Timeline Synthetic Speaker',
  pages: [page('zoom-speaker-peer', 'https://zoom.us/wc/123/join', [
    control({ tag: 'input', id: 'input-for-name', type: 'text', value_present: false }),
    control({ text: 'Join' }),
  ])],
});
assert.equal(zoomSpeakerPeerName.phase, 'fill_display_name');
assert.equal(zoomSpeakerPeerName.value, 'Timeline Synthetic Speaker');

const zoomTargetMeeting = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  meetingUrl: 'https://us05web.zoom.us/j/987654321?pwd=real-token',
  pages: [
    page('zoom-probe', 'https://app.zoom.us/wc/123456789/join', [
      control({ tag: 'input', id: 'input-for-name', type: 'text', value_present: false }),
    ]),
    page('zoom-real', 'https://app.zoom.us/wc/987654321/join', [
      control({ tag: 'input', id: 'input-for-name', type: 'text', value_present: false }),
    ]),
  ],
});
assert.equal(zoomTargetMeeting.target_id, 'zoom-real');

const zoomJoin = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  completedActionKeys: ['zoom-prejoin:fill_display_name'],
  pages: [page('zoom-prejoin', 'https://zoom.us/wc/123/join', [
    control({ tag: 'input', type: 'text', placeholder: '您的姓名', value_present: true }),
    control({ text: '加入' }),
  ])],
});
assert.equal(zoomJoin.phase, 'join_meeting');

const zoomJoinAfterGlobalFill = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  completedActionKeys: ['phase:fill_display_name'],
  pages: [page('zoom-prejoin-next', 'https://zoom.us/wc/123/join', [
    control({ tag: 'input', id: 'input-for-name', type: 'text', value_present: true }),
    control({ text: '加入' }),
  ])],
});
assert.equal(zoomJoinAfterGlobalFill.phase, 'join_meeting');

const zoomAudio = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  activeMeeting: true,
  pages: [page('zoom-live', 'https://zoom.us/wc/123/join', [control({ text: '使用电脑音频加入' })])],
});
assert.equal(zoomAudio.phase, 'join_audio');

const zoomAudioPaused = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  autoJoinAudio: false,
  activeMeeting: true,
  pages: [page('zoom-live', 'https://zoom.us/wc/123/join', [control({ text: '使用电脑音频加入' })])],
});
assert.equal(zoomAudioPaused, null);

const zoomLeave = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  autoLeave: true,
  activeMeeting: true,
  speakerSeen: true,
  leaveReady: true,
  completedActionKeys: ['zoom-live:join_audio'],
  pages: [page('zoom-live', 'https://zoom.us/wc/123/join', [control({ aria: 'Leave Meeting' })])],
});
assert.equal(zoomLeave.phase, 'leave_meeting');

const zoomCoreLeave = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  autoLeave: true,
  activeMeeting: true,
  completionSeen: true,
  leaveReady: true,
  pages: [page('zoom-core-live', 'https://zoom.us/wc/123/join', [control({ aria: 'Leave Meeting' })])],
});
assert.equal(zoomCoreLeave.phase, 'leave_meeting');

const teamsChineseLeave = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  autoLeave: true,
  activeMeeting: true,
  completionSeen: true,
  leaveReady: true,
  pages: [page('teams-live-cn', 'https://teams.live.com/v2/', [control({ aria: '退出', text: '离开' })])],
});
assert.equal(teamsChineseLeave.phase, 'leave_meeting');

const zoomConfirmLeave = chooseThreePlatformBrowserAutomationAction({
  platform: 'zoom',
  autoJoin: true,
  autoLeave: true,
  activeMeeting: true,
  speakerSeen: true,
  leaveReady: true,
  completedActionKeys: ['zoom-live:join_audio', 'zoom-live:leave_meeting'],
  pages: [page('zoom-live', 'https://zoom.us/wc/123/join', [
    control({ aria: 'Leave Meeting' }),
    control({ text: 'Leave Meeting', in_dialog: false }),
  ])],
});
assert.equal(zoomConfirmLeave.phase, 'confirm_leave_meeting');

const meetJoin = chooseThreePlatformBrowserAutomationAction({
  platform: 'google-meet',
  autoJoin: true,
  pages: [page('meet-prejoin', 'https://meet.google.com/abc-defg-hij', [control({ text: 'Ask to join' })])],
});
assert.equal(meetJoin.phase, 'join_meeting');

const meetAnonymousName = chooseThreePlatformBrowserAutomationAction({
  platform: 'google-meet',
  autoJoin: true,
  displayName: 'Timeline Adapter Acceptance',
  pages: [page('meet-anonymous', 'https://meet.google.com/abc-defg-hij', [
    control({ tag: 'input', type: 'text', placeholder: 'Your name', value_present: false }),
    control({ text: 'Ask to join' }),
  ])],
});
assert.equal(meetAnonymousName.phase, 'fill_display_name');
assert.equal(meetAnonymousName.value, 'Timeline Adapter Acceptance');

const meetAnonymousJoin = chooseThreePlatformBrowserAutomationAction({
  platform: 'google-meet',
  autoJoin: true,
  completedActionKeys: ['phase:fill_display_name'],
  pages: [page('meet-anonymous', 'https://meet.google.com/abc-defg-hij', [
    control({ tag: 'input', type: 'text', placeholder: 'Your name', value_present: true }),
    control({ text: 'Ask to join' }),
  ])],
});
assert.equal(meetAnonymousJoin.phase, 'join_meeting');

const teamsContinue = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  pages: [page('teams-join', 'https://teams.microsoft.com/l/meetup-join/abc', [control({ text: 'Continue on this browser' })])],
});
assert.equal(teamsContinue.phase, 'continue_browser');

const teamsAnonymousName = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  displayName: 'Timeline Adapter Acceptance',
  pages: [page('teams-anonymous', 'https://teams.live.com/meet/1234567890123?p=sample', [
    control({ tag: 'input', type: 'text', placeholder: 'Type your name', value_present: false }),
    control({ text: 'Join now' }),
  ])],
});
assert.equal(teamsAnonymousName.phase, 'fill_display_name');

const teamsAnonymousJoin = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  completedActionKeys: ['phase:fill_display_name'],
  pages: [page('teams-anonymous', 'https://teams.live.com/meet/1234567890123?p=sample', [
    control({ tag: 'input', type: 'text', placeholder: 'Type your name', value_present: true }),
    control({ text: 'Join now' }),
  ])],
});
assert.equal(teamsAnonymousJoin.phase, 'join_meeting');

const teamsError = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  pages: [page('teams-error', 'https://teams.microsoft.com/v2/', [control({ text: 'Retry' }), control({ text: 'Clear cache and retry' })], 'Microsoft Teams')],
});
assert.equal(teamsError, null);

const teamsPersonalMeetingCard = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  pages: [page('teams-personal-home', 'https://teams.live.com/v2/', [
    control({ text: '创建会议链接' }),
    control({ text: '使用会议 ID 加入' }),
    control({ text: '加入' }),
  ], '开会 | Microsoft Teams')],
});
assert.equal(teamsPersonalMeetingCard.phase, 'open_meeting_card');
assert.deepEqual(teamsPersonalMeetingCard.selector, { text: '加入' });

const teamsPersonalPreJoin = chooseThreePlatformBrowserAutomationAction({
  platform: 'teams',
  autoJoin: true,
  pages: [page('teams-personal-prejoin', 'https://teams.live.com/v2/', [
    control({ text: '立即加入' }),
  ], 'Meeting join | Microsoft Teams')],
});
assert.equal(teamsPersonalPreJoin.phase, 'join_meeting');
assert.deepEqual(teamsPersonalPreJoin.selector, { text: '立即加入' });

const expression = buildThreePlatformBrowserActionExpression(meetJoin);
assert.match(expression, /Ask to join/);
assert.match(expression, /node\.click\(\)/);

const guarded = spawnSync(process.execPath, [
  fileURLToPath(new URL('../scripts/start-three-platform-live-acceptance.mjs', import.meta.url)),
  '--platform=zoom',
  '--auto-join=true',
  '--startup-only=true',
  '--skip-build=true',
], { encoding: 'utf8' });
assert.equal(guarded.status, 2);
assert.match(guarded.stderr, /--allow-external-actions=true/);

const sharedZoomMeetingGuard = spawnSync(process.execPath, [
  fileURLToPath(new URL('../scripts/start-three-platform-live-acceptance.mjs', import.meta.url)),
  '--platform=zoom',
  '--speaker-peer=true',
  '--auto-join=true',
  '--allow-external-actions=true',
  '--startup-only=true',
  '--skip-build=true',
], { encoding: 'utf8' });
assert.equal(sharedZoomMeetingGuard.status, 2);
assert.match(sharedZoomMeetingGuard.stderr, /shared Zoom meeting URL/);

console.log('ok three platform live browser automation');
