import assert from 'node:assert/strict';

import {
  detectMeetingFromUrl,
  detectMeetingPlatformFromUrl,
  stableMeetingIdFromUrl,
} from '../packages/meeting-timeline-sdk/adapters/meeting-url.mjs';
import { normalizeLocalDetectorEvent } from '../packages/meeting-timeline-sdk/adapters/local-detector.mjs';

const google = detectMeetingFromUrl('https://meet.google.com/abc-defg-hij?authuser=0');
assert.equal(google.platform, 'google_meet');
assert.equal(google.meeting_id, 'abc-defg-hij');
assert.equal(google.external_meeting_id, 'abc-defg-hij');
assert.equal(google.confidence, 'high');
assert.equal(detectMeetingPlatformFromUrl('https://meet.google.com/abc-defg-hij'), 'google_meet');

const teams = detectMeetingFromUrl('https://teams.microsoft.com/l/meetup-join/19%3Ameeting_MjQxODI%40thread.v2/0?context=%7B%7D');
assert.equal(teams.platform, 'microsoft_teams');
assert.equal(teams.meeting_id, '19:meeting_MjQxODI@thread.v2');
assert.equal(teams.confidence, 'high');

const zoom = detectMeetingFromUrl('https://us06web.zoom.us/j/987654321?pwd=secret');
assert.equal(zoom.platform, 'zoom');
assert.equal(zoom.meeting_id, '987654321');

const zoomNative = detectMeetingFromUrl('zoommtg://zoom.us/join?confno=987654321&pwd=secret');
assert.equal(zoomNative.platform, 'zoom');
assert.equal(zoomNative.meeting_id, '987654321');
assert.equal(zoomNative.confidence, 'high');

const zoomWebClientJoin = detectMeetingFromUrl('https://app.zoom.us/wc/join/987654321?pwd=secret');
assert.equal(zoomWebClientJoin.platform, 'zoom');
assert.equal(zoomWebClientJoin.meeting_id, '987654321');

const teamsNative = detectMeetingFromUrl('msteams://teams.microsoft.com/l/meetup-join/19%3Ameeting_NATIVE%40thread.v2/0');
assert.equal(teamsNative.platform, 'microsoft_teams');
assert.equal(teamsNative.meeting_id, '19:meeting_NATIVE@thread.v2');

const lark = detectMeetingFromUrl({
  window: {
    url: 'https://vc.feishu.cn/j/123456789',
    title: 'Feishu meeting window',
  },
});
assert.equal(lark.platform, 'lark');
assert.equal(lark.meeting_id, '123456789');
assert.equal(lark.title, 'Feishu meeting window');

const larkAppLink = detectMeetingFromUrl('https://applink.feishu.cn/client/videochat/open?meeting_id=lark-meeting-001');
assert.equal(larkAppLink.platform, 'lark');
assert.equal(larkAppLink.meeting_id, 'lark-meeting-001');

const larkNative = detectMeetingFromUrl('lark://client/byteview/join?meeting_no=lark-native-001');
assert.equal(larkNative.platform, 'lark');
assert.equal(larkNative.meeting_id, 'lark-native-001');

const webex = detectMeetingFromUrl('https://example.webex.com/meet/ada');
assert.equal(webex.platform, 'webex');
assert.equal(webex.meeting_id, 'meet-ada');

const webexNative = detectMeetingFromUrl('webex://meet/ada');
assert.equal(webexNative.platform, 'webex');
assert.equal(webexNative.meeting_id, 'meet-ada');

assert.equal(stableMeetingIdFromUrl('https://unknown.example.com/meeting/abc'), 'unknown.example.com-meeting-abc');
assert.equal(detectMeetingFromUrl('not a url'), null);

const localDetectorGoogle = normalizeLocalDetectorEvent({
  type: 'meeting_started',
  url: 'https://meet.google.com/abc-defg-hij',
  start_time_ms: 1_782_442_800_000,
});
assert.equal(localDetectorGoogle[0].meeting.platform, 'google_meet');
assert.equal(localDetectorGoogle[0].meeting.meeting_id, 'abc-defg-hij');
assert.equal(localDetectorGoogle[0].meeting.external_meeting_id, 'abc-defg-hij');

const localDetectorTeams = normalizeLocalDetectorEvent({
  type: 'active_speaker',
  window: {
    url: 'https://teams.microsoft.com/l/meetup-join/19%3Ameeting_XYZ%40thread.v2/0',
  },
  occurred_at_ms: 1_782_442_810_000,
  speaker: { name: 'Ada' },
});
assert.equal(localDetectorTeams[0].type, 'speaker_started');
assert.equal(localDetectorTeams[0].meeting.platform, 'microsoft_teams');
assert.equal(localDetectorTeams[0].meeting.meeting_id, '19:meeting_XYZ@thread.v2');
assert.equal(localDetectorTeams[0].speaker_name, 'Ada');

console.log('ok meeting URL detection helpers');
