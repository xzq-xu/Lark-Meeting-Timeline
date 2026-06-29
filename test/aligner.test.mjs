import assert from 'node:assert/strict';
import {
  buildTimeline,
  demoTimeline,
  extractLarkMeetingPatch,
  normalizeAnnotationEvent,
  normalizeLarkEventPayload,
  normalizeSequence,
} from '../src/normalize.mjs';

const timeline = demoTimeline();
assert.equal(timeline.meeting.platform, 'lark');
assert.equal(timeline.segments.length, 5);
assert.equal(timeline.sequence.length, 4);

const why = timeline.alignments.find((x) => x.sequence_id === 'mark-why-1');
assert.ok(why, 'why mark should exist');
assert.equal(why.active_segment_id, 'seg-2');
assert.ok(why.confidence >= 0.9);

const screenShare = timeline.alignments.find((x) => x.sequence_id === 'mark-risk-1');
assert.ok(screenShare.events.some((x) => x.event.type === 'screen_share_start'));

const absoluteSequence = normalizeSequence([
  {
    id: 'abs-1',
    ts: '2026-06-26T02:03:25.000Z',
    label: 'absolute timestamp',
  },
], timeline.meeting);
assert.equal(absoluteSequence[0].time_ms, 205_000);

const rebuilt = buildTimeline({
  meeting: timeline.meeting,
  segments: timeline.segments,
  events: timeline.events,
  sequence: absoluteSequence,
});
assert.equal(rebuilt.alignments[0].active_segment_id, 'seg-3');

const liveOnly = buildTimeline({
  meeting: {
    platform: 'lark',
    meeting_id: 'live-only',
    meeting_url: 'https://vc.feishu.cn/j/demo',
    title: 'annotation only',
    start_time: '2026-06-26T03:00:00.000Z',
    pending_binding: true,
    source: 'annotation_fallback',
  },
  segments: [],
  events: [{ id: 'start', time_ms: 0, type: 'meeting_start', label: '会议开始', source: 'lark_event' }],
  sequence: [{ id: 'live-mark', time_ms: 30_000, kind: 'mark', label: '实时标注' }],
});
assert.equal(liveOnly.segments.length, 0);
assert.equal(liveOnly.meeting.meeting_url, 'https://vc.feishu.cn/j/demo');
assert.equal(liveOnly.meeting.pending_binding, true);
assert.equal(liveOnly.meeting.source, 'annotation_fallback');
assert.equal(liveOnly.alignments.length, 1);
assert.equal(liveOnly.alignments[0].active_segment_id, null);
assert.equal(liveOnly.alignments[0].events[0].event.type, 'meeting_start');

const localSimulation = buildTimeline({
  meeting: {
    platform: 'lark',
    meeting_id: 'local-sim',
    title: 'local simulation',
    start_time: '2026-06-26T03:00:00.000Z',
    source: 'local_simulation',
  },
  events: [
    {
      id: 'evt-live-meeting-end',
      time_ms: 2000,
      type: 'meeting_end',
      label: '本地模拟结束',
      source: 'local_simulation',
      metadata: { raw_type: 'local.live_meeting.ended' },
    },
  ],
});
assert.equal(localSimulation.events[0].label, '本地模拟结束');
assert.equal(localSimulation.events[0].source, 'local_simulation');

const reservePending = buildTimeline({
  meeting: {
    platform: 'lark',
    meeting_id: 'reserve-001',
    title: 'reserve pending',
    start_time: '2026-06-26T03:00:00.000Z',
    source: 'lark_reserve_pending',
    pending_binding: true,
    reserve_id: 'reserve-001',
  },
  events: [],
  sequence: [],
});
assert.equal(reservePending.meeting.pending_binding, true);
assert.equal(reservePending.meeting.source, 'lark_reserve_pending');
assert.equal(reservePending.events.length, 0);

const openAnnotation = normalizeAnnotationEvent({
  id: 'epaper-mark-001',
  source: 'hanwang_epaper',
  time_ms: 45_000,
  kind: 'handwriting_trigger',
  text_candidates: ['why?', 'why'],
  intent: 'question',
  mark: { action: 'freehand' },
  strokes: [[{ x: 0.1, y: 0.2, t: 0 }]],
}, liveOnly.meeting, Date.parse('2026-06-26T03:00:45.000Z'));
assert.equal(openAnnotation.id, 'epaper-mark-001');
assert.equal(openAnnotation.source, 'hanwang_epaper');
assert.equal(openAnnotation.label, 'why?');
assert.deepEqual(openAnnotation.payload.text_candidates, ['why?', 'why']);
assert.equal(openAnnotation.payload.mark.action, 'freehand');
assert.equal(openAnnotation.payload.strokes.length, 1);

const capturedAnnotation = normalizeAnnotationEvent({
  id: 'epaper-mark-captured',
  source: 'hanwang_epaper',
  captured_at_ms: Date.parse('2026-06-26T03:00:30.000Z'),
  timestamp_ms: Date.parse('2026-06-26T03:09:00.000Z'),
  label: '采集时间优先',
}, liveOnly.meeting, Date.parse('2026-06-26T03:10:00.000Z'));
assert.equal(capturedAnnotation.time_ms, 30_000);
assert.equal(capturedAnnotation.payload.timing.captured_at, Date.parse('2026-06-26T03:00:30.000Z'));
assert.equal(capturedAnnotation.time_source, 'captured_at');
assert.equal(capturedAnnotation.payload.timing.source_field, 'captured_at_ms');
assert.equal(capturedAnnotation.payload.timing.server_receive_delay_ms, 570_000);

const relativeOverrideAnnotation = normalizeAnnotationEvent({
  id: 'epaper-mark-relative',
  source: 'hanwang_epaper',
  time_ms: 12_345,
  captured_at_ms: Date.parse('2026-06-26T03:00:30.000Z'),
  label: '显式相对时间优先',
}, liveOnly.meeting, Date.parse('2026-06-26T03:10:00.000Z'));
assert.equal(relativeOverrideAnnotation.time_ms, 12_345);
assert.equal(relativeOverrideAnnotation.time_source, 'explicit_time');
assert.equal(relativeOverrideAnnotation.payload.timing.source_field, 'time_ms');

const strokeTimeAnnotation = normalizeAnnotationEvent({
  id: 'epaper-mark-stroke-time',
  source: 'hanwang_epaper',
  strokes: [[
    { x: 0.1, y: 0.2, t: Date.parse('2026-06-26T03:00:10.000Z') },
    { x: 0.3, y: 0.4, t: Date.parse('2026-06-26T03:00:15.000Z') },
  ]],
  label: 'stroke 绝对时间',
}, liveOnly.meeting, Date.parse('2026-06-26T03:10:00.000Z'));
assert.equal(strokeTimeAnnotation.time_ms, 15_000);
assert.equal(strokeTimeAnnotation.time_source, 'stroke_point_time');
assert.equal(strokeTimeAnnotation.payload.timing.captured_at_ms, Date.parse('2026-06-26T03:00:15.000Z'));

const cachedBeforeEndAnnotation = normalizeAnnotationEvent({
  id: 'cached-before-end',
  source: 'hanwang_epaper',
  payload: {
    timing: {
      captured_at_ms: Date.parse('2026-06-26T03:00:18.000Z'),
    },
  },
  label: '会中采集会后发送',
}, {
  ...liveOnly.meeting,
  end_time: '2026-06-26T03:00:20.000Z',
}, Date.parse('2026-06-26T03:01:00.000Z'));
assert.equal(cachedBeforeEndAnnotation.time_ms, 18_000);
assert.equal(cachedBeforeEndAnnotation.payload.timing.source_field, 'payload.timing.captured_at_ms');
assert.equal(cachedBeforeEndAnnotation.payload.timing.after_meeting_end_ms, 0);

const endedBeforeLateMark = buildTimeline({
  meeting: {
    platform: 'lark',
    meeting_id: 'ended-before-late-mark',
    title: 'event order',
    start_time: '2026-06-26T03:00:00.000Z',
    end_time: '2026-06-26T03:00:20.000Z',
  },
  events: [{ id: 'end', time_ms: 20_000, type: 'meeting_end', label: '会议结束' }],
  sequence: [{ id: 'late-mark', time_ms: 45_000, kind: 'mark', label: '迟到标注' }],
});
assert.equal(endedBeforeLateMark.events[0].time_ms, 20_000);
assert.equal(endedBeforeLateMark.sequence[0].time_ms, 45_000);

const larkMeetingStartPayload = {
  header: {
    event_id: 'evt-real-meeting-start-001',
    event_type: 'vc.meeting.meeting_started_v1',
    create_time: '1782442800',
  },
  event: {
    meeting_id: 'om_real_001',
    topic: '真实飞书会议',
    meeting_url: 'https://vc.feishu.cn/j/real-demo',
    start_time: '1782442800',
  },
};
const larkMeetingPatch = extractLarkMeetingPatch(larkMeetingStartPayload, {});
assert.equal(larkMeetingPatch.meeting_id, 'om_real_001');
assert.equal(larkMeetingPatch.title, '真实飞书会议');
assert.equal(larkMeetingPatch.meeting_url, 'https://vc.feishu.cn/j/real-demo');
assert.equal(larkMeetingPatch.start_time, '2026-06-26T03:00:00.000Z');
assert.equal(larkMeetingPatch.end_time, null);
const larkStartEvent = normalizeLarkEventPayload(larkMeetingStartPayload, larkMeetingPatch);
assert.equal(larkStartEvent.type, 'meeting_start');
assert.equal(larkStartEvent.time_ms, 0);

const directMeetingStartPayload = {
  header: {
    event_id: 'evt-direct-meeting-start-001',
    event_type: 'vc.meeting.all_meeting_started_v1',
    create_time: '1782442800000',
  },
  event: {
    meeting: {
      id: 'om_direct_001',
      topic: '用户直接开启飞书会议',
      url: 'https://vc.feishu.cn/j/direct-demo',
      start_time: '1782442800',
    },
  },
};
const directMeetingPatch = extractLarkMeetingPatch(directMeetingStartPayload, {});
assert.equal(directMeetingPatch.meeting_id, 'om_direct_001');
assert.equal(directMeetingPatch.title, '用户直接开启飞书会议');
assert.equal(directMeetingPatch.meeting_url, 'https://vc.feishu.cn/j/direct-demo');
assert.equal(directMeetingPatch.start_time, '2026-06-26T03:00:00.000Z');
const directStartEvent = normalizeLarkEventPayload(directMeetingStartPayload, directMeetingPatch);
assert.equal(directStartEvent.type, 'meeting_start');
assert.equal(directStartEvent.time_ms, 0);

const meetingInfoStartPayload = {
  header: {
    event_id: 'evt-meeting-info-start-001',
    event_type: 'vc.meeting.all_meeting_started_v1',
    create_time: '1782442800',
  },
  event: {
    meeting_info: {
      open_meeting_id: 'om_info_001',
      meeting_no: '987654321',
      topic: 'meeting_info 形态会议',
      url: 'https://vc.feishu.cn/j/meeting-info-demo',
      begin_time: '1782442800',
    },
  },
};
const meetingInfoPatch = extractLarkMeetingPatch(meetingInfoStartPayload, {});
assert.equal(meetingInfoPatch.meeting_id, 'om_info_001');
assert.equal(meetingInfoPatch.external_meeting_id, 'om_info_001');
assert.equal(meetingInfoPatch.title, 'meeting_info 形态会议');
assert.equal(meetingInfoPatch.meeting_url, 'https://vc.feishu.cn/j/meeting-info-demo');
assert.equal(meetingInfoPatch.start_time, '2026-06-26T03:00:00.000Z');
const meetingInfoStartEvent = normalizeLarkEventPayload(meetingInfoStartPayload, meetingInfoPatch);
assert.equal(meetingInfoStartEvent.type, 'meeting_start');
assert.equal(meetingInfoStartEvent.metadata.meeting_id, 'om_info_001');

const directMeetingEndPayload = {
  header: {
    event_id: 'evt-direct-meeting-end-001',
    event_type: 'vc.meeting.all_meeting_ended_v1',
  },
  event: {
    meeting: {
      id: 'om_direct_001',
      topic: '用户直接开启飞书会议',
      url: 'https://vc.feishu.cn/j/direct-demo',
      start_time: '1782442800',
      end_time: '1782443400',
    },
  },
};
const directEndPatch = extractLarkMeetingPatch(directMeetingEndPayload, directMeetingPatch);
assert.equal(directEndPatch.end_time, '2026-06-26T03:10:00.000Z');
const directEndEvent = normalizeLarkEventPayload(directMeetingEndPayload, directEndPatch);
assert.equal(directEndEvent.type, 'meeting_end');
assert.equal(directEndEvent.time_ms, 600_000);

const joinMeetingPayload = {
  header: {
    event_id: 'evt-join-meeting-001',
    event_type: 'vc.meeting.join_meeting_v1',
    create_time: '1782442860',
  },
  event: {
    meeting: {
      id: 'om_join_001',
      topic: '用户直接开启后加入会议',
      meeting_no: '123456789',
      start_time: '1782442800',
    },
  },
};
const joinMeetingPatch = extractLarkMeetingPatch(joinMeetingPayload, {});
assert.equal(joinMeetingPatch.meeting_id, 'om_join_001');
assert.equal(joinMeetingPatch.start_time, '2026-06-26T03:00:00.000Z');
const joinMeetingEvent = normalizeLarkEventPayload(joinMeetingPayload, joinMeetingPatch);
assert.equal(joinMeetingEvent.type, 'participant_join');
assert.equal(joinMeetingEvent.time_ms, 60_000);

console.log('ok lark timeline alignment');
