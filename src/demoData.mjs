export const demoMeeting = {
  platform: 'lark',
  meeting_id: 'demo-lark-meeting-001',
  minute_token: 'demo-minute-token',
  title: '电子纸标注与会议时间轴对齐评审',
  start_time: '2026-06-26T02:00:00.000Z',
  end_time: '2026-06-26T02:42:00.000Z',
  timezone: 'Asia/Shanghai',
};

export const demoTranscript = [
  {
    id: 'seg-1',
    start_ms: 18_000,
    end_ms: 72_000,
    speaker_id: 'u_zhang',
    speaker_name: '张三',
    text: '今天我们先验证飞书会议时间轴，重点是把妙记、会议事件和电子纸笔迹放到同一个时间线上。',
  },
  {
    id: 'seg-2',
    start_ms: 95_000,
    end_ms: 145_000,
    speaker_id: 'u_li',
    speaker_name: '李四',
    text: '如果用户在会议里写了 why 或者画了重点，我们需要知道当时正在讨论什么内容。',
  },
  {
    id: 'seg-3',
    start_ms: 184_000,
    end_ms: 246_000,
    speaker_id: 'u_xu',
    speaker_name: '徐智强',
    text: '先不要把标注格式设计得太重，第一版只要能把数据序列贴回会议上下文即可。',
  },
  {
    id: 'seg-4',
    start_ms: 318_000,
    end_ms: 382_000,
    speaker_id: 'u_chen',
    speaker_name: '陈五',
    text: '飞书这边有会议事件和妙记产物，但导出妙记权限需要确认，所以要保留手动导入和音频转写的 fallback。',
  },
  {
    id: 'seg-5',
    start_ms: 516_000,
    end_ms: 588_000,
    speaker_id: 'u_li',
    speaker_name: '李四',
    text: '验证通过以后再接 OAuth、租户授权和持久化。现在先跑通真实 minute token 到内部 timeline 的链路。',
  },
];

export const demoEvents = [
  {
    id: 'evt-start',
    time_ms: 0,
    type: 'meeting_start',
    label: '会议开始',
    source: 'lark_event',
    metadata: { raw_type: 'vc.meeting.started' },
  },
  {
    id: 'evt-recording',
    time_ms: 42_000,
    type: 'recording_start',
    label: '录制开始',
    source: 'lark_event',
    metadata: { raw_type: 'vc.recording.started' },
  },
  {
    id: 'evt-share-start',
    time_ms: 302_000,
    type: 'screen_share_start',
    label: '开始共享屏幕',
    source: 'lark_event',
    metadata: { raw_type: 'vc.screen_share.started' },
  },
  {
    id: 'evt-share-end',
    time_ms: 452_000,
    type: 'screen_share_end',
    label: '结束共享屏幕',
    source: 'lark_event',
    metadata: { raw_type: 'vc.screen_share.ended' },
  },
];

export const demoSequence = [
  {
    id: 'mark-why-1',
    time_ms: 118_000,
    kind: 'handwriting_trigger',
    label: '手写 why?',
    payload: {
      text_candidates: ['why?', 'why', 'w hy?'],
      intent: 'question',
    },
  },
  {
    id: 'mark-circle-1',
    time_ms: 206_000,
    kind: 'mark',
    label: '圈出“格式不要太重”',
    payload: {
      action: 'enclosure',
      target_text: '标注格式设计得太重',
    },
  },
  {
    id: 'mark-risk-1',
    time_ms: 342_000,
    kind: 'attention',
    label: '标注：权限风险',
    payload: {
      text: '权限?',
      intent: 'risk_check',
    },
  },
  {
    id: 'mark-next-1',
    time_ms: 552_000,
    kind: 'todo',
    label: '写下 next: OAuth',
    payload: {
      text: 'next: OAuth',
      intent: 'follow_up',
    },
  },
];
