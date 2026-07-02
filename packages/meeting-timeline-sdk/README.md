# Meeting Timeline SDK

把外部项目里的会议检测、电子纸手写标记、桌面观察器事件，插入到统一会议时间轴。

这个 SDK 只依赖当前服务暴露的稳定协议，不依赖飞书 SDK：

- `POST /api/meeting-session/start`
- `POST /api/annotations`
- `POST /api/annotations/batch`
- `POST /api/meeting-session/end`
- `GET /api/annotations/status?id=...`
- `GET /api/stream`

## 安装

本仓库内直接用相对路径：

```js
import { createMeetingTimelineClient } from './packages/meeting-timeline-sdk/index.mjs';
```

给别的项目使用时，可以先把 `packages/meeting-timeline-sdk` 作为独立包复制过去，或用本地 file 依赖：

```json
{
  "dependencies": {
    "@ai-annotation/meeting-timeline-sdk": "file:../meeting-timeline-sdk"
  }
}
```

## 基本用法

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';

const timeline = createMeetingTimelineClient({
  baseUrl: 'http://localhost:8787',
  source: 'hanwang_epaper',
  deviceId: 'hanwang-device-001',
});

await timeline.startMeeting({
  platform: 'lark',
  meetingId: 'meeting-001',
  title: '产品评审',
  meetingUrl: 'https://vc.feishu.cn/j/example',
  startTimeMs: Date.now(),
  detectorSource: 'desktop_meeting_observer',
});

await timeline.insertMark({
  id: 'mark-001',
  capturedAtMs: Date.now(),
  kind: 'handwriting_trigger',
  label: 'why?',
  textCandidates: ['why?', 'why'],
  intent: 'question',
  strokes: [],
});

await timeline.endMeeting({
  endTimeMs: Date.now(),
});
```

## 平台 adapter

SDK 还提供轻量 adapter，把不同会议平台事件归一化为统一 signal，再应用到同一条时间轴协议：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { applyMeetingSignals } from '@ai-annotation/meeting-timeline-sdk/adapters/core';
import { normalizeGoogleMeetEvent } from '@ai-annotation/meeting-timeline-sdk/adapters/google-meet';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });

const signals = normalizeGoogleMeetEvent(req.body);
await applyMeetingSignals(timeline, signals);
```

当前内置归一化器：

- `@ai-annotation/meeting-timeline-sdk/adapters/local-detector`
- `@ai-annotation/meeting-timeline-sdk/adapters/lark`
- `@ai-annotation/meeting-timeline-sdk/adapters/google-meet`
- `@ai-annotation/meeting-timeline-sdk/adapters/microsoft-teams`
- `@ai-annotation/meeting-timeline-sdk/adapters/zoom`
- `@ai-annotation/meeting-timeline-sdk/adapters/webex`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-url`
- `@ai-annotation/meeting-timeline-sdk/adapters/local-observer`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-registry`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest`
- `@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-kit`
- `@ai-annotation/meeting-timeline-sdk/adapters/signal-reconciler`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-http`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-fixtures`
- `@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan`
- `@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-onboarding`
- `@ai-annotation/meeting-timeline-sdk/adapters/transcript`
- `@ai-annotation/meeting-timeline-sdk/adapters/webhook-security`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-setup`

`meeting_started` 会调用 `startMeeting`，`meeting_ended` 会调用 `endMeeting`。`participant_joined/left`、`speaker_started/ended` 和 `artifact_ready` 默认不会写入用户标注流；如果需要临时显示参会人位置，可以给 `applyMeetingSignals` 传 `{ participantAsAnnotation: true }`，如果需要显示发言人位置，可以传 `{ speakerAsAnnotation: true }`，或者用 `onParticipantSignal` / `onSpeakerSignal` / `onArtifactSignal` 接到自己的服务端轨道。`subscription_lifecycle` 表示平台订阅自身的过期、移除、暂停、漏投或重新授权要求，默认不画到会议轴；需要接入诊断时传 `onSubscriptionLifecycleSignal` 处理。

如果外部项目希望先用一个入口跑通，可以用 `timeline-bridge`。它组合了本地观察、平台事件 ingest、信号校准、实时标注和会后转写导入：

```js
import { createMeetingTimelineBridge } from '@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge';

const bridge = createMeetingTimelineBridge({
  baseUrl: 'http://localhost:8787',
  source: 'hanwang_epaper',
  deviceId: 'hanwang-001',
});

await bridge.observeCandidates(browserWindowsSnapshot, { observedAtMs: Date.now() });
await bridge.ingest('google-meet', req.body);
await bridge.insertMark({ capturedAtMs: Date.now(), label: 'why?' });
await bridge.importTranscript({
  platform: 'google_meet',
  meeting: { meetingId: 'conference-record-id' },
  raw: googleTranscriptEntries,
});
```

如果业务项目要接入多个会议平台，推荐从 `platform-kit` 开始。它把 `timeline-bridge`、webhook router、平台 setup/onboarding、fixture acceptance 组合成一个入口；底层 normalizer、验签、artifact fetch 仍然可以按需单独 import：

```js
import { createMeetingPlatformTimelineKit } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-kit';

const meetingKit = createMeetingPlatformTimelineKit({
  baseUrl: 'https://timeline.example.com',
  basePath: '/api/platform-events',
  env: process.env,
  source: 'desktop_meeting_observer',
  verify: true,
  reconcile: true,
});

const setup = meetingKit.platform('google-meet');
// setup.webhook_route === '/api/platform-events/google-meet'
// setup.integration_plan.recommended_mode === 'hybrid_local_observer_first'

const fixtureReport = meetingKit.fixtureAcceptance('google-meet', {
  requiredCoverage: ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'],
});

await meetingKit.handleWebhook({
  method: req.method,
  url: req.url,
  headers: req.headers,
  body: req.body,
  rawBody: req.rawBody,
});

await meetingKit.insertMark({ capturedAtMs: Date.now(), label: 'why?' });
```

在标准 Web `Request` / `Response` 环境里，可以直接让 kit 读取 raw body 并返回 `Response`。这样 Zoom / Webex 签名校验仍然能拿到原始 body：

```js
export async function POST(request) {
  return meetingKit.handleFetchRequest(request);
}

export async function GET(request) {
  return meetingKit.handleFetchRequest(request); // Microsoft Graph validationToken / router status
}
```

如果不使用 `platform-kit`，也可以单独创建 HTTP handler：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { createMeetingPlatformFetchHandler } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-http';

const timeline = createMeetingTimelineClient({ baseUrl: 'https://timeline.example.com' });
const handlePlatformWebhook = createMeetingPlatformFetchHandler(timeline, {
  basePath: '/api/platform-events',
  verify: true,
  reconcile: true,
});

export default handlePlatformWebhook;
```

Local detector adapter 支持桌面观察器、浏览器扩展、汉王宿主 App 或人工控制器上报 `meeting_started` / `meeting_ended`，payload 可以带 `detected_platform: 'google_meet'` 或 `meeting.platform: 'zoom'` 表示真实会议来源；如果只传窗口 URL，SDK 会用 `meeting-url` 自动识别 Google Meet / Teams / Zoom / Lark / Webex 的平台和会议 ID。Local detector 也支持 `active_speaker`、`speaker_started`、`speaker_ended` 这类发言人信号，用来低延迟标出发言人位置。Lark adapter 支持 `vc.meeting.all_meeting_started_v1`、`vc.meeting.all_meeting_ended_v1`、`vc.meeting.meeting_started_v1`、`vc.meeting.meeting_ended_v1`、`vc.meeting.join_meeting_v1`、`vc.meeting.leave_meeting_v1`，并会把 `minute_token` 透传到会议轴，便于会后妙记导入。Google Meet adapter 同时支持已经解包的 Workspace Events CloudEvent，以及 Pub/Sub 默认 wrapped push body。wrapped body 会自动 base64 解码 `message.data`，所以 webhook handler 可以直接把 `req.body` 传给 `normalizeGoogleMeetEvent(req.body)`。Google Workspace Events 的 `subscription.v1.suspended`、`subscription.v1.expirationReminder`、`subscription.v1.expired` 会归一化为 `subscription_lifecycle`。Microsoft Graph change notifications 的 `lifecycleEvent` 值 `reauthorizationRequired`、`subscriptionRemoved`、`missed` 也会归一化为 `subscription_lifecycle`。Webex adapter 支持 `meetings` started/ended、`meetingParticipants` joined/left、`recordings` created/updated、`meetingTranscripts` created。

宿主服务如果要按平台名动态接 webhook，可以直接用 registry：

```js
import { meetingPlatformEventAdapterFor } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry';

const adapter = meetingPlatformEventAdapterFor(req.params.platform);
const signals = adapter.normalize(req.body, { receivedAtMs: Date.now() });
```

如果外部项目只想“收到平台 webhook 后直接落到会议轴”，可以用更高层的 `platform-ingest`：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import {
  diagnosePlatformEvent,
  ingestPlatformEvent,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });

const diagnostic = diagnosePlatformEvent('google-meet', req.body);
if (!diagnostic.actionable) {
  console.warn('Platform event did not produce timeline signals', diagnostic.issues);
}

await ingestPlatformEvent(timeline, 'google-meet', req.body, {
  receivedAtMs: Date.now(),
  participantAsAnnotation: true,
});
```

`diagnosePlatformEvent()` 不写入 timeline，只返回 adapter、`signal_types`、`coverage`、`meetings` 和 `issues`，用于验收真实样本事件是否能建轴、是否只是订阅生命周期事件、是否缺少 meeting URL 导致和本地观察器对齐变弱。`ingestPlatformEvent()` 内部会按平台名选择 normalizer，把原始事件转成 `NormalizedMeetingSignal[]`，再调用 `startMeeting`、`endMeeting` 或可选的 participant/artifact handler。对于 Google Meet / Teams / Zoom / Webex，新项目可以优先接这一层，只有需要自定义事件验签、补拉详情或 artifact 导入时再下钻到 registry/core。

如果要验收一组真实平台样本，而不是单条事件，可以用 `platform-acceptance`。它会合并 setup readiness、integration plan 和样本事件诊断，输出 `accepted / blocked / pending_samples / missing_required_coverage`：

```js
import { buildPlatformAcceptanceReport } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance';

const report = buildPlatformAcceptanceReport('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
    GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub-pusher@demo.iam.gserviceaccount.com',
  },
  requireEndEvent: true,
  samples: {
    google_meet: [
      { label: 'started', body: googleStartedWebhook },
      { label: 'ended', body: googleEndedWebhook },
    ],
  },
});

console.log(report.status, report.missing_required_coverage, report.issues);
```

如果真实 webhook 还没打通，但要先验证宿主项目的 SDK 接入、CI gate 和时间轴写入路径，可以用 `platform-fixtures` 生成各平台的原始事件样本。fixture 仍然走对应 normalizer 和 acceptance report，不会绕过适配层：

```js
import { buildPlatformAcceptanceReport } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance';
import { buildPlatformFixtureAcceptanceInput } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-fixtures';

const fixtureInput = buildPlatformFixtureAcceptanceInput({
  baseUrl: 'https://timeline.example.com',
});

const googleFixtureReport = buildPlatformAcceptanceReport('google-meet', {
  ...fixtureInput,
  requireEndEvent: true,
  requiredCoverage: ['meeting_start', 'meeting_end', 'participant_track', 'artifact_ready'],
});

// googleFixtureReport.accepted === true
// fixtureInput.samples contains local_detector, lark, google_meet, microsoft_teams, zoom, webex
```

如果要给配置页、接入向导或 CI 验收生成一份“这个会议平台现在能不能接进 timeline”的总报告，可以用 `platform-onboarding`。它会合并 `permission plan`、`integration plan`、真实样本 `acceptance` 和 `artifact-plan`：

```js
import { buildMeetingPlatformOnboardingReport } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-onboarding';

const onboarding = buildMeetingPlatformOnboardingReport('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  features: ['axis', 'participants', 'transcript', 'recording', 'security'],
  samples: {
    google_meet: [
      { label: 'started', body: googleStartedWebhook },
      { label: 'transcript ready', body: googleTranscriptWebhook },
    ],
  },
});

// onboarding.status === 'ready' | 'blocked_by_setup' | 'needs_real_samples' | 'needs_more_coverage'
// onboarding.runtime_contract.annotation_time_field === 'captured_at_ms'
// onboarding.next_actions includes missing env/config steps and post-meeting artifact fetch actions
```

如果宿主同时接本地观察器和官方 webhook，建议用有状态的 `createReconciledPlatformEventIngestor()`。它会过滤 exact duplicate、重复 speaker 信号，并处理“本地观察先建轴，Google Meet / Teams / Zoom / Webex 官方事件后到用于校准”的优先级。匹配时会同时看 `meeting_url`、`meeting_id`、`external_meeting_id`，所以 Zoom 本地数字会议号和官方 `uuid`、Teams Graph resource 里的 `joinWebUrl` 都可以对齐到同一场会议：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { createReconciledPlatformEventIngestor } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const ingestor = createReconciledPlatformEventIngestor(timeline);

await ingestor.ingest('local-detector', {
  type: 'meeting_started',
  detected_platform: 'google_meet',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  start_time_ms: Date.now(),
});

await ingestor.ingest('google-meet', req.body); // 后到的官方 start 可校准本地轴
```

如果外部项目想直接复用完整 HTTP webhook 入口，可以用 `platform-webhook-handler`。它会处理 Microsoft Graph `validationToken`、Zoom `endpoint.url_validation`、Zoom/Webex/Teams/Google Pub/Sub 验证，再把事件交给 `ingestPlatformEvent()`。真实产品如果同时接本地观察器和官方 webhook，建议直接打开 `reconcile: true`，handler 会复用同一个有状态 reconciler，把本地先到的轴和官方后到的事件对齐，并在响应里返回 `reconciliation` 供调试：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { createPlatformWebhookHandler } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const handleWebhook = createPlatformWebhookHandler(timeline, {
  reconcile: true,
  googleMeet: {
    expectedAudience: 'https://timeline.example.com/api/platform-events/google-meet',
    serviceAccountEmail: 'pubsub-pusher@demo.iam.gserviceaccount.com',
  },
  microsoftTeams: {
    clientState: process.env.MICROSOFT_GRAPH_CLIENT_STATE,
  },
  zoom: {
    secretToken: process.env.ZOOM_WEBHOOK_SECRET_TOKEN,
  },
  webex: {
    secret: process.env.WEBEX_WEBHOOK_SECRET,
  },
});

// Express/Fastify/Koa/Node 原生都可以套这一层；关键是传入原始 rawBody。
const result = await handleWebhook({
  platform: 'google-meet',
  method: req.method,
  url: req.url,
  headers: req.headers,
  body: req.body,
  rawBody: req.rawBody,
});

res.status(result.status).set(result.headers).send(result.body);
```

验收真实平台样本时可以走同一个 handler 的 dry-run 模式：传 `dryRun: true`、`diagnoseOnly: true`，或在 URL query 里加 `?diagnose=1` / `?dry_run=1`。handler 仍会先做平台验证，但只返回 `diagnostic`，不会调用 `startMeeting` / `endMeeting` / `insertMark`，也不会污染共享 reconciler 状态：

```js
const diagnosticResult = await handleWebhook({
  platform: 'google-meet',
  method: req.method,
  url: `${req.url}?diagnose=1`,
  headers: req.headers,
  body: req.body,
  rawBody: req.rawBody,
});

console.log(diagnosticResult.body.diagnostic.coverage);
console.log(diagnosticResult.body.diagnostic.issues);
```

如果多个 HTTP endpoint 分别接 Google Meet、Teams、Zoom、Webex，但最终写同一个 timeline，可以把同一个 handler 或 `createReconciledPlatformEventIngestor()` 传给这些 endpoint，避免每个 endpoint 维护一份去重状态。`handleWebhook.getReconciliationState()` 可以用于诊断当前活跃会议、已处理 fingerprint 和最近发言人信号；测试或切换账号时可调用 `handleWebhook.resetReconciliationState()` 清空状态。

如果外部项目希望 SDK 直接管理一组平台 webhook 路由，可以用 `platform-webhook-router`。它不绑定 Express/Fastify/Koa，只负责把请求路径匹配到平台，并复用同一个 handler/reconciler：

```js
import { createMeetingPlatformWebhookRouter } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router';

const router = createMeetingPlatformWebhookRouter(timeline, {
  basePath: '/api/platform-events',
  baseUrl: 'https://timeline.example.com',
  verify: true,
  reconcile: true,
  googleMeet: {
    expectedAudience: 'https://timeline.example.com/api/platform-events/google-meet',
  },
  zoom: {
    secretToken: process.env.ZOOM_WEBHOOK_SECRET_TOKEN,
  },
});

// GET /api/platform-events/status
// GET /api/platform-events/setup
// POST /api/platform-events/google-meet
// GET/POST /api/platform-events/teams
const response = await router({
  method: req.method,
  url: req.url,
  headers: req.headers,
  body: req.body,
  rawBody: req.rawBody,
});
```

本地观察器或汉王宿主 App 也可以走同一个 handler，只是默认验证结果会是 `platform_verification_not_configured`：

```js
await handleWebhook({
  platform: 'local-detector',
  body: {
    type: 'meeting_started',
    detected_platform: 'google_meet',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
    start_time_ms: Date.now(),
  },
});
```

浏览器扩展、桌面观察器或电子纸宿主 App 不需要等官方 webhook。推荐直接用 `local-observer` 的高层入口：宿主只要定期喂窗口 URL / 标题 / 可见状态快照，SDK 会判断会议开始/结束并写入 timeline：

```js
import { createLocalMeetingTimelineObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/local-observer';

const observer = createLocalMeetingTimelineObserver(timeline, { source: 'desktop_observer' });
await observer.observe({
  url: 'https://meet.google.com/abc-defg-hij',
  observedAtMs: Date.now(),
});
```

如果宿主拿到的是浏览器扩展/桌面观察器的多窗口多标签快照，可以直接用 `observeCandidates()`。SDK 会展开 `windows[].tabs[]`、筛出 Google Meet / Teams / Zoom / Webex / Lark 候选，并选择最可信的会议标签；浏览器 tab 的 `active: false` 不会被误判成会议结束：

```js
await observer.observeCandidates({
  windows: [{
    focused: true,
    tabs: [
      { active: true, url: 'https://mail.example.com', title: 'Mail' },
      { active: false, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' },
    ],
  }],
}, {
  observedAtMs: Date.now(),
});
```

如果接入方只想拿 signal 自己处理，也可以用低层状态机：

```js
import { createLocalMeetingObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/local-observer';

const observer = createLocalMeetingObserver({ source: 'desktop_observer' });
const observed = observer.observe({
  url: 'https://meet.google.com/abc-defg-hij',
  observedAtMs: Date.now(),
});

console.log(observed.signals);
```

低层状态机同样支持 `observeCandidates()`，返回值会带 `selection.candidates` 和 `selection.selectedSnapshot`，方便调试为什么选中了某个会议窗口。

如果宿主已经知道事件类型，也可以直接走 `local-detector`：

```js
await ingestPlatformEvent(timeline, 'local-detector', {
  type: 'meeting_started',
  url: 'https://meet.google.com/abc-defg-hij',
  start_time_ms: Date.now(),
});
```

实时发言人轨也建议先走 `local-detector`。Google Meet、Teams、Zoom、Webex 的官方 webhook 主要提供会议开始/结束、参会人、录制和转写产物，不应假设它们能低延迟提供 active speaker；会后可以再用 transcript segment 的 `speaker_name` 回填。

```js
await ingestPlatformEvent(timeline, 'local-detector', {
  type: 'active_speaker',
  detected_platform: 'google_meet',
  meeting_url: 'https://meet.google.com/abc-defg-hij',
  occurred_at_ms: Date.now(),
  speaker: {
    id: 'speaker-ada',
    name: 'Ada',
  },
}, {
  speakerAsAnnotation: true,
});
```

## 会后转写导入

事件 adapter 只负责告诉你 transcript/recording 已生成；正文内容建议会后拉取后再导入。`artifact-plan` 可以把 `artifact_ready` signal 转成平台相关的补拉/导入计划，告诉宿主应该用哪个 provider API、哪个 transcript normalizer，以及是否只是录制 metadata：

```js
import { buildArtifactImportPlans } from '@ai-annotation/meeting-timeline-sdk/adapters/artifact-plan';

const plans = buildArtifactImportPlans(platformIngestResult);
// plans[0].fetch.strategy === 'google_meet_rest_transcript_entries'
// plans[0].transcript_import.normalizer === 'normalizeGoogleMeetTranscriptEntries'
```

如果宿主已经有对应平台的 OAuth access token，可以继续用 `artifact-fetch` 生成请求或直接拉取并导入。它不会绑定 Google/Microsoft/Zoom/Webex 的官方 SDK，默认只用注入的 `fetch`：

```js
import { fetchAndImportArtifactTranscript } from '@ai-annotation/meeting-timeline-sdk/adapters/artifact-fetch';

const result = await fetchAndImportArtifactTranscript(timeline, plans[0], {
  accessToken: providerAccessToken,
  fetchImpl: fetch,
});

// Google Meet 会请求 meet.googleapis.com/v2/.../transcripts/.../entries
// Teams 会请求 graph.microsoft.com/v1.0/.../transcripts/.../content
// Zoom/Webex 会优先使用 artifact_url / download_url 直链
```

SDK 提供平台中性的 `importTranscript()`，也提供更高层的 `importPlatformTranscript()`，会按 `platform` 自动选择 normalizer 并导入：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { importPlatformTranscript } from '@ai-annotation/meeting-timeline-sdk/adapters/transcript';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });

await importPlatformTranscript(timeline, {
  platform: 'google_meet',
  meeting: {
    platform: 'google_meet',
    meetingId: 'conference-record-id',
    startTimeMs: meetingStartMs,
  },
  raw: googleMeetTranscriptEntryList,
});
```

`adapters/transcript` 当前支持：

- `normalizeGoogleMeetTranscriptEntries(raw)`
- `normalizeMicrosoftTeamsTranscript(raw)`，支持 JSON segments 或 WebVTT/timed text
- `normalizeZoomTranscript(raw)`，支持 Zoom VTT 或 JSON segments
- `normalizeWebexTranscript(raw)`，支持 Webex VTT/text 或 JSON snippets/segments
- `buildPlatformTranscriptImportPayload(input)`，按 `platform` 自动选择 normalizer 并生成导入 body
- `importPlatformTranscript(client, input)`，按 `platform` 自动选择 normalizer，并调用 `client.importTranscript()` 导入

## 平台接入配置

`platform-setup` 提供只读 manifest 和订阅 request body builder，方便宿主项目生成配置页或自动化脚本：

```js
import {
  MEETING_PLATFORM_KEYS,
  allPlatformCapabilityContracts,
  allPlatformPermissionPlans,
  buildPlatformIntegrationPlan,
  buildPlatformPermissionPlan,
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildMicrosoftGraphSubscriptionRenewalRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildWebexWebhookRequests,
  buildZoomEventSubscriptionRequest,
  evaluatePlatformSubscriptionMaintenance,
  evaluatePlatformSetupReadiness,
  normalizeMeetingPlatform,
  platformCapabilityContract,
  platformSetupManifest,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup';

const platform = normalizeMeetingPlatform('google-meet'); // google_meet
const supportedPlatforms = MEETING_PLATFORM_KEYS; // local_detector, lark, google_meet, microsoft_teams, zoom, webex

const google = buildGoogleMeetWorkspaceSubscriptionRequest({
  targetResource: '//cloudidentity.googleapis.com/users/me',
  pubsubTopic: 'projects/demo/topics/meet-events',
});

const teams = buildMicrosoftTeamsMeetingCallSubscriptionRequest({
  joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/...',
  notificationUrl: 'https://timeline.example.com/api/platform-events/teams',
  clientState: process.env.MICROSOFT_GRAPH_CLIENT_STATE,
});

const zoom = buildZoomEventSubscriptionRequest({
  webhookUrl: 'https://timeline.example.com/api/platform-events/zoom',
});

const webex = buildWebexWebhookRequests({
  targetUrl: 'https://timeline.example.com/api/platform-events/webex',
  secret: process.env.WEBEX_WEBHOOK_SECRET,
  ownedBy: 'org',
});

const manifest = platformSetupManifest('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

const googleCapabilities = platformCapabilityContract('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

const allCapabilities = allPlatformCapabilityContracts({
  baseUrl: 'https://timeline.example.com',
});

const googlePlan = buildPlatformIntegrationPlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  subscription: {
    name: 'subscriptions/google-subscription-id',
    expireTime: '2026-06-28T02:00:00.000Z',
  },
});

// googlePlan.recommended_mode === 'hybrid_local_observer_first'
// googlePlan.realtime_axis.primary === 'local_observer'
// googlePlan.provider_events.endpoint === 'https://timeline.example.com/api/platform-events/google-meet'
// googlePlan.post_meeting_transcript.strategy === 'import_after_meeting_ends'

const permissionPlan = buildPlatformPermissionPlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  features: ['realtime-axis', 'participants', 'transcript', 'recording', 'security'],
});

// permissionPlan.required_scopes includes meetings.space.readonly and drive.meet.readonly
// permissionPlan.missing_security_env shows whether GOOGLE_PUBSUB_OIDC_AUDIENCE is still missing

const permissionPlans = allPlatformPermissionPlans({
  baseUrl: 'https://timeline.example.com',
  env: process.env,
});

const readiness = evaluatePlatformSetupReadiness('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
});

const teamsMaintenance = evaluatePlatformSubscriptionMaintenance('teams', {
  id: 'graph-subscription-id',
  expirationDateTime: '2026-06-28T02:00:00.000Z',
});

if (teamsMaintenance.renewal_due) {
  const request = buildMicrosoftGraphSubscriptionRenewalRequest({
    subscriptionId: 'graph-subscription-id',
  });
  // PATCH https://graph.microsoft.com/v1.0${request.path}
}
```

`platformSetupManifest()` 会暴露 Google Workspace subscription lifecycle event types、Microsoft Graph lifecycle events、Webex webhook resources。Teams 订阅 request 默认把 `lifecycleNotificationUrl` 指向同一个 webhook endpoint；如果宿主项目用独立 lifecycle endpoint，可以显式传 `lifecycleNotificationUrl` 覆盖。Webex 的 `buildWebexWebhookRequests()` 会按默认资源生成多条 webhook 创建请求，因为 Webex firehose 不覆盖 meetings started/ended 和 meetingParticipants joined/left。

`platformCapabilityContract()` 是给宿主项目做接入决策的机器可读能力表：每个平台会声明实时建轴、参会人轨、发言人轨、会后转写、录制、订阅生命周期、实时转写是否可用，以及对应 SDK normalizer 和 fallback 建议。`buildPlatformPermissionPlan()` 用目标能力反推需要开启的 provider 权限、OAuth scope、webhook 安全环境变量和事件类型，适合配置页或验收脚本先检查“scope/事件/签名密钥是否和目标功能匹配”。`buildPlatformIntegrationPlan()` 会进一步把 capability、manifest、readiness 和 subscription maintenance 合成推荐接入路径：默认策略是 `hybrid_local_observer_first`，也就是本地 URL/window 观察优先建立低延迟会议轴，Google Meet / Teams / Zoom / Webex / Lark 官方事件随后校准或补充 participant/artifact，转写统一在会后导入，不把 transcript 当作实时标注前置依赖。

## Webhook 验证工具

真实接 Zoom / Microsoft Graph / Google Pub/Sub push 时，建议先在 webhook 层完成平台验证，再把 payload 交给 normalizer：

```js
import {
  buildZoomUrlValidationResponse,
  microsoftGraphValidationResponse,
  verifyGooglePubSubOidcJwt,
  verifyMicrosoftGraphClientState,
  verifyWebexWebhookEvent,
  verifyZoomWebhookEvent,
} from '@ai-annotation/meeting-timeline-sdk/adapters/webhook-security';

if (req.query.validationToken) {
  res.type('text/plain').send(microsoftGraphValidationResponse(new URL(req.url, 'https://callback.example')));
  return;
}

if (req.body.event === 'endpoint.url_validation') {
  res.json(buildZoomUrlValidationResponse(req.body.payload, { secretToken: process.env.ZOOM_WEBHOOK_SECRET_TOKEN }));
  return;
}

const verification = verifyZoomWebhookEvent({
  headers: req.headers,
  rawBody: req.rawBody,
  secretToken: process.env.ZOOM_WEBHOOK_SECRET_TOKEN,
});
if (!verification.ok) throw new Error(verification.reason);

const graphCheck = verifyMicrosoftGraphClientState(req.body, {
  clientState: process.env.MICROSOFT_GRAPH_CLIENT_STATE,
});
if (!graphCheck.ok) throw new Error(graphCheck.reason);

const googleCheck = await verifyGooglePubSubOidcJwt({
  headers: req.headers,
  expectedAudience: process.env.GOOGLE_PUBSUB_OIDC_AUDIENCE,
  serviceAccountEmail: process.env.GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL,
});
if (!googleCheck.ok) throw new Error(googleCheck.reason);

const webexCheck = verifyWebexWebhookEvent({
  headers: req.headers,
  rawBody: req.rawBody,
  secret: process.env.WEBEX_WEBHOOK_SECRET,
});
if (!webexCheck.ok) throw new Error(webexCheck.reason);
```

当前 server demo 会读取这些环境变量：

- `ZOOM_WEBHOOK_SECRET_TOKEN`：启用 Zoom URL validation 和事件签名校验。
- `MICROSOFT_GRAPH_CLIENT_STATE`：启用 Microsoft Graph change notification `clientState` 校验。
- `GOOGLE_PUBSUB_OIDC_AUDIENCE`：启用 Google Pub/Sub authenticated push OIDC JWT 校验，匹配 JWT `aud` claim。
- `GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL`：可选，匹配 Pub/Sub push subscription 里配置的 service account email claim。
- `GOOGLE_PUBSUB_BEARER_TOKEN`：没有配置 OIDC 时，对 Google Pub/Sub push 做轻量 bearer 校验，主要用于本地或网关前置鉴权兜底。
- `WEBEX_WEBHOOK_SECRET`：启用 Webex `X-Spark-Signature` HMAC-SHA1 校验。

## 首条标记内联建轴

如果外部项目不方便单独调用 `startMeeting`，可以在第一条标记里带 `meetingSession`：

```js
await timeline.insertMark({
  id: 'mark-inline-001',
  capturedAtMs: Date.now(),
  label: 'follow up',
  textCandidates: ['follow up'],
  meetingSession: {
    meetingId: 'meeting-001',
    title: '产品评审',
    startTimeMs: Date.now() - 30_000,
    detectorSource: 'host_app',
  },
});
```

## 关键约束

- 推荐总是传 `capturedAtMs`。这是“标到正确时间轴位置”的核心字段。
- `capturedAtMs` 应该是设备/宿主采集到标记完成时的绝对时间，不是上传时间。
- 若缺失 `capturedAtMs`，服务端只能退回到收到请求的时间，实时性和会后补传都会变差。
- `insertMark` 默认会校验 `capturedAtMs`；如确实要允许服务端收包时间兜底，可传 `{ requireCapturedAt: false }`。

## 批量写入

```js
await timeline.insertMarks([
  { id: 'm1', capturedAtMs: t1, label: 'why?', textCandidates: ['why?'] },
  { id: 'm2', capturedAtMs: t2, label: '重点', intent: 'attention' },
]);
```

## 查询单条标记状态

```js
const status = await timeline.getAnnotationStatus('mark-001');
```

状态里会包含是否落在真实会议轴、归一化后的会议内时间、是否晚于会议结束等信息。
