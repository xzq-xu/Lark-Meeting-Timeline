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
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler`
- `@ai-annotation/meeting-timeline-sdk/adapters/transcript`
- `@ai-annotation/meeting-timeline-sdk/adapters/webhook-security`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-setup`

`meeting_started` 会调用 `startMeeting`，`meeting_ended` 会调用 `endMeeting`。`participant_joined/left`、`speaker_started/ended` 和 `artifact_ready` 默认不会写入用户标注流；如果需要临时显示参会人位置，可以给 `applyMeetingSignals` 传 `{ participantAsAnnotation: true }`，如果需要显示发言人位置，可以传 `{ speakerAsAnnotation: true }`，或者用 `onParticipantSignal` / `onSpeakerSignal` / `onArtifactSignal` 接到自己的服务端轨道。`subscription_lifecycle` 表示平台订阅自身的过期、移除、暂停、漏投或重新授权要求，默认不画到会议轴；需要接入诊断时传 `onSubscriptionLifecycleSignal` 处理。

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
import { ingestPlatformEvent } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });

await ingestPlatformEvent(timeline, 'google-meet', req.body, {
  receivedAtMs: Date.now(),
  participantAsAnnotation: true,
});
```

`ingestPlatformEvent()` 内部会按平台名选择 normalizer，把原始事件转成 `NormalizedMeetingSignal[]`，再调用 `startMeeting`、`endMeeting` 或可选的 participant/artifact handler。对于 Google Meet / Teams / Zoom / Webex，新项目可以优先接这一层，只有需要自定义事件验签、补拉详情或 artifact 导入时再下钻到 registry/core。

如果外部项目想直接复用完整 HTTP webhook 入口，可以用 `platform-webhook-handler`。它会处理 Microsoft Graph `validationToken`、Zoom `endpoint.url_validation`、Zoom/Webex/Teams/Google Pub/Sub 验证，再把事件交给 `ingestPlatformEvent()`：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { createPlatformWebhookHandler } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const handleWebhook = createPlatformWebhookHandler(timeline, {
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

浏览器扩展、桌面观察器或电子纸宿主 App 不需要等官方 webhook，可以先用 `local-observer` 把窗口 URL 快照转成 meeting signal，再交给 `applyMeetingSignals()`：

```js
import { applyMeetingSignals } from '@ai-annotation/meeting-timeline-sdk/adapters/core';
import { createLocalMeetingObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/local-observer';

const observer = createLocalMeetingObserver({ source: 'desktop_observer' });
const observed = observer.observe({
  url: 'https://meet.google.com/abc-defg-hij',
  observedAtMs: Date.now(),
});

await applyMeetingSignals(timeline, observed.signals);
```

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

事件 adapter 只负责告诉你 transcript/recording 已生成；正文内容建议会后拉取后再导入。SDK 提供平台中性的 `importTranscript()`，也提供更高层的 `importPlatformTranscript()`，会按 `platform` 自动选择 normalizer 并导入：

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

`platformCapabilityContract()` 是给宿主项目做接入决策的机器可读能力表：每个平台会声明实时建轴、参会人轨、发言人轨、会后转写、录制、订阅生命周期、实时转写是否可用，以及对应 SDK normalizer 和 fallback 建议。当前策略是实时标注只依赖 meeting axis，转写统一标记为 `post_meeting` 导入，不把 transcript 当作实时链路前置依赖。

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
