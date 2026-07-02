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

- `@ai-annotation/meeting-timeline-sdk/adapters/google-meet`
- `@ai-annotation/meeting-timeline-sdk/adapters/microsoft-teams`
- `@ai-annotation/meeting-timeline-sdk/adapters/zoom`
- `@ai-annotation/meeting-timeline-sdk/adapters/webhook-security`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-setup`

`meeting_started` 会调用 `startMeeting`，`meeting_ended` 会调用 `endMeeting`。`participant_joined/left` 和 `artifact_ready` 默认不会写入用户标注流；如果需要临时显示参会人位置，可以给 `applyMeetingSignals` 传 `{ participantAsAnnotation: true }`，或者用 `onParticipantSignal` / `onArtifactSignal` 接到自己的服务端轨道。`subscription_lifecycle` 表示平台订阅自身的过期、移除、暂停、漏投或重新授权要求，默认不画到会议轴；需要接入诊断时传 `onSubscriptionLifecycleSignal` 处理。

Google Meet adapter 同时支持已经解包的 Workspace Events CloudEvent，以及 Pub/Sub 默认 wrapped push body。wrapped body 会自动 base64 解码 `message.data`，所以 webhook handler 可以直接把 `req.body` 传给 `normalizeGoogleMeetEvent(req.body)`。Google Workspace Events 的 `subscription.v1.suspended`、`subscription.v1.expirationReminder`、`subscription.v1.expired` 会归一化为 `subscription_lifecycle`。Microsoft Graph change notifications 的 `lifecycleEvent` 值 `reauthorizationRequired`、`subscriptionRemoved`、`missed` 也会归一化为 `subscription_lifecycle`。

## 平台接入配置

`platform-setup` 提供只读 manifest 和订阅 request body builder，方便宿主项目生成配置页或自动化脚本：

```js
import {
  buildGoogleMeetWorkspaceSubscriptionRequest,
  buildMicrosoftGraphSubscriptionRenewalRequest,
  buildMicrosoftTeamsMeetingCallSubscriptionRequest,
  buildZoomEventSubscriptionRequest,
  evaluatePlatformSubscriptionMaintenance,
  evaluatePlatformSetupReadiness,
  platformSetupManifest,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-setup';

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

const manifest = platformSetupManifest('google-meet', {
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

`platformSetupManifest()` 会暴露 Google Workspace subscription lifecycle event types 和 Microsoft Graph lifecycle events。Teams 订阅 request 默认把 `lifecycleNotificationUrl` 指向同一个 webhook endpoint；如果宿主项目用独立 lifecycle endpoint，可以显式传 `lifecycleNotificationUrl` 覆盖。

## Webhook 验证工具

真实接 Zoom / Microsoft Graph / Google Pub/Sub push 时，建议先在 webhook 层完成平台验证，再把 payload 交给 normalizer：

```js
import {
  buildZoomUrlValidationResponse,
  microsoftGraphValidationResponse,
  verifyGooglePubSubOidcJwt,
  verifyMicrosoftGraphClientState,
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
```

当前 server demo 会读取这些环境变量：

- `ZOOM_WEBHOOK_SECRET_TOKEN`：启用 Zoom URL validation 和事件签名校验。
- `MICROSOFT_GRAPH_CLIENT_STATE`：启用 Microsoft Graph change notification `clientState` 校验。
- `GOOGLE_PUBSUB_OIDC_AUDIENCE`：启用 Google Pub/Sub authenticated push OIDC JWT 校验，匹配 JWT `aud` claim。
- `GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL`：可选，匹配 Pub/Sub push subscription 里配置的 service account email claim。
- `GOOGLE_PUBSUB_BEARER_TOKEN`：没有配置 OIDC 时，对 Google Pub/Sub push 做轻量 bearer 校验，主要用于本地或网关前置鉴权兜底。

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
