# 多会议平台时间轴适配方案

本文档定义会议时间轴 SDK 后续如何适配 Google Meet、Microsoft Teams、Zoom 等会议软件。目标不是为每个平台重写一套时间轴，而是把不同平台的会议事件归一化为同一套 Meeting Timeline Signal，再调用现有 SDK 协议建轴、插入标注、结束轴和导入会后产物。

## 目标

- SDK 核心保持平台无关：只关心会议轴、标注、会后转写和状态订阅。
- 各会议平台差异下沉到 adapter：OAuth、Webhook、事件验签、权限、事件字段映射都不污染核心协议。
- 实时标注优先低延迟：电子纸/桌面端检测到正在开会时，可以先调用本地建轴；官方会议事件用于校准、补齐和会后追溯。
- 转写暂不要求实时：会议中只需要标注能实时落到正确时间，转写和录制可以在会议结束后导入。

## 非目标

- 不做跨平台实时转写 SDK。
- 不让标注设备直接对接 Google、Microsoft、Zoom 的 OAuth 和 Webhook。
- 不把平台原始事件格式暴露给上层业务。
- 不依赖某个平台的会议开始事件作为唯一建轴来源，因为真实事件投递可能有延迟、权限缺失或订阅失效。

## 现有稳定协议

当前 SDK 已经围绕以下服务端协议封装：

```text
POST /api/meeting-session/start
POST /api/annotations
POST /api/annotations/batch
POST /api/meeting-session/end
GET  /api/annotations/status?id=...
GET  /api/stream
```

跨平台适配不要改变这组主协议。新增平台只需要把外部事件归一化后调用：

```js
await timeline.startMeeting(...)
await timeline.insertMark(...)
await timeline.endMeeting(...)
```

## 总体架构

```text
Google Meet Events / Teams Graph / Zoom Webhook / Webex Webhook / Lark Events
                 |
                 v
        platform adapter
        - verify
        - normalize
        - dedupe
        - timestamp mapping
                 |
                 v
        NormalizedMeetingSignal
                 |
                 v
        meeting timeline SDK
                 |
                 v
        local timeline server
        - meeting axis
        - annotations
        - participants
        - post-meeting artifacts
```

关键原则：

1. 官方会议平台事件只是一种建轴来源，不是唯一来源。
2. 标注设备永远只写 `captured_at_ms`，不需要理解会议平台。
3. 会议轴的 `start_time_ms` 和 `end_time_ms` 必须来自事件发生时间或本地检测时间，不能用服务端收包时间替代。
4. 会后转写、录制、智能纪要是 artifact，不阻塞实时标注。

## 归一化事件模型

建议后续在 SDK 或服务端 adapter 层引入这个平台无关模型：

```ts
type MeetingPlatform =
  | 'lark'
  | 'google_meet'
  | 'microsoft_teams'
  | 'zoom'
  | 'webex'
  | 'manual'
  | 'local_detector';

interface NormalizedMeetingIdentity {
  platform: MeetingPlatform;
  meeting_id: string;
  external_meeting_id?: string;
  meeting_url?: string;
  title?: string;
  organizer_id?: string;
  organizer_name?: string;
}

type NormalizedMeetingSignal =
  | {
      type: 'meeting_started';
      meeting: NormalizedMeetingIdentity;
      occurred_at_ms: number;
      source_event_id?: string;
      source: 'webhook' | 'long_connection' | 'polling' | 'local_detector' | 'manual';
      raw?: unknown;
    }
  | {
      type: 'meeting_ended';
      meeting: NormalizedMeetingIdentity;
      occurred_at_ms: number;
      source_event_id?: string;
      source: 'webhook' | 'long_connection' | 'polling' | 'local_detector' | 'manual';
      raw?: unknown;
    }
  | {
      type: 'participant_joined' | 'participant_left';
      meeting: NormalizedMeetingIdentity;
      occurred_at_ms: number;
      participant_id?: string;
      participant_name?: string;
      source_event_id?: string;
      raw?: unknown;
    }
  | {
      type: 'artifact_ready';
      meeting: NormalizedMeetingIdentity;
      artifact_kind: 'transcript' | 'recording' | 'smart_notes';
      occurred_at_ms: number;
      artifact_id?: string;
      artifact_url?: string;
      source_event_id?: string;
      raw?: unknown;
    }
  | {
      type: 'subscription_lifecycle';
      platform: MeetingPlatform;
      lifecycle_type:
        | 'expiration_reminder'
        | 'expired'
        | 'suspended'
        | 'reauthorization_required'
        | 'subscription_removed'
        | 'missed'
        | string;
      occurred_at_ms: number;
      subscription_id?: string;
      subscription_name?: string;
      expires_at_ms?: number;
      resource?: string;
      source_event_id?: string;
      raw?: unknown;
    };
```

映射到现有 SDK：

| Normalized signal | SDK/server action |
| --- | --- |
| `meeting_started` | `timeline.startMeeting({ platform, meeting_id, start_time_ms })` |
| `meeting_ended` | `timeline.endMeeting({ meeting_id, end_time_ms })` |
| `participant_joined/left` | 后续新增 participant track，或先作为 event row 绘制 |
| `artifact_ready: transcript` | 先写入事件轨道表示转写已生成；正文 segment 通过 `adapters/transcript` + `POST /api/import/transcript` 会后导入 |
| `artifact_ready: recording` | 写入事件轨道并保存 artifact 元数据，不进入用户标注主链路 |
| `subscription_lifecycle` | 写入平台接入诊断状态，不进入会议时间轴 |

## 平台适配矩阵

| 平台 | 实时建轴 | 参会人位置 | 会后转写/录制 | 主要风险 |
| --- | --- | --- | --- | --- |
| 飞书/Lark | 当前已接入长连接事件和会议扫描兜底 | join/leave 事件可做 speaker/participant track | 妙记会后导入 | 企业权限、事件投递延迟 |
| Google Meet | 本地 `meeting-apps`/browser 预设先建轴；Workspace Events conference started/ended 校准 | 本地 active speaker tile 或 participant joined/left events | Meet REST `conferenceRecords.transcripts/entries`、recordings、smartNotes | Workspace 权限、订阅目标限制、部分参与者只能收到有限事件 |
| Microsoft Teams | 本地 `meeting-apps`/native 预设先建轴；Graph `meetingCallEvents` 或 Teams bot 校准 | 本地 active speaker tile、Graph rosterUpdated 或 Teams bot participant events | Graph transcript/recording notifications，会后获取内容 | Graph 应用权限、rich notification 加密、订阅最长 3 天、租户管理员可能关闭 transcript API |
| Zoom | 本地 `meeting-apps`/native 预设先建轴；Zoom Meeting webhooks 校准 | 本地 active speaker tile 或 meeting participant webhook | `recording.completed` 后取录制和转写文件 | HTTPS webhook、事件 scope、3 秒响应要求、云录制/转写设置 |
| Cisco Webex | Webex webhooks 的 meetings started/ended | `meetingParticipants` joined/left webhook | `meetingTranscripts` created、recordings created/updated | webhook payload 可能只有元数据、完整内容需 REST 补拉、FedRAMP 支持范围不同 |

## Google Meet 适配

官方能力：

- Google Workspace Events API 可以订阅 Meet conference start/end、participant join/leave、recording/transcript/smart note start/end/file generated 等事件。
- 订阅目标可以是 meeting space，也可以是 user。user 目标会收到该用户拥有的会议空间相关事件。
- Meet REST API 提供 `conferenceRecords`、participants、participantSessions、recordings、smartNotes、transcripts、transcript entries 等资源。

推荐接入方式：

1. P0：先支持本地/桌面检测器建轴。浏览器扩展或 WebView 用 `adapters/meeting-app-capture` 低成本读取 Google Meet 的 URL、按钮文案、tile/ariaLabel、active speaker 信息，再交给 `adapters/meeting-apps` preset 归一成 `meeting_started` / `speaker_started`，最后通过 `meeting-source` 调用 `startMeeting({ platform: 'google_meet', start_time_ms })`。
2. P1：接 Google Workspace Events API，处理：
   - `google.workspace.meet.conference.v2.started`
   - `google.workspace.meet.conference.v2.ended`
   - `google.workspace.meet.participant.v2.joined`
   - `google.workspace.meet.participant.v2.left`
   - `google.workspace.meet.transcript.v2.fileGenerated`
3. P2：会议结束后用 Meet REST API 拉取 transcript entries 和 recording 元数据，通过 `normalizeGoogleMeetTranscriptEntries()` 导入时间轴 transcript track。

注意点：

- 官方事件适合校准轴和会后补齐，但不应阻塞电子纸标注实时落轴。
- Pub/Sub push 默认是 wrapped JSON，事件在 `message.data`，需要 base64 解码；当前 Google adapter 已支持 wrapped push body 和已解包 CloudEvent 两种输入。
- Workspace Events 订阅生命周期事件 `subscription.v1.suspended`、`subscription.v1.expirationReminder`、`subscription.v1.expired` 是平台健康信号，不对应某场会议；adapter 会归一化为 `subscription_lifecycle`，服务端只更新 `/api/platform-events/status`。
- 事件 payload 可能只给 resource name，需要再调 REST API 获取详情。
- Transcript entry 和 Google Docs transcript 可能不完全一致，需要保留 provider/source 字段。

参考：

- [Subscribe to Google Meet events](https://developers.google.com/workspace/events/guides/events-meet)
- [Google Meet REST API reference](https://developers.google.com/workspace/meet/api/reference/rest/v2)
- [Google Meet transcript entries](https://developers.google.com/workspace/meet/api/reference/rest/v2/conferenceRecords.transcripts.entries/get)

## Microsoft Teams 适配

官方能力有两条路线：

### 路线 A：Microsoft Graph meetingCallEvents

Graph change notifications 支持对 Teams online meeting 订阅 call started、call ended 和 roster updated。订阅资源形如：

```text
/communications/onlineMeetings(joinWebUrl='{url-encoded-joinWebUrl}')/meetingCallEvents
```

推荐用于后端平台级接入。

优点：

- 能拿会议真实开始、结束和 roster 变化。
- 适合没有安装 Teams bot 的会议。

风险：

- 需要 application permission，例如 `OnlineMeetings.Read.All` 或 `OnlineMeetings.ReadWrite.All`。
- 官方建议 active meeting call 用 rich notifications；basic notification 信息太少。
- 订阅最长 3 天，需要自动续订。
- `lifecycleNotificationUrl` 会收到 `reauthorizationRequired`、`subscriptionRemoved`、`missed` 等 lifecycle notification。它们不对应具体会议；adapter 会归一化为 `subscription_lifecycle`，用于提示重新授权、重建订阅或补拉漏投事件。

### 路线 B：Teams app/bot meeting events

Teams SDK 可以让会议内 app/bot 接收 meetingStart、meetingEnd、participant join/leave 等事件。

推荐用于企业内深度集成或会议侧边栏应用。

优点：

- 会议内上下文更明确。
- 对 participant track 更直接。

风险：

- 要做 Teams app manifest、bot、RSC 权限和安装流程。
- 用户/企业部署成本比 Graph webhook 更高。

会后转写：

- Graph 支持 transcript/recording change notifications。
- transcript 通常在会议结束后生成，适合作为 artifact 导入，不适合作为实时标注链路依赖。
- 租户管理员可以关闭 Graph transcript API，需在诊断里明确暴露。

参考：

- [Microsoft Graph meeting call event notifications](https://learn.microsoft.com/en-us/graph/changenotifications-for-onlinemeeting)
- [Microsoft Teams meeting events](https://learn.microsoft.com/en-us/microsoftteams/platform/teams-sdk/in-depth-guides/meeting-events)
- [Teams transcript and recording notifications](https://learn.microsoft.com/en-us/graph/teams-changenotifications-callrecording-and-calltranscript)
- [Teams transcript/recording overview](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/overview-transcripts)

## Zoom 适配

官方能力：

- Zoom Meeting webhooks 可以接收会议和 webinar 事件。
- Webhook body 包含 `event`、`event_ts`、`payload`。
- Zoom webhook 需要公开 HTTPS endpoint，能接收 JSON POST，并在要求时间内返回成功状态。
- 事件订阅受 scope 控制，例如 participant joined 和 recording completed 需要相应 read scope。

推荐接入方式：

1. P0：本地/桌面检测器建轴，会议 URL 或窗口标题作为 `meeting_url/title`。
2. P1：接 Zoom webhook，把 meeting started/ended、participant joined/left 归一化为 Meeting Signal。
3. P2：`recording.completed` 后导入录制和转写 artifact。

注意点：

- Zoom webhook 事件应使用 `event_ts` 作为 `occurred_at_ms`，不要用接收时间。
- endpoint 必须快速确认，耗时处理放到队列。
- 缺少 scope 时事件订阅会失败，需要在 adapter diagnostics 暴露。

参考：

- [Zoom Meetings webhooks](https://developers.zoom.us/docs/api/meetings/events/)
- [Using Zoom webhooks](https://developers.zoom.us/docs/api/webhooks/)

## Cisco Webex 适配

官方能力：

- Webex webhook 支持 meeting related resources：`meetings`、`meetingParticipants`、`recordings`、`meetingTranscripts` 等。
- `meetings` 支持 `started` / `ended`，`meetingParticipants` 支持 `joined` / `left`，`meetingTranscripts` 支持 `created`。
- `all` firehose 不包含 meetings `started/ended` 和 meetingParticipants `joined/left`，所以必须为这些资源/事件单独注册 webhook。
- 创建 webhook 时可设置 `secret`，Webex 会用 `X-Spark-Signature` 发送 JSON payload 的 HMAC-SHA1 签名。

推荐接入方式：

1. P0：仍然保留 local detector path，保证电子纸标注立即落到当前轴。
2. P1：注册 Webex meeting webhooks：
   - `meetings:started`
   - `meetings:ended`
   - `meetingParticipants:joined`
   - `meetingParticipants:left`
   - `recordings:created`
   - `recordings:updated`
   - `meetingTranscripts:created`
3. P2：会议结束后用 Meeting Transcripts API 或 transcript download link 拉取 VTT/text，通过 `normalizeWebexTranscript()` 和 `/api/import/transcript` 导入。

注意点：

- Webex webhook 的 `data` 可能只包含元数据或资源 ID，完整标题、参与者、转写正文、录制下载地址可能需要宿主项目使用 OAuth token 再调 REST API 补齐。
- Webex Meeting Transcripts API 在 2026 年 1 月更新后支持 Cisco AI Assistant 生成的 transcripts，但具体租户设置仍会影响是否可用。
- Webex webhook 会长期运行，不像 Graph/Workspace Events 那样需要短周期续订；但是目标 URL 连续失败会导致 webhook 被禁用，需要监控状态。
- FedRAMP 环境下会议相关 webhook/meetingTranscripts 支持范围要单独核验，不能默认和商业环境一致。

参考：

- [Webex Webhooks guide](https://developer.webex.com/messaging/docs/api/guides/webhooks)
- [Create a Webhook](https://developer.webex.com/messaging/docs/api/v1/webhooks/create-a-webhook)
- [Webex Meeting Transcripts](https://developer.webex.com/docs/api/v1/meeting-transcripts)

## 统一接入优先级

建议按以下顺序开发：

1. **保留并强化 local detector path**
   - 宿主应用、桌面观察器、电子纸 companion app 发现“用户已经在会议中”时，直接调用本地 `startMeeting`。
   - 对 Google Meet / Teams / Zoom / Lark / Webex 的浏览器 DOM，先用 `adapters/meeting-app-capture` 采集按钮、participant tile、ariaLabel 和音量/发言状态，再用 `adapters/meeting-apps` preset 归一化。
   - 对桌面 Accessibility 快照，直接交给 `adapters/meeting-apps` 或 `adapters/native-meeting`；它们负责识别 Leave/Join 按钮、participant tile、ariaLabel 和 active speaker。
   - 这是跨平台最低延迟、最低权限依赖的路径。

2. **抽出 adapter core**
   - SDK 已增加 `NormalizedMeetingSignal` 的运行时归一化和 `applyMeetingSignal(client, signal)` / `applyMeetingSignals(client, signals)`。
   - 先不绑定任何具体平台，保证 server 和外部项目都可以复用。

3. **Google Meet adapter**
   - 因为 Google Workspace Events 的会议 start/end、participant、transcript artifact 模型最接近我们的抽象。
   - 先做 webhook/event normalization，不急着做完整 transcript 导入。

4. **Microsoft Teams adapter**
   - 先 Graph `meetingCallEvents`，后 Teams app/bot。
   - 必须把 rich notification、订阅续期、tenant admin 禁用 transcript 的状态写进 diagnostics。

5. **Zoom adapter**
   - 先做 webhook receiver 和签名校验。
   - `recording.completed` 只作为会后 artifact，不影响实时标注验收。

6. **Webex adapter**
   - 先做 meetings/meetingParticipants/recordings/meetingTranscripts webhook normalization。
   - 用 `X-Spark-Signature` 验证请求；transcript 正文仍走会后导入。

## SDK/服务端边界

已新增但不替换现有接口：

```text
POST /api/platform-events/:platform
GET  /api/platform-events/setup
GET  /api/platform-events/:platform/setup
GET  /api/platform-events/:platform/status
GET  /api/platform-events/status
```

`POST /api/platform-events/:platform` 当前支持 `local-detector`、`lark`/`feishu`、`google-meet`、`teams`、`zoom`、`webex` 及其别名。服务端会用 SDK adapter 归一化原始事件，`meeting_started` 进入 `POST /api/meeting-session/start` 同一套建轴逻辑，`meeting_ended` 进入 `POST /api/meeting-session/end` 同一套结束逻辑。`participant_joined/left` 会进入会议 `events` 轨道，不写入用户标注流；同一平台、同一参会人、同一 join/leave 类型在默认 3 秒窗口内会被过滤为重复事件。会后 transcript/recording/smart notes 的 `artifact_ready` signal 也会进入会议 `events` 轨道，默认 5 秒窗口内按 artifact id/url 去重；transcript 正文通过 SDK `adapters/transcript` 归一化后调用 `POST /api/import/transcript` 导入。`subscription_lifecycle` 只更新平台状态，不会创建会议轴，也不会写入用户标注或会议事件轨道。

`platformCapabilityContract(platform)` 会输出平台能力契约，供宿主项目决定接入路径：

```ts
{
  platform: 'local_detector' | 'lark' | 'google_meet' | 'microsoft_teams' | 'zoom' | 'webex',
  realtime_axis: { status: 'supported_best_effort', fallback: 'local_detector_recommended_...' },
  participant_track: { status: 'supported_best_effort' },
  post_meeting_transcript: { status: 'supported', availability: 'post_meeting', import_endpoint: '/api/import/transcript' },
  recording: { status: 'supported' | 'metadata_supported' },
  subscription_lifecycle: { status: 'supported' | 'not_applicable' },
  realtime_transcript: { status: 'not_supported' },
  sdk_modules: { events, transcript, setup, security },
  limitations: string[]
}
```

关键约束是：实时标注只依赖 `realtime_axis`，转写统一通过 `post_meeting_transcript` 会后导入；任何平台的实时 transcript 都不作为 P0/P1 链路前置条件。

SDK 还导出 `MEETING_PLATFORM_KEYS`、`MEETING_PLATFORM_ALIASES` 和 `normalizeMeetingPlatform()`，宿主项目应从这里读取平台列表和别名映射。新增平台时必须同时补齐 event adapter、setup manifest、capability contract、endpoint、安全配置，以及适用的 transcript normalizer，并通过 `test/sdk-platform-conformance.test.mjs`。

真实 webhook 接入的安全层也已经放进 SDK：

- `@ai-annotation/meeting-timeline-sdk/adapters/webhook-security`
- Zoom：支持 `endpoint.url_validation` challenge response；配置 `ZOOM_WEBHOOK_SECRET_TOKEN` 后校验 `x-zm-request-timestamp` 和 `x-zm-signature`。
- Microsoft Graph / Teams：支持 `validationToken` 纯文本响应；配置 `MICROSOFT_GRAPH_CLIENT_STATE` 后校验通知里的 `clientState`。
- Google Meet / Pub/Sub：配置 `GOOGLE_PUBSUB_OIDC_AUDIENCE` 后会校验 authenticated push 的 Google-signed OIDC JWT，并可用 `GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL` 限定 service account email；未配置 OIDC 时可用 `GOOGLE_PUBSUB_BEARER_TOKEN` 做轻量 bearer gate。
- Webex：配置 `WEBEX_WEBHOOK_SECRET` 后校验 `X-Spark-Signature` HMAC-SHA1。

`GET /api/platform-events/status` 会返回每个平台最近一次 `last_verification`，用于区分“未配置所以跳过校验”和“签名/状态不匹配被拒绝”。同时会返回 `lifecycle_event_count` 和 `last_lifecycle`，用于观察订阅过期提醒、暂停、移除、漏投和重新授权要求。

`GET /api/platform-events/setup` 会返回 Google Meet、Microsoft Teams、Zoom、Webex 的接入 manifest：默认事件类型/资源、endpoint、权限/环境变量要求和操作步骤，同时包含 readiness 诊断，检查 endpoint 是否是 HTTPS/localhost、必需安全环境变量是否已配置。`GET /api/platform-events/:platform/setup` 可按平台返回，并支持用 query 生成订阅 request body，例如 Teams 的 `join_web_url` + `client_state`，Google 的 `target_resource` + `pubsub_topic`，或 Webex 的 `webex_subscription_name` + `webex_secret`。同一接口还会基于 `subscription_expires_at`、`subscription_id`、`subscription_name` 等 query 返回 maintenance 建议，用于 Graph / Workspace Events 订阅续期调度；Zoom/Webex 返回无需短周期续订。

SDK 包结构建议：

```text
packages/meeting-timeline-sdk/
  index.mjs
  index.d.ts
  adapters/
    platform-setup.mjs
    platform-setup.d.ts
    platform-registry.mjs
    platform-registry.d.ts
    platform-ingest.mjs
    platform-ingest.d.ts
    webhook-security.mjs
    webhook-security.d.ts
    transcript.mjs
    transcript.d.ts
    core.mjs
    core.d.ts
    local-detector.mjs
    local-detector.d.ts
    lark.mjs
    lark.d.ts
    google-meet.mjs
    google-meet.d.ts
    microsoft-teams.mjs
    microsoft-teams.d.ts
    zoom.mjs
    zoom.d.ts
    webex.mjs
    webex.d.ts
```

`adapters/core` 只做平台无关的信号定义、校验和应用：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { normalizeGoogleMeetEvent } from '@ai-annotation/meeting-timeline-sdk/adapters/google-meet';
import { applyMeetingSignal } from '@ai-annotation/meeting-timeline-sdk/adapters/core';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const signals = normalizeGoogleMeetEvent(req.body);

for (const signal of signals) {
  await applyMeetingSignal(timeline, signal);
}
```

对外部宿主项目，推荐优先使用更高层的 `platform-ingest`，把“平台名 + 原始事件”直接接入时间轴：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { ingestPlatformEvent } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });

await ingestPlatformEvent(timeline, req.params.platform, req.body, {
  receivedAtMs: Date.now(),
  participantAsAnnotation: true,
});
```

这层只负责选择平台 adapter、归一化事件并调用 `startMeeting` / `endMeeting` / 可选参会人标注。平台验签、OAuth、REST 补拉和会后 artifact 下载仍由宿主 webhook 层处理。

## 诊断字段

每个平台 adapter 都应该输出统一 diagnostics，方便现场定位问题：

```ts
interface PlatformAdapterDiagnostics {
  platform: MeetingPlatform;
  receiver_ready: boolean;
  auth_ready: boolean;
  subscription_ready: boolean;
  last_event_at_ms?: number;
  last_event_type?: string;
  last_delivery_delay_ms?: number;
  last_error_code?: string;
  next_action?: string;
}
```

核心诊断判断：

- 事件接收器是否可达。
- OAuth/app permission 是否满足。
- 官方事件是否真的投递过。
- 最近一次投递延迟是多少。
- 当前会议轴是否来自官方事件、本地检测器、扫描兜底或手动建轴。

## 产品判断

对我们的产品来说，最可靠的跨平台方案不是“等官方会议事件来了再建轴”，而是：

1. 本地检测器或宿主应用先建轴，保证电子纸标注实时落点。
2. 官方会议事件后到时做校准和去重。
3. 会后 transcript/recording 作为 artifact 补齐上下文。

这与当前飞书联调中遇到的问题一致：官方事件投递可能慢、权限可能不清晰，但用户写下标注的时间不能等。因此 SDK 的主方向应当是统一时间轴协议和本地低延迟入口，平台 adapter 负责增强可信度和补充会议信息。
