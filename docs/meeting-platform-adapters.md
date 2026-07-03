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

1. P0：先支持本地/桌面检测器建轴。浏览器扩展 content script 优先用 `adapters/meeting-app-browser-runtime`，Electron WebView 或自定义宿主可直接用 `adapters/meeting-app-runtime`；runtime 内部由 `meeting-app-monitor` 驱动 `meeting-app-capture` 低成本读取 Google Meet 的 URL、按钮文案、tile/ariaLabel、active speaker 信息，再交给 `adapters/meeting-apps` preset 归一成 `meeting_started` / `speaker_started`，最后通过 `meeting-source` 调用 `startMeeting({ platform: 'google_meet', start_time_ms })`。monitor 负责轮询、去重、限流和 keep-alive，避免调用方自己处理 active speaker 稳定窗口。
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
   - 对 Google Meet / Teams / Zoom / Lark / Webex 的浏览器 DOM，浏览器扩展优先用 `adapters/meeting-app-browser-runtime`，其他宿主用 `adapters/meeting-app-runtime`；内部用 `meeting-app-monitor` 管理轮询和 keep-alive，用 `meeting-app-capture` 采集按钮、participant tile、ariaLabel 和音量/发言状态，再用 `meeting-apps` preset 归一化。
   - 每个平台进入 P0 前都要跑真实 DOM 采样 gate：`npm run meeting-app:extension:build` 生成并构建扩展，在真实会议页保存 `window.__meetingTimelineLiveCapture.exportRecords()` 或 `evidencePackage()`，再用 `npm run meeting-app:evidence-gate -- --input=<evidence.json>` 验收。默认不允许 fixture 兜底，只有 `production_ready=true` 才算该平台的本地观察路径可交付。多平台状态用 `npm run meeting-app:evidence-matrix` 汇总，逐项追踪 Google Meet / Teams / Zoom / Webex / Lark 的真实 DOM 缺口。
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

如果宿主项目只需要一个更直接的决策对象，而不是完整 setup/rollout 细节，可以调用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy`。`buildMeetingPlatformAdaptationStrategyMatrix()` 会把每个平台收敛成同一组字段：实时轴主来源、provider 事件是否阻塞实时、转写是否阻塞实时、发言人位置来源、pilot gate、production gate 和 handoff package 入口。这里的策略结论固定为：Google Meet、Teams、Zoom、Webex、Lark 的 provider 事件都不能作为当前标注的唯一低延迟时钟；真实用户边写边标注时，必须优先使用本地观察或 host detector 建轴，provider 事件用于回填和审计。

SDK 还导出 `MEETING_PLATFORM_KEYS`、`MEETING_PLATFORM_ALIASES` 和 `normalizeMeetingPlatform()`，宿主项目应从这里读取平台列表和别名映射。新增平台时必须同时补齐 event adapter、setup manifest、capability contract、endpoint、安全配置，以及适用的 transcript normalizer，并通过 `test/sdk-platform-conformance.test.mjs`。

真实 webhook 接入的安全层也已经放进 SDK：

- `@ai-annotation/meeting-timeline-sdk/adapters/webhook-security`
- Zoom：支持 `endpoint.url_validation` challenge response；配置 `ZOOM_WEBHOOK_SECRET_TOKEN` 后校验 `x-zm-request-timestamp` 和 `x-zm-signature`。
- Microsoft Graph / Teams：支持 `validationToken` 纯文本响应；配置 `MICROSOFT_GRAPH_CLIENT_STATE` 后校验通知里的 `clientState`。
- Google Meet / Pub/Sub：配置 `GOOGLE_PUBSUB_OIDC_AUDIENCE` 后会校验 authenticated push 的 Google-signed OIDC JWT，并可用 `GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL` 限定 service account email；未配置 OIDC 时可用 `GOOGLE_PUBSUB_BEARER_TOKEN` 做轻量 bearer gate。
- Webex：配置 `WEBEX_WEBHOOK_SECRET` 后校验 `X-Spark-Signature` HMAC-SHA1。

`GET /api/platform-events/status` 会返回每个平台最近一次 `last_verification`，用于区分“未配置所以跳过校验”和“签名/状态不匹配被拒绝”。同时会返回 `lifecycle_event_count` 和 `last_lifecycle`，用于观察订阅过期提醒、暂停、移除、漏投和重新授权要求。

`GET /api/platform-events/setup` 会返回 Google Meet、Microsoft Teams、Zoom、Webex 的接入 manifest：默认事件类型/资源、endpoint、权限/环境变量要求和操作步骤，同时包含 readiness 诊断，检查 endpoint 是否是 HTTPS/localhost、必需安全环境变量是否已配置。`GET /api/platform-events/:platform/setup` 可按平台返回，并支持用 query 生成订阅 request body，例如 Teams 的 `join_web_url` + `client_state`，Google 的 `target_resource` + `pubsub_topic`，或 Webex 的 `webex_subscription_name` + `webex_secret`。同一接口还会基于 `subscription_expires_at`、`subscription_id`、`subscription_name` 等 query 返回 maintenance 建议，用于 Graph / Workspace Events 订阅续期调度；Zoom/Webex 返回无需短周期续订。

SDK 侧如果要把 provider 订阅创建交给另一个项目执行，优先用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-subscription-handoff` 或仓库命令 `npm run meeting-platform:subscription-handoff`。它在 `platform-provider-connection` 的安全、权限和事件映射基础上，进一步输出 Google Workspace Events、Microsoft Graph、Zoom、Webex 的创建请求或控制台配置摘要，并把 Lark/飞书标为 `manual_setup`。这份 handoff 的职责是“订阅能不能创建、缺哪些参数/secret、创建后要保存哪些 id/expiration”，不是实时轴来源；即使 `ready_to_create=true`，用户边写边标注仍必须使用宿主本地观察或设备上报的 `captured_at_ms`。

跨平台正式推进时，宿主项目应调用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout` 做统一 gate。它会把官方 provider event gate 和本地会议 App DOM gate 合成一个 `status`：`production_ready` 表示低延迟本地观察和 provider 回填都可用；`realtime_ready_provider_pending` 表示标注可以实时落轴，但 provider 事件仍待接；`provider_ready_collect_local_evidence` 表示官方事件可用但还缺真实会议页 DOM 证据；`needs_live_dom_and_provider_evidence` 表示两条路径都还缺真实样本。这个状态比单独看 setup manifest 更接近上线决策。`platform-strategy` 是 rollout 的产品化读法，适合给宿主应用决定“现在能不能先开 pilot”；`platform-live-adapter` 的 `buildMeetingPlatformLiveAdapterMatrix()` 和 `createMeetingPlatformLiveAdapterSuite()` 会把 Google Meet / Teams / Zoom / Webex / Lark 收敛成一个宿主可直接消费的多平台接入面板和 adapter 工厂，`buildMeetingPlatformLiveAdapterReadiness()` / `buildMeetingPlatformLiveAdapterReadinessMatrix()` 则把 adapter 方法、`captured_at_ms` 时间戳约束、provider/transcript 非阻塞策略和 evidence package 复验压成单一上线验收对象。`buildMeetingPlatformAdaptationRunbook()` 会进一步给出现场采样步骤：安装 capture runtime、采 active speaker / ended DOM 快照、采 provider start/end 事件、运行本地 DOM gate、运行 provider event gate，最后合成 rollout plan。真实会议进行中，宿主项目优先使用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter`：`observeMeetingApp()` 负责本地观察建轴和本地证据记录，`ingestProvider()` 负责官方事件回填和 provider 证据记录，`insertAnnotation()` 内部会调用 `platform-realtime-annotation`，先做设备时间校准、会议身份绑定和 annotation intake，再按结果执行插入、建轴后插入或 pending。只需要证据不需要写时间轴时，可以退到 `@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session`：本地 DOM 快照和 provider webhook 到达时分别写入同一个 session，`summary()` 会实时返回 `can_insert_realtime_marks`、`provider_reconcile_ready`、correlation 状态和下一步动作，避免等会议结束后才知道证据不够。单场真实会议采样完成后，用 `live.exportPackage()`、`session.exportPackage()` 或 `@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package` 把 provider capture records、本地 DOM record set、rollout plan、handoff 摘要和脱敏环境 key 封成一个可复验交接包；这个包可以交给另一个项目继续接入，也可以作为 pilot/production gate 的审计证据。包内会带 `platform-evidence-correlation` 结果，用 meeting id / URL 或时间窗口证明 provider 事件和本地 DOM 记录来自同一场会议。收到包的一方应调用 `verifyMeetingPlatformEvidencePackage()` 或仓库命令 `npm run meeting-platform:evidence-package` 重新计算 rollout 和 correlation，不直接信任包内旧结论；如果要直接验收“SDK live adapter 是否可交给宿主项目上线”，用 `npm run meeting-platform:live-readiness` 生成 `meeting_platform_live_readiness_report`。仓库内的 `npm run meeting-platform:rollout-matrix` 会批量读取 `data/provider-evidence/`、`data/meeting-app-evidence/` 和 `data/meeting-platform-evidence-packages/`，适合做多平台验收看板。

给其他项目开始接入时，优先交付 `@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package`。它不是新的事件解析器，而是把 `platform-adapter-contract`、`platform-runtime-profile`、`platform-field-capture`、`meeting-app-extension`、`platform-live-adapter` readiness 和 `platform-handoff-readiness` 汇总成单个平台的机器可读接入包：包括 extension matches、provider transport/start/end events、`insertAnnotation` 写入契约、speaker marker 滤波参数、会后 transcript 非阻塞声明、证据文件路径、SDK import 和 CLI 命令。`buildMeetingPlatformAdaptationPackageMatrix()` 默认可覆盖 Google Meet、Microsoft Teams、Zoom、Webex、Lark；仓库命令 `npm run meeting-platform:adaptation-package` 会把每个平台 package 写到 `data/meeting-platform-adaptation-packages/`，并输出 `meeting_platform_adaptation_package_report`。这份 package 的 `readiness.sdk_wiring_ready=true` 只表示 SDK 接口和实时策略可交给宿主项目接入；production 结论仍以后续真实 evidence package / handoff readiness 为准。

当宿主项目已经准备写浏览器扩展、WebView preload 或 native host 时，再交付 `@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle`。它在 adaptation package 上继续展开运行时 glue：`meeting-app-content-script` manifest/start options、`meeting-app-browser-runtime` preset、URL matches/host permissions、extension message 示例、mutation observer 和 speaker filter 参数、host ingest endpoints。Google Meet、Teams、Zoom、Webex、Lark 的 bundle 都保留同一个实时原则：provider event 只做校准/回填，transcript 只做会后 artifact，实时标注必须由 `captured_at_ms` 直接落当前会议轴。仓库命令 `npm run meeting-platform:runtime-bundle` 会把这些 bundle 写到 `data/meeting-platform-runtime-bundles/`，供另一个工程直接消费。

下游开始选平台或做接入面板时，先读 `@ai-annotation/meeting-timeline-sdk/adapters/platform-registry` 的 `buildMeetingPlatformRegistryManifest()`。这张 manifest 汇总 normalizer、runtime bundle、provider security verifier、insert endpoint、`captured_at_ms`、provider/transcript 非阻塞状态和 host endpoints；`assertMeetingPlatformRegistryManifest()` 可作为 CI gate，确保这些平台都满足 SDK 接入契约。仓库命令 `npm run meeting-platform:registry` 会输出同一份报告。它适合放在配置页或 CI 里作为“当前 Google Meet / Teams / Zoom / Webex / Lark 是否具备 SDK 接入面”的第一层总览。

给其他项目真正写代码时，优先用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime`，而不是让接入方直接拼 `platform-kit`、live adapter、runtime bundle、registry 和 timeline view。`createMeetingPlatformIntegrationRuntime()` 暴露统一方法：`observeMeetingApp(platform, snapshot)`、`ingestProvider(platform, event)`、`insertAnnotation(platform, mark)`、`speakerTrack()`、`participantTrack()`、`timelineView()`、`runtimeBundles()`、`registry()` 和 `handoffReadiness()`。浏览器扩展或 WebView preload 用 `createMeetingPlatformIntegrationBrowserRuntime()`，它会从当前 `window.location`/DOM capture profile 自动识别 Google Meet、Teams、Zoom、Webex、Lark，并把 content-script sample、provider event 和标注消息路由到同一个 integration runtime。`buildMeetingPlatformIntegrationRuntimeManifest()` / `assertMeetingPlatformIntegrationRuntimeManifest()` 只验 SDK 接线、runtime bundle、`captured_at_ms`、provider/transcript 非阻塞策略，适合作为 host handoff gate；真实 DOM/provider 证据仍然必须走 `platform-real-intake` 或 `platform-handoff-readiness`，不能用 runtime manifest 代替 production ready。

如果下游需要一个可改造的 host scaffold，而不是自己拼这些模块，用 `platform-host-integration`。生成的 host wrapper 会同时暴露 `runtimeBundles()`、`handoffBundle()`、`adapterContracts()`、`insertAnnotation()` 和 provider webhook route；`/api/meeting-platform/runtime-bundles` 可作为扩展/WebView/native host 拉取 runtime bundle 的统一入口。

适配开发早期可以先跑 `@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-sample`。`runMeetingPlatformAdapterSample('google-meet')` 会用 fixture 级本地会议页快照、provider start/end/participant/transcript 事件和一条 `capturedAtMs` 标注，完整调用 live adapter 的 `observeMeetingApp()`、`insertAnnotation()`、`ingestProvider()`、`exportPackage()` 和 `verify()`，产出 `meeting_platform_adapter_sample`。`runMeetingPlatformAdapterSampleMatrix()` 默认覆盖 Google Meet、Microsoft Teams、Zoom、Webex、Lark；仓库命令 `npm run meeting-platform:adapter-sample` 会把每个平台 sample 写到 `data/meeting-platform-adapter-samples/`。这不是生产证据，只是 adapter 开发的第一道 smoke gate：证明 SDK 调用面、时间戳契约、provider 非阻塞策略和 evidence package 复验链路同时跑通。通过 sample 后，仍然需要真实 DOM snapshot 和真实 provider capture records 才能进入 `production_ready` 结论。

开始真实会议采样前，用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake` 生成现场接入计划。`buildMeetingPlatformFieldIntakePlan(platform, { baseUrl, env, evidenceDir })` 会把 provider connection、field capture manifest、collector config 和 real-intake gate 串成一个 `meeting_platform_field_intake_plan`：里面明确列出 provider endpoint、缺失安全环境变量、需要采的本地 observer 快照、provider start/end coverage、`captured_at_ms` 标注样本要求、输出文件路径、CLI 命令和 operator steps。`buildMeetingPlatformFieldIntakeMatrix()` 默认覆盖 Google Meet、Microsoft Teams、Zoom、Webex、Lark；仓库命令 `npm run meeting-platform:field-intake` 会写出 `data/meeting-platform-field-intake-plans/*.json` 和 `data/meeting-platform-field-intake-report.json`。这一步的定位是“采样前的执行单”，不是验收结论：如果它显示 `provider_setup_missing_env`，先补 provider 安全配置；如果显示 `needs_local_observer_capture` 或 `needs_provider_event_capture`，按计划采真实会议页和官方事件证据。

真实会议页快照采到后，先跑本地 observer 诊断，再进入平台 evidence package 验收。SDK 的 `buildMeetingAppDomAdaptationDiagnosisMatrix()` 会把 Google Meet、Teams、Zoom、Webex、Lark 的 selector 命中、`meeting_started`、`speaker_started`、`meeting_ended` 分成独立字段；仓库命令 `npm run meeting-app:evidence-matrix` 也会在每个平台 row 中输出 `dom_diagnosis`，用于快速判断是 controls/participants/active speaker 选择器没命中，还是 observer 状态转换没闭合。

如果产品侧只需要在时间轴上标出“谁在什么时候开始发言”，不要走转写链路。SDK 的 `@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track` 会把连续 active-speaker samples 或宿主已去抖的 `speaker_started/speaker_ended` signals 压成 `speaker_track` marks，并输出 `dropped_duplicate_count`、`dropped_short_segment_count`、`open_segment_count` 等滤波诊断。默认策略是候选发言人稳定后再落 marker、短发言段不写、同一发言人近距离重复合并；Google Meet / Teams / Zoom / Webex / Lark 都使用同一套本地观察优先策略，provider 事件不作为实时发言人来源，转写 speaker label 只允许会后回填。

参会人位置单独走 `@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track`。它接收本地 roster snapshots 或已归一化的 `participant_joined/participant_left` signals，输出 `participant_track` marks，并通过 `duplicate_window_ms`、`suppress_reconnect_gap_ms`、`leave_stable_ms` 过滤重复事件、短暂断线重连和未稳定离会。首帧 roster 默认只建立基线，不写“加入”marker；Google Meet participant.v2、Teams rosterUpdated/bot events、Zoom participant webhooks、Webex meetingParticipants、Lark join/leave 都只能作为参会人轨来源或校准证据，不能替代本地建轴，也不能阻塞电子纸实时标注。

跨项目复用时间轴 UI 时，不要复制 demo 里的 SVG 计算，改用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view`。它把会议轴、用户标注、speaker/participant/artifact/transcript 轨合成 renderer-agnostic view model：`rails` 定义轨道，`viewport` 定义当前窗口，`ticks` 定义刻度，`markers[*].x_ratio` 定义在窗口中的横向位置，`uncalibrated_markers` 明确列出缺少可靠时间的标注。该层只做本地数据整理和缩放窗口计算，不拉 provider、不等 transcript，也不参与建轴；Google Meet / Teams / Zoom / Webex / Lark 的 UI 都应该消费同一份 view model。

跨设备写入标注前，宿主应先调用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync` 或等价算法校准设备时间。该 adapter 使用 `/api/time` 的 midpoint 采样契约计算 `clock_offset_ms`，输出 `recommended_offset_ms`、RTT、uncertainty、是否超过 500ms 推荐 skew，以及已校准的 `calibrated_annotation.captured_at_ms`。这一步是平台无关的：Google Meet、Teams、Zoom、Webex、Lark 的 provider 事件都不能替代设备端笔迹结束时刻；provider 事件最多用于会议轴回填和审计。

跨平台绑定会议身份时，用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding` 把本地观察、provider signal、当前轴和 annotation 汇总成一个决策。它比较 `meeting_id`、`external_meeting_id`、URL 稳定 id、平台、标题和开始时间窗口，输出 `bound_to_current_axis`、`open_axis_from_local_observer`、`open_axis_from_provider_start`、`pending_binding`、`binding_conflict` 或 `insufficient_identity`。这层的作用是防止晚到 provider start 把用户正在写的当前会议轴切错，也让 Google Meet / Teams / Zoom / Webex / Lark 都使用同一套绑定证据和冲突策略。

跨项目插入实时标注时，不要让宿主项目复刻 demo 服务端的 pending/open-session/after-end 分支，改用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake`。这个 adapter 接收当前会议轴状态和一条标注，唯一强约束是标注必须带可靠 `captured_at_ms` 或绝对 stroke 时间戳；它会输出 `ready_to_insert_current_axis`、`start_open_session_then_insert`、`pending_real_meeting`、`needs_device_captured_at`、`after_meeting_end` 等状态，并附带 `insert_payload`、`open_session_payload` 和下一步动作。这里同样不拉 provider、不等 transcript，provider start/end 只用于后续回填或校准；设备端一旦写完笔迹，应把 `ink_end_at_ms`/`captured_at_ms` 随标注一起上报。

给其他项目接 SDK 时，优先暴露更高层的 `@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation`。它顺序组合 `platform-clock-sync`、`platform-session-binding` 和 `platform-annotation-intake`，把一条设备标注变成 `insert_mark`、`start_meeting_session + insert_mark`、`store_pending_mark`、`run_clock_sync_before_realtime_insert` 或 `do_not_insert_until_meeting_identity_conflict_is_resolved`。这层是“会议中边写边标到真实时间轴”的最小复用单元：宿主只需要执行返回的 action，不需要知道 demo 服务端内部如何处理当前轴、待绑定、晚到事件和冲突。

会后转写、录制和智能纪要统一走 `@ai-annotation/meeting-timeline-sdk/adapters/platform-artifact-handoff`。它把各平台的 `artifact_ready` signal 转成 fetch/import 请求骨架，并输出 transcript normalizer、token env、source event types 和 import endpoint。这个 handoff 的验收条件是“会后能补拉和导入”，不是“实时可用”：`realtime_blocking_count` 必须为 0，实时标注不能等待 transcript/recording/smart notes 生成，导入后的 transcript 只进入会后轨道或搜索索引，不允许重写用户当时的 `captured_at_ms` 标注位置。

进入真实平台验证时，用 `@ai-annotation/meeting-timeline-sdk/adapters/platform-real-intake` 做第二道 gate。`buildMeetingPlatformRealEvidenceIntakeReport(platform, input)` 会同时检查真实 provider capture records、真实会议 App DOM/native observer records、`captured_at_ms`、fixture/synthetic 证据污染、evidence package 复验和 live adapter readiness；`assertMeetingPlatformRealEvidenceIntakeMatrix()` 可一次覆盖 Google Meet、Microsoft Teams、Zoom、Webex、Lark。这个入口默认拒绝 fixture evidence，只有 provider 安全配置、真实会议页观察和证据包复验都通过时才返回 `accepted=true`，适合作为“可以交给宿主项目进入 pilot/production”的 SDK 级验收对象。仓库命令 `npm run meeting-platform:real-intake` 会读取 `data/provider-evidence/`、`data/meeting-app-evidence/` 和 `data/meeting-platform-evidence-packages/`，输出 `data/meeting-platform-real-intake-report.json`；现场采样时也可以显式传 `--provider-dir`、`--dom-dir`、`--package-dir`、`--platforms` 和 `--fail-on-incomplete=true`。

给其他项目交付接入能力时，最后再跑 `@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness`。它不是新的事件解析器，而是把 `platform-adapter-contract`、`platform-provider-connection`、`meeting-app-profile` DOM 诊断、`platform-field-intake`、`platform-real-intake` 和 `platform-live-adapter` readiness 汇总成一个交付状态。`buildMeetingPlatformHandoffReadinessMatrix()` 默认覆盖 Google Meet、Microsoft Teams、Zoom、Webex、Lark，每行输出 `handoff_ready`、`pilot_ready`、`production_ready`、`local_observer_ready`、`provider_reconcile_ready`、缺失 provider env、DOM/provider 证据数量和下一步动作。状态含义是：`needs_local_observer_evidence` 说明还不能把实时标注交给宿主；`pilot_ready_provider_setup_pending` 说明本地 observer 可以先试点但 provider 安全配置未完成；`pilot_ready_provider_reconcile_pending` 说明可以试点并继续采 provider 回填证据；`production_ready` 说明本地观察、provider start/end、证据包复验和 live adapter readiness 都通过。仓库命令 `npm run meeting-platform:handoff-readiness` 会读取 `data/meeting-platform-evidence-packages/`，输出 `data/meeting-platform-handoff-readiness-report.json`，适合作为下游项目接入会议 timeline SDK 前的单一验收入口。

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
    platform-integration-runtime.mjs
    platform-integration-runtime.d.ts
    platform-ingest.mjs
    platform-ingest.d.ts
    platform-rollout.mjs
    platform-rollout.d.ts
    platform-evidence-correlation.mjs
    platform-evidence-correlation.d.ts
    platform-evidence-session.mjs
    platform-evidence-session.d.ts
    platform-live-adapter.mjs
    platform-live-adapter.d.ts
    platform-adapter-contract.mjs
    platform-adapter-contract.d.ts
    platform-adapter-sample.mjs
    platform-adapter-sample.d.ts
    platform-field-intake.mjs
    platform-field-intake.d.ts
    platform-real-intake.mjs
    platform-real-intake.d.ts
    platform-host-integration.mjs
    platform-host-integration.d.ts
    platform-evidence-package.mjs
    platform-evidence-package.d.ts
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
