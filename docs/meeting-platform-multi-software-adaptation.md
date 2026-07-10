# 会议软件适配快照（Google Meet、Teams、Zoom、Webex、Lark）

你问的核心是：`SDK`方向下，Google Meet 等会议软件怎么真正落地进时间轴能力。结论先说在前面——**架构层面已经具备“统一路线”**，现在差的是把每个平台的真实采样证据补齐后放到生产上线。

## 统一结论

- 5 个目标平台都应走同一条主线：**本地观察（低延迟）优先 + provider 回填（非阻塞）**。
- 5 个平台的 `platform-adapter-route` 当前都可给出 `local_observer_first_provider_reconcile` 路线。
- `meeting_platform:crosswalk/strategy/adapter-route/conformance` 都返回“可接入/可复用”，但都标记为 `needs_live_dom_evidence`，说明**现场采样证据**和**host 绑定**还没全部完成。
- 重点：
  - 不要把官方事件当实时主时钟。
  - `meeting_started / meeting_ended` 要尽量由本地候选观察/本地检测器触发。
  - provider 事件用于“回填、修正、合规”，只做同步/对账，不阻塞实时插标。

## 当前平台快照（从命令结果提炼）

### 一眼看懂表

| 平台 | 主采样面 | Provider 主线 | 实时策略 | 主线状态 | 下一步首要动作 |
| --- | --- | --- | --- | --- | --- |
| Google Meet | `browser_extension` | `google_workspace_events_pubsub`（Workspace Events） | `local_observer_first_provider_reconcile` | `ready_for_pilot` / `realtime:false` | 采集真实 meeting page live DOM 与 active speaker 快照 |
| Microsoft Teams | `desktop_or_browser_observer` | `microsoft_graph_change_notifications` | `local_observer_first_provider_reconcile` | `ready_for_pilot` / `realtime:false` | 采集会议开始/结束真实 DOM 与候选窗口快照 |
| Zoom | `native_detector` | `zoom_meeting_webhooks` | `local_observer_first_provider_reconcile` | `ready_for_pilot` / `realtime:false` | 采集本地加入/离会快照、主时轴验证 |
| Webex | `browser_extension_or_native_detector` | `webex_webhooks` | `local_observer_first_provider_reconcile` | `ready_for_pilot` / `realtime:false` | 补齐本地DOM活动窗口样本与 provider start/end 样本 |
| Lark / 飞书 | `browser_extension_or_desktop_observer` | `lark_long_connection_or_event_callback` | `local_observer_first_provider_reconcile` | `ready_for_pilot` / `realtime:false` | 除本地候选观察外，还需补 provider 正式回填样本 |

说明：`realtime:false` 并非“不能用”，是因为当前未达到 `生产-ready` 的证据链，不影响先跑 P0 的实时写入方案。

## 各平台适配点（开发顺序）

- **Google Meet**
  - 路线：`content script / browser_extension -> local_observer -> axis`，provider 作为回填。
  - 必要字段：会议开始、结束事件，会议 URL/主题，会议 ID，`publish_time`/`event_time` 时间归一。
  - 难点：Workspace 订阅作用域、Pub/Sub 鉴权、会议信息有时依赖 REST 补全。

- **Microsoft Teams**
  - 路线：桌面/Web 统一走 `local_observer`，provider 走 Graph `change notifications`。
  - 必要字段：`callStarted/callEnded`、`rosterUpdated`、`clientState` / 订阅生命周期。
  - 难点：租户权限和订阅续期、`subscriptionRemoved/reauthorizationRequired/missed` 处理。

- **Zoom**
  - 路线：桌面优先（native detector），provider 走 Meeting webhooks。
  - 必要字段：`meeting.started` / `meeting.ended` / `participant_joined|left` / `recording.completed`。
  - 难点：webhook 快速响应、签名鉴权、转写依赖 cloud recording。

- **Webex**
  - 路线：浏览器/原生观察器均可，provider 走 meetings/participants/webhooks。
  - 必要字段：`meetings.started|ended`、`meetingParticipants.joined|left`、`recordings / meetingTranscripts` 产物 URL。
  - 难点：部分 payload 需要 REST 补齐；签名与 replay 校验稳定性。

- **Lark / 飞书**
  - 路线：现网可走长连接事件 + 本地观察器；provider 回填与分钟稿回灌。
  - 必要字段：`meeting_started / meeting_ended`、`join_meeting / leave_meeting`、minutes/transcript export token。
  - 难点：事件订阅依赖租户授权、部分分钟内容是会后。

## 建议的实施顺序

现场操作、量化阈值和失败归因统一以 [多会议平台 P0 现场验收单](meeting-platform-p0-field-acceptance.md) 为准。

### P0（你现在可以发起）
1. 统一接入 `local_observer_axis`（所有平台）。
2. 统一标注上报：每个 mark 都带 `captured_at_ms`。
3. 优先把 `meeting-platform:strategy` 与 `meeting-platform:runtime-bundle` 通过 CI 验证通过。

### P1（同一周内补齐）
1. 对 Google Meet + Teams + Webex + Zoom 分别补本地 DOM 快照采样：`active_meeting / active_speaker / end_meeting`。
2. 配置 provider 环境变量与签名校验。
3. 运行 `meeting-platform:field-intake` + `meeting-platform:handoff-readiness`。

### P2（生产前）
1. 补会议 artifact 的回填（转写/录制）并验证 `transcript import`。
2. 输出 `meeting-platform-evidence-package`，跑 `meeting-platform:live-readiness --require-production-ready=false` 作为 pilot。
3. 与 SDK 消费端对接时固定 `captured_at_ms` 时钟策略，禁止等待 provider 决定是否落标。

## 建议你直接执行的命令

```bash
npm run meeting-platform:strategy -- --platforms=google-meet,teams,zoom,webex,lark
npm run meeting-platform:adapter-route -- --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787
npm run meeting-platform:crosswalk -- --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787 --json=true
npm run meeting-platform:conformance -- --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787
npm run meeting-platform:field-intake -- --platforms=google-meet,teams,zoom,webex,lark --base-url=http://localhost:8787
npm run meeting-platform:handoff-readiness -- --platforms=google-meet,teams,zoom,webex,lark
```

## 当前关键判断（给你下项目决定）
- 如果你的下一步目标是“先上线可用”，不要先追求“完美 provider 事件”。
  先把每个平台 `local_observer_axis` 跑起来，让标注稳定对时。
- 如果目标是“后续能横向扩展到更多会议软件”，下一项动作是把 Google Meet 的 adapter profile 固化为模板，按同样 `surface_order + route` 方式新增平台即可。
