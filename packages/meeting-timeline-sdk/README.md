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

提交给外部项目接入前，建议在仓库根目录跑一次包级 smoke：

```bash
npm run sdk:package-smoke
```

这个检查会先 `npm pack`，再创建一个临时 consumer 项目，通过 `@ai-annotation/meeting-timeline-sdk` 和 `@ai-annotation/meeting-timeline-sdk/adapters/*` 导入公开入口，避免 SDK 只在 monorepo 相对路径下可用。

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
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-session-discovery`
- `@ai-annotation/meeting-timeline-sdk/adapters/active-speaker`
- `@ai-annotation/meeting-timeline-sdk/adapters/browser-meeting`
- `@ai-annotation/meeting-timeline-sdk/adapters/native-meeting`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-monitor`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-extension`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-capability`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-integration-package`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-pipeline`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate`
- `@ai-annotation/meeting-timeline-sdk/adapters/meeting-source`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-registry`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest`
- `@ai-annotation/meeting-timeline-sdk/adapters/timeline-bridge`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-kit`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event`
- `@ai-annotation/meeting-timeline-sdk/adapters/signal-reconciler`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-handler`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-webhook-router`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-http`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-node`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-acceptance`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-capture`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-gate`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-correlation`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-host-integration`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-subscription-handoff`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-artifact-handoff`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract`
- `@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package`
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

如果接入方不是只验证单条链路，而是要把 Google Meet、Teams、Zoom、Webex、Lark 做成同一个会议标注 runtime，优先用 `platform-integration-runtime`。它是薄 facade：内部仍使用 `platform-kit`、live adapter、runtime bundle、registry 和 handoff readiness，但调用方面只需要面向统一方法。

```js
import {
  createMeetingPlatformIntegrationRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';

const runtime = createMeetingPlatformIntegrationRuntime({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom'],
});

const resolution = runtime.resolvePlatform({
  url: 'https://meet.google.com/abc-defg-hij', // also accepts zoommtg://, msteams://, lark://, webex://
  title: 'Google Meet',
});
// resolution.platform === 'google_meet'; resolution.strategy.primary_axis_source === 'local_observer'

const candidateResolution = runtime.resolvePlatformCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});
// candidateResolution.selected_resolution.platform === 'google_meet'

await runtime.observePlatformCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});
// observes the selected supported platform and emits meeting_started / meeting_ended signals.

await runtime.observeMeetingApp('google-meet', domSnapshot);
await runtime.ingestProvider('google-meet', workspaceEvent);
await runtime.insertAnnotation('google-meet', {
  annotation: {
    label: 'why?',
    captured_at_ms: Date.now(),
  },
});

const viewModel = runtime.timelineView('google-meet', {
  meeting: currentMeeting,
  annotations: currentMarks,
});
```

浏览器扩展或 WebView preload 可以再包一层 `createMeetingPlatformIntegrationBrowserRuntime()`。它会从当前 `window.location`/DOM capture profile 自动识别 Google Meet、Teams、Zoom、Webex、Lark，然后把 content-script sample、provider event 和手写标注路由到同一个 integration runtime：

```js
import {
  createMeetingPlatformIntegrationBrowserRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';

const browserRuntime = createMeetingPlatformIntegrationBrowserRuntime({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

const currentPlatform = browserRuntime.resolvePlatform();

await browserRuntime.sample();
await browserRuntime.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: {
      label: 'why?',
      captured_at_ms: Date.now(),
    },
  },
});
```

如果注入点就是浏览器扩展 content script，可以直接安装平台版 bridge。它会创建 browser runtime、监听 extension message，并保留平台检测状态，background 或 native host 只需要发送统一的 `meeting_timeline.insert_mark` / `meeting_timeline.sample` 消息：

```js
import {
  installMeetingPlatformIntegrationContentScriptBridge,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-integration-runtime';

installMeetingPlatformIntegrationContentScriptBridge({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
```

`runtime.manifest()` 只证明 SDK 接线、runtime bundle、adapter route、候选会议观察、speaker/participant 位置轨、`captured_at_ms`、provider/transcript 非阻塞策略已经满足 host handoff；真实会议页 DOM 和官方事件证据仍然要用 `platform-real-intake` / `platform-handoff-readiness` 验收，不能用静态 runtime manifest 冒充 production ready。正式交付给另一个项目时用 `await runtime.runManifest({ requireHandoffReady: true, ...evidenceByPlatform })` 或 `runMeetingPlatformIntegrationRuntimeManifest()`：它会运行 handoff readiness 和 runtime host replay，只有每个平台的 `runtime_host_replay_accepted=true` 且 `handoff_ready=true` 时才会让 `host_integration_ready=true`。manifest 会带上 `adapter_route_matrix`、`speaker_track_matrix`、`participant_track_matrix` 和 `handoff_readiness_matrix`，并在任一平台的适配路线、发言人轨、参会人轨或 runtime replay 不满足交付要求时直接报错。

跨项目投递到 host 的统一 HTTP envelope 用 `platform-runtime-event`。Google Meet 扩展、Teams WebView preload、Zoom native helper 都可以只构造同一类事件包，再发到 `/api/meeting-platform/runtime-events`；host 侧 `handleRuntimeEvent()` 会分发到 observe、candidate observation、provider ingest、insert annotation、speaker/participant track、timeline view、adapter route/blueprint inspection 或可执行 handoff gate：

```js
import {
  buildMeetingPlatformRuntimeEventPlan,
  createMeetingPlatformRuntimeEventClient,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-event';

const googleRuntimePlan = buildMeetingPlatformRuntimeEventPlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
// googleRuntimePlan.actions 明确列出 observe/candidate-observe/provider/annotation/speaker/participant/view/adapter-route/adapter-blueprint/run gate 各 action 的 producer、必填字段和 client method。
// googleRuntimePlan.realtime_contract.provider_events_required_for_realtime === false。
// googleRuntimePlan.realtime_contract.transcript_required_for_realtime === false。

const runtimeEvents = createMeetingPlatformRuntimeEventClient({
  baseUrl: 'https://timeline.example.com',
});

await runtimeEvents.observePlatformCandidates({
  windows: [{
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet', active: true }],
  }],
});

await runtimeEvents.insertAnnotation('google-meet', {
  annotation: {
    label: 'why?',
    captured_at_ms: Date.now(),
  },
  current_meeting: currentMeeting,
});

const blueprints = await runtimeEvents.adapterBlueprints({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

await runtimeEvents.runManifest({ requireHandoffReady: true, google_meet: { evidencePackage } });
await runtimeEvents.runHandoffReadiness({ platforms: ['google-meet'], target: 'production' });
```

如果使用 SDK 生成的浏览器扩展脚手架，background worker 已内置同一条链路：向扩展发送 `meeting_timeline.observe_candidates`，它会查询当前浏览器 tabs，构造成 `observe_platform_candidates` runtime event 投递给 host。这样 Google Meet、Teams、Zoom、Webex、Lark/飞书都可以先用同一种“候选会议窗口观察”方式建轴，content script 只负责更细粒度的页面内观察。

如果要给另一个项目批量交付动作契约，可以直接导出 runtime event plan：

```bash
npm run meeting-platform:runtime-event-plan -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-runtime-event-plans \
  --report-file=data/meeting-platform-runtime-event-plan-report.json
```

多平台正式接入前，可以先用 `platform-strategy` 输出机器可读策略。它把 Google Meet、Teams、Zoom、Webex、Lark 的共性收敛成同一条原则：实时标注轴由本地观察或 host detector 先建，provider webhook 只做 reconcile/backfill，post-meeting transcript 只做会后导入，不阻塞当前标注：

```js
import {
  buildMeetingPlatformAdaptationStrategyMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-strategy';

const matrix = buildMeetingPlatformAdaptationStrategyMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex'],
});

console.log(matrix.rows);
console.log(matrix.rows.map((row) => ({
  platform: row.platform,
  next: row.next_phase,
  providerPath: row.provider_path,
  permissionRisk: row.permission_risk,
})));

const googleStrategy = matrix.strategies.find((item) => item.platform === 'google_meet');
console.log(googleStrategy.adaptation_playbook.phases.map((phase) => ({
  id: phase.id,
  priority: phase.priority,
  status: phase.status,
})));
```

`adaptation_playbook` 是跨会议软件适配时最应该交给宿主项目看的字段：

- `phases`：P0 本地轴、P0 实时标注写入、P1 provider reconcile、P1 发言人/参会者 marker、P2 会后 transcript/recording 回填、P3 production evidence gate。
- `integration_path`：Google Meet 是 `google_workspace_events_pubsub`，Teams 是 `microsoft_graph_change_notifications`，Zoom/Webex 是 webhook，Lark 是长连接或事件回调。
- `risk_profile`：权限风险、事件延迟风险、发言人实时缺口、会后转写可用性，以及对应 mitigation。
- `matrix.rows`：给接入面板用的扁平字段，包括 `next_phase`、`provider_path`、`permission_risk`、`provider_reconcile_required`、`speaker_realtime_gap`、`post_meeting_backfill_supported`。

如果另一个项目只想知道“这个会议软件应该接哪条链路”，优先用 `platform-adapter-route`。它把 strategy/runtime profile 收敛成宿主可直接消费的 route：第一优先级永远是本地观察或 host detector 建当前会议轴，第二步用 `captured_at_ms` 插入标注，发言人只写位置 marker，provider webhook 做非阻塞校准，会后 transcript/recording 只做回填：

```js
import {
  buildMeetingPlatformAdapterRoute,
  buildMeetingPlatformAdapterRouteMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-route';

const google = buildMeetingPlatformAdapterRoute('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

console.log(google.route_order);
console.log(google.adapter_surfaces.primary);
console.log(google.entrypoints.browser_extension.matches);

const routeMatrix = buildMeetingPlatformAdapterRouteMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
```

`platform-adapter-route`、`platform-adapter-contract` 和 `platform-adaptation-package` 都会透出同一份 surface 指南：`adapter_surfaces`、`launch_requirements`、`evidence_thresholds`、`fallback_policy`。下游项目不必自己判断 Google Meet/Teams/Zoom/Webex/Lark 的优先入口；例如 Google Meet 会优先 `browser_extension`，Zoom 会优先 `native_detector`，provider webhook 只作为非阻塞 reconcile 和生产证据。

如果要把 SDK 给另一个项目做真实接入，建议优先交付 `platform-adapter-blueprint`。它比 route 更靠近“接入契约”：同时包含 browser extension / native detector / provider reconcile / post-meeting artifact 四类 surface 的候选输入字段、实时建轴证据、发言人 marker 约束、`captured_at_ms` 标注写入要求、验收 gate 和平台风险说明。Google Meet 会明确是浏览器优先，Teams/Zoom 会明确 native detector 的优先级，所有平台都会声明 provider 和 transcript 不阻塞实时标注：

```js
import {
  buildMeetingPlatformAdapterBlueprint,
  buildMeetingPlatformAdapterBlueprintMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-blueprint';

const meetBlueprint = buildMeetingPlatformAdapterBlueprint('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

console.log(meetBlueprint.surfaces.browser_extension.evidence.required_fields);
console.log(meetBlueprint.surfaces.provider_reconcile.blocks_realtime); // false

const blueprintMatrix = buildMeetingPlatformAdapterBlueprintMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// 也可以通过 kit / SDK facade 调用：
// kit.platformAdapterBlueprint('zoom')
// sdk.adapterBlueprintMatrix()
```

批量交付时再用 `platform-registry` 做最后一层静态验收。Registry manifest 现在会统计 `adapter_blueprint_ready_count`，并要求每个平台都有 ready 的 blueprint、主 surface、SDK import、host endpoint，且 blueprint 里 provider/transcript 都不能阻塞实时标注，时间戳字段必须保持 `captured_at_ms`。这样下游项目接 Google Meet/Teams/Zoom/Webex/Lark 时，不会只拿到事件 normalizer，却缺少真实 runtime 接入契约。

也可以直接导出给另一个项目消费：

```bash
npm run meeting-platform:adapter-blueprint -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-adapter-blueprints \
  --report-file=data/meeting-platform-adapter-blueprint-report.json
```

也可以直接导出策略报告，给别的项目做平台选择或接入面板：

```bash
npm run meeting-platform:strategy -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --report-file=data/meeting-platform-strategy-report.json
```

宿主项目需要真正落地时，推荐再读一层 `platform-runtime-profile`。它把策略转成运行时可执行约束：谁先建轴、结束事件如何兜底、provider 事件是否阻塞实时标注、发言人位置用什么滤波参数。这个 profile 不依赖实时转写；发言人只作为时间轴 marker 写入，正文仍然等会后 transcript import：

```js
import {
  buildMeetingPlatformRuntimeProfileMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-profile';

const runtime = buildMeetingPlatformRuntimeProfileMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// 每个平台都要求：标注用 captured_at_ms；provider webhook 和 transcript 不阻塞实时落轴。
console.log(runtime.rows.map((row) => ({
  platform: row.platform,
  surface: row.primary_surface,
  start: row.start_create_on,
  end: row.end_create_on,
  speakerStableMs: row.speaker_min_stable_ms,
})));
```

也可以从 `platform-kit` 读取同一份 profile：

```js
const googleProfile = kit.platformRuntimeProfile('google-meet');
// googleProfile.axis.end.fallbacks 描述 Workspace Events 晚到时，本地观察器如何先闭合会议轴。
// googleProfile.speaker_markers.filter 可直接传给 active-speaker observer。
// googleProfile.adapter_surfaces 描述推荐接入面：Google Meet 优先 browser_extension，Zoom 优先 native_detector。
// googleProfile.launch_requirements/evidence_thresholds/fallback_policy 可直接用于接入 UI、preflight gate 和采样验收。
```

采真实 Google Meet / Teams / Zoom / Webex 样本时，用 `platform-field-capture` 生成采样清单。它面向现场工具，而不是最终验收：告诉采集器至少要采 active speaker DOM、meeting ended DOM、provider start/end 事件，以及这些数据最后应该导出到哪个 evidence package：

```js
import {
  buildMeetingPlatformFieldCapturePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';

const plan = buildMeetingPlatformFieldCapturePlan('zoom', {
  baseUrl: 'https://timeline.example.com',
});

// plan.checklist 可直接渲染为采样 UI；plan.missing_items 表示当前 evidence package 还缺什么。
// 如果传入 evidencePackage，SDK 会返回 production_ready / pilot_ready_provider_pending 等状态。
```

如果现场采样工具已经拿到了 DOM snapshot、provider webhook 记录或 provider sample，可以直接用一站式 bundle 把原始采样输入转换成可复验的 evidence package，并同步得到采样缺口：

```js
import {
  buildMeetingPlatformFieldEvidenceBundle,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';

const bundle = buildMeetingPlatformFieldEvidenceBundle('google-meet', {
  meetingAppRecordSet,
  providerRecords,
}, {
  baseUrl: 'https://timeline.example.com',
});

// bundle.evidence_package 可落盘交给 CI 或下游项目。
// bundle.evidence_package.adapter_route 会说明本平台的实时接入路线：
// 本地观察先建轴、标注按 captured_at_ms 写入、provider/transcript 只做非阻塞回填。
// bundle.field_capture_plan.missing_items 表示还缺哪些真实会议样本。
// bundle.verification.passed 表示是否已经达到 production-ready 验收。
```

在宿主项目里也可以通过 `kit.platformFieldEvidenceBundle()` 和 `kit.platformFieldEvidenceMatrix()` 调用同一套逻辑，用于批量比较 Google Meet / Teams / Zoom / Webex 的真实采样进度。

现场第一次接入某个会议软件时，建议先生成 field intake plan。它把 provider connection、field capture manifest、collector config 和 real-intake gate 串成一份执行单：哪些环境变量缺失、provider endpoint 是什么、需要采哪些本地 observer 快照、provider start/end coverage 要求是什么、原始 JSON 和 evidence package 应写到哪里、下一步跑哪个命令：

```js
import {
  buildMeetingPlatformFieldIntakeMatrix,
  buildMeetingPlatformFieldIntakePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-intake';

const googleIntake = buildMeetingPlatformFieldIntakePlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  evidenceDir: 'data',
});

const matrix = buildMeetingPlatformFieldIntakeMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// googleIntake.operator_steps 可直接渲染成现场采样流程。
// matrix.rows 可用于多会议软件接入看板。
```

现场采样工具如果需要一份机器可读的接入说明，先生成 manifest。它会把“要采哪些 provider 事件 / DOM 快照、原始 JSON 允许哪些形态、输出文件写到哪里、用哪个 CLI 验收”放在同一个对象里：

```js
import {
  buildMeetingPlatformFieldCaptureManifest,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';

const manifest = buildMeetingPlatformFieldCaptureManifest('teams', {
  baseUrl: 'https://timeline.example.com',
});

// manifest.input_contract.accepted_inputs 是现场工具允许输出的 JSON 形态。
// manifest.file_contract.files 给出 raw input、bundle、evidence package 的默认路径。
// manifest.automation.commands.build_field_evidence 可直接交给 CI 或现场采样脚本执行。
```

采样端运行时更适合读取 collector config。它在 manifest 基础上补了浏览器 URL match patterns、content script 配置、timeline ingest endpoint 和实时标注策略：

```js
import {
  buildMeetingPlatformFieldCollectorConfig,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-field-capture';

const collector = buildMeetingPlatformFieldCollectorConfig('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

// collector.browser_observer.matches 可直接用于 Chrome extension / WebView preload 的白名单。
// collector.timeline_ingest.endpoints.insertMark 是实时标注写入 endpoint。
// collector.storage.files.field_evidence_input 是采样端应写入的 raw JSON 路径。
```

如果下游项目只想拿“一个平台如何接入会议时间轴”的最终契约，直接用 `platform-adapter-contract`。它会把 runtime profile、provider connection、collector config、候选会议观察和证据验收条件收敛成一个对象：

```js
import {
  assertMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContract,
  buildMeetingPlatformAdapterContractAcceptanceMatrix,
  buildMeetingPlatformAdapterContractMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-contract';

const google = buildMeetingPlatformAdapterContract('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

// google.realtime_axis.rules 明确：本地观察先建轴，provider 事件只做 reconcile。
// google.annotations.endpoints.insertMark 是实时标注写入 endpoint。
// google.annotations.endpoints.runtimeEvents 是 observe/provider/annotation 的统一 runtime event endpoint。
// google.annotations.runtime_event.client_factory 指向 createMeetingPlatformRuntimeEventClient。
// google.candidate_observation 定义 meeting_timeline.observe_candidates -> observe_platform_candidates -> /api/meeting-platform/observe-candidates。
// google.provider_observer.events 列出 Google Workspace Events 需要监听的 started/ended/participant/artifact 事件。
// google.local_observer.matches 可交给浏览器扩展或 WebView preload 白名单。
// google.evidence.missing_items 表示当前离 production-ready 还缺哪些真实会议样本。

const matrix = buildMeetingPlatformAdapterContractMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

console.log(matrix.rows);

const acceptance = buildMeetingPlatformAdapterContractAcceptanceMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// 默认 target=contract，只检查结构和实时策略：captured_at_ms、provider 不阻塞实时、
// transcript 不阻塞实时、insertMark endpoint、候选观察、浏览器匹配、provider start/end 事件等。
console.log(acceptance.accepted_count);

assertMeetingPlatformAdapterContract(google);
```

`platform-kit` 也暴露同一能力：`kit.platformAdapterContract('teams')` 和 `kit.platformAdapterContractMatrix()`。外部项目如果要做多会议软件适配面板，优先读这个 contract；只有实际采样、验收、导出 evidence package 时才下钻到 `platform-field-capture`。

如果要把这份契约交给另一个项目，不需要写 glue code，直接导出 JSON：

```sh
npm run meeting-platform:adapter-contract -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --out-dir=data/meeting-platform-adapter-contracts \
  --report-file=data/meeting-platform-adapter-contract-report.json
```

每个平台会生成一份 `{platform}.json`，例如 `google_meet.json`。其中 `annotations.endpoints` 给实时标注写入地址，`candidate_observation` 给 background/native host 的多窗口候选观察契约，`local_observer` 给浏览器扩展或 native host 的页面内观察配置，`provider_observer` 给官方事件订阅/校准配置，`evidence.missing_items` 表示离 production-ready 还缺哪些真实会议样本。

CLI 报告会同时输出 `acceptance`。默认 `--acceptance-target=contract` 只检查 contract 是否能被宿主项目安全接入；如果要把真实证据也纳入 gate，可以用 `--acceptance-target=production --fail-on-rejected=true`，此时缺真实 DOM / provider start-end 样本的平台会失败。

如果宿主项目需要“直接接入包”而不是单独拼 contract、runtime profile、extension matches、provider setup、runtime event plan 和 live readiness，用 `platform-adaptation-package`。它把 Google Meet / Teams / Zoom / Webex / Lark 的本地观察、provider 回填、实时标注、runtime event 动作表、发言人 marker、会后转写、证据路径、SDK import 和命令行验收压成一个机器可读对象：

```js
import {
  buildMeetingPlatformAdaptationPackage,
  buildMeetingPlatformAdaptationPackageMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adaptation-package';

const googlePackage = buildMeetingPlatformAdaptationPackage('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

// googlePackage.extension.matches 可直接给浏览器扩展 / WebView preload。
// googlePackage.extension.permissions 与 googlePackage.candidate_observation 可直接给 background/native host 生成候选会议观察器。
// googlePackage.annotation_pipeline.insert_endpoint 是设备端实时标注写入地址。
// googlePackage.annotation_pipeline.runtime_event_plan 是 observe/provider/annotation/speaker/view/run gate 的动作契约。
// googlePackage.runtime_event_plan.examples.insert_annotation 是外部项目插入标注的样例 envelope。
// googlePackage.adaptation_playbook.next_phase 是接入面板下一步应该做的阶段。
// googlePackage.adaptation_playbook.integration_path.path 是 provider 侧接法，例如 google_workspace_events_pubsub。
// googlePackage.adaptation_playbook.risk_profile 描述权限、延迟、发言人 marker 和会后回填风险。
// googlePackage.provider_observer.required_for_realtime === false。
// googlePackage.transcript.blocks_realtime_annotation === false。

const packageMatrix = buildMeetingPlatformAdaptationPackageMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// packageMatrix.candidate_observer_count 表示多少平台已经暴露 observe_platform_candidates 候选观察契约。
console.log(packageMatrix.rows.map((row) => ({
  platform: row.platform,
  next: row.next_phase,
  provider: row.provider_path,
  backfill: row.post_meeting_backfill_supported,
})));
```

`platform-kit` 也暴露同一层：`kit.platformAdaptationPackage('google-meet')` 和 `kit.platformAdaptationPackageMatrix()`。CLI 可直接导出每个平台的 package JSON：

```sh
npm run meeting-platform:adaptation-package -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --out-dir=data/meeting-platform-adaptation-packages \
  --report-file=data/meeting-platform-adaptation-package-report.json
```

这份 package 的定位是“交给另一个项目开始接入”的 SDK 汇总，不替代真实会议采样；`readiness.sdk_wiring_ready=true` 只说明协议和 SDK 调用面可接，是否能 production 仍要看 evidence package / handoff readiness。

如果宿主项目只关心会议软件页面侧适配，可以直接用 `meeting-app-adapter-integration-package`。它把静态 handoff package、capability report、execution plan、runtime delivery、entrypoints、命令和 evidence contract 合成一个对象，适合作为 Google Meet、Teams、Zoom、Webex、Lark 适配任务的交接输入：

```js
import {
  buildMeetingAppAdapterIntegrationPackage,
  buildMeetingAppAdapterIntegrationPackageMatrix,
  createMeetingPlatformTimelineKit,
  detectMeetingPlatformForBrowser,
} from '@ai-annotation/meeting-timeline-sdk';

const pkg = buildMeetingAppAdapterIntegrationPackage('google-meet', {
  input: liveDomSnapshot,
  evidence: liveEvidence,
});
const matrix = buildMeetingAppAdapterIntegrationPackageMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  inputs: liveDomSnapshotsByPlatform,
  evidenceByPlatform,
});

console.log(pkg.runtime_delivery.adapter_route, pkg.entrypoints, pkg.integration_steps, matrix.rows);
```

SDK 主入口会直接暴露 `platform-kit`、`meeting-app-adapter-integration-package`、`platform-integration-runtime` 和 `platform-runtime-event` 这几层；需要极细粒度 tree-shaking 时，仍可以从 `@ai-annotation/meeting-timeline-sdk/adapters/*` subpath 导入。

宿主项目也可以用更高层的 `createMeetingAppTimelineSdk()`：

```js
import { createMeetingAppTimelineSdk } from '@ai-annotation/meeting-timeline-sdk';

const meetingSdk = createMeetingAppTimelineSdk({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

const detected = meetingSdk.detect({ url: location.href, title: document.title });
await meetingSdk.observeMeetingApp(detected.platform, liveDomSnapshot, { remote: true });
await meetingSdk.insertAnnotation(detected.platform, mark, { remote: true });

const handoff = meetingSdk.handoff({ url: location.href, title: document.title }, {
  surface: 'browser-extension',
});
const hostPackage = meetingSdk.hostPackage({
  surfaces: ['browser-extension', 'native-detector'],
});
const connectorPackage = meetingSdk.connectorPackage({
  surfaces: ['browser-extension', 'native-detector'],
  observeTracks: true,
});
const connectorReleaseGate = meetingSdk.connectorReleaseGate(connectorPackage);
const connectorPlatformRoadmap = meetingSdk.connectorPlatformRoadmap(connectorPackage);
const connectorAdapterMatrix = meetingSdk.connectorAdapterMatrix(connectorPackage);
const capability = meetingSdk.meetingAppAdapterCapability('google-meet', liveDomSnapshot);
const executionPlan = meetingSdk.meetingAppAdapterExecutionPlan(capability);
const capabilityMatrix = meetingSdk.meetingAppAdapterCapabilityMatrix();

const adaptationPackage = meetingSdk.platformAdaptationPackage('google-meet');
const consumerHandoff = meetingSdk.platformConsumerHandoff();
const implementationHandoff = meetingSdk.platformImplementationHandoff('google-meet');
const adapterExportPackage = meetingSdk.platformAdapterExportPackage('google-meet');
const adapterImportPlan = meetingSdk.platformAdapterImportPlan(adapterExportPackage, {
  availableFiles: adapterExportPackage.host_files.map((file) => file.path),
});
const adapterInstallManifest = meetingSdk.platformAdapterInstallManifest([adapterImportPlan]);
const adapterLaunchPlan = meetingSdk.platformAdapterLaunchPlan(adapterInstallManifest, {
  url: 'https://meet.google.com/abc-defg-hij',
});
const adapterSession = meetingSdk.platformAdapterSession(adapterLaunchPlan);
const adapterRunner = meetingSdk.platformAdapterRunner(adapterInstallManifest);
const adapterMessageBridge = meetingSdk.platformAdapterMessageBridge(adapterInstallManifest);
const adapterSmoke = await meetingSdk.platformAdapterSmoke(adapterInstallManifest, {
  platforms: ['google-meet', 'zoom', 'teams'],
});
const runtimeBundle = meetingSdk.platformRuntimeBundle('google-meet');
const routeMatrix = meetingSdk.platformAdapterRouteMatrix();
const blueprintMatrix = meetingSdk.platformAdapterBlueprintMatrix();
const adapterDecision = meetingSdk.platformAdapterDecision({
  url: location.href,
  title: document.title,
});
const startupPlan = meetingSdk.platformAdapterStartupPlan({
  url: location.href,
  title: document.title,
});

console.log(connectorReleaseGate.accepted, connectorReleaseGate.rows);
console.log(connectorPlatformRoadmap.recommended_first_platform, connectorPlatformRoadmap.rows);
console.log(connectorAdapterMatrix.runtime_invariants, connectorAdapterMatrix.rows);
console.log(consumerHandoff.sdk_facade_handoff.minimal_realtime_flow);
console.log(consumerHandoff.sdk_facade_handoff.surface_wiring.browser_extension);
console.log(consumerHandoff.surface_coverage_matrix.rows);
console.log(consumerHandoff.adaptation_roadmap.rows);
console.log(capability.timeline_capabilities);
console.log(executionPlan.steps);
console.log(capabilityMatrix.rows);
console.log(implementationHandoff.implementation_flow);
console.log(adapterExportPackage.host_files);
console.log(adapterImportPlan.install_steps);
console.log(adapterInstallManifest.platform_registry);
console.log(adapterLaunchPlan.runtime_actions);
console.log(adapterSmoke.accepted, adapterSmoke.rows);
console.log(blueprintMatrix.rows);
console.log(adapterDecision.selected_surface, adapterDecision.runtime_actions);
console.log(startupPlan.install_target, startupPlan.actions, startupPlan.message_contract);
await adapterSession.observeAxis();
await adapterSession.insertAnnotation({
  label: 'why?',
  captured_at_ms: Date.now(),
});
await adapterRunner.open({ url: 'https://meet.google.com/abc-defg-hij' });
await adapterRunner.insertAnnotation({
  label: 'follow up',
  captured_at_ms: Date.now(),
});
await adapterMessageBridge.handleMessage({
  type: 'meeting_timeline.observe_candidates',
  payload: {
    tabs: [{ url: 'https://meet.google.com/abc-defg-hij', active: true }],
  },
});
await adapterMessageBridge.handleMessage({
  type: 'meeting_timeline.insert_mark',
  payload: {
    mark: { label: 'from extension message', captured_at_ms: Date.now() },
  },
});
```

这个 facade 也直接暴露宿主集成需要的机器可读适配产物：`adapterProfile()`、`observerPlan()`、`selectAdapter()`、`handoff()`、`handoffMatrix()`、`handoffAcceptance()`、`hostPackage()`、`connectorPackage()`、`connectorReleaseGate()`、`connectorPlatformRoadmap()`、`connectorAdapterMatrix()`、`connectorHostAdapterConfig()`、`resolveConnectorHostAdapterConfig()`、`connectorHostAdapterBootstrapPlan()`、`connectorHostAdapterBootstrapPlanMatrix()`、`connectorHostAdapterBootstrapPlanMatrixAcceptanceReport()`、`providerReplayReport()`、`providerReplayMatrix()`、`meetingAppAdapterCapability()`、`meetingAppAdapterCapabilityMatrix()`、`meetingAppAdapterExecutionPlan()`、`meetingAppAdapterExecutionPlanMatrix()`、`platformAdaptationPackage()`、`platformConsumerHandoff()`、`platformImplementationHandoff()`、`platformAdapterAuthoringPlan()`、`platformAdapterPortfolio()`、`platformAdapterAcceptanceChecklist()`、`platformAdapterExportPackage()`、`platformAdapterImportPlan()`、`platformAdapterInstallManifest()`、`platformAdapterLaunchPlan()`、`platformAdapterSession()`、`platformAdapterRunner()`、`platformAdapterMessageBridge()`、`platformAdapterSmoke()`、`platformRuntimeBundle()`、`platformAdapterRoute()`、`platformAdapterBlueprint()`、`platformAdapterDecision()`、`platformAdapterStartupPlan()`、`platformAdapterDecisionMatrix()`、`platformAdapterStartupPlanMatrix()`、`platformAdaptationStrategy()` 和 `platformConnectorHub()`。外部项目可以先用 `hostPackage()` 拿到 Google Meet / Teams / Zoom / Webex / Lark 的 content-script/WebView/native detector 安装目标、runtime options、speaker/participant track 参数、CI gates 和 `captured_at_ms` 非阻塞标注契约；如果需要给接入面板一个更完整的单平台契约，就用 `platformAdapterBlueprint()` 拿到 surface 证据字段、验收 gates 和平台风险；再按当前会议 URL 用 `platformAdapterDecision()` 判断当前窗口应该先启用 browser extension、native detector 还是 provider reconcile；随后用 `platformAdapterStartupPlan()` 拿到当前 surface 的 `install_target`、bridge 安装函数、`meeting_timeline.observe_candidates` / `meeting_timeline.insert_mark` 消息契约、启动动作和代码引用。这个决策/启动层会保留 `provider_events_block_realtime=false` 和 `transcript_blocks_realtime=false`，避免把会后 provider/transcript 流误当作实时轴。`handoff()` 可以继续选择单场会议的启动配置。`connectorHostAdapterBootstrapPlan()` 则面向已经拿到 connector package 的宿主项目，把当前 URL/tab/window 解析、host config 加载、adapter 安装、`observe_platform_candidates` 和 `insert_annotation` 顺序压成一个最小启动计划；`connectorHostAdapterBootstrapPlanMatrix()` 会为 Google Meet / Teams / Zoom / Webex / Lark 逐平台生成同样的启动计划，适合 CI 或接入面板一次性验收多会议软件，`connectorHostAdapterBootstrapPlanMatrixAcceptanceReport()` 则给这份矩阵独立验收结果。`meetingAppAdapterCapability()` 会把某个平台当前的 realtime axis、speaker track、participant track、annotation timeline、post-meeting transcript、recording 和 subscription lifecycle 压成一个可读报告；`meetingAppAdapterExecutionPlan()` 会把这份报告转成接入步骤，明确 provider reconcile 和 transcript import 不阻塞实时标注。`platformAdaptationPackage()` 是单平台接入包，包含 adaptation playbook、provider path、runtime event plan 和风险画像；`platformConsumerHandoff()` 是多平台宿主验收包；`platformImplementationHandoff()` 是单平台落地执行单，直接列出安装 surface、bridge、adapter preflight、runtime events、provider reconcile、验收命令和 production gaps；`platformAdapterAuthoringPlan()` 是新会议软件接入前的作者计划，会给出平台 key、browser match、provider normalizer、fixture、测试和验收命令；`platformAdapterPortfolio()` 是多会议软件接入组合表，把 P0 本地实时轴、P1 provider reconcile、P2 会后 artifact、官方文档、证据要求和第一条命令压成一张下游项目可读的总表；`platformAdapterAcceptanceChecklist()` 是面向 CI / 接入面板的目标清单，按 `static`、`pilot`、`production` 明确哪些条目已过、哪些证据缺失；`platformAdapterExportPackage()` 是把单个平台交给另一个项目落地时的文件索引和 artifact 交付包；`platformAdapterImportPlan()` 是下游项目拿到 export package 后的导入校验和安装步骤；`platformAdapterInstallManifest()` 会把多个 import plan 合成宿主项目统一注册表，列出每个平台的 selected surface、content-script matches、SDK 方法和 provider reconcile 约束；`platformAdapterLaunchPlan()` 会基于当前 URL / platform / surface 从 install manifest 生成运行时启动计划，明确第一步启动哪个 surface、先调用 `observePlatformCandidates()` 还是 WebView bridge、以及实时标注必须用 `captured_at_ms`；`platformAdapterSession()` 会消费 accepted launch plan，封装 `observeAxis()`、`insertAnnotation()`、`speakerTrack()`、`participantTrack()` 和 `providerReconcile()`，并在默认情况下拒绝未建轴就写实时标注；`platformAdapterRunner()` 会从 install manifest 和当前会议 URL 一步完成 launch plan、session 创建和本地轴观察，适合下游宿主作为长期持有的 per-window runtime；`platformAdapterMessageBridge()` 会把 `meeting_timeline.observe_candidates`、`meeting_timeline.insert_mark`、`meeting_timeline.speaker_track` 等本地消息直接路由到 runner，适合扩展 background、WebView preload、native detector 在没有 HTTP runtime event endpoint 时先跑通本地链路；`platformAdapterSmoke()` 会用 Google Meet / Teams / Zoom / Webex / Lark 的 fixture URL 驱动 message bridge，验证 `observe_candidates -> insert_mark -> speaker_track -> participant_track -> provider_event` 本地链路、`captured_at_ms` 保留和 provider 非阻塞约束，适合下游项目接 SDK 后先放进 CI；`platformRuntimeBundle()` 是扩展/WebView/native runtime 可执行配置；`connectorPackage()` 会把 host package、按 surface 拆分的 observer plan、scheduler config、provider replay matrix、浏览器扩展 scaffold、runtime event plan 和硬契约收在一起，`connectorReleaseGate()` 会把这份包聚合验收成 pilot/production gate，`connectorPlatformRoadmap()` 会按 Google Meet、Zoom、Teams、Webex、Lark 给出推荐 first surface、install target、release 状态和下一步，`connectorAdapterMatrix()` 则把每个平台的 selected surface、install step、input sources、runtime event 顺序、SDK facade methods 和 evidence contract 压成宿主工程可直接装配的表，适合直接给另一个项目落地。

`connectorAdapterMatrix().rows[*].provider_replay` 会把同一平台的 provider replay 验收状态、覆盖范围、runtime event 数和 `provider_events_block_realtime=false` 一起放到宿主装配表里；下游项目不需要再从 `provider-replay-matrix.json` 反查每个平台是否通过基础 provider 回填验证。

`platformConsumerHandoff().sdk_facade_handoff` 是给下游工程师看的最小接线清单：它只假设对方 import 包根并创建 `createMeetingAppTimelineSdk({ baseUrl, platforms })`，然后列出 `observePlatformCandidates()`、`insertAnnotation()`、`speakerTrack()`、`ingestProvider()` 的调用顺序，以及 browser extension、Electron WebView、native detector、provider adapter 四类接入面的消息/方法映射。Google Meet 的 provider path 会显示为 `google_workspace_events_pubsub`，Teams 为 `microsoft_graph_change_notifications`，Zoom/Webex 为 webhook，但实时标注仍统一走本地轴和 `captured_at_ms`。

`platformConsumerHandoff().surface_coverage_matrix` 是给产品/工程排期看的覆盖表：每个平台都会列出 browser extension、WebView preload、native detector、provider reconcile、post-meeting backfill、lightweight connector、speaker track、participant track 是否已具备 SDK 接线能力。它不代表真实生产证据已完成，但能快速判断 Google Meet、Teams、Zoom、Webex、Lark 分别应该从哪个接入面先落地。

`platformConsumerHandoff().adaptation_roadmap` 是给 SDK 消费方的适配路线图：默认优先级是 Google Meet、Zoom、Teams、Webex、Lark，并为每个平台给出 `recommended_first_surface`、`next_phase`、`provider_path`、`provider_permission_risk` 和 `production_gaps`。这个顺序可以通过 `priorityPlatformOrder` 覆盖；当前建议先从 browser extension / WebView 的本地轴落地实时标注，provider 事件只作为 reconcile/backfill，不阻塞用户边开会边写标注。

`recommended_first_surface` 会读取同一份 `adapter_surfaces` 指南，而不是固定写死为浏览器扩展：Google Meet 默认是 `browser_extension`，Zoom 默认是 `native_detector`，Teams 的 `desktop_observer` 会映射成宿主可启动的 `native_detector`。因此接入配置页可以直接展示 `surface_coverage_matrix.rows` 和 `adaptation_roadmap.rows`，不需要另写平台判断。

`platformImplementationHandoff('google-meet')` 是给接入工程师看的单平台执行单：它把 `adaptation_roadmap`、`platformRuntimeBundle()`、adapter route 和 adapter preflight 合成一份 JSON，包含 `implementation_flow`、`install_surface`、`adapter_preflight`、`runtime_events`、`provider_reconcile`、`contracts` 和 `acceptance.commands`。默认 URL-only 状态会是 `adapter_preflight.status === 'needs_live_page_evidence'`：这表示安装和接线方式已明确，但下游必须在真实会议窗口里补 DOM snapshot、candidate tabs/windows、进程/Accessibility/音频状态，直到 `adapter_preflight.realtime_annotation_ready === true` 后才能第一次 `insertAnnotation()`。如果要批量交付给另一个项目，可用：

```sh
npm run meeting-platform:implementation-handoff -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-implementation-handoffs \
  --report-file=data/meeting-platform-implementation-handoff-report.json
```

如果要先评估一个 SDK 尚未内置的新会议软件，用 `platformAdapterAuthoringPlan('Acme Rooms')` 或 CLI 生成接入计划。它不会假设 provider 事件能实时到达，而是把本地候选观察、`captured_at_ms` 标注写入、provider 事件后验 reconcile 和 post-meeting artifact 分开，便于把 Google Meet / Teams / Zoom / Webex / Lark 的模式扩展到新平台：

```sh
npm run meeting-platform:adapter-authoring -- \
  --platforms=google-meet,"Acme Rooms" \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-adapter-authoring \
  --report-file=data/meeting-platform-adapter-authoring-report.json
```

如果要给另一个项目一个“多会议软件接入看板”的单一 JSON，优先用 `platformAdapterPortfolio()` 或 CLI。它会为每个平台列出 `adapter_surfaces`、`launch_requirements`、`p0_realtime_axis`、`p1_provider_reconcile`、`p2_post_meeting_backfill`、`official_doc_count`、`implementation_ready`、`pilot_ready`、`production_ready` 和 `next_action`。这里的 `recommended_first_surface` 会复用 runtime profile 的平台画像：Google Meet 默认 `browser_extension`，Zoom 默认 `native_detector`，Teams 的 desktop observer 会映射到宿主可启动的 native detector；自定义平台仍保留 authoring plan 的浏览器优先骨架，避免误套内置平台策略：

```sh
npx meeting-platform-adapter-portfolio \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-adapter-portfolio \
  --report-file=data/meeting-platform-adapter-portfolio-report.json
```

在当前 demo 仓库内也可以继续用 `npm run meeting-platform:adapter-portfolio`。发布后的 SDK 会把同名 bin 暴露到 `node_modules/.bin/meeting-platform-adapter-portfolio`，所以下游宿主项目不需要复制本仓库的 `scripts/` 目录。

portfolio 解决“该接哪条路”，acceptance checklist 解决“当前能不能过某个目标”。`target=static` 只检查 SDK 静态契约；`target=pilot` 还要求真实本地会议观察证据；`target=production` 还要求 provider reconcile、生产证据和可选 runtime replay：

```sh
npm run meeting-platform:adapter-acceptance-checklist -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=pilot \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-adapter-acceptance-checklists \
  --report-file=data/meeting-platform-adapter-acceptance-checklist-report.json
```

如果已经决定把某个平台交给另一个项目接入，用 `platformAdapterExportPackage()` 或 CLI 生成“可拿走”的文件包。它会把 `authoring-plan.json`、`acceptance-checklist.json`、`implementation-handoff.json`、`adapter-blueprint.json`、`runtime-bundle.json`、`adapter-contract.json`、`provider-connection.json` 和 `adaptation-package.json` 组织到同一平台目录；主 `adapter-export-package.json` 只保留索引、SDK import、surface entrypoint、setup order 和核心硬约束，详细 artifact 拆成独立 JSON，避免一个对象过大：

```sh
npm run meeting-platform:adapter-export-package -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=static \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-platform-adapter-export-packages \
  --report-file=data/meeting-platform-adapter-export-package-report.json
```

这个 export package 和独立 SDK 的关系是：当前仍在同一个仓库内交付，但输出结构已经按下游项目消费设计。其他项目只需要读 `adapter-export-package.json`，按 `host_files` 取对应 artifact，并优先实现 `setup_order` 的 P0 本地轴流程：安装本地 surface 后，先跑 `adapter_preflight`，用 live DOM snapshot、candidate tabs/windows 或进程/Accessibility/音频状态把 `adapter_preflight.realtime_annotation_ready` 验到 `true`；再 `observePlatformCandidates()` 绑定当前会议，最后用 `insertAnnotation(platform, { captured_at_ms, ...mark })` 写实时标注。URL-only 会保持 `needs_live_page_evidence`，不会误报可写入；provider 事件、转写和录制产物只作为 reconcile/backfill。

安装 SDK 包后也可以直接运行同名 bin：

```sh
npx meeting-platform-adapter-export-package \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=static \
  --base-url=https://timeline.example.com \
  --out-dir=meeting-platform-adapter-export-packages
```

下游项目拿到这份目录后，用 `platformAdapterImportPlan()` 或 CLI 做导入校验。它会确认 `adapter-export-package.json` 的硬约束、adapter preflight gate、所选 surface 是否可用、`host_files` 是否齐全，并自动读取同目录里的 `adapter-blueprint.json`，把 primary surface、surface order、第一条验收 gate 和 provider/transcript 非阻塞约束压进 import plan：

```sh
npx meeting-platform-adapter-import-plan \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=static \
  --dir=meeting-platform-adapter-export-packages \
  --out-dir=meeting-platform-adapter-import-plans
```

`adapter-import-plan.json` 的关键字段是 `selected_surface`、`adapter_blueprint`、`adapter_preflight`、`host_file_coverage`、`readiness.issues` 和 `install_steps`。对于 Google Meet / Teams / Zoom / Webex / Lark，默认 surface 会优先选择本地可观测路径：Google Meet 通常是 `browser_extension`，Zoom/桌面优先场景会落到 `native_detector`；如果接入方明确要 Electron WebView 或 native detector，可以传 `--surface=webview-preload` 或 `--surface=native-detector`，旧的 `--surface=native-host` 仍作为兼容别名。如果是尚未内置的新会议软件，默认不会把它当成可运行适配器；只有显式 `allowCustomAuthoring` 时，import plan 才会把它视为“可继续作者接入”的计划。

多个平台的 import plan 验收通过后，再用 `platformAdapterInstallManifest()` 或 CLI 生成宿主统一安装清单。它不会替代真实会议软件的运行时代码，而是把下游项目需要注册的 surfaces 合并成一份 manifest：browser extension 的 `content_scripts`/`host_permissions`、WebView/native detector 的 platform registry、`adapter_blueprints` 索引、`adapter_preflight` registry、provider reconcile 的非阻塞约束、以及 `run_adapter_preflight_before_realtime_session -> observePlatformCandidates() -> insertAnnotation()` 的顺序约束。

```sh
npx meeting-platform-adapter-install-manifest \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=static \
  --dir=meeting-platform-adapter-import-plans \
  --out-file=meeting-platform-adapter-install-manifest.json
```

`adapter-install-manifest.json` 的关键字段是 `platform_registry`、`adapter_blueprints.rows`、`browser_extension.content_scripts`、`native_detector.rows`、`selected_surfaces`、`provider_reconcile.rows`、`install_sequence` 和 `readiness.issues`。如果你指定了 `--platforms`，某个平台缺少 `adapter-import-plan.json` 会让 CLI report 的 `ok=false`，避免漏装某个会议软件时仍然显示可发布。

宿主项目在真正打开会议窗口前，可以先用 `platformAdapterStartupPlan()` 或 SDK CLI 生成启动计划。它不依赖 install manifest，适合给 Google Meet/Teams/Zoom/Webex/Lark 的扩展、WebView preload 或 native detector 做“第一步该启动什么”的运行时配置：Google Meet 会落到 `browser_extension`，Teams/Zoom 会落到 `native_detector`，每个平台都会列出 `observe_axis`、`insert_realtime_annotation`、speaker/participant marker、provider reconcile 的顺序，并保留 `captured_at_ms` 和 provider/transcript 非阻塞约束。

```sh
npx meeting-platform-adapter-startup \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=meeting-platform-adapter-startup \
  --report-file=meeting-platform-adapter-startup-report.json

npx meeting-platform-adapter-startup \
  --platform=google-meet \
  --url=https://meet.google.com/abc-defg-hij \
  --out-file=google-meet-startup-plan.json
```

`adapter-startup-report.json` 的关键字段是 `realtime_startup_ready_count`、`rows[*].selected_surface`、`install_target`、`observe_action`、`insert_action` 和 `provider_events_block_realtime=false`。默认报告只保留轻量摘要，完整 per-platform startup plan 写到 `--out-dir`；如果 CI 或调试需要在报告里带压缩版 plan summary，再显式加 `--include-plans=true`。这份计划用于启动本地实时轴；如果 `selected_surface=provider_reconcile`，报告会明确它不能作为实时标注的 primary surface。

startup plan 只能说明“应该启动哪个 surface”；真正打开会议窗口后，还需要用 `platformAdapterPreflight()` 或 SDK CLI 验证 live DOM/native evidence 是否足够建实时轴。Google Meet 这类 browser surface 可以传 DOM snapshot；Teams/Zoom 这类 native-first surface 可以传窗口、进程、Accessibility 或音频通话状态。URL-only preflight 会返回 `needs_live_page_evidence`，不会误报 realtime ready：

```sh
npx meeting-platform-adapter-preflight \
  --mode=candidates \
  --input-file=current-meeting-candidates.json \
  --base-url=https://timeline.example.com \
  --report-file=meeting-platform-adapter-preflight-report.json
```

`adapter-preflight-report.json` 的关键字段是 `mode`、`status`、`accepted_count`、`realtime_ready_count`、`live_evidence_ready_count`、`rows[*].selected_surface` 和 `rows[*].issue_codes`。`mode=matrix` 适合 CI 一次检查 Google Meet / Teams / Zoom / Webex / Lark 的 evidence 覆盖，`mode=candidates` 适合扩展 background 或 native detector 把当前窗口列表交给 SDK 选择可建轴的候选。

运行时打开某个会议窗口后，用 `platformAdapterLaunchPlan()` 或 CLI 把当前 URL / 显式 platform 映射到具体启动动作。它会消费 `adapter-install-manifest.json` 和其中的 `adapter_blueprints` 索引，自动识别 Google Meet / Zoom / Teams / Webex / Lark URL，选择已安装的 surface，并输出 `adapter_blueprint` 摘要与 `runtime_actions`。如果只选择 `provider-reconcile`，launch plan 会拒绝作为实时启动面，因为 provider 只能会后对齐，不能替代本地轴。

```sh
npx meeting-platform-adapter-launch-plan \
  --manifest-file=meeting-platform-adapter-install-manifest.json \
  --url=https://meet.google.com/abc-defg-hij \
  --out-file=meeting-platform-adapter-launch-plan.json
```

`adapter-launch-plan.json` 的关键字段是 `platform`、`selected_surface`、`surface_entrypoint`、`runtime_actions`、`axis_contract` 和 `readiness.issues`。宿主项目可以把它作为最后一层运行时自检：先按 `surface_entrypoint` 启动 content script / WebView preload / native detector，再按 `runtime_actions` 绑定当前会议轴并写入实时标注。

如果宿主项目不想自己维护调用顺序，可以直接用 `createMeetingPlatformAdapterSession(launchPlan, client)` 或 facade 的 `platformAdapterSession(launchPlan)`。session 的硬约束是：先调用 `observeAxis()` 让本地候选会议建轴，再调用 `insertAnnotation({ captured_at_ms, ...mark })` 写入当前标注；如果未建轴就写入，默认会抛出 `adapter_session_axis_not_observed`。同一个 session 也暴露 `speakerTrack()`、`participantTrack()` 和 `providerReconcile()`，用于把发言人位置、参会者轨迹和会后 provider 事件放进同一条会议时间轴。

如果宿主项目连 launch plan 和 session 生命周期也不想自己维护，用 `createMeetingPlatformAdapterRunner(adapterInstallManifest, client)` 或 facade 的 `platformAdapterRunner(adapterInstallManifest)`。runner 的 `open({ url })` 会先生成 accepted launch plan，再创建 session，并默认立刻 `observeAxis()`；之后同一个 runner 直接 `insertAnnotation()`、`speakerTrack()`、`participantTrack()`。这更适合浏览器扩展 content script、Electron WebView preload 或 native detector 的 per-window runtime。

如果宿主项目的边界已经是消息，而不是直接调用 SDK 方法，用 `createMeetingPlatformAdapterMessageBridge(adapterInstallManifest, client)` 或 facade 的 `platformAdapterMessageBridge(adapterInstallManifest)`。它接收 `meeting_timeline.observe_candidates` 建轴，接收 `meeting_timeline.insert_mark` / `meeting_timeline.insert_annotation` 写入标注，也支持 `meeting_timeline.speaker_track`、`meeting_timeline.participant_track` 和 `meeting_timeline.provider_event`。如果需要在打开 session 前先证明候选窗口有 live evidence，用 `meeting_timeline.candidate_launch_plan` 只取严格启动计划，或用 `meeting_timeline.open_candidate_session` 在候选 preflight 通过后直接打开 session；URL-only 候选会返回 `needs_live_page_evidence`，不会冒充实时 ready。这层适合扩展 background、WebView preload、native detector 先用本地 SDK 跑通；如果宿主已经有跨进程 HTTP endpoint，则继续用 `platform-runtime-event`。

把 SDK 交给另一个宿主项目时，可以先跑 `platformAdapterSmoke(adapterInstallManifest)` 或子模块的 `runMeetingPlatformAdapterSmoke()`。它会为 Google Meet / Teams / Zoom / Webex / Lark 生成 fixture 会议 URL，创建 message bridge，按顺序发送 `meeting_timeline.observe_candidates`、`meeting_timeline.insert_mark`、`meeting_timeline.speaker_track`、`meeting_timeline.participant_track` 和 `meeting_timeline.provider_event`，并返回 `meeting_platform_adapter_smoke_report`。这个报告只验证本地实时标注链路：已先建轴、`captured_at_ms` 未丢失、speaker/participant 位置可写、provider reconcile 不阻塞实时标注，并确认每个平台的 `adapter_blueprint` 已进入 message bridge 启动链路；它不要求真实会议开始/结束事件，也不要求转写内容。

同一条 smoke 也有 CLI，适合宿主 CI 直接跑。没有 `--manifest-file` 时，CLI 会按 `--platforms` 生成静态 install manifest；有 manifest 时会用真实交付清单验收：

```sh
npx meeting-platform-adapter-smoke \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=static \
  --json=true

npx meeting-platform-adapter-smoke \
  --manifest-file=meeting-platform-adapter-install-manifest.json \
  --platforms=google-meet,zoom \
  --fail-on-blocked=true
```

需要把这份 connector package 直接落盘给另一个宿主项目时，可以用 SDK 自带 CLI：

```sh
npx meeting-app-connector-package \
  --platforms=google-meet,teams,zoom,webex,lark \
  --surfaces=browser-extension,native-detector \
  --base-url=https://timeline.example.com \
  --out-dir=meeting-app-connector-package \
  --report-file=meeting-app-connector-package-report.json
```

CLI 会输出 `connector-package.json`、`connector-handoff.json`、`connector-adoption-index.json`、`connector-field-intake-index.json`、`connector-release-gate.json`、`connector-adapter-matrix.json`、`connector-adapter-matrix-acceptance.json`、`host-adapter-config-index.json`、`host-adapter-configs/{platform}.json`、`host-adapter-bootstrap-plan-matrix.json`、`host-adapter-bootstrap-plan-matrix-acceptance.json`、`connector-platform-roadmap.json`、`connector-bridge-handoff.json`、`connector-bridge-handoff-acceptance.json`、`connector-bridge-smoke-report.json`、`connector-quickstart.md`、`host-install-checklist.json`、`host-install-checklist-acceptance.json`、`connector-smoke-plan.json`、`connector-smoke-plan-acceptance.json`、`connector-smoke-run-report.json`、`host-package.json`、handoff/acceptance 矩阵、`adapter-blueprint-matrix.json`、`startup-plan-matrix.json`、按 surface 拆分的 observer plan / scheduler config、runtime event plan，以及可作为起点的浏览器扩展 scaffold。报告会剥离 scaffold 文件内容，避免 CI 日志和交接摘要过大；实际文件会写在 `extension/` 目录下。下游工程应优先打开 `connector-quickstart.md`、`connector-release-gate.json`、`connector-adapter-matrix.json`、`host-adapter-config-index.json`、`host-adapter-configs/{platform}.json`、`host-adapter-bootstrap-plan-matrix.json`、`host-adapter-bootstrap-plan-matrix-acceptance.json`、`connector-platform-roadmap.json`、`connector-adoption-index.json`、`connector-field-intake-index.json`、`connector-bridge-handoff.json` 和 `startup-plan-matrix.json`，用 release gate 先判断整包是否能进入实时标注试点或生产交付；其中 host bootstrap acceptance 已并入 release gate，避免 URL/window 识别、config 加载、adapter 安装或 observe-before-insert 顺序不满足时仍误判整包可用。用 adapter matrix 判断每个平台的 selected surface、install step、`observe_platform_candidates -> insert_annotation` 顺序、SDK 方法和 evidence contract，用 host adapter config 给单平台宿主直接加载 runtime endpoint、`captured_at_ms` 契约、bridge contract、SDK facade 和 provider replay 状态，用 host adapter bootstrap plan matrix 检查 URL/window 识别、config 加载、adapter 安装、observe candidates 和 insert annotation 的顺序是否可直接接入，并用对应 acceptance report 做独立 CI gate；用 platform roadmap 判断 Google Meet、Zoom、Teams、Webex、Lark 的推荐接入顺序、first surface、install target 和下一步，用 adoption index 判断每个平台的 P0 轴来源、安装目标、实时标注状态、bridge 状态和生产前 evidence 缺口，再用 field intake index 看到真实采样文件路径、必需 active/end 快照、provider start/end 覆盖和 real-intake 验证命令，最后用 selected surface / install target 决定先装 content script、WebView preload 还是 native detector；CI 或接入面板可以直接读取 `host-install-checklist.json`，逐平台检查 install target、runtime action client method 和 `captured_at_ms` 契约，再用 `host-install-checklist-acceptance.json` 做独立 gate；bridge 接入可以用 `connector-bridge-handoff-acceptance.json` 独立检查 SDK module、content-script bridge factory、消息类型、runtime action 和 host endpoint 要求，并用 `connector-bridge-smoke-report.json` 确认统一 `meeting_timeline.*` 消息能 dry-run 到 `observe_platform_candidates` / `insert_annotation` runtime event；接入方真正试跑时读取 `connector-smoke-plan.json`，按 observe candidates、insert annotation、speaker/participant marker 的顺序发送事件，并用 `connector-smoke-plan-acceptance.json` 作为 smoke gate；`connector-smoke-run-report.json` 是 SDK 生成时的 dry-run 结果，只证明计划能映射到 runtime client 方法，不替代真实会议现场 evidence。

下游项目拿到 connector package 后，不需要自己猜哪些字段是硬约束，可以直接用 `meeting-app-connector-package` 子模块验收、生成交接摘要，或按平台取宿主配置：

```js
import {
  assertMeetingAppTimelineConnectorPackage,
  buildMeetingAppTimelineConnectorAdoptionIndex,
  buildMeetingAppTimelineConnectorAdapterMatrix,
  buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport,
  buildMeetingAppTimelineConnectorFieldIntakeIndex,
  buildMeetingAppTimelineConnectorBridgeHandoff,
  buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport,
  buildMeetingAppTimelineConnectorHandoff,
  buildMeetingAppTimelineHostAdapterBootstrapPlan,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix,
  buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport,
  buildMeetingAppTimelineHostAdapterConfig,
  buildMeetingAppTimelineHostAdapterConfigIndex,
  buildMeetingAppTimelineConnectorPackageAcceptanceReport,
  buildMeetingAppTimelineConnectorReleaseGate,
  buildMeetingAppTimelineConnectorSmokePlan,
  buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport,
  createMeetingAppTimelineConnectorRuntimeClient,
  resolveMeetingAppTimelineHostAdapterConfig,
  runMeetingAppTimelineConnectorBridgeSmoke,
  runMeetingAppTimelineConnectorSmokePlan,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-connector-package';

const acceptance = buildMeetingAppTimelineConnectorPackageAcceptanceReport(connectorPackage, {
  target: 'realtime',
});
assertMeetingAppTimelineConnectorPackage(connectorPackage);
const handoff = buildMeetingAppTimelineConnectorHandoff(connectorPackage);
const adoptionIndex = buildMeetingAppTimelineConnectorAdoptionIndex(connectorPackage);
const fieldIntakeIndex = buildMeetingAppTimelineConnectorFieldIntakeIndex(connectorPackage);
const hostConfig = buildMeetingAppTimelineHostAdapterConfig(connectorPackage, 'google-meet');
const hostConfigIndex = buildMeetingAppTimelineHostAdapterConfigIndex(connectorPackage);
const resolvedHostConfig = resolveMeetingAppTimelineHostAdapterConfig(connectorPackage, {
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', active: true }],
});
const bootstrapPlan = buildMeetingAppTimelineHostAdapterBootstrapPlan(connectorPackage, {
  tabs: [{ url: 'https://meet.google.com/abc-defg-hij', active: true }],
});
const bootstrapPlanMatrix = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrix(connectorPackage);
const bootstrapPlanMatrixAcceptance = buildMeetingAppTimelineHostAdapterBootstrapPlanMatrixAcceptanceReport(bootstrapPlanMatrix);
const bridgeHandoff = buildMeetingAppTimelineConnectorBridgeHandoff(connectorPackage);
const bridgeAcceptance = buildMeetingAppTimelineConnectorBridgeHandoffAcceptanceReport(bridgeHandoff);
const bridgeSmoke = await runMeetingAppTimelineConnectorBridgeSmoke(bridgeHandoff);
const smokePlan = buildMeetingAppTimelineConnectorSmokePlan(connectorPackage);
const smokeAcceptance = buildMeetingAppTimelineConnectorSmokePlanAcceptanceReport(smokePlan);
const smokeRun = await runMeetingAppTimelineConnectorSmokePlan(smokePlan);
const releaseGate = buildMeetingAppTimelineConnectorReleaseGate(connectorPackage, {
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  bridgeHandoffAcceptance: bridgeAcceptance,
  bridgeSmokeReport: bridgeSmoke,
  smokePlan,
  smokePlanAcceptance: smokeAcceptance,
  smokeRunReport: smokeRun,
  hostAdapterBootstrapPlanMatrixAcceptanceReport: bootstrapPlanMatrixAcceptance,
});
const adapterMatrix = buildMeetingAppTimelineConnectorAdapterMatrix(connectorPackage, {
  releaseGate,
  adoptionIndex,
  fieldIntakeIndex,
  bridgeHandoff,
  smokePlan,
});
const adapterMatrixAcceptance = buildMeetingAppTimelineConnectorAdapterMatrixAcceptanceReport(adapterMatrix);
const runtime = createMeetingAppTimelineConnectorRuntimeClient(connectorPackage, { fetch });
await runtime.insertAnnotation('google-meet', {
  id: 'mark-001',
  label: 'why?',
  captured_at_ms: Date.now(),
});
```

默认 `target: 'realtime'` 会检查 `captured_at_ms`、provider/transcript 非阻塞、runtime event endpoint、`observe_meeting_app` / `insert_annotation` / speaker / participant 轨道动作、每个 surface 的 observer plan 和 scheduler config，以及浏览器扩展 scaffold。`target: 'production'` 会额外要求真实会议 live snapshot evidence，因此不会把静态包误判成生产可上线。

`createMeetingAppTimelineConnectorRuntimeClient()` 会从 connector package 读取 runtime event endpoint 和各平台支持的 action，发送前先校验 action/platform 是否在 package 的 runtime plan 中。外部项目可以把它用于 Google Meet content script、Teams WebView preload、Zoom/Webex native helper 或 Lark 长连接代理，统一发 `observe_meeting_app`、`observe_platform_candidates`、`insert_annotation`、`speaker_track` 和 `participant_track`，不用手写 action 路由表。

如果宿主项目已经确定只接一个会议平台，可以用更轻的 `meeting-platform-connector` façade。它把 registry、provider connection、browser observer、adapter route、runtime event endpoint 和 normalizer 压成单个平台对象；`createMeetingPlatformConnectorRuntime()` 会固定平台名，所以 Google Meet/Teams/Zoom/Webex/Lark 的接入侧只需要调用 `insertAnnotation(mark)`、`observeMeetingApp(snapshot)`、`observePlatformCandidates(env)` 或 `normalizeProviderEvent(raw)`：

```js
import {
  assertMeetingPlatformConnector,
  buildMeetingPlatformConnector,
  createMeetingPlatformConnectorRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector';

const connector = buildMeetingPlatformConnector('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
assertMeetingPlatformConnector(connector);

const runtime = createMeetingPlatformConnectorRuntime(connector, { fetch });
await runtime.insertAnnotation({
  id: 'mark-001',
  label: 'why?',
  captured_at_ms: Date.now(),
});
const signals = runtime.normalizeProviderEvent(googleWorkspaceEventBody);
```

如果宿主项目要同时适配多种会议软件，用 `createMeetingPlatformConnectorHub()` 更合适。它会先看显式 `platform/provider`，再从 `url`、`meeting_url`、`tab.url`、`tabs[].url`、`windows[].tabs[].url` 里识别 Google Meet、Teams、Zoom、Webex 或 Lark，并把标注自动路由到对应平台 runtime：

```js
import {
  createMeetingPlatformConnectorHub,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector';

const hub = createMeetingPlatformConnectorHub({
  baseUrl: 'https://timeline.example.com',
  fetch,
});

await hub.insertAnnotation({
  url: location.href,
  title: document.title,
}, {
  id: 'mark-001',
  label: 'why?',
  captured_at_ms: Date.now(),
});

await hub.observePlatformCandidates({
  tabs: [{ active: true, url: location.href, title: document.title }],
});
```

如果注入点是浏览器扩展 content script、Electron WebView preload 或移动端内嵌 WebView，可以直接使用 connector browser runtime。它会复用 `meeting-app-browser-runtime` 的 DOM 采样、MutationObserver 和 speaker/participant track 逻辑，但底层写入走 connector hub，不要求宿主提供完整 timeline client：

```js
import {
  installMeetingPlatformConnectorContentScriptBridge,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-connector';

installMeetingPlatformConnectorContentScriptBridge({
  baseUrl: 'https://timeline.example.com',
  fetch,
});
```

安装后，background worker 或宿主 WebView 只要发送 `meeting_timeline.insert_mark` / `meeting_timeline.sample` / `meeting_timeline.sample_tracks` / `meeting_timeline.preflight_current_window` 消息；bridge 会从当前 `location.href` 识别 Google Meet、Teams、Zoom、Webex 或 Lark，并投递统一的 `insert_annotation`、`observe_meeting_app`、`speaker_track` 或 `participant_track` runtime event。`preflight_current_window` 不会访问服务端，它只在当前 content script / WebView preload 内捕获 DOM，并返回这一个会议窗口是否已经具备实时建轴和标注插入条件。浏览器扩展 background 还支持 `meeting_timeline.preflight_candidates`：它先 query 候选会议 tabs，再逐个转发 `preflight_current_window` 到对应 content script，最后返回 per-tab readiness 汇总，适合 popup 或 native host 一次性判断多场 Google Meet / Teams / Zoom / Webex / Lark 窗口。

`platform-kit` 也暴露同一入口：`kit.platformConnector('google-meet')`、`kit.platformConnectorMatrix()`、`kit.platformConnectorHub()`、`kit.resolvePlatformConnector(input)`、`kit.platformConnectorAcceptance(connector)`、`kit.createPlatformConnectorRuntime(connector, { fetch })`、`kit.createPlatformConnectorHub({ fetch })`、`kit.createPlatformConnectorBrowserRuntime({ fetch })` 和 `kit.createPlatformConnectorContentScriptBridge({ fetch })`。这适合宿主项目已经统一使用 `createMeetingPlatformTimelineKit()`，但仍希望按平台懒加载 Google Meet / Teams / Zoom / Webex / Lark connector runtime，或者直接把当前浏览器/会议窗口状态交给 SDK 自动分发。

交付给其他项目时，`platform-consumer-handoff` 会把这条轻量路径写进 `lightweight_connector_handoff`：包括 `meeting-platform-connector` 模块名、hub/browser runtime/content-script bridge factory、`meeting_timeline.insert_mark` / `meeting_timeline.sample` / `meeting_timeline.sample_tracks` / `meeting_timeline.preflight_current_window` 消息类型、host runtime event endpoint 和 `captured_at_ms` 契约。这样接入方可以明确选择两条路线：完整 `platform-integration-runtime` 用于 host 级编排，或轻量 connector bridge 用于浏览器扩展、Electron WebView preload、移动端 WebView 的页面侧事件投递。

`platform-kit` 同样暴露这一层：`kit.meetingAppAdapterIntegrationPackage('google-meet')` 和 `kit.meetingAppAdapterIntegrationPackageMatrix()`。CI 里可以用 `assertMeetingAppAdapterIntegrationPackage()`、`assertMeetingAppAdapterIntegrationPackageMatrix()` 或 kit 上的同名方法做 gate；默认 target 是 `pilot`，如果传 `target: 'production'`，则必须补齐真实会议 evidence package、provider start/end reconcile 和 handoff readiness 之后才会通过。

同一层也可以从 CLI 直接导出，默认覆盖 Google Meet、Teams、Zoom、Webex、Lark：

```sh
npm run meeting-app:adapter-integration-package -- \
  --base-url=https://timeline.example.com \
  --out-dir=data/meeting-app-adapter-integration-packages \
  --report-file=data/meeting-app-adapter-integration-package-report.json
```

每个平台目录会包含 `runtime-delivery.json` 和 `integration-package.json`；报告里的 `target_accepted_count` 可用来区分 pilot / realtime / production gate。

安装 SDK 包以后，这个导出能力也可以直接作为 npm bin 使用：

```sh
npx meeting-app-adapter-integration-package \
  --platforms=google-meet,teams,zoom,webex,lark \
  --base-url=https://timeline.example.com \
  --out-dir=meeting-app-adapter-integration-packages \
  --report-file=meeting-app-adapter-integration-package-report.json
```

如果下游项目要直接启动浏览器扩展、WebView preload 或 native host runtime，用 `platform-runtime-bundle`。它在 `platform-adaptation-package` 基础上再补一层可执行运行时配置：content script manifest、浏览器 URL matches、adapter route、`meeting-app-browser-runtime` preset、`platform-integration-runtime` content-script bridge 安装参数、轻量 `meeting-platform-connector` bridge 安装参数、mutation observer / speaker filter 参数、extension message 示例、host ingest endpoints，以及 `captured_at_ms` 写入契约：

```js
import {
  buildMeetingPlatformRuntimeBundle,
  buildMeetingPlatformRuntimeBundleMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-runtime-bundle';

const googleRuntime = buildMeetingPlatformRuntimeBundle('google-meet', {
  baseUrl: 'https://timeline.example.com',
});

// googleRuntime.browser.manifest 可交给扩展构建器。
// googleRuntime.adapter_route.routes[0].route === 'local_observer_axis'。
// googleRuntime.browser.permissions 和 googleRuntime.messaging.background_message_types 包含候选会议观察所需的 tabs + meeting_timeline.observe_candidates。
// googleRuntime.runtime.content_script_bridge.options 可直接传给平台级 content-script bridge。
// googleRuntime.runtime.lightweight_connector_bridge.options 可直接传给 installMeetingPlatformConnectorContentScriptBridge。
// googleRuntime.messaging.lightweight_connector_message_types 是轻量 bridge 接收的 meeting_timeline.* 消息集合。
// googleRuntime.messaging.examples.content_script_insert_annotation 是外部插入标注的消息格式。
// googleRuntime.messaging.examples.observe_candidates 是 background/native host 触发 observe_platform_candidates 的消息格式。
// googleRuntime.host.endpoints.insertMark 是实时标注写入地址。
// googleRuntime.host.endpoints.runtimeEvents 是统一 runtime event envelope 写入地址。
// googleRuntime.messaging.runtime_event.client_factory 指向 createMeetingPlatformRuntimeEventClient。
// googleRuntime.messaging.runtime_event.plan 是外部宿主接入 observe/provider/annotation/speaker/participant/view/adapter-route 的动作表。
// googleRuntime.readiness.provider_required_for_realtime === false。
// googleRuntime.readiness.transcript_blocks_realtime === false。

const runtimeMatrix = buildMeetingPlatformRuntimeBundleMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
// runtimeMatrix.rows[*].candidate_observation_ready 可直接用于下游适配看板。
```

`platform-kit` 同样暴露 `kit.platformRuntimeBundle('google-meet')` 和 `kit.platformRuntimeBundleMatrix()`。CLI 可批量导出每个平台的 runtime bundle：

```sh
npm run meeting-platform:runtime-bundle -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --out-dir=data/meeting-platform-runtime-bundles \
  --report-file=data/meeting-platform-runtime-bundle-report.json
```

这层适合交给另一个工程直接落地运行时 glue code；它仍然不代表 production-ready，真实会议 DOM / provider 样本仍以后续 evidence package 和 handoff readiness 为准。

也可以直接导出每个平台一份 manifest 文件，给 Chrome 扩展、本地 host 或 provider recorder 读取：

```sh
npm run meeting-platform:field-intake -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --out-dir=data/meeting-platform-field-intake-plans \
  --report-file=data/meeting-platform-field-intake-report.json
```

这一步输出的是“采样前执行单”，不是生产验收；最终仍要跑 `meeting-platform:real-intake` 或 `platform-live-adapter` readiness。

给其他项目交付时，建议在 evidence package 生成后再跑一层 `platform-handoff-readiness`。evidence package 自身已经带 `adapter_route`、summary 和 verification 字段，用于说明“本地观察建轴、标注写入、发言人位置、provider 回填、会后 artifact”这几条路线是否满足实时非阻塞原则；handoff readiness 会复验同一条 adapter route，并把 adapter contract、provider connection、candidate observation、DOM 诊断、field intake、real-intake gate 和 live adapter readiness 合成一个状态：`adapter_route_blocked`、`needs_candidate_observation_contract`、`needs_local_observer_evidence`、`pilot_ready_provider_setup_pending`、`pilot_ready_provider_reconcile_pending` 或 `production_ready`。这个对象适合给宿主项目做接入面板：它不要求读完整 runbook，但能直接回答“Google Meet / Teams / Zoom / Webex / Lark 现在能不能接入，缺什么，下一条命令是什么”。其中 `adapter_route_ready` 和 `candidate_observation_ready` 都是硬门槛：前者保证本地轴优先、`captured_at_ms` 写入、provider/transcript 非阻塞，后者保证能通过 `meeting_timeline.observe_candidates` / `observe_platform_candidates` / `/api/meeting-platform/observe-candidates` 把当前会议窗口候选送进 host；provider event 和 transcript 仍然只做回填或会后处理。

```js
import {
  runMeetingPlatformHandoffReadinessMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-handoff-readiness';

const matrix = await runMeetingPlatformHandoffReadinessMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  google_meet: {
    evidencePackage: googleEvidencePackage,
  },
}, {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  target: 'production',
});

// matrix.rows 每行都有 handoff_ready / pilot_ready / production_ready，
// 以及 adapter_route_ready、adapter_first_route、candidate_observation_ready、
// runtime_host_replay_accepted、provider_missing_env、dom_record_count、next_actions。
```

CLI 入口：

```sh
npm run meeting-platform:handoff-readiness -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --package-dir=data/meeting-platform-evidence-packages \
  --report-file=data/meeting-platform-handoff-readiness-report.json
```

```sh
npm run meeting-platform:field-manifest -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex \
  --out-dir=data/meeting-platform-field-manifests \
  --report-file=data/meeting-platform-field-manifest-report.json
```

对应 CLI 可直接把现场采样目录转换成 bundle 和 evidence package。目录里的 JSON 可以是原始采样输入、按平台分组的对象、已有 evidence package，或上一次导出的 bundle：

```sh
npm run meeting-platform:field-evidence -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex \
  --dir=data/meeting-platform-field-evidence \
  --package-dir=data/meeting-platform-evidence-packages \
  --bundle-dir=data/meeting-platform-field-evidence-bundles \
  --report-file=data/meeting-platform-field-evidence-report.json
```

默认会要求 production-ready；如果只是验证“能否实时落轴”，可以加 `--require-production-ready=false`，此时 DOM-only 但缺 provider 回填的平台会按 pilot 状态通过。

对应 CLI 可以直接放到现场采样工具或 CI：

```sh
npm run meeting-platform:field-capture -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex \
  --dir=data/meeting-platform-evidence-packages \
  --report-file=data/meeting-platform-field-capture-report.json
```

默认报告会读取已有 evidence package 并输出每个平台的 `missing_items`。如果只是生成空采样计划，用 `--plan-only=true`；如果要让 CI 在仍缺采样时失败，用 `--fail-on-incomplete=true`。

真实采样交接时，`platform-evidence-correlation` 会检查 provider 事件和本地 DOM 记录是否来自同一场会议。它优先用 meeting id / URL 匹配；没有共享 id 时会退到同平台时间窗口匹配。`verifyMeetingPlatformEvidencePackage()` 默认会要求 correlation 通过，避免把不同会议的 provider 样本和 DOM 样本混成一个 production-ready 包。

业务项目如果要直接接入“会议中边写边标注”，优先用 `platform-live-adapter`。它把本地会议 App 观察、provider webhook 回填、实时标注插入和 evidence session 绑成一个对象；`observeMeetingApp()` 会同时尝试建轴并记录本地证据，`ingestProvider()` 会同时回填 provider 事件并记录 provider 证据，`insertAnnotation()` 会先走 `platform-realtime-annotation` 决策链，再按 action 执行建轴、插入或 pending。

```js
import { createMeetingPlatformLiveAdapter } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter';

const live = createMeetingPlatformLiveAdapter('google-meet', timelineClient, {
  baseUrl: 'https://timeline.example.com',
  requireMeetingEnd: false,
});

await live.observeMeetingApp(activeDomSnapshot, {
  capturedAtMs: Date.now(),
});

const inserted = await live.insertAnnotation({
  capturedAtMs: Date.now(),
  label: 'why?',
  kind: 'question',
}, {
  // 默认允许先用 capturedAtMs 低延迟落轴；量产设备可设为 true，要求先完成 /api/time 校准。
  requireClockSync: false,
});

// inserted.result.pipeline.status === 'ready_to_insert' 表示已写入当前轴。
// inserted.result.pipeline.status === 'start_axis_then_insert' 表示 SDK 已先建轴再写入。
// inserted.result.pipeline.status === 'pending_real_meeting' 表示缺会议身份，应暂存并等待真实会议绑定。
// inserted.live_evidence.can_insert_realtime_marks === true 表示当前本地观察足以支撑实时落轴。
// live.exportPackage() 可在现场采样结束后交给另一个项目复验。
```

多平台宿主可以用 suite/matrix 先生成 Google Meet / Teams / Zoom / Webex / Lark 的统一接入面板，再按平台懒加载 live adapter：

```js
import {
  buildMeetingPlatformLiveAdapterMatrix,
  createMeetingPlatformLiveAdapterSuite,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter';

const matrix = buildMeetingPlatformLiveAdapterMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

const suite = createMeetingPlatformLiveAdapterSuite(timelineClient, {
  baseUrl: 'https://timeline.example.com',
  platforms: matrix.platforms,
});

const zoomLive = suite.adapter('zoom');
const googleLive = suite.adapter('google-meet');

const readiness = suite.readiness('google-meet', {
  target: 'production',
  evidencePackage: googleLive.exportPackage(),
});

// readiness.status === 'ready' 表示 adapter 方法、captured_at_ms 时间戳约束、
// 非阻塞 provider/transcript 策略和 evidence package 复验都通过。

suite.assertReadiness('google-meet', {
  target: 'production',
  evidencePackage: googleLive.exportPackage(),
});
```

会议进行中可以用 `platform-evidence-session` 持续收集本地 DOM 和 provider webhook 证据。它的定位是 live object：每采到一个窗口快照或 webhook，就调用 `summary()` 看当前是否能实时落标注、是否还缺 provider reconcile 证据，以及后续能否导出 handoff 包：

```js
import { createMeetingPlatformEvidenceSession } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-session';

const session = createMeetingPlatformEvidenceSession('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
});

session.captureMeetingAppSnapshot(activeDomSnapshot, {
  phase: 'active',
  capturedAtMs: Date.now(),
});

const live = session.summary();
// live.can_insert_realtime_marks === true 表示本地观察证据已经足够支持当前标注落轴。
// live.provider_reconcile_ready === false 时，provider webhook 仍可后补，不阻塞当前标注。

session.captureProviderWebhook(reqLikeObject);

const packageForHandoff = session.exportPackage();
const verification = session.verify();
```

如果宿主项目拿到的是桌面窗口、浏览器标签页或 native app 进程快照，先用 `meeting-session-discovery` 把这些低层信号转成统一会议候选，再交给 observer 建轴。这个路径适合 Google Meet 浏览器页，也适合 Zoom / Teams / Lark / Webex native app 没有 webhook 或 webhook 延迟较高的情况：

```js
import { createMeetingSessionTimelineDiscovery } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-session-discovery';

const discovery = createMeetingSessionTimelineDiscovery(timeline, {
  source: 'desktop_session_discovery',
});

await discovery.observeEnvironment({
  windows: [{
    id: 'browser-1',
    focused: true,
    application: { name: 'Google Chrome', bundleId: 'com.google.Chrome' },
    tabs: [
      { id: 'mail', active: false, url: 'https://mail.google.com', title: 'Inbox' },
      { id: 'meet', active: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Review - Google Meet' },
    ],
  }, {
    id: 'zoom-1',
    title: 'Zoom Meeting',
    application: { name: 'zoom.us', bundleId: 'us.zoom.xos' },
    inMeeting: true,
  }],
}, { observedAtMs: Date.now() });
```

`meeting-session-discovery` 会优先使用 URL 中的真实会议 ID；如果 native app 没有 URL，会在明确 `inMeeting: true` 或标题/进程足够像会议窗口时生成稳定的本地 `native-{platform}-...` 会话 ID。后续收到官方 webhook 时可以继续用 `signal-reconciler` 或 `platform-ingest` 做校准/回填。

浏览器会议推荐用更高层的 `browser-meeting`。它面向浏览器扩展、Electron wrapper、自动化采集器或桌面宿主进程，输入可以是 tabs/windows/page/dom 快照；SDK 会同时完成会议会话发现和 active speaker 滤波：

```js
import { createBrowserMeetingTimelineObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/browser-meeting';

const browserMeetings = createBrowserMeetingTimelineObserver(timeline, {
  source: 'browser_extension',
  speakerOptions: {
    minStableMs: 300,
    switchStableMs: 400,
    endIdleMs: 1500,
  },
  applyOptions: { speakerAsAnnotation: true },
});

await browserMeetings.observe({
  windows: [{
    id: 'win-1',
    focused: true,
    tabs: [{
      id: 'meet-tab',
      active: true,
      audible: true,
      url: 'https://meet.google.com/abc-defg-hij',
      title: 'Review - Google Meet',
      page: {
        inMeeting: true,
        activeSpeaker: { id: 'speaker-ada', name: 'Ada', speaking: true },
      },
    }],
  }],
  observedAtMs: Date.now(),
});
```

这条路径同样适用于 Teams Web、Zoom Web、Webex Web 和 Lark/Feishu Web：平台和会议 ID 优先从 URL 解析，`page.activeSpeaker` / `dom.activeSpeaker` / `participants[].speaking` 会归一化成 `speaker_started/ended`。如果官方 webhook 之后到达，再用 reconciler 校准；实时标注不要等待官方事件。

如果宿主拿到的是更贴近真实浏览器扩展或桌面 Accessibility 的“脏输入”，例如 Google Meet 的 tile `ariaLabel`、Teams 的 `Leave` 按钮、Zoom 的窗口控件和 participant tile，推荐先走 `meeting-apps` 预设层。它会按 Google Meet / Microsoft Teams / Zoom / Lark / Webex 的常见 DOM/AX 线索，把按钮文案、tile 列表、发言状态和音量值归一成 `browser-meeting` 可识别的 `inMeeting`、`participants`、`activeSpeaker`：

```js
import { createMeetingAppObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-apps';

const meetingApps = createMeetingAppObserver({
  source: 'browser_extension',
  speakerOptions: { minStableMs: 300, switchStableMs: 400, endIdleMs: 1500 },
});

const result = meetingApps.observe({
  tabs: [{
    active: true,
    audible: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Review - Google Meet',
    page: {
      buttons: [{ ariaLabel: 'Leave call' }],
      tiles: [
        { dataset: { participantId: 'ada' }, ariaLabel: 'Ada Lovelace is speaking' },
      ],
    },
  }],
}, { observedAtMs: Date.now() });

// result.signals -> meeting_started / speaker_started
```

`meeting-apps` 是轻量 preset，不依赖具体浏览器扩展 SDK，也不要求实时 OCR 或实时转写。它的作用是把 Google Meet 等会议软件的本地可观测状态变成统一 meeting signal；官方 provider webhook 仍然走 `google-meet` / `microsoft-teams` / `zoom` 等 adapter 做校准和会后 artifact。

外部宿主项目刚接入时，可以先用 `meetingAppAdapterFit()` 做输入自检。它不会写时间轴，只判断当前 tabs/windows/DOM/AX 快照能不能被 SDK 识别成会议应用、能不能建立实时轴、能不能产出发言人/参会人轨：

```js
const fit = kit.meetingAppAdapterFit({
  tabs: [{
    active: true,
    audible: true,
    url: 'https://meet.google.com/abc-defg-hij',
    title: 'Review - Google Meet',
    page: {
      buttons: [{ ariaLabel: 'Leave call' }],
      tiles: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
    },
  }],
}, { platform: 'google-meet' });

// fit.ready_for_realtime_axis === true
// fit.ready_for_speaker_track === true
// fit.next_actions 会提示缺 meeting identity、active speaker、participant 或 end sample。
```

多平台宿主可以把不同会议软件的样本放进 `meetingAppAdapterFitMatrix()`，先比较 Google Meet / Teams / Zoom / Lark / Webex 哪些输入已经能接入，哪些还缺本地观察字段。这个报告适合作为接入面板或 CI preflight；生产发布前仍然要继续跑 live snapshot diagnosis 和 handoff readiness。

fit 通过后，宿主可以再生成 `meetingAppRuntimeObserverPlan()`，把“怎么持续观察”固化成机器可读计划。它会把 runtime preset、DOM/AX 输入契约、采样节流、active speaker 稳定窗口、会议消失后的 end grace window 和 provider/transcript 非阻塞规则放在同一个对象里：

```js
const observerPlan = kit.meetingAppRuntimeObserverPlan({
  url: 'https://meet.google.com/abc-defg-hij',
  page: {
    controls: [{ label: 'Leave call' }],
    participants: [{ id: 'ada', ariaLabel: 'Ada Lovelace is speaking' }],
  },
}, {
  platform: 'google-meet',
  surface: 'browser-extension',
});

// observerPlan.cadence.changed_observe_every_ms 控制 DOM 变化触发后的节流。
// observerPlan.cadence.meeting_missing_end_grace_ms 控制候选会议消失后多久补 meeting_ended。
// observerPlan.trigger_policy 明确 start/end/speaker/keep-alive 的触发方式。
```

`meetingAppRuntimeObserverPlanMatrix()` 可以一次生成多平台计划。浏览器扩展和 WebView 默认用 mutation observer + 低频 keep-alive；native detector 默认按窗口/Accessibility/音频快照变化触发。它仍然不替代真实 DOM 证据，只是把外部项目的观察循环和节流参数标准化，避免每个宿主自己猜采样频率。

如果要把 Google Meet / Teams / Zoom / Webex / Lark 的本地观察能力交给另一个宿主项目实现，优先生成 `meeting-app-adapter-manifest`。它是比 observer plan 更接近交付边界的清单：包含浏览器匹配规则、扩展权限、content script manifest、MutationObserver 采样参数、DOM/AX selector 数量、`observe_candidates` 消息契约、`captured_at_ms` 时间戳字段、宿主 HTTP endpoints、验收命令和 live snapshot 前置要求：

```js
import {
  buildMeetingAppAdapterManifestMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-manifest';

const matrix = buildMeetingAppAdapterManifestMatrix({
  baseUrl: 'http://localhost:8787',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.rows 可直接渲染为接入清单；manifest.accepted 只代表静态 SDK contract 通过。
// readiness.requires_live_snapshot_before_production 会提醒正式发布前仍需真实会议 DOM 快照验证。
```

CLI 入口会同时写出汇总报告和每个平台的 manifest JSON，适合交给 Chrome/Firefox 扩展、Electron WebView、Android WebView 或桌面 Accessibility 采集器实现：

```bash
npm run meeting-app:adapter-manifest
```

默认输出到 `data/meeting-app-adapter-manifests/` 和 `data/meeting-app-adapter-manifest-report.json`。这个清单不要求 provider webhook 或实时转写；provider 事件和 transcript 仍然只作为校准/会后处理，不能阻塞实时标注。

如果要适配一个 SDK 暂未内置的会议软件，不要先改核心归一化器，先用 `meeting-app-adapter-spec` 把新增平台的接入边界描述清楚。spec 会检查 URL match、扩展权限、控制按钮 selector、参会人/发言人 selector、MutationObserver track selector、`captured_at_ms` 时间戳、provider/transcript 非阻塞规则和必要信号：

```js
import {
  assertMeetingAppAdapterSpec,
  buildMeetingAppAdapterSpec,
  buildMeetingAppAdapterSpecTemplate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-spec';

const template = buildMeetingAppAdapterSpecTemplate({
  adapter_key: 'whereby',
  display_name: 'Whereby',
  matches: ['https://whereby.com/*'],
});

const spec = buildMeetingAppAdapterSpec({
  ...template,
  control_selectors: ['[aria-label*="Leave" i]', '[data-testid*="toolbar" i]'],
  participant_selectors: ['[data-participant-id]', '[aria-label*="speaking" i]'],
  text_selectors: ['[role="status"]', '[aria-live]'],
  mutation_track_selectors: ['[data-participant-id]', '[role="status"]'],
});

assertMeetingAppAdapterSpec(spec);
```

spec 通过只表示“这个会议软件有可实现的本地观察 contract”；它不会假装我们已经支持真实 DOM。下一步仍然要用 content script / WebView / Accessibility 采样，把 live snapshot 喂给 `meetingAppAdapterFit()`、runtime host replay 和 handoff readiness。

CLI 入口可以批量生成内置平台 spec、加载自定义 spec JSON，并输出汇总报告：

```bash
npm run meeting-app:adapter-spec
npm run meeting-app:adapter-spec -- --spec-file=data/whereby-spec.json
npm run meeting-app:adapter-spec -- --template-adapter-key=slack-huddle --template-file=data/slack-huddle-spec-template.json
```

默认输出到 `data/meeting-app-adapter-specs/` 和 `data/meeting-app-adapter-spec-report.json`。报告中的 `custom_count`、`capture_selector_ready_count`、`mutation_observer_ready_count` 可以作为新增会议软件进入真实采样前的静态 gate。

spec 通过后，用 `meeting-app-adapter-runtime-config` 把它转成真正可交给浏览器扩展、WebView preload 或 Electron content script 的 runtime options。这个转换会关闭内置 `runtimePreset`，保留自定义 selector、MutationObserver 参数、`captured_at_ms` 契约和 provider/transcript 非阻塞规则：

```js
import {
  buildMeetingAppAdapterRuntimeConfig,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-runtime-config';
import { installMeetingAppContentScriptBridge } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script';

const runtimeConfig = buildMeetingAppAdapterRuntimeConfig(spec, {
  windowMessaging: true,
  allowedOrigins: ['https://whereby.com'],
});

installMeetingAppContentScriptBridge(timeline, runtimeConfig.content_script_options);
```

CLI 入口可以把内置平台和自定义 spec 批量转成 runtime config：

```bash
npm run meeting-app:adapter-runtime-config
npm run meeting-app:adapter-runtime-config -- --spec-file=data/whereby-spec.json
```

默认输出到 `data/meeting-app-adapter-runtime-configs/` 和 `data/meeting-app-adapter-runtime-config-report.json`。报告中的 `capture_ready_count`、`mutation_ready_count`、`content_script_ready_count` 用来确认配置是否已经能交给宿主运行时，但仍不替代真实 live snapshot evidence。

如果要把 Google Meet / Teams / Zoom / Webex / Lark 或自定义会议软件交给另一个宿主项目落地，可以生成完整 handoff package：

```js
import {
  buildMeetingAppAdapterHandoffPackage,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package';

const handoffPackage = buildMeetingAppAdapterHandoffPackage('google-meet');
```

```bash
npm run meeting-app:adapter-handoff-package
npm run meeting-app:adapter-handoff-package -- --spec-file=data/whereby-spec.json
```

默认输出到 `data/meeting-app-adapter-handoff-packages/` 和 `data/meeting-app-adapter-handoff-package-report.json`。每个平台目录里包含 `adapter-spec.json`、`runtime-config.json`、`extension-manifest-fragment.json`、`verification-plan.json`、`integration-readme.md`，内置平台还会包含 `adapter-manifest.json`。这份 package 的 `validation.required_live_evidence` 和 `verification-plan.json` 明确要求真实 DOM snapshot、candidate observation、speaker/participant track 和当前轴标注插入验证。

宿主项目补齐 evidence 后，可以用 SDK 或 CLI 复验 package：

```js
import {
  buildMeetingAppAdapterVerificationReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-handoff-package';

const report = buildMeetingAppAdapterVerificationReport(handoffPackage, {
  evidence,
});
```

```bash
npm run meeting-app:adapter-verify
```

CLI 默认读取 `data/meeting-app-adapter-handoff-packages/` 和 `data/meeting-app-adapter-evidence/`，输出 `data/meeting-app-adapter-verification-report.json`。报告会给出 `static_ready`、`live_evidence_ready`、`pilot_ready`、`production_ready` 和每个缺失证据项。

面向 Google Meet、Teams、Zoom、Webex、Lark 这类多会议软件接入面板时，优先用 capability report 判断每个平台该走哪条路径：官方 provider 事件、本地/浏览器 observation、会后 transcript backfill，还是混合模式。

```js
import {
  buildMeetingAppAdapterCapabilityMatrix,
  buildMeetingAppAdapterCapabilityReport,
  buildMeetingAppAdapterExecutionPlan,
  buildMeetingAppAdapterExecutionPlanMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-adapter-capability';

const googleCapability = buildMeetingAppAdapterCapabilityReport('google-meet', {
  input: liveGoogleMeetSnapshot,
  evidence: googleLiveEvidence,
});

const googlePlan = buildMeetingAppAdapterExecutionPlan(googleCapability);

const capabilityMatrix = buildMeetingAppAdapterCapabilityMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  inputs: liveSnapshotsByPlatform,
  evidenceByAdapter,
});

const executionPlanMatrix = buildMeetingAppAdapterExecutionPlanMatrix({
  capabilityMatrix,
});
```

`recommended_mode` 会明确给出 `hybrid_local_observer_first`、`local_observer_axis_with_provider_backfill`、`provider_event_primary_with_local_snapshot_required` 或 `post_meeting_artifact_only`；`timeline_capabilities` 会分别列出 `realtime_axis`、`speaker_track`、`participant_track`、`annotation_timeline`、`post_meeting_transcript` 和 `recording` 的 provider/local 状态。

`execution_plan` 则把状态翻译成可执行步骤：`capture_meeting_app_snapshot` 负责真实会议页/窗口 observation，`start_or_reconcile_meeting_axis` 负责本地低延迟建轴或 provider 事件校准，`insert_annotation_on_current_axis` 负责按 `captured_at_ms` 落标注，`emit_speaker_position_markers` / `track_visible_participants` 负责时间轴位置轨，`configure_provider_reconcile` 和 `import_post_meeting_artifacts` 只做回填和会后处理。

```bash
npm run meeting-app:adapter-capability
```

CLI 默认输出 `data/meeting-app-adapter-capability-report.json`。如果传 `-- --input-file=live-snapshots.json --evidence-file=evidence-by-adapter.json`，报告会把现场 observation 和 verification evidence 一起纳入 `pilot_ready` / `production_ready` 判断。

同一份 `runtimeConfig.browser_runtime_options` 也可以直接传给 `createMeetingAppBrowserRuntime()`；`runtimeConfig.capture_options` 可以直接传给 `captureMeetingAppDomSnapshot()` 做手动采样。也就是说，新会议软件的接入路径是 `adapter spec -> runtime config -> handoff package -> verification report -> capability matrix -> handoff readiness`。

如果宿主不想自己解释 `trigger_policy`，可以直接用 `meeting-app-observer-scheduler`。它消费 observer plan 和现有 runtime，把 DOM mutation、native snapshot change、keep-alive、active speaker follow-up、candidate missing end grace 统一映射为 `runtime.sample()` / `runtime.sampleTracks()` 调用：

```js
import { createMeetingAppBrowserRuntime } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime';
import { createMeetingAppObserverScheduler } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-observer-scheduler';

const runtime = createMeetingAppBrowserRuntime(timeline, {
  window,
  runtimePreset: 'google_meet',
});
const scheduler = createMeetingAppObserverScheduler(runtime, observerPlan, {
  runtimeInputMode: 'provider',
});

await scheduler.triggerChanged();
scheduler.triggerCandidateMissing(undefined, { schedule: true });
```

更推荐给外部项目的入口是 `meeting-platform-runtime-host`。它把 runtime bundle、browser runtime、observer scheduler 和宿主侧 MutationObserver/lifecycle 组合好，宿主只需要提供 timeline client 和平台名：

```js
import { createMeetingTimelineClient } from '@ai-annotation/meeting-timeline-sdk';
import { createMeetingPlatformRuntimeHost } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host';

const timeline = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const host = createMeetingPlatformRuntimeHost(timeline, 'google-meet', {
  window,
  document,
});

host.start(); // 安装 change observer + keep-alive scheduler。
host.changed(); // 宿主也可以从自己的 DOM/AX/native 变化回调显式触发。
host.candidateMissing(); // 当前会议页面/窗口消失时，按 grace window 补 meeting_ended。
```

同一个 API 也适用于 Teams / Zoom / Webex / Lark。平台差异主要落在 runtime bundle 的 selector、host permission、cadence 和 provider reconcile 描述里；实时标注仍然坚持本地观察优先，provider event 和 transcript 只做后处理校准。

如果要把接入任务交给另一个项目，可以生成 handoff 包：

```js
const handoff = kit.platformRuntimeHostHandoff('google-meet');

// handoff.package_entry 指向 SDK import。
// handoff.host_hooks 列出宿主需要转发的 surface open/change/missing/close 钩子。
// handoff.acceptance.checks 是接入前的最小验收条件。
```

也可以直接导出给外部项目：

```bash
npm run meeting-platform:runtime-host-handoff
```

默认会为 Google Meet / Teams / Zoom / Webex / Lark 生成总报告和每个平台的 handoff JSON。

接入前可以用 `meeting-platform-runtime-host-verifier` 做 SDK 闭环验收。它用平台 fixture 构造一个可切换的 browser runtime 环境，驱动真实 runtime host 写入 `meeting_started`、`speaker_started` 和 `meeting_ended`，确认 timestamp 没有漂移，并且默认把 speaker track 写成标记：

```js
import { runMeetingPlatformRuntimeHostVerificationMatrix } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-platform-runtime-host-verifier';

const matrix = await runMeetingPlatformRuntimeHostVerificationMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.accepted_count === matrix.platform_count 表示 SDK 接入 contract 闭环通过。
```

也可以直接导出报告：

```bash
npm run meeting-platform:runtime-host-verify
```

这个验证器不替代真实会议软件的现场 evidence capture；它的作用是证明外部项目拿到 SDK 后，runtime host、browser runtime、observer scheduler、speaker marker 和 timeline client 的最小链路没有断。

浏览器扩展或 WebView 里可以再往前接一层 `meeting-app-capture`。它只读取 DOM 文本、按钮、`aria-label`、participant tile、常见 `data-participant-*` / `data-user-*` / `data-person-*` 属性和音量/发言状态，输出 `meeting-apps` 可识别的快照；不截图、不 OCR、不读取转写正文：

```js
import { captureMeetingAppDomSnapshot } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-capture';
import { createMeetingSourceAggregator } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-source';

const meetingSources = createMeetingSourceAggregator(timeline, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 300, switchStableMs: 400, endIdleMs: 1500 },
});

const snapshot = captureMeetingAppDomSnapshot({ document, location, window }, {
  observedAtMs: Date.now(),
  browserName: 'Chrome',
  platform: 'google_meet', // 可省略；SDK 会尽量从 URL/title 自动选择 Google Meet/Teams/Zoom/Lark/Webex profile。
  includeShadowDom: true, // 仅在目标页面使用 open shadow root 时打开；默认关闭以控制采样成本。
});

await meetingSources.observeMeetingApp(snapshot, { observedAtMs: snapshot.observedAtMs });
```

这条链路是 Google Meet / Teams Web / Zoom Web / Webex Web / Lark/Feishu Web 的推荐 P0 接入：先用本地 DOM 状态低延迟建轴和标发言人位置；Google Workspace Events、Microsoft Graph、Zoom/Webex/Lark webhook 晚到后再进入 provider adapter 做 reconcile。

外部项目优先使用更高层的 `meeting-app-runtime`，它把 client、`meeting-source` 和 DOM monitor 组合好，适合浏览器扩展 content script、Electron WebView 或内嵌浏览器宿主直接接入：

```js
import { createMeetingAppTimelineRuntime } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-runtime';
import { resolveMeetingAppRuntimeAdapterProfile } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';

const profile = resolveMeetingAppRuntimeAdapterProfile({ document, location, window });
if (!profile.detected) throw new Error('当前页面不是已支持的会议应用');

const runtime = createMeetingAppTimelineRuntime({
  baseUrl: 'http://localhost:8787',
  source: 'browser_dom_runtime',
}, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 300, switchStableMs: 400, endIdleMs: 1500 },
  captureOptions: { browserName: 'Chrome', ...profile.capture.options },
  trackRuntimeOptions: profile.tracks.runtime_options,
  sampleIntervalMs: 1000,
  unchangedObserveEveryMs: 1000,
});

runtime.start(() => ({ document, location, window }));
runtime.startTracks(() => ({ document, location, window }));

// 如果宿主自己已经采好了快照，也可以直接走同一个 runtime 的 track 入口。
await runtime.observeMeetingAppTracks(domSnapshots, {
  speakerTrackOptions: { minStableMs: 250, endIdleMs: 500 },
  participantTrackOptions: { leaveStableMs: 500 },
});

await runtime.insertMark({
  id: crypto.randomUUID(),
  capturedAtMs: Date.now(),
  kind: 'handwriting_trigger',
  label: 'why?',
});

// 官方 provider 事件晚到后仍可进入同一个 reconciler 校准：
await runtime.ingestProvider('google-meet', googleWorkspaceEventBody);
```

如果宿主已经统一使用 `platform-kit`，同一个 profile 可以从 `kit.meetingAppRuntimeAdapterProfile(input)` 获取，避免下游项目同时依赖多个底层模块。多平台接入面板可以用 `kit.meetingAppRuntimeAdapterProfileMatrix({ platforms: ['google-meet', 'teams', 'zoom'] })` 一次拿到每个平台的 extension matches、capture profile、runtime preset 和 speaker/participant track 默认阈值。宿主需要从浏览器 tabs、Electron windows 或本地 app 快照里自动选择当前会议时，用 `kit.selectMeetingAppRuntimeAdapter(input)`：

```js
const selection = kit.selectMeetingAppRuntimeAdapter({
  tabs: [
    { active: false, url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting' },
    { active: true, audible: true, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' },
  ],
}, {
  platforms: ['google-meet', 'teams', 'zoom'],
});

const handoff = kit.meetingAppRuntimeAdapterHandoff(selection, {
  surface: 'browser-extension',
  baseUrl: 'http://localhost:8787',
});

if (handoff.readiness.ready_to_start) {
  const runtime = createMeetingAppTimelineRuntime({
    baseUrl: 'http://localhost:8787',
    ...handoff.runtime.options,
  }, {
    captureOptions: handoff.runtime.capture_options,
    trackRuntimeOptions: handoff.tracks.options,
  });
  runtime.start(() => ({ document, location, window }));
}
```

`meetingAppRuntimeAdapterHandoff()` 是推荐给宿主项目保存/传递的边界对象：`surface` 可选 `browser-extension`、`electron-webview`、`webview`、`native-detector`，输出包含安装目标、权限、runtime options、capture options、speaker/participant track options、时间戳字段和非阻塞规则。

如果宿主要先评估多个会议软件和多个接入面，可以生成 handoff matrix：

```js
const matrix = kit.meetingAppRuntimeAdapterHandoffMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex'],
  surfaces: ['browser-extension', 'electron-webview', 'native-detector'],
});

console.table(matrix.rows.map((row) => ({
  platform: row.platform,
  surface: row.surface,
  ready: row.ready_to_start,
  install: row.install_target,
})));

const acceptance = kit.meetingAppRuntimeAdapterHandoffMatrixAcceptance(matrix);
if (!acceptance.accepted) throw new Error(acceptance.issues.map((item) => item.code).join(', '));
```

`meetingAppRuntimeAdapterHandoffAcceptance()` 和 `meetingAppRuntimeAdapterHandoffMatrixAcceptance()` 可以作为宿主项目的 CI gate：它们会检查 `captured_at_ms`、非阻塞 provider/transcript、权限范围、surface 安装能力、runtime/track options 是否完整。没有真实 DOM 快照前会保留 `production_requires_live_snapshot` warning，但不阻塞试点接入。

如果要把这套配置交给另一个宿主项目，而不是让对方重新拼 matrix、acceptance、SDK import 和上线检查，可以直接生成 host package：

```js
const hostPackage = kit.meetingAppRuntimeAdapterHostPackage({
  platforms: ['google-meet', 'teams', 'zoom'],
  surfaces: ['browser-extension', 'electron-webview', 'native-detector'],
  primarySurface: 'browser-extension',
});

if (!hostPackage.accepted) {
  throw new Error(hostPackage.next_actions.join(', '));
}
```

`hostPackage` 会同时带上 `handoff_matrix`、`handoff_acceptance`、推荐 SDK imports、宿主入口方法、CI gates、实时标注时间字段、provider/transcript 非阻塞规则和 rollout checklist。外部项目可以把它当成机器可读交付物，先接 `adapter-selection -> runtime-handoff -> handoff-ci-gate` 三个入口，后续再补真实 DOM 采样证据。

如果是在浏览器扩展 content script、内嵌浏览器或 Electron WebView 里运行，可以用 `meeting-app-content-script` 直接安装浏览器侧 bridge。它会创建 `meeting-app-browser-runtime`，自动读取当前 `document/location/window`，安装扩展消息监听，并把 background script 或宿主转发来的标注消息写入时间轴：

```js
import { installMeetingAppContentScriptBridge } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script';

const bridge = installMeetingAppContentScriptBridge({
  baseUrl: 'http://localhost:8787',
}, {
  runtimePreset: 'google_meet',
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 300, switchStableMs: 400, endIdleMs: 1500 },
  sampleIntervalMs: 10000, // MutationObserver 负责低延迟触发，低频轮询只做兜底。
});

// background script 发送 { type: 'meeting_timeline.insert_mark', payload: { mark } } 即可。
// 页面卸载或扩展停用时调用 bridge.dispose() 清理 listener 和 monitor。
```

如果宿主不是 Chrome/Firefox 扩展，而是 WebView 或页面内脚本，可以打开 `windowMessaging`，用 `window.postMessage()` 投递同样的 `meeting_timeline.*` 消息；默认只接收同源消息，避免误消费页面里的其他事件：

```js
const bridge = installMeetingAppContentScriptBridge({ baseUrl: 'http://localhost:8787' }, {
  runtimePreset: 'microsoft_teams',
  windowMessaging: true,
  extensionMessaging: false,
});
```

底层仍可直接使用 `meeting-app-browser-runtime`；它提供 `handleMessage()`、`installMutationObserver()`、`flushMutationObserver()` 等低层控制点，适合需要自定义扩展生命周期的项目。

`runtimePreset` 当前支持 `google_meet`、`microsoft_teams`、`zoom`、`lark` 和 `webex`。preset 会自动合并对应 `meeting-app-capture` DOM profile、MutationObserver track selectors、字幕/聊天/转写噪声过滤、发言人稳定 follow-up 和低频兜底轮询；调用方仍然可以覆盖 `mutationTrackSelectors`、`mutationIgnoreSelectors`、`mutationDebounceMs`、`speakerStableFollowupMs`、`captureOptions` 等字段。也可以单独读取预设用于浏览器扩展配置面板：

```js
import {
  MEETING_APP_BROWSER_RUNTIME_PRESETS,
  meetingAppBrowserRuntimePreset,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime';

const preset = meetingAppBrowserRuntimePreset('microsoft-teams');
// preset.captureOptions.platform === 'microsoft_teams'
// MEETING_APP_BROWSER_RUNTIME_PRESETS.google_meet.observeMutations === true
```

接入浏览器扩展、Electron WebView 或桌面 Accessibility 采集器前，可以先跑 `meeting-app-fixtures` 的本地验收样本。它覆盖 Google Meet、Teams Web、Zoom Web、Webex Web、Lark/Feishu Web，并验证平台识别、会议 ID、入会态、active speaker、`meeting_started` 和 `speaker_started`：

```js
import { buildMeetingAppFixtureAcceptanceReport } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixtures';

const report = buildMeetingAppFixtureAcceptanceReport();
// report.accepted === true 表示本地 DOM/AX 归一化链路基线通过
```

如果还要验证这些 fixture 是否足以驱动时间轴上的发言人轨和参会人轨，用 `meeting-app-fixture-tracks`。它会把 fixture 的 active speaker signal 喂给 `platform-speaker-track`，把 roster 变化喂给 `platform-participant-track`，确认五个平台都能生成不依赖 transcript 文本的 `speaker_track` / `participant_track` marks：

```js
import {
  buildMeetingAppFixtureTrackReadinessReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-fixture-tracks';

const trackReport = buildMeetingAppFixtureTrackReadinessReport();
// trackReport.accepted === true 表示 fixture 级轨道连通性通过
```

真实网页或 WebView 已经能持续采样时，用 `meeting-app-track-pipeline` 把快照序列直接转换成发言人轨、参会人轨和可插入时间轴的 marks。它只依赖本地页面观察结果，不要求 provider webhook，也不要求实时转写：

```js
import {
  buildMeetingAppTrackPipeline,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-pipeline';

const pipeline = buildMeetingAppTrackPipeline(domSnapshots, {
  speakerTrackOptions: {
    minStableMs: 250,
    endIdleMs: 500,
  },
  participantTrackOptions: {
    leaveStableMs: 500,
  },
});

// pipeline.marks 可以批量插到当前会议轴；mark.intent 会是 speaker_track / participant_track
// pipeline.coverage.transcript_required === false
// pipeline.coverage.provider_event_required === false
```

如果宿主已经在循环采样真实会议页，可以直接用 `meeting-app-track-runtime` 管增量插入和去重。runtime 会保留最近一段快照，重复 observe 同一批快照不会重复写入同一个 mark；默认优先调用 `insertMarks` 批量接口，缺少批量接口时退回逐条 `insertMark`：

```js
import {
  createMeetingTimelineClient,
} from '@ai-annotation/meeting-timeline-sdk';
import {
  createMeetingAppTrackRuntime,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-track-runtime';

const client = createMeetingTimelineClient({ baseUrl: 'http://localhost:8787' });
const runtime = createMeetingAppTrackRuntime(client, {
  speakerTrackOptions: { minStableMs: 250, endIdleMs: 500 },
  participantTrackOptions: { leaveStableMs: 500 },
});

await runtime.observe(domSnapshot);
// 返回值里的 new_marks 是这次新增并已写入的 speaker_track / participant_track marks
```

正式接入 Google Meet / Teams Web / Zoom Web / Webex Web / Lark Web 前，建议再跑 `meeting-app-gate`。它不检查官方 webhook 权限，而是检查本地会议 App 路径是否满足实时标注：browser runtime preset、DOM capture profile、MutationObserver track/ignore selectors、建轴、发言人位置和结束信号。fixture-only 只能证明 SDK wiring；要证明生产可用，需要用 `meeting-app-snapshot-recorder` 记录真实采集到的 DOM snapshots：

```js
import {
  buildMeetingAppLaunchGate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate';
import {
  createMeetingAppSnapshotRecorder,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder';

const fixtureGate = buildMeetingAppLaunchGate('google-meet', {
  allowFixtureProduction: true,
  requireProductionReady: false,
});
// fixtureGate.status === 'warning'

const recorder = createMeetingAppSnapshotRecorder({ captureProfile: 'google_meet' });
recorder.capture({ document, location, window }, { phase: 'active', label: 'live-active' });
// 用户离开会议或回到 prejoin 页面后再采一帧：
recorder.capture({ document, location, window }, { phase: 'ended', label: 'live-ended' });

const liveGate = buildMeetingAppLaunchGate('google-meet', {
  recordSet: recorder.exportRecords(),
});
// liveGate.production_ready === true 时，才表示真实 Google Meet DOM 适配已验收。
```

如果要同时看 Google Meet / Teams / Zoom / Webex / Lark 的真实快照适配情况，可以用 DOM adaptation diagnosis matrix。它会把 selector 命中、建轴、发言人和结束态拆成独立字段，方便判断是 URL/selector 问题，还是 observer 状态转换问题：

```js
import {
  buildMeetingAppDomAdaptationDiagnosisMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-profile';

const diagnosis = buildMeetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  snapshots: capturedLiveMeetingSnapshotsByPlatform,
});

// diagnosis.rows[*].controls_matched / participants_matched / active_speaker_matched
// 分别对应 toolbar、参会人/发言人选择器是否适配真实页面。
// diagnosis.rows[*].meeting_started / speaker_started / meeting_ended
// 对应本地 observer 是否能产生时间轴 start、speaker marker 和 end。
```

如果宿主项目只需要在当前会议窗口启动前做一次低成本检查，用 `platform-adapter-preflight`。它会先跑 URL/title/platform/native context 的 startup plan，再按 surface 选择证据诊断：browser extension / WebView 用 live DOM snapshots，native detector 用窗口、进程、Accessibility 或音频通话状态。最后统一压成 `static_startup_ready`、`live_evidence_ready`、`meeting_start_ready`、`speaker_track_ready`、`meeting_end_ready` 和 `realtime_annotation_ready`。没有真实 DOM 或 native-window 证据时不会宣称 realtime ready；provider events 和 transcript 都只作为非阻塞回填。

```js
import {
  buildMeetingPlatformAdapterCandidatePreflight,
  buildMeetingPlatformAdapterCurrentWindowPreflight,
  buildMeetingPlatformAdapterPreflight,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-adapter-preflight';

const preflight = buildMeetingPlatformAdapterPreflight({
  url: 'https://meet.google.com/abc-defg-hij',
  snapshots: currentWindowLiveSnapshots,
}, {
  requireSpeakerTrack: true,
});

// preflight.status === 'ready_for_realtime_annotations'
// preflight.readiness.realtime_annotation_ready === true
// preflight.readiness.meeting_end_ready 可以单独提示“结束态还没验证”。

const nativePreflight = buildMeetingPlatformAdapterPreflight({
  platform: 'zoom',
  window: { title: 'Zoom Meeting', active: true, inMeeting: true, meeting_id: '987654321' },
  process: { name: 'Zoom Workplace' },
  audio: { call_active: true },
});
// nativePreflight.summary.evidence_kind === 'native_window'
// nativePreflight.summary.selected_surface === 'native_detector'

const currentWindowPreflight = buildMeetingPlatformAdapterCurrentWindowPreflight({
  window,
  document,
}, {
  requireSpeakerTrack: true,
});
// content script / WebView preload 可以直接用这个入口，SDK 会先 capture 当前 DOM 再 preflight。

// 如果已经安装 connector bridge，也可以通过消息触发同一件事：
const response = await bridge.dispatchMessage({
  type: 'meeting_timeline.preflight_current_window',
  payload: { options: { requireSpeakerTrack: true } },
});
// response.result.accepted 表示当前窗口是否能作为实时标注轴来源。
```

如果宿主拿到的是浏览器扩展 background、Electron preload 或桌面观察器的一批候选窗口/标签，用 `buildMeetingPlatformAdapterCandidatePreflight()` 一次性展开 `windows[].tabs[]` / `tabs[]` / `candidates[]`。URL-only 候选只会验证 startup plan 并返回 `needs_live_page_evidence`；浏览器候选需要明确携带当前 `document` 或 live snapshot，native 候选需要携带 platform/process/window/call state，才会被判定为 `ready_for_realtime_annotations`，避免误把静态 URL 匹配当成可实时建轴。候选不是简单按输入顺序选择：SDK 会按已通过实时预检、当前窗口/live DOM/native evidence、活跃标签或焦点窗口、meeting start 和 speaker track 证据打分排序，并在 `rows[*].selection_score`、`rows[*].selection_rank`、`rows[*].selection_reason` 里保留原因，方便 popup、native helper 或接入面板解释为什么选择某一个 Google Meet / Teams / Zoom / Webex / Lark 窗口。

```js
const candidatePreflight = buildMeetingPlatformAdapterCandidatePreflight({
  windows: [{
    id: 'chrome-main',
    tabs: [
      { id: 7, url: 'https://meet.google.com/abc-defg-hij', title: 'Google Meet' },
      { id: 8, url: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting', title: 'Teams' },
      { id: 9, url: location.href, title: document.title, document },
    ],
  }],
}, {
  requireSpeakerTrack: true,
});

// candidatePreflight.rows[*] 会保留 window_id / tab_id / url / platform / status / selection_score。
// candidatePreflight.selected_platform 是当前最适合建实时标注轴的平台。
// candidatePreflight.accepted === true 表示至少一个候选已经可实时插入 captured_at_ms 标注。
```

如果已经有 `platformAdapterInstallManifest()`，可以进一步用 `platformAdapterCandidateLaunchPlan()` 把“候选窗口验收”和“启动 runtime surface”合成一步。这个对象的 `accepted` 同时要求候选 preflight 通过、install manifest 可用、平台已注册、surface 可启动；通过后把内部的 `launch_plan` 交给 `platformAdapterSession()` 或 runner 即可。

```js
const launch = sdk.platformAdapterCandidateLaunchPlan(installManifest, {
  candidates: [{
    url: 'https://meet.google.com/abc-defg-hij',
    snapshots: currentWindowLiveSnapshots,
  }],
}, {
  requireSpeakerTrack: true,
});

if (launch.accepted) {
  const session = sdk.platformAdapterSession(launch.launch_plan);
  await session.observeAxis({
    candidates: [launch.selected_candidate],
    captured_at_ms: Date.now(),
  });
}
```

如果希望 SDK 帮你管理轮询、去重和 keep-alive，可以直接用 `meeting-app-monitor`。它会高频低成本采集 DOM，但只有在页面状态变化、或到达 keep-alive 间隔时才把样本送给 `meeting-source`；即使 DOM 不变，也会按间隔继续送样本，避免 active speaker 的 `minStableMs` 因过度去重而无法触发：

```js
import { createMeetingAppDomMonitor } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-monitor';

const monitor = createMeetingAppDomMonitor(meetingSources, {
  sampleIntervalMs: 1000,
  minObserveIntervalMs: 250,
  unchangedObserveEveryMs: 1000,
  captureOptions: { browserName: 'Chrome' },
});

monitor.start(() => ({ document, location, window }));

// tab hidden, user leaves the page, or extension unloads:
monitor.stop();
```

桌面客户端推荐用 `native-meeting`。它面向 macOS Accessibility、Windows UI Automation、Electron shell 或宿主进程采集到的 app/window/process/audio 快照；适合 Zoom、Teams、Lark/Feishu、Webex 桌面端：

```js
import { createNativeMeetingTimelineObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/native-meeting';

const nativeMeetings = createNativeMeetingTimelineObserver(timeline, {
  source: 'native_desktop_observer',
  speakerOptions: {
    minStableMs: 300,
    switchStableMs: 400,
    endIdleMs: 1500,
  },
  applyOptions: { speakerAsAnnotation: true },
});

await nativeMeetings.observe({
  applications: [{
    name: 'zoom.us',
    bundleId: 'us.zoom.xos',
    windows: [{
      id: 'zoom-call-1',
      title: 'Daily Standup - Zoom Meeting',
      focused: true,
      inMeeting: true,
      accessibility: {
        callActive: true,
        activeSpeaker: { id: 'speaker-ada', name: 'Ada', speaking: true },
      },
    }],
  }],
  observedAtMs: Date.now(),
});
```

没有 URL 的 native app 会优先使用显式 `meetingId`；如果宿主拿不到会议 ID，SDK 会基于平台、会议窗口标题或窗口 ID 生成 `native-{platform}-...` 会话 ID。这个 ID 可以先保证实时标注落轴，后续官方 webhook 或会后转写到达后再做 reconcile。

如果一个项目同时接浏览器扩展、桌面客户端观察器和官方 webhook，推荐用 `meeting-source` 作为总入口。它给 browser/native/local/provider 共用同一个 reconciler，避免重复建轴，并允许官方事件晚到后校准本地轴：

```js
import { createMeetingSourceAggregator } from '@ai-annotation/meeting-timeline-sdk/adapters/meeting-source';

const meetingSources = createMeetingSourceAggregator(timeline, {
  applyOptions: { speakerAsAnnotation: true },
  speakerOptions: { minStableMs: 300, switchStableMs: 400, endIdleMs: 1500 },
  reconcileOptions: { duplicateWindowMs: 60_000 },
});

await meetingSources.observeBrowser(browserSnapshot, { observedAtMs: Date.now() });
await meetingSources.observeNative(nativeAppSnapshot, { observedAtMs: Date.now() });
await meetingSources.observeMeetingApp(meetingAppDomOrAxSnapshot, { observedAtMs: Date.now() });
await meetingSources.ingestProvider('google-meet', googleWorkspaceEventBody);

await meetingSources.insertMark({
  capturedAtMs: Date.now(),
  kind: 'handwriting_trigger',
  label: 'why?',
  intent: 'question',
});
```

`meeting-source` 的返回值会带 `rawSignals`、`signals`、`reconciliation.skipped` 和 `diagnostic`，方便在接入现场判断是“未识别出会议窗口”、还是“被 provider 事件接管后本地重复信号被跳过”。

如果业务项目要接入多个会议平台，推荐从 `platform-kit` 开始。它把 `timeline-bridge`、webhook router、平台 setup/onboarding、provider fixture acceptance 和本地 meeting-app fixture acceptance 组合成一个入口；底层 normalizer、验签、artifact fetch 仍然可以按需单独 import：

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

const localAppReport = meetingKit.meetingAppFixtureAcceptance();
// localAppReport.accepted === true 表示五个平台的本地 DOM/AX 建轴和发言人标注基线通过

const appGate = meetingKit.meetingAppLaunchGate('google-meet', {
  allowFixtureProduction: true,
  requireProductionReady: false,
});
// appGate.runtime_ready === true 表示 browser runtime preset、capture profile 和 mutation 配置齐备。

const appProfile = meetingKit.meetingAppIntegrationProfile('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
// appProfile 聚合了这个平台的 extension host 权限、DOM 采集 selectors、browser runtime preset、
// 本地实时事件模型、官方 provider 事件回填策略和 launch gate readiness。

const appMatrix = meetingKit.meetingAppIntegrationMatrix();
// appMatrix.rows 可以直接用于产品侧展示 Google Meet / Teams / Zoom / Lark / Webex 的接入状态。

const deploymentManifest = meetingKit.meetingAppDeploymentManifest('google-meet', {
  baseUrl: 'https://timeline.example.com',
});
// deploymentManifest 是给外部项目消费的一页式接入契约：
// 它集中包含 profile、runtime_config、extension_install_plan、live_snapshot_capture_plan 和 production_gate。
// Google Meet、Teams、Zoom、Lark/Feishu、Webex 都走同一个结构，只换平台 key。

const manifestAcceptance = meetingKit.meetingAppDeploymentManifestAcceptance(deploymentManifest);
// manifestAcceptance.accepted === true 表示契约结构完整、可以交付外部项目接入；
// manifestAcceptance.production_ready === true 才表示已经用真实会议 DOM 快照通过生产 gate。

const liveEvidence = meetingKit.meetingAppLiveEvidencePackage({
  platforms: ['google-meet'],
  snapshots: capturedLiveMeetingSnapshots,
});
// liveEvidence.summary.rows 会列出每个平台的 record_count、missing_required_coverage 和 next_actions。
// 它适合放进现场采样工具或 CI，把真实 Google Meet / Teams / Zoom 页面快照变成生产 gate 结果。

const domDiagnosis = meetingKit.meetingAppDomAdaptationDiagnosis('google-meet', {
  snapshots: capturedLiveMeetingSnapshots,
});
// domDiagnosis 更适合现场调试：它会拆开 selector_probe、observer_probe 和 runtime_probe，
// 明确指出当前真实页面快照是否命中 controls / participants / active speaker / meeting ended。
// Google Meet、Teams、Zoom、Lark/Feishu、Webex 都可以用同一个诊断入口，只换平台 key。

const domDiagnosisMatrix = meetingKit.meetingAppDomAdaptationDiagnosisMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  snapshots: capturedLiveMeetingSnapshotsByPlatform,
});
// domDiagnosisMatrix.rows 可以直接做成多会议软件本地 observer 适配看板。

const runtimeConfig = meetingKit.meetingAppRuntimeAdapterConfig('google-meet');
// runtimeConfig 可以交给浏览器 extension content script、Electron WebView preload 或桌面宿主：
// 它包含 bridge_options、runtime_options、capture_options、extension host 权限和 message_types。

const runtimeAcceptance = meetingKit.meetingAppRuntimeAdapterAcceptance(runtimeConfig);
// runtimeAcceptance.accepted === true 表示配置的 host 权限、消息协议、runtime preset、capture selectors
// 和 startup 字段满足 SDK 运行前置条件；生产前仍需用真实会议 DOM 快照通过 launch gate。

const capturePlan = meetingKit.meetingAppLiveSnapshotCapturePlan('google-meet');
// capturePlan.required_snapshots 描述现场工具至少要采 active speaker 与 meeting ended 两类快照。

const runtimeValidation = meetingKit.meetingAppRuntimeAdapterValidation(runtimeConfig, {
  snapshots: capturedLiveMeetingSnapshots,
});
// runtimeValidation.production_ready === true 才表示该平台的 runtime config 已经被真实会议 DOM 快照证明。

const extensionPlan = meetingKit.meetingAppExtensionInstallPlan({
  platforms: ['google-meet', 'microsoft-teams'],
  js: ['meeting-app-content-script.bundle.js'],
});
// extensionPlan.manifest.content_scripts[0].matches 是 Google Meet / Teams 的白名单注入规则。
// extensionPlan.content_script_adapter 指向 platform-integration-runtime，低层 meeting-app-content-script 仍作为内部兼容模块保留。
// 默认也支持 Zoom、Lark/Feishu、Webex，且不会生成 <all_urls> 这种过宽权限。
// 对完整 scaffold/runtime profile，extension.permissions 会包含 storage + tabs；tabs 用于 background 观察当前浏览器候选会议窗口。

const extensionScaffold = meetingKit.meetingAppExtensionScaffold({
  platforms: ['google-meet'],
  baseUrl: 'https://timeline.example.com',
});
// extensionScaffold.files 包含 package.json、build.mjs、manifest.json、src/content-script.entry.mjs、src/background.entry.mjs 和 README.md。
// content script 入口用平台级 SDK bridge 监听会议网页；background worker 把 start/end/mark 调用转发到 timeline 服务。
// content script 注入后会发送 meeting_timeline.extension_attached；background 可用 meeting_timeline.extension_status 查询最近注入状态。
// popup/native host 也可以发 meeting_timeline.observe_candidates，让 background 查询 tabs 并投递 observe_platform_candidates runtime event。
// scaffold 默认还会生成 src/live-capture.entry.mjs，在页面上暴露 window.__meetingTimelineLiveCapture。
// 现场验证时可在真实 Google Meet / Teams / Zoom 页面调用 captureActive()、captureEnded()、evidencePackage()、diagnose()。

const attachedMessage = meetingKit.meetingAppExtensionAttachedMessage({
  platform: 'google-meet',
  capturedAtMs: Date.now(),
  url: 'https://meet.google.com/abc-defg-hij',
});

const markMessage = meetingKit.meetingAppExtensionClientCallMessage('insertMark', {
  label: 'why?',
  captured_at_ms: Date.now(),
}, {
  platform: 'google-meet',
});
// 业务项目接入不同会议软件时，优先复用这些 helper，而不是手写 meeting_timeline.* 字符串。

const extensionAcceptance = meetingKit.meetingAppExtensionAcceptance({
  platforms: ['google-meet'],
  baseUrl: 'https://timeline.example.com',
});
// extensionAcceptance.accepted === true 表示 manifest、权限、构建入口、SDK bridge 和平台域名覆盖都通过基础验收。

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

在 Node `http` 或 Express/Koa 类服务里，可以用 `platform-node` 直接接 `IncomingMessage`/`ServerResponse`。它会读取请求流并保留 raw body 给 Zoom/Webex 签名校验：

```js
import { createServer } from 'node:http';
import { createMeetingPlatformNodeHandler } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-node';

const handleNodeWebhook = createMeetingPlatformNodeHandler(timeline, {
  basePath: '/api/platform-events',
  verify: true,
  reconcile: true,
});

createServer((req, res) => {
  handleNodeWebhook(req, res).catch((error) => {
    res.statusCode = 500;
    res.end(error.message);
  });
}).listen(8788);
```

Express 风格可以直接挂 middleware：

```js
import { createMeetingPlatformExpressMiddleware } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-node';

app.use('/api/platform-events', createMeetingPlatformExpressMiddleware(timeline, {
  basePath: '/api/platform-events',
  verify: true,
  reconcile: true,
}));
```

Local detector adapter 支持桌面观察器、浏览器扩展、汉王宿主 App 或人工控制器上报 `meeting_started` / `meeting_ended`，payload 可以带 `detected_platform: 'google_meet'` 或 `meeting.platform: 'zoom'` 表示真实会议来源；如果只传窗口 URL，SDK 会用 `meeting-url` 自动识别 Google Meet / Teams / Zoom / Lark / Webex 的平台和会议 ID。Local detector 也支持 `active_speaker`、`speaker_started`、`speaker_ended` 这类发言人信号，用来低延迟标出发言人位置。会议 App fixture 支持 `state: 'active'` 和 `state: 'prejoin'`，`diagnoseMeetingAppFixtureLifecycle()` 会验证同一平台从入会、发言人到离会是否都能产生 timeline 信号。Lark adapter 支持 `vc.meeting.all_meeting_started_v1`、`vc.meeting.all_meeting_ended_v1`、`vc.meeting.meeting_started_v1`、`vc.meeting.meeting_ended_v1`、`vc.meeting.join_meeting_v1`、`vc.meeting.leave_meeting_v1`，并会把 `minute_token` 透传到会议轴，便于会后妙记导入。Google Meet adapter 同时支持已经解包的 Workspace Events CloudEvent，以及 Pub/Sub 默认 wrapped push body。wrapped body 会自动 base64 解码 `message.data`，所以 webhook handler 可以直接把 `req.body` 传给 `normalizeGoogleMeetEvent(req.body)`。Google Workspace Events 的 `subscription.v1.suspended`、`subscription.v1.expirationReminder`、`subscription.v1.expired` 会归一化为 `subscription_lifecycle`。Microsoft Graph change notifications 的 `lifecycleEvent` 值 `reauthorizationRequired`、`subscriptionRemoved`、`missed` 也会归一化为 `subscription_lifecycle`。Webex adapter 支持 `meetings` started/ended、`meetingParticipants` joined/left、`recordings` created/updated、`meetingTranscripts` created。

宿主服务如果要按平台名动态接 webhook，可以直接用 registry：

```js
import {
  assertMeetingPlatformRegistryManifest,
  buildMeetingPlatformRegistryManifest,
  meetingPlatformEventAdapterFor,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-registry';

const adapter = meetingPlatformEventAdapterFor(req.params.platform);
const signals = adapter.normalize(req.body, { receivedAtMs: Date.now() });

const manifest = buildMeetingPlatformRegistryManifest({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// manifest.rows 是下游项目的选型表：normalizer、runtime、candidate observation、provider、insert endpoint、runtime action 数和非阻塞规则。
// manifest.candidate_observer_count 必须等于 manifest.platform_count，才表示每个平台都能用 observe-candidates 做 host-level 建轴。
// manifest.entries[*] 进一步包含 SDK import、runtime bundle、runtime event plan、provider security verifier 和 host endpoints。
assertMeetingPlatformRegistryManifest(manifest);
```

也可以直接导出 registry 报告：

```sh
npm run meeting-platform:registry -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --report-file=data/meeting-platform-registry-report.json
```

如果要在下游项目接入前做一层更严格的 SDK 静态验收，用 `platform-conformance`。它会把 normalizer、setup manifest、capability contract、adapter contract、adapter route、runtime bundle、registry、candidate observation、`captured_at_ms` 和 provider/transcript 非阻塞规则合成一份 `meeting_platform_conformance_report`；`accepted=true` 只说明 SDK 调用面和实时策略可交付，production 仍以后续真实 DOM/provider evidence package 和 handoff readiness 为准：

```js
import {
  assertMeetingPlatformConformanceReport,
  buildMeetingPlatformConformanceReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-conformance';

const conformanceOptions = {
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
};

const conformance = buildMeetingPlatformConformanceReport(conformanceOptions);
assertMeetingPlatformConformanceReport(conformanceOptions);
```

```sh
npm run meeting-platform:conformance -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --report-file=data/meeting-platform-conformance-report.json
```

如果要把 SDK 交给另一个项目接入，优先导出 `platform-consumer-handoff`。它不是单个平台 demo，而是一份机器可读的 consumer index：包含 Google Meet / Teams / Zoom / Webex / Lark 的 SDK module、kit/host 方法、HTTP endpoints、CLI 命令、启动顺序、硬契约、每个平台的静态 readiness、adapter blueprint、adapter startup plan、adapter preflight live-evidence gate，以及 production evidence/handoff readiness 是否还缺。默认只把静态 SDK 接入门禁作为硬失败；如果需要正式交付验收，可以打开 `requireHandoffReady` 或 `requireProductionReady`：

```js
import {
  assertMeetingPlatformConsumerHandoff,
  buildMeetingPlatformConsumerHandoff,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-consumer-handoff';

const handoff = buildMeetingPlatformConsumerHandoff({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

assertMeetingPlatformConsumerHandoff(handoff);
// handoff.entrypoints.http_endpoints.consumer_handoff
// handoff.hard_contracts.timestamp_field === 'captured_at_ms'
// handoff.adapter_startup_plan_matrix.rows[*].selected_surface 告诉宿主先启动 browser/native/webview 哪个 surface。
// handoff.adapter_preflight_matrix.rows[*].status === 'needs_live_page_evidence' 表示还没传真实 DOM/native evidence，不能写实时标注。
// handoff.rows[*].adapter_startup_insert_action === 'insertAnnotation'
// handoff.rows[*].production_ready 用来提示还缺哪些真实会议证据。
```

```sh
npm run meeting-platform:consumer-handoff -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --report-file=data/meeting-platform-consumer-handoff-report.json
```

如果下游项目是直接安装 SDK 包，也可以不依赖本仓库的 npm script，直接运行包内 CLI：

```sh
npx meeting-platform-consumer-handoff \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --report-file=data/meeting-platform-consumer-handoff-report.json
```

`consumer-handoff` 会把 `meeting-platform:adapter-startup` 和 `meeting-platform:adapter-preflight` 都列入 `entrypoints.commands` 和 `boot_order`。因此下游项目不需要自己在 `platform-adapter-blueprint`、`platform-adapter-startup`、`platform-adapter-preflight`、`platform-host-integration` 之间猜调用顺序：先读 blueprint 确认可接入面，再读 startup plan 启动本地实时 surface，然后用真实 DOM snapshot、当前窗口列表、进程/Accessibility/音频状态跑 preflight；只有 `adapter_preflight_matrix.rows[*].realtime_annotation_ready === true` 后，才按 `observePlatformCandidates` → `insertAnnotation` 的顺序写入 `captured_at_ms` 标注。URL-only preflight 会稳定返回 `needs_live_page_evidence`，不会误报可实时写入。provider webhook 和转写仍然只做 reconcile/backfill，不作为实时标注前置条件。

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

如果另一个项目还没接好官方 webhook，但需要先验证“provider 原始事件能不能进入统一 runtime contract”，可以用 provider replay。它不会插入用户标注，也不会把 provider 当成实时落标注前置条件，只把 Google Meet / Teams / Zoom / Webex / Lark 的原始事件样本跑过 normalizer、diagnostic 和 `provider_event` runtime event：

```js
import {
  assertMeetingPlatformProviderReplayMatrix,
  buildMeetingPlatformProviderReplayReport,
  sampleMeetingPlatformProviderEvents,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-ingest';

const samples = sampleMeetingPlatformProviderEvents('google-meet');
const report = buildMeetingPlatformProviderReplayReport('google-meet', samples);

console.log(report.coverage.meeting_start, report.coverage.meeting_end);
// report.runtime_contract.provider_events_block_realtime === false

assertMeetingPlatformProviderReplayMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});
```

同一能力也随 SDK 包发布为命令行入口，适合下游项目在还没有完整 host 服务时先验 provider 样本：

```sh
npx meeting-platform-provider-replay \
  --platforms=google-meet,teams,zoom,webex,lark \
  --records-file=data/provider-events.json \
  --out-file=data/provider-replay-matrix.json \
  --report-file=data/provider-replay-report.json
```

`records-file` 可以是 `{ "recordsByPlatform": { "google_meet": [...] } }`，也可以在只验单个平台时直接传事件数组；不传文件时命令会使用 SDK 内置样本跑通 CI wiring。`createMeetingAppTimelineSdk()` 也直接暴露同一层：`sdk.providerReplayReport()`、`sdk.assertProviderReplayReport()`、`sdk.providerReplayMatrix()` 和 `sdk.assertProviderReplayMatrix()`。这适合放进下游项目 CI，证明 provider 样本能被 SDK 识别为开始、结束、参会人或 artifact 事件；真正的实时标注仍必须使用本地观察建轴和 `captured_at_ms`。

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

真实 webhook 接通后，建议同时用 `platform-capture` 捕获一小批原始事件做回放验收。默认记录 parsed body、脱敏 headers、raw body 的 SHA-256；只有显式传 `includeRawBody: true` 才保存 raw body：

```js
import {
  buildPlatformCaptureAcceptanceReport,
  capturePlatformWebRequest,
  parsePlatformCaptureJsonl,
  serializePlatformCaptureRecord,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-capture';

export async function POST(request) {
  const capture = await capturePlatformWebRequest(request, {
    basePath: '/api/platform-events',
  });
  await appendLine('platform-events.jsonl', serializePlatformCaptureRecord(capture));
  return meetingKit.handleFetchRequest(request);
}

const records = parsePlatformCaptureJsonl(await readText('platform-events.jsonl'));
const googleReplay = buildPlatformCaptureAcceptanceReport('google-meet', records, {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  requireEndEvent: true,
});
```

部署前或 CI 里可以用 `platform-gate` 做硬性验收。默认要求真实捕获记录或真实样本；fixture 只能证明 SDK wiring，除非显式允许，否则不会被当成生产证据：

```js
import {
  assertPlatformLaunchGate,
  buildPlatformLaunchGate,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-gate';

const gate = buildPlatformLaunchGate('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  records,
  requireEndEvent: true,
  requireParticipants: true,
});

if (!gate.production_ready) {
  console.error(gate.blocking_issues, gate.next_actions);
}

// gate 会同时检查 setup、candidate observation runtime contract、真实样本 coverage 和 fixture 污染。
assertPlatformLaunchGate('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  records,
  requireEndEvent: true,
});
```

如果要给配置页、接入向导或 CI 验收生成一份“这个会议平台现在能不能接进 timeline”的总报告，可以用 `platform-onboarding`。它会合并 `permission plan`、`integration plan`、候选观察 gate、真实样本 `acceptance` 和 `artifact-plan`：

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

// onboarding.status === 'ready' | 'blocked_by_setup' | 'blocked_by_runtime_contract' | 'needs_real_samples' | 'needs_more_coverage'
// onboarding.runtime_contract.annotation_time_field === 'captured_at_ms'
// onboarding.candidate_observation_gate.accepted === true 表示 observe-candidates 建轴入口可交给宿主项目。
// onboarding.next_actions includes missing env/config steps, candidate observation fixes, and post-meeting artifact fetch actions
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

如果宿主拿到的是连续 UI/音频/DOM 采样，不要每帧都直接写时间轴。先用 `active-speaker` 做去抖和切换滤波，默认候选发言人稳定约 300ms 才发出 `speaker_started`，静音持续约 1500ms 才发出 `speaker_ended`：

```js
import { createActiveSpeakerTimelineObserver } from '@ai-annotation/meeting-timeline-sdk/adapters/active-speaker';

const speakers = createActiveSpeakerTimelineObserver(timeline, {
  source: 'browser_dom_observer',
  minStableMs: 300,
  switchStableMs: 400,
  endIdleMs: 1500,
  applyOptions: { speakerAsAnnotation: true },
});

await speakers.observe({
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    meeting_url: 'https://meet.google.com/abc-defg-hij',
  },
  activeSpeaker: {
    id: 'speaker-ada',
    name: 'Ada',
    speaking: true,
  },
  observedAtMs: Date.now(),
});
```

短暂误检会停留在候选状态，不会落到时间轴；真正切换发言人时，SDK 会用候选第一次出现的时间作为 `captured_at_ms`，避免为了滤波把视觉落点推迟几百毫秒。

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

如果宿主项目不想直接把每个 `speaker_started` 写入用户标注流，可以用 `platform-speaker-track` 先生成独立的发言人位置轨。它只输出 `speaker_track` marks，不包含转写正文；输入可以是连续 active-speaker samples，也可以是宿主已经去抖过的 `speaker_started/speaker_ended` signals：

```js
import {
  buildMeetingPlatformSpeakerTrack,
  buildMeetingPlatformSpeakerTrackMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-speaker-track';

const matrix = buildMeetingPlatformSpeakerTrackMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.rows[*].provider_events_block_realtime === false
// matrix.rows[*].source_order[0] === 'local_active_speaker_observer'

const track = buildMeetingPlatformSpeakerTrack('google-meet', {
  samples: [
    {
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      activeSpeaker: { id: 'speaker-ada', name: 'Ada', speaking: true },
      observedAtMs: Date.now(),
    },
  ],
}, {
  minStableMs: 500,
  switchStableMs: 700,
  endIdleMs: 1500,
  minSegmentMs: 800,
});

// track.marks 可批量 insertMarks() 到会议时间轴的 speaker rail。
// track.diagnostics 会列出 dropped_duplicate_count / dropped_short_segment_count。
```

对应 CLI 可用于现场采样或 CI 报告：

```sh
npm run meeting-platform:speaker-track -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --samples-file=data/meeting-platform-speaker-samples.json \
  --out-dir=data/meeting-platform-speaker-tracks \
  --report-file=data/meeting-platform-speaker-track-report.json
```

参会人位置轨用 `platform-participant-track`。它只输出 `participant_track` marks，表示谁在什么时间加入或离开；输入可以是本地 roster snapshots，也可以是平台 adapter 已经归一化的 `participant_joined/participant_left` signals。默认不会把首帧 roster 当成“刚加入”，会过滤短窗口内的重复事件，并抑制短暂断线重连造成的 leave/join 抖动：

```js
import {
  buildMeetingPlatformParticipantTrack,
  buildMeetingPlatformParticipantTrackMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-participant-track';

const matrix = buildMeetingPlatformParticipantTrackMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.rows[*].provider_events_block_realtime === false
// matrix.rows[*].source_order[0] === 'local_roster_observer'

const track = buildMeetingPlatformParticipantTrack('google-meet', {
  snapshots: [
    {
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      observed_at_ms: Date.now(),
      participants: [{ id: 'ada', name: 'Ada' }],
    },
  ],
}, {
  leaveStableMs: 1500,
  suppressReconnectGapMs: 10000,
});

// track.marks 可批量 insertMarks() 到会议时间轴的 participant rail。
// track.diagnostics 会列出 dropped_duplicate_count / suppressed_reconnect_pair_count / pending_leave_count。
```

对应 CLI：

```sh
npm run meeting-platform:participant-track -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --input=data/meeting-platform-participant-input.json \
  --out-dir=data/meeting-platform-participant-tracks \
  --report-file=data/meeting-platform-participant-track-report.json
```

如果另一个项目需要直接渲染时间轴，优先用 `platform-timeline-view` 把会议轴、用户标注、speaker/participant/artifact/transcript 轨道合成 renderer-agnostic 数据包。它不输出 SVG/Canvas/HTML，只输出 rails、viewport、ticks、marker `x_ratio`、可见 marker 和未校准 marker：

```js
import {
  buildMeetingPlatformTimelineView,
  buildMeetingPlatformTimelineViewMatrix,
  zoomMeetingPlatformTimelineViewport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-timeline-view';

const matrix = buildMeetingPlatformTimelineViewMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.rows[*].provider_events_block_realtime === false
// matrix.rows[*].transcript_blocks_realtime === false

const view = buildMeetingPlatformTimelineView('google-meet', {
  meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: Date.now(),
    duration_ms: 10 * 60 * 1000,
  },
  annotations: [{ id: 'note-1', label: 'why?', time_ms: 90_000 }],
  speakerTrack: { marks: speakerTrack.marks },
  participantTrack: { marks: participantTrack.marks },
  artifactHandoff: { rows: artifactHandoff.rows },
}, {
  viewportStartMs: 60_000,
  viewportDurationMs: 240_000,
});

// view.visible_markers[*].x_ratio 可直接映射到任意 UI 宽度。
const zoomedViewport = zoomMeetingPlatformTimelineViewport(view.viewport, 2);
```

对应 CLI：

```sh
npm run meeting-platform:timeline-view -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --input=data/meeting-platform-timeline-view-input.json \
  --out-dir=data/meeting-platform-timeline-views \
  --report-file=data/meeting-platform-timeline-view-report.json
```

设备或宿主项目真正插入标注前，应先用 `platform-clock-sync` 把设备本地时间校准成服务端会议轴时间。它沿用 `/api/time` 的 midpoint 算法：`clock_offset_ms = server_time_ms - ((client_send_at_ms + client_receive_at_ms) / 2)`，再把原始 `device_mark_end_ms` 转成 `captured_at_ms`。这个模块同样和会议平台无关；Google Meet、Teams、Zoom、Webex、Lark 都使用同一套时钟契约：

```js
import {
  buildMeetingPlatformClockSyncReport,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-clock-sync';

const clock = buildMeetingPlatformClockSyncReport('google-meet', {
  samples: [
    {
      client_send_at_ms: sendAtMs,
      server_time_ms: serverTimeMs,
      client_receive_at_ms: receiveAtMs,
    },
  ],
  annotation: {
    id: 'note-1',
    label: 'why?',
    captured_at_ms: deviceMarkEndMs,
  },
});

// clock.calibrated_annotation.captured_at_ms 才应该交给 annotation intake。
```

对应 CLI：

```sh
npm run meeting-platform:clock-sync -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --input=data/meeting-platform-clock-sync-input.json \
  --out-dir=data/meeting-platform-clock-sync \
  --report-file=data/meeting-platform-clock-sync-report.json
```

跨 Google Meet、Teams、Zoom、Webex、Lark 适配时，还需要把本地观察、provider event 和当前标注绑定到同一场会议。`platform-session-binding` 会比较 `meeting_id`、`external_meeting_id`、URL 派生稳定 id、平台、标题和时间窗口，输出是否绑定当前轴、是否应该从本地观察或 provider start 开轴、是否 pending，或是否存在冲突：

```js
import {
  buildMeetingPlatformSessionBinding,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-session-binding';

const binding = buildMeetingPlatformSessionBinding('google-meet', {
  current_meeting: currentAxis,
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: Date.now(),
  },
  signals: providerSignals,
  annotation: clock.calibrated_annotation,
});

if (binding.status === 'bound_to_current_axis') {
  await timeline.insertMark(decision.insert_payload);
}
if (binding.should_start_axis) {
  await timeline.startMeeting(binding.start_payload);
}
```

对应 CLI：

```sh
npm run meeting-platform:session-binding -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --input=data/meeting-platform-session-binding-input.json \
  --out-dir=data/meeting-platform-session-binding \
  --report-file=data/meeting-platform-session-binding-report.json
```

正式交给外部项目接入时，推荐直接用 `platform-realtime-annotation`，它把 `platform-clock-sync`、`platform-session-binding` 和 `platform-annotation-intake` 合成一条决策链。输入是一条设备标注、设备/服务端时钟同步样本、本地会议观察或 provider signal；输出是明确动作：`insert_mark`、`start_meeting_session + insert_mark`、`store_pending_mark`、`run_clock_sync_before_realtime_insert` 或冲突阻断。这样宿主项目不用复刻 demo 的“当前轴/待绑定/开轴/冲突/缺时间戳”分支。

```js
import {
  buildMeetingPlatformRealtimeAnnotation,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-realtime-annotation';

const decision = buildMeetingPlatformRealtimeAnnotation('google-meet', {
  clock_sync: {
    offset_ms: 12,
    rtt_ms: 48,
  },
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: meetingStartMs,
  },
  local_observer: {
    url: 'https://meet.google.com/abc-defg-hij',
    observed_at_ms: Date.now(),
  },
  annotation: {
    id: 'mark-001',
    label: 'why?',
    captured_at_ms: deviceInkEndMs,
    strokes: rawStrokes,
  },
});

if (decision.actions.includes('start_meeting_session')) {
  await timeline.startMeeting(decision.start_payload);
}
if (decision.actions.includes('insert_mark')) {
  await timeline.insertMark(decision.insert_payload);
}
if (decision.actions.includes('store_pending_mark')) {
  await pendingStore.put(decision.pending_payload);
}
```

对应 CLI 可以批量导出多平台链路报告：

```sh
npm run meeting-platform:realtime-annotation -- \
  --input=data/meeting-platform-realtime-annotation-input.json \
  --out-dir=data/meeting-platform-realtime-annotation \
  --report-file=data/meeting-platform-realtime-annotation-report.json
```

如果下游项目只需要“把当前手写/标注插到会议时间轴上”，不要复制 demo 服务端里的条件判断，直接用 `platform-annotation-intake`。它不拉 provider、不等转写，只检查标注是否携带可靠 `captured_at_ms`，再根据当前会议轴状态给出动作：直接插入当前轴、先开一个 open session 再插入、进入 pending 等真实会议 start 回填，或者标记为缺时间戳/会后审计。

```js
import {
  buildMeetingPlatformAnnotationIntake,
  buildMeetingPlatformAnnotationIntakeMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-annotation-intake';

const matrix = buildMeetingPlatformAnnotationIntakeMatrix({
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.provider_blocking_count === 0
// matrix.transcript_blocking_count === 0

const decision = buildMeetingPlatformAnnotationIntake('google-meet', {
  current_meeting: {
    platform: 'google_meet',
    meeting_id: 'abc-defg-hij',
    start_time_ms: meetingStartMs,
  },
  annotation: {
    id: 'note-1',
    label: 'why?',
    captured_at_ms: inkEndAtMs,
    text_candidates: ['why?', 'why'],
    strokes,
  },
});

if (decision.status === 'ready_to_insert_current_axis') {
  await timeline.insertMark(decision.insert_payload);
}
if (decision.status === 'start_open_session_then_insert') {
  await timeline.startMeeting(decision.open_session_payload);
  await timeline.insertMark(decision.insert_payload);
}
```

对应 CLI：

```sh
npm run meeting-platform:annotation-intake -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --input=data/meeting-platform-annotation-input.json \
  --out-dir=data/meeting-platform-annotation-intake \
  --report-file=data/meeting-platform-annotation-intake-report.json
```

## 会后转写导入

事件 adapter 只负责告诉你 transcript/recording 已生成；正文内容建议会后拉取后再导入。`artifact-plan` 可以把 `artifact_ready` signal 转成平台相关的补拉/导入计划，告诉宿主应该用哪个 provider API、哪个 transcript normalizer，以及是否只是录制 metadata：

如果要把 Google Meet / Teams / Zoom / Webex / Lark 的会后产物能力交给另一个项目，优先用 `platform-artifact-handoff`。它会把 transcript、recording、smart notes 的事件来源、fetch strategy、normalizer、token env、导入 endpoint 和“绝不阻塞实时标注”的约束收成统一对象：

```js
import {
  buildMeetingPlatformArtifactHandoff,
  buildMeetingPlatformArtifactHandoffMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-artifact-handoff';

const matrix = buildMeetingPlatformArtifactHandoffMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
});

// matrix.realtime_blocking_count === 0
// matrix.rows[*].post_meeting_only === true

const handoff = buildMeetingPlatformArtifactHandoff('google-meet', {
  signals: [
    {
      type: 'artifact_ready',
      meeting: { platform: 'google_meet', meeting_id: 'abc-defg-hij' },
      occurred_at_ms: Date.now(),
      artifact_kind: 'transcript',
      artifact_id: 'transcript-1',
    },
  ],
}, {
  baseUrl: 'https://timeline.example.com',
});

// handoff.fetch_requests 是会后补拉请求骨架；Authorization 会被 redacted。
// handoff.transcript_import_count > 0 时，补拉后用 transcript adapter 导入。
```

命令行版本适合接入项目或 CI 生成会后产物交接报告：

```sh
npm run meeting-platform:artifact-handoff -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --signals-file=data/meeting-platform-artifact-signals.json \
  --out-dir=data/meeting-platform-artifact-handoffs \
  --report-file=data/meeting-platform-artifact-handoff-report.json
```

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

进入真实适配推进时，用 `platform-rollout` 把官方事件 gate 和本地 DOM gate 合成一个上线决策：

```js
import { buildMeetingPlatformRolloutPlan } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';

const plan = buildMeetingPlatformRolloutPlan('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
  },
  providerRecords: capturedGoogleWorkspaceEvents,
  meetingAppRecordSet: capturedGoogleMeetDomRecordSet,
});

// plan.status === 'production_ready' 表示本地低延迟观察和 provider 事件回填都已通过。
// plan.ready_for_realtime_annotations === true 表示至少有一条路径可实时落标注。
```

`buildMeetingPlatformRolloutSummary()` 可同时汇总 Google Meet、Teams、Zoom、Webex、Lark 的 `production_ready`、`ready_for_realtime_annotations` 和 `next_actions`。如果只有本地 DOM 证据通过，会进入 `realtime_ready_provider_pending`；如果只有 provider 事件通过，会进入 `provider_ready_collect_local_evidence`，提醒继续采真实会议页 DOM。

现场推进多平台适配时，可以再拿 `buildMeetingPlatformAdaptationRunbook()` 生成采样清单。它不会把 fixture 当成生产证据，而是把本地 DOM 快照、provider event 样本、验证命令和最终 rollout gate 放到同一份结构里：

```js
import { buildMeetingPlatformAdaptationRunbook } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-rollout';

const runbook = buildMeetingPlatformAdaptationRunbook('zoom', {
  baseUrl: 'https://timeline.example.com',
});

// runbook.local_dom.required_snapshots 描述 active speaker / meeting ended 等必须采的真实页面状态。
// runbook.provider_events.required_coverage 描述 provider webhook 至少要证明 meeting_start / meeting_end。
// runbook.steps 可以直接渲染成接入向导或 CI checklist。
```

现场接入时，推荐把一次真实会议的本地 DOM 证据、provider webhook 证据和 rollout 判断打包成 `platform-evidence-package`。包默认只保存环境变量 key，不保存 secret value；provider 事件仍沿用 `platform-capture` 的 header 脱敏和 raw body 默认不保存策略：

```js
import { buildMeetingPlatformEvidencePackage } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package';

const evidencePackage = buildMeetingPlatformEvidencePackage('google-meet', {
  providerRecords,
  meetingAppRecordSet,
}, {
  baseUrl: 'https://timeline.example.com',
  env: {
    GOOGLE_PUBSUB_OIDC_AUDIENCE: 'https://timeline.example.com/api/platform-events/google-meet',
  },
});

// evidencePackage.rollout_plan.status === 'production_ready' 才表示本地实时观察和 provider 回填都已验证。
// evidencePackage.handoff 可直接交给另一个项目继续接入或复验。
```

收到另一个项目交来的包时，不要直接信任包内旧的 `rollout_plan`，用 verify 重新计算：

```js
import { verifyMeetingPlatformEvidencePackage } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package';

const verification = verifyMeetingPlatformEvidencePackage(evidencePackage, {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
});

// verification.passed === true 表示当前环境下仍满足 production_ready。
```

如果采样过程在宿主里分步发生，可以用 builder 累积：

```js
import { createMeetingPlatformEvidencePackageBuilder } from '@ai-annotation/meeting-timeline-sdk/adapters/platform-evidence-package';

const evidence = createMeetingPlatformEvidencePackageBuilder('zoom', {
  baseUrl: 'https://timeline.example.com',
});

evidence.captureProviderWebhook(reqLikeObject);
evidence.addMeetingAppRecord(activeSnapshotRecord);
evidence.addMeetingAppRecord(endedSnapshotRecord);

const handoff = evidence.exportPackage();
```

如果要把会议时间轴能力交给另一个项目接入，优先生成 `platform-live-adapter` 的 handoff bundle。它不是一份纯文档，而是机器可读的接入契约：包含 SDK import、宿主必须提供的输入、候选会议观察合同、实时标注时间字段、证据路径、CI 命令、单平台 readiness 和多平台矩阵。

```js
import {
  buildMeetingPlatformLiveAdapterHandoff,
  buildMeetingPlatformLiveAdapterHandoffBundle,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-live-adapter';

const googleHandoff = buildMeetingPlatformLiveAdapterHandoff('google-meet', {
  baseUrl: 'https://timeline.example.com',
  evidencePackage,
});

// googleHandoff.host_contract.annotation_timestamp_field === 'captured_at_ms'
// googleHandoff.candidate_observation_contract.runtime_event_action === 'observe_platform_candidates'
// googleHandoff.commands.validate_live_readiness === 'npm run meeting-platform:live-readiness'
// googleHandoff.evidence_paths.evidence_package 指向可复验的 provider + DOM 证据包。

const bundle = buildMeetingPlatformLiveAdapterHandoffBundle({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'microsoft-teams', 'zoom', 'webex', 'lark'],
});

// bundle.handoffs 可以直接交给外部项目生成接入 checklist。
// bundle.candidate_observer_count 必须等于 bundle.platform_count。
// bundle.readiness_matrix 可以作为 CI gate，保证实时标注不依赖会后转写或 provider 事件。
```

如果已经使用 `platform-kit`，同样可以从高层入口拿到这份交付物：

```js
const kit = createMeetingPlatformTimelineKit(client, {
  baseUrl: 'https://timeline.example.com',
});

const handoff = kit.platformLiveAdapterHandoff('zoom');
const bundle = kit.platformLiveAdapterHandoffBundle({
  platforms: ['google-meet', 'zoom'],
});
```

官方 provider 侧接入可以先生成 `platform-provider-connection` 包。它会把 Google Meet / Teams / Zoom / Webex / Lark 的事件、scope/permission、安全校验、订阅请求和 SDK 验收命令放到同一份结构里，同时明确 provider 事件只做 reconcile/backfill，不阻塞实时标注：

```js
import {
  buildMeetingPlatformProviderConnectionPack,
  buildMeetingPlatformProviderConnectionMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-provider-connection';

const googleProvider = buildMeetingPlatformProviderConnectionPack('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  subscription: {
    targetResource: '//cloudidentity.googleapis.com/users/me',
    pubsubTopic: 'projects/demo/topics/meet-events',
  },
});

// googleProvider.event_mapping 会列出 Workspace Events 到 meeting_started/artifact_ready 等 SDK signal 的映射。
// googleProvider.security.verifier === 'verifyGooglePubSubOidcJwt'
// googleProvider.realtime_annotation_policy.provider_events_block_realtime === false

const providerMatrix = buildMeetingPlatformProviderConnectionMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'microsoft-teams', 'zoom', 'webex'],
});
```

对应的命令行报告可以在接入项目或 CI 里跑：

```sh
npm run meeting-platform:provider-connection -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex \
  --subscriptions-file=data/provider-subscriptions.json \
  --report-file=data/meeting-platform-provider-connection-report.json
```

`subscriptions-file` 可以是平台到订阅参数的 JSON map；例如 `google_meet.pubsubTopic`、`microsoft_teams.joinWebUrl`、`zoom.webhookUrl`、`webex.targetUrl`。报告会列出缺失的 secret/env、官方文档链接、事件映射和 `provider_events_block_realtime=false` 的实时标注约束。

如果接入方已经拿到各平台的订阅参数，需要更直接地生成“可以创建订阅”的交接包，用 `platform-subscription-handoff`。它会复用 `platform-provider-connection` 的权限、安全和事件映射判断，但输出面向执行的请求列表：Google Workspace Events 的 `/v1beta/subscriptions` body、Microsoft Graph `/subscriptions` body、Zoom 事件订阅配置摘要、Webex 每个 webhook 的 `/v1/webhooks` body。Lark/飞书当前仍按控制台或长连接配置归为 `manual_setup`，不会伪造不存在的统一创建 API。

```js
import {
  buildMeetingPlatformSubscriptionHandoff,
  buildMeetingPlatformSubscriptionHandoffMatrix,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-subscription-handoff';

const handoff = buildMeetingPlatformSubscriptionHandoff('google-meet', {
  baseUrl: 'https://timeline.example.com',
  env: process.env,
  subscription: {
    targetResource: '//cloudidentity.googleapis.com/users/me',
    pubsubTopic: 'projects/demo/topics/meet-events',
    ttl: '86400s',
  },
});

// handoff.ready_to_create === true 时，handoff.requests 可交给平台接入脚本执行。
// handoff.handoff_contract.annotation_timestamp_field === 'captured_at_ms'

const matrix = buildMeetingPlatformSubscriptionHandoffMatrix({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'teams', 'zoom', 'webex', 'lark'],
  subscriptions: {
    google_meet: { targetResource: '//cloudidentity.googleapis.com/users/me', pubsubTopic: 'projects/demo/topics/meet-events' },
    microsoft_teams: { joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/demo' },
    zoom: { webhookUrl: 'https://timeline.example.com/api/platform-events/zoom' },
    webex: { targetUrl: 'https://timeline.example.com/api/platform-events/webex' },
  },
});
```

对应命令适合作为下游项目的订阅准备 gate：

```sh
npm run meeting-platform:subscription-handoff -- \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --subscriptions-file=data/provider-subscriptions.json \
  --out-dir=data/meeting-platform-subscription-handoffs \
  --report-file=data/meeting-platform-subscription-handoff-report.json
```

报告里 `ready_to_create_count` 表示能直接创建订阅的平台数，`manual_setup_count` 表示必须去平台控制台或长连接配置的平台数，`security_blocked_count` 和 `parameter_missing_count` 直接给 CI/接入面板做失败原因。即使 provider 订阅已全部可创建，实时标注仍然以宿主捕获的 `captured_at_ms` 为准；provider 事件只做 start/end/artifact 的回填和审计。

如果要给另一个项目一个更完整的“可改造骨架”，用 `platform-host-integration` 生成 host scaffold。它会输出 `package.json`、timeline client、host wrapper、per-platform adapter entries、framework-neutral HTTP route、strategy/handoff/readiness/runtime-bundle/平台解析/extension-plan/integration-runtime 脚本和 README。生成的 host wrapper 会先创建 `createMeetingPlatformIntegrationRuntime()`，并暴露 `host.platformAdapter(platform)` / `host.platformAdapters()`；下游项目可以优先接统一 runtime，再按需下钻到 Google Meet、Teams、Zoom、Webex、Lark 各自的 resolve、observeCandidates、insertAnnotation、speakerTrack、participantTrack 或 provider reconcile 入口：

```sh
npx meeting-platform-host-integration \
  --base-url=https://timeline.example.com \
  --platforms=google-meet,teams,zoom,webex,lark \
  --package-name=meeting-platform-host \
  --out-dir=./meeting-platform-host \
  --report-file=./meeting-platform-host-report.json
```

当前仓库内也可以直接跑：

```sh
npm run meeting-platform:host-integration
```

CLI 报告里的 `adapter_runtime_ready`、`candidate_observation_ready`、`meeting_track_ready` 和 `platform_conformance_ready` 必须为 `true`，才说明这个 scaffold 能作为下游项目的可复用 SDK 接入骨架。

```js
import {
  buildMeetingPlatformHostIntegrationScaffold,
  assertMeetingPlatformHostIntegrationScaffold,
} from '@ai-annotation/meeting-timeline-sdk/adapters/platform-host-integration';

const scaffold = buildMeetingPlatformHostIntegrationScaffold({
  baseUrl: 'https://timeline.example.com',
  platforms: ['google-meet', 'microsoft-teams', 'zoom'],
  packageName: 'meeting-platform-host',
});

assertMeetingPlatformHostIntegrationScaffold(scaffold);
// scaffold.files 里包含 src/meeting-platform-host.mjs、src/http-routes.mjs 和 src/platform-adapters/{platform}.mjs。
// scaffold.plan.adapter_runtime_contract 会列出每个平台的 adapter file、selected surface、observe/insert 方法和 captured_at_ms 契约。
// scaffold.plan.candidate_observation_contract 会列出每个平台的 observe-candidates 覆盖率、tabs 权限和 message type。
// host.platformAdapter('google_meet').insertAnnotation({ label: 'why?', captured_at_ms: Date.now() }) 是每个平台的薄运行时入口。
// host.adapterRoutes() 和 /api/meeting-platform/adapter-routes 可把 Google Meet/Teams/Zoom 的本地观察、provider 回填、会后转写路线交给宿主决策。
// host.adaptationStrategyMatrix() 和 /api/meeting-platform/strategy 可给宿主先做平台适配决策。
// host.resolvePlatform(input) 和 /api/meeting-platform/resolve 可用会议 URL/window/title 判断当前应启用哪个平台适配器。
// host.resolvePlatformCandidates(input) 和 /api/meeting-platform/resolve-candidates 可从多窗口/多标签页快照中选择当前会议。
// host.observePlatformCandidates(input) 和 /api/meeting-platform/observe-candidates 可把当前会议出现/消失写成实时轴事件。
// host.integrationRuntimeSummary() 和 /api/meeting-platform/integration-runtime 可给宿主接入面板读取。
// host.integrationRuntimeManifest() 和 /api/meeting-platform/integration-runtime/manifest 可给 CI/handoff gate 读取静态 SDK 接线状态。
// scaffold.plan.endpoints/commands 会列出下面两个正式验收入口，方便宿主接入面板或 CI 自动发现。
// host.runIntegrationRuntimeManifest() 和 /api/meeting-platform/integration-runtime/run-manifest 会实际运行 handoff readiness/runtime replay gate。
// host.runHandoffReadiness() 和 /api/meeting-platform/handoff-readiness 可单独复验真实证据包是否能交给宿主上线。
// /api/meeting-platform/runtime-events 可让扩展/WebView/native host 统一投递 observe/insert/provider 事件。
// host.runtimeBundles() 和 /api/meeting-platform/runtime-bundles 可直接给扩展/WebView/native host 读取。
// host.extensionInstallPlan() 和 /api/meeting-platform/extension-plan 可直接给扩展构建器读取。
```

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
