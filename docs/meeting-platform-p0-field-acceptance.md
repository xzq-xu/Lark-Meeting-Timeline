# 多会议平台 P0 现场验收单

> 本文只验收真实会议软件观察器，不是 SDK 发布门禁。SDK 发布按 [会议时间轴 SDK 发布验收单](meeting-timeline-sdk-release-acceptance.md) 执行，不要求电子纸、Android 或 ADB。

## P0 只验证什么

P0 只验证一条产品主链路：会议软件在本机进入会议后，本地观察器立即建立当前会议轴；任意标注生产者产生标注时，SDK 按标注发生时的 `captured_at_ms` 写入当前轴；离会后本地观察器关闭该轴；下一场会议不得继承上一场标注。

以下能力不阻塞 P0：

- 官方 provider 的开始、结束和参会人事件；
- 实时或会后转写；
- 录制、妙记、智能纪要等会后产物；
- 发言内容识别。发言人轨只记录“谁在何时发言”。

## 统一通过标准

每个平台必须连续完成 3 场真实会议，每场至少插入 5 条标注，并产生至少 2 次持续 3 秒以上的发言人切换。

| 指标 | 浏览器观察器 | 原生观察器 | 判定方式 |
| --- | ---: | ---: | --- |
| 会议开始识别 | 最大 1,500 ms | 最大 2,500 ms | `axis_started_at_ms - operator_join_at_ms` |
| 标注可见延迟 | P95 不超过 300 ms，单条不超过 800 ms | 同左 | `annotation_visible_at_ms - annotation_captured_at_ms` |
| 标注轴位置误差 | 绝对值最大 500 ms | 同左 | 实际位置与 `captured_at_ms - axis_start_ms` 的差 |
| 发言人标记延迟 | 稳定窗口后最大 1,500 ms | 同左 | 不等待转写，只测滤波后的 marker |
| 会议结束识别 | 最大 6,500 ms | 最大 7,500 ms | `axis_ended_at_ms - operator_leave_at_ms` |
| 丢失或重复标注 | 0 | 0 | 每场输入 5 条，轴上必须恰好 5 条 |
| 跨会议污染 | 0 | 0 | 新轴中上一场标注数量必须为 0 |

浏览器结束阈值包含当前 4 秒离会防抖；原生观察器包含当前 5 秒离会防抖。阈值不是 provider 延迟预算，provider 晚到不影响 P0 结果。

## 平台执行顺序

| 顺序 | 平台 | P0 首选采样面 | 备用面 | P0 特别检查 |
| ---: | --- | --- | --- | --- |
| 1 | Google Meet | Chrome/Edge 扩展 | 桌面窗口观察器 | `meet.google.com` 页面内建轴、活动发言人和离会态 |
| 2 | Zoom | 原生窗口/可访问性检测器 | Web 版扩展 | 客户端窗口出现/消失不能依赖 webhook |
| 3 | Microsoft Teams | 原生窗口/可访问性检测器 | Web 版扩展 | 新 Teams 与 Web Teams 的窗口标题/URL 都要命中 |
| 4 | Webex | Chrome/Edge 扩展 | 原生窗口检测器 | Web 会议页离会后必须闭轴 |
| 5 | 飞书/Lark | Chrome/Edge 扩展 | 桌面窗口观察器 | 长连接只回填；本地观察器必须先建轴 |

先完成 Google Meet 的 3 场连续通过，再复用同一流程扩到其余平台。不要五个平台同时调 selector，否则无法判断回归来自公共层还是平台层。

## 会前准备

先确认 SDK 本体已经通过发布门禁：

```bash
npm run sdk:release-acceptance
```

现场只替换会议观察面，不替换公共标注协议。标注可以由 SDK 标准生产者、网页测试面板或真实设备应用产生；不得要求 ADB 才能开始会议平台验收。

先验证五个平台的静态契约。该命令不需要真实会议，也不会要求 provider 凭据：

```bash
npm run meeting-platform:adapter-acceptance-checklist -- \
  --target=static \
  --fail-on-incomplete=true
```

浏览器平台生成并安装扩展：

```bash
npm run meeting-app:extension:build
```

在 `chrome://extensions` 或 `edge://extensions` 开启开发者模式，加载 `data/meeting-app-extension`。Teams 和 Zoom 使用原生客户端时，由宿主 detector 发送等价的 `observe_platform_candidates`、`speaker_track` 和结束信号。

## 每场会议怎么跑

1. 点击会议页面的“加入”按钮。扩展会在 click capture 阶段自动记录 `operator_join_at_ms`；若会议证据尚未创建，服务端先排队，建轴后再回算开始延迟。
2. 在会议开始后的不同时间点通过标准生产者人工插入 5 条标注。每条在操作发生时立即写 `captured_at_ms`，不能在服务端收到时补时间。真实设备可替换标准生产者，但不是必需条件。
3. 让两位发言人各持续发言至少 3 秒，检查时间轴只有滤波后的发言人位置，没有短促抖动产生的重复 marker。
4. 浏览器扩展会在建轴时自动保存 active DOM 快照。调试 selector 时可额外调用：

```js
window.__meetingTimelineLiveCapture.captureActive()
```

5. 点击“离开/结束”按钮。扩展自动记录 `operator_leave_at_ms`，本地结束防抖完成后自动保存 ended DOM 快照并回算结束延迟。调试时可额外调用：

```js
window.__meetingTimelineLiveCapture.captureEnded()
```

6. 仅在需要导出浏览器侧诊断副本时调用：

```js
window.__meetingTimelineLiveCapture.evidencePackage()
```

7. 服务端会按 `<platform>-<meeting-id>.json` 自动写入 `data/meeting-platform-field-evidence/`，其中包含实际 `observer_surface`、开始/结束引用、标注延迟、轴位置误差和发言人 marker 延迟。不要人工改写实测字段。
8. 第 2、3 场开始后立即检查当前轴，任何上一场标注残留都直接判失败。

## 构建并验收证据

下面以 Google Meet 为例。P0 明确关闭 production/provider 要求：

```bash
npm run meeting-platform:field-evidence -- \
  --platforms=google-meet \
  --require-production-ready=false \
  --require-correlation=false \
  --fail-on-incomplete=true

npm run meeting-platform:adapter-acceptance-checklist -- \
  --platforms=google-meet \
  --target=pilot \
  --package-dir=data/meeting-platform-evidence-packages \
  --fail-on-incomplete=true
```

第二条命令现在会读取 `--package-dir` 里的最新真实 evidence package；不再像旧版本一样只生成一份永远缺现场数据的 pilot 清单。

其余平台只替换 `--platforms`：`teams`、`zoom`、`webex`、`lark`。五个平台全部采完后再跑总门禁：

```bash
npm run meeting-platform:adapter-acceptance-checklist -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --target=pilot \
  --package-dir=data/meeting-platform-evidence-packages \
  --fail-on-incomplete=true
```

现场指标的权威汇总命令为：

```bash
npm run meeting-platform:p0-field-report -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --dir=data/meeting-platform-field-evidence \
  --fail-on-incomplete=true
```

该报告只取每个平台按开始时间排序的最新 3 场，逐场检查开始/结束延迟、5 条标注、标注 P95 与硬上限、轴位置误差、两位发言人及 marker 延迟、重复/丢失和跨会议污染。Zoom/Teams 使用网页备用面时按浏览器阈值判定，不会沿用更宽松的原生阈值。

## 失败怎么归因

| 现象 | 先查哪里 | 不应采取的动作 |
| --- | --- | --- |
| 会议开始无轴 | URL/窗口候选、content script/preload、`observe_platform_candidates` | 等 provider start |
| 标注落在未来或错误位置 | 输入设备时钟、`captured_at_ms`、轴开始时间 | 用服务端接收时间覆盖原时间 |
| 离会后轴不结束 | 本地 controls/window 消失信号、4/5 秒 grace timer | 等会后转写或 provider end |
| 新会议出现旧标注 | meeting identity、当前轴切换、持久化查询条件 | 清空全部历史数据掩盖绑定错误 |
| 发言人 marker 抖动 | 500/700 ms 稳定窗口、1,200 ms 同人合并、短段过滤 | 引入实时转写来判断发言人 |

## P0 完成定义

单个平台只有同时满足以下条件才标记为 P0 通过：

- 静态 adapter contract 通过；
- 连续 3 场真实会议全部满足统一阈值；
- 本地开始、标注、发言人位置、结束四类证据齐全；
- 新会议没有任何上一场标注；
- 在关闭 provider 或 provider 延迟到达时仍能完成整条链路；
- pilot checklist 返回 `accepted=true`。

P0 通过后才进入 P1：配置 provider 安全凭据并采集同一会议的 start/end 事件，用于轴边界校准和审计，不改变实时标注的原始 `captured_at_ms`。

这里的 “P0 通过” 只表示该平台观察器通过现场兼容性验证。SDK 包是否可发布由独立发布门禁判断；两份报告不得互相冒充。
