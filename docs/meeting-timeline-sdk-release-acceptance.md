# 会议时间轴 SDK 发布验收单

## 验收边界

本验收单判断 `@ai-annotation/meeting-timeline-sdk` 是否能交付给其他项目使用。SDK 不要求部署到汉王、电子纸或 Android 设备。

SDK 的必验职责：

- 建立和结束当前会议轴；
- 接收任意来源的标注，并以源端 `captured_at_ms` 定位；
- 为无 ID 的标注生成稳定 ID；
- 网络失败时以同一 ID 重试，失败项保留待发送；
- 保证时间轴上不丢失、不重复、不跨会议污染；
- 接收已滤波的发言人位置；
- 提供五个平台的统一 adapter 契约和可安装 npm 包。

以下内容不阻塞 SDK 发布：

- 汉王或其他具体硬件应用接入；
- 各会议软件真实页面 selector 的现场兼容性；
- provider OAuth、webhook、转写和会后产物；
- 发言内容识别。

这些能力分别进入“设备集成验收”和“平台现场验收”，不得反向成为 SDK 包发布的前置条件。

## 发布门禁

在仓库根目录执行：

```bash
npm run sdk:release-acceptance
```

命令连续执行五层检查：

1. 标准生产者单元测试；
2. `/producer` 浏览器入口依赖扫描，禁止转入根入口或任何 `node:` 内置模块；
3. 发布验收脚本自身的回归测试；
4. `npm pack` 后从独立 consumer 项目导入根入口、`/producer` 和 adapter 子路径；
5. 五平台静态 adapter 门禁，以及隔离服务上的 5 平台 × 3 场端到端契约验收。

默认报告写入：

```text
data/meeting-timeline-sdk-release-acceptance-report.json
```

报告必须明确包含：

```text
scope = deterministic_sdk_contract_and_transport
field_pilot_claimed = false
device_required = false
accepted = true
accepted_platform_count = 5
accepted_run_count = 15
```

## 统一指标

每个平台执行 3 个隔离会话，每场通过标准生产者写入 5 条标注，并写入 2 个不同发言人的稳定 marker。

| 指标 | 门槛 |
| --- | ---: |
| 建轴延迟 | 最大 1,500 ms |
| 标注可见延迟 | P95 不超过 300 ms，单条不超过 800 ms |
| 标注轴位置误差 | 绝对值最大 500 ms |
| 发言人 marker 延迟 | 最大 1,500 ms |
| 闭轴延迟 | 最大 6,500 ms |
| 丢失标注 | 0 |
| 可见重复标注 | 0 |
| 跨会议污染 | 0 |

第一场还会模拟“服务端已提交、客户端响应丢失”。生产者必须以同一 ID 重试，证据中的 `delivery_attempt_count` 至少为 2，但 `visible_instance_count` 必须仍为 1。

## 包交付要求

外部项目至少可以使用以下入口：

```js
import {
  createMeetingTimelineClient,
  createMeetingTimelineAnnotationProducer,
} from '@ai-annotation/meeting-timeline-sdk';

import {
  createMeetingTimelineAnnotationProducer as createProducer,
} from '@ai-annotation/meeting-timeline-sdk/producer';
```

标准生产者必须满足：

- 默认来源是 `sdk_annotation_producer`，不伪装成具体设备；
- `capture()` 只捕获，不发送；
- `publish()` 捕获并发送；
- `enqueue()` 和 `flush()` 支持失败后恢复；
- 4xx 参数错误不盲目重试；408、425、429、5xx 和网络错误允许退避重试；
- 每次重试复用原 ID 与原 `captured_at_ms`。

## 与现场验收的关系

SDK 发布门禁证明公共协议、包导出、传输可靠性、时间轴计算和隔离逻辑成立。它不证明某天更新后的 Meet、Teams、Zoom、Webex 或 Lark 页面仍能被 selector 正确识别。

真实平台兼容性继续使用 [多会议平台 P0 现场验收单](meeting-platform-p0-field-acceptance.md)。现场验收可以使用标准生产者完成 5 条人工标注，不需要连接电子纸。硬件团队接入后，再单独验证设备时钟、离线持久化、功耗和 Android 生命周期。
