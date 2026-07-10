# 会议时间轴 SDK 进度快照（2026-07-09，更新于 2026-07-10）

## 当前结论

多平台本地实时链路的代码和验收基础设施已经收敛，但现场 P0 尚未完成。不要把静态门禁通过等同于真实会议通过。

当前权威状态：

- 分支：`codex/meeting-timeline-sdk`；
- 静态 adapter checklist：Google Meet、Teams、Zoom、Webex、Lark 共 5/5 通过；
- 浏览器扩展：五平台 MV3 产物构建通过，浏览器 bundle 不含 Node 加密依赖；
- 自动证据：支持开始/结束 DOM 快照、操作者加入/离开时间、标注延迟与轴误差、发言人 marker 延迟、跨会议隔离；
- SDK 回归与安装 smoke：通过；
- 真实现场报告：0/5 平台、0/15 场会议通过。

## 已完成

1. 实时主链路固定为 `local observer -> current axis -> captured_at_ms annotation`。
2. provider 事件和会后转写只做回填，不阻塞 P0。
3. 扩展会捕获加入/离开点击；开始证据尚未创建时，服务端会暂存时间引用并在建轴后回算延迟。
4. 每场会议自动生成 `meeting_platform_field_evidence_input`，按 meeting ID 隔离。
5. `meeting-platform:p0-field-report` 只取每个平台最新连续 3 场，执行统一硬门禁。
6. Teams、Webex、Lark 的混合 surface 名称已在导入边界归一为实际 browser/native 入口。

## 尚未完成

以下项目必须由真实外部状态提供，当前没有证据：

- Chrome 尚未加载 `data/meeting-app-extension`；
- 汉王设备未出现在 `adb devices -l`；
- 五个平台尚未各完成连续 3 场真实会议；
- 每场 5 条真实设备标注和两位发言人切换尚未采集；
- 因没有现场记录，延迟、轴误差、重复/丢失和跨会议污染尚不能判定通过。

## 权威命令

静态门禁：

```bash
npm run meeting-platform:adapter-acceptance-checklist -- \
  --target=static \
  --fail-on-incomplete=true
```

现场门禁：

```bash
npm run meeting-platform:p0-field-report -- \
  --platforms=google-meet,teams,zoom,webex,lark \
  --dir=data/meeting-platform-field-evidence \
  --fail-on-incomplete=true
```

现场执行标准见 [多会议平台 P0 现场验收单](meeting-platform-p0-field-acceptance.md)。只有第二条命令报告 5/5 平台通过，才允许声明 P0 完成。
