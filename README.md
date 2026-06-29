# Lark Meeting Timeline Demo

这个子项目只验证一件事：会议中实时把外部标注事件贴到会议时间轴上；会议结束后再用飞书/Lark 妙记转写补全发言上下文。

当前支持：

- 飞书/Lark `tenant_access_token` 获取
- 通过 `minute_token` 拉取妙记转写
- 通过飞书/Lark 长连接在本地接收真实事件
- 接收飞书事件回调的 URL verification 和 plaintext event
- 收到飞书直开会议开始事件后自动创建无转写的实时会议时间轴
- 开放会议会话协议 `POST /api/meeting-session/start`，供桌面观察器、汉王端或人工入口在真实会议开始时建轴
- 本地手动开始/结束实时会议，用作没有公网 webhook 时的 fallback
- 会中实时写入外部标注事件，并通过 SSE 自动刷新页面
- 开放标注接口 `POST /api/annotations`，供后续墨水屏/手写设备接入
- 真实验收 probe 期间可自动扫描当前账号近期真实会议作为事件未投递时的建轴兜底
- 手动导入飞书转写 JSON
- 手动导入任意数据序列 JSON
- 在浏览器中可视化 transcript、meeting events、external sequence 的对齐结果

## 启动

```bash
cd /Users/xzq/Documents/Codex/2026-06-15/pdf-ai/work/lark-meeting-timeline-demo
cp .env.example .env
npm run check
npm run start
```

`npm run start` 会设置 `REAL_DEMO_AUTO_ARM=1`：服务启动后自动进入真实等待状态，并启用当前用户会议扫描兜底；租户级泛扫描默认关闭，避免误绑定同租户其他人的会议。默认不会自动写入验收标注、虚拟墨水屏标注或设备流标注，确保每场真实会议的标注序列只来自本场真实设备/接口输入。需要关闭这个默认行为做底层调试时，用：

```bash
npm run start:plain
```

面向产品目标的一键演示命令是：

```bash
npm run demo:live
```

它会进入真实等待态、打开飞书授权页、等待 `vc:meeting.search:read` callback、触发当前用户会议扫描兜底、等待真实会议轴出现，并通过开放标注接口写入验收标注。命令结束会写出 `data/live-demo-report.json`；这个命令对应“用户直接开启飞书会议后，时间轴实时出现标注，文字转写会后处理”的主验收链路。

开会验收时可以另开一个终端观察真实证据流：

```bash
npm run monitor:real-demo
```

如果遇到“页面时间轴看起来不对”，先留存事件投递诊断报告：

```bash
npm run event:report
```

它会写出 `data/event-delivery-report.json`，区分本机长连接/HTTP 接收器是否就绪、事件 handler 是否完整注册、飞书云端是否真的投递过事件、事件是否已经建轴，以及当前是否满足真实轴和实时标注验收。

如果要判断“现在这台机器是否适合直接开会演示”，用：

```bash
npm run onsite:status
```

它会写出 `data/onsite-status-report.json`，同时给出 `minimum_start_ready` 和 `practical_start_ready`。前者表示事件通道或扫描兜底至少有一条入口可尝试；后者要求当前用户扫描兜底已授权可用，或已经观察到飞书真实事件投递，更适合现场演示前做硬卡口。

如果只是排查当前用户会议扫描授权，用：

```bash
npm run auth:status
```

它会写出 `data/auth-meeting-scan-status-report.json`，展示 `vc:meeting.search:read` 是否已授权、OAuth callback 是否已经回到本机、当前用户扫描是否开启，并检查授权 URL 的跳转链是否仍保留 localhost callback。默认只报告状态，不把未授权视为命令失败；需要硬卡口时加 `--require-ready=true`。需要直接打开授权页并等待 callback 时，用：

```bash
npm run auth:open
```

授权完成后脚本会自动触发一次当前用户会议扫描。

现场整体验收优先用这条命令。它会先做设备接入只读预检，确认开放标注端点、CORS、时钟同步、SSE 和验收摘要都可用；预检通过后再打开飞书 OAuth，等待授权，进入真实会议等待，并在真实会议轴出现后通过开放标注接口写入一条验收标注：

```bash
npm run accept:onsite
```

命令完成或失败时会写出两份报告：`data/onsite-device-preflight-report.json` 保存开放标注端点、CORS、时钟同步、SSE 和验收摘要的只读预检结果；`data/onsite-acceptance-report.json` 保存 `product_acceptance_complete` 对应的底层证据：真实会议轴来源、开放标注数量、SSE 广播数量和每个验收项的 `ok` 状态。这些文件用于事后确认本次现场验收，而不是替代页面实时显示。

如果这次现场要严格证明“飞书会议开始事件本身完成投递并建轴”，用：

```bash
npm run accept:onsite:strict
```

它同样会先做设备预检和 OAuth 授权，但最终要求 `event_axis=yes`；如果只是扫描兜底建轴，产品链路可能已经可用，但 strict 命令会失败并写出 `data/onsite-strict-device-preflight-report.json` 和 `data/onsite-strict-event-acceptance-report.json` 供排查开放平台事件订阅。

现场验收优先用这条命令。它会进入真实等待态，在真实会议轴出现后，通过开放标注接口 `POST /api/annotations` 自动写入一条标注，并一直等到“真实会议轴 + 标注落轴 + 页面实时广播”全部满足：

```bash
npm run accept:real-meeting
```

如果输出里提示当前用户 OAuth 过期，或者你希望事件订阅未投递时用当前账号会议扫描兜底，就用授权版：

```bash
npm run accept:real-meeting:auth
```

这条命令会打开飞书 OAuth，等待授权完成，然后触发当前账号会议扫描；扫描兜底能验证“直接开会后自动建真实轴并实时标注”，但不能证明开放平台会议开始事件已经投递。
`accept:real-meeting` 和 `accept:real-meeting:auth` 会写出 `data/real-meeting-acceptance-report.json`。

如果使用 `start:plain`，或想手动重新 armed 当前服务，可以用：

```bash
npm run monitor:real-demo -- --prepare
```

如果希望真实会议轴出现后，由命令行通过开放标注接口自动写入一条验收标注：

```bash
npm run monitor:real-demo -- --prepare --auto-mark
```

如果当前用户 OAuth 过期，并且你希望同时打开重新授权页：

```bash
npm run monitor:real-demo -- --prepare --open-auth
```

更短的授权兜底命令：

```bash
npm run auth:meeting-scan
```

它会打开飞书 OAuth，等待授权完成，随后立即触发一次当前账号会议扫描；如果你已经在会议中且事件订阅没有投递，这一步会尝试用扫描结果建真实会议轴。

如果希望命令在授权完成前不要继续进入验收监听：

```bash
npm run monitor:real-demo -- --prepare --open-auth --wait-auth --auto-mark
```

如果这次要严格验证“飞书会议开始事件本身完成建轴”，加上：

```bash
npm run monitor:real-demo -- --prepare --open-auth --wait-auth --auto-mark --require-event-axis
```

快速体检当前状态：

```bash
npm run monitor:real-demo -- --once
```

监控输出里的 `axis_mode` 和 `axis_source` 用来区分建轴来源：

```text
axis_mode=not_built axis_source=none
axis_mode=meeting_start_event axis_source=lark_ws_event
axis_mode=lark_passive_meeting_scan axis_source=lark_passive_meeting_scan
```

只有 `event_axis=yes`，或者 `axis_mode=meeting_start_event` 且 `axis_source=lark_ws_event` / `lark_http_event`，才说明飞书会议开始事件本身完成建轴；`lark_passive_meeting_scan` / `lark_probe_auto_search` 是当前账号会议扫描兜底，仍可验证“直接开会后自动建轴并实时标注”，但它不是事件订阅投递证据。`--require-event-axis` 会把扫描兜底完成视为未严格通过，适合排查开放平台事件订阅是否真的投递。

打开：

```text
http://localhost:8787
```

OAuth 回调和页面地址都建议使用 `localhost`，不要混用 `127.0.0.1`，否则飞书会把它们视为不同回调地址。

没有飞书凭证也能跑样例数据。真实接入时把 `.env` 里的 `LARK_APP_ID`、`LARK_APP_SECRET`、`LARK_BASE_URL` 配好。

## 开放会议会话协议

现在产品主路径不再被飞书官方事件投递卡住。统一抽象是：只要某个可信触发源确认“用户打开了真实会议”，就调用本地会话协议建轴；标注仍统一走 `POST /api/annotations`，转写仍会后导入。

```bash
curl -sS -X POST http://localhost:8787/api/meeting-session/start \
  -H 'content-type: application/json; charset=utf-8' \
  -d '{
    "platform": "lark",
    "title": "真实飞书会议",
    "meeting_url": "https://vc.feishu.cn/j/...",
    "start_time_ms": 1782442800000,
    "detector_source": "desktop_meeting_observer"
  }'
```

`start_time_ms` / `start_time` 最好由检测器提供；如果不传，服务端会把收到请求的时刻作为会话零点，这适合“会议刚打开就立即调用”的场景。页面里的“开放会话建轴”按钮调用的就是这条协议。未来可以把飞书官方事件、当前账号扫描、桌面端观察器、汉王端会议状态都接到同一个入口。

端侧也可以把建轴和第一条标注合并成一次请求。只要第一条标注显式带 `meeting_session`，服务端会先建开放会议轴，再按同一请求里的 `captured_at_ms` 把标注落到会议相对时间：

```bash
curl -sS -X POST http://localhost:8787/api/annotations \
  -H 'content-type: application/json; charset=utf-8' \
  -d '{
    "id": "epaper-mark-001",
    "source": "hanwang_epaper",
    "captured_at_ms": 1782442808000,
    "kind": "handwriting_trigger",
    "label": "why?",
    "meeting_session": {
      "platform": "lark",
      "meeting_id": "lark-session-001",
      "title": "真实飞书会议",
      "meeting_url": "https://vc.feishu.cn/j/...",
      "start_time_ms": 1782442800000,
      "detector_source": "hanwang_host_app"
    }
  }'
```

已有真实会议轴时，内联 `meeting_session` 默认不会覆盖当前轴；如果确实要切换到另一场会议，需要显式传 `force_meeting_session=true`。

验收分两层：

- `product_acceptance_complete=true`：会议轴已建立，开放标注已落轴，并通过 SSE 实时广播。开放会话轴可以完成这一层。
- `strict_event_acceptance_complete=true`：在产品验收基础上，会议轴还必须来自飞书 `meeting_start` 官方事件。这个只用于证明开放平台事件投递本身。

结束会话可调用：

```bash
curl -sS -X POST http://localhost:8787/api/meeting-session/end \
  -H 'content-type: application/json; charset=utf-8' \
  -d '{"end_time_ms":1782446400000}'
```

## 真实飞书账号接入

先在飞书开放平台创建应用，配置 OAuth 回调地址：

```text
http://localhost:8787/api/auth/lark/callback
```

如果飞书控制台不接受本地回调地址，就用 ngrok/cloudflared 暴露本地服务，并把 `.env` 里的 `LARK_REDIRECT_URI` 改成公网 HTTPS 回调。

`.env` 至少需要：

```text
LARK_BASE_URL=https://open.feishu.cn
LARK_APP_ID=cli_xxx
LARK_APP_SECRET=xxx
LARK_REDIRECT_URI=http://localhost:8787/api/auth/lark/callback
LARK_EVENT_CALLBACK_URL=
LARK_WS_EVENTS=1
```

然后重启服务。页面里有两个用户 OAuth 入口：

- `登录飞书账号（妙记）`：只请求妙记相关 scope，用于会后同步转写。
- `授权会议扫描兜底`：请求 `vc:meeting.search:read`，只用于飞书事件未投递时按当前账号扫描正在进行的会议。

实时会议建轴主路径不靠这两个登录按钮拿权限，而是靠飞书开放平台应用权限：在开发者后台开通并发布 `vc:meeting.all_meeting:readonly`，再通过长连接事件接收 `vc.meeting.all_meeting_started_v1`。

用户 OAuth token 过期后，页面会显示“授权已过期”。如果飞书返回过 `refresh_token`，可点“刷新授权”；如果当前授权没有 `refresh_token`，需要重新登录飞书账号。

授权入口支持按需追加 OAuth scope，例如启用当前账号会议扫描兜底：

```text
GET /api/auth/lark/start?scope=vc:meeting.search:read
```

页面里的“授权会议扫描兜底”会使用这个入口。授权 URL 会把 `.env` 的 `LARK_OAUTH_SCOPES` 和追加 scope 合并去重；修改 scope 后必须重新登录一次，已有 token 不会自动补 scope。
这个接口的 JSON 会显式返回 `callback_url`、`state_created_at` 和 `state_expires_at`，真实联调时优先确认 `callback_url` 是否与飞书开放平台配置完全一致，例如 `http://localhost:8787/api/auth/lark/callback`。

真实同步妙记时，支持直接粘贴：

```text
minute_token
```

或者粘贴完整妙记链接，服务端会尝试解析 token。

服务端同步接口：

```text
POST /api/lark/sync-minute
```

有用户 OAuth 授权时优先用 `user_access_token`；没有用户授权时回退到 `tenant_access_token`。

注意：飞书“导出妙记文字记录”权限可能需要在开放平台申请并发布到测试企业/租户。如果接口返回权限不足，不是登录失败，而是应用权限或妙记导出能力没有开通。

## 飞书事件接入

优先使用飞书/Lark 长连接事件模式。服务启动后会用 `LARK_APP_ID` / `LARK_APP_SECRET` 建立 WebSocket 长连接；这种模式本地即可接收事件，不需要公网 HTTPS、ngrok 或 cloudflared。页面“飞书事件接入”会显示长连接状态：

```text
长连接已连接：本机可直接接收飞书事件并自动建轴
```

如果这里显示 `failed` 或一直 `connecting`，通常需要在飞书开放平台把事件订阅模式切到长连接。demo 会接收长连接推来的所有事件，然后只处理会议相关 payload。

`GET /api/lark/config` 会返回 `ws_event_receiver.registered_event_types`。当前 demo 显式注册：

```text
vc.meeting.all_meeting_started_v1
vc.meeting.all_meeting_ended_v1
vc.meeting.join_meeting_v1
vc.meeting.leave_meeting_v1
vc.meeting.meeting_started_v1
vc.meeting.meeting_ended_v1
```

主路径仍是 `vc.meeting.all_meeting_started_v1`。如果企业会议开始事件未投递，但 `vc.meeting.join_meeting_v1` 可投递，demo 会把首次真实加入会议事件作为建轴兜底，并使用事件里的 `meeting.start_time` 计算时间轴零点。因此如果状态是 `connected` 且上述事件已注册，但 `recent_ws_event_count` 仍为 0，问题通常不在本地代码，而在开放平台还没有把事件投递给这个应用：检查事件订阅方式是否已保存为长连接、事件是否已添加并发布、应用可见范围、测试企业发布状态和权限开通。

本目标的主路径是“用户直接在飞书里开会”，需要订阅企业会议事件：

```text
vc.meeting.all_meeting_started_v1
vc.meeting.all_meeting_ended_v1
```

对应应用身份权限：

```text
vc:meeting.all_meeting:readonly
```

注意：`vc.meeting.meeting_started_v1` / `vc.meeting.meeting_ended_v1` 更偏向 OpenAPI 预约会议场景；`vc:reserve` 只用于 demo 自己创建会议，不是“用户直接开会自动建轴”的主路径。

排查事件是否真的进入 demo：

```text
GET /api/lark/config
GET /api/lark/delivery-diagnostics
GET /api/lark/events-log
GET /api/lark/ws-parser-self-test
GET /api/lark/meeting-event-requirements
GET /api/lark/real-demo/progress
GET /api/lark/real-demo/progress-stream
```

`delivery-diagnostics` 会把链路拆成应用凭证、长连接状态、本地已注册事件、真实投递证据、本机 HTTP 模拟证据和下一步动作。它还会返回 `real_meeting_event_audit` 和 `open_platform_checklist`：前者给出当前根因状态，例如 `no_event_delivery_observed`、`wrong_event_subscription`、`parser_payload_gap`、`event_delivery_ok`；后者把事件接收方式、开放平台订阅保存与版本发布、直开会议事件订阅、join/leave 兜底事件、直开会议权限发布、应用可见范围拆成机器可读核查项。

`events-log` 会记录最近 50 条事件的 `event_type`、来源通道、是否被判定为会议相关、是否进入时间轴处理，并持久化到 `data/lark-events-log.json`。敏感字段如 token、ticket、secret 会被隐藏。`ws-parser-self-test` 不改状态，只用 SDK parser 和本地 normalize 跑一条合成直开会议开始事件与一条 join 兜底事件；如果它通过但长连接 `connected` 后 `real_meeting_event_audit.status=no_event_delivery_observed`，说明本地接收和解析基本正常。此时如果你已经直接开过一次飞书会议但真实事件计数仍为 0，通常需要检查开放平台事件订阅是否已保存为长连接、订阅事件是否已添加并发布、权限发布、应用可见范围或测试企业发布状态。

`real-demo/progress` 是非阻塞的验收进度接口，适合开会时轮询。`real-demo/progress-stream` 是同一份进度证据的 SSE 版本，页面会用它实时刷新“等待真实轴 / 已建真实轴 / 已收到标注”的状态。它们都会返回当前 `status`、`completion_evidence`、`real_meeting_event_audit`、设备流状态、SSE 广播证据和下一步动作；`real-demo/monitor` 则是阻塞等待接口，适合自动化测试或“等到完成/超时再返回”的场景。

如果只想看现场是否可验收，读摘要接口：

```bash
curl -s http://localhost:8787/api/lark/real-demo/acceptance | jq
```

它会返回 `verdict`、`product_acceptance_complete`、`strict_event_acceptance_complete`、`recommended_command` 和 `missing_product_requirements`。这里把“产品验收”和“严格事件投递验收”分开：扫描兜底可以完成产品验收，因为用户确实是直接开启真实飞书会议后建真实轴并实时标注；但只有 `strict_event_acceptance_complete=true` 才证明飞书会议开始事件本身完成投递和建轴。

命令行可用 `npm run monitor:real-demo` 直接订阅 `real-demo/progress-stream`。它只在关键状态变化时打印一行，适合真实开会时观察 `events`、`real_axis`、`event_axis`、`axis_mode`、`axis_source`、`annotations`、`requirements` 是否发生变化；同一行也会显示 `auth` 和 `scan`，用于区分“飞书事件尚未投递”和“当前用户会议扫描兜底因 OAuth/scope 不可用”。正常 `npm start` 已经自动 armed；加 `--prepare` 会手动调用 `POST /api/lark/real-demo/prepare`，适合 `start:plain` 或需要重置当前等待窗口时使用。加 `--auto-mark` 会在真实会议轴出现且还没有标注时，通过开放标注接口 `POST /api/annotations` 自动写入一条验收标注；这条路径等价于未来墨水屏设备上传标注。加 `--require-event-axis` 会要求完成轴来自飞书会议开始事件，而不是扫描兜底。加 `--open-auth` 会在需要当前用户会议扫描兜底时打开飞书 OAuth 授权页；加 `--wait-auth` 会轮询等待 OAuth 真的可用，超时后以非 0 状态退出，避免在兜底不可用时误判。脚本不会代替你登录或确认授权。

如果需要关闭长连接，设置：

```text
LARK_WS_EVENTS=0
```

然后使用 HTTP webhook 模式。

## HTTP Webhook 备用模式

本地服务入口：

```text
POST /api/lark/events
```

飞书开放平台配置 HTTP 事件订阅时需要公网 HTTPS，可用 ngrok/cloudflared 暴露本地端口。PoC 阶段建议先关闭回调加密，只验证 plaintext event；加密回调可以在真实租户确认后补 AES 解密。

如果不用长连接，而要用 HTTP webhook 验证“直接开启真实飞书会议后自动建轴”，需要：

1. 用公网 HTTPS 暴露本地服务，例如 `https://your-tunnel.example.com`。
2. 在 `.env` 写入：

```text
LARK_EVENT_CALLBACK_URL=https://your-tunnel.example.com/api/lark/events
```

3. 在飞书开放平台事件订阅里把回调 URL 配成同一个地址。
4. 订阅 `vc.meeting.all_meeting_started_v1` / `vc.meeting.all_meeting_ended_v1`，并确认应用已开通 `vc:meeting.all_meeting:readonly`。

页面顶部的“链路验收状态”和“飞书事件回调”面板会显示当前真实链路状态。如果还是 `http://localhost:8787/api/lark/events`，真实飞书云端无法直接打到本机；长连接已连接但没有事件日志时，说明开放平台还没有把会议事件投递给这个应用。

机器可读验收状态：

```text
GET /api/readiness
GET /api/lark/acceptance-report
GET /api/lark/real-meeting-probe
POST /api/lark/real-meeting-probe/start
POST /api/lark/real-meeting-probe/auto-bind
POST /api/lark/real-meeting-probe/reset
```

`/api/readiness` 只把长连接投递，或公网 HTTPS webhook 投递，计为真实飞书事件。本机 `curl http://localhost:8787/api/lark/events` 只用于验证解析和对齐链路，页面会标记为“本机 HTTP 模拟轴”，不会让真实链路验收通过。
`/api/readiness` 同时返回 `passive_meeting_scan` 和兼容别名 `passive_scan`，二者内容相同，用于判断当前账号会议扫描兜底是否已开启、是否被 OAuth/scope/cooldown 卡住。

`/api/lark/acceptance-report` 是机器可读验收单，会返回主路径必须事件、权限中文名、官方文档链接、当前证据和下一项未完成动作。当前主路径必须项是：

```text
事件：vc.meeting.all_meeting_started_v1 / vc.meeting.all_meeting_ended_v1
权限：vc:meeting.all_meeting:readonly（获取所有视频会议信息）
```

验收报告里的 `current_validation` 专门表示“本次 probe 窗口”的严格证据：如果已经点击“启动验收探针”，只有 `probe_started_at` 之后收到的真实会议事件才会让 `real_event_after_probe=true`；如果事件未投递但当前账号会议扫描自动绑定成功，会显示 `auto_search_binding_after_probe=true`。历史事件日志和旧标注仍保留在 `current_evidence`，但不会让本次验收误判通过。

真实验收主流程：

1. 保持 demo 服务运行，确认页面显示长连接 connected，或已配置公网 HTTPS webhook。
2. 直接在飞书客户端开启一次会议；不需要先点页面按钮建轴。
3. 如果应用收到 `vc.meeting.all_meeting_started_v1`，页面会自动建出“飞书长连接事件轴”。如果只收到 `vc.meeting.join_meeting_v1`，demo 也会作为兜底建轴，但诊断会明确标记为会议上下文兜底事件。
4. 通过 `POST /api/annotations`、`POST /api/annotations/batch`，或页面“写入验收标注”，写入一条带 `captured_at_ms` 的标注，确认它实时出现在统一时间轴。
5. 页面里的“自动验收标注”默认关闭。开启后，真实会议轴建立时会通过同一个开放标注接口自动写入一条验收标注，用来验证“直接开会后时间轴实时出现标注”的闭环；正式接入墨水屏时仍应由设备调用 `POST /api/annotations`。
6. 页面“进入真实等待状态”和 `npm run start` 的自动 armed 不会立即创建会议轴；若当前用户 OAuth 具备 `vc:meeting.search:read`，会在后台扫描当前用户正在进行的真实会议并绑定。扫描建轴会在页面明确标记为兜底来源，不会冒充飞书 meeting_start 事件。
7. 如果长连接 connected 但开会后没有真实事件，优先检查事件订阅方式是否已保存为长连接、事件是否已添加并发布、应用可见范围，以及 `vc:meeting.all_meeting:readonly` 权限是否已发布；如果当前用户扫描兜底不可用，点击“授权会议扫描兜底”补 `vc:meeting.search:read`。

如果你先开了飞书会议，之后才点击“启动事件等待”或重置页面，probe 会只认点击之后的新事件。这时右侧会提示“最近真实会议开始事件早于本次等待窗口”。可以点“恢复最近真实轴”，它会从本地已经收到的飞书 start/end 事件恢复当前时间轴；这不是模拟会议，也不会创建新的飞书事件。

事件订阅卡住时，可以用 probe 自动扫描、手动点击“扫描我的真实会议”，或“按会议号/链接绑定真实会议”作为兜底验证。当前 demo 把“应用身份扫描租户近期会议”保留为实验入口；如果飞书返回 `Invalid access token for authorization`，不要把它当成可靠兜底，优先走当前用户 OAuth 的 `vc:meeting.search:read`。

按会议号/链接会调用：

```text
GET /open-apis/vc/v1/meetings/list_by_no
```

需要权限：

```text
vc:meeting.search:read
```

按会议号/链接兜底需要用户粘贴会议链接或会议号，因此不等同于“直接开会自动建轴”。probe 自动扫描不需要粘贴会议号，但依赖当前登录用户 OAuth scope `vc:meeting.search:read`；验收报告会明确区分 `real_event_after_probe` 和 `auto_search_binding_after_probe`。

收到会议开始事件时，服务端会自动创建一条实时会议时间轴；后续外部标注会按会议 `start_time` 落到相对时间上。会议事件字段在不同飞书事件里可能略有差异，demo 会尽量从 `meeting_id`、`topic/title`、`meeting_url/join_url`、`start_time`、`minute_token` 等常见字段中归一化。

本地可以用下面的命令模拟真实飞书会议开始事件：

```bash
curl -X POST http://localhost:8787/api/lark/events \
  -H 'content-type: application/json' \
  -d '{
    "header": {
      "event_id": "evt-real-meeting-start-001",
      "event_type": "vc.meeting.all_meeting_started_v1",
      "create_time": "1782442800"
    },
    "event": {
      "meeting": {
        "id": "om_real_001",
        "topic": "真实飞书会议",
        "url": "https://vc.feishu.cn/j/real-demo",
        "start_time": "1782442800"
      }
    }
  }'
```

## 同步妙记

浏览器页面里填 `minute_token`，调用：

```text
POST /api/lark/sync-minute
```

机器可读转写状态：

```text
GET /api/transcript-status
```

它会返回 `status`、`segment_count`、`meeting_ended`、`minute_token_present`、`next_action` 等字段。会议中没有转写是正常状态，`realtime_blocking=false`；实时标注会先进入时间轴，会议结束后再通过 `POST /api/lark/sync-minute` 或 `POST /api/import/lark-transcript` 补齐转写。

服务会请求：

```text
POST /open-apis/auth/v3/tenant_access_token/internal
GET  /open-apis/minutes/v1/minutes/:minute_token/transcript
```

注意：飞书“导出妙记”权限可能对新应用受限，所以 demo 同时保留了手动导入转写 JSON 的入口。

## 任意数据序列格式

先用一个很松的格式，后面可以替换成 HMP：

```json
[
  {
    "id": "mark-1",
    "time_ms": 195000,
    "kind": "handwriting_question",
    "label": "用户写下 why?",
    "payload": {
      "text": "why?"
    }
  }
]
```

也支持绝对时间：

```json
[
  {
    "id": "mark-2",
    "ts": "2026-06-26T02:18:15.000Z",
    "label": "圈出方案风险"
  }
]
```

如果使用绝对时间，系统会用会议 `start_time` 转成相对会议时间。

## 会中实时写入

P0 不依赖实时转写。推荐路径有两条：

- 主路径：用户直接开启飞书会议，`vc.meeting.all_meeting_started_v1` 到达后自动建轴；`vc.meeting.join_meeting_v1` 可作为事件订阅兜底建轴。
- 备选路径：点击“创建飞书会议”，由 demo 调用飞书 VC 预约 API 创建真实会议链接；这一步只进入“等待真实会议开始”状态，不会伪造会议开始时间。后续由真实会议开始事件、同步活跃会议、扫描当前账号真实会议或会议号绑定来建轴。

备选路径创建飞书会议需要应用身份权限：

```text
vc:reserve
```

可以去开放平台权限页申请：

```text
https://open.feishu.cn/app/<your_app_id>/auth?q=vc:reserve&op_from=openapi&token_type=tenant
```

创建成功后，页面会自动打开飞书会议链接；用户进入会议后可以点“同步活跃会议”把真实 active meeting id/start_time 拉回时间轴。如果事件订阅尚未投递，也可以点“扫描我的真实会议”按当前 OAuth 用户扫描近两小时会议并绑定真实 start_time；这个兜底需要用户授权 scope `vc:meeting.search:read`，修改 scope 后需要重新登录飞书账号。

如果还不能创建真实飞书会议，也可以先用页面折叠区里的按钮或下面的接口手动开始一个本地模拟会议。注意：本地模拟只用于验证对齐链路，不代表真实飞书会议；服务端不会让本地模拟结束事件默认写入真实飞书事件轴。

```bash
curl -X POST http://localhost:8787/api/live/start-meeting \
  -H 'content-type: application/json' \
  -d '{
    "title": "项目评审会",
    "meeting_url": "https://vc.feishu.cn/j/xxxxxxxx"
  }'
```

电子纸端每完成一个 mark 后，直接写入开放标注接口。如果端侧同时知道当前会议上下文，建议在第一条标注里带 `meeting_session`，这样不需要单独先调用 `POST /api/meeting-session/start`：

```bash
curl -X POST http://localhost:8787/api/annotations \
  -H 'content-type: application/json' \
  -d '{
    "id": "epaper-mark-001",
    "source": "hanwang_epaper",
    "captured_at_ms": 1782460012345,
    "kind": "handwriting_trigger",
    "label": "手写 why?",
    "text_candidates": ["why?", "why"],
    "intent": "question",
    "mark": {
      "action": "freehand"
    },
    "strokes": [],
    "payload": {
      "page_id": "optional-page-id"
    }
  }'
```

设备端接入前可以先跑只读预检。它会读取 `GET /api/annotation-ingest-info`，检查开放标注 CORS、时钟同步、SSE 状态和产品验收摘要，不会写入任何标注：

```bash
npm run device:preflight
```

命令会写出 `data/device-preflight-report.json`，便于确认墨水屏接入前置条件是否满足；它不写入标注，也不会改变会议轴。

如果要验证“设备写入后 SSE 实时广播能被页面同源收到”，可以跑端到端设备 roundtrip。它会先订阅 `GET /api/stream`，再写入一条开放标注，最后等待 SSE state 中出现同一个 annotation id：

```bash
npm run device:roundtrip
```

也可以用仓库里的设备端等价调试客户端。它只调用开放接口，不读取服务端内部状态，适合模拟未来汉王端上传一条标注：

```bash
npm run device:mark -- \
  --label="why?" \
  --captured-at-ms=1782460012345 \
  --wait-real-axis \
  --timeout-ms=60000
```

如果当前还没有真实会议轴，这条标注会先进入 `pending_real_meeting`；收到飞书真实会议开始事件或扫描兜底绑定后，同一个稳定 `id` 会被重算到真实会议轴。去掉 `--wait-real-axis` 可以只检查首次 ACK。

如果要用设备客户端模拟“第一条标注同时建轴”，加 `--meeting-session=true` 和会议上下文字段：

```bash
npm run device:mark -- \
  --label="why?" \
  --captured-at-ms=1782442808000 \
  --meeting-session=true \
  --meeting-id=lark-session-001 \
  --meeting-title="真实飞书会议" \
  --meeting-url="https://vc.feishu.cn/j/..." \
  --meeting-start-ms=1782442800000 \
  --wait-real-axis \
  --timeout-ms=60000
```

`device:mark` 未显式传 `--captured-at-ms` / `--captured-at` 时，会先读取 `GET /api/annotation-ingest-info` 里的 `clock_sync.endpoint`，调用 `GET /api/time?client_send_at_ms=...` 估算 `clock_offset_ms`，再把本机采集时间修正成 `captured_at_ms`。显式传入采集时间时不会二次修正；如需关闭默认同步，可加 `--sync-clock=false`。

真实设备接入时优先传 `captured_at_ms`，含义是“这次标注/笔迹完成采集的设备绝对时间”。服务端会根据会议 `start_time` 计算相对会议时间。兼容字段包括 `captured_at`、`timing.captured_at_ms`、`payload.timing.captured_at_ms`、`ink_end_at_ms`、`timestamp_ms`、`created_at_ms`、`device_time_ms`。如果只上传 stroke 点，点里的 `t` / `ts` / `timestamp_ms` 必须是 Unix 绝对时间，服务端会用最大点时间兜底推导采集结束时间。`time_ms` 只建议用于离线回放、批量导入或人工调试覆盖，因为它会被视为“已经换算好的会议相对时间”。

服务端会在每条标注里保留 `time_source` 和 `payload.timing`，页面也会显示“采集时间 / stroke 时间 / 显式相对时间 / 服务端收到时间”。如果标注是在会中采集、会后才上传，只要带了真实 `captured_at_ms`，它仍会落在会议结束之前。缺少采集时间时才会退回 `server_received_at`，页面会标出上传延迟和是否晚于会议结束；这类标注会被保存，但不会计入“真实会议轴已可靠标注”的验收统计。

会议轴只应由开放会议会话协议、真实飞书会议开始事件、真实会议查询绑定、或明确的本地模拟开始创建；点击“启动事件等待”或“启动验收探针”只启动等待/验收窗口，不会创建会议轴。只有带 `captured_at_ms` / `captured_at` / stroke 绝对时间的标注，才允许在真实会议事件到达前创建 pending 轴；缺少采集时间的标注不会创建 pending 轴。会议结束后后到的标注不会再自动开启新 pending 轴，而是先归到当前明确会议轴，除非请求显式传 `force_new_session` 或 `new_session` 且同时提供绝对采集时间。会议结束事件也不会再被服务端自动移动到最后一个标注之后：如果有标注显示在结束之后，说明该标注的采集时间确实晚于结束事件，或者设备上报的时间源不一致，需要优先排查 `captured_at_ms`。

如果设备本地时钟不可信，开会前可以请求 `GET /api/time?client_send_at_ms=<device_now_ms>` 计算设备时间与服务端时间的 offset。设备收到响应时记下 `client_receive_at_ms`，用：

```text
clock_offset_ms = server_time_ms - ((client_send_at_ms + client_receive_at_ms) / 2)
captured_at_ms = device_mark_end_ms + clock_offset_ms
```

再把修正后的绝对时间填入 `captured_at_ms`。`GET /api/annotation-ingest-info` 的 `clock_sync` 字段也会返回这套算法、同步 URL 和建议的最大时钟偏差。

运行时回归验证：

```bash
npm run check
```

这会启动一个隔离的临时服务，模拟 `vc.meeting.all_meeting_started_v1` 进入 `/api/lark/events`，随后通过 `POST /api/annotations` 写入开放标注，并确认 `GET /api/stream` 的 SSE 能实时推送时间轴更新。这个测试证明 demo 内部链路可用，但不等同于飞书云端真实投递验收；真实验收仍以页面诊断里的 `real_event_count > 0` 或 probe 收到真实 `vc.meeting.all_meeting_started_v1` 为准。

真实页面验收的默认路径是：`npm start` 后保持页面或 monitor 打开，直接在飞书客户端开会；如果飞书事件或扫描没有及时建轴，可以点击“开放会话建轴”或让未来桌面/汉王端调用 `POST /api/meeting-session/start`。页面显示真实会议轴后，确认设备流标注实时出现在统一时间轴。“启动验收探针”只是缩小本次验收窗口的可选辅助，不会建轴；“写入验收标注”只在真实会议轴已建立且未结束时可用，用来手动补一条开放标注。只有真实轴和开放标注实时落轴都通过，才算满足“直接开会 + 实时标注”的端到端目标。

如果飞书事件投递不稳定，页面在 probe waiting 时会自动尝试“扫描我的真实会议”绑定当前账号近期会议；也可以手动点击该按钮重试。这个兜底需要当前 OAuth token 包含 `vc:meeting.search:read`。如果页面显示缺少该 scope，点击“授权会议扫描兜底”，授权完成后重新开始一次真实验收。

页面还提供“应用身份扫描近期会议”，调用 `POST /api/lark/bind-tenant-latest-meeting`，尝试使用应用身份在近 15 分钟时间窗内按标题/时间扫描真实会议并建轴。飞书可能会拒绝租户 token 访问该搜索接口；如果诊断显示 `invalid_token`，改用“授权会议扫描兜底”获得当前用户 OAuth 的 `vc:meeting.search:read` 后，再走“扫描我的真实会议”。

验收等待状态会持久化到 `data/real-meeting-probe.json`；服务重启后页面仍会显示上一次等待窗口是否还有效。页面里的“清除验收状态”或 `POST /api/lark/real-meeting-probe/reset` 会把 probe 恢复到 idle，适合权限配置完成后重新开始一次干净验收，或回到 demo 样例查看。

`链路验收状态` 的最终 `ready` 以“真实会议轴已建立 + 开放标注接口写入过实时标注”为核心判定。`vc.meeting.all_meeting_started_v1` 是否真实投递仍保留为单独诊断项；如果开放会话协议或扫描兜底已经绑定真实会议并完成实时标注，整体链路也可以判定通过，但页面会明确显示该轴来源，不会冒充会议开始事件投递。

飞书用户 OAuth 不是实时标注链路的硬前提；它只影响会后妙记同步、当前账号会议手动扫描等兜底能力。直接会议事件主路径只依赖应用凭证、事件订阅和长连接/HTTP 回调。

机器可读 schema：

```text
GET /annotation-schema.json
```

设备端接入发现接口：

```text
GET /api/annotation-ingest-info
```

该接口会返回 `POST /api/annotations`、`POST /api/annotations/batch`、`GET /api/time`、`GET /api/stream`、`GET /api/stream/status`、兼容别名 `GET /api/stream-status`、`GET /api/lark/real-demo/acceptance`、schema URL、有效会议状态、`annotation_route` 和最小 payload 示例。开放标注接口支持 `OPTIONS` 预检和 CORS，方便后续墨水屏、独立调试页或 Android WebView 直接接入。

`annotation_route.mode` 用来告诉设备端下一条标注会怎样落轴：

- `create_pending_on_first_annotation`：当前没有真实会议轴，下一条带绝对采集时间的标注会先创建 pending 轴，等待飞书真实会议事件后重绑定。
- `append_to_pending_meeting`：已经处在 pending 轴，继续实时追加标注。
- `append_to_real_meeting`：已经有真实飞书会议轴，标注会按 `captured_at_ms` 归一化到真实会议时间。
- `append_to_local_simulation`：仅本地调试轴，不代表真实飞书会议。

如果当前只是内置 demo 样例，`current_meeting.meeting_id` 会返回 `null`，并在 `ignored_current_meeting` 里说明被忽略的样例轴，避免设备误以为 demo 会议是真实会议。设备也可以通过 header 传身份：`x-hmp-device-id` / `x-device-id` 会写入 `device_id`，`x-hmp-device-type` / `x-device-type` 可作为默认 `source`。

设备端或调试端可以读 `real_demo_acceptance_url` 判断整条产品链路是否闭环：`product_acceptance_complete=true` 表示真实会议轴、开放标注落轴、实时 SSE 广播都已经满足；`strict_event_acceptance_complete=true` 额外表示会议轴来自飞书会议开始事件，而不是扫描兜底。

`POST /api/annotations` 的响应会保留完整 `item` 和 `state`，同时提供给设备端使用的简短 `ack`：

```json
{
  "ack": {
    "accepted": true,
    "annotation_id": "epaper-mark-001",
    "operation": "created",
    "idempotent": true,
    "idempotency_key_source": "id",
    "warnings": [],
    "binding_state": "pending_real_meeting",
    "normalized_time_ms": 0,
    "time_source": "captured_at",
    "meeting_id": "pending-live-...",
    "pending_binding": true,
    "on_real_axis": false,
    "created_pending_timeline": true,
    "replaced_existing": false
  }
}
```

设备端应给每条标注生成稳定 `id`。网络抖动后重试同一个 `id` 会做 upsert：时间轴里仍只有一条标注，`ack.operation` 会变成 `updated`，`ack.replaced_existing=true`。这样墨水屏端可以安全重传，不会在会议时间轴上刷出重复标注。

如果设备端没有传 `id` / `annotation_id` / `mark_id` / `event_id`，服务端仍会接收并生成临时 id，但 `ack.idempotent=false`，`ack.warnings` 会包含 `missing_stable_id`。这种请求不适合自动重试，因为重传会被视为新标注。

如果首次 ACK 显示 `binding_state=pending_real_meeting`，设备端可以按稳定 id 查询后续绑定状态。推荐设备端使用单条状态接口：

```text
GET /api/annotation-status?id=epaper-mark-001
```

它会返回面向设备决策的扁平字段，例如 `status` / `next_action` / `on_real_axis` / `requires_device_captured_at` / `after_meeting_end_ms`。关键 `status` 枚举：

- `pending_real_meeting`：标注已存入 pending 轴，等待真实飞书会议开始事件或扫描兜底后重绑定。
- `real_axis_bound`：标注已经绑定真实会议轴，可视为实时标注链路成功。
- `needs_device_captured_at`：标注已保存，但缺少可靠采集时间；设备应使用同一个稳定 id 重传并带上 `captured_at_ms` 或绝对 stroke 时间。
- `after_meeting_end`：标注落点晚于已知会议结束时间，不计入真实会中标注。
- `not_found`：服务端没有找到该稳定 id。

旧的统计/调试查询仍保留：

```text
GET /api/annotation-bindings?id=epaper-mark-001
```

返回里的 `found` / `item.binding_state` / `item.on_real_axis` 也可以判断这条标注是否已经被真实飞书会议轴接管。`GET /api/annotation-ingest-info` 里的 `annotation_status_url` 和 `binding_lookup_url` 分别给出了两个 URL 模板。

旧接口仍兼容：

```text
POST /api/live/sequence-event
```

但后续墨水屏端建议只接 `POST /api/annotations`。

批量上传可以使用：

```text
POST /api/annotations/batch
```

请求体可以是数组，也可以是 `{ "annotations": [...] }` / `{ "items": [...] }`。单批最多 200 条，服务端会按数组顺序归一化，最后只广播一次 SSE；响应里的 `acks[]` 与每条输入一一对应。批量上传尤其要带每条标注自己的 `captured_at_ms`，否则会后同步只能按服务端收到批量请求的时间落点。

如果标注先到、飞书会议开始事件后到，且标注带了绝对采集时间，服务端会先创建一条 `pending_binding=true` 的“等待飞书会议事件”实时轴，让标注立即出现在页面上；一旦收到真实 `meeting_start`，会用原始 annotation payload 重新按真实会议 `start_time` 计算 `time_ms`，并把标注迁移到真实会议时间轴上。

真实会议验收 probe 等待期间，如果页面还停在旧的本地模拟轴或本机 HTTP 模拟轴上，带绝对采集时间的新开放标注会进入 pending 轴，不会继续污染旧轴。真实会议开始事件或真实会议查询绑定完成后，pending 标注会按真实 `start_time` 重算位置。

浏览器页面通过：

```text
GET /api/stream
```

保持 SSE 长连接。服务端收到新标注、飞书事件或新转写后，会重新对齐并推送最新 timeline。

这意味着：

```text
飞书会议开始事件/本地开始实时会议
        ↓
创建无转写实时会议时间轴
        ↓
电子纸写字/标注完成
        ↓
POST /api/annotations
        ↓
服务端按会议 start_time 归一化
        ↓
重算附近转写片段与会议事件
        ↓
页面实时出现标注和对应上下文
```

飞书妙记转写可能不是严格实时产物；如果飞书只在会后生成完整转写，这条链路仍可先实时显示“标注发生在会议第几分钟、附近有什么会议事件”，等会后妙记同步后再自动补齐对应发言上下文。
