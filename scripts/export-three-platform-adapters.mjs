#!/usr/bin/env node

import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import { chmod, copyFile, cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { basename, dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const args = new Map(process.argv.slice(2).map((raw) => {
  const [key, ...rest] = raw.replace(/^--/, '').split('=');
  return [key, rest.length ? rest.join('=') : 'true'];
}));
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = resolve(String(args.get('out-dir') ?? 'data/three-platform-adapters'));
const baseUrl = String(args.get('base-url') ?? process.env.MEETING_TIMELINE_BASE_URL ?? 'http://localhost:8787').replace(/\/+$/, '');
const offline = args.get('offline') === 'true';
const jsonOutput = args.get('json') === 'true';

async function command(program, commandArgs, options = {}) {
  const startedAt = Date.now();
  try {
    const result = await execFileAsync(program, commandArgs, {
      cwd: options.cwd ?? repoRoot,
      env: options.env ?? process.env,
      maxBuffer: 30 * 1024 * 1024,
    });
    return {
      ok: true,
      command: [program, ...commandArgs].join(' '),
      duration_ms: Date.now() - startedAt,
      stdout: String(result.stdout ?? '').trim(),
      stderr: String(result.stderr ?? '').trim(),
    };
  } catch (error) {
    return {
      ok: false,
      command: [program, ...commandArgs].join(' '),
      duration_ms: Date.now() - startedAt,
      exit_code: error?.code ?? 1,
      stdout: String(error?.stdout ?? '').trim(),
      stderr: String(error?.stderr ?? '').trim(),
      error: String(error?.message ?? error),
    };
  }
}

async function sha256(path) {
  const body = await readFile(path);
  return { path: relative(outDir, path), bytes: body.length, sha256: createHash('sha256').update(body).digest('hex') };
}

async function writeLaunchers(desktopDir, tarballName) {
  const macPath = join(desktopDir, 'start-macos.command');
  const windowsPath = join(desktopDir, 'start-windows.cmd');
  const installMacPath = join(desktopDir, 'install-macos.command');
  const installWindowsPath = join(desktopDir, 'install-windows.cmd');
  await writeFile(installMacPath, [
    '#!/bin/sh',
    'set -eu',
    'cd "$(dirname "$0")"',
    `npm install -g "./${tarballName}"`,
    'meeting-timeline-desktop-adapter --diagnose --platforms=teams,zoom',
    '',
  ].join('\n'), 'utf8');
  await writeFile(macPath, [
    '#!/bin/sh',
    'set -eu',
    'BASE_URL="${MEETING_TIMELINE_BASE_URL:-http://localhost:8787}"',
    'exec meeting-timeline-desktop-adapter --base-url="$BASE_URL" --platforms=teams,zoom --interval-ms=750 --evidence-file="${MEETING_TIMELINE_EVIDENCE_FILE:-./desktop-adapter-evidence.jsonl}"',
    '',
  ].join('\n'), 'utf8');
  await writeFile(installWindowsPath, [
    '@echo off',
    'cd /d %~dp0',
    `call npm install -g ".\\${tarballName}"`,
    'if errorlevel 1 exit /b %errorlevel%',
    'call meeting-timeline-desktop-adapter --diagnose --platforms=teams,zoom',
    '',
  ].join('\r\n'), 'utf8');
  await writeFile(windowsPath, [
    '@echo off',
    'if "%MEETING_TIMELINE_BASE_URL%"=="" set MEETING_TIMELINE_BASE_URL=http://localhost:8787',
    'call meeting-timeline-desktop-adapter --base-url="%MEETING_TIMELINE_BASE_URL%" --platforms=teams,zoom --interval-ms=750 --evidence-file="desktop-adapter-evidence.jsonl"',
    '',
  ].join('\r\n'), 'utf8');
  await Promise.all([chmod(macPath, 0o755), chmod(installMacPath, 0o755)]);
  return [installMacPath, macPath, installWindowsPath, windowsPath];
}

async function writeReadme(tarballName) {
  const path = join(outDir, 'README.md');
  await writeFile(path, [
    '# Google Meet / Microsoft Teams / Zoom 适配器交付包',
    '',
    '这不是要求业务项目继续实现 adapter 的接口文档，而是可直接安装的运行时交付物。',
    '',
    '## Web 会议',
    '',
    '1. 打开 `chrome://extensions` 或 `edge://extensions`。',
    '2. 开启开发者模式，选择“加载已解压的扩展程序”。',
    '3. 选择本目录下的 `browser-extension`。',
    '4. 打开 Google Meet、Teams Web 或 Zoom Web 会议；content script 会自动监听会议开始、结束和稳定发言人。',
    '5. 扩展 popup 中的 `Capture active` / `Capture ended` 只用于验收取证，不是启动自动监听所需的按钮。真实验收时写入一条设备标注，会议结束后再用 `Export evidence` 导出证据。',
    '',
    '同一扩展已经内置三个平台的 URL 匹配、DOM/无障碍语义归一化、开始/结束检测、发言人滤波、时间轴写入和证据导出。',
    '',
    `同一个 SDK tarball 也内置了这份预构建扩展。安装 \`desktop-host/${tarballName}\` 后可运行 \`meeting-timeline-adapters --out-dir=./meeting-timeline-browser-extension\` 导出，不需要回到源码仓库构建。`,
    '',
    '## Teams / Zoom 桌面客户端',
    '',
    `SDK 安装包位于 \`desktop-host/${tarballName}\`。macOS 双击 \`install-macos.command\`，Windows 运行 \`install-windows.cmd\`。`,
    '',
    '安装后：',
    '',
    '- macOS 双击 `start-macos.command`。首次运行需要在“系统设置 > 隐私与安全性 > 辅助功能”允许终端或宿主应用。',
    '- Windows 运行 `start-windows.cmd`，并保证 adapter 与 Teams/Zoom 处于同一登录用户和权限级别。',
    '- 默认连接 `http://localhost:8787`；远程服务可通过 `MEETING_TIMELINE_BASE_URL` 修改。',
    '',
    'desktop host 已包含进程/窗口扫描、会议控件识别、两帧开始去抖、四帧结束去抖、两帧发言人稳定滤波和证据 JSONL，不需要调用方自己构造 `windows[]`。',
    '',
    '## 发布状态',
    '',
    '构建、安装启动和合成状态机测试通过不等于真实会议生产验收。以 `release-manifest.json` 为准：只有真实会议中的开始、发言人、标注、结束四类证据齐全后，单个平台的 `production_ready` 才能变为 `true`。',
    '',
    '在源码仓库中可运行自动验收器：`npm run meeting-platform:live-acceptance -- --platform=google-meet --synthetic-audio=true`，平台也可传 `teams` 或 `zoom`；Zoom 默认打开官方 `zoom.us/test` 测试会议入口。测试人员只需在启动的隔离浏览器中登录、加入会议并离会，不需要实现 adapter；macOS 的 `--synthetic-audio=true` 会注入固定语音并自动产生稳定发言段，其他系统可传 `--fake-audio-file=/absolute/path/to/mono.wav`。显式授权测试程序改变外部会议状态时，可加 `--auto-join=true --auto-leave=true --allow-external-actions=true` 自动处理可见的入会、电脑音频和离会控件。三平台完成后运行 `npm run meeting-platform:live-acceptance:verify`。',
    '',
  ].join('\n'), 'utf8');
  return path;
}

async function main() {
  await rm(outDir, { recursive: true, force: true });
  const extensionBuildDir = join(outDir, '.browser-extension-build');
  const extensionDir = join(outDir, 'browser-extension');
  const desktopDir = join(outDir, 'desktop-host');
  await Promise.all([mkdir(extensionBuildDir, { recursive: true }), mkdir(extensionDir, { recursive: true }), mkdir(desktopDir, { recursive: true })]);

  const extensionArgs = [
    'scripts/export-meeting-app-extension.mjs',
    `--out-dir=${extensionBuildDir}`,
    `--base-url=${baseUrl}`,
    '--platforms=google-meet,teams,zoom',
    '--build=true',
    `--offline=${offline}`,
    '--json=true',
  ];
  const extensionCommand = await command(process.execPath, extensionArgs, {
    env: { ...process.env, NODE_ENV: 'production' },
  });
  let extensionReport = null;
  try { extensionReport = JSON.parse(extensionCommand.stdout); } catch {}
  if (!extensionCommand.ok || extensionReport?.ok !== true) {
    throw new Error(`Browser extension build failed: ${extensionCommand.stderr || extensionCommand.stdout || extensionCommand.error}`);
  }

  const extensionFiles = ['manifest.json', 'content-script.js', 'live-capture.js', 'background.js', 'popup.js', 'popup.html', 'popup.css', 'README.md'];
  await Promise.all(extensionFiles.map((file) => copyFile(join(extensionBuildDir, file), join(extensionDir, file))));
  const extensionZip = join(outDir, 'meeting-timeline-browser-extension.zip');
  const zip = await command('zip', ['-q', '-r', extensionZip, '.'], { cwd: extensionDir });
  if (!zip.ok) throw new Error(`Extension zip failed: ${zip.stderr || zip.error}`);

  const sdkPackageBuildDir = join(outDir, '.sdk-package-build');
  await cp(join(repoRoot, 'packages/meeting-timeline-sdk'), sdkPackageBuildDir, { recursive: true });
  const bundledExtensionDir = join(sdkPackageBuildDir, 'runtime', 'browser-extension');
  await mkdir(dirname(bundledExtensionDir), { recursive: true });
  await cp(extensionDir, bundledExtensionDir, { recursive: true });
  const stagedPackageJsonPath = join(sdkPackageBuildDir, 'package.json');
  const stagedPackageJson = JSON.parse(await readFile(stagedPackageJsonPath, 'utf8'));
  stagedPackageJson.meetingTimelineAdapters = {
    platforms: ['google_meet', 'microsoft_teams', 'zoom'],
    browser_extension: 'runtime/browser-extension',
    desktop_host_binary: 'meeting-timeline-desktop-adapter',
    browser_export_binary: 'meeting-timeline-adapters',
  };
  await writeFile(stagedPackageJsonPath, `${JSON.stringify(stagedPackageJson, null, 2)}\n`, 'utf8');

  const pack = await command('npm', ['pack', sdkPackageBuildDir, '--json', `--pack-destination=${desktopDir}`]);
  if (!pack.ok) throw new Error(`SDK pack failed: ${pack.stderr || pack.error}`);
  const packRows = JSON.parse(pack.stdout);
  const tarballName = packRows[0]?.filename;
  if (!tarballName) throw new Error('SDK pack did not report a tarball filename');
  const tarballPath = join(desktopDir, basename(tarballName));
  const tarList = await command('tar', ['-tzf', tarballPath]);
  const requiredPackageFiles = [
    'package/bin/meeting-timeline-desktop-adapter.mjs',
    'package/cli/desktop-meeting-adapter.mjs',
    'package/adapters/desktop-meeting-host.mjs',
    'package/adapters/desktop-scanners/macos.jxa',
    'package/adapters/desktop-scanners/windows.ps1',
    'package/bin/meeting-timeline-adapters.mjs',
    'package/runtime/browser-extension/manifest.json',
    'package/runtime/browser-extension/content-script.js',
    'package/runtime/browser-extension/live-capture.js',
    'package/runtime/browser-extension/background.js',
  ];
  const missingPackageFiles = requiredPackageFiles.filter((file) => !tarList.stdout.split('\n').includes(file));
  if (!tarList.ok || missingPackageFiles.length > 0) {
    throw new Error(`Desktop package is missing: ${missingPackageFiles.join(', ')}`);
  }

  const desktopInstallDir = join(outDir, '.desktop-host-install');
  const installArgs = ['install', '-g', '--prefix', desktopInstallDir, tarballPath, '--no-audit', '--no-fund'];
  if (offline) installArgs.push('--offline');
  const desktopInstall = await command('npm', installArgs);
  if (!desktopInstall.ok) throw new Error(`Desktop package install failed: ${desktopInstall.stderr || desktopInstall.error}`);
  const desktopBinary = process.platform === 'win32'
    ? join(desktopInstallDir, 'meeting-timeline-desktop-adapter.cmd')
    : join(desktopInstallDir, 'bin', 'meeting-timeline-desktop-adapter');
  const desktopStart = await command(desktopBinary, [
    '--diagnose',
    `--base-url=${baseUrl}`,
    '--platforms=teams,zoom',
  ]);
  let desktopDiagnosis = null;
  try { desktopDiagnosis = JSON.parse(desktopStart.stdout); } catch {}
  const desktopInstallStartOk = desktopStart.ok
    && desktopDiagnosis?.schema === 'desktop_meeting_adapter_diagnosis'
    && desktopDiagnosis?.platforms?.includes('microsoft_teams')
    && desktopDiagnosis?.platforms?.includes('zoom');
  if (!desktopInstallStartOk) throw new Error(`Desktop package start failed: ${desktopStart.stderr || desktopStart.stdout || desktopStart.error}`);

  const packagedExtensionDir = join(outDir, '.packaged-extension-smoke');
  const adaptersBinary = process.platform === 'win32'
    ? join(desktopInstallDir, 'meeting-timeline-adapters.cmd')
    : join(desktopInstallDir, 'bin', 'meeting-timeline-adapters');
  const packagedExtensionExport = await command(adaptersBinary, [
    `--out-dir=${packagedExtensionDir}`,
    '--json=true',
  ]);
  let packagedExtensionReport = null;
  let packagedExtensionManifest = null;
  try { packagedExtensionReport = JSON.parse(packagedExtensionExport.stdout); } catch {}
  try { packagedExtensionManifest = JSON.parse(await readFile(join(packagedExtensionDir, 'manifest.json'), 'utf8')); } catch {}
  const packagedExtensionExportOk = packagedExtensionExport.ok
    && packagedExtensionReport?.ok === true
    && packagedExtensionManifest?.manifest_version === 3
    && packagedExtensionReport?.platforms?.length === 3;
  if (!packagedExtensionExportOk) {
    throw new Error(`Packaged browser extension export failed: ${packagedExtensionExport.stderr || packagedExtensionExport.stdout || packagedExtensionExport.error}`);
  }
  await rm(desktopInstallDir, { recursive: true, force: true });
  await rm(packagedExtensionDir, { recursive: true, force: true });
  await rm(sdkPackageBuildDir, { recursive: true, force: true });

  const launchers = await writeLaunchers(desktopDir, basename(tarballPath));
  const readmePath = await writeReadme(basename(tarballPath));
  const artifactPaths = [
    ...extensionFiles.map((file) => join(extensionDir, file)),
    extensionZip,
    tarballPath,
    ...launchers,
    readmePath,
  ];
  const artifacts = await Promise.all(artifactPaths.map(sha256));
  const report = {
    type: 'three_platform_adapter_release',
    schema: 'three_platform_adapter_release',
    schema_version: 1,
    generated_at: new Date().toISOString(),
    ok: true,
    delivery_ready: true,
    production_ready: false,
    base_url: baseUrl,
    out_dir: outDir,
    platforms: ['google_meet', 'microsoft_teams', 'zoom'],
    browser_extension: {
      build_ok: true,
      install_dir: relative(outDir, extensionDir),
      zip: relative(outDir, extensionZip),
      manifest: extensionReport.manifest,
      build_gate: extensionReport.build_gate,
    },
    desktop_host: {
      package_ok: true,
      install_start_ok: desktopInstallStartOk,
      package: relative(outDir, tarballPath),
      supported_platforms: ['microsoft_teams', 'zoom'],
      supported_os: ['darwin', 'win32'],
      required_package_files: requiredPackageFiles,
      missing_package_files: missingPackageFiles,
      install_gate: {
        install: { ok: desktopInstall.ok, duration_ms: desktopInstall.duration_ms, command: desktopInstall.command },
        start: { ok: desktopStart.ok, duration_ms: desktopStart.duration_ms, command: desktopStart.command },
        diagnosis: desktopDiagnosis,
      },
      bundled_browser_extension: {
        ok: packagedExtensionExportOk,
        command: packagedExtensionExport.command,
        platforms: packagedExtensionReport.platforms,
        manifest_version: packagedExtensionManifest.manifest_version,
      },
    },
    adapters: [
      { platform: 'google_meet', installable: true, startable: true, web_ready: true, desktop_ready: false, real_meeting_accepted: false, production_ready: false },
      { platform: 'microsoft_teams', installable: true, startable: true, web_ready: true, desktop_ready: true, real_meeting_accepted: false, production_ready: false },
      { platform: 'zoom', installable: true, startable: true, web_ready: true, desktop_ready: true, real_meeting_accepted: false, production_ready: false },
    ],
    artifacts,
    remaining_gate: 'Capture real meeting_started, speaker_started, realtime annotation, and meeting_ended evidence for each platform.',
  };
  await writeFile(join(outDir, 'release-manifest.json'), `${JSON.stringify(report, null, 2)}\n`, 'utf8');
  await rm(extensionBuildDir, { recursive: true, force: true });
  if (jsonOutput) console.log(JSON.stringify(report, null, 2));
  else {
    console.log(`three_platform_adapter_release | ok=yes | dir=${outDir}`);
    console.log(`browser_extension=${relative(repoRoot, extensionDir)}`);
    console.log(`desktop_package=${relative(repoRoot, tarballPath)}`);
    console.log('production_ready=no (real meeting evidence pending)');
  }
  return report;
}

main().catch((error) => {
  const report = {
    type: 'three_platform_adapter_release',
    ok: false,
    out_dir: outDir,
    error: String(error?.message ?? error),
  };
  console.error(jsonOutput ? JSON.stringify(report, null, 2) : `three_platform_adapter_release | ok=no | error=${report.error}`);
  process.exitCode = 2;
});
