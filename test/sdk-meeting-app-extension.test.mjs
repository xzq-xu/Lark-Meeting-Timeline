import assert from 'node:assert/strict';

import {
  MEETING_APP_EXTENSION_PLATFORM_KEYS,
  MEETING_APP_EXTENSION_PROFILES,
  assertMeetingAppExtensionScaffold,
  buildMeetingAppContentScriptManifest,
  buildMeetingAppExtensionBackgroundSource,
  buildMeetingAppExtensionBuildSource,
  buildMeetingAppExtensionContentScriptSource,
  buildMeetingAppExtensionInstallPlan,
  buildMeetingAppExtensionMatchPatterns,
  buildMeetingAppExtensionPackageJson,
  buildMeetingAppExtensionScaffold,
  buildMeetingAppExtensionScaffoldAcceptanceReport,
  meetingAppExtensionProfile,
  normalizeMeetingAppExtensionPlatform,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-extension.mjs';

assert.deepEqual(MEETING_APP_EXTENSION_PLATFORM_KEYS, [
  'google_meet',
  'microsoft_teams',
  'zoom',
  'lark',
  'webex',
]);
assert.equal(MEETING_APP_EXTENSION_PROFILES.google_meet.matches.includes('https://meet.google.com/*'), true);
assert.equal(normalizeMeetingAppExtensionPlatform('google-meet'), 'google_meet');
assert.equal(normalizeMeetingAppExtensionPlatform('teams'), 'microsoft_teams');
assert.equal(normalizeMeetingAppExtensionPlatform('feishu'), 'lark');

const google = meetingAppExtensionProfile('google-meet');
assert.equal(google.platform, 'google_meet');
assert.deepEqual(google.matches, ['https://meet.google.com/*']);

const patterns = buildMeetingAppExtensionMatchPatterns();
assert.equal(patterns.type, 'meeting_app_extension_match_patterns');
assert.equal(patterns.platforms.length, 5);
assert.equal(patterns.matches.includes('https://meet.google.com/*'), true);
assert.equal(patterns.matches.includes('https://teams.microsoft.com/*'), true);
assert.equal(patterns.matches.includes('https://*.zoom.com/*'), true);
assert.equal(patterns.matches.includes('https://vc.feishu.cn/*'), true);
assert.equal(patterns.matches.includes('https://*.webex.com/*'), true);
assert.equal(patterns.matches.includes('<all_urls>'), false);
assert.deepEqual(patterns.content_scripts[0].js, ['meeting-app-content-script.bundle.js']);

const customPatterns = buildMeetingAppExtensionMatchPatterns(['google_meet', 'google-meet'], {
  js: 'content.js',
  extraMatches: ['https://meet.google.com/*', 'https://meet.example.test/*'],
  extraHostPermissions: ['https://meet.example.test/*'],
});
assert.deepEqual(customPatterns.platforms, ['google_meet']);
assert.deepEqual(customPatterns.content_scripts[0].js, ['content.js']);
assert.equal(customPatterns.matches.filter((item) => item === 'https://meet.google.com/*').length, 1);
assert.equal(customPatterns.matches.includes('https://meet.example.test/*'), true);
assert.equal(customPatterns.host_permissions.includes('https://meet.example.test/*'), true);

const googleManifest = buildMeetingAppContentScriptManifest({
  platforms: ['google_meet'],
  js: ['content.js'],
  permissions: ['storage', 'storage'],
});
assert.equal(googleManifest.manifest_version, 3);
assert.deepEqual(googleManifest.permissions, ['storage']);
assert.deepEqual(googleManifest.host_permissions, ['https://meet.google.com/*']);
assert.deepEqual(googleManifest.content_scripts[0].matches, ['https://meet.google.com/*']);
assert.deepEqual(googleManifest.content_scripts[0].js, ['content.js']);
assert.equal(googleManifest.content_scripts[0].run_at, 'document_idle');

const plan = buildMeetingAppExtensionInstallPlan({
  platforms: ['google_meet', 'microsoft_teams'],
  js: ['content.js'],
});
assert.equal(plan.type, 'meeting_app_extension_install_plan');
assert.deepEqual(plan.platforms, ['google_meet', 'microsoft_teams']);
assert.equal(plan.content_script_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-content-script');
assert.equal(plan.browser_runtime_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-browser-runtime');
assert.equal(plan.snapshot_recorder_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-snapshot-recorder');
assert.equal(plan.launch_gate_adapter, '@ai-annotation/meeting-timeline-sdk/adapters/meeting-app-gate');
assert.equal(plan.runtime_contract.timestamp_field, 'captured_at_ms');
assert.deepEqual(plan.manifest.content_scripts[0].js, ['content.js']);

const contentScriptSource = buildMeetingAppExtensionContentScriptSource({
  platforms: ['google_meet', 'microsoft_teams'],
});
assert.match(contentScriptSource, /installMeetingAppContentScriptBridge/);
assert.match(contentScriptSource, /meet\.google\.com/);
assert.match(contentScriptSource, /teams\.microsoft\.com/);
assert.match(contentScriptSource, /startRuntime: true/);

const backgroundSource = buildMeetingAppExtensionBackgroundSource({
  baseUrl: 'https://timeline.example.com/',
});
assert.match(backgroundSource, /const BASE_URL = "https:\/\/timeline\.example\.com";/);
assert.match(backgroundSource, /meeting_timeline\.extension_attached/);
assert.match(backgroundSource, /meeting_timeline\.extension_status/);
assert.match(backgroundSource, /STATUS_STORAGE_KEY/);
assert.match(backgroundSource, /setStorageValue/);
assert.match(backgroundSource, /\/api\/meeting-session\/start/);
assert.match(backgroundSource, /\/api\/annotations\/batch/);

const packageJson = buildMeetingAppExtensionPackageJson({
  packageName: 'demo-meeting-extension',
  sdkDependencyVersion: 'workspace:*',
});
assert.equal(packageJson.name, 'demo-meeting-extension');
assert.equal(packageJson.scripts.build, 'node build.mjs');
assert.equal(packageJson.dependencies['@ai-annotation/meeting-timeline-sdk'], 'workspace:*');
assert.equal(packageJson.devDependencies.esbuild, '^0.25.0');

const buildSource = buildMeetingAppExtensionBuildSource({
  outputScript: 'content-script.js',
  backgroundScript: 'background.js',
});
assert.match(buildSource, /from 'esbuild'/);
assert.match(buildSource, /src\/content-script\.entry\.mjs/);
assert.match(buildSource, /src\/background\.entry\.mjs/);
assert.match(buildSource, /format: "iife"/);
assert.match(buildSource, /format: "esm"/);

const scaffold = buildMeetingAppExtensionScaffold({
  platforms: ['google_meet'],
  baseUrl: 'https://timeline.example.com',
  outputScript: 'content-script.js',
});
assert.equal(scaffold.type, 'meeting_app_extension_scaffold');
assert.equal(scaffold.validation.uses_all_urls, false);
assert.equal(scaffold.manifest.background.service_worker, 'background.js');
assert.equal(scaffold.manifest.background.type, 'module');
assert.equal(scaffold.manifest.permissions.includes('storage'), true);
assert.equal(scaffold.manifest.host_permissions.includes('https://timeline.example.com/*'), true);
assert.deepEqual(scaffold.manifest.content_scripts[0].js, ['content-script.js']);
assert.equal(scaffold.bundle.background_input, 'src/background.entry.mjs');
assert.equal(scaffold.files.find((file) => file.path === 'package.json').mime, 'application/json');
assert.match(scaffold.files.find((file) => file.path === 'package.json').content, /"esbuild"/);
assert.match(scaffold.files.find((file) => file.path === 'build.mjs').content, /content-script\.js/);
assert.equal(scaffold.files.find((file) => file.path === 'manifest.json').mime, 'application/json');
assert.match(scaffold.files.find((file) => file.path === 'src/content-script.entry.mjs').content, /meeting-app-content-script/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /runtimeApi\(\)\?\.onMessage/);
assert.match(scaffold.files.find((file) => file.path === 'src/background.entry.mjs').content, /extension_status/);
assert.match(scaffold.files.find((file) => file.path === 'README.md').content, /npm run build/);

const scaffoldReport = buildMeetingAppExtensionScaffoldAcceptanceReport(scaffold);
assert.equal(scaffoldReport.accepted, true);
assert.deepEqual(scaffoldReport.platforms, ['google_meet']);
assert.equal(scaffoldReport.accepted_platform_count, 1);
assert.equal(scaffoldReport.manifest.uses_all_urls, false);
assert.equal(scaffoldReport.manifest.permissions.includes('storage'), true);
assert.deepEqual(scaffoldReport.issues, []);
assert.equal(assertMeetingAppExtensionScaffold(scaffold).accepted, true);

const unsafeScaffold = buildMeetingAppExtensionScaffold({ platforms: ['google_meet'] });
const unsafeManifestFile = unsafeScaffold.files.find((file) => file.path === 'manifest.json');
const unsafeManifest = JSON.parse(unsafeManifestFile.content);
unsafeManifest.host_permissions = ['<all_urls>'];
unsafeManifestFile.content = `${JSON.stringify(unsafeManifest, null, 2)}\n`;
const unsafeReport = buildMeetingAppExtensionScaffoldAcceptanceReport(unsafeScaffold);
assert.equal(unsafeReport.accepted, false);
assert.equal(unsafeReport.issues.some((item) => item.code === 'overbroad_host_permission'), true);
assert.throws(
  () => assertMeetingAppExtensionScaffold(unsafeScaffold),
  /Meeting app extension scaffold acceptance failed/,
);

assert.throws(
  () => meetingAppExtensionProfile('unknown-meeting'),
  /Unsupported meeting app extension platform/,
);

console.log('ok meeting app extension');
