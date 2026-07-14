import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const repoRoot = new URL('..', import.meta.url);
const tmpDir = await mkdtemp(join(tmpdir(), 'meeting-platform-subscription-handoff-script-'));
const inputDir = join(tmpDir, 'input');
const reportFile = join(tmpDir, 'subscription-handoff-report.json');
const outDir = join(tmpDir, 'handoffs');
const baseUrl = 'https://timeline.example.com';

await mkdir(inputDir, { recursive: true });
const subscriptionsFile = join(inputDir, 'subscriptions.json');
await writeFile(subscriptionsFile, JSON.stringify({
  google_meet: {
    targetResource: '//cloudidentity.googleapis.com/users/me',
    pubsubTopic: 'projects/demo/topics/meet-events',
  },
  microsoft_teams: {
    joinWebUrl: 'https://teams.microsoft.com/l/meetup-join/fixture',
    notificationUrl: `${baseUrl}/api/platform-events/teams`,
    clientState: 'client-state',
    now: '2026-06-26T02:00:00.000Z',
  },
  zoom: {
    webhookUrl: `${baseUrl}/api/platform-events/zoom`,
    accountId: 'zoom-account-1',
  },
  webex: {
    targetUrl: `${baseUrl}/api/platform-events/webex`,
    secret: 'webex-secret',
    ownedBy: 'org',
  },
}, null, 2), 'utf8');

const env = {
  ...process.env,
  GOOGLE_PUBSUB_OIDC_AUDIENCE: `${baseUrl}/api/platform-events/google-meet`,
  GOOGLE_PUBSUB_SERVICE_ACCOUNT_EMAIL: 'pubsub@example.iam.gserviceaccount.com',
  MICROSOFT_GRAPH_CLIENT_STATE: 'client-state',
  ZOOM_WEBHOOK_SECRET_TOKEN: 'zoom-secret',
  WEBEX_WEBHOOK_SECRET: 'webex-secret',
};

const { stdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-subscription-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=google-meet,teams,zoom,webex,lark',
  `--subscriptions-file=${subscriptionsFile}`,
  `--report-file=${reportFile}`,
  `--out-dir=${outDir}`,
], {
  cwd: repoRoot,
  env,
});

assert.match(stdout, /meeting_platform_subscription_handoff_report/);
assert.match(stdout, /ready_to_create=4\/5/);
assert.match(stdout, /requests=10/);
const report = JSON.parse(await readFile(reportFile, 'utf8'));
assert.equal(report.type, 'meeting_platform_subscription_handoff_report');
assert.equal(report.ok, true);
assert.equal(report.ready_to_create_count, 4);
assert.equal(report.manual_setup_count, 1);
assert.equal(report.request_count, 10);
assert.equal(report.rows.find((row) => row.platform === 'google_meet').status, 'subscription_request_ready');
assert.equal(report.rows.find((row) => row.platform === 'lark').status, 'manual_setup');
assert.equal(report.written_files.length, 5);

const webexHandoff = JSON.parse(await readFile(join(outDir, 'webex.json'), 'utf8'));
assert.equal(webexHandoff.requests.length, 7);
assert.equal(webexHandoff.requests[0].path, '/v1/webhooks');

const { stdout: jsonStdout } = await execFileAsync(process.execPath, [
  'scripts/meeting-platform-subscription-handoff.mjs',
  `--base-url=${baseUrl}`,
  '--platforms=zoom',
  `--subscriptions-file=${subscriptionsFile}`,
  '--json=true',
], {
  cwd: repoRoot,
  env: {
    ...process.env,
  },
});
const blockedReport = JSON.parse(jsonStdout);
assert.equal(blockedReport.ok, false);
assert.equal(blockedReport.security_blocked_count, 1);
assert.equal(blockedReport.rows[0].missing_env.includes('ZOOM_WEBHOOK_SECRET_TOKEN'), true);

await assert.rejects(
  execFileAsync(process.execPath, [
    'scripts/meeting-platform-subscription-handoff.mjs',
    `--base-url=${baseUrl}`,
    '--platforms=zoom',
    `--subscriptions-file=${subscriptionsFile}`,
    '--fail-on-not-ready=true',
  ], {
    cwd: repoRoot,
    env: {
      ...process.env,
    },
  }),
  (error) => error.code === 2,
);

console.log('ok meeting platform subscription handoff script');
