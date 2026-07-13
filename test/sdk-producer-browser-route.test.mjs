import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import net from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

async function freePort() {
  const server = net.createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = server.address().port;
  server.close();
  await once(server, 'close');
  return port;
}

async function waitForServer(baseUrl) {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/sdk/producer.mjs`);
      if (response.ok) return;
    } catch {
      // Retry while the isolated server starts.
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error('server did not expose SDK producer module');
}

const port = await freePort();
const dataDir = await mkdtemp(join(tmpdir(), 'sdk-browser-route-'));
const baseUrl = `http://127.0.0.1:${port}`;
const child = spawn(process.execPath, ['src/server.mjs'], {
  cwd: new URL('..', import.meta.url),
  env: {
    ...process.env,
    PORT: String(port),
    TIMELINE_DATA_DIR: dataDir,
    LARK_WS_EVENTS: '0',
    REAL_DEMO_AUTO_ARM: '0',
  },
  stdio: ['ignore', 'pipe', 'pipe'],
});

try {
  await waitForServer(baseUrl);
  const producerResponse = await fetch(`${baseUrl}/sdk/producer.mjs`);
  const producerSource = await producerResponse.text();
  assert.equal(producerResponse.headers.get('content-type'), 'text/javascript; charset=utf-8');
  assert.match(producerSource, /createMeetingTimelineAnnotationProducer/);
  assert.match(producerSource, /from '.\/errors\.mjs'/);
  assert.doesNotMatch(producerSource, /node:/);

  const errorsResponse = await fetch(`${baseUrl}/sdk/errors.mjs`);
  assert.equal(errorsResponse.ok, true);
  assert.match(await errorsResponse.text(), /class MeetingTimelineSdkError/);

  const appSource = await (await fetch(`${baseUrl}/app.js`)).text();
  assert.match(appSource, /from '\/sdk\/producer\.mjs'/);
  assert.match(appSource, /standardAnnotationProducer\.publish\(body\)/);
} finally {
  child.kill('SIGTERM');
  await once(child, 'exit').catch(() => {});
  await rm(dataDir, { recursive: true, force: true });
}

console.log('ok sdk producer browser route');
