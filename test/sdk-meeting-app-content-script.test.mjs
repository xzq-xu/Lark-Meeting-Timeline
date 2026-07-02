import assert from 'node:assert/strict';

import {
  createMeetingAppContentScriptBridge,
  installMeetingAppContentScriptBridge,
} from '../packages/meeting-timeline-sdk/adapters/meeting-app-content-script.mjs';

function fakeExtensionRuntime() {
  const listeners = [];
  return {
    onMessage: {
      addListener(listener) {
        listeners.push(listener);
      },
      removeListener(listener) {
        const index = listeners.indexOf(listener);
        if (index >= 0) listeners.splice(index, 1);
      },
    },
    listenerCount() {
      return listeners.length;
    },
    emit(message, sender = { tab: { id: 1 } }) {
      if (listeners.length === 0) return Promise.resolve({ consumed: false });
      return new Promise((resolve) => {
        const consumed = listeners[0](message, sender, (response) => resolve({ consumed, response }));
        if (consumed !== true) resolve({ consumed, response: undefined });
      });
    },
  };
}

function fakeWindow() {
  const listeners = new Map();
  const posted = [];
  const win = {
    location: { href: 'https://meet.google.com/abc-defg-hij', origin: 'https://meet.google.com' },
    addEventListener(name, listener) {
      const rows = listeners.get(name) ?? [];
      rows.push(listener);
      listeners.set(name, rows);
    },
    removeEventListener(name, listener) {
      listeners.set(name, (listeners.get(name) ?? []).filter((item) => item !== listener));
    },
    postMessage(payload, origin) {
      posted.push({ payload, origin });
    },
    dispatchMessage(data, origin = 'https://meet.google.com') {
      for (const listener of listeners.get('message') ?? []) {
        listener({ data, origin, source: win });
      }
    },
    listenerCount(name) {
      return (listeners.get(name) ?? []).length;
    },
    posted,
  };
  return win;
}

function wait(ms = 0) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function fakeRuntime() {
  const calls = [];
  return {
    calls,
    async handleMessage(message, options) {
      calls.push({ method: 'handleMessage', message, options });
      return { handled: true, action: 'test', type: message.type };
    },
    start(options) {
      calls.push({ method: 'start', options });
      return { running: true };
    },
    stop() {
      calls.push({ method: 'stop' });
      return { running: false };
    },
    dispose() {
      calls.push({ method: 'dispose' });
      return { running: false };
    },
    getState() {
      return { monitor: { running: true } };
    },
  };
}

const extensionRuntime = fakeExtensionRuntime();
const runtime = fakeRuntime();
const bridge = createMeetingAppContentScriptBridge({}, {
  runtime,
  extensionRuntime,
});

const startState = bridge.start({ startRuntime: true });
assert.equal(startState.installs[0].installed, true);
assert.equal(extensionRuntime.listenerCount(), 1);
assert.equal(runtime.calls.at(-1).method, 'start');
assert.equal(bridge.getState().messaging.extension_installed, true);

const unsupported = await extensionRuntime.emit({ type: 'unrelated.message' });
assert.equal(unsupported.consumed, false);
assert.equal(runtime.calls.filter((item) => item.method === 'handleMessage').length, 0);

const extensionResponse = await extensionRuntime.emit({
  type: 'meeting_timeline.insert_mark',
  payload: { mark: { id: 'mark-1', label: 'why?' } },
});
assert.equal(extensionResponse.consumed, true);
assert.equal(extensionResponse.response.handled, true);
assert.equal(extensionResponse.response.action, 'test');
assert.equal(runtime.calls.at(-1).options.sender.tab.id, 1);

const directResponse = await bridge.dispatchMessage({ type: 'insert_mark', payload: { id: 'direct' } });
assert.equal(directResponse.handled, true);
assert.equal(runtime.calls.at(-1).message.type, 'insert_mark');

const removedExtension = bridge.removeMessaging('extension');
assert.equal(removedExtension.removed, 1);
assert.equal(extensionRuntime.listenerCount(), 0);
assert.equal(bridge.getState().messaging.extension_installed, false);

const win = fakeWindow();
const windowRuntime = fakeRuntime();
const windowBridge = createMeetingAppContentScriptBridge({}, {
  runtime: windowRuntime,
  window: win,
  windowMessaging: true,
  extensionMessaging: false,
});
windowBridge.start({ startRuntime: false });
assert.equal(win.listenerCount('message'), 1);
win.dispatchMessage({
  id: 'request-1',
  type: 'meeting_timeline.sample',
});
await wait();
assert.equal(windowRuntime.calls.filter((item) => item.method === 'handleMessage').length, 1);
assert.equal(win.posted.length, 1);
assert.equal(win.posted[0].payload.type, 'meeting_timeline.response');
assert.equal(win.posted[0].payload.request_id, 'request-1');
assert.equal(win.posted[0].origin, 'https://meet.google.com');

win.dispatchMessage({
  type: 'meeting_timeline.response',
  result: { handled: true },
});
await wait();
assert.equal(windowRuntime.calls.filter((item) => item.method === 'handleMessage').length, 1);

win.dispatchMessage({
  type: 'meeting_timeline.sample',
}, 'https://evil.example');
await wait();
assert.equal(windowRuntime.calls.filter((item) => item.method === 'handleMessage').length, 1);

const disposedWindow = windowBridge.dispose();
assert.equal(disposedWindow.messaging.removed, 1);
assert.equal(win.listenerCount('message'), 0);
assert.equal(windowRuntime.calls.at(-1).method, 'dispose');

const installedRuntime = fakeRuntime();
const installedExtension = fakeExtensionRuntime();
const installed = installMeetingAppContentScriptBridge({}, {
  runtime: installedRuntime,
  extensionRuntime: installedExtension,
});
assert.equal(installedExtension.listenerCount(), 1);
assert.equal(installedRuntime.calls.some((item) => item.method === 'start'), true);
installed.dispose();
assert.equal(installedExtension.listenerCount(), 0);

console.log('ok meeting app content script bridge');
