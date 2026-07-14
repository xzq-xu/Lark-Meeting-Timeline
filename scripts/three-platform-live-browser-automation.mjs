import { normalizeMeetingPlatform } from '../packages/meeting-timeline-sdk/adapters/platform-setup.mjs';

const ERROR_HINTS = Object.freeze([
  'error',
  'something went wrong',
  'retry',
  'clear cache',
  '错误',
  '出错了',
]);

const LABELS = Object.freeze({
  continue_browser: [
    'continue on this browser',
    'join on the web instead',
    '在此浏览器上继续',
    '改为在 web 上加入',
  ],
  join_from_browser: [
    'join from your browser',
    'join from browser',
    '从浏览器加入',
    '通过浏览器加入',
  ],
  join: [
    'join now',
    'ask to join',
    'join meeting',
    'join',
    '立即加入',
    '申请加入',
    '加入会议',
    '加入',
  ],
  join_audio: [
    'join audio by computer',
    'join with computer audio',
    'computer audio',
    'join audio',
    '使用电脑音频加入',
    '通过计算机加入音频',
    '加入音频',
  ],
  leave: [
    'leave call',
    'leave meeting',
    'hang up',
    'leave',
    '离开通话',
    '退出通话',
    '离开会议',
    '挂断',
  ],
  name: [
    'your name',
    'enter your name',
    'display name',
    'name',
    '输入姓名',
    '您的姓名',
    '显示名称',
    '姓名',
  ],
});

function text(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function visibleControls(page = {}) {
  return (page.controls ?? []).filter((item) => item?.visible !== false && item?.disabled !== true);
}

function controlCorpus(control = {}) {
  return [control.text, control.aria, control.title, control.placeholder]
    .map(text)
    .filter(Boolean);
}

function exactLabel(control, labels) {
  const corpus = controlCorpus(control);
  return labels.some((label) => corpus.includes(text(label)));
}

function includesLabel(control, labels) {
  const corpus = controlCorpus(control);
  return labels.some((label) => corpus.some((value) => value.includes(text(label))));
}

function pageError(page = {}) {
  const corpus = text(`${page.title ?? ''} ${visibleControls(page).flatMap(controlCorpus).join(' ')}`);
  return ERROR_HINTS.some((hint) => corpus.includes(text(hint)));
}

function pageMatchesPlatform(page = {}, platform) {
  let host = '';
  try { host = new URL(String(page.url ?? '')).hostname.toLowerCase(); } catch {}
  if (platform === 'google_meet') return host === 'meet.google.com';
  if (platform === 'microsoft_teams') return host === 'teams.microsoft.com'
    || host.endsWith('.teams.microsoft.com')
    || host === 'teams.live.com'
    || host.endsWith('.teams.live.com')
    || host === 'teams.cloud.microsoft'
    || host.endsWith('.teams.cloud.microsoft');
  if (platform === 'zoom') return host === 'zoom.us' || host.endsWith('.zoom.us') || host === 'zoom.com' || host.endsWith('.zoom.com');
  return false;
}

function meetingJoinSurface(page = {}, platform) {
  let path = '';
  try { path = new URL(String(page.url ?? '')).pathname.toLowerCase(); } catch {}
  if (platform === 'google_meet') return /^\/[a-z]{3}-[a-z]{4}-[a-z]{3}\/?$/.test(path);
  if (platform === 'microsoft_teams') return path.includes('meetup-join') || path.includes('/meet/') || path.includes('/pre-join/');
  if (platform === 'zoom') return path.includes('/wc/') || path.includes('/j/') || path.includes('/join/');
  return false;
}

function controlSelector(control = {}) {
  if (control.id) return { id: String(control.id) };
  if (control.aria) return { aria: String(control.aria) };
  if (control.placeholder) return { placeholder: String(control.placeholder) };
  if (control.text) return { text: String(control.text) };
  return null;
}

function action(page, phase, type, control, value) {
  const selector = controlSelector(control);
  if (!selector) return null;
  return {
    key: `${page.target_id}:${phase}`,
    target_id: page.target_id,
    page_url: page.url,
    phase,
    type,
    selector,
    value,
  };
}

function firstControl(page, predicate) {
  return visibleControls(page).find(predicate) ?? null;
}

function nextJoinAction(page, platform, completed) {
  if (platform === 'zoom') {
    let path = '';
    try { path = new URL(String(page.url ?? '')).pathname.toLowerCase(); } catch {}
    if (path === '/test' || path.endsWith('/test')) {
      const start = firstControl(page, (item) => ['scheduleMtg', 'btnJoinTest'].includes(item.id)
        || exactLabel(item, ['start a new meeting', 'join a test meeting', '开始新会议', '加入测试会议']));
      if (start && !completed.has(`${page.target_id}:zoom_test_start`)) return action(page, 'zoom_test_start', 'click', start);
      return null;
    }
  }

  const continueBrowser = firstControl(page, (item) => exactLabel(item, LABELS.continue_browser));
  if (continueBrowser && !completed.has(`${page.target_id}:continue_browser`)) {
    return action(page, 'continue_browser', 'click', continueBrowser);
  }
  const joinFromBrowser = firstControl(page, (item) => exactLabel(item, LABELS.join_from_browser));
  if (joinFromBrowser && !completed.has(`${page.target_id}:join_from_browser`)) {
    return action(page, 'join_from_browser', 'click', joinFromBrowser);
  }
  if (!meetingJoinSurface(page, platform)) return null;

  const nameInput = firstControl(page, (item) => (
    item.tag === 'input'
    && ['text', '', undefined].includes(item.type)
    && item.value_present !== true
    && !text(item.value)
    && includesLabel(item, LABELS.name)
  ));
  if (nameInput && !completed.has(`${page.target_id}:fill_display_name`)) {
    return action(page, 'fill_display_name', 'fill', nameInput, 'Timeline Adapter Acceptance');
  }

  const join = firstControl(page, (item) => (
    ['button', 'a'].includes(item.tag)
    && exactLabel(item, LABELS.join)
  ));
  if (join && !completed.has(`${page.target_id}:join_meeting`)) return action(page, 'join_meeting', 'click', join);
  return null;
}

export function chooseThreePlatformBrowserAutomationAction(input = {}) {
  const platform = normalizeMeetingPlatform(input.platform);
  const completed = new Set(input.completedActionKeys ?? []);
  const pages = (input.pages ?? [])
    .filter((page) => pageMatchesPlatform(page, platform))
    .filter((page) => !pageError(page));

  if (input.activeMeeting === true) {
    for (const page of pages) {
      const joinAudio = firstControl(page, (item) => exactLabel(item, LABELS.join_audio));
      if (input.autoJoin === true && joinAudio && !completed.has(`${page.target_id}:join_audio`)) {
        return action(page, 'join_audio', 'click', joinAudio);
      }
    }
    if (input.autoLeave === true && input.speakerSeen === true && input.leaveReady === true) {
      for (const page of pages) {
        const initialLeaveKey = `${page.target_id}:leave_meeting`;
        if (completed.has(initialLeaveKey)) {
          const confirmation = firstControl(page, (item) => item.in_dialog === true && exactLabel(item, LABELS.leave));
          if (confirmation && !completed.has(`${page.target_id}:confirm_leave_meeting`)) {
            return action(page, 'confirm_leave_meeting', 'click', confirmation);
          }
          continue;
        }
        const leave = firstControl(page, (item) => exactLabel(item, LABELS.leave));
        if (leave && !completed.has(`${page.target_id}:leave_meeting`)) {
          return action(page, 'leave_meeting', 'click', leave);
        }
      }
    }
    return null;
  }

  if (input.autoJoin !== true) return null;
  for (const page of pages) {
    const candidate = nextJoinAction(page, platform, completed);
    if (candidate) return candidate;
  }
  return null;
}

export function buildThreePlatformBrowserActionExpression(action) {
  const serialized = JSON.stringify(action);
  return `(() => {
    const action = ${serialized};
    const controls = Array.from(document.querySelectorAll('button, a, input'));
    const visible = (node) => {
      const rect = node.getBoundingClientRect();
      const style = getComputedStyle(node);
      return rect.width > 0 && rect.height > 0 && style.visibility !== 'hidden' && style.display !== 'none' && !node.disabled;
    };
    const match = (node) => {
      if (!visible(node)) return false;
      if (action.selector.id != null) return node.id === action.selector.id;
      if (action.selector.aria != null) return node.getAttribute('aria-label') === action.selector.aria;
      if (action.selector.placeholder != null) return node.getAttribute('placeholder') === action.selector.placeholder;
      if (action.selector.text != null) return (node.innerText || '').replace(/\\s+/g, ' ').trim() === action.selector.text;
      return false;
    };
    const node = controls.find(match);
    if (!node) return { ok: false, reason: 'control_not_found', phase: action.phase };
    if (action.type === 'fill') {
      const descriptor = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
      descriptor.set.call(node, action.value);
      node.dispatchEvent(new Event('input', { bubbles: true }));
      node.dispatchEvent(new Event('change', { bubbles: true }));
    } else {
      node.click();
    }
    return { ok: true, phase: action.phase, type: action.type, url: location.href };
  })()`;
}
