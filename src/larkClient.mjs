const DEFAULT_BASE_URL = 'https://open.feishu.cn';
const DEFAULT_AUTHORIZE_PATH = '/open-apis/authen/v1/index';

export function createLarkClient(env = process.env) {
  const baseUrl = (env.LARK_BASE_URL || DEFAULT_BASE_URL).replace(/\/$/, '');
  const appId = env.LARK_APP_ID;
  const appSecret = env.LARK_APP_SECRET;
  const authorizePath = env.LARK_OAUTH_AUTHORIZE_PATH || DEFAULT_AUTHORIZE_PATH;
  const redirectUri = env.LARK_REDIRECT_URI || 'http://localhost:8787/api/auth/lark/callback';
  const oauthScopes = env.LARK_OAUTH_SCOPES || '';
  let tenantTokenCache = null;
  let appTokenCache = null;

  async function request(path, options = {}) {
    const url = `${baseUrl}${path}`;
    const response = await fetch(url, {
      ...options,
      headers: {
        'content-type': 'application/json; charset=utf-8',
        ...(options.headers ?? {}),
      },
    });
    const text = await response.text();
    let json = null;
    try {
      json = text ? JSON.parse(text) : null;
    } catch {
      json = { raw: text };
    }
    if (!response.ok) {
      const message = json?.msg || json?.message || text || `HTTP ${response.status}`;
      throw new Error(`Lark request failed: ${message}`);
    }
    return json;
  }

  async function tenantAccessToken() {
    if (!appId || !appSecret) {
      throw new Error('Missing LARK_APP_ID or LARK_APP_SECRET');
    }
    const now = Date.now();
    if (tenantTokenCache && tenantTokenCache.expires_at_ms - now > 60_000) return tenantTokenCache.token;
    const json = await request('/open-apis/auth/v3/tenant_access_token/internal', {
      method: 'POST',
      body: JSON.stringify({ app_id: appId, app_secret: appSecret }),
    });
    if (json.code !== 0) throw new Error(`Lark token error: ${json.msg || json.message || json.code}`);
    const token = json.tenant_access_token;
    tenantTokenCache = {
      token,
      expires_at_ms: now + Math.max(60, Number(json.expire ?? 7200) - 120) * 1000,
    };
    return token;
  }

  async function appAccessToken() {
    if (!appId || !appSecret) {
      throw new Error('Missing LARK_APP_ID or LARK_APP_SECRET');
    }
    const now = Date.now();
    if (appTokenCache && appTokenCache.expires_at_ms - now > 60_000) return appTokenCache.token;
    const json = await request('/open-apis/auth/v3/app_access_token/internal', {
      method: 'POST',
      body: JSON.stringify({ app_id: appId, app_secret: appSecret }),
    });
    if (json.code !== 0) throw new Error(`Lark app token error: ${json.msg || json.message || json.code}`);
    const token = json.app_access_token;
    appTokenCache = {
      token,
      expires_at_ms: now + Math.max(60, Number(json.expire ?? 7200) - 120) * 1000,
    };
    return token;
  }

  async function authedGet(path, query = {}) {
    const token = await tenantAccessToken();
    return getWithToken(path, token, query);
  }

  async function getWithToken(path, token, query = {}) {
    const qs = new URLSearchParams();
    for (const [key, value] of Object.entries(query)) {
      if (value != null && value !== '') qs.set(key, String(value));
    }
    const suffix = qs.size ? `?${qs.toString()}` : '';
    return request(`${path}${suffix}`, {
      method: 'GET',
      headers: { authorization: `Bearer ${token}` },
    });
  }

  function scopeList(value) {
    return String(value ?? '')
      .split(/[\s,]+/)
      .map((scope) => scope.trim())
      .filter(Boolean);
  }

  function createAuthorizeUrl(state, opts = {}) {
    if (!appId) throw new Error('Missing LARK_APP_ID');
    const params = new URLSearchParams({
      app_id: appId,
      redirect_uri: redirectUri,
      state,
    });
    const scopes = [
      ...(opts.ignoreDefaultScopes ? [] : scopeList(oauthScopes)),
      ...scopeList(opts.scope),
      ...scopeList(opts.scopes),
      ...(Array.isArray(opts.extraScopes) ? opts.extraScopes.flatMap(scopeList) : []),
    ];
    const dedupedScopes = [...new Set(scopes)];
    if (dedupedScopes.length) params.set('scope', dedupedScopes.join(' '));
    return `${baseUrl}${authorizePath}?${params.toString()}`;
  }

  function normalizeOAuthToken(json) {
    const data = json?.data ?? json;
    const accessToken = data?.access_token ?? data?.user_access_token;
    if (!accessToken) throw new Error(`Lark OAuth token response missing access_token: ${JSON.stringify(json)}`);
    return {
      access_token: accessToken,
      refresh_token: data?.refresh_token ?? null,
      expires_in: Number(data?.expires_in ?? data?.expire ?? 7200),
      refresh_expires_in: Number(data?.refresh_expires_in ?? data?.refresh_expire ?? 0),
      token_type: data?.token_type ?? 'Bearer',
      scope: data?.scope ?? null,
      raw: json,
    };
  }

  async function exchangeOAuthCodeV2(code) {
    const json = await request('/open-apis/authen/v2/oauth/token', {
      method: 'POST',
      body: JSON.stringify({
        grant_type: 'authorization_code',
        client_id: appId,
        client_secret: appSecret,
        code,
        redirect_uri: redirectUri,
      }),
    });
    if (json.code != null && json.code !== 0) {
      throw new Error(`Lark OAuth v2 token error: ${json.msg || json.message || json.code}`);
    }
    return normalizeOAuthToken(json);
  }

  async function exchangeOAuthCodeLegacy(code) {
    const token = await appAccessToken();
    const json = await request('/open-apis/authen/v1/access_token', {
      method: 'POST',
      headers: { authorization: `Bearer ${token}` },
      body: JSON.stringify({
        grant_type: 'authorization_code',
        code,
      }),
    });
    if (json.code != null && json.code !== 0) {
      throw new Error(`Lark OAuth legacy token error: ${json.msg || json.message || json.code}`);
    }
    return normalizeOAuthToken(json);
  }

  async function exchangeOAuthCode(code) {
    if (!appId || !appSecret) {
      throw new Error('Missing LARK_APP_ID or LARK_APP_SECRET');
    }
    try {
      return await exchangeOAuthCodeV2(code);
    } catch (v2Error) {
      try {
        const token = await exchangeOAuthCodeLegacy(code);
        token.fallback = 'legacy';
        return token;
      } catch (legacyError) {
        throw new Error(`${v2Error.message}; fallback failed: ${legacyError.message}`);
      }
    }
  }

  async function refreshOAuthToken(refreshToken) {
    if (!appId || !appSecret) {
      throw new Error('Missing LARK_APP_ID or LARK_APP_SECRET');
    }
    if (!refreshToken) throw new Error('refresh_token is required');
    const json = await request('/open-apis/authen/v2/oauth/token', {
      method: 'POST',
      body: JSON.stringify({
        grant_type: 'refresh_token',
        client_id: appId,
        client_secret: appSecret,
        refresh_token: refreshToken,
      }),
    });
    if (json.code != null && json.code !== 0) {
      throw new Error(`Lark OAuth refresh error: ${json.msg || json.message || json.code}`);
    }
    return normalizeOAuthToken(json);
  }

  async function fetchUserInfo(userAccessToken) {
    return getWithToken('/open-apis/authen/v1/user_info', userAccessToken);
  }

  async function fetchMinuteTranscript(minuteToken) {
    if (!minuteToken) throw new Error('minute_token is required');
    const encoded = encodeURIComponent(minuteToken);
    return authedGet(`/open-apis/minutes/v1/minutes/${encoded}/transcript`, {
      need_speaker: true,
      need_timestamp: true,
    });
  }

  async function fetchMinuteTranscriptWithUserToken(minuteToken, userAccessToken) {
    if (!minuteToken) throw new Error('minute_token is required');
    if (!userAccessToken) throw new Error('user_access_token is required');
    const encoded = encodeURIComponent(minuteToken);
    return getWithToken(`/open-apis/minutes/v1/minutes/${encoded}/transcript`, userAccessToken, {
      need_speaker: true,
      need_timestamp: true,
    });
  }

  async function searchMinutesWithUserToken(userAccessToken, opts = {}) {
    if (!userAccessToken) throw new Error('user_access_token is required');
    const body = {
      page_size: Math.min(Math.max(Number(opts.page_size ?? 20), 1), 200),
    };
    if (opts.page_token) body.page_token = opts.page_token;
    if (opts.query) body.query = String(opts.query);
    if (opts.keyword) body.keyword = String(opts.keyword);
    if (opts.start_time || opts.end_time) {
      body.time_range = {};
      if (opts.start_time) body.time_range.start_time = opts.start_time;
      if (opts.end_time) body.time_range.end_time = opts.end_time;
    }
    return request('/open-apis/minutes/v1/minutes/search', {
      method: 'POST',
      headers: { authorization: `Bearer ${userAccessToken}` },
      body: JSON.stringify(body),
    });
  }

  async function fetchMinuteMedia(minuteToken) {
    if (!minuteToken) throw new Error('minute_token is required');
    const encoded = encodeURIComponent(minuteToken);
    return authedGet(`/open-apis/minutes/v1/minutes/${encoded}/media`);
  }

  async function createMeetingReserve(opts = {}) {
    const token = await tenantAccessToken();
    const now = Date.now();
    const endTimeSeconds = Math.floor((opts.end_time_ms ?? now + 2 * 60 * 60 * 1000) / 1000);
    const body = {
      end_time: String(opts.end_time ?? endTimeSeconds),
      meeting_settings: {
        topic: opts.topic ?? opts.title ?? '实时标注会议',
        meeting_initial_type: opts.meeting_initial_type ?? 1,
        meeting_connect: opts.meeting_connect ?? true,
        auto_record: opts.auto_record ?? false,
      },
    };
    if (opts.password) body.meeting_settings.password = String(opts.password);
    if (opts.owner_id) body.owner_id = String(opts.owner_id);
    return request('/open-apis/vc/v1/reserves/apply', {
      method: 'POST',
      headers: { authorization: `Bearer ${token}` },
      body: JSON.stringify(body),
    });
  }

  async function fetchReserveActiveMeeting(reserveId) {
    if (!reserveId) throw new Error('reserve_id is required');
    const token = await tenantAccessToken();
    const encoded = encodeURIComponent(reserveId);
    return getWithToken(`/open-apis/vc/v1/reserves/${encoded}/get_active_meeting`, token, {
      with_participants: true,
    });
  }

  async function listMeetingsByNoWithToken(meetingNo, token, opts = {}) {
    if (!meetingNo) throw new Error('meeting_no is required');
    if (!token) throw new Error('access token is required');
    return getWithToken('/open-apis/vc/v1/meetings/list_by_no', token, {
      meeting_no: meetingNo,
      start_time: opts.start_time,
      end_time: opts.end_time,
      page_size: opts.page_size ?? 10,
      page_token: opts.page_token,
    });
  }

  async function listMeetingsByNo(meetingNo, opts = {}) {
    const token = await tenantAccessToken();
    return listMeetingsByNoWithToken(meetingNo, token, opts);
  }

  async function searchMeetingsWithToken(token, opts = {}) {
    if (!token) throw new Error('access token is required');
    const body = {};
    if (opts.query) body.query = String(opts.query);
    const meetingFilter = {};
    if (opts.organizer_ids?.length) meetingFilter.organizer_ids = opts.organizer_ids.map(String);
    if (opts.participant_ids?.length) meetingFilter.participant_ids = opts.participant_ids.map(String);
    if (opts.open_room_ids?.length) meetingFilter.open_room_ids = opts.open_room_ids.map(String);
    if (opts.start_time || opts.end_time) {
      meetingFilter.start_time = {};
      if (opts.start_time) meetingFilter.start_time.start_time = String(opts.start_time);
      if (opts.end_time) meetingFilter.start_time.end_time = String(opts.end_time);
    }
    if (Object.keys(meetingFilter).length) body.meeting_filter = meetingFilter;
    const qs = new URLSearchParams();
    qs.set('page_size', String(Math.min(Math.max(Number(opts.page_size ?? 10), 1), 50)));
    if (opts.page_token) qs.set('page_token', String(opts.page_token));
    return request(`/open-apis/vc/v1/meetings/search?${qs.toString()}`, {
      method: 'POST',
      headers: { authorization: `Bearer ${token}` },
      body: JSON.stringify(body),
    });
  }

  async function searchMeetings(opts = {}) {
    const token = await tenantAccessToken();
    return searchMeetingsWithToken(token, opts);
  }

  async function fetchMeetingDetailWithToken(meetingId, token) {
    if (!meetingId) throw new Error('meeting_id is required');
    if (!token) throw new Error('access token is required');
    return getWithToken(`/open-apis/vc/v1/meetings/${encodeURIComponent(meetingId)}`, token);
  }

  async function fetchMeetingDetail(meetingId) {
    const token = await tenantAccessToken();
    return fetchMeetingDetailWithToken(meetingId, token);
  }

  return {
    baseUrl,
    redirectUri,
    isConfigured: Boolean(appId && appSecret),
    tenantAccessToken,
    appAccessToken,
    createAuthorizeUrl,
    exchangeOAuthCode,
    refreshOAuthToken,
    fetchUserInfo,
    searchMinutesWithUserToken,
    fetchMinuteTranscript,
    fetchMinuteTranscriptWithUserToken,
    fetchMinuteMedia,
    createMeetingReserve,
    fetchReserveActiveMeeting,
    listMeetingsByNo,
    listMeetingsByNoWithToken,
    searchMeetings,
    searchMeetingsWithToken,
    fetchMeetingDetail,
    fetchMeetingDetailWithToken,
  };
}

export function extractMinuteToken(input) {
  const text = String(input || '').trim();
  if (!text) return '';
  try {
    const url = new URL(text);
    const pathParts = url.pathname.split('/').filter(Boolean);
    const known = ['minutes', 'minute', 'minutes_token', 'minutesToken'];
    for (let i = 0; i < pathParts.length; i++) {
      if (known.includes(pathParts[i]) && pathParts[i + 1]) return pathParts[i + 1];
    }
    for (const key of ['minute_token', 'token', 'object_token']) {
      const value = url.searchParams.get(key);
      if (value) return value;
    }
    return pathParts.at(-1) || text;
  } catch {
    return text;
  }
}

export function extractMeetingNo(input) {
  const text = String(input || '').trim();
  if (!text) return '';
  try {
    const url = new URL(text);
    const pathParts = url.pathname.split('/').filter(Boolean);
    const joinIndex = pathParts.findIndex((part) => ['j', 'join', 'meeting'].includes(part));
    if (joinIndex >= 0 && pathParts[joinIndex + 1]) return pathParts[joinIndex + 1].replace(/\D/g, '') || pathParts[joinIndex + 1];
    for (const key of ['meeting_no', 'meetingNo', 'no']) {
      const value = url.searchParams.get(key);
      if (value) return value.replace(/\D/g, '') || value;
    }
    const last = pathParts.at(-1);
    if (last) return last.replace(/\D/g, '') || last;
  } catch {
    // fall through to plain-text extraction
  }
  const digits = text.match(/\d{6,}/g);
  return digits ? digits.join('') : text;
}
