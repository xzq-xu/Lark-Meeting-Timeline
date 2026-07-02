import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import { buildArtifactImportPlan } from './artifact-plan.mjs';
import { importPlatformTranscript } from './transcript.mjs';

const DEFAULT_ENDPOINTS = Object.freeze({
  google_meet: 'https://meet.googleapis.com',
  microsoft_teams: 'https://graph.microsoft.com/v1.0',
});

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function bearerHeaders(token) {
  return token ? { authorization: `Bearer ${token}` } : {};
}

function stripLeadingSlash(value = '') {
  return String(value).replace(/^\/+/, '');
}

function endpointFor(platform, options = {}) {
  return String(
    firstNonEmpty(
      options.endpoint,
      options.baseUrl,
      options.base_url,
      options.endpoints?.[platform],
      DEFAULT_ENDPOINTS[platform],
      '',
    ),
  ).replace(/\/+$/, '');
}

function tokenFor(platform, options = {}) {
  return firstNonEmpty(
    options.token,
    options.accessToken,
    options.access_token,
    options.oauthToken,
    options.oauth_token,
    options.tokens?.[platform],
    options[`${platform}Token`],
    options[`${platform}_token`],
  );
}

function textAcceptFor(plan = {}, options = {}) {
  return firstNonEmpty(
    options.accept,
    options.contentType,
    options.content_type,
    plan.platform === 'microsoft_teams' ? 'text/vtt' : undefined,
  );
}

function asPlan(input = {}, options = {}) {
  if (input.fetch || input.transcript_import || input.action) return input;
  return buildArtifactImportPlan(input, options);
}

function withQuery(url, query = {}) {
  const parsed = new URL(url);
  for (const [key, value] of Object.entries(query)) {
    if (value != null && value !== '') parsed.searchParams.set(key, String(value));
  }
  return parsed.toString();
}

function googleRequest(plan = {}, options = {}) {
  const resource = stripLeadingSlash(plan.fetch?.resource);
  if (!resource) {
    throw new MeetingTimelineSdkError('Google Meet artifact fetch requires fetch.resource');
  }
  return {
    method: 'GET',
    url: withQuery(`${endpointFor('google_meet', options)}/v2/${resource}`, {
      pageSize: options.pageSize ?? options.page_size,
      pageToken: options.pageToken ?? options.page_token,
    }),
    headers: {
      ...bearerHeaders(tokenFor('google_meet', options)),
      ...(options.headers ?? {}),
    },
    response_type: 'json',
  };
}

function teamsTranscriptResource(plan = {}, options = {}) {
  const meetingId = firstNonEmpty(plan.meeting?.meeting_id, options.meetingId, options.meeting_id);
  const transcriptId = firstNonEmpty(plan.artifact_id, options.transcriptId, options.transcript_id, options.artifactId, options.artifact_id);
  const userId = firstNonEmpty(options.userId, options.user_id);
  if (options.resource) return stripLeadingSlash(options.resource);
  if (!meetingId || !transcriptId) {
    throw new MeetingTimelineSdkError('Microsoft Teams transcript fetch requires meeting_id and transcript artifact_id');
  }
  const owner = userId ? `users/${encodeURIComponent(userId)}` : 'me';
  return `${owner}/onlineMeetings/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}/content`;
}

function microsoftTeamsRequest(plan = {}, options = {}) {
  return {
    method: 'GET',
    url: `${endpointFor('microsoft_teams', options)}/${teamsTranscriptResource(plan, options)}`,
    headers: {
      ...bearerHeaders(tokenFor('microsoft_teams', options)),
      accept: textAcceptFor(plan, options),
      ...(options.headers ?? {}),
    },
    response_type: 'text',
  };
}

function directDownloadRequest(platform, plan = {}, options = {}) {
  const url = firstNonEmpty(plan.fetch?.url, plan.artifact_url, options.url);
  if (!url) {
    throw new MeetingTimelineSdkError(`${platform} artifact fetch requires artifact_url or fetch.url`);
  }
  const token = tokenFor(platform, options);
  return {
    method: 'GET',
    url,
    headers: {
      ...bearerHeaders(token),
      ...(options.headers ?? {}),
    },
    response_type: plan.artifact_kind === 'recording' ? 'arrayBuffer' : 'text',
  };
}

export function buildArtifactFetchRequest(input = {}, options = {}) {
  const plan = asPlan(input, options);
  if (plan.status === 'ignored' || plan.status === 'unsupported_artifact') {
    throw new MeetingTimelineSdkError(`Artifact fetch is not available for plan status: ${plan.status}`);
  }
  if (plan.platform === 'google_meet' && plan.artifact_kind === 'transcript') return googleRequest(plan, options);
  if (plan.platform === 'microsoft_teams' && plan.artifact_kind === 'transcript') return microsoftTeamsRequest(plan, options);
  if (plan.platform === 'zoom' || plan.platform === 'webex') return directDownloadRequest(plan.platform, plan, options);
  return directDownloadRequest(plan.platform ?? 'provider', plan, options);
}

async function readResponseBody(response, responseType) {
  if (responseType === 'json') return response.json();
  if (responseType === 'arrayBuffer') return response.arrayBuffer();
  return response.text();
}

function responseHeaders(response) {
  if (!response?.headers) return {};
  if (typeof response.headers.entries === 'function') return Object.fromEntries(response.headers.entries());
  return response.headers;
}

export async function fetchArtifactContent(input = {}, options = {}) {
  const plan = asPlan(input, options);
  const request = buildArtifactFetchRequest(plan, options);
  const fetchImpl = options.fetchImpl ?? options.fetch ?? globalThis.fetch;
  if (typeof fetchImpl !== 'function') {
    throw new MeetingTimelineSdkError('fetchArtifactContent requires fetchImpl or global fetch');
  }
  const response = await fetchImpl(request.url, {
    method: request.method,
    headers: request.headers,
  });
  const body = await readResponseBody(response, request.response_type);
  if (!response.ok) {
    throw new MeetingTimelineSdkError(`Artifact fetch failed with HTTP ${response.status}`, {
      status: response.status,
      body,
      request,
    });
  }
  return compactObject({
    platform: plan.platform,
    artifact_kind: plan.artifact_kind,
    request,
    status: response.status,
    headers: responseHeaders(response),
    content_type: response.headers?.get?.('content-type'),
    body,
    plan,
  });
}

export async function fetchAndImportArtifactTranscript(client, input = {}, options = {}) {
  const fetched = await fetchArtifactContent(input, options);
  const raw = firstNonEmpty(options.raw, fetched.body);
  const response = await importPlatformTranscript(client, {
    platform: fetched.platform,
    meeting: fetched.plan.meeting,
    raw,
    source: options.source,
    language: options.language,
  }, options.importOptions ?? options.import_options ?? {});
  return {
    fetched,
    import_result: response,
  };
}
