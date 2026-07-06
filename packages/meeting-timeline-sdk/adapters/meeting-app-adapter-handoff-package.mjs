import { MeetingTimelineSdkError, compactObject } from '../index.mjs';
import {
  buildMeetingAppAdapterManifest,
} from './meeting-app-adapter-manifest.mjs';
import {
  MEETING_APP_ADAPTER_SPEC_SCHEMA,
  buildMeetingAppAdapterSpec,
} from './meeting-app-adapter-spec.mjs';
import {
  MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA,
  buildMeetingAppAdapterRuntimeConfig,
} from './meeting-app-adapter-runtime-config.mjs';

export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA = 'meeting_app_adapter_handoff_package';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA = 'meeting_app_adapter_handoff_package_matrix';
export const MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION = 1;

function firstNonEmpty(...values) {
  return values.find((value) => value != null && value !== '');
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (value && typeof value !== 'string' && typeof value[Symbol.iterator] === 'function') return Array.from(value);
  return value == null ? [] : [value];
}

function unique(values = []) {
  return [...new Set(values.filter((value) => value != null && value !== '').map((value) => String(value)))];
}

function issue(severity, code, message, details = {}) {
  return compactObject({ severity, code, message, ...details });
}

function packageInputs(options = {}) {
  const specs = asArray(firstNonEmpty(options.packages, options.configs, options.adapters, options.adapter_specs, options.adapterSpecs, options.specs));
  const platforms = asArray(firstNonEmpty(options.platforms, options.platform_keys, options.platformKeys, []));
  return [...platforms, ...specs];
}

function cleanOptions(options = {}) {
  return {
    ...options,
    packages: undefined,
    configs: undefined,
    adapters: undefined,
    adapter_specs: undefined,
    adapterSpecs: undefined,
    specs: undefined,
    platforms: undefined,
    platform_keys: undefined,
    platformKeys: undefined,
  };
}

function maybeBuiltInManifest(spec = {}, options = {}) {
  if (spec.source !== 'built_in_manifest') return undefined;
  try {
    return buildMeetingAppAdapterManifest(spec.platform ?? spec.adapter_key, options);
  } catch {
    return undefined;
  }
}

function integrationReadme({ spec, runtimeConfig, manifest }) {
  return [
    `# ${spec.display_name || spec.adapter_key} Meeting Adapter Handoff`,
    '',
    'This package is the host-side contract for adding this meeting app to Meeting Timeline annotations.',
    '',
    '## Files',
    '- `adapter-spec.json`: static URL, permission, selector, MutationObserver and nonblocking contract.',
    '- `runtime-config.json`: options for browser extension, WebView preload or Electron content script runtime.',
    '- `extension-manifest-fragment.json`: match patterns, permissions and content script fragment.',
    manifest ? '- `adapter-manifest.json`: built-in adapter manifest used to derive this package.' : '',
    '',
    '## Minimal Integration',
    '1. Install `@ai-annotation/meeting-timeline-sdk` in the host project.',
    '2. Load `runtime-config.json` and pass `content_script_options` to `installMeetingAppContentScriptBridge()`.',
    '3. Pass `browser_runtime_options` to `createMeetingAppBrowserRuntime()` when observing inside a page context.',
    '4. Pass `capture_options` to `captureMeetingAppDomSnapshot()` when taking manual live snapshots.',
    '5. Keep provider events and transcript import nonblocking for realtime annotation placement.',
    '',
    '## Validation',
    '- Capture live DOM snapshots before production rollout.',
    '- Verify speaker and participant track extraction on real meetings.',
    '- Verify annotation insertions keep `captured_at_ms` and land on the current meeting axis.',
    '',
    '## Contract',
    `- adapter_key: ${spec.adapter_key}`,
    `- timestamp_field: ${runtimeConfig.contracts?.timestamp_field}`,
    `- provider_events_block_realtime: ${runtimeConfig.contracts?.provider_events_block_realtime}`,
    `- transcript_blocks_realtime: ${runtimeConfig.contracts?.transcript_blocks_realtime}`,
    '',
  ].filter(Boolean).join('\n');
}

function packageIssues({ spec, runtimeConfig, files }) {
  return [
    spec.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
      ? undefined
      : issue('error', 'missing_adapter_spec', 'Handoff package requires an adapter spec.'),
    spec.accepted === true
      ? undefined
      : issue('error', 'adapter_spec_not_accepted', 'Adapter spec must pass static acceptance.', {
        adapter_key: spec.adapter_key,
        spec_issue_count: spec.issue_count,
      }),
    runtimeConfig.schema === MEETING_APP_ADAPTER_RUNTIME_CONFIG_SCHEMA
      ? undefined
      : issue('error', 'missing_runtime_config', 'Handoff package requires a runtime config.'),
    runtimeConfig.accepted === true
      ? undefined
      : issue('error', 'runtime_config_not_accepted', 'Runtime config must pass static acceptance.', {
        adapter_key: runtimeConfig.adapter_key,
        runtime_issue_count: runtimeConfig.issue_count,
      }),
    runtimeConfig.readiness?.content_script_ready === true
      ? undefined
      : issue('error', 'content_script_not_ready', 'Handoff package requires content script match patterns.'),
    runtimeConfig.readiness?.capture_options_ready === true
      ? undefined
      : issue('error', 'capture_options_not_ready', 'Handoff package requires control and participant selectors.'),
    runtimeConfig.readiness?.mutation_observer_ready === true
      ? undefined
      : issue('error', 'mutation_observer_not_ready', 'Handoff package requires mutation observation selectors.'),
    files.some((file) => file.path === 'runtime-config.json')
      ? undefined
      : issue('error', 'missing_runtime_config_file', 'Handoff file list must include runtime-config.json.'),
  ].filter(Boolean);
}

export function buildMeetingAppAdapterHandoffPackage(specOrPlatform = {}, options = {}) {
  const packageOptions = cleanOptions(options);
  const spec = specOrPlatform?.schema === MEETING_APP_ADAPTER_SPEC_SCHEMA
    ? specOrPlatform
    : buildMeetingAppAdapterSpec(specOrPlatform, packageOptions);
  const runtimeConfig = buildMeetingAppAdapterRuntimeConfig(spec, packageOptions);
  const manifest = maybeBuiltInManifest(spec, packageOptions);
  const files = [
    {
      path: 'adapter-spec.json',
      media_type: 'application/json',
      schema: spec.schema,
      content: spec,
    },
    manifest
      ? {
        path: 'adapter-manifest.json',
        media_type: 'application/json',
        schema: manifest.schema,
        content: manifest,
      }
      : undefined,
    {
      path: 'runtime-config.json',
      media_type: 'application/json',
      schema: runtimeConfig.schema,
      content: runtimeConfig,
    },
    {
      path: 'extension-manifest-fragment.json',
      media_type: 'application/json',
      schema: 'browser_extension_manifest_fragment',
      content: runtimeConfig.extension,
    },
    {
      path: 'integration-readme.md',
      media_type: 'text/markdown',
      schema: 'meeting_app_adapter_handoff_readme',
      content: integrationReadme({ spec, runtimeConfig, manifest }),
    },
  ].filter(Boolean);
  const issues = packageIssues({ spec, runtimeConfig, files });
  const blocking = issues.filter((item) => item.severity === 'error');
  const nextActions = unique([
    ...(spec.next_actions ?? []),
    ...(runtimeConfig.next_actions ?? []),
    'capture_live_dom_snapshots_before_production_rollout',
    'verify_speaker_track_on_real_meeting',
    'verify_annotation_insert_on_current_axis',
  ]);
  return {
    type: 'meeting_app_adapter_handoff_package',
    schema: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted: blocking.length === 0,
    adapter_key: spec.adapter_key,
    display_name: spec.display_name,
    source: spec.source,
    surface: spec.surface,
    artifact_count: files.length,
    file_count: files.length,
    files,
    file_paths: files.map((file) => file.path),
    adapter_spec: spec,
    adapter_manifest: manifest,
    runtime_config: runtimeConfig,
    extension_manifest_fragment: runtimeConfig.extension,
    consumer_entrypoints: runtimeConfig.entrypoints,
    contracts: {
      timestamp_field: runtimeConfig.contracts?.timestamp_field,
      provider_events_block_realtime: runtimeConfig.contracts?.provider_events_block_realtime,
      transcript_blocks_realtime: runtimeConfig.contracts?.transcript_blocks_realtime,
      local_observer_may_start_axis: runtimeConfig.contracts?.local_observer_may_start_axis,
      required_signals: runtimeConfig.contracts?.required_signals ?? [],
      output_intents: runtimeConfig.contracts?.output_intents ?? [],
    },
    validation: {
      static_ready: blocking.length === 0,
      spec_accepted: spec.accepted === true,
      runtime_config_accepted: runtimeConfig.accepted === true,
      capture_ready: runtimeConfig.readiness?.capture_options_ready === true,
      mutation_ready: runtimeConfig.readiness?.mutation_observer_ready === true,
      content_script_ready: runtimeConfig.readiness?.content_script_ready === true,
      requires_live_snapshot_before_production: true,
      required_live_evidence: [
        'live_dom_snapshot',
        'candidate_observation',
        'speaker_track',
        'participant_track',
        'annotation_insert_current_axis',
      ],
    },
    implementation_steps: [
      'load_runtime_config_json_in_host_project',
      'install_content_script_bridge_with_content_script_options',
      'create_browser_runtime_with_browser_runtime_options',
      'capture_live_snapshots_with_capture_options',
      'run_handoff_readiness_with_live_snapshot_evidence',
    ],
    issue_count: issues.length,
    blocking_count: blocking.length,
    warning_count: issues.filter((item) => item.severity === 'warning').length,
    issues,
    next_actions: nextActions,
  };
}

export function buildMeetingAppAdapterHandoffPackageMatrix(options = {}) {
  const packageOptions = cleanOptions(options);
  const packages = packageInputs(options).map((input) => buildMeetingAppAdapterHandoffPackage(input, packageOptions));
  return {
    type: 'meeting_app_adapter_handoff_package_matrix',
    schema: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA,
    schema_version: MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA_VERSION,
    accepted: packages.every((pkg) => pkg.accepted === true),
    package_count: packages.length,
    accepted_count: packages.filter((pkg) => pkg.accepted === true).length,
    custom_count: packages.filter((pkg) => pkg.source !== 'built_in_manifest').length,
    built_in_count: packages.filter((pkg) => pkg.source === 'built_in_manifest').length,
    runtime_config_ready_count: packages.filter((pkg) => pkg.validation.runtime_config_accepted).length,
    content_script_ready_count: packages.filter((pkg) => pkg.validation.content_script_ready).length,
    live_evidence_required_count: packages.filter((pkg) => pkg.validation.requires_live_snapshot_before_production).length,
    rows: packages.map((pkg) => ({
      adapter_key: pkg.adapter_key,
      display_name: pkg.display_name,
      source: pkg.source,
      accepted: pkg.accepted,
      file_count: pkg.file_count,
      runtime_config_accepted: pkg.validation.runtime_config_accepted,
      content_script_ready: pkg.validation.content_script_ready,
      capture_ready: pkg.validation.capture_ready,
      mutation_ready: pkg.validation.mutation_ready,
      blocking_count: pkg.blocking_count,
      warning_count: pkg.warning_count,
      first_next_action: pkg.next_actions?.[0],
    })),
    packages,
    next_actions: unique(packages.flatMap((pkg) => pkg.next_actions ?? [])),
  };
}

export function assertMeetingAppAdapterHandoffPackage(packageOrSpec = {}, options = {}) {
  const handoffPackage = packageOrSpec?.schema === MEETING_APP_ADAPTER_HANDOFF_PACKAGE_SCHEMA
    ? packageOrSpec
    : buildMeetingAppAdapterHandoffPackage(packageOrSpec, options);
  if (handoffPackage.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter handoff package acceptance failed', {
      code: 'meeting_app_adapter_handoff_package_rejected',
      adapter_key: handoffPackage.adapter_key,
      issues: handoffPackage.issues,
      next_actions: handoffPackage.next_actions,
    });
  }
  return handoffPackage;
}

export function assertMeetingAppAdapterHandoffPackageMatrix(matrixOrOptions = {}, options = {}) {
  const matrix = matrixOrOptions?.schema === MEETING_APP_ADAPTER_HANDOFF_PACKAGE_MATRIX_SCHEMA
    ? matrixOrOptions
    : buildMeetingAppAdapterHandoffPackageMatrix({ ...matrixOrOptions, ...options });
  if (matrix.accepted !== true) {
    throw new MeetingTimelineSdkError('Meeting app adapter handoff package matrix acceptance failed', {
      code: 'meeting_app_adapter_handoff_package_matrix_rejected',
      rows: matrix.rows,
      next_actions: matrix.next_actions,
    });
  }
  return matrix;
}
