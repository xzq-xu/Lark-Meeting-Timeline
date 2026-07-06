import type { MeetingPlatformAdapterAuthoringOptions } from './platform-adapter-authoring.mjs';

export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_ITEM_SCHEMA: 'meeting_platform_adapter_portfolio_item';
export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA: 'meeting_platform_adapter_portfolio';
export const MEETING_PLATFORM_ADAPTER_PORTFOLIO_SCHEMA_VERSION: 1;

export interface MeetingPlatformAdapterPortfolioOptions extends MeetingPlatformAdapterAuthoringOptions {
  includeArtifacts?: boolean;
  include_artifacts?: boolean;
}

export interface MeetingPlatformAdapterPortfolioItem {
  type: 'meeting_platform_adapter_portfolio_item';
  schema: 'meeting_platform_adapter_portfolio_item';
  schema_version: 1;
  platform: string;
  display_name: string;
  built_in: boolean;
  adapter_status: string;
  recommended_first_surface: string;
  p0_realtime_axis: Record<string, unknown>;
  p1_provider_reconcile: Record<string, unknown>;
  p2_post_meeting_backfill: Record<string, unknown>;
  implementation?: Record<string, unknown>;
  evidence_requirements: Record<string, unknown>;
  commands: Record<string, string>;
  artifacts?: Record<string, unknown>;
  next_actions: string[];
}

export interface MeetingPlatformAdapterPortfolio {
  type: 'meeting_platform_adapter_portfolio';
  schema: 'meeting_platform_adapter_portfolio';
  schema_version: 1;
  platform_count: number;
  built_in_count: number;
  external_authoring_count: number;
  browser_surface_ready_count: number;
  provider_reconcile_count: number;
  implementation_ready_count: number;
  pilot_ready_count: number;
  production_ready_count: number;
  platforms: string[];
  rows: Array<Record<string, unknown>>;
  items: MeetingPlatformAdapterPortfolioItem[];
  next_actions: string[];
}

export function buildMeetingPlatformAdapterPortfolioItem(
  platform: string,
  options?: MeetingPlatformAdapterPortfolioOptions,
): MeetingPlatformAdapterPortfolioItem;

export function buildMeetingPlatformAdapterPortfolio(
  options?: MeetingPlatformAdapterPortfolioOptions,
): MeetingPlatformAdapterPortfolio;
