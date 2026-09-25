/** Browser-safe settings surface. Keep Node-backed transports out of renderer bundles. */
export type {
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  CodexModelCatalog,
} from './types'
export {
  AI_PROVIDERS,
  DEFAULT_MAX_OUTPUT_TOKENS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  clampMaxOutputTokens,
  providerRequiresApiKey,
} from './providers'
export { LOCAL_ENDPOINT_PRESETS, isLocalEndpoint, presetForBaseUrl } from './local'
export type { LocalEndpointPreset } from './local'
export type { ListModelsResult } from './list-models'
export { getProviderAdapter, modelLacksVision } from './registry'
export { AI_MEDIA_PROVIDERS, imageGenerationAvailable, mediaAnalysisAvailable } from './media'
export { AI_SEARCH_PROVIDERS } from './search-settings'
