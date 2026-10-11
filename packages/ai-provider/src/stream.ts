import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { withOutputCapFallback } from './output-cap'
import { MAX_MAX_OUTPUT_TOKENS, UNLIMITED_OUTPUT_TOKENS } from './providers'
import { streamAnthropic } from './protocols/anthropic'
import { streamGemini } from './protocols/gemini'
import { streamOpenAiCompatible } from './protocols/openai-compatible'
import { streamCodexAppServer } from './codex-app-server'
import type { StreamCallbacks } from './protocols/shared'
import { getProviderAdapter, type AiProtocol } from './registry'
import type { AiProviderConfig, AiProviderId } from './types'

export { streamAnthropic } from './protocols/anthropic'
export { streamGemini } from './protocols/gemini'
export { streamOpenAiCompatible } from './protocols/openai-compatible'
export { AiCreditsError, sseLines } from './protocols/shared'
export type { StreamCallbacks } from './protocols/shared'

/** route a streaming, tool-calling-capable turn by provider id */
export async function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const endpoint = getProviderAdapter(provider).resolveEndpoint(config)
  const { baseUrl } = endpoint
  if (endpoint.protocol === 'codex-app-server') {
    return streamCodexAppServer(config, system, messages, tools, maxTokens, cb)
  }
  const protocol: Exclude<AiProtocol, 'codex-app-server'> = endpoint.protocol
  // "No limit": Anthropic's API insists on max_tokens, so ask for the ceiling (a rejection teaches
  // the real one); the other dialects simply leave the field off
  const unlimited = maxTokens === UNLIMITED_OUTPUT_TOKENS
  const requested = unlimited && protocol === 'anthropic' ? MAX_MAX_OUTPUT_TOKENS : maxTokens
  const run =
    unlimited && protocol !== 'anthropic'
      ? (fn: (cap: number) => Promise<void>) => fn(0)
      : (fn: (cap: number) => Promise<void>) =>
          withOutputCapFallback(baseUrl, config.model, requested, fn)
  return run((cap) => {
    switch (protocol) {
      case 'anthropic':
        return streamAnthropic(config, system, messages, tools, cap, cb, baseUrl)
      case 'gemini':
        return streamGemini(config, system, messages, tools, cap, cb, baseUrl, {
          omitTemperature: endpoint.omitTemperature,
        })
      case 'openai-compatible':
        return streamOpenAiCompatible(baseUrl, config, system, messages, tools, cap, cb, {
          omitTemperature: endpoint.omitTemperature,
          useMaxCompletionTokens: endpoint.useMaxCompletionTokens,
          bodyExtras: endpoint.bodyExtras,
          patientTimeouts: endpoint.patientTimeouts,
        })
    }
  })
}
