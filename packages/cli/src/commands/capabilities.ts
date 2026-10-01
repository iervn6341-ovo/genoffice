import {
  activeMediaProvider,
  activeSearchProvider,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
} from '@genoffice/ai-provider'
import { readAiSettingsFile } from '@genoffice/ai-search'
import { aiSettingsPath, prepareCloud } from '../cloud'
import type { CommandDef } from '../registry'
import { appLaunch } from '../resources'

export const capabilitiesCommand: CommandDef = {
  name: 'capabilities',
  summary:
    'Report which cloud features (search, image search, image generation, media analysis) are configured in GenOffice, and whether the app is installed.',
  usage: 'capabilities',
  async run(_args, ctx) {
    await prepareCloud(ctx.env)
    const settings = readAiSettingsFile(aiSettingsPath(ctx.env))
    const publicSearch = true
    const searchProvider = activeSearchProvider(settings)
    const keyedSearch = searchProvider !== 'auto'
    const search = publicSearch || keyedSearch
    const imageSearch = publicSearch || searchProvider === 'serper'
    const imageGeneration = imageGenerationAvailable(settings)
    const mediaAnalysis = mediaAnalysisAvailable(settings)
    const via = (byok: string | null | undefined) => (byok ? byok : publicSearch ? 'auto' : null)
    const detail = {
      search: {
        available: search,
        via: keyedSearch ? searchProvider : publicSearch ? 'auto' : null,
      },
      image_search: {
        available: imageSearch,
        via: searchProvider === 'serper' ? 'serper' : publicSearch ? 'auto' : null,
      },
      image_generation: {
        available: imageGeneration,
        via: imageGeneration ? via(activeMediaProvider(settings, 'image')) : null,
      },
      media_analysis: {
        available: mediaAnalysis,
        via: mediaAnalysis ? via(activeMediaProvider(settings, 'analysis')) : null,
      },
      app: { available: appLaunch(ctx.env) !== null },
      settings_path: aiSettingsPath(ctx.env),
    }
    const on = Object.entries(detail)
      .filter(([k, v]) => k !== 'settings_path' && (v as { available: boolean }).available)
      .map(([k]) => k)
    return {
      summary: on.length
        ? `configured: ${on.join(', ')}`
        : 'no cloud feature configured; the app is not installed',
      detail,
    }
  },
}
