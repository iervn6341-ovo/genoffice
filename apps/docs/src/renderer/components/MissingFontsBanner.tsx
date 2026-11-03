import { memo, useMemo } from 'react'
import type { FontSubstitution } from '../font-check'
import { useI18n } from '../i18n/locale'

export interface MissingFontsBannerProps {
  missing: FontSubstitution[]
  isDark?: boolean
  onViewDetails: () => void
  onDismiss: () => void
}

export const MissingFontsBanner = memo(function MissingFontsBanner({
  missing,
  isDark = false,
  onViewDetails,
  onDismiss,
}: MissingFontsBannerProps) {
  const { lang, t } = useI18n()
  const isZhTW = lang === 'zh-TW'
  const isZh = lang === 'zh'

  const fontNamesSummary = useMemo(() => {
    const list = missing.slice(0, 3).map((m) => m.name)
    const suffix = missing.length > 3 ? '…' : ''
    return list.join('、') + suffix
  }, [missing])

  const ui = useMemo(
    () => ({
      bannerText: isZhTW
        ? `文件包含未安裝的字型（${fontNamesSummary}），已自動以相容替代字型顯示以維持排版。`
        : isZh
        ? `文档包含未安装的字体（${fontNamesSummary}），已自动以兼容替代字体显示以保持排版。`
        : `This document contains fonts not installed on your system (${fontNamesSummary}); substitutes are applied to maintain layout.`,
      viewBtn: isZhTW ? '檢視替代對照表' : isZh ? '查看替代对照表' : 'View Font Mapping',
      dismissBtn: isZhTW ? '不再提示' : isZh ? '不再提示' : 'Dismiss',
    }),
    [isZhTW, isZh, fontNamesSummary],
  )

  if (missing.length === 0) return null

  return (
    <div
      className={`missing-fonts-banner ${isDark ? 'theme-dark' : 'theme-light'}`}
      role="status"
    >
      <div className="missing-fonts-banner-left">
        <span className="missing-fonts-banner-icon" aria-hidden>
          🔤
        </span>
        <span className="missing-fonts-banner-text">{ui.bannerText}</span>
      </div>
      <div className="missing-fonts-banner-actions">
        <button
          className="missing-fonts-banner-view-btn"
          onClick={onViewDetails}
        >
          {ui.viewBtn}
        </button>
        <button
          className="missing-fonts-banner-dismiss-btn"
          onClick={onDismiss}
          title={ui.dismissBtn}
          aria-label={t('appClose')}
        >
          ✕
        </button>
      </div>
    </div>
  )
})
