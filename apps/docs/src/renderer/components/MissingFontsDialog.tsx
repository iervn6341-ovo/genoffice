import { memo, useMemo, useState } from 'react'
import type { FontSubstitution } from '../font-check'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import { defaultEastAsiaFontFor } from '../font-list'

export interface MissingFontsDialogProps {
  missing: FontSubstitution[]
  filePath?: string | null
  fileName?: string
  isDark?: boolean
  onClose: () => void
}

/**
 * Resolves a readable substitute font name and description for a missing font.
 */
function resolveSubstituteDetails(item: FontSubstitution, isZhTW: boolean, isZh: boolean) {
  const original = item.name
  let substitute = item.substitute

  // Fallback to recommended compatible bundled/system face
  if (!substitute) {
    if (/明體|新細明體|細明體|pmingliu|mingliu/i.test(original)) {
      substitute = 'GenOffice Songti TC'
    } else if (/宋体|simsun|nsimsun/i.test(original)) {
      substitute = 'GenOffice Songti SC'
    } else if (/正黑體|jhenghei|pingfang tc/i.test(original)) {
      substitute = 'PingFang TC / GenOffice Sans TC'
    } else if (/微软雅黑|yahei|pingfang sc/i.test(original)) {
      substitute = 'PingFang SC / GenOffice Sans SC'
    } else if (/標楷體|kai|dfkai/i.test(original)) {
      substitute = 'BiauKai / KaiTi'
    } else {
      substitute = defaultEastAsiaFontFor(isZhTW ? 'zh-TW' : 'zh') || 'GenOffice Songti TC'
    }
  }

  // Determine font family category description
  let category = isZhTW ? '相容替代字型' : isZh ? '兼容替代字体' : 'Compatible Substitute'
  if (/明體|宋體|songti|mingliu|serif/i.test(substitute)) {
    category = isZhTW ? '內建明體 / 宋體相容' : isZh ? '内置明体 / 宋体兼容' : 'Serif Compatible'
  } else if (/黑體|sans|gothic/i.test(substitute)) {
    category = isZhTW ? '內建黑體相容' : isZh ? '内置黑体兼容' : 'Sans-serif Compatible'
  } else if (/楷/i.test(substitute)) {
    category = isZhTW ? '楷書體相容' : isZh ? '楷体兼容' : 'KaiTI Compatible'
  }

  return { substitute, category }
}

export const MissingFontsDialog = memo(function MissingFontsDialog({
  missing,
  filePath,
  fileName,
  isDark = false,
  onClose,
}: MissingFontsDialogProps) {
  const { lang, t } = useI18n()
  const isZhTW = lang === 'zh-TW'
  const isZh = lang === 'zh'
  const modalKeys = useModalKeys(onClose)

  const docKey = filePath || fileName || 'default'
  const storageKey = `aidocs.suppressFontModal.${docKey}`

  const [dontShowAgain, setDontShowAgain] = useState(() => {
    try {
      return localStorage.getItem(storageKey) === '1'
    } catch {
      return false
    }
  })

  const ui = useMemo(
    () => ({
      title: isZhTW ? '文件字型相容性提示' : isZh ? '文档字体兼容性提示' : 'Font Compatibility Notice',
      desc: isZhTW
        ? '此文件使用了本機未安裝的字型。GenOffice 已自動匹配高相容度的替代字型，以確保文字排版、換行與段落結構與原始文件完全一致。'
        : isZh
        ? '此文档使用了本机未安装的字体。GenOffice 已自动匹配高兼容度的替代字体，以保持排版与版面结构一致。'
        : 'This document uses fonts not installed on this system. GenOffice has automatically matched compatible substitute fonts to preserve layout and pagination.',
      colOriginal: isZhTW ? '原始文件字型' : isZh ? '原始文档字体' : 'Original Document Font',
      colSubstitute: isZhTW ? '相容替代字型' : isZh ? '兼容替代字体' : 'Substitute Font',
      colStatus: isZhTW ? '排版相容性狀態' : isZh ? '排版兼容状态' : 'Layout Compatibility',
      activeStatus: isZhTW ? '已自動替代顯示' : isZh ? '已自动替代显示' : 'Substituted Active',
      dontShowAgainLabel: isZhTW ? '此文件不再自動彈出提示' : isZh ? '此文档不再自动弹出提示' : "Don't show again for this document",
      okBtn: isZhTW ? '我知道了' : isZh ? '我知道了' : 'Got it',
      installedBadge: isZhTW ? '系統未安裝' : isZh ? '系统未安装' : 'Not installed',
    }),
    [isZhTW, isZh],
  )

  const handleConfirm = () => {
    try {
      if (dontShowAgain) {
        localStorage.setItem(storageKey, '1')
      } else {
        localStorage.removeItem(storageKey)
      }
    } catch {
      // Ignore storage write issues
    }
    onClose()
  }

  return (
    <div
      className="modal-backdrop missing-fonts-modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className={`modal missing-fonts-modal ${isDark ? 'theme-dark' : 'theme-light'}`}>
        <div className="missing-fonts-header">
          <div className="missing-fonts-title-wrap">
            <span className="missing-fonts-title-icon" aria-hidden>
              🔤
            </span>
            <h2>{ui.title}</h2>
          </div>
          <button
            className="missing-fonts-close-btn"
            onClick={onClose}
            aria-label={t('appClose')}
          >
            ✕
          </button>
        </div>

        <p className="missing-fonts-desc">{ui.desc}</p>

        <div className="missing-fonts-table-wrap">
          <table className="missing-fonts-table">
            <thead>
              <tr>
                <th>{ui.colOriginal}</th>
                <th>{ui.colSubstitute}</th>
                <th>{ui.colStatus}</th>
              </tr>
            </thead>
            <tbody>
              {missing.map((item) => {
                const { substitute, category } = resolveSubstituteDetails(item, isZhTW, isZh)
                return (
                  <tr key={item.name}>
                    <td className="font-original-cell">
                      <span className="font-name">{item.name}</span>
                      <span className="font-missing-tag">{ui.installedBadge}</span>
                    </td>
                    <td className="font-substitute-cell">
                      <span className="font-arrow">➔</span>
                      <span className="substitute-name">{substitute}</span>
                    </td>
                    <td className="font-status-cell">
                      <span className="status-badge">{category}</span>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>

        <div className="missing-fonts-footer">
          <label className="missing-fonts-checkbox-label">
            <input
              type="checkbox"
              checked={dontShowAgain}
              onChange={(e) => setDontShowAgain(e.target.checked)}
            />
            <span>{ui.dontShowAgainLabel}</span>
          </label>

          <button className="btn-primary missing-fonts-confirm-btn" onClick={handleConfirm}>
            {ui.okBtn}
          </button>
        </div>
      </div>
    </div>
  )
})
