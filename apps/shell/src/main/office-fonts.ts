import { copyFileSync, existsSync, mkdirSync, readdirSync } from 'node:fs'
import { homedir, platform } from 'node:os'
import { join } from 'node:path'
import { shell } from 'electron'
import type { AcquireOfficeFontResult } from '../shared/home-api'

/** Known Office font file mappings (lowercase font family -> candidate filenames) */
const OFFICE_FONT_FILE_MAP: Record<string, string[]> = {
  // Traditional Chinese
  '微軟正黑體': ['MSJH.ttf', 'MSJHBD.ttf'],
  'microsoft jhenghei': ['MSJH.ttf', 'MSJHBD.ttf'],
  'microsoft jhenghei ui': ['MSJH.ttf', 'MSJHBD.ttf'],
  '新細明體': ['mingliu.ttc', 'mingliub.ttc'],
  '細明體': ['mingliu.ttc', 'mingliub.ttc'],
  'pmingliu': ['mingliu.ttc', 'mingliub.ttc'],
  'mingliu': ['mingliu.ttc', 'mingliub.ttc'],
  'pmingliu-extb': ['mingliu.ttc', 'mingliub.ttc'],
  'mingliu-extb': ['mingliu.ttc', 'mingliub.ttc'],
  '標楷體': ['DFKai-SB.ttf', 'kaiu.ttf', 'Kaiti.ttf'],
  'dfkai-sb': ['DFKai-SB.ttf', 'kaiu.ttf', 'Kaiti.ttf'],
  'biaukai': ['DFKai-SB.ttf', 'kaiu.ttf', 'Kaiti.ttf'],

  // Simplified Chinese
  '微軟雅黑': ['msyh.ttc', 'msyhbd.ttc', 'msyhl.ttc'],
  'microsoft yahei': ['msyh.ttc', 'msyhbd.ttc', 'msyhl.ttc'],
  'microsoft yahei ui': ['msyh.ttc', 'msyhbd.ttc', 'msyhl.ttc'],
  '宋體': ['Simsun.ttc', 'simsunb.ttf'],
  '宋体': ['Simsun.ttc', 'simsunb.ttf'],
  'simsun': ['Simsun.ttc', 'simsunb.ttf'],
  'simsun-extb': ['Simsun.ttc', 'simsunb.ttf'],
  'nsimsun': ['Simsun.ttc', 'simsunb.ttf'],
  '黑體': ['SimHei.ttf'],
  '黑体': ['SimHei.ttf'],
  'simhei': ['SimHei.ttf'],
  '楷體': ['Kaiti.ttf'],
  '楷体': ['Kaiti.ttf'],
  'kaiti': ['Kaiti.ttf'],
  '仿宋': ['Fangsong.ttf'],
  'fangsong': ['Fangsong.ttf'],
  '等線': ['Deng.ttf', 'Dengb.ttf', 'Dengl.ttf'],
  '等线': ['Deng.ttf', 'Dengb.ttf', 'Dengl.ttf'],
  'dengxian': ['Deng.ttf', 'Dengb.ttf', 'Dengl.ttf'],
  '等線 light': ['Dengl.ttf', 'Deng.ttf'],
  '等线 light': ['Dengl.ttf', 'Deng.ttf'],
  'dengxian light': ['Dengl.ttf', 'Deng.ttf'],

  // Japanese Office
  'yu gothic': ['YuGothR.ttc', 'YuGothM.ttc', 'YuGothB.ttc', 'YuGothL.ttc'],
  'yu gothic medium': ['YuGothM.ttc', 'YuGothR.ttc'],
  'yu mincho': ['yumin.ttf', 'yumindb.ttf', 'yuminl.ttf'],
  'meiryo': ['meiryo.ttc', 'meiryob.ttc'],
  'meiryo ui': ['meiryo.ttc', 'meiryob.ttc'],
  'ms gothic': ['msgothic.ttc'],
  'ms pgothic': ['msgothic.ttc'],
  'ms ui gothic': ['msgothic.ttc'],
  'ms mincho': ['msmincho.ttc'],
  'ms pmincho': ['msmincho.ttc'],

  // Korean Office
  'malgun gothic': ['malgun.ttf', 'malgunbd.ttf', 'malgunsl.ttf'],
  'batang': ['batang.ttc'],
  'batangche': ['batang.ttc'],
  'gulim': ['gulim.ttc'],
  'gulimche': ['gulim.ttc'],
  'dotum': ['gulim.ttc'],
  'dotumche': ['gulim.ttc'],
  'gungsuh': ['batang.ttc'],

  // Western / Office core
  'calibri': [
    'Calibri.ttf',
    'Calibrib.ttf',
    'Calibrii.ttf',
    'Calibriz.ttf',
    'calibril.ttf',
    'calibrili.ttf',
  ],
  'calibri light': ['calibril.ttf', 'calibrili.ttf', 'Calibri.ttf'],
  'cambria': ['Cambria.ttc', 'Cambriab.ttf', 'Cambriai.ttf', 'Cambriaz.ttf'],
  'cambria math': ['Cambria.ttc'],
  'aptos': [
    'Aptos.ttf',
    'Aptos-Bold.ttf',
    'Aptos-Italic.ttf',
    'Aptos-Bold-Italic.ttf',
    'Aptos-Black.ttf',
    'Aptos-Black-Italic.ttf',
    'Aptos-ExtraBold.ttf',
    'Aptos-ExtraBold-Italic.ttf',
    'Aptos-Light.ttf',
    'Aptos-Light-Italic.ttf',
    'Aptos-SemiBold.ttf',
    'Aptos-SemiBold-Italic.ttf',
    'Aptos-Narrow.ttf',
    'Aptos-Narrow-Bold.ttf',
    'Aptos-Narrow-Italic.ttf',
    'Aptos-Narrow-Bold-Italic.ttf',
  ],
  'aptos display': ['Aptos.ttf', 'Aptos-Bold.ttf'],
  'aptos narrow': [
    'Aptos-Narrow.ttf',
    'Aptos-Narrow-Bold.ttf',
    'Aptos-Narrow-Italic.ttf',
    'Aptos-Narrow-Bold-Italic.ttf',
  ],
  'consolas': ['Consola.ttf', 'Consolab.ttf', 'Consolai.ttf', 'Consolaz.ttf'],
  'segoe ui': [
    'segoeui.ttf',
    'segoeuib.ttf',
    'seguihis.ttf',
    'seguisym.ttf',
    'Segoe Print Bold.ttf',
    'Segoe Script Bold.ttf',
  ],
  'arial': ['arial.ttf', 'arialbd.ttf', 'arialbi.ttf', 'ariali.ttf', 'ariblk.ttf'],
  'times new roman': ['times.ttf', 'timesbd.ttf', 'timesbi.ttf', 'timesi.ttf'],
  'tahoma': ['tahoma.ttf', 'tahomabd.ttf'],
  'verdana': ['Verdana.ttf', 'Verdana Bold.ttf', 'Verdana Italic.ttf', 'Verdana Bold Italic.ttf'],
}

/** Get the user-level fonts directory for the current operating system */
export function getUserFontsDirectory(): string {
  const currentPlatform = platform()
  if (currentPlatform === 'darwin') {
    return join(homedir(), 'Library', 'Fonts')
  }
  if (currentPlatform === 'win32') {
    const localAppData =
      process.env.LOCALAPPDATA || join(homedir(), 'AppData', 'Local')
    return join(localAppData, 'Microsoft', 'Windows', 'Fonts')
  }
  return join(homedir(), '.local', 'share', 'fonts')
}

/** Get candidate directories where Microsoft Office font files may reside */
function getCandidateOfficeDirectories(): string[] {
  const currentPlatform = platform()
  if (currentPlatform === 'darwin') {
    return [
      '/Applications/Microsoft Word.app/Contents/Resources/DFonts',
      '/Applications/Microsoft Excel.app/Contents/Resources/DFonts',
      '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts',
      '/Applications/Microsoft Outlook.app/Contents/Resources/DFonts',
      '/Applications/Microsoft 365.app/Contents/Resources/DFonts',
      join(homedir(), 'Applications/Microsoft Word.app/Contents/Resources/DFonts'),
      join(homedir(), 'Applications/Microsoft Excel.app/Contents/Resources/DFonts'),
      join(homedir(), 'Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts'),
      '/Library/Fonts/Microsoft',
    ]
  }
  if (currentPlatform === 'win32') {
    return [
      'C:\\Program Files\\Microsoft Office\\root\\vfs\\Fonts',
      'C:\\Program Files (x86)\\Microsoft Office\\root\\vfs\\Fonts',
      'C:\\Windows\\Fonts',
    ]
  }
  return []
}

/**
 * Clean and simplify a font family string for fuzzy matching against font filenames
 */
function cleanFontToken(str: string): string {
  return str.toLowerCase().replace(/[\s\-_]+/g, '')
}

/**
 * Find candidate font files in candidate Office directories
 */
function findCandidateFontFiles(
  fontFamily: string,
  officeDirs: string[],
): { srcDir: string; fileName: string }[] {
  const normKey = fontFamily.trim().toLowerCase()
  const explicitFiles = OFFICE_FONT_FILE_MAP[normKey] || []
  const token = cleanFontToken(fontFamily)

  const found: { srcDir: string; fileName: string }[] = []
  const seenFiles = new Set<string>()

  for (const dir of officeDirs) {
    if (!existsSync(dir)) continue
    let dirEntries: string[] = []
    try {
      dirEntries = readdirSync(dir)
    } catch {
      continue
    }

    // 1. Check explicit mappings
    for (const targetFile of explicitFiles) {
      const match = dirEntries.find(
        (entry) => entry.toLowerCase() === targetFile.toLowerCase(),
      )
      if (match && !seenFiles.has(match.toLowerCase())) {
        seenFiles.add(match.toLowerCase())
        found.push({ srcDir: dir, fileName: match })
      }
    }

    // 2. Fuzzy match if no explicit files matched or to discover additional family faces
    if (token.length >= 3) {
      for (const entry of dirEntries) {
        if (!/\.(ttf|ttc|otf)$/i.test(entry)) continue
        const entryToken = cleanFontToken(entry.replace(/\.(ttf|ttc|otf)$/i, ''))
        if (
          (entryToken.startsWith(token) || token.startsWith(entryToken)) &&
          !seenFiles.has(entry.toLowerCase())
        ) {
          seenFiles.add(entry.toLowerCase())
          found.push({ srcDir: dir, fileName: entry })
        }
      }
    }
  }

  return found
}

/**
 * Handle acquiring an Office font and copying it into the user fonts directory.
 */
export async function handleAcquireOfficeFont(
  fontFamily: string,
): Promise<AcquireOfficeFontResult> {
  const trimmed = fontFamily.trim()
  if (!trimmed) {
    return {
      success: false,
      fontName: trimmed,
      reason: 'error',
      message: 'Font family name is empty.',
    }
  }

  const officeDirs = getCandidateOfficeDirectories()
  const existingOfficeDirs = officeDirs.filter((d) => existsSync(d))

  // If no Office directory exists on the machine
  if (existingOfficeDirs.length === 0) {
    return {
      success: false,
      fontName: trimmed,
      reason: 'no-office',
      message: 'Microsoft Office installation was not detected on this system.',
    }
  }

  const matches = findCandidateFontFiles(trimmed, existingOfficeDirs)
  if (matches.length === 0) {
    return {
      success: false,
      fontName: trimmed,
      reason: 'not-found',
      message: `Font file for "${trimmed}" was not found in local Office installation.`,
    }
  }

  const userFontsDir = getUserFontsDirectory()
  try {
    if (!existsSync(userFontsDir)) {
      mkdirSync(userFontsDir, { recursive: true })
    }
  } catch (err) {
    return {
      success: false,
      fontName: trimmed,
      reason: 'error',
      message: `Failed to create fonts directory: ${String(err)}`,
    }
  }

  let copiedCount = 0
  for (const { srcDir, fileName } of matches) {
    const srcPath = join(srcDir, fileName)
    const destPath = join(userFontsDir, fileName)
    try {
      copyFileSync(srcPath, destPath)
      copiedCount++
    } catch {
      // Continue copying remaining variant files if one fails
    }
  }

  if (copiedCount === 0) {
    return {
      success: false,
      fontName: trimmed,
      reason: 'error',
      message: `Failed to copy font files for "${trimmed}".`,
    }
  }

  return {
    success: true,
    fontName: trimmed,
    copiedCount,
    destPath: userFontsDir,
    message: `Successfully acquired ${copiedCount} font file(s) for "${trimmed}".`,
  }
}

/**
 * Handle opening the user fonts directory in system file explorer.
 */
export async function handleOpenFontsFolder(): Promise<void> {
  const dir = getUserFontsDirectory()
  try {
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true })
    }
    await shell.openPath(dir)
  } catch {
    // Silently handle explorer errors
  }
}
