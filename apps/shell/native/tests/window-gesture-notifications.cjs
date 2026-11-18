// Run after the shell build on macOS 27.2. No injected input or frame changes.
const { createRequire } = require('node:module')
const { resolve, join } = require('node:path')
const { mkdtempSync, writeFileSync, rmSync } = require('node:fs')
const { tmpdir } = require('node:os')
const { execFileSync } = require('node:child_process')
const assert = require('node:assert/strict')
const root = resolve(__dirname, '../../../..')
const req = createRequire(join(root, 'package.json'))

async function main() {
  if (process.platform !== 'darwin') return console.log('AppKit checks skipped')
  const scratch = mkdtempSync(join(tmpdir(), 'genoffice-native-notifications-'))
  const modulePath = join(scratch, 'notification-tests.node')
  let app
  try {
    const motionTest = join(scratch, 'motion-test')
    execFileSync(
      'xcrun',
      ['clang++', '-std=c++17', join(__dirname, 'window-gesture-motion.cpp'), '-o', motionTest],
      { stdio: 'inherit' },
    )
    execFileSync(motionTest, [], { stdio: 'inherit' })
    execFileSync(
      'xcrun',
      [
        'clang++',
        '-std=c++17',
        '-fobjc-arc',
        '-bundle',
        '-undefined',
        'dynamic_lookup',
        '-DNAPI_VERSION=8',
        `-I${req('node-api-headers').include_dir}`,
        '-framework',
        'AppKit',
        join(__dirname, 'window-gesture-notifications.mm'),
        '-o',
        modulePath,
      ],
      { stdio: 'inherit' },
    )
    writeFileSync(
      join(scratch, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, autoSave: false }),
    )
    const env = {
      ...process.env,
      GENOFFICE_USER_DATA: scratch,
      GENOFFICE_E2E_BACKGROUND: '1',
      GENOFFICE_NO_SPARE_VIEW: '1',
    }
    delete env.ELECTRON_RUN_AS_NODE
    app = await req('playwright')._electron.launch({
      executablePath: req('electron'),
      args: [join(root, 'apps/shell')],
      env,
    })
    await app.firstWindow()
    const result = await app.evaluate(({ BrowserWindow }, modulePath) => {
      if (!process.getSystemVersion().startsWith('27.2')) return null
      const req = process.getBuiltinModule('module').createRequire(process.cwd() + '/package.json')
      return JSON.parse(
        req(modulePath).run(BrowserWindow.getAllWindows()[0].getNativeWindowHandle()),
      )
    }, modulePath)
    if (!result) return console.log('AppKit checks skipped outside macOS 27.2')
    for (const [name, passed] of Object.entries(result)) assert.equal(passed, true, name)
    console.log(`AppKit notification checks passed (${Object.keys(result).length} checks)`)
  } finally {
    if (app) await app.close()
    rmSync(scratch, { recursive: true, force: true })
  }
}
main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
