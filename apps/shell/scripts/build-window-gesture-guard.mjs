import { execFileSync } from 'node:child_process'
import { mkdirSync, rmSync, statSync, existsSync, renameSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

if (process.platform === 'darwin') {
  const script = fileURLToPath(import.meta.url)
  const root = resolve(dirname(script), '..')
  const source = resolve(root, 'native/window-gesture-guard.mm')
  const motion = resolve(root, 'native/window-gesture-motion.h')
  const output = resolve(root, 'out/native/window-gesture-guard.node')
  mkdirSync(dirname(output), { recursive: true })
  if (
    !existsSync(output) ||
    statSync(output).mtimeMs <
      Math.max(statSync(source).mtimeMs, statSync(motion).mtimeMs, statSync(script).mtimeMs)
  ) {
    const require = createRequire(import.meta.url)
    const include = require('node-api-headers').include_dir
    const slices = []
    const merged = `${output}.build-${process.pid}`
    try {
      for (const arch of ['arm64', 'x86_64']) {
        const slice = `${output}.${arch}`
        slices.push(slice)
        execFileSync(
          'xcrun',
          [
            'clang++',
            '-std=c++17',
            '-fobjc-arc',
            '-DNAPI_VERSION=8',
            '-arch',
            arch,
            '-mmacosx-version-min=12.0',
            '-bundle',
            '-undefined',
            'dynamic_lookup',
            '-framework',
            'AppKit',
            `-I${include}`,
            source,
            '-o',
            slice,
          ],
          { stdio: 'inherit' },
        )
      }
      execFileSync('lipo', ['-create', ...slices, '-output', merged], { stdio: 'inherit' })
      // A running test/released app may have the previous module mapped.
      renameSync(merged, output)
    } finally {
      for (const slice of slices) rmSync(slice, { force: true })
      rmSync(merged, { force: true })
    }
  }
}
