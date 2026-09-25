import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    // Node 25+ defines its own localStorage (undefined without --localstorage-file), which
    // shadows jsdom's and breaks every storage test; turn Node's off
    execArgv: ['--no-experimental-webstorage'],
    testTimeout: 20000,
  },
})
