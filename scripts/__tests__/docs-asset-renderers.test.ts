import { execFileSync } from 'node:child_process'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)
const tsxCli = join(repoRoot, 'node_modules', 'tsx', 'dist', 'cli.mjs')

describe('documentation asset checks', () => {
  for (const script of ['render-hero-workbook-api.ts', 'render-social-preview.ts']) {
    it(`${script} verifies committed assets without a system renderer`, () => {
      const emptyPath = mkdtempSync(join(tmpdir(), 'bilig-empty-path-'))

      try {
        expect(() =>
          execFileSync(process.execPath, [tsxCli, join(repoRoot, 'scripts', script), '--check'], {
            cwd: repoRoot,
            env: {
              ...process.env,
              PATH: emptyPath,
            },
            stdio: 'pipe',
          }),
        ).not.toThrow()
      } finally {
        rmSync(emptyPath, { recursive: true, force: true })
      }
    })
  }
})
