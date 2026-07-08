import { execFileSync } from 'node:child_process'
import { resolve } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = resolve(new URL('../..', import.meta.url).pathname)

describe('repository garbage policy', () => {
  it('passes the repository garbage gate', () => {
    const output = execFileSync('bun', ['scripts/check-repo-garbage.ts'], {
      cwd: repoRoot,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    })

    expect(output).toContain('Repository garbage policy passed.')
  })
})
