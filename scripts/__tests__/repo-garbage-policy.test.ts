import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

const repoRoot = new URL('../..', import.meta.url)
const tsIgnoreDirective = ['@ts-', 'ignore'].join('')
const tsExpectErrorDirective = ['@ts-', 'expect-error'].join('')
const oxlintDisableDirective = ['oxlint-', 'disable'].join('')

function gitLsFiles(patterns: readonly string[] = []): string[] {
  return execFileSync('git', ['ls-files', ...patterns], {
    cwd: repoRoot,
    encoding: 'utf8',
  })
    .split('\n')
    .filter((line) => line.length > 0)
    .filter((line) => existsSync(join(repoRoot.pathname, line)))
}

function readTrackedFile(path: string): string {
  return readFileSync(join(repoRoot.pathname, path), 'utf8')
}

function trackedSourceFiles(patterns: readonly string[]): string[] {
  return gitLsFiles(patterns).filter((path) => /\.(?:[cm]?[tj]sx?|json|md|ya?ml)$/u.test(path))
}

describe('repository garbage policy', () => {
  it('does not track disposable archive or output artifacts', () => {
    expect(gitLsFiles(['docs/archive/**'])).toEqual([])
    expect(gitLsFiles(['output/**'])).toEqual([])
    expect(gitLsFiles(['docs/superpowers/plans/**'])).toEqual([])
  })

  it('keeps internal dominance and scorecard artifacts out of the tracked repo', () => {
    const allowedScorecardPaths = new Set(['.github/workflows/scorecard.yml', '.github/workflows/scorecard-code-scanning.yml'])
    const blocked = gitLsFiles().filter((path) => {
      if (allowedScorecardPaths.has(path)) {
        return false
      }
      return /(?:dominance|leadership|public-workbook-corpus-scorecard|fixture-corpus-scorecard|fidelity-scorecard|slo-scorecard)/u.test(
        path,
      )
    })

    expect(blocked).toEqual([])
  })

  it('keeps default browser E2E files free of registered skips', () => {
    const e2eFiles = gitLsFiles(['e2e/tests/*.pw.ts']).filter((path) => !path.includes('/prod-smoke.'))
    const violations = e2eFiles.flatMap((path) => {
      const source = readTrackedFile(path)
      return /(?:\btest\.skip\b|\bdescribe\.skip\b|\bit\.skip\b|skip\.bind\(test\))/u.test(source) ? [path] : []
    })

    expect(violations).toEqual([])
  })

  it('keeps browser entrypoints and web worker imports free of static Node builtins', () => {
    const browserFiles = [
      'packages/xlsx/src/browser.ts',
      'packages/excel-import/src/browser.ts',
      'apps/web/src/workbook-import-preview.worker.ts',
    ]
    const violations = browserFiles.flatMap((path) => {
      const source = readTrackedFile(path)
      return /(?:from\s+['"]node:|import\s*\(\s*['"]node:)/u.test(source) ? [path] : []
    })

    expect(violations).toEqual([])
  })

  it('does not import Node-only workbook import surfaces from web runtime code', () => {
    const webRuntimeFiles = trackedSourceFiles(['apps/web/src/**']).filter((path) => !path.includes('/__tests__/'))
    const violations = webRuntimeFiles.flatMap((path) => {
      const source = readTrackedFile(path)
      return source.includes("from '@bilig/excel-import'") || source.includes('from "@bilig/excel-import"') ? [path] : []
    })

    expect(violations).toEqual([])
  })

  it('requires lint and type suppressions to document the runtime reason', () => {
    const files = trackedSourceFiles(['apps/**', 'packages/**', 'scripts/**', 'e2e/**']).filter(
      (path) => !path.includes('/dist/') && !path.endsWith('.json'),
    )
    const violations: string[] = []
    for (const path of files) {
      const lines = readTrackedFile(path).split('\n')
      lines.forEach((line, index) => {
        if (line.includes(tsIgnoreDirective)) {
          violations.push(`${path}:${String(index + 1)} uses ${tsIgnoreDirective}`)
        }
        if (
          line.includes(tsExpectErrorDirective) &&
          !/(runtime guard|JS caller|JS callers|js-caller|legacy JavaScript|bare strings)/iu.test(line)
        ) {
          violations.push(`${path}:${String(index + 1)} has an unjustified ${tsExpectErrorDirective}`)
        }
        if (line.includes(oxlintDisableDirective) && !line.includes('--')) {
          violations.push(`${path}:${String(index + 1)} has an unjustified ${oxlintDisableDirective}`)
        }
      })
    }

    expect(violations).toEqual([])
  })
})
