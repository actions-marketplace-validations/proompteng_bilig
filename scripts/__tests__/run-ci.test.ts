import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { resolveCiProfile, resolveCiSkipBrowserGates } from '../run-ci-config.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

function isStringRecord(value: unknown): value is Record<string, string> {
  return typeof value === 'object' && value !== null && Object.values(value).every((entry) => typeof entry === 'string')
}

function readPackageScripts(): Record<string, string> {
  const parsed: unknown = JSON.parse(readFileSync(resolve(repoRoot, 'package.json'), 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || !('scripts' in parsed) || !isStringRecord(parsed.scripts)) {
    throw new Error('package.json must define string scripts.')
  }
  return parsed.scripts
}

describe('run-ci', () => {
  it('defaults to the fast CI profile and accepts explicit profiles', () => {
    expect(resolveCiProfile({})).toBe('fast')
    expect(resolveCiProfile({ BILIG_CI_PROFILE: 'fast' })).toBe('fast')
    expect(resolveCiProfile({ BILIG_CI_PROFILE: 'full' })).toBe('full')
  })

  it('rejects malformed CI profiles instead of silently downgrading gates', () => {
    expect(() => resolveCiProfile({ BILIG_CI_PROFILE: 'ful' })).toThrow('BILIG_CI_PROFILE must be "fast" or "full", got ful')
  })

  it('resolves the browser gate skip flag strictly', () => {
    expect(resolveCiSkipBrowserGates({})).toBe(false)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: '1' })).toBe(true)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'true' })).toBe(true)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: '0' })).toBe(false)
    expect(resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'false' })).toBe(false)
    expect(() => resolveCiSkipBrowserGates({ BILIG_CI_SKIP_BROWSER: 'yes' })).toThrow(
      'BILIG_CI_SKIP_BROWSER must be "1", "true", "0", or "false" when set, got yes',
    )
  })

  it('serializes generated checks and avoids pnpm for direct preflight gates', () => {
    const source = readFileSync(resolve(repoRoot, 'scripts/run-ci.ts'), 'utf8')

    expect(source).toContain('const generatedSourceChecks: readonly CiTask[] = [')
    expect(source).toContain("await runSequential('generated-source checks', generatedSourceChecks)")
    expect(source).toContain("const semanticFastGate = pnpm('semantic correctness fast gate', 'test:semantic:fast')")
    expect(source).toContain("await runSequential('semantic correctness checks', [semanticFastGate])")
    expect(source).toContain("await runSequential('static package build prerequisites'")
    expect(source).toContain("bunScript('protocol check', 'scripts/gen-protocol.ts', '--check')")
    expect(source).toContain(
      "direct('protocol package build for generated-source imports', workspaceBin('tsc'), '-p', 'packages/protocol/tsconfig.json')",
    )
    expect(source).toContain('const wasmBuildTask: CiTask = {')
    expect(source).toContain(
      "const xlsxCorrectnessLane = directPackageScript('correctness workbook import/export', 'test:correctness:xlsx')",
    )
    expect(source).toContain("directPackageScript('correctness Desktop Excel oracle harness', 'test:correctness:excel-oracle')")
    expect(source).toContain("await runSequential('workbook correctness checks', [xlsxCorrectnessLane, excelOracleCorrectnessLane])")
    expect(source).toContain('function vitestChunkEnv(defaultChunkSize: string): Readonly<Record<string, string>>')
    expect(source).toContain("BILIG_VITEST_FILE_CHUNK_SIZE: process.env['BILIG_VITEST_FILE_CHUNK_SIZE'] ?? defaultChunkSize")
    expect(source).toContain("withEnv(directPackageScript('correctness core', 'test:correctness:core'), vitestChunkEnv('10'))")
    expect(source).toContain("withEnv(directPackageScript('correctness formula', 'test:correctness:formula'), vitestChunkEnv('3'))")
    expect(source).not.toContain(
      [
        "withEnv(\n  directPackageScript('correctness XLSX fixture corpus', 'test",
        'correctness',
        "corpus'),\n  vitestChunkEnv('10'),\n)",
      ].join(':'),
    )
    expect(source).not.toContain('research:xlsx-fixture-corpus:resume-financial:check')
    expect(source).not.toContain('research:xlsx-fixture-corpus:test')
    expect(source).not.toContain('XLSX fixture corpus completion audit check')
    expect(source).not.toContain(['bilig', 'dominance', 'scorecard check'].join(' '))
    expect(source).not.toContain(['bilig', 'dominance', 'audit check'].join(' '))
    expect(source).not.toContain('auditability scorecard check')
    expect(source).not.toContain('reliability scorecard check')
    expect(source).not.toContain('collaboration scorecard check')
    expect(source).not.toContain('automation scorecard check')
    expect(source).not.toContain('security posture scorecard check')
    expect(source).not.toContain('formula dominance check')
    expect(source).not.toContain('scripts/gen-formula-dominance-snapshot.ts')
    expect(source).toContain('formula compatibility snapshot check')
    expect(source).toContain("directPackageScript('xlsx import synthetic memory gate', 'xlsx-import:memory-gate:synthetic')")
    expect(source).not.toContain("directPackageScript('XLSX fixture corpus memory gate', 'research:xlsx-fixture-corpus:memory-gate')")
    expect(source).not.toContain('UI same-corpus XLSX fixture check')
    expect(source).not.toContain('scripts/capture-ui-responsiveness-same-corpus.ts')
    expect(source).not.toContain('WorkPaper TrueCalc scalar benchmark check')
    expect(source).not.toContain('WorkPaper xlsx-calc benchmark check')
    expect(source).not.toContain('WorkPaper IronCalc Rust benchmark check')
    expect(source).not.toContain('WorkPaper Univer benchmark check')
    expect(source).not.toContain('UI responsiveness live browser scorecard check')
    expect(source).not.toContain('scripts/gen-ui-responsiveness-live-browser-scorecard.ts')
    expect(source).toContain("bunScript('agent discovery docs check', 'scripts/sync-agent-discovery-docs.ts', '--check')")
    expect(source).toContain("await runSequential('static direct checks'")
    expect(source).toMatch(
      /direct\(\s*'dependency policy check',\s*workspaceBin\('vitest'\),\s*'run',\s*'scripts\/__tests__\/dependency-policy\.test\.ts',\s*'--configLoader',\s*'runner',?\s*\)/u,
    )
    expect(source).not.toContain("pnpm('protocol check'")
    expect(source).not.toContain("pnpm('wasm build'")
    expect(source).not.toContain("pnpm('correctness XLSX fixture corpus'")
    expect(source).not.toContain(['test', 'correctness', 'corpus'].join(':'))
    expect(source).not.toContain("await runStage('generated-source checks'")
    expect(source).not.toContain("await runStage('static package build prerequisites'")
    expect(source).not.toContain("await runStage('static direct checks'")
  })

  it('runs the CI orchestrator through tsx instead of bun', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const scripts = readPackageScripts()

    expect(scripts['ci']).toBe('BILIG_CI_PROFILE=fast tsx scripts/run-ci.ts')
    expect(scripts['ci']).not.toContain('BILIG_CI_SKIP_BROWSER')
    expect(scripts['ci:core']).toBe('BILIG_CI_PROFILE=fast BILIG_CI_SKIP_BROWSER=1 tsx scripts/run-ci.ts')
    expect(scripts['ci:github']).toBe('BILIG_CI_PROFILE=fast BILIG_CI_SKIP_BROWSER=1 tsx scripts/run-ci.ts')
    expect(scripts['ci:full']).toBe('BILIG_CI_PROFILE=full tsx scripts/run-ci.ts')
    expect(scripts['auditability:check']).toBeUndefined()
    expect(scripts['reliability:check']).toBeUndefined()
    expect(scripts['collaboration:check']).toBeUndefined()
    expect(scripts['automation:check']).toBeUndefined()
    expect(scripts['security:posture:check']).toBeUndefined()
    expect(packageJson).toContain('"xlsx-import:memory-gate:synthetic": "bun scripts/xlsx-import-memory-gate.ts --synthetic-only"')
    expect(packageJson).not.toContain('"research:xlsx-fixture-corpus:')
    expect(packageJson).not.toContain('"research:xlsx-fixture-corpus:memory-gate:synthetic"')
    expect(scripts[['test', 'correctness', 'corpus'].join(':')]).toBeUndefined()
    expect(scripts['test:correctness:xlsx']).toContain('packages/excel-fixtures/src/__tests__/oracle-harness.test.ts')
    expect(scripts['test:correctness:xlsx']).not.toContain('scripts/__tests__/xlsx-fixture-corpus')
    expect(scripts['research:xlsx-fixture-corpus:test']).toBeUndefined()
    expect(scripts[['ui', 'same-corpus', 'fixture', 'generate'].join(':')]).toBeUndefined()
    expect(scripts[['ui', 'same-corpus', 'fixture', 'check'].join(':')]).toBeUndefined()
    expect(scripts['research:ui-same-corpus:fixture:generate']).toBeUndefined()
    expect(scripts['research:ui-same-corpus:fixture:check']).toBeUndefined()
    expect(scripts['workpaper:bench:truecalc:check']).toBeUndefined()
    expect(scripts['workpaper:bench:xlsx-calc:check']).toBeUndefined()
    expect(scripts['workpaper:bench:ironcalc-rust:check']).toBeUndefined()
    expect(scripts['workpaper:bench:univer:check']).toBeUndefined()
    expect(scripts[['research', 'workpaper', 'bench', 'truecalc', 'check'].join(':')]).toBeUndefined()
    expect(scripts[['research', 'workpaper', 'bench', 'xlsx-calc', 'check'].join(':')]).toBeUndefined()
    expect(scripts[['research', 'workpaper', 'bench', 'ironcalc-rust', 'check'].join(':')]).toBeUndefined()
    expect(scripts[['research', 'workpaper', 'bench', 'univer', 'check'].join(':')]).toBeUndefined()
  })

  it('guards broad pre-push lint through the same resource gate', () => {
    const packageJson = readFileSync(resolve(repoRoot, 'package.json'), 'utf8')
    const prePushSource = readFileSync(resolve(repoRoot, 'scripts/run-pre-push.ts'), 'utf8')

    expect(packageJson).toContain('"pre-commit": "corepack pnpm@10.32.1 hooks:pre-commit"')
    expect(packageJson).toContain('"pre-push": "corepack pnpm@10.32.1 hooks:pre-push"')
    expect(prePushSource).toContain("assertLocalCiResourceGuardAllowsRun(rootDir, process.env, { runLabel: 'pre-push lint' })")
    expect(prePushSource).toContain("await run('corepack', ['pnpm@10.32.1', 'lint'])")
  })

  it('uses the deterministic CI profile for the GitHub release gate', () => {
    const workflow = readFileSync(resolve(repoRoot, '.github/workflows/ci.yml'), 'utf8')

    expect(workflow).toContain('pnpm run ci:github')
  })
})
