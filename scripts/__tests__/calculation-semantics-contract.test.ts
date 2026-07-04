import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildCalculationSemanticsContract, parseCalculationSemanticsContract } from '../gen-calculation-semantics-contract.ts'
import { readJsonObject } from '../json-contract-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('calculation semantics contract', () => {
  it('covers every committed canonical and workbook-semantics fixture', () => {
    const contract = buildCalculationSemanticsContract()

    expect(contract.summary.allCommittedFormulaSemanticsCovered).toBe(true)
    expect(contract.summary.canonicalFormulaFixtureCount).toBeGreaterThan(0)
    expect(contract.summary.coveredCanonicalFixtureCount).toBe(contract.summary.canonicalFormulaFixtureCount)
    expect(contract.summary.workbookSemanticsFixtureCount).toBe(12)
    expect(contract.summary.coveredWorkbookSemanticsFixtureCount).toBe(12)
    expect(contract.summary.coveredWorkbookSemanticsCategories).toEqual([
      'defined-names',
      'cross-sheet-references',
      'structured-references',
      'what-if-analysis',
      'dynamic-array-spills',
      'error-semantics',
    ])
    expect(contract.summary.missingCanonicalFixtureIds).toEqual([])
    expect(contract.summary.missingWorkbookSemanticsFixtureIds).toEqual([])
    expect(contract.summary.fixtureRegistryAligned).toBe(true)
    expect(contract.coverage.stableFormulaFixtureIds).toContain('lookup-reference:offset-basic')
    expect(contract.coverage.deterministicVolatileFixtureIds).toEqual([
      'date-time:now-volatile',
      'date-time:today-volatile',
      'volatile:rand-basic',
    ])
  })

  it('keeps the checked-in generated artifact aligned with the live fixture corpus', () => {
    const artifact = parseCalculationSemanticsContract(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/calculation-semantics-contract.json')),
    )
    const current = buildCalculationSemanticsContract(artifact.generatedAt)

    expect(artifact.summary).toEqual(current.summary)
    expect(artifact.coverage).toEqual(current.coverage)
  })
})
