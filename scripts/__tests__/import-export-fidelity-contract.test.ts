import { describe, expect, it } from 'vitest'

import { buildImportExportFidelityContract, validateImportExportFidelityContract } from '../gen-import-export-fidelity-contract.ts'

describe('import/export fidelity contract', () => {
  it('generates a checked artifact from real CSV and XLSX import/export round trips', async () => {
    const contract = await buildImportExportFidelityContract('2026-05-06T08:00:00.000Z')

    expect(contract).toMatchObject({
      schemaVersion: 1,
      suite: 'import-export-fidelity',
      generatedAt: '2026-05-06T08:00:00.000Z',
      summary: {
        allRequiredCasesPassed: true,
        csvRoundTripPassed: true,
        xlsxImportPassed: true,
        xlsxSnapshotRoundTripPassed: true,
      },
    })
    expect(contract.cases.map((entry) => entry.id)).toEqual([
      'csv-import-preview',
      'csv-engine-roundtrip',
      'xlsx-import-preview',
      'xlsx-snapshot-roundtrip-values-formulas-formats',
      'xlsx-snapshot-roundtrip-dimensions-merges',
      'xlsx-snapshot-roundtrip-freeze-panes',
      'xlsx-snapshot-roundtrip-filters',
      'xlsx-snapshot-roundtrip-sorts',
      'xlsx-snapshot-roundtrip-sheet-protection',
      'xlsx-snapshot-roundtrip-protected-ranges',
      'xlsx-snapshot-roundtrip-data-validations',
      'xlsx-snapshot-roundtrip-tables',
      'xlsx-snapshot-roundtrip-charts',
      'xlsx-snapshot-roundtrip-pivots',
      'xlsx-formula-context-audit',
      'xlsx-pivot-cache-semantics',
      'xlsx-external-data-provenance',
      'xlsx-macro-payload-preserved-without-execution',
      'xlsx-runtime-feature-policy-warning',
    ])
    expect(contract.cases.every((entry) => entry.required && entry.passed)).toBe(true)
    expect(contract.cases.find((entry) => entry.id === 'xlsx-macro-payload-preserved-without-execution')).toMatchObject({
      format: 'xlsx',
      direction: 'import-export-import',
      coveredFeatures: [
        'xlsx.macros.detectedNoExecution',
        'xlsx.macros.payloadRoundtrip',
        'xlsx.macros.codeNameRoundtrip',
        'xlsx.runtimeFeaturePolicyWarnings',
      ],
      missingFeatures: [],
    })
    expect(contract.summary.coveredFeatures).toEqual([
      'csv.import',
      'csv.preview',
      'csv.export',
      'csv.roundtrip',
      'xlsx.import',
      'xlsx.preview',
      'xlsx.export',
      'xlsx.roundtrip',
      'xlsx.values',
      'xlsx.formulas',
      'xlsx.formulaAudit.context',
      'xlsx.formulaAudit.cacheStatus',
      'xlsx.numberFormats',
      'xlsx.workbookProperties',
      'xlsx.calculationSettings',
      'xlsx.calculationSettings.calcChainDiagnostics',
      'xlsx.definedNames',
      'xlsx.comments',
      'xlsx.styles',
      'xlsx.conditionalFormats.roundtrip',
      'xlsx.rowColumnDimensions',
      'xlsx.merges',
      'xlsx.freezePanes.roundtrip',
      'xlsx.filters.roundtrip',
      'xlsx.sorts.roundtrip',
      'xlsx.sheetProtection.roundtrip',
      'xlsx.protectedRanges.roundtrip',
      'xlsx.dataValidations.roundtrip',
      'xlsx.tables.roundtrip',
      'xlsx.charts.roundtrip',
      'xlsx.pivots.roundtrip',
      'xlsx.pivots.cacheSemantics',
      'xlsx.pivots.externalCacheOnlySemantics',
      'xlsx.multiSheet',
      'xlsx.macros.detectedNoExecution',
      'xlsx.macros.payloadRoundtrip',
      'xlsx.macros.codeNameRoundtrip',
      'xlsx.externalData.provenance',
      'xlsx.runtimeFeaturePolicyWarnings',
    ])
    expect(contract.summary.unsupportedFeatures).toEqual([])
    expect(contract.summary.declinedRuntimeFeatures).toEqual(['xlsx.macros.execution'])
    expect(contract.semanticLedger).toContainEqual({
      feature: 'xlsx.macros.execution',
      disposition: 'declined-runtime',
      reason: 'Bilig preserves macro payload metadata but intentionally never executes workbook macros.',
    })
    expect(contract.semanticLedger).toContainEqual({
      feature: 'xlsx.values',
      disposition: 'preserved',
      reason: 'Preserved by required import/export fidelity case evidence.',
    })
  })

  it('keeps unsupported and declined import/export semantics explicit', async () => {
    const contract = await buildImportExportFidelityContract('test-generated')

    expect(contract.summary.unsupportedFeatures).toEqual([])
    expect(contract.summary.declinedRuntimeFeatures).toEqual(['xlsx.macros.execution'])
    expect(new Set(contract.semanticLedger.map((entry) => entry.disposition))).toEqual(new Set(['preserved', 'declined-runtime']))
  })

  it('rejects stale artifacts missing required fidelity cases', async () => {
    const contract = await buildImportExportFidelityContract('2026-05-06T08:00:00.000Z')
    const staleContract = {
      ...contract,
      cases: contract.cases.filter((entry) => entry.id !== 'xlsx-snapshot-roundtrip-dimensions-merges'),
    }

    expect(() => validateImportExportFidelityContract(staleContract)).toThrow(
      'Import/export fidelity contract is missing required case: xlsx-snapshot-roundtrip-dimensions-merges',
    )
  })

  it('rejects artifacts whose summary feature coverage drifts from case evidence', async () => {
    const contract = await buildImportExportFidelityContract('2026-05-06T08:00:00.000Z')
    const missingFeatureContract = {
      ...contract,
      summary: {
        ...contract.summary,
        coveredFeatures: contract.summary.coveredFeatures.filter((feature) => feature !== 'xlsx.styles'),
      },
    }
    const extraFeatureContract = {
      ...contract,
      summary: {
        ...contract.summary,
        coveredFeatures: [...contract.summary.coveredFeatures, 'xlsx.unbackedClaim'],
      },
    }

    expect(() => validateImportExportFidelityContract(missingFeatureContract)).toThrow(
      'Import/export fidelity contract summary is missing covered feature: xlsx.styles',
    )
    expect(() => validateImportExportFidelityContract(extraFeatureContract)).toThrow(
      'Import/export fidelity contract summary reports uncovered feature: xlsx.unbackedClaim',
    )
  })
})
