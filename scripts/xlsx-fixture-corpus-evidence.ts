import type { XlsxFixtureCorpusCase } from './xlsx-fixture-corpus-types.ts'

export const xlsxFixtureImportWarningClassifierEvidence = 'import-warning-classifier=2026-05-08-full-precision-formula-aware'
export const xlsxFixturePivotClassifierEvidence = 'pivot-classifier=2026-05-08-external-cache-warning'
export const xlsxFixtureResourceLimitClassifierEvidence = 'resource-limit-classifier=2026-05-17-native-streaming-xlsx-footprint'
export const xlsxFixtureFormulaOracleCacheClassifierEvidence =
  'formula-oracle-cache-classifier=2026-05-12-independent-recalculation-cross-check'

export type XlsxFixtureCorpusEvidenceRefreshReason =
  | 'missing-used-range-evidence'
  | 'missing-import-warning-classifier-evidence'
  | 'missing-pivot-classifier-evidence'
  | 'missing-resource-limit-classifier-evidence'
  | 'missing-formula-oracle-cache-classifier-evidence'

export function xlsxFixtureCorpusCaseNeedsEvidenceRefresh(entry: XlsxFixtureCorpusCase): boolean {
  return xlsxFixtureCorpusCaseEvidenceRefreshReasons(entry).length > 0
}

export function xlsxFixtureCorpusCaseEvidenceRefreshReasons(
  entry: XlsxFixtureCorpusCase,
): readonly XlsxFixtureCorpusEvidenceRefreshReason[] {
  const reasons: XlsxFixtureCorpusEvidenceRefreshReason[] = []
  if (!hasXlsxFixtureCorpusUsedRangeEvidence(entry)) {
    reasons.push('missing-used-range-evidence')
  }
  if (hasImportWarningUnsupportedClassification(entry) && !hasCurrentImportWarningClassifierEvidence(entry)) {
    reasons.push('missing-import-warning-classifier-evidence')
  }
  if (hasPivotUnsupportedClassification(entry) && !hasCurrentPivotClassifierEvidence(entry)) {
    reasons.push('missing-pivot-classifier-evidence')
  }
  if (hasResourceLimitUnsupportedClassification(entry) && !hasCurrentResourceLimitClassifierEvidence(entry)) {
    reasons.push('missing-resource-limit-classifier-evidence')
  }
  if (hasFormulaOracleCacheUnsupportedClassification(entry) && !hasCurrentFormulaOracleCacheClassifierEvidence(entry)) {
    reasons.push('missing-formula-oracle-cache-classifier-evidence')
  }
  return reasons
}

export function hasXlsxFixtureCorpusUsedRangeEvidence(entry: XlsxFixtureCorpusCase): boolean {
  return entry.workbookMetadata.dimensions.every((dimension) => {
    if (!Object.hasOwn(dimension, 'usedRange')) {
      return false
    }
    const range = dimension.usedRange
    if (dimension.nonEmptyCellCount === 0) {
      return range === null
    }
    return (
      range !== null &&
      range !== undefined &&
      range.startRow >= 0 &&
      range.startColumn >= 0 &&
      range.endRow >= range.startRow &&
      range.endColumn >= range.startColumn &&
      dimension.rowCount >= range.endRow + 1 &&
      dimension.columnCount >= range.endColumn + 1
    )
  })
}

export function hasImportWarningUnsupportedClassification(entry: XlsxFixtureCorpusCase): boolean {
  return hasImportWarningUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

export function hasImportWarningUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.import.warning:'))
}

export function hasPivotUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.pivots.'))
}

export function hasResourceLimitUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.xlsxFixtureCorpus.resourceLimit:'))
}

export function hasFormulaOracleCacheUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.xlsxFixtureCorpus.formulaOracleCache:'))
}

function hasPivotUnsupportedClassification(entry: XlsxFixtureCorpusCase): boolean {
  return hasPivotUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasResourceLimitUnsupportedClassification(entry: XlsxFixtureCorpusCase): boolean {
  return hasResourceLimitUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasFormulaOracleCacheUnsupportedClassification(entry: XlsxFixtureCorpusCase): boolean {
  return hasFormulaOracleCacheUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasCurrentImportWarningClassifierEvidence(entry: XlsxFixtureCorpusCase): boolean {
  return entry.evidence.includes(xlsxFixtureImportWarningClassifierEvidence)
}

function hasCurrentPivotClassifierEvidence(entry: XlsxFixtureCorpusCase): boolean {
  return entry.evidence.includes(xlsxFixturePivotClassifierEvidence)
}

function hasCurrentResourceLimitClassifierEvidence(entry: XlsxFixtureCorpusCase): boolean {
  return entry.evidence.includes(xlsxFixtureResourceLimitClassifierEvidence)
}

function hasCurrentFormulaOracleCacheClassifierEvidence(entry: XlsxFixtureCorpusCase): boolean {
  return entry.evidence.includes(xlsxFixtureFormulaOracleCacheClassifierEvidence)
}
