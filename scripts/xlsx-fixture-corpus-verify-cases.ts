import { emptyFeatureCounts } from './xlsx-fixture-corpus-workbook.ts'
import type { XlsxFixtureArtifact, XlsxFixtureCorpusCase } from './xlsx-fixture-corpus-types.ts'

export function artifactBaseEvidence(artifact: XlsxFixtureArtifact): string[] {
  return [
    `source=${artifact.sourceUrl}`,
    `license=${artifact.license.title}`,
    `sha256=${artifact.sha256}`,
    ...(artifact.topicEvidence ?? []).map((entry) => `topic=${entry}`),
  ]
}

export function failedCase(
  artifact: XlsxFixtureArtifact,
  status: 'failed' | 'error',
  evidence: readonly string[],
  errors: readonly string[],
): XlsxFixtureCorpusCase {
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status,
    passed: false,
    featureCounts: emptyFeatureCounts(),
    workbookMetadata: { workbookName: artifact.fileName, sheetNames: [], dimensions: [] },
    validation: {
      importPassed: false,
      formulaOraclePassed: false,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: false,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [],
    evidence: [...evidence, ...errors],
  }
}
