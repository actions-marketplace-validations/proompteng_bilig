import type { XlsxFixtureCorpusCase, XlsxFixtureCorpusScorecard, XlsxFixtureManifest } from './xlsx-fixture-corpus-types.ts'

export function buildXlsxFixtureCorpusScorecardFromCases(args: {
  readonly manifest: XlsxFixtureManifest
  readonly cases: readonly XlsxFixtureCorpusCase[]
  readonly generatedAt?: string
}): XlsxFixtureCorpusScorecard {
  const passedWorkbookCount = args.cases.filter((entry) => entry.status === 'passed').length
  const failedWorkbookCount = args.cases.filter((entry) => entry.status === 'failed').length
  const errorWorkbookCount = args.cases.filter((entry) => entry.status === 'error').length
  const unsupportedWorkbookCount = args.cases.filter((entry) => entry.status === 'unsupported').length
  const formulaOracleComparisonCount = args.cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0)
  return {
    schemaVersion: 1,
    suite: 'xlsx-fixture-corpus',
    generatedAt: args.generatedAt ?? new Date().toISOString(),
    summary: {
      targetWorkbookCount: args.manifest.targetWorkbookCount,
      sourceCount: args.manifest.sources.length,
      cachedWorkbookCount: args.manifest.artifacts.length,
      importedWorkbookCount: args.cases.filter((entry) => entry.validation.importPassed).length,
      passedWorkbookCount,
      failedWorkbookCount,
      errorWorkbookCount,
      unsupportedWorkbookCount,
      formulaOracleComparisonCount,
      formulaOracleMatchCount: countFormulaOracleMatches(args.cases),
      structuralSmokeRunCount: args.cases.filter((entry) => entry.validation.structuralSmokePassed !== null).length,
      allCachedWorkbooksPassed: args.cases.every((entry) => entry.passed),
      remainingToTarget: Math.max(0, args.manifest.targetWorkbookCount - args.manifest.artifacts.length),
    },
    cases: args.cases,
  }
}

export function validateXlsxFixtureCorpusScorecard(scorecard: XlsxFixtureCorpusScorecard): void {
  if (scorecard.schemaVersion !== 1 || scorecard.suite !== 'xlsx-fixture-corpus') {
    throw new Error('Unexpected XLSX fixture corpus scorecard header')
  }
  if (!Number.isInteger(scorecard.summary.targetWorkbookCount) || scorecard.summary.targetWorkbookCount <= 0) {
    throw new Error('XLSX fixture corpus scorecard has an invalid target workbook count')
  }
  if (scorecard.cases.length !== scorecard.summary.cachedWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard case count does not match cached workbook count')
  }
  if (scorecard.summary.remainingToTarget !== Math.max(0, scorecard.summary.targetWorkbookCount - scorecard.summary.cachedWorkbookCount)) {
    throw new Error('XLSX fixture corpus scorecard remaining target count is stale')
  }
  const passedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'passed').length
  const failedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'failed').length
  const errorWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'error').length
  const unsupportedWorkbookCount = scorecard.cases.filter((entry) => entry.status === 'unsupported').length
  if (scorecard.summary.passedWorkbookCount !== passedWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard passed workbook count is stale')
  }
  if (scorecard.summary.failedWorkbookCount !== failedWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard failed workbook count is stale')
  }
  if (scorecard.summary.errorWorkbookCount !== errorWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard error workbook count is stale')
  }
  if (scorecard.summary.unsupportedWorkbookCount !== unsupportedWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard unsupported workbook count is stale')
  }
  const importedWorkbookCount = scorecard.cases.filter((entry) => entry.validation.importPassed).length
  if (scorecard.summary.importedWorkbookCount !== importedWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard imported workbook count is stale')
  }
  const formulaOracleComparisonCount = scorecard.cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0)
  if (scorecard.summary.formulaOracleComparisonCount !== formulaOracleComparisonCount) {
    throw new Error('XLSX fixture corpus scorecard formula oracle comparison count is stale')
  }
  if (scorecard.summary.formulaOracleMatchCount !== countFormulaOracleMatches(scorecard.cases)) {
    throw new Error('XLSX fixture corpus scorecard formula oracle match count is stale')
  }
  if (scorecard.summary.allCachedWorkbooksPassed !== scorecard.cases.every((entry) => entry.passed)) {
    throw new Error('XLSX fixture corpus scorecard pass summary is stale')
  }
  if (!scorecard.summary.allCachedWorkbooksPassed) {
    throw new Error('XLSX fixture corpus scorecard has cached workbooks that did not pass')
  }
}

export function validateXlsxFixtureCorpusScorecardManifestCoverage(args: {
  readonly scorecard: XlsxFixtureCorpusScorecard
  readonly manifest: XlsxFixtureManifest
}): void {
  if (args.scorecard.summary.targetWorkbookCount !== args.manifest.targetWorkbookCount) {
    throw new Error('XLSX fixture corpus scorecard target count does not match the manifest')
  }
  if (args.scorecard.summary.sourceCount !== args.manifest.sources.length) {
    throw new Error('XLSX fixture corpus scorecard source count does not match the manifest')
  }
  if (args.scorecard.summary.cachedWorkbookCount !== args.manifest.artifacts.length) {
    throw new Error('XLSX fixture corpus scorecard cached workbook count does not match the manifest')
  }
  if (args.scorecard.cases.length !== args.manifest.artifacts.length) {
    throw new Error('XLSX fixture corpus scorecard cases do not cover every manifest artifact')
  }

  args.manifest.artifacts.forEach((artifact, index) => {
    const corpusCase = args.scorecard.cases[index]
    if (
      !corpusCase ||
      corpusCase.id !== artifact.id ||
      corpusCase.sourceId !== artifact.sourceId ||
      corpusCase.sourceUrl !== artifact.sourceUrl ||
      corpusCase.sha256 !== artifact.sha256 ||
      corpusCase.byteSize !== artifact.byteSize
    ) {
      throw new Error(`XLSX fixture corpus scorecard case ${artifact.id} does not match the manifest artifact`)
    }
  })
}

export function countFormulaOracleMatches(cases: readonly XlsxFixtureCorpusCase[]): number {
  return cases.reduce(
    (sum, entry) => sum + Math.max(0, entry.validation.formulaOracleComparisons - entry.validation.formulaOracleMismatches.length),
    0,
  )
}
