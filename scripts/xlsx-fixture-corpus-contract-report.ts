import type { XlsxFixtureCorpusCase, XlsxFixtureCorpusContractReport, XlsxFixtureManifest } from './xlsx-fixture-corpus-types.ts'

export function buildXlsxFixtureCorpusContractReportFromCases(args: {
  readonly manifest: XlsxFixtureManifest
  readonly cases: readonly XlsxFixtureCorpusCase[]
  readonly generatedAt?: string
}): XlsxFixtureCorpusContractReport {
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

export function validateXlsxFixtureCorpusContractReport(contractReport: XlsxFixtureCorpusContractReport): void {
  if (contractReport.schemaVersion !== 1 || contractReport.suite !== 'xlsx-fixture-corpus') {
    throw new Error('Unexpected XLSX fixture corpus contract report header')
  }
  if (!Number.isInteger(contractReport.summary.targetWorkbookCount) || contractReport.summary.targetWorkbookCount <= 0) {
    throw new Error('XLSX fixture corpus contract report has an invalid target workbook count')
  }
  if (contractReport.cases.length !== contractReport.summary.cachedWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report case count does not match cached workbook count')
  }
  if (
    contractReport.summary.remainingToTarget !==
    Math.max(0, contractReport.summary.targetWorkbookCount - contractReport.summary.cachedWorkbookCount)
  ) {
    throw new Error('XLSX fixture corpus contract report remaining target count is stale')
  }
  const passedWorkbookCount = contractReport.cases.filter((entry) => entry.status === 'passed').length
  const failedWorkbookCount = contractReport.cases.filter((entry) => entry.status === 'failed').length
  const errorWorkbookCount = contractReport.cases.filter((entry) => entry.status === 'error').length
  const unsupportedWorkbookCount = contractReport.cases.filter((entry) => entry.status === 'unsupported').length
  if (contractReport.summary.passedWorkbookCount !== passedWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report passed workbook count is stale')
  }
  if (contractReport.summary.failedWorkbookCount !== failedWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report failed workbook count is stale')
  }
  if (contractReport.summary.errorWorkbookCount !== errorWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report error workbook count is stale')
  }
  if (contractReport.summary.unsupportedWorkbookCount !== unsupportedWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report unsupported workbook count is stale')
  }
  const importedWorkbookCount = contractReport.cases.filter((entry) => entry.validation.importPassed).length
  if (contractReport.summary.importedWorkbookCount !== importedWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report imported workbook count is stale')
  }
  const formulaOracleComparisonCount = contractReport.cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0)
  if (contractReport.summary.formulaOracleComparisonCount !== formulaOracleComparisonCount) {
    throw new Error('XLSX fixture corpus contract report formula oracle comparison count is stale')
  }
  if (contractReport.summary.formulaOracleMatchCount !== countFormulaOracleMatches(contractReport.cases)) {
    throw new Error('XLSX fixture corpus contract report formula oracle match count is stale')
  }
  if (contractReport.summary.allCachedWorkbooksPassed !== contractReport.cases.every((entry) => entry.passed)) {
    throw new Error('XLSX fixture corpus contract report pass summary is stale')
  }
  if (!contractReport.summary.allCachedWorkbooksPassed) {
    throw new Error('XLSX fixture corpus contract report has cached workbooks that did not pass')
  }
}

export function validateXlsxFixtureCorpusContractReportManifestCoverage(args: {
  readonly contractReport: XlsxFixtureCorpusContractReport
  readonly manifest: XlsxFixtureManifest
}): void {
  if (args.contractReport.summary.targetWorkbookCount !== args.manifest.targetWorkbookCount) {
    throw new Error('XLSX fixture corpus contract report target count does not match the manifest')
  }
  if (args.contractReport.summary.sourceCount !== args.manifest.sources.length) {
    throw new Error('XLSX fixture corpus contract report source count does not match the manifest')
  }
  if (args.contractReport.summary.cachedWorkbookCount !== args.manifest.artifacts.length) {
    throw new Error('XLSX fixture corpus contract report cached workbook count does not match the manifest')
  }
  if (args.contractReport.cases.length !== args.manifest.artifacts.length) {
    throw new Error('XLSX fixture corpus contract report cases do not cover every manifest artifact')
  }

  args.manifest.artifacts.forEach((artifact, index) => {
    const corpusCase = args.contractReport.cases[index]
    if (
      !corpusCase ||
      corpusCase.id !== artifact.id ||
      corpusCase.sourceId !== artifact.sourceId ||
      corpusCase.sourceUrl !== artifact.sourceUrl ||
      corpusCase.sha256 !== artifact.sha256 ||
      corpusCase.byteSize !== artifact.byteSize
    ) {
      throw new Error(`XLSX fixture corpus contract report case ${artifact.id} does not match the manifest artifact`)
    }
  })
}

export function countFormulaOracleMatches(cases: readonly XlsxFixtureCorpusCase[]): number {
  return cases.reduce(
    (sum, entry) => sum + Math.max(0, entry.validation.formulaOracleComparisons - entry.validation.formulaOracleMismatches.length),
    0,
  )
}
