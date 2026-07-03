import type { CellValue } from '../packages/protocol/src/types.js'

export type XlsxFixtureSourceKind = 'direct-url' | 'ckan-resource' | 'github-contents'
export type XlsxFixtureCaseStatus = 'passed' | 'failed' | 'error' | 'unsupported'

export interface XlsxFixtureLicenseEvidence {
  readonly spdxId: string | null
  readonly title: string
  readonly evidenceUrl: string | null
}

export interface XlsxFixtureSource {
  readonly id: string
  readonly kind: XlsxFixtureSourceKind
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly discoveredAt: string
  readonly license: XlsxFixtureLicenseEvidence
  readonly topicEvidence?: readonly string[]
  readonly portal?: string
  readonly datasetId?: string
  readonly resourceId?: string
}

export interface XlsxFixtureArtifact {
  readonly id: string
  readonly sourceId: string
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
  readonly cachePath: string
  readonly sha256: string
  readonly byteSize: number
  readonly workbookFingerprint: string
  readonly fetchedAt: string
  readonly license: XlsxFixtureLicenseEvidence
  readonly topicEvidence?: readonly string[]
}

export interface XlsxFixtureManifest {
  readonly schemaVersion: 1
  readonly corpus: 'xlsx-fixture-corpus'
  readonly targetWorkbookCount: number
  readonly generatedAt: string
  readonly sources: readonly XlsxFixtureSource[]
  readonly artifacts: readonly XlsxFixtureArtifact[]
  readonly fetchState?: XlsxFixtureFetchState
}

export interface XlsxFixtureFetchState {
  readonly exhaustedSourceIds: readonly string[]
}

export interface XlsxFixtureFeatureCounts {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly chartCount: number
  readonly pivotCount: number
  readonly mergeCount: number
  readonly styleRangeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount: number
  readonly macroPayloadCount: number
  readonly warningCount: number
}

export type XlsxFixtureVerificationPhase =
  | 'read-cache'
  | 'inspect-footprint'
  | 'import-xlsx'
  | 'formula-oracle'
  | 'round-trip'
  | 'structural-smoke'

export interface XlsxFixtureVerificationPhaseTiming {
  readonly phase: XlsxFixtureVerificationPhase
  readonly elapsedMs: number
}

export interface XlsxFixtureValidationSummary {
  readonly importPassed: boolean
  readonly formulaOraclePassed: boolean
  readonly formulaOracleComparisons: number
  readonly formulaOracleMismatches: readonly string[]
  readonly roundTripPassed: boolean
  readonly structuralSmokePassed: boolean | null
}

export interface XlsxFixtureExternalReferenceSummary {
  readonly linkedWorkbookCount: number
  readonly formulaDependencyCount: number
  readonly cachedValueDependencyCount: number
}

export interface XlsxFixtureCorpusCase {
  readonly id: string
  readonly sourceId: string
  readonly sourceUrl: string
  readonly fileName: string
  readonly sha256: string
  readonly byteSize: number
  readonly license: XlsxFixtureLicenseEvidence
  readonly status: XlsxFixtureCaseStatus
  readonly passed: boolean
  readonly elapsedMs?: number
  readonly peakRssBytes?: number | null
  readonly phaseTimings?: readonly XlsxFixtureVerificationPhaseTiming[]
  readonly externalWorkbookReferences?: XlsxFixtureExternalReferenceSummary
  readonly featureCounts: XlsxFixtureFeatureCounts
  readonly workbookMetadata: {
    readonly workbookName: string
    readonly sheetNames: readonly string[]
    readonly dimensions: readonly {
      readonly sheetName: string
      readonly rowCount: number
      readonly columnCount: number
      readonly nonEmptyCellCount: number
      readonly usedRange?: {
        readonly startRow: number
        readonly startColumn: number
        readonly endRow: number
        readonly endColumn: number
      } | null
    }[]
  }
  readonly validation: XlsxFixtureValidationSummary
  readonly unsupportedFeatureClassifications: readonly string[]
  readonly evidence: readonly string[]
}

export interface XlsxFixtureCorpusScorecard {
  readonly schemaVersion: 1
  readonly suite: 'xlsx-fixture-corpus'
  readonly generatedAt: string
  readonly summary: {
    readonly targetWorkbookCount: number
    readonly sourceCount: number
    readonly cachedWorkbookCount: number
    readonly importedWorkbookCount: number
    readonly passedWorkbookCount: number
    readonly failedWorkbookCount: number
    readonly errorWorkbookCount: number
    readonly unsupportedWorkbookCount: number
    readonly formulaOracleComparisonCount: number
    readonly formulaOracleMatchCount: number
    readonly structuralSmokeRunCount: number
    readonly allCachedWorkbooksPassed: boolean
    readonly remainingToTarget: number
  }
  readonly cases: readonly XlsxFixtureCorpusCase[]
}

export interface FormulaOracle {
  readonly sheetName: string
  readonly address: string
  readonly expected: CellValue
}

export interface FormulaOracleValidationResult {
  readonly comparisons: number
  readonly mismatches: readonly string[]
}

export interface BuildScorecardArgs {
  readonly manifest: XlsxFixtureManifest
  readonly cacheDir: string
  readonly generatedAt?: string
  readonly manifestPath?: string
  readonly isolatedVerification?: boolean
  readonly structuralSmokeSampleLimit?: number
  readonly verifyConcurrency?: number
  readonly verifyTimeoutMs?: number
  readonly verifyMaxRssBytes?: number
  readonly verifyRssCheckIntervalMs?: number
  readonly verifyMaxCellCount?: number
  readonly reusableCases?: readonly XlsxFixtureCorpusCase[]
  readonly onCaseVerified?: (progress: XlsxFixtureCorpusVerificationProgress) => void
}

export interface XlsxFixtureCorpusVerificationProgress {
  readonly completedCount: number
  readonly totalCount: number
  readonly latestCase: XlsxFixtureCorpusCase
}

export interface DiscoverCkanArgs {
  readonly manifest: XlsxFixtureManifest
  readonly portalBases: readonly string[]
  readonly query: string
  readonly limit: number
  readonly rowsPerRequest: number
  readonly discoveredAt?: string
  readonly requiredTopic?: 'financial-workpapers' | 'recent-2025-2026-workbooks'
}

export interface FetchCorpusArgs {
  readonly manifest: XlsxFixtureManifest
  readonly cacheDir: string
  readonly limit: number
  readonly fetchedAt?: string
  readonly maxBytes?: number
  readonly downloadTimeoutMs?: number
  readonly fetchBatchSize?: number
  readonly fetchConcurrency?: number
  readonly fingerprintTimeoutMs?: number
  readonly fingerprintMaxRssBytes?: number
  readonly fingerprintRssCheckIntervalMs?: number
  readonly isolatedFingerprinting?: boolean
  readonly onArtifactsCommitted?: (
    manifest: XlsxFixtureManifest,
    progress: XlsxFixtureCorpusFetchCheckpointProgress,
  ) => void | Promise<void>
  readonly sourceIds?: readonly string[]
}

export interface XlsxFixtureCorpusFetchCheckpointProgress {
  readonly artifactCount: number
  readonly exhaustedSourceCount: number
  readonly committedArtifactCount: number
  readonly exhaustedSourceDelta: number
  readonly failedSourceCount: number
  readonly duplicateHashSourceCount: number
  readonly duplicateFingerprintSourceCount: number
  readonly failedSourceSamples: readonly XlsxFixtureCorpusFetchFailureSample[]
}

export interface XlsxFixtureCorpusFetchFailureSample {
  readonly sourceId: string
  readonly fileName: string
  readonly error: string
}

export interface CkanPageRequest {
  readonly portalBase: string
  readonly query: string
  readonly rowsPerRequest: number
  readonly start: number
}

export interface CkanPageResult {
  readonly portalBase: string
  readonly packages: readonly Record<string, unknown>[]
}

export interface WorkbookDownloadResult {
  readonly source: XlsxFixtureSource
  readonly bytes: Uint8Array | null
  readonly sha256: string | null
  readonly workbookFingerprint: string | null
  readonly error: string | null
  readonly retryableFailure?: boolean
}
