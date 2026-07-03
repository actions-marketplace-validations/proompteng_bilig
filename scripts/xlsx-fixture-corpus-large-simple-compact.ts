import {
  tryInspectLargeSimpleXlsxHeadless,
  type LargeSimpleXlsxHeadlessInspectResult,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource, type XlsxZipEntries } from '@bilig/xlsx/zip-reader'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { xlsxFixtureResourceLimitClassifierEvidence, hasResourceLimitUnsupportedClassifications } from './xlsx-fixture-corpus-evidence.ts'
import type { XlsxFixtureCorpusWorkerOptions } from './xlsx-fixture-corpus-footprint.ts'
import { largeSimpleImportPhaseTelemetryEvidence } from './xlsx-fixture-corpus-large-simple-evidence.ts'
import {
  formulaOracleFormulaCountResourceLimitPreflight,
  type ResourceLimitPreflight,
  roundTripResourceLimitPreflight,
  structuralSmokeResourceLimitPreflight,
  unsupportedResourceLimitCase,
} from './xlsx-fixture-corpus-resource-limits.ts'
import { timeVerificationPhase, type startVerificationRuntimeMetrics } from './xlsx-fixture-corpus-verification-metrics.ts'
import { isZipWorkbookSource } from './xlsx-fixture-corpus-xlsx-byte-source.ts'
import type {
  XlsxFixtureArtifact,
  XlsxFixtureCorpusCase,
  XlsxFixtureFeatureCounts,
  XlsxFixtureValidationSummary,
} from './xlsx-fixture-corpus-types.ts'
import type { WorkbookFootprint } from './xlsx-fixture-corpus-workbook.ts'

declare const Bun:
  | {
      gc(force?: boolean): void
    }
  | undefined

export type LargeSimpleUnsupportedFeatureClassifier = (
  snapshot: WorkbookSnapshot,
  warnings: readonly string[],
  featureCounts: XlsxFixtureFeatureCounts,
  options: { readonly extraClassifications?: readonly string[] },
) => string[]

export function shouldUseCompactLargeSimpleVerification(
  artifact: XlsxFixtureArtifact,
  footprint: WorkbookFootprint,
  runStructuralSmoke: boolean,
): boolean {
  return (
    footprint.largeSimpleXlsxImport?.eligible === true &&
    shouldUseCompactLargeSimpleFeatureCounts(artifact, footprint.featureCounts, runStructuralSmoke)
  )
}

export function verifyLargeSimpleWorkbookCompactPreflight(args: {
  readonly artifact: XlsxFixtureArtifact
  readonly source: XlsxZipByteSource
  readonly baseEvidence: readonly string[]
  readonly classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier
  readonly maxCellCount: number
  readonly minByteLength: number
  readonly runStructuralSmoke: boolean
  readonly runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>
  readonly workerOptions: XlsxFixtureCorpusWorkerOptions
}): XlsxFixtureCorpusCase | null {
  if (args.source.byteLength < args.minByteLength) {
    return null
  }
  const zip = readLargeSimpleVerifierZipEntries(args.source)
  if (!zip) {
    return null
  }
  args.workerOptions.onPhase?.('import-xlsx')
  const startedAt = performance.now()
  const inspected = tryInspectLargeSimpleHeadless({
    byteLength: args.source.byteLength,
    fileName: args.artifact.fileName,
    zip,
    options: {
      afterWorksheetScan: collectGarbage,
      minByteLength: 0,
      releaseZipSource: true,
    },
  })
  if (!inspected) {
    return null
  }
  const featureCounts = featureCountsFromLargeSimpleStats(inspected.stats)
  const footprint = footprintFromLargeSimpleInspect(inspected, featureCounts)
  const fullSnapshotResourceLimit =
    featureCounts.cellCount > args.maxCellCount ? largeSimpleFullSnapshotResourceLimit(featureCounts, args.maxCellCount) : null
  const recordImportTiming = (): void => {
    args.runtimeMetrics.phaseTimings.push({ phase: 'import-xlsx', elapsedMs: Math.max(0, Math.round(performance.now() - startedAt)) })
  }
  if (fullSnapshotResourceLimit && !shouldUseCompactLargeSimpleFeatureCounts(args.artifact, featureCounts, args.runStructuralSmoke)) {
    recordImportTiming()
    return unsupportedResourceLimitCase(args.artifact, args.baseEvidence, footprint, args.maxCellCount)
  }
  if (!shouldUseCompactLargeSimpleFeatureCounts(args.artifact, featureCounts, args.runStructuralSmoke)) {
    return null
  }
  recordImportTiming()
  return buildLargeSimpleCompactCase(
    args.artifact,
    inspected,
    featureCounts,
    args.baseEvidence,
    args.runStructuralSmoke,
    args.classifyUnsupportedFeatures,
    fullSnapshotResourceLimit ? [fullSnapshotResourceLimit] : [],
  )
}

export async function verifyLargeSimpleWorkbookCompact(args: {
  readonly artifact: XlsxFixtureArtifact
  readonly source: XlsxZipByteSource
  readonly footprint: WorkbookFootprint
  readonly baseEvidence: readonly string[]
  readonly classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier
  readonly runStructuralSmoke: boolean
  readonly runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>
  readonly workerOptions: XlsxFixtureCorpusWorkerOptions
}): Promise<XlsxFixtureCorpusCase | null> {
  const inspected = await timeVerificationPhase(args.runtimeMetrics, args.workerOptions, 'import-xlsx', () => {
    const zip = readLargeSimpleVerifierZipEntries(args.source)
    return zip
      ? tryInspectLargeSimpleHeadless({
          byteLength: args.source.byteLength,
          fileName: args.artifact.fileName,
          zip,
          options: {
            afterWorksheetScan: collectGarbage,
            minByteLength: 0,
            releaseZipSource: true,
          },
        })
      : null
  })
  if (!inspected) {
    return null
  }
  const featureCounts = mergeFeatureCounts(featureCountsFromLargeSimpleStats(inspected.stats), args.footprint.featureCounts)
  return buildLargeSimpleCompactCase(
    args.artifact,
    inspected,
    featureCounts,
    args.baseEvidence,
    args.runStructuralSmoke,
    args.classifyUnsupportedFeatures,
  )
}

function tryInspectLargeSimpleHeadless(args: {
  readonly byteLength: number
  readonly fileName: string
  readonly zip: XlsxZipEntries
  readonly options: Parameters<typeof tryInspectLargeSimpleXlsxHeadless>[3]
}): LargeSimpleXlsxHeadlessInspectResult | null {
  try {
    return tryInspectLargeSimpleXlsxHeadless({ byteLength: args.byteLength }, args.fileName, args.zip, args.options)
  } catch {
    return null
  }
}

function readLargeSimpleVerifierZipEntries(source: XlsxZipByteSource) {
  return isZipWorkbookSource(source) ? readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source)) : null
}

export function borrowXlsxZipByteSource(source: XlsxZipByteSource): XlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    readRange: (start, end) => source.readRange(start, end),
    ...(source.readRangeInto
      ? {
          readRangeInto: (start: number, end: number, target: Uint8Array) => source.readRangeInto!(start, end, target),
        }
      : {}),
    ...(source.inflateRawRange ? { inflateRawRange: (start: number, end: number) => source.inflateRawRange!(start, end) } : {}),
  }
}

function shouldUseCompactLargeSimpleFeatureCounts(
  artifact: XlsxFixtureArtifact,
  counts: XlsxFixtureFeatureCounts,
  runStructuralSmoke: boolean,
): boolean {
  const formulaOracleResourceLimit = formulaOracleFormulaCountResourceLimitPreflight(counts)
  const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, counts)
  const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(counts) : null
  return (
    counts.cellCount > 0 &&
    (counts.formulaCellCount === 0 || formulaOracleResourceLimit !== null) &&
    roundTripResourceLimit !== null &&
    (!runStructuralSmoke || structuralSmokeResourceLimit !== null)
  )
}

function buildLargeSimpleCompactCase(
  artifact: XlsxFixtureArtifact,
  inspected: LargeSimpleXlsxHeadlessInspectResult,
  featureCounts: XlsxFixtureFeatureCounts,
  baseEvidence: readonly string[],
  runStructuralSmoke: boolean,
  classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier,
  extraResourceLimits: readonly ResourceLimitPreflight[] = [],
): XlsxFixtureCorpusCase {
  const metadata: XlsxFixtureCorpusCase['workbookMetadata'] = {
    workbookName: inspected.workbookName,
    sheetNames: inspected.sheetNames,
    dimensions: inspected.stats.dimensions,
  }
  collectGarbage()
  const formulaOracleResourceLimit = formulaOracleFormulaCountResourceLimitPreflight(featureCounts)
  const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, featureCounts)
  const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(featureCounts) : null
  if (
    (featureCounts.formulaCellCount > 0 && !formulaOracleResourceLimit) ||
    !roundTripResourceLimit ||
    (runStructuralSmoke && !structuralSmokeResourceLimit)
  ) {
    throw new Error('Large-simple compact verification requires resource-skipped round-trip and structural phases.')
  }
  const phaseResourceLimitClassifications = [
    ...extraResourceLimits.map((entry) => entry.classification),
    ...(formulaOracleResourceLimit ? [formulaOracleResourceLimit.classification] : []),
    roundTripResourceLimit.classification,
    ...(structuralSmokeResourceLimit ? [structuralSmokeResourceLimit.classification] : []),
  ]
  const phaseResourceLimitEvidence = [
    ...extraResourceLimits.flatMap((entry) => entry.evidence),
    ...(formulaOracleResourceLimit?.evidence ?? []),
    ...roundTripResourceLimit.evidence,
    ...(structuralSmokeResourceLimit?.evidence ?? []),
  ]
  const unsupportedFeatureClassifications = classifyUnsupportedFeatures(
    minimalSnapshotFromLargeSimpleInspect(inspected),
    inspected.warnings,
    featureCounts,
    { extraClassifications: phaseResourceLimitClassifications },
  )
  const validation: XlsxFixtureValidationSummary = {
    importPassed: true,
    formulaOraclePassed: true,
    formulaOracleComparisons: 0,
    formulaOracleMismatches: [],
    roundTripPassed: true,
    structuralSmokePassed: null,
  }
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: unsupportedFeatureClassifications.length > 0 ? 'unsupported' : 'passed',
    passed: true,
    featureCounts,
    workbookMetadata: metadata,
    validation,
    unsupportedFeatureClassifications,
    evidence: [
      ...baseEvidence,
      `sheets=${String(featureCounts.sheetCount)}`,
      `cells=${String(featureCounts.cellCount)}`,
      `formulas=${String(featureCounts.formulaCellCount)}`,
      ...largeSimpleImportPhaseTelemetryEvidence(inspected.stats),
      ...(hasResourceLimitUnsupportedClassifications(unsupportedFeatureClassifications)
        ? [xlsxFixtureResourceLimitClassifierEvidence, ...phaseResourceLimitEvidence]
        : []),
    ],
  }
}

function largeSimpleFullSnapshotResourceLimit(featureCounts: XlsxFixtureFeatureCounts, maxCellCount: number): ResourceLimitPreflight {
  return {
    classification: `xlsx.xlsxFixtureCorpus.resourceLimit:cellCount>${String(maxCellCount)}`,
    evidence: [
      'rss-limit-phase=import-xlsx',
      `Full public snapshot materialization skipped because workbook has ${String(
        featureCounts.cellCount,
      )} cells, above verifier max ${String(maxCellCount)}; headless large-simple visitor import supplied counts and metadata without public cell arrays.`,
    ],
  }
}

function minimalSnapshotFromLargeSimpleInspect(inspected: LargeSimpleXlsxHeadlessInspectResult): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: inspected.workbookName },
    sheets: inspected.sheetNames.map((sheetName, order) => ({
      id: order + 1,
      name: sheetName,
      order,
      cells: [],
    })),
  }
}

function footprintFromLargeSimpleInspect(
  inspected: LargeSimpleXlsxHeadlessInspectResult,
  featureCounts: XlsxFixtureFeatureCounts,
): WorkbookFootprint {
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: inspected.workbookName,
      sheetNames: inspected.sheetNames,
      dimensions: inspected.stats.dimensions,
    },
    externalWorkbookReferences: [],
    largeSimpleXlsxImport: { eligible: true, blockers: [] },
  }
}

function featureCountsFromLargeSimpleStats(stats: LargeSimpleXlsxImportStats): XlsxFixtureFeatureCounts {
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: stats.warningCount,
  }
}

function mergeFeatureCounts(
  importedFeatureCounts: XlsxFixtureFeatureCounts,
  footprintFeatureCounts: XlsxFixtureFeatureCounts,
): XlsxFixtureFeatureCounts {
  return {
    sheetCount: Math.max(importedFeatureCounts.sheetCount, footprintFeatureCounts.sheetCount),
    cellCount: Math.max(importedFeatureCounts.cellCount, footprintFeatureCounts.cellCount),
    formulaCellCount: Math.max(importedFeatureCounts.formulaCellCount, footprintFeatureCounts.formulaCellCount),
    valueCellCount: Math.max(importedFeatureCounts.valueCellCount, footprintFeatureCounts.valueCellCount),
    definedNameCount: Math.max(importedFeatureCounts.definedNameCount, footprintFeatureCounts.definedNameCount),
    tableCount: Math.max(importedFeatureCounts.tableCount, footprintFeatureCounts.tableCount),
    chartCount: Math.max(importedFeatureCounts.chartCount, footprintFeatureCounts.chartCount),
    pivotCount: Math.max(importedFeatureCounts.pivotCount, footprintFeatureCounts.pivotCount),
    mergeCount: Math.max(importedFeatureCounts.mergeCount, footprintFeatureCounts.mergeCount),
    styleRangeCount: Math.max(importedFeatureCounts.styleRangeCount, footprintFeatureCounts.styleRangeCount),
    conditionalFormatCount: Math.max(importedFeatureCounts.conditionalFormatCount, footprintFeatureCounts.conditionalFormatCount),
    dataValidationCount: Math.max(importedFeatureCounts.dataValidationCount, footprintFeatureCounts.dataValidationCount),
    macroPayloadCount: Math.max(importedFeatureCounts.macroPayloadCount, footprintFeatureCounts.macroPayloadCount),
    warningCount: Math.max(importedFeatureCounts.warningCount, footprintFeatureCounts.warningCount),
  }
}

function collectGarbage(): void {
  if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') {
    Bun.gc(true)
    return
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}
