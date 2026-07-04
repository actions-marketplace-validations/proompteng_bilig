import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import type { XlsxZipByteSource } from '@bilig/xlsx/zip-reader'
import {
  externalPivotCachesWarning,
  externalWorkbookReferencesWarning,
  importXlsx,
  macroExecutionDeclinedWarning,
  manualCalculationModeWarning,
  precisionAsDisplayedCalculationWarning,
  volatileFormulasWarning,
} from '../packages/excel-import/src/index.js'
import { importXlsxFromZipByteSource } from '../packages/excel-import/src/xlsx-byte-source-import.js'
import { detachImportedXlsxSourceBytes } from '../packages/excel-import/src/xlsx-source-bytes.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { buildXlsxFixtureCorpusContractReportFromCases } from './xlsx-fixture-corpus-contract-report.ts'
import {
  hasFormulaOracleCacheUnsupportedClassifications,
  hasImportWarningUnsupportedClassifications,
  hasPivotUnsupportedClassifications,
  hasResourceLimitUnsupportedClassifications,
  xlsxFixtureFormulaOracleCacheClassifierEvidence,
  xlsxFixtureImportWarningClassifierEvidence,
  xlsxFixturePivotClassifierEvidence,
  xlsxFixtureResourceLimitClassifierEvidence,
} from './xlsx-fixture-corpus-evidence.ts'
import { summarizeExternalWorkbookReferences, unsupportedWorkbookMetadataEvidence } from './xlsx-fixture-corpus-external-links.ts'
import { inspectWorkbookFootprintIsolated, type XlsxFixtureCorpusWorkerOptions } from './xlsx-fixture-corpus-footprint.ts'
import {
  classifyUnsupportedLocaleDecimalCommaFormulaOracle,
  localeDecimalCommaFormulaOracleUnsupportedClassification,
  type FormulaOracleMismatchDetail,
} from './xlsx-fixture-corpus-formula-oracle-classifiers.ts'
import { validateXlsxFixtureManifest } from './xlsx-fixture-corpus-json.ts'
import {
  shouldUseCompactLargeSimpleVerification,
  verifyLargeSimpleWorkbookCompact,
  verifyLargeSimpleWorkbookCompactPreflight,
} from './xlsx-fixture-corpus-large-simple-compact.ts'
import {
  formulaOracleFormulaCountResourceLimitPreflight,
  formulaOracleResourceLimitPreflight,
  importResourceLimitPreflight,
  roundTripResourceLimitPreflight,
  structuralSmokeResourceLimitPreflight,
  unsupportedPreflightResourceLimitCase,
  unsupportedResourceLimitCase,
  type ResourceLimitPreflight,
} from './xlsx-fixture-corpus-resource-limits.ts'
import type {
  BuildContractReportArgs,
  FormulaOracle,
  FormulaOracleValidationResult,
  XlsxFixtureArtifact,
  XlsxFixtureCaseStatus,
  XlsxFixtureCorpusCase,
  XlsxFixtureCorpusContractReport,
  XlsxFixtureFeatureCounts,
  XlsxFixtureValidationSummary,
} from './xlsx-fixture-corpus-types.ts'
import {
  startVerificationRuntimeMetrics,
  timeVerificationPhase,
  withVerificationRuntimeMetrics,
} from './xlsx-fixture-corpus-verification-metrics.ts'
import { artifactBaseEvidence, failedCase } from './xlsx-fixture-corpus-verify-cases.ts'
import { indexReusableXlsxFixtureCorpusCases } from './xlsx-fixture-corpus-verify-checkpoint.ts'
import { verifyCachedWorkbookArtifactIsolated } from './xlsx-fixture-corpus-verify-isolated.ts'
import {
  cellValuesMatchOracle,
  countImportedWorkbookFeatures,
  countWorkbookFeatures,
  extractFormulaOraclesFromXlsxByteSource,
  formatCellValue,
  importedWorkbookMetadata,
  inspectWorkbookFootprint,
  isUnsupportedCycleOracleMismatch,
  sha256HexSync,
} from './xlsx-fixture-corpus-workbook.ts'
import { FileBackedXlsxZipByteSource, sha256XlsxZipByteSourceHex } from './xlsx-fixture-corpus-xlsx-byte-source.ts'

export { verifyCachedWorkbookArtifactIsolated } from './xlsx-fixture-corpus-verify-isolated.ts'
export { localeDecimalCommaFormulaOracleUnsupportedClassification }

declare const Bun:
  | {
      gc(force?: boolean): void
    }
  | undefined

const xlsxFixtureCorpusFootprintWorkerScriptPath = fileURLToPath(new URL('./xlsx-fixture-corpus-footprint-worker.ts', import.meta.url))

export const defaultVerifyTimeoutMs = 180_000
export const defaultVerifyConcurrency = 1
export const defaultVerifyMaxRssBytes = 1536 * 1024 * 1024
export const defaultVerifyMaxCellCount = 1_500_000
export const isolatedFootprintByteThreshold = 1_000_000

const externalWorkbookRoundTripSkipEvidence =
  'Round-trip projection skipped because external workbook links are not recalculated during XLSX import.'
const macroRoundTripSkipEvidence = 'Round-trip projection skipped because macro execution is intentionally declined during XLSM import.'
export const rawPivotPartUnsupportedClassification = 'xlsx.pivots.rawPartNotSemanticallyImported'
export const externalPivotCacheUnsupportedClassification = 'xlsx.pivots.externalCacheNotSemanticallyImported'
export const staleFormulaCacheUnsupportedClassification = 'xlsx.xlsxFixtureCorpus.formulaOracleCache:independentRecalcMatched'
export const externalLinkTransitiveFormulaUnsupportedClassification = 'xlsx.externalLinks.transitiveFormulaDependenciesUnsupported'
export const nativeFormulaOracleUnavailableUnsupportedClassification = 'xlsx.xlsxFixtureCorpus.formulaOracle:nativeXmlCacheUnavailable'

interface DetailedFormulaOracleValidationResult extends FormulaOracleValidationResult {
  readonly mismatchDetails: readonly FormulaOracleMismatchDetail[]
  readonly skippedUnsupportedFormulaCount: number
  readonly resourceLimit?: ResourceLimitPreflight
  readonly unsupportedNativeFormulaOracle?: UnsupportedFormulaOracleCacheClassification
}

interface UnsupportedFormulaOracleCacheClassification {
  readonly unsupported: boolean
  readonly classifications: readonly string[]
  readonly evidence: readonly string[]
}

export async function buildXlsxFixtureCorpusContractReport(args: BuildContractReportArgs): Promise<XlsxFixtureCorpusContractReport> {
  validateXlsxFixtureManifest(args.manifest)
  const structuralSmokeSampleLimit = args.structuralSmokeSampleLimit ?? 50
  const verifyConcurrency = Math.max(1, Math.trunc(args.verifyConcurrency ?? defaultVerifyConcurrency))
  const verificationManifestPath = args.manifestPath
  const isolatedVerification = args.isolatedVerification === true && verificationManifestPath !== undefined
  const verifyTimeoutMs = Math.max(1, Math.trunc(args.verifyTimeoutMs ?? defaultVerifyTimeoutMs))
  const verifyMaxRssBytes = capVerifyMaxRssBytes(Math.max(1, Math.trunc(args.verifyMaxRssBytes ?? defaultVerifyMaxRssBytes)))
  const verifyMaxCellCount = Math.max(1, Math.trunc(args.verifyMaxCellCount ?? defaultVerifyMaxCellCount))
  const verifyRssCheckIntervalMs = Math.max(100, Math.trunc(args.verifyRssCheckIntervalMs ?? 250))
  const reusableCasesById = indexReusableXlsxFixtureCorpusCases({
    manifest: args.manifest,
    cases: args.reusableCases ?? [],
    structuralSmokeSampleLimit,
  })
  let completedCount = 0
  const reportVerifiedCase = (verifiedCase: XlsxFixtureCorpusCase): XlsxFixtureCorpusCase => {
    completedCount += 1
    args.onCaseVerified?.({
      completedCount,
      totalCount: args.manifest.artifacts.length,
      latestCase: verifiedCase,
    })
    return verifiedCase
  }
  const cases = await mapWithConcurrency(args.manifest.artifacts, verifyConcurrency, (artifact, index) => {
    const reusableCase = reusableCasesById.get(artifact.id)
    if (reusableCase) {
      return Promise.resolve(reportVerifiedCase(reusableCase))
    }
    const runStructuralSmoke = index < structuralSmokeSampleLimit
    const verifiedCasePromise =
      isolatedVerification && verificationManifestPath
        ? verifyCachedWorkbookArtifactIsolated({
            artifact,
            cacheDir: args.cacheDir,
            manifestPath: verificationManifestPath,
            runStructuralSmoke,
            timeoutMs: verifyTimeoutMs,
            maxRssBytes: verifyMaxRssBytes,
            maxCellCount: verifyMaxCellCount,
            rssCheckIntervalMs: verifyRssCheckIntervalMs,
          })
        : verifyCachedWorkbookArtifact(artifact, args.cacheDir, runStructuralSmoke, verifyMaxCellCount, {
            timeoutMs: verifyTimeoutMs,
            maxRssBytes: verifyMaxRssBytes,
            rssCheckIntervalMs: verifyRssCheckIntervalMs,
          })
    return verifiedCasePromise.then(reportVerifiedCase)
  })
  return buildXlsxFixtureCorpusContractReportFromCases({
    manifest: args.manifest,
    generatedAt: args.generatedAt,
    cases,
  })
}

export async function verifyCachedWorkbookArtifact(
  artifact: XlsxFixtureArtifact,
  cacheDir: string,
  runStructuralSmoke: boolean,
  maxCellCount: number,
  workerOptions: XlsxFixtureCorpusWorkerOptions,
): Promise<XlsxFixtureCorpusCase> {
  const runtimeMetrics = startVerificationRuntimeMetrics()
  const finishCase = (corpusCase: XlsxFixtureCorpusCase): XlsxFixtureCorpusCase =>
    withVerificationRuntimeMetrics(corpusCase, runtimeMetrics)
  const cachePath = join(cacheDir, artifact.cachePath)
  const baseEvidence = artifactBaseEvidence(artifact)
  if (!existsSync(cachePath)) {
    return finishCase(failedCase(artifact, 'error', baseEvidence, [`Missing cached workbook file: ${artifact.cachePath}`]))
  }
  const source = new FileBackedXlsxZipByteSource(cachePath)
  try {
    const actualHash = await timeVerificationPhase(runtimeMetrics, workerOptions, 'read-cache', () => sha256XlsxZipByteSourceHex(source))
    if (actualHash !== artifact.sha256) {
      return finishCase(
        failedCase(artifact, 'failed', baseEvidence, [
          `Cached workbook hash mismatch: expected ${artifact.sha256}, received ${actualHash}`,
        ]),
      )
    }
    const preflightCompactCase = verifyLargeSimpleWorkbookCompactPreflight({
      artifact,
      source,
      baseEvidence,
      classifyUnsupportedFeatures,
      maxCellCount,
      minByteLength: 0,
      runStructuralSmoke,
      runtimeMetrics,
      workerOptions,
    })
    if (preflightCompactCase) {
      return finishCase(preflightCompactCase)
    }
    const footprint = await timeVerificationPhase(runtimeMetrics, workerOptions, 'inspect-footprint', () =>
      source.byteLength >= isolatedFootprintByteThreshold
        ? inspectWorkbookFootprintIsolated({
            bytes: new Uint8Array(),
            filePath: cachePath,
            fileName: artifact.fileName,
            scriptPath: xlsxFixtureCorpusFootprintWorkerScriptPath,
            options: workerOptions,
          })
        : inspectWorkbookFootprint(readAllSourceBytes(source), artifact.fileName),
    )
    if (!footprint) {
      return finishCase(
        failedCase(artifact, 'error', baseEvidence, [
          'Workbook footprint subprocess did not return a valid footprint.',
          'The workbook was isolated in a subprocess so the corpus verification run could continue.',
        ]),
      )
    }
    if (footprint.featureCounts.cellCount > maxCellCount) {
      return finishCase(unsupportedResourceLimitCase(artifact, baseEvidence, footprint, maxCellCount))
    }
    const importResourceLimit = importResourceLimitPreflight(artifact, footprint)
    if (importResourceLimit) {
      return finishCase(unsupportedPreflightResourceLimitCase(artifact, baseEvidence, footprint, importResourceLimit))
    }
    collectGarbage()
    if (shouldUseCompactLargeSimpleVerification(artifact, footprint, runStructuralSmoke)) {
      const compactImportCase = await verifyLargeSimpleWorkbookCompact({
        artifact,
        source,
        footprint,
        baseEvidence,
        classifyUnsupportedFeatures,
        runStructuralSmoke,
        runtimeMetrics,
        workerOptions,
      })
      if (compactImportCase) {
        return finishCase(compactImportCase)
      }
    }
    const { imported, featureCounts, metadata } = await timeVerificationPhase(runtimeMetrics, workerOptions, 'import-xlsx', () => {
      const importedWorkbook = importXlsxFromZipByteSource(source, artifact.fileName, {
        attachSourceReaderForUntouchedExport: false,
      })
      const importedFeatureCounts = countImportedWorkbookFeatures(importedWorkbook)
      return {
        imported: importedWorkbook,
        featureCounts: mergeImportedAndFootprintFeatureCounts(importedFeatureCounts, footprint.featureCounts),
        metadata: importedWorkbookMetadata(importedWorkbook),
      }
    })
    const formulaOracleFormulaCountResourceLimit = formulaOracleFormulaCountResourceLimitPreflight({
      formulaCellCount: featureCounts.formulaCellCount,
    })
    const formulaOracleResourceLimit =
      formulaOracleFormulaCountResourceLimit ??
      (featureCounts.formulaCellCount === 0 ? null : formulaOracleResourceLimitPreflight(imported.snapshot))
    const formulaOracleBlockingWarning = hasFormulaOracleBlockingWarning(imported.warnings)
    const formulaOracleNeedsWorkbookBytes =
      !formulaOracleResourceLimit && footprint.featureCounts.formulaCellCount > 0 && !formulaOracleBlockingWarning
    if (!formulaOracleNeedsWorkbookBytes) {
      collectGarbage()
    }
    const {
      formulaOracleValidation,
      unsupportedFormulaOracleWarning,
      unsupportedFormulaOracleCache,
      unsupportedNativeFormulaOracle,
      unsupportedExternalLinkFormulaOracle,
      formulaOracleSourceResourceLimit,
    } = formulaOracleResourceLimit
      ? {
          formulaOracleValidation: emptyFormulaOracleValidation(),
          unsupportedFormulaOracleWarning: false,
          unsupportedFormulaOracleCache: emptyUnsupportedFormulaOracleCacheClassification(),
          unsupportedNativeFormulaOracle: emptyUnsupportedFormulaOracleCacheClassification(),
          unsupportedExternalLinkFormulaOracle: emptyUnsupportedFormulaOracleCacheClassification(),
          formulaOracleSourceResourceLimit: null,
        }
      : await timeVerificationPhase(runtimeMetrics, workerOptions, 'formula-oracle', async () => {
          const nextFormulaOracleValidation =
            footprint.featureCounts.formulaCellCount === 0 || formulaOracleBlockingWarning
              ? emptyFormulaOracleValidation()
              : await validateFormulaOracles(
                  imported.snapshot,
                  source,
                  artifact.fileName,
                  unsupportedFormulaDependencyKeys(imported.snapshot),
                )
          collectGarbage()
          const nextUnsupportedFormulaOracleWarning = hasUnsupportedPrecisionAsDisplayedOracleWarning(
            imported.warnings,
            nextFormulaOracleValidation,
          )
          const nextUnsupportedFormulaOracleCache = nextUnsupportedFormulaOracleWarning
            ? emptyUnsupportedFormulaOracleCacheClassification()
            : await classifyUnsupportedFormulaOracleCache(imported.snapshot, nextFormulaOracleValidation)
          const nextUnsupportedNativeFormulaOracle =
            nextFormulaOracleValidation.unsupportedNativeFormulaOracle ?? emptyUnsupportedFormulaOracleCacheClassification()
          const nextUnsupportedExternalLinkFormulaOracle = classifyUnsupportedExternalLinkFormulaOracle(
            imported.snapshot,
            nextFormulaOracleValidation,
          )
          return {
            formulaOracleValidation: nextFormulaOracleValidation,
            unsupportedFormulaOracleWarning: nextUnsupportedFormulaOracleWarning,
            unsupportedFormulaOracleCache: nextUnsupportedFormulaOracleCache,
            unsupportedNativeFormulaOracle: nextUnsupportedNativeFormulaOracle,
            unsupportedExternalLinkFormulaOracle: nextUnsupportedExternalLinkFormulaOracle,
            formulaOracleSourceResourceLimit: nextFormulaOracleValidation.resourceLimit ?? null,
          }
        })
    const unsupportedLocaleDecimalCommaFormulaOracle = classifyUnsupportedLocaleDecimalCommaFormulaOracle(
      imported.snapshot,
      formulaOracleValidation,
    )
    const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, featureCounts)
    const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(featureCounts) : null
    const phaseResourceLimitClassifications = [
      ...(formulaOracleResourceLimit ? [formulaOracleResourceLimit.classification] : []),
      ...(formulaOracleSourceResourceLimit ? [formulaOracleSourceResourceLimit.classification] : []),
      ...(roundTripResourceLimit ? [roundTripResourceLimit.classification] : []),
      ...(structuralSmokeResourceLimit ? [structuralSmokeResourceLimit.classification] : []),
    ]
    const phaseResourceLimitEvidence = [
      ...(formulaOracleResourceLimit?.evidence ?? []),
      ...(formulaOracleSourceResourceLimit?.evidence ?? []),
      ...(roundTripResourceLimit?.evidence ?? []),
      ...(structuralSmokeResourceLimit?.evidence ?? []),
    ]
    collectGarbage()
    const unsupportedFeatureClassifications = classifyUnsupportedFeatures(imported.snapshot, imported.warnings, featureCounts, {
      supportedImportWarnings: supportedFormulaOracleImportWarnings(imported.warnings, formulaOracleValidation),
      extraClassifications: [
        ...unsupportedFormulaOracleCache.classifications,
        ...unsupportedNativeFormulaOracle.classifications,
        ...unsupportedExternalLinkFormulaOracle.classifications,
        ...unsupportedLocaleDecimalCommaFormulaOracle.classifications,
        ...phaseResourceLimitClassifications,
      ],
    })
    const roundTripSkipEvidence = roundTripValidationSkipEvidence(imported.warnings)
    const externalWorkbookReferences = summarizeExternalWorkbookReferences(imported.snapshot)
    const unsupportedWorkbookEvidence = unsupportedWorkbookMetadataEvidence(imported.snapshot, formulaOracleValidation)
    const shouldRunStructuralSmoke = runStructuralSmoke && !structuralSmokeResourceLimit
    const roundTripValidation = await timeVerificationPhase(runtimeMetrics, workerOptions, 'round-trip', () =>
      roundTripSkipEvidence || roundTripResourceLimit
        ? { passed: true, structuralSmokeSnapshot: shouldRunStructuralSmoke ? imported.snapshot : undefined }
        : roundTripsSupportedSemantics(detachImportedWorkbookSnapshot(imported), {
            retainRoundTrippedSnapshot: shouldRunStructuralSmoke,
          }),
    )
    const roundTripPassed = roundTripValidation.passed
    collectGarbage()
    let structuralSmokeSnapshot = shouldRunStructuralSmoke ? roundTripValidation.structuralSmokeSnapshot : undefined
    const structuralSmokePassed = await timeVerificationPhase(runtimeMetrics, workerOptions, 'structural-smoke', () =>
      structuralSmokeSnapshot ? runStructuralSmokeOps(structuralSmokeSnapshot) : runStructuralSmoke ? null : null,
    )
    structuralSmokeSnapshot = undefined
    collectGarbage()
    const formulaOraclePassed =
      unsupportedFormulaOracleWarning ||
      unsupportedFormulaOracleCache.unsupported ||
      unsupportedNativeFormulaOracle.unsupported ||
      unsupportedExternalLinkFormulaOracle.unsupported ||
      unsupportedLocaleDecimalCommaFormulaOracle.unsupported ||
      formulaOracleValidation.mismatches.length === 0
    const validation: XlsxFixtureValidationSummary = {
      importPassed: true,
      formulaOraclePassed,
      formulaOracleComparisons: formulaOracleValidation.comparisons,
      formulaOracleMismatches: formulaOraclePassed ? [] : formulaOracleValidation.mismatches,
      roundTripPassed,
      structuralSmokePassed,
    }
    const passed =
      validation.importPassed &&
      validation.formulaOraclePassed &&
      validation.roundTripPassed &&
      (validation.structuralSmokePassed === null || validation.structuralSmokePassed)
    const status: XlsxFixtureCaseStatus = passed ? (unsupportedFeatureClassifications.length > 0 ? 'unsupported' : 'passed') : 'failed'
    return finishCase({
      id: artifact.id,
      sourceId: artifact.sourceId,
      sourceUrl: artifact.sourceUrl,
      fileName: artifact.fileName,
      sha256: artifact.sha256,
      byteSize: artifact.byteSize,
      license: artifact.license,
      status,
      passed,
      ...(externalWorkbookReferences ? { externalWorkbookReferences } : {}),
      featureCounts,
      workbookMetadata: metadata,
      validation,
      unsupportedFeatureClassifications,
      evidence: [
        ...baseEvidence,
        `sheets=${String(featureCounts.sheetCount)}`,
        `cells=${String(featureCounts.cellCount)}`,
        `formulas=${String(featureCounts.formulaCellCount)}`,
        ...(featureCounts.pivotCount > 0 ? [`pivots=${String(featureCounts.pivotCount)}`] : []),
        ...unsupportedWorkbookEvidence,
        ...(hasImportWarningUnsupportedClassifications(unsupportedFeatureClassifications)
          ? [xlsxFixtureImportWarningClassifierEvidence]
          : []),
        ...(hasPivotUnsupportedClassifications(unsupportedFeatureClassifications) ? [xlsxFixturePivotClassifierEvidence] : []),
        ...(hasResourceLimitUnsupportedClassifications(unsupportedFeatureClassifications)
          ? [xlsxFixtureResourceLimitClassifierEvidence, ...phaseResourceLimitEvidence]
          : []),
        ...(hasFormulaOracleCacheUnsupportedClassifications(unsupportedFeatureClassifications)
          ? [xlsxFixtureFormulaOracleCacheClassifierEvidence, ...unsupportedFormulaOracleCache.evidence]
          : []),
        ...unsupportedNativeFormulaOracle.evidence,
        ...unsupportedExternalLinkFormulaOracle.evidence,
        ...unsupportedLocaleDecimalCommaFormulaOracle.evidence,
        ...(roundTripSkipEvidence ? [roundTripSkipEvidence] : []),
        ...validationEvidence(validation),
      ],
    })
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    return finishCase(failedCase(artifact, 'error', baseEvidence, [message]))
  } finally {
    source.release()
  }
}

export function capVerifyMaxRssBytes(value: number): number {
  const normalizedValue = Math.max(1, Math.trunc(value))
  if (normalizedValue > defaultVerifyMaxRssBytes) {
    throw new Error(
      `XLSX fixture corpus verification RSS limits above ${String(Math.ceil(defaultVerifyMaxRssBytes / 1024 / 1024))} MiB are disabled because workbook workers can hang interactive hosts.`,
    )
  }
  return normalizedValue
}

async function validateFormulaOracles(
  snapshot: WorkbookSnapshot,
  source: XlsxZipByteSource,
  fileName: string,
  skippedFormulaKeys: ReadonlySet<string> = new Set(),
): Promise<DetailedFormulaOracleValidationResult> {
  try {
    const sourceFormulaOracles = extractFormulaOraclesFromXlsxByteSource(source, fileName)
    if (sourceFormulaOracles === null) {
      return {
        ...emptyFormulaOracleValidation(),
        unsupportedNativeFormulaOracle: {
          unsupported: true,
          classifications: [nativeFormulaOracleUnavailableUnsupportedClassification],
          evidence: [
            'formula-oracle-native-cache-reader=unsupported',
            'SheetJS formula oracle fallback is disabled for XLSX fixture corpus verification.',
          ],
        },
      }
    }
    const allFormulaOracles = sourceFormulaOracles
    const formulaOracles = allFormulaOracles.filter(
      (oracle) => !skippedFormulaKeys.has(formulaOracleCellKey(oracle.sheetName, oracle.address)),
    )
    const mismatchDetails = await compareFormulaOracles(snapshot, formulaOracles)
    return {
      comparisons: formulaOracles.length,
      mismatches: mismatchDetails.map((mismatch) => mismatch.message),
      mismatchDetails,
      skippedUnsupportedFormulaCount: allFormulaOracles.length - formulaOracles.length,
    }
  } catch (error) {
    return {
      comparisons: 0,
      mismatches: [`Formula oracle check failed: ${error instanceof Error ? error.message : String(error)}`],
      mismatchDetails: [],
      skippedUnsupportedFormulaCount: 0,
    }
  }
}

function emptyFormulaOracleValidation(): DetailedFormulaOracleValidationResult {
  return { comparisons: 0, mismatches: [], mismatchDetails: [], skippedUnsupportedFormulaCount: 0 }
}

function emptyUnsupportedFormulaOracleCacheClassification(): UnsupportedFormulaOracleCacheClassification {
  return { unsupported: false, classifications: [], evidence: [] }
}

async function classifyUnsupportedFormulaOracleCache(
  snapshot: WorkbookSnapshot,
  validation: DetailedFormulaOracleValidationResult,
): Promise<UnsupportedFormulaOracleCacheClassification> {
  void snapshot
  void validation
  return emptyUnsupportedFormulaOracleCacheClassification()
}

function classifyUnsupportedExternalLinkFormulaOracle(
  snapshot: WorkbookSnapshot,
  validation: DetailedFormulaOracleValidationResult,
): UnsupportedFormulaOracleCacheClassification {
  const unsupportedFormulaDependencies = snapshot.workbook.metadata?.unsupportedFormulaDependencies ?? []
  if (unsupportedFormulaDependencies.length === 0 || validation.mismatchDetails.length === 0) {
    return emptyUnsupportedFormulaOracleCacheClassification()
  }
  return {
    unsupported: true,
    classifications: [externalLinkTransitiveFormulaUnsupportedClassification],
    evidence: [
      `External-linked workbook formula oracle mismatches were classified as unsupported transitive linked-workbook dependencies: ${String(validation.mismatchDetails.length)} mismatches.`,
      ...validation.mismatchDetails
        .slice(0, 25)
        .map((mismatch) => `external-transitive-formula=${mismatch.sheetName}!${mismatch.address} ${mismatch.message}`),
    ],
  }
}

export function hasFormulaOracleBlockingWarning(warnings: readonly string[]): boolean {
  return warnings.some(
    (warning) =>
      warning === macroExecutionDeclinedWarning || warning === manualCalculationModeWarning || warning === volatileFormulasWarning,
  )
}

function hasUnsupportedPrecisionAsDisplayedOracleWarning(warnings: readonly string[], validation: FormulaOracleValidationResult): boolean {
  return warnings.includes(precisionAsDisplayedCalculationWarning) && validation.comparisons > 0 && validation.mismatches.length > 0
}

function supportedFormulaOracleImportWarnings(warnings: readonly string[], validation: FormulaOracleValidationResult): readonly string[] {
  return warnings.includes(precisionAsDisplayedCalculationWarning) && validation.comparisons > 0 && validation.mismatches.length === 0
    ? [precisionAsDisplayedCalculationWarning]
    : []
}

function roundTripValidationSkipEvidence(warnings: readonly string[]): string | null {
  if (warnings.includes(externalWorkbookReferencesWarning)) {
    return externalWorkbookRoundTripSkipEvidence
  }
  if (warnings.includes(macroExecutionDeclinedWarning)) {
    return macroRoundTripSkipEvidence
  }
  return null
}

function validationEvidence(validation: XlsxFixtureValidationSummary): string[] {
  const evidence: string[] = []
  if (!validation.formulaOraclePassed) {
    evidence.push(...validation.formulaOracleMismatches.slice(0, 25))
  }
  if (!validation.roundTripPassed) {
    evidence.push('Round-trip projection failed')
  }
  if (validation.structuralSmokePassed === false) {
    evidence.push('Structural smoke operations failed')
  }
  return evidence
}

async function compareFormulaOracles(
  snapshot: WorkbookSnapshot,
  oracles: readonly FormulaOracle[],
): Promise<FormulaOracleMismatchDetail[]> {
  if (oracles.length === 0) {
    return []
  }
  const { SpreadsheetEngine } = await import('../packages/core/src/engine.js')
  const engine = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: `xlsx-fixture-corpus-${stableId(snapshot.workbook.name)}`,
  })
  await engine.ready()
  engine.importSnapshot(snapshot)
  engine.recalculateNow()
  const mismatches: FormulaOracleMismatchDetail[] = []
  for (const oracle of oracles) {
    const actual = engine.getCellValue(oracle.sheetName, oracle.address)
    if (!cellValuesMatchOracle(actual, oracle.expected)) {
      const explanation = engine.explainCell(oracle.sheetName, oracle.address)
      if (isUnsupportedCycleOracleMismatch(actual, oracle.expected, explanation.inCycle)) {
        continue
      }
      const message = `${oracle.sheetName}!${oracle.address} expected ${formatCellValue(oracle.expected)} got ${formatCellValue(actual)}`
      mismatches.push({ sheetName: oracle.sheetName, address: oracle.address, expected: oracle.expected, actual, message })
    }
  }
  return mismatches
}

function formulaOracleCellKey(sheetName: string, address: string): string {
  return `${sheetName}\u0000${address}`
}

interface RoundTripSemanticValidationResult {
  readonly passed: boolean
  readonly structuralSmokeSnapshot?: WorkbookSnapshot
}

export async function roundTripsSupportedSemantics(
  snapshot: WorkbookSnapshot,
  options: { readonly retainRoundTrippedSnapshot?: boolean } = {},
): Promise<RoundTripSemanticValidationResult> {
  try {
    const [{ exportXlsx }, { roundTripSemanticsDigest }] = await Promise.all([
      import('../packages/excel-import/src/index.js'),
      import('./xlsx-fixture-corpus-roundtrip.ts'),
    ])
    const workbookName = snapshot.workbook.name
    const expectedDigest = roundTripSemanticsDigest(snapshot)
    detachImportedXlsxSourceBytes(snapshot)
    collectGarbage()
    const exported = exportXlsx(snapshot)
    snapshot = createDetachedWorkbookSnapshot(workbookName)
    collectGarbage()
    const roundTrippedSnapshot = importXlsx(exported, `${workbookName}.xlsx`).snapshot
    const actualDigest = roundTripSemanticsDigest(roundTrippedSnapshot)
    return {
      passed: actualDigest === expectedDigest,
      ...(options.retainRoundTrippedSnapshot ? { structuralSmokeSnapshot: roundTrippedSnapshot } : {}),
    }
  } catch {
    return { passed: false }
  }
}

function detachImportedWorkbookSnapshot(imported: ReturnType<typeof importXlsx>): WorkbookSnapshot {
  const snapshot = imported.snapshot
  imported.snapshot = createDetachedWorkbookSnapshot(snapshot.workbook.name)
  collectGarbage()
  return snapshot
}

function createDetachedWorkbookSnapshot(workbookName: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: workbookName },
    sheets: [],
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

function readAllSourceBytes(source: XlsxZipByteSource): Uint8Array {
  return source.readRange(0, source.byteLength)
}

async function runStructuralSmokeOps(snapshot: WorkbookSnapshot): Promise<boolean | null> {
  try {
    const sheetName = findStructuralSmokeSheetName(snapshot)
    if (!sheetName) {
      return null
    }
    const { SpreadsheetEngine } = await import('../packages/core/src/engine.js')
    const engine = new SpreadsheetEngine({ workbookName: `${snapshot.workbook.name}-structural-smoke` })
    engine.importSnapshot(cloneWorkbookSnapshotForStructuralSmoke(snapshot))
    engine.insertRows(sheetName, 0, 1)
    engine.deleteRows(sheetName, 0, 1)
    engine.recalculateNow()
    return true
  } catch {
    return false
  }
}

function findStructuralSmokeSheetName(snapshot: WorkbookSnapshot): string | null {
  const sheet = snapshot.sheets.find((entry) => !entry.metadata?.sheetProtection && (entry.metadata?.protectedRanges?.length ?? 0) === 0)
  return sheet?.name ?? null
}

export function cloneWorkbookSnapshotForStructuralSmoke(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  const workbookMetadata = cloneStructuralSmokeWorkbookMetadata(snapshot.workbook.metadata)
  return {
    version: snapshot.version,
    workbook: {
      name: snapshot.workbook.name,
      ...(workbookMetadata ? { metadata: workbookMetadata } : {}),
    },
    sheets: snapshot.sheets.map((sheet) => {
      const metadata = cloneStructuralSmokeSheetMetadata(sheet.metadata)
      return {
        ...(sheet.id !== undefined ? { id: sheet.id } : {}),
        name: sheet.name,
        order: sheet.order,
        ...(metadata ? { metadata } : {}),
        cells: sheet.cells.map((cell) => ({
          address: cell.address,
          ...(cell.row !== undefined ? { row: cell.row } : {}),
          ...(cell.col !== undefined ? { col: cell.col } : {}),
          ...(cell.value !== undefined ? { value: structuredClone(cell.value) } : {}),
          ...(cell.formula !== undefined ? { formula: cell.formula } : {}),
          ...(cell.format !== undefined ? { format: cell.format } : {}),
        })),
      }
    }),
  }
}

function cloneStructuralSmokeWorkbookMetadata(
  metadata: WorkbookSnapshot['workbook']['metadata'],
): WorkbookSnapshot['workbook']['metadata'] {
  if (!metadata) {
    return undefined
  }
  const cloned: NonNullable<WorkbookSnapshot['workbook']['metadata']> = {}
  for (const [key, value] of Object.entries(metadata)) {
    if (isPackageArtifactMetadataKey(key)) {
      continue
    }
    Reflect.set(cloned, key, structuredClone(value))
  }
  return Object.keys(cloned).length > 0 ? cloned : undefined
}

function cloneStructuralSmokeSheetMetadata(
  metadata: WorkbookSnapshot['sheets'][number]['metadata'],
): WorkbookSnapshot['sheets'][number]['metadata'] {
  if (!metadata) {
    return undefined
  }
  const cloned: NonNullable<WorkbookSnapshot['sheets'][number]['metadata']> = {}
  for (const [key, value] of Object.entries(metadata)) {
    if (isPackageArtifactMetadataKey(key)) {
      continue
    }
    Reflect.set(cloned, key, structuredClone(value))
  }
  return Object.keys(cloned).length > 0 ? cloned : undefined
}

function isPackageArtifactMetadataKey(key: string): boolean {
  return key.endsWith('Artifacts')
}

export function mergeImportedAndFootprintFeatureCounts(
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

export function classifyUnsupportedFeatures(
  snapshot: WorkbookSnapshot,
  warnings: readonly string[],
  featureCounts: XlsxFixtureFeatureCounts = countWorkbookFeatures(snapshot, warnings),
  options: { readonly supportedImportWarnings?: readonly string[]; readonly extraClassifications?: readonly string[] } = {},
): string[] {
  const classifications = new Set<string>()
  const supportedImportWarnings = new Set(options.supportedImportWarnings ?? [])
  if ((snapshot.workbook.metadata?.macroPayloads?.length ?? 0) > 0) {
    classifications.add('xlsx.macros.execution.declined')
  }
  if (featureCounts.pivotCount > (snapshot.workbook.metadata?.pivots?.length ?? 0)) {
    classifications.add(
      warnings.includes(externalPivotCachesWarning) ? externalPivotCacheUnsupportedClassification : rawPivotPartUnsupportedClassification,
    )
  }
  if ((snapshot.workbook.metadata?.unsupportedFormulaDependencies?.length ?? 0) > 0) {
    classifications.add('xlsx.externalLinks.formulaDependenciesUnsupported')
  }
  if ((snapshot.workbook.metadata?.externalWorkbookReferences?.length ?? 0) > 0) {
    classifications.add('xlsx.externalLinks.workbookReferencesPreserved')
  }
  if ((snapshot.workbook.metadata?.unsupportedPivots?.length ?? 0) > 0) {
    const hasExternalUnsupportedPivot = snapshot.workbook.metadata?.unsupportedPivots?.some((pivot) => pivot.kind === 'external-cache')
    classifications.add(hasExternalUnsupportedPivot ? externalPivotCacheUnsupportedClassification : rawPivotPartUnsupportedClassification)
  }
  for (const warning of warnings) {
    if (supportedImportWarnings.has(warning)) {
      continue
    }
    classifications.add(`xlsx.import.warning:${warning}`)
  }
  for (const classification of options.extraClassifications ?? []) {
    classifications.add(classification)
  }
  return [...classifications].toSorted()
}

function unsupportedFormulaDependencyKeys(snapshot: WorkbookSnapshot): ReadonlySet<string> {
  return new Set(
    (snapshot.workbook.metadata?.unsupportedFormulaDependencies ?? []).map((entry) => formulaOracleCellKey(entry.sheetName, entry.address)),
  )
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}

async function mapWithConcurrency<T, R>(
  items: readonly T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = []
  let nextIndex = 0
  const workerCount = Math.min(Math.max(1, concurrency), items.length)
  const runNext = async (): Promise<void> => {
    const index = nextIndex
    nextIndex += 1
    if (index >= items.length) {
      return
    }
    results[index] = await mapper(items[index], index)
    await runNext()
  }
  await Promise.all(Array.from({ length: workerCount }, () => runNext()))
  return results
}
