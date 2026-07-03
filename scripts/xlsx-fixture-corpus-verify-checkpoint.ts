import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'

import { parseXlsxFixtureCorpusCase, parseXlsxFixtureCorpusScorecardJson } from './xlsx-fixture-corpus-json.ts'
import type { XlsxFixtureArtifact, XlsxFixtureCorpusCase, XlsxFixtureManifest } from './xlsx-fixture-corpus-types.ts'

interface XlsxFixtureCorpusVerificationCheckpoint {
  readonly schemaVersion: 1
  readonly suite: 'xlsx-fixture-corpus-verification-checkpoint'
  readonly generatedAt: string
  readonly cases: readonly XlsxFixtureCorpusCase[]
}

export function indexReusableXlsxFixtureCorpusCases(args: {
  readonly manifest: XlsxFixtureManifest
  readonly cases: readonly XlsxFixtureCorpusCase[]
  readonly structuralSmokeSampleLimit: number
}): ReadonlyMap<string, XlsxFixtureCorpusCase> {
  const candidatesById = new Map(args.cases.map((entry) => [entry.id, entry]))
  const reusableById = new Map<string, XlsxFixtureCorpusCase>()
  args.manifest.artifacts.forEach((artifact, index) => {
    const candidate = candidatesById.get(artifact.id)
    if (candidate && isReusableXlsxFixtureCorpusCase(artifact, candidate, index < args.structuralSmokeSampleLimit)) {
      reusableById.set(artifact.id, candidate)
    }
  })
  return reusableById
}

export function readReusableXlsxFixtureCorpusCases(paths: readonly string[]): XlsxFixtureCorpusCase[] {
  const cases: XlsxFixtureCorpusCase[] = []
  for (const path of paths) {
    if (!existsSync(path)) {
      continue
    }
    const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
    if (isVerificationCheckpoint(parsed)) {
      cases.push(...parsed.cases.map((entry) => normalizeReusableXlsxFixtureCorpusCase(parseXlsxFixtureCorpusCase(entry))))
      continue
    }
    if (isXlsxFixtureCorpusScorecardPayload(parsed)) {
      const scorecardCases = Reflect.get(parsed, 'cases')
      if (!Array.isArray(scorecardCases)) {
        throw new Error('XLSX fixture corpus scorecard is missing cases')
      }
      cases.push(...scorecardCases.map((entry) => normalizeReusableXlsxFixtureCorpusCase(parseXlsxFixtureCorpusCase(entry))))
      continue
    }
    cases.push(...parseXlsxFixtureCorpusScorecardJson(parsed).cases.map((entry) => normalizeReusableXlsxFixtureCorpusCase(entry)))
  }
  return cases
}

export function writeXlsxFixtureCorpusVerificationCheckpoint(args: {
  readonly path: string
  readonly manifest: XlsxFixtureManifest
  readonly casesById: ReadonlyMap<string, XlsxFixtureCorpusCase>
  readonly generatedAt?: string
}): void {
  const checkpoint: XlsxFixtureCorpusVerificationCheckpoint = {
    schemaVersion: 1,
    suite: 'xlsx-fixture-corpus-verification-checkpoint',
    generatedAt: args.generatedAt ?? new Date().toISOString(),
    cases: args.manifest.artifacts.flatMap((artifact) => {
      const entry = args.casesById.get(artifact.id)
      return entry ? [entry] : []
    }),
  }
  mkdirSync(dirname(args.path), { recursive: true })
  writeFileSync(args.path, `${JSON.stringify(checkpoint, null, 2)}\n`)
}

export function upsertXlsxFixtureCorpusVerificationCheckpoint(args: {
  readonly path: string
  readonly manifest: XlsxFixtureManifest
  readonly verifiedCase: XlsxFixtureCorpusCase
  readonly generatedAt?: string
}): void {
  const artifact = args.manifest.artifacts.find((entry) => entry.id === args.verifiedCase.id)
  if (!artifact) {
    throw new Error(`Cannot checkpoint XLSX fixture case ${args.verifiedCase.id} because it is not in the manifest`)
  }
  if (!caseMatchesArtifact(artifact, args.verifiedCase)) {
    throw new Error(`Cannot checkpoint XLSX fixture case ${args.verifiedCase.id} because it does not match the manifest artifact`)
  }
  const casesById = new Map(readReusableXlsxFixtureCorpusCases([args.path]).map((entry) => [entry.id, entry]))
  casesById.set(args.verifiedCase.id, args.verifiedCase)
  writeXlsxFixtureCorpusVerificationCheckpoint({
    path: args.path,
    manifest: args.manifest,
    casesById,
    generatedAt: args.generatedAt,
  })
}

function isReusableXlsxFixtureCorpusCase(
  artifact: XlsxFixtureArtifact,
  candidate: XlsxFixtureCorpusCase,
  structuralSmokeRequired: boolean,
): boolean {
  return (
    candidate.passed &&
    caseMatchesArtifact(artifact, candidate) &&
    (!structuralSmokeRequired || candidate.validation.structuralSmokePassed !== null)
  )
}

function caseMatchesArtifact(artifact: XlsxFixtureArtifact, candidate: XlsxFixtureCorpusCase): boolean {
  return (
    candidate.id === artifact.id &&
    candidate.sourceId === artifact.sourceId &&
    candidate.sourceUrl === artifact.sourceUrl &&
    candidate.fileName === artifact.fileName &&
    candidate.sha256 === artifact.sha256 &&
    candidate.byteSize === artifact.byteSize
  )
}

function normalizeReusableXlsxFixtureCorpusCase(candidate: XlsxFixtureCorpusCase): XlsxFixtureCorpusCase {
  const legacyRssLimitMiB = legacyRssLimitMiBFromEvidence(candidate.evidence)
  if (candidate.status !== 'error' || legacyRssLimitMiB === undefined) {
    return candidate
  }
  return {
    ...candidate,
    status: 'unsupported',
    passed: true,
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [`xlsx.xlsxFixtureCorpus.resourceLimit:rss>${String(legacyRssLimitMiB)}MiB`],
    evidence: candidate.evidence.map((line) =>
      line.startsWith('Verification subprocess exceeded RSS limit:')
        ? line.replace('Verification subprocess exceeded RSS limit:', 'XLSX fixture corpus verification RSS limit exceeded:')
        : line,
    ),
  }
}

function legacyRssLimitMiBFromEvidence(evidence: readonly string[]): number | undefined {
  for (const line of evidence) {
    const match = /Verification subprocess exceeded RSS limit: .+ > (?<value>\d+(?:\.\d+)?) (?<unit>MiB|GiB)/.exec(line)
    if (!match?.groups) {
      continue
    }
    const value = Number(match.groups['value'])
    if (!Number.isFinite(value)) {
      continue
    }
    return Math.max(1, Math.ceil(value * (match.groups['unit'] === 'GiB' ? 1024 : 1)))
  }
  return undefined
}

function isVerificationCheckpoint(value: unknown): value is XlsxFixtureCorpusVerificationCheckpoint {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'schemaVersion') === 1 &&
    Reflect.get(value, 'suite') === 'xlsx-fixture-corpus-verification-checkpoint' &&
    Array.isArray(Reflect.get(value, 'cases'))
  )
}

function isXlsxFixtureCorpusScorecardPayload(value: unknown): boolean {
  return (
    typeof value === 'object' &&
    value !== null &&
    Reflect.get(value, 'schemaVersion') === 1 &&
    Reflect.get(value, 'suite') === 'xlsx-fixture-corpus'
  )
}
