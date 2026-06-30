#!/usr/bin/env bun

import { readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { formatJsonForRepo } from './scorecard-format.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const outputPath = join(repoRoot, 'docs', 'public-evidence.json')
const checkMode = process.argv.includes('--check')

interface LaneScorecard {
  readonly lane: string
  readonly comparableCount: number
  readonly workpaperWins: number
  readonly hyperformulaWins: number
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly worstWorkpaperToHyperFormulaMeanRatio: number
  readonly worstMeanRatioWorkload: string
  readonly worstWorkpaperToHyperFormulaP95Ratio: number
  readonly worstP95RatioWorkload: string
}

interface PublicEvidence {
  readonly schemaVersion: 1
  readonly package: {
    readonly name: string
    readonly version: string
    readonly releasePleaseManifestVersion: string
    readonly mcpServerVersion: string
    readonly mcpPackageVersion: string
  }
  readonly workpaperVsHyperFormula: {
    readonly artifactPath: string
    readonly generatedAt: string
    readonly sampleCount: number
    readonly warmupCount: number
    readonly workpaperPackageVersion: string
    readonly hyperformulaVersion: string
    readonly hyperformulaCommit: string
    readonly overall: LaneScorecard
    readonly publicLane: LaneScorecard
    readonly holdout: LaneScorecard
    readonly meanAndP95WinCount: number
    readonly p95HoldoutCount: number
  }
}

function asRecord(value: unknown, context: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${context} must be an object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function readString(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value !== 'string' || value.length === 0) {
    throw new Error(`${context}.${key} must be a non-empty string`)
  }
  return value
}

function readNumber(record: Record<string, unknown>, key: string, context: string): number {
  const value = record[key]
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`${context}.${key} must be a finite number`)
  }
  return value
}

function readJsonRecord(path: string, context: string): Promise<Record<string, unknown>> {
  return readFile(path, 'utf8').then((content) => asRecord(JSON.parse(content) as unknown, context))
}

function readLane(value: unknown, context: string): LaneScorecard {
  const record = asRecord(value, context)
  return {
    lane: readString(record, 'lane', context),
    comparableCount: readNumber(record, 'comparableCount', context),
    workpaperWins: readNumber(record, 'workpaperWins', context),
    hyperformulaWins: readNumber(record, 'hyperformulaWins', context),
    directionalMeanRatioGeomean: readNumber(record, 'directionalMeanRatioGeomean', context),
    directionalP95RatioGeomean: readNumber(record, 'directionalP95RatioGeomean', context),
    worstWorkpaperToHyperFormulaMeanRatio: readNumber(record, 'worstWorkpaperToHyperFormulaMeanRatio', context),
    worstMeanRatioWorkload: readString(record, 'worstMeanRatioWorkload', context),
    worstWorkpaperToHyperFormulaP95Ratio: readNumber(record, 'worstWorkpaperToHyperFormulaP95Ratio', context),
    worstP95RatioWorkload: readString(record, 'worstP95RatioWorkload', context),
  }
}

function readExcludedFamilies(benchmarkArtifact: Record<string, unknown>): readonly string[] {
  const scorecard = asRecord(benchmarkArtifact['scorecard'], 'workpaper benchmark artifact.scorecard')
  const excludedFamilies = scorecard['excludedFamilies']
  if (!Array.isArray(excludedFamilies) || excludedFamilies.some((entry) => typeof entry !== 'string')) {
    throw new Error('workpaper benchmark artifact.scorecard.excludedFamilies must be an array of strings')
  }
  return [...excludedFamilies]
}

function isExcludedWorkload(workload: string, excludedFamilies: readonly string[]): boolean {
  return excludedFamilies.some((family) => workload.includes(family))
}

function countMeanAndP95WorkpaperWins(benchmarkArtifact: Record<string, unknown>): number {
  const results = benchmarkArtifact['results']
  if (!Array.isArray(results)) {
    throw new Error('workpaper benchmark artifact.results must be an array')
  }

  const excludedFamilies = readExcludedFamilies(benchmarkArtifact)
  let count = 0
  for (const [index, entry] of results.entries()) {
    const result = asRecord(entry, `workpaper benchmark artifact.results[${index.toString()}]`)
    if (result['comparable'] !== true) {
      continue
    }
    const workload = readString(result, 'workload', `workpaper benchmark artifact.results[${index.toString()}]`)
    if (isExcludedWorkload(workload, excludedFamilies)) {
      continue
    }

    const comparison = asRecord(result['comparison'], `workpaper benchmark artifact.results[${index.toString()}].comparison`)
    const meanRatio = readNumber(
      comparison,
      'workpaperToHyperFormulaMeanRatio',
      `workpaper benchmark artifact.results[${index.toString()}].comparison`,
    )
    const p95Ratio = readNumber(
      comparison,
      'workpaperToHyperFormulaP95Ratio',
      `workpaper benchmark artifact.results[${index.toString()}].comparison`,
    )
    const verificationEquivalent = comparison['verificationEquivalent']
    if (verificationEquivalent !== false && meanRatio < 1 && p95Ratio < 1) {
      count += 1
    }
  }

  return count
}

function headline(lane: Pick<LaneScorecard, 'workpaperWins' | 'comparableCount'>): string {
  return `${lane.workpaperWins.toString()}/${lane.comparableCount.toString()}`
}

function ratio3(value: number): string {
  return `${value.toFixed(3).replace(/0+$/u, '').replace(/\.$/u, '')}x`
}

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function requireIncludesIgnoringWhitespace(haystack: string, needle: string, context: string): void {
  if (!normalizeWhitespace(haystack).includes(normalizeWhitespace(needle))) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/gu, ' ').trim()
}

function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include ${needle}`)
  }
}

function requireBenchmarkTableRow(content: string, lane: string, comparableCount: number, workpaperWins: number, context: string): void {
  const pattern = new RegExp(
    `\\| ${lane}\\s+\\|\\s+\`${comparableCount.toString()}\`\\s+\\|\\s+\`${workpaperWins.toString()}\`\\s+\\|`,
    'u',
  )
  if (!pattern.test(content)) {
    throw new Error(
      `${context} is missing ${lane} benchmark row for ${comparableCount.toString()} comparable workloads and ${workpaperWins.toString()} WorkPaper wins`,
    )
  }
}

async function buildEvidence(): Promise<PublicEvidence> {
  const [packageManifest, serverManifest, workpaperPackageManifest, workpaperServerManifest, releasePleaseManifest, benchmarkArtifact] =
    await Promise.all([
      readJsonRecord(join(repoRoot, 'packages', 'headless', 'package.json'), 'packages/headless/package.json'),
      readJsonRecord(join(repoRoot, 'packages', 'headless', 'server.json'), 'packages/headless/server.json'),
      readJsonRecord(join(repoRoot, 'packages', 'workpaper', 'package.json'), 'packages/workpaper/package.json'),
      readJsonRecord(join(repoRoot, 'packages', 'workpaper', 'server.json'), 'packages/workpaper/server.json'),
      readJsonRecord(join(repoRoot, '.release-please-manifest.json'), '.release-please-manifest.json'),
      readJsonRecord(
        join(repoRoot, 'packages', 'benchmarks', 'baselines', 'workpaper-vs-hyperformula.json'),
        'workpaper benchmark artifact',
      ),
    ])

  const packageName = readString(packageManifest, 'name', 'packages/headless/package.json')
  const packageVersion = readString(packageManifest, 'version', 'packages/headless/package.json')
  const workpaperPackageName = readString(workpaperPackageManifest, 'name', 'packages/workpaper/package.json')
  const workpaperPackageVersion = readString(workpaperPackageManifest, 'version', 'packages/workpaper/package.json')
  const releasePleaseVersion = readString(releasePleaseManifest, 'packages/headless', '.release-please-manifest.json')
  const serverVersion = readString(serverManifest, 'version', 'packages/headless/server.json')
  const serverPackages = serverManifest['packages']
  if (!Array.isArray(serverPackages)) {
    throw new Error('packages/headless/server.json.packages must be an array')
  }
  const npmServerPackage = serverPackages
    .map((entry) => asRecord(entry, 'packages/headless/server.json.packages[]'))
    .find((entry) => entry['identifier'] === packageName)
  if (!npmServerPackage) {
    throw new Error(`packages/headless/server.json is missing npm package entry for ${packageName}`)
  }
  const mcpPackageVersion = readString(npmServerPackage, 'version', 'packages/headless/server.json package entry')

  const workpaperServerVersion = readString(workpaperServerManifest, 'version', 'packages/workpaper/server.json')
  const workpaperServerPackages = workpaperServerManifest['packages']
  if (!Array.isArray(workpaperServerPackages)) {
    throw new Error('packages/workpaper/server.json.packages must be an array')
  }
  const workpaperNpmServerPackage = workpaperServerPackages
    .map((entry) => asRecord(entry, 'packages/workpaper/server.json.packages[]'))
    .find((entry) => entry['identifier'] === workpaperPackageName)
  if (!workpaperNpmServerPackage) {
    throw new Error(`packages/workpaper/server.json is missing npm package entry for ${workpaperPackageName}`)
  }
  const workpaperMcpPackageVersion = readString(workpaperNpmServerPackage, 'version', 'packages/workpaper/server.json package entry')

  const benchmark = asRecord(benchmarkArtifact['benchmark'], 'workpaper benchmark artifact.benchmark')
  const engines = asRecord(benchmarkArtifact['engines'], 'workpaper benchmark artifact.engines')
  const workpaperEngine = asRecord(engines['workpaper'], 'workpaper benchmark artifact.engines.workpaper')
  const hyperformulaEngine = asRecord(engines['hyperformula'], 'workpaper benchmark artifact.engines.hyperformula')
  const scorecard = asRecord(benchmarkArtifact['scorecard'], 'workpaper benchmark artifact.scorecard')
  const scorecards = asRecord(scorecard['scorecards'], 'workpaper benchmark artifact.scorecard.scorecards')
  const overall = readLane(scorecards['overall'], 'workpaper benchmark artifact.scorecard.scorecards.overall')
  const meanAndP95WinCount = countMeanAndP95WorkpaperWins(benchmarkArtifact)

  const alignedVersionEntries = [
    ['.release-please-manifest.json', releasePleaseVersion],
    ['packages/headless/server.json', serverVersion],
    ['packages/headless/server.json package entry', mcpPackageVersion],
    ['packages/workpaper/package.json', workpaperPackageVersion],
    ['packages/workpaper/server.json', workpaperServerVersion],
    ['packages/workpaper/server.json package entry', workpaperMcpPackageVersion],
  ] as const
  for (const [context, version] of alignedVersionEntries) {
    if (version !== packageVersion) {
      throw new Error(`${context} version ${version} must match ${packageName}@${packageVersion}`)
    }
  }

  return {
    schemaVersion: 1,
    package: {
      name: packageName,
      version: packageVersion,
      releasePleaseManifestVersion: releasePleaseVersion,
      mcpServerVersion: serverVersion,
      mcpPackageVersion,
    },
    workpaperVsHyperFormula: {
      artifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      generatedAt: readString(benchmarkArtifact, 'generatedAt', 'workpaper benchmark artifact'),
      sampleCount: readNumber(benchmark, 'sampleCount', 'workpaper benchmark artifact.benchmark'),
      warmupCount: readNumber(benchmark, 'warmupCount', 'workpaper benchmark artifact.benchmark'),
      workpaperPackageVersion: readString(workpaperEngine, 'version', 'workpaper engine'),
      hyperformulaVersion: readString(hyperformulaEngine, 'version', 'hyperformula engine'),
      hyperformulaCommit: readString(hyperformulaEngine, 'commit', 'hyperformula engine'),
      overall,
      publicLane: readLane(scorecards['public'], 'workpaper benchmark artifact.scorecard.scorecards.public'),
      holdout: readLane(scorecards['holdout'], 'workpaper benchmark artifact.scorecard.scorecards.holdout'),
      meanAndP95WinCount,
      p95HoldoutCount: overall.comparableCount - meanAndP95WinCount,
    },
  }
}

async function assertPublicSurfaces(evidence: PublicEvidence): Promise<void> {
  const benchmark = evidence.workpaperVsHyperFormula
  const overall = benchmark.overall
  const publicLane = benchmark.publicLane
  const holdout = benchmark.holdout
  const meanHeadline = headline(overall)
  const publicHeadline = headline(publicLane)
  const holdoutHeadline = headline(holdout)
  const meanAndP95Headline = `${benchmark.meanAndP95WinCount.toString()}/${overall.comparableCount.toString()}`
  const p95Ratio = ratio3(overall.worstWorkpaperToHyperFormulaP95Ratio)
  const forbiddenPublicProofTokens = [
    ['headless-performance', 'leadership-scorecard.json'].join('-'),
    ['pnpm headless', 'performance:check'].join(':'),
    ['headless performance', 'leadership', 'scorecard'].join(' '),
    ['checked', 'leadership', 'scorecard'].join(' '),
    ['all-provider', 'scorecard'].join(' '),
    ['goal status', '`'].join(' '),
  ] as const
  const scannedPaths = [
    'package.json',
    'scripts/run-ci.ts',
    '.github/workflows/headless-package.yml',
    'README.md',
    'packages/headless/README.md',
    'docs/index.html',
    'docs/what-workpaper-benchmark-proves.md',
    'docs/headless-workpaper-benchmark-evidence.md',
    'docs/production-adoption-checklist-headless-workpaper.md',
    'docs/llms.txt',
  ] as const
  const scannedContents = await Promise.all(scannedPaths.map(async (path) => [path, await readFile(join(repoRoot, path), 'utf8')] as const))
  for (const [path, content] of scannedContents) {
    for (const token of forbiddenPublicProofTokens) {
      requireNotIncludes(content, token, path)
    }
  }

  const [readme, headlessReadme, index, benchmarkExplainer, benchmarkEvidence, hyperformulaAlternative, svgCard] = await Promise.all([
    readFile(join(repoRoot, 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'index.html'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'what-workpaper-benchmark-proves.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'headless-workpaper-benchmark-evidence.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'hyperformula-alternative-headless-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, 'docs', 'assets', 'workpaper-benchmark-card.svg'), 'utf8'),
  ])

  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
  ] as const) {
    requireIncludes(content, `[\`${meanHeadline}\` comparable WorkPaper mean wins]`, path)
    requireIncludes(content, `\`${overall.worstP95RatioWorkload}\``, path)
    requireIncludes(content, `\`${p95Ratio}\``, path)
  }

  requireIncludes(index, `<strong>${meanHeadline}</strong>`, 'docs/index.html')
  requireIncludes(
    index,
    `${overall.workpaperWins.toString()} of ${overall.comparableCount.toString()} comparable mean-latency rows`,
    'docs/index.html',
  )
  requireIncludesIgnoringWhitespace(
    index,
    `${overall.worstP95RatioWorkload} is the current worst p95 row: <code>${p95Ratio}</code>`,
    'docs/index.html',
  )
  requireIncludes(index, 'workpaper-vs-hyperformula.json', 'docs/index.html')
  requireIncludes(index, 'pnpm workpaper:bench:competitive:check', 'docs/index.html')

  for (const [path, content] of [
    ['docs/what-workpaper-benchmark-proves.md', benchmarkExplainer],
    ['docs/headless-workpaper-benchmark-evidence.md', benchmarkEvidence],
  ] as const) {
    requireIncludes(content, 'workpaper-vs-hyperformula.json', path)
    requireIncludes(content, `\`${meanHeadline}\` mean-latency wins`, path)
    requireBenchmarkTableRow(content, 'Overall', overall.comparableCount, overall.workpaperWins, path)
    requireBenchmarkTableRow(content, 'Public', publicLane.comparableCount, publicLane.workpaperWins, path)
    requireBenchmarkTableRow(content, 'Holdout', holdout.comparableCount, holdout.workpaperWins, path)
    requireIncludes(content, `generated at \`${benchmark.generatedAt}\``, path)
    requireIncludes(content, `\`${overall.directionalMeanRatioGeomean.toString()}\``, path)
    requireIncludes(content, `\`${overall.directionalP95RatioGeomean.toString()}\``, path)
    requireIncludes(content, `\`${meanAndP95Headline}\` workloads winning both`, path)
    requireIncludes(content, `\`${overall.worstP95RatioWorkload}\``, path)
    requireIncludes(content, `\`${overall.worstWorkpaperToHyperFormulaP95Ratio.toString()}\``, path)
  }

  requireIncludes(hyperformulaAlternative, `\`${meanHeadline}\` mean wins`, 'docs/hyperformula-alternative-headless-workpaper.md')
  requireIncludes(
    hyperformulaAlternative,
    `\`${publicHeadline}\` public-lane mean wins`,
    'docs/hyperformula-alternative-headless-workpaper.md',
  )
  requireIncludes(
    hyperformulaAlternative,
    `\`${holdoutHeadline}\` holdout-lane mean wins`,
    'docs/hyperformula-alternative-headless-workpaper.md',
  )

  requireIncludes(svgCard, `>${meanAndP95Headline}</text>`, 'docs/assets/workpaper-benchmark-card.svg')
  requireIncludes(svgCard, 'pnpm workpaper:bench:competitive:check', 'docs/assets/workpaper-benchmark-card.svg')
  requireIncludes(svgCard, `${overall.worstP95RatioWorkload} p95: ${p95Ratio}`, 'docs/assets/workpaper-benchmark-card.svg')
}

const evidence = await buildEvidence()
const rendered = formatJsonForRepo(`${JSON.stringify(evidence, null, 2)}\n`)

if (checkMode) {
  const current = await readFile(outputPath, 'utf8')
  if (current !== rendered) {
    throw new Error('docs/public-evidence.json is out of date. Run: pnpm public:evidence:generate')
  }
  await assertPublicSurfaces(evidence)
  console.log(
    JSON.stringify(
      {
        ok: true,
        outputPath,
        version: evidence.package.version,
        workpaperMeanWins: headline(evidence.workpaperVsHyperFormula.overall),
        workpaperMeanAndP95Wins: `${evidence.workpaperVsHyperFormula.meanAndP95WinCount.toString()}/${evidence.workpaperVsHyperFormula.overall.comparableCount.toString()}`,
      },
      null,
      2,
    ),
  )
} else {
  await writeFile(outputPath, rendered)
  console.log(
    JSON.stringify(
      {
        mode: 'write',
        outputPath,
        version: evidence.package.version,
        workpaperMeanWins: headline(evidence.workpaperVsHyperFormula.overall),
      },
      null,
      2,
    ),
  )
}
