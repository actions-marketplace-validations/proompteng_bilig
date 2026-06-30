import { readFileSync } from 'node:fs'
import { join } from 'node:path'

export interface OperatorWorkflowEvidence {
  readonly packageJsonPath: string
  readonly runCiPath: string
  readonly auditGeneratorPath: string
  readonly scorecardGeneratorPath: string
  readonly dominancePackageScriptsAbsent: boolean
  readonly googleSheetsTenXClaimGateScriptAbsent: boolean
  readonly publicClaimsCheckScriptPresent: boolean
  readonly runCiDominanceChecksAbsent: boolean
  readonly runCiPublicClaimsCheckPresent: boolean
  readonly generatedSourceChecksSerialized: boolean
  readonly promptArtifactAuditCoupledToLiveStatus: boolean
}

export function loadOperatorWorkflowEvidence(rootDir: string): OperatorWorkflowEvidence {
  const packageJsonPath = join(rootDir, 'package.json')
  const runCiPath = join(rootDir, 'scripts', 'run-ci.ts')
  const auditGeneratorPath = join(rootDir, 'scripts', 'bilig-dominance-audit.ts')
  const packageJson = parsePackageJson(readFileSync(packageJsonPath, 'utf8'))
  const runCiSource = readFileSync(runCiPath, 'utf8')
  const auditGeneratorSource = readFileSync(auditGeneratorPath, 'utf8')
  return {
    packageJsonPath: 'package.json',
    runCiPath: 'scripts/run-ci.ts',
    auditGeneratorPath: 'scripts/bilig-dominance-audit.ts',
    scorecardGeneratorPath: 'scripts/gen-bilig-dominance-scorecard.ts',
    dominancePackageScriptsAbsent: !Object.keys(packageJson.scripts).some((scriptName) => scriptName.startsWith('dominance:')),
    googleSheetsTenXClaimGateScriptAbsent: packageJson.scripts['google-sheets-10x:claim:check'] === undefined,
    publicClaimsCheckScriptPresent: packageJson.scripts['claims:check'] === 'bun scripts/check-public-claims.ts',
    runCiDominanceChecksAbsent: !runCiSource.includes(
      "bunScript('bilig dominance scorecard check', 'scripts/gen-bilig-dominance-scorecard.ts', '--check')",
    ),
    runCiPublicClaimsCheckPresent: runCiSource.includes("bunScript('public claims check', 'scripts/check-public-claims.ts')"),
    generatedSourceChecksSerialized: runCiSource.includes('Keep generated-source checks serialized'),
    promptArtifactAuditCoupledToLiveStatus:
      auditGeneratorSource.includes('buildBiligDominanceStatusFromArgs') &&
      auditGeneratorSource.includes('livePublicWorkbookCorpus') &&
      auditGeneratorSource.includes('validateBiligDominancePromptArtifactAudit'),
  }
}

export function operatorWorkflowGaps(evidence: OperatorWorkflowEvidence): string[] {
  return [
    ...(evidence.dominancePackageScriptsAbsent ? [] : ['package.json still exposes default dominance scripts']),
    ...(evidence.googleSheetsTenXClaimGateScriptAbsent ? [] : ['package.json still exposes the Google Sheets 10x claim gate script']),
    ...(evidence.publicClaimsCheckScriptPresent ? [] : ['package.json is missing the claims:check script']),
    ...(evidence.runCiDominanceChecksAbsent ? [] : ['run-ci still executes dominance scorecard checks']),
    ...(evidence.runCiPublicClaimsCheckPresent ? [] : ['run-ci does not execute claims:check']),
    ...(evidence.generatedSourceChecksSerialized ? [] : ['generated-source CI checks are not serialized']),
    ...(evidence.promptArtifactAuditCoupledToLiveStatus ? [] : ['prompt-to-artifact audit is not coupled to live dominance status']),
  ]
}

function parsePackageJson(source: string): { readonly scripts: Record<string, string> } {
  const parsed: unknown = JSON.parse(source)
  if (!isRecord(parsed)) {
    throw new Error('package.json must be an object')
  }
  const scripts = parsed['scripts']
  if (!isRecord(scripts)) {
    throw new Error('package.json scripts must be an object')
  }
  return {
    scripts: Object.fromEntries(Object.entries(scripts).filter((entry): entry is [string, string] => typeof entry[1] === 'string')),
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}
