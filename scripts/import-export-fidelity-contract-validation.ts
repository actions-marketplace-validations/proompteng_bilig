import type { ImportExportFidelityCase, ImportExportFidelityContract } from './gen-import-export-fidelity-contract.ts'
import {
  buildImportExportSemanticLedger,
  importExportDeclinedRuntimeFeatures,
  importExportUnsupportedFeatures,
  type ImportExportSemanticDisposition,
  type ImportExportSemanticLedgerEntry,
} from './import-export-semantic-loss-ledger.ts'

const requiredCaseIds = [
  'csv-import-preview',
  'csv-engine-roundtrip',
  'xlsx-import-preview',
  'xlsx-snapshot-roundtrip-values-formulas-formats',
  'xlsx-snapshot-roundtrip-dimensions-merges',
  'xlsx-snapshot-roundtrip-freeze-panes',
  'xlsx-snapshot-roundtrip-filters',
  'xlsx-snapshot-roundtrip-sorts',
  'xlsx-snapshot-roundtrip-sheet-protection',
  'xlsx-snapshot-roundtrip-protected-ranges',
  'xlsx-snapshot-roundtrip-data-validations',
  'xlsx-snapshot-roundtrip-tables',
  'xlsx-snapshot-roundtrip-charts',
  'xlsx-snapshot-roundtrip-pivots',
  'xlsx-formula-context-audit',
  'xlsx-pivot-cache-semantics',
  'xlsx-external-data-provenance',
  'xlsx-macro-payload-preserved-without-execution',
  'xlsx-runtime-feature-policy-warning',
] as const

export function parseImportExportFidelityContract(value: unknown): ImportExportFidelityContract {
  const record = toRecord(value, 'import/export fidelity contract')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'import-export-fidelity') {
    throw new Error('Unexpected import/export fidelity contract header')
  }
  const source = recordField(record, 'source', 'import/export fidelity source')
  const summary = recordField(record, 'summary', 'import/export fidelity summary')
  return {
    schemaVersion: 1,
    suite: 'import-export-fidelity',
    generatedAt: stringField(record, 'generatedAt', 'import/export fidelity generatedAt'),
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-import-export-fidelity-contract.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/excel-import'),
      enginePackage: literalField(source, 'enginePackage', 'packages/core'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed', 'import/export fidelity allRequiredCasesPassed'),
      csvRoundTripPassed: booleanField(summary, 'csvRoundTripPassed', 'import/export fidelity csvRoundTripPassed'),
      xlsxImportPassed: booleanField(summary, 'xlsxImportPassed', 'import/export fidelity xlsxImportPassed'),
      xlsxSnapshotRoundTripPassed: booleanField(
        summary,
        'xlsxSnapshotRoundTripPassed',
        'import/export fidelity xlsxSnapshotRoundTripPassed',
      ),
      coveredFeatures: stringArrayField(summary, 'coveredFeatures', 'import/export fidelity coveredFeatures'),
      unsupportedFeatures: stringArrayField(summary, 'unsupportedFeatures', 'import/export fidelity unsupportedFeatures'),
      declinedRuntimeFeatures: stringArrayField(summary, 'declinedRuntimeFeatures', 'import/export fidelity declinedRuntimeFeatures'),
    },
    semanticLedger: arrayField(record, 'semanticLedger', 'import/export fidelity semanticLedger').map(parseSemanticLedgerEntry),
    cases: arrayField(record, 'cases', 'import/export fidelity cases').map(parseImportExportFidelityCase),
  }
}

export function validateImportExportFidelityContract(contract: ImportExportFidelityContract): void {
  for (const id of requiredCaseIds) {
    const entry = requiredCase(contract.cases, id)
    if (!entry.required) {
      throw new Error(`Import/export fidelity contract required case is not marked required: ${id}`)
    }
    if (!entry.passed) {
      throw new Error(`Import/export fidelity contract contains a failed required case: ${id}`)
    }
    if (entry.missingFeatures.length > 0) {
      throw new Error(`Import/export fidelity contract required case reports missing features: ${id}`)
    }
  }
  if (!contract.summary.allRequiredCasesPassed) {
    throw new Error('Import/export fidelity contract summary reports failed required cases')
  }
  if (!contract.summary.csvRoundTripPassed || !contract.summary.xlsxImportPassed || !contract.summary.xlsxSnapshotRoundTripPassed) {
    throw new Error('Import/export fidelity contract summary is missing required CSV/XLSX pass coverage')
  }
  validateSummaryCoveredFeatures(contract)
  validateSemanticLedger(contract)
  const unsupportedFeatures = importExportUnsupportedFeatures(contract.semanticLedger)
  for (const feature of unsupportedFeatures) {
    if (!contract.summary.unsupportedFeatures.includes(feature)) {
      throw new Error(`Import/export fidelity contract is missing unsupported feature disclosure: ${feature}`)
    }
  }
  if (contract.summary.unsupportedFeatures.length !== unsupportedFeatures.length) {
    throw new Error('Import/export fidelity contract reports unexpected unsupported import/export features')
  }
  const declinedRuntimeFeatures = importExportDeclinedRuntimeFeatures(contract.semanticLedger)
  for (const feature of declinedRuntimeFeatures) {
    if (!contract.summary.declinedRuntimeFeatures.includes(feature)) {
      throw new Error(`Import/export fidelity contract is missing declined runtime feature disclosure: ${feature}`)
    }
  }
}

function validateSemanticLedger(contract: ImportExportFidelityContract): void {
  const expectedLedger = buildImportExportSemanticLedger(contract.summary.coveredFeatures)
  if (JSON.stringify(contract.semanticLedger) !== JSON.stringify(expectedLedger)) {
    throw new Error('Import/export fidelity semantic ledger is stale against the current feature evidence')
  }
  const dispositions = new Set(contract.semanticLedger.map((entry) => entry.disposition))
  for (const requiredDisposition of ['preserved', 'declined-runtime'] satisfies ImportExportSemanticDisposition[]) {
    if (!dispositions.has(requiredDisposition)) {
      throw new Error(`Import/export fidelity semantic ledger is missing ${requiredDisposition} entries`)
    }
  }
}

function validateSummaryCoveredFeatures(contract: ImportExportFidelityContract): void {
  const caseFeatures = new Set(contract.cases.flatMap((entry) => entry.coveredFeatures))
  const summaryFeatures = new Set(contract.summary.coveredFeatures)
  for (const feature of caseFeatures) {
    if (!summaryFeatures.has(feature)) {
      throw new Error(`Import/export fidelity contract summary is missing covered feature: ${feature}`)
    }
  }
  for (const feature of summaryFeatures) {
    if (!caseFeatures.has(feature)) {
      throw new Error(`Import/export fidelity contract summary reports uncovered feature: ${feature}`)
    }
  }
}

function requiredCase(cases: readonly ImportExportFidelityCase[], id: string): ImportExportFidelityCase {
  const entry = cases.find((candidate) => candidate.id === id)
  if (!entry) {
    throw new Error(`Import/export fidelity contract is missing required case: ${id}`)
  }
  return entry
}

function parseImportExportFidelityCase(value: unknown): ImportExportFidelityCase {
  const record = toRecord(value, 'import/export fidelity case')
  return {
    id: stringField(record, 'id', 'import/export fidelity case id'),
    format: parseFormat(stringField(record, 'format', 'import/export fidelity case format')),
    direction: parseDirection(stringField(record, 'direction', 'import/export fidelity case direction')),
    required: booleanField(record, 'required', 'import/export fidelity case required'),
    passed: booleanField(record, 'passed', 'import/export fidelity case passed'),
    coveredFeatures: stringArrayField(record, 'coveredFeatures', 'import/export fidelity case coveredFeatures'),
    missingFeatures: stringArrayField(record, 'missingFeatures', 'import/export fidelity case missingFeatures'),
    evidence: stringField(record, 'evidence', 'import/export fidelity case evidence'),
  }
}

function parseSemanticLedgerEntry(value: unknown): ImportExportSemanticLedgerEntry {
  const record = toRecord(value, 'import/export fidelity semantic ledger entry')
  return {
    feature: stringField(record, 'feature', 'import/export fidelity semantic ledger feature'),
    disposition: parseSemanticDisposition(stringField(record, 'disposition', 'import/export fidelity semantic ledger disposition')),
    reason: stringField(record, 'reason', 'import/export fidelity semantic ledger reason'),
  }
}

function parseSemanticDisposition(value: string): ImportExportSemanticDisposition {
  if (value === 'preserved' || value === 'unsupported' || value === 'declined-runtime') {
    return value
  }
  throw new Error(`Unexpected import/export semantic ledger disposition: ${value}`)
}

function parseFormat(value: string): ImportExportFidelityCase['format'] {
  if (value === 'csv' || value === 'xlsx') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity format: ${value}`)
}

function parseDirection(value: string): ImportExportFidelityCase['direction'] {
  if (value === 'import' || value === 'export-import' || value === 'import-export-import') {
    return value
  }
  throw new Error(`Unexpected import/export fidelity direction: ${value}`)
}

function recordField(value: Record<string, unknown>, field: string, name: string): Record<string, unknown> {
  return toRecord(value[field], name)
}

function arrayField(value: Record<string, unknown>, field: string, name: string): unknown[] {
  const fieldValue = value[field]
  if (!Array.isArray(fieldValue)) {
    throw new Error(`Expected ${name} to be an array`)
  }
  return fieldValue
}

function stringArrayField(value: Record<string, unknown>, field: string, name: string): string[] {
  const fieldValue = arrayField(value, field, name)
  if (!fieldValue.every((entry) => typeof entry === 'string')) {
    throw new Error(`Expected ${name} to contain only strings`)
  }
  return fieldValue
}

function stringField(value: Record<string, unknown>, field: string, name: string): string {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'string') {
    throw new Error(`Expected ${name} to be a string`)
  }
  return fieldValue
}

function booleanField(value: Record<string, unknown>, field: string, name: string): boolean {
  const fieldValue = value[field]
  if (typeof fieldValue !== 'boolean') {
    throw new Error(`Expected ${name} to be a boolean`)
  }
  return fieldValue
}

function literalField<const T extends string>(value: Record<string, unknown>, field: string, expected: T): T {
  if (value[field] !== expected) {
    throw new Error(`Expected ${field} to be ${expected}`)
  }
  return expected
}

function toRecord(value: unknown, name: string): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${name} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}
