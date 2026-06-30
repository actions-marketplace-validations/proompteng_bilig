#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import {
  compatibilityFamilies,
  formulaCompatibilityRegistry,
  type CompatibilityFamily,
  type CompatibilityStatus,
  type FormulaCompatibilityEntry,
} from '../packages/formula/src/compatibility.ts'
import { formulaInventory, formulaInventorySummary } from '../packages/formula/src/generated/formula-inventory.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

interface FormulaCompatibilitySnapshot {
  schemaVersion: 1
  formulaBreadth: {
    officeListed: FormulaCompatibilityRatio
    tracked: FormulaCompatibilityRatio
    missingOfficeFunctions: string[]
  }
  canonical: {
    nonProductionRows: FormulaCompatibilityRow[]
    statusCounts: Record<CompatibilityStatus, number>
    summary: FormulaCompatibilityRatio
  }
  families: FormulaCompatibilityFamily[]
  strategicFamilies: FormulaCompatibilityFamily[]
}

interface FormulaCompatibilityRatio {
  percent: number
  production: number
  total: number
}

interface FormulaCompatibilityRow {
  family: CompatibilityFamily
  formula: string
  id: string
  notes?: string
  status: CompatibilityStatus
  wasmStatus: FormulaCompatibilityEntry['wasmStatus']
}

interface FormulaCompatibilityFamily {
  family: CompatibilityFamily
  nonProductionRows: FormulaCompatibilityRow[]
  statusCounts: Record<CompatibilityStatus, number>
  summary: FormulaCompatibilityRatio
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'formula', 'src', '__tests__', 'fixtures', 'formula-compatibility-snapshot.json')
const isCheckMode = process.argv.includes('--check')

const snapshot = buildSnapshot()
const serializedSnapshot = formatJsonForRepo(`${JSON.stringify(snapshot, null, 2)}\n`)

if (isCheckMode) {
  if (!existsSync(outputPath)) {
    throw new Error(
      `Missing generated formula compatibility snapshot at ${outputPath}. Run: bun scripts/gen-formula-compatibility-snapshot.ts`,
    )
  }

  const currentSnapshot = readFileSync(outputPath, 'utf8')
  if (currentSnapshot !== serializedSnapshot) {
    throw new Error(`Generated formula compatibility snapshot is out of date. Run: bun scripts/gen-formula-compatibility-snapshot.ts`)
  }
} else {
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, serializedSnapshot)
}

console.log(
  JSON.stringify(
    {
      outputPath,
      mode: isCheckMode ? 'check' : 'write',
      canonicalRows: snapshot.canonical.summary.total,
      canonicalProductionRows: snapshot.canonical.summary.production,
      openCanonicalRows: snapshot.canonical.nonProductionRows.map((row) => row.id),
      officeBreadth: snapshot.formulaBreadth.officeListed,
      trackedBreadth: snapshot.formulaBreadth.tracked,
    },
    null,
    2,
  ),
)

function buildSnapshot(): FormulaCompatibilitySnapshot {
  const officeListed = formulaInventory.filter((entry) => entry.inOfficeList)
  const canonicalRows = formulaCompatibilityRegistry.filter((entry) => entry.scope === 'canonical')
  const canonicalProduction = canonicalRows.filter((entry) => entry.status === 'implemented-wasm-production')
  const canonicalNonProduction = canonicalRows.filter((entry) => entry.status !== 'implemented-wasm-production').map(toCompatibilityRow)

  return {
    schemaVersion: 1,
    formulaBreadth: {
      officeListed: ratio(formulaInventorySummary.officeListedRegisteredInCodebase, officeListed.length),
      tracked: ratio(formulaInventorySummary.registeredInCodebase, formulaInventorySummary.total),
      missingOfficeFunctions: officeListed.filter((entry) => !entry.registeredInCodebase).map((entry) => entry.name),
    },
    canonical: {
      summary: ratio(canonicalProduction.length, canonicalRows.length),
      statusCounts: countStatuses(canonicalRows),
      nonProductionRows: canonicalNonProduction,
    },
    families: compatibilityFamilies.map((family) => buildFamilySummary(family, canonicalRows)),
    strategicFamilies: [
      buildFamilySummary('dynamic-array', canonicalRows),
      buildFamilySummary('names', canonicalRows),
      buildFamilySummary('tables', canonicalRows),
      buildFamilySummary('structured-reference', canonicalRows),
      buildFamilySummary('lambda', canonicalRows),
    ],
  }
}

function buildFamilySummary(family: CompatibilityFamily, canonicalRows: readonly FormulaCompatibilityEntry[]): FormulaCompatibilityFamily {
  const familyRows = canonicalRows.filter((entry) => entry.family === family)
  const familyProduction = familyRows.filter((entry) => entry.status === 'implemented-wasm-production')

  return {
    family,
    summary: ratio(familyProduction.length, familyRows.length),
    statusCounts: countStatuses(familyRows),
    nonProductionRows: familyRows.filter((entry) => entry.status !== 'implemented-wasm-production').map(toCompatibilityRow),
  }
}

function countStatuses(entries: readonly FormulaCompatibilityEntry[]): Record<CompatibilityStatus, number> {
  const counts: Record<CompatibilityStatus, number> = {
    unsupported: 0,
    seeded: 0,
    'implemented-js': 0,
    'implemented-js-and-wasm-shadow': 0,
    'implemented-wasm-production': 0,
    blocked: 0,
  }

  for (const entry of entries) {
    counts[entry.status] += 1
  }

  return counts
}

function ratio(production: number, total: number): FormulaCompatibilityRatio {
  return {
    production,
    total,
    percent: total === 0 ? 0 : Number(((production / total) * 100).toFixed(1)),
  }
}

function toCompatibilityRow(entry: FormulaCompatibilityEntry): FormulaCompatibilityRow {
  return {
    id: entry.id,
    family: entry.family,
    formula: entry.formula,
    status: entry.status,
    wasmStatus: entry.wasmStatus,
    notes: entry.notes,
  }
}
