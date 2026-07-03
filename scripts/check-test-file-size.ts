#!/usr/bin/env bun

import { readdir, readFile, stat } from 'node:fs/promises'
import path from 'node:path'

import { parseTestMaxLines } from './test-file-size-config.js'

interface TestSizeDebtEntry {
  readonly maxLines: number
  readonly path: string
  readonly reason: string
  readonly targetMaxLines: number
}

interface FileSizeRecord {
  readonly lineCount: number
  readonly relativePath: string
}

const root = process.cwd()
const maxLines = parseTestMaxLines(process.env['BILIG_TEST_MAX_LINES'])
const roots = ['apps', 'packages', 'scripts', 'e2e']
const ignoredDirNames = new Set(['node_modules', 'dist', 'build', 'generated', 'coverage'])
const testFilePatterns = [/\.test\.tsx?$/u, /\.pw\.ts$/u]
const debtFilePath = path.join(root, 'scripts/test-file-size-debt.json')

const debtEntries = await readDebtEntries()
const debtByPath = new Map(debtEntries.map((entry) => [entry.path, entry]))
const checkedFiles: FileSizeRecord[] = []
const violations: string[] = []

await Promise.all(roots.map(async (relativeRoot) => walk(path.join(root, relativeRoot))))

checkedFiles.sort((left, right) => right.lineCount - left.lineCount || left.relativePath.localeCompare(right.relativePath))

const checkedByPath = new Map(checkedFiles.map((record) => [record.relativePath, record]))

for (const [relativePath, entry] of debtByPath) {
  const checked = checkedByPath.get(relativePath)
  if (checked === undefined) {
    violations.push(`${relativePath}: stale debt entry points to a missing test file`)
    continue
  }
  if (checked.lineCount <= maxLines) {
    violations.push(
      `${relativePath}: stale debt entry is now ${String(checked.lineCount)} lines; remove it from scripts/test-file-size-debt.json`,
    )
    continue
  }
  if (checked.lineCount > entry.maxLines) {
    violations.push(`${relativePath}: ${String(checked.lineCount)} lines exceeds debt ceiling ${String(entry.maxLines)}`)
  }
}

for (const checked of checkedFiles) {
  if (checked.lineCount <= maxLines) {
    continue
  }
  if (debtByPath.has(checked.relativePath)) {
    continue
  }
  violations.push(
    `${checked.relativePath}: ${String(checked.lineCount)} lines exceeds ${String(maxLines)} and is not tracked as test-size debt`,
  )
}

if (violations.length > 0) {
  console.error(`Test file size check failed. Limit is ${String(maxLines)} lines unless tracked in scripts/test-file-size-debt.json.`)
  for (const violation of violations.toSorted()) {
    console.error(`- ${violation}`)
  }
  process.exit(1)
}

const largest = checkedFiles[0]
console.log(
  `Test file size check passed (${String(checkedFiles.length)} files, ${
    debtEntries.length
  } tracked oversized files, max ${largest ? `${largest.relativePath} at ${String(largest.lineCount)} lines` : '0 lines'}).`,
)

async function readDebtEntries(): Promise<TestSizeDebtEntry[]> {
  const parsed: unknown = JSON.parse(await readFile(debtFilePath, 'utf8'))
  if (!Array.isArray(parsed)) {
    throw new Error('scripts/test-file-size-debt.json must contain an array.')
  }

  const entries = parsed.map(parseDebtEntry)
  const sortedPaths = entries.map((entry) => entry.path).toSorted()
  const actualPaths = entries.map((entry) => entry.path)
  if (actualPaths.join('\n') !== sortedPaths.join('\n')) {
    throw new Error('scripts/test-file-size-debt.json must be sorted by path.')
  }

  const seen = new Set<string>()
  for (const entry of entries) {
    if (seen.has(entry.path)) {
      throw new Error(`Duplicate test-size debt entry: ${entry.path}`)
    }
    seen.add(entry.path)
  }

  return entries
}

function parseDebtEntry(value: unknown): TestSizeDebtEntry {
  if (!isUnknownRecord(value)) {
    throw new Error('Each test-size debt entry must be an object.')
  }

  const entry = {
    maxLines: parsePositiveInteger(value['maxLines'], 'maxLines'),
    path: parseRelativePath(value['path']),
    reason: parseNonEmptyString(value['reason'], 'reason'),
    targetMaxLines: parsePositiveInteger(value['targetMaxLines'], 'targetMaxLines'),
  }

  if (entry.targetMaxLines > maxLines) {
    throw new Error(`${entry.path}: targetMaxLines must be <= ${String(maxLines)}`)
  }
  if (entry.maxLines <= maxLines) {
    throw new Error(`${entry.path}: maxLines must be > ${String(maxLines)} for tracked test-size debt`)
  }

  return entry
}

function isUnknownRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function parsePositiveInteger(value: unknown, fieldName: string): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value <= 0) {
    throw new Error(`Test-size debt ${fieldName} must be a positive safe integer.`)
  }
  return value
}

function parseNonEmptyString(value: unknown, fieldName: string): string {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`Test-size debt ${fieldName} must be a non-empty string.`)
  }
  return value
}

function parseRelativePath(value: unknown): string {
  const relativePath = parseNonEmptyString(value, 'path')
  if (path.isAbsolute(relativePath) || relativePath.includes('\\') || relativePath.split('/').includes('..')) {
    throw new Error(`Test-size debt path must be a repository-relative POSIX path, got ${relativePath}`)
  }
  return relativePath
}

async function walk(currentPath: string): Promise<void> {
  const currentStat = await stat(currentPath)
  if (currentStat.isDirectory()) {
    if (ignoredDirNames.has(path.basename(currentPath))) {
      return
    }
    const entries = await readdir(currentPath, { withFileTypes: true })
    await Promise.all(entries.map(async (entry) => walk(path.join(currentPath, entry.name))))
    return
  }

  const relativePath = path.relative(root, currentPath)
  if (!testFilePatterns.some((pattern) => pattern.test(relativePath))) {
    return
  }

  checkedFiles.push({
    lineCount: countLines(await readFile(currentPath, 'utf8')),
    relativePath,
  })
}

function countLines(content: string): number {
  if (content.length === 0) {
    return 0
  }
  return content.split(/\r\n|\r|\n/u).length
}
