#!/usr/bin/env bun

import { readdir, readFile, stat } from 'node:fs/promises'
import path from 'node:path'

import { parseTestMaxLines } from './test-file-size-config.js'

interface FileSizeRecord {
  readonly lineCount: number
  readonly relativePath: string
}

const root = process.cwd()
const maxLines = parseTestMaxLines(process.env['BILIG_TEST_MAX_LINES'])
const roots = ['apps', 'packages', 'scripts', 'e2e']
const ignoredDirNames = new Set(['node_modules', 'dist', 'build', 'generated', 'coverage'])
const testFilePatterns = [/\.test\.tsx?$/u, /\.pw\.ts$/u]

const checkedFiles: FileSizeRecord[] = []
const violations: string[] = []

await Promise.all(roots.map(async (relativeRoot) => walk(path.join(root, relativeRoot))))

checkedFiles.sort((left, right) => right.lineCount - left.lineCount || left.relativePath.localeCompare(right.relativePath))

for (const checked of checkedFiles) {
  if (checked.lineCount <= maxLines) {
    continue
  }
  violations.push(`${checked.relativePath}: ${String(checked.lineCount)} lines exceeds ${String(maxLines)}`)
}

if (violations.length > 0) {
  console.error(`Test file size check failed. Limit is ${String(maxLines)} lines.`)
  for (const violation of violations.toSorted()) {
    console.error(`- ${violation}`)
  }
  process.exit(1)
}

const largest = checkedFiles[0]
console.log(
  `Test file size check passed (${String(checkedFiles.length)} files, max ${
    largest ? `${largest.relativePath} at ${String(largest.lineCount)} lines` : '0 lines'
  }).`,
)

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
