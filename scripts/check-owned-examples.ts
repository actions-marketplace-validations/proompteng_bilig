#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))

const retainedExamples = [
  'examples/headless-workpaper',
  'examples/recalc-bridge-workflows',
  'examples/serverless-workpaper-api',
  'examples/xlsx-cache-doctor-ci',
  'examples/xlsx-recalculation-node',
  'integrations/n8n-nodes-workpaper',
] as const

const requiredPackageScripts: Readonly<Record<string, readonly string[]>> = {
  'examples/headless-workpaper': ['agent:mcp-xlsx-risk-preflight', 'agent:verify', 'typecheck'],
  'examples/recalc-bridge-workflows': ['smoke'],
  'examples/serverless-workpaper-api': ['test', 'typecheck'],
  'examples/xlsx-recalculation-node': ['start', 'typecheck'],
  'integrations/n8n-nodes-workpaper': ['check'],
}

function gitTrackedFiles(): string[] {
  return execFileSync('git', ['ls-files', '-z'], { cwd: repoRoot, encoding: 'utf8' })
    .split('\0')
    .filter((path) => path.length > 0)
    .toSorted()
}

function readJsonObject(path: string): Record<string, unknown> {
  const value: unknown = JSON.parse(readFileSync(join(repoRoot, path), 'utf8'))
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${path} must be a JSON object`)
  }
  return Object.fromEntries(Object.entries(value))
}

function readStringRecord(value: unknown, context: string): Record<string, string> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error(`${context} must be a JSON object`)
  }
  return Object.fromEntries(Object.entries(value).filter((entry): entry is [string, string] => typeof entry[1] === 'string'))
}

function main(): void {
  const trackedFiles = gitTrackedFiles()
  const violations: string[] = []

  for (const dir of retainedExamples) {
    if (!trackedFiles.some((path) => path.startsWith(`${dir}/`))) {
      violations.push(`${dir}: retained example/integration directory is missing`)
    }
    if (!existsSync(join(repoRoot, dir, 'README.md'))) {
      violations.push(`${dir}: retained example/integration must have README.md`)
    }
  }

  for (const [dir, scripts] of Object.entries(requiredPackageScripts)) {
    const manifest = readJsonObject(`${dir}/package.json`)
    const packageScripts = readStringRecord(manifest['scripts'], `${dir}/package.json scripts`)
    for (const script of scripts) {
      if (!packageScripts[script]) {
        violations.push(`${dir}/package.json: missing required script ${script}`)
      }
    }
  }

  const forbiddenTrackedPatterns = [
    /^examples\/xlsx-cache-doctor-ci\/reports\//u,
    /^examples\/[^/]+\/node_modules\//u,
    /^integrations\/[^/]+\/node_modules\//u,
  ] as const
  for (const path of trackedFiles) {
    for (const pattern of forbiddenTrackedPatterns) {
      if (pattern.test(path)) {
        violations.push(`${path}: generated or installed example artifact is tracked`)
      }
    }
  }

  if (violations.length > 0) {
    console.error(`Owned example check failed:\n${violations.map((violation) => `- ${violation}`).join('\n')}`)
    process.exit(1)
  }
  console.log('Owned example check passed.')
}

main()
