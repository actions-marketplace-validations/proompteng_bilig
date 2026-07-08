#!/usr/bin/env bun

import { spawn, type ChildProcess } from 'node:child_process'
import { readFileSync, statSync } from 'node:fs'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

interface BuildAttemptResult {
  readonly code: number | null
  readonly elapsedMs: number
  readonly signal: string | null
  readonly timedOut: boolean
}

type TerminationSignal = 'SIGTERM' | 'SIGKILL'

const repoRoot = fileURLToPath(new URL('..', import.meta.url))
const webRoot = resolve(repoRoot, 'apps/web')
const viteCli = resolve(repoRoot, 'node_modules/vite/bin/vite.js')
const timeoutMs = readPositiveIntegerEnv('BILIG_WEB_BUNDLE_BUILD_TIMEOUT_MS', 120_000)
const attempts = readPositiveIntegerEnv('BILIG_WEB_BUNDLE_BUILD_ATTEMPTS', 2)
const killGraceMs = readPositiveIntegerEnv('BILIG_WEB_BUNDLE_BUILD_KILL_GRACE_MS', 5_000)
const entryScriptMaxBytes = readPositiveIntegerEnv('BILIG_WEB_BUNDLE_ENTRY_SCRIPT_MAX_BYTES', 128 * 1024)
const entryStylesheetMaxBytes = readPositiveIntegerEnv('BILIG_WEB_BUNDLE_ENTRY_STYLESHEET_MAX_BYTES', 128 * 1024)
const command = [process.execPath, viteCli, 'build', '--configLoader', 'runner'] as const

function readPositiveIntegerEnv(name: string, fallback: number): number {
  const raw = process.env[name]
  if (raw === undefined) {
    return fallback
  }
  const parsed = Number.parseInt(raw, 10)
  if (!Number.isSafeInteger(parsed) || parsed <= 0 || String(parsed) !== raw) {
    throw new Error(`${name} must be a positive base-10 integer, got ${JSON.stringify(raw)}`)
  }
  return parsed
}

function formatSeconds(ms: number): string {
  return `${(ms / 1000).toFixed(1)}s`
}

function terminateProcessTree(child: ChildProcess, signal: TerminationSignal): void {
  const pid = child.pid
  if (pid === undefined) {
    return
  }
  try {
    if (process.platform !== 'win32') {
      process.kill(-pid, signal)
      return
    }
  } catch {
    // Fall through to direct child termination below. Some hosts deny process-group signals.
  }
  child.kill(signal)
}

function runBuildAttempt(attempt: number): Promise<BuildAttemptResult> {
  return new Promise((resolveAttempt, rejectAttempt) => {
    const startedAt = performance.now()
    let timedOut = false
    let settled = false
    let killFallback: ReturnType<typeof setTimeout> | null = null
    const child = spawn(command[0], command.slice(1), {
      cwd: webRoot,
      detached: process.platform !== 'win32',
      env: process.env,
      stdio: 'inherit',
    })

    const timeout = setTimeout(() => {
      timedOut = true
      console.error(
        `[web-bundle] attempt ${String(attempt)}/${String(attempts)} timed out after ${formatSeconds(timeoutMs)}; terminating Vite build`,
      )
      terminateProcessTree(child, 'SIGTERM')
      killFallback = setTimeout(() => {
        terminateProcessTree(child, 'SIGKILL')
      }, killGraceMs)
    }, timeoutMs)

    child.once('error', (error) => {
      clearTimeout(timeout)
      if (killFallback) {
        clearTimeout(killFallback)
      }
      rejectAttempt(error)
    })

    child.once('close', (code, signal) => {
      if (settled) {
        return
      }
      settled = true
      clearTimeout(timeout)
      if (killFallback) {
        clearTimeout(killFallback)
      }
      resolveAttempt({
        code,
        elapsedMs: performance.now() - startedAt,
        signal,
        timedOut,
      })
    })
  })
}

function htmlAttributeValues(html: string, tagName: 'link' | 'script', attributeName: 'href' | 'src', requiredToken?: string): string[] {
  const values: string[] = []
  const tagPattern = new RegExp(`<${tagName}\\b[^>]*>`, 'giu')
  for (const tagMatch of html.matchAll(tagPattern)) {
    const tag = tagMatch[0]
    if (requiredToken && !tag.includes(requiredToken)) {
      continue
    }
    const attributePattern = new RegExp(`\\b${attributeName}=["']([^"']+)["']`, 'iu')
    const attributeMatch = attributePattern.exec(tag)
    const value = attributeMatch?.[1]
    if (value) {
      values.push(value)
    }
  }
  return values
}

function localDistAssetBytes(publicPath: string): number | undefined {
  if (!publicPath.startsWith('/assets/')) {
    return undefined
  }
  return statSync(resolve(webRoot, 'dist', publicPath.slice(1))).size
}

function checkWebBundleEntryBudget(): number {
  const indexHtml = readFileSync(resolve(webRoot, 'dist/index.html'), 'utf8')
  const violations: string[] = []
  const modulePreloads = htmlAttributeValues(indexHtml, 'link', 'href', 'rel="modulepreload"')
  for (const modulePreload of modulePreloads) {
    violations.push(`${modulePreload}: modulepreload is disabled for the bounded startup shell`)
  }
  for (const scriptPath of htmlAttributeValues(indexHtml, 'script', 'src', 'type="module"')) {
    const bytes = localDistAssetBytes(scriptPath)
    if (bytes !== undefined && bytes > entryScriptMaxBytes) {
      violations.push(`${scriptPath}: entry script is ${String(bytes)} bytes, limit is ${String(entryScriptMaxBytes)}`)
    }
  }
  for (const stylesheetPath of htmlAttributeValues(indexHtml, 'link', 'href', 'rel="stylesheet"')) {
    const bytes = localDistAssetBytes(stylesheetPath)
    if (bytes !== undefined && bytes > entryStylesheetMaxBytes) {
      violations.push(`${stylesheetPath}: entry stylesheet is ${String(bytes)} bytes, limit is ${String(entryStylesheetMaxBytes)}`)
    }
  }
  if (violations.length === 0) {
    return 0
  }
  console.error(`[web-bundle] entry bundle budget failed:\n${violations.map((violation) => `- ${violation}`).join('\n')}`)
  return 1
}

async function runBuildWithRetries(attempt: number): Promise<number> {
  const result = await runBuildAttempt(attempt)
  if (result.code === 0 && !result.timedOut) {
    return checkWebBundleEntryBudget()
  }
  const reason = result.timedOut
    ? `timed out after ${formatSeconds(timeoutMs)}`
    : result.signal
      ? `failed with signal ${result.signal}`
      : `failed with exit ${String(result.code)}`
  if (attempt < attempts) {
    console.error(`[web-bundle] attempt ${String(attempt)}/${String(attempts)} ${reason}; retrying`)
    return runBuildWithRetries(attempt + 1)
  }
  console.error(`[web-bundle] attempt ${String(attempt)}/${String(attempts)} ${reason}`)
  return result.code ?? 1
}

process.exit(await runBuildWithRetries(1))
