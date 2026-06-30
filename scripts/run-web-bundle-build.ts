#!/usr/bin/env bun

import { spawn, type ChildProcess } from 'node:child_process'
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

async function runBuildWithRetries(attempt: number): Promise<number> {
  const result = await runBuildAttempt(attempt)
  if (result.code === 0 && !result.timedOut) {
    return 0
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
