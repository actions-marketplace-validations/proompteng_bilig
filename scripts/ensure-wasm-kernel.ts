import { spawnSync, type SpawnSyncReturns } from 'node:child_process'
import { existsSync, readdirSync, statSync, type Dirent, type Stats } from 'node:fs'
import { join } from 'node:path'
import { workspaceRootDir } from './workspace-resolution.js'

interface SpawnSyncLike {
  (
    command: string,
    args: readonly string[],
    options: {
      cwd: string
      env: NodeJS.ProcessEnv | undefined
      stdio: 'pipe'
    },
  ): SpawnSyncReturns<Buffer>
}

interface EnsureWasmKernelArtifactOptions {
  readonly rootDir?: string
  readonly existsSync?: (path: string) => boolean
  readonly spawnSync?: SpawnSyncLike
  readonly statSync?: (path: string) => Pick<Stats, 'mtimeMs'>
  readonly env?: NodeJS.ProcessEnv
  readonly outputPaths?: readonly string[]
  readonly sourcePaths?: readonly string[]
}

export function resolveWasmKernelArtifactPath(rootDir = workspaceRootDir): string {
  return join(rootDir, 'packages/wasm-kernel/build/release.wasm')
}

export function resolveWasmKernelOutputPaths(rootDir = workspaceRootDir): string[] {
  return [resolveWasmKernelArtifactPath(rootDir), join(rootDir, 'packages/wasm-kernel/dist/index.js')]
}

export function resolveWasmKernelSourcePaths(rootDir = workspaceRootDir): string[] {
  const statFile = statSync
  const readDirectory = readdirSync
  const roots = [
    join(rootDir, 'packages/wasm-kernel/assembly'),
    join(rootDir, 'packages/wasm-kernel/src'),
    join(rootDir, 'packages/wasm-kernel/asconfig.json'),
    join(rootDir, 'packages/wasm-kernel/package.json'),
    join(rootDir, 'packages/wasm-kernel/scripts'),
    join(rootDir, 'packages/wasm-kernel/tsconfig.json'),
  ]
  return roots.flatMap((path) => collectWasmKernelSourceFiles(path, statFile, readDirectory))
}

function formatSpawnFailure(result: SpawnSyncReturns<Buffer>): string {
  if (result.error) {
    return result.error.message
  }
  const stderr = result.stderr.toString().trim()
  const stdout = result.stdout.toString().trim()
  return stderr || stdout || `Exited with status ${String(result.status)}`
}

export function ensureWasmKernelArtifact(options: EnsureWasmKernelArtifactOptions = {}): string {
  const rootDir = options.rootDir ?? workspaceRootDir
  const artifactPath = resolveWasmKernelArtifactPath(rootDir)
  const artifactExists = options.existsSync ?? existsSync
  const outputPaths = options.outputPaths ?? resolveWasmKernelOutputPaths(rootDir)
  const sourcePaths = options.sourcePaths ?? resolveWasmKernelSourcePaths(rootDir)
  const statFile = options.statSync ?? statSync
  if (wasmKernelOutputsAreFresh({ outputPaths, sourcePaths, existsSync: artifactExists, statSync: statFile })) {
    return artifactPath
  }

  const spawnBuild = options.spawnSync ?? spawnSync
  const buildResult = spawnBuild('pnpm', ['wasm:build'], {
    cwd: rootDir,
    env: options.env ?? process.env,
    stdio: 'pipe',
  })
  if (buildResult.status !== 0) {
    throw new Error(`Failed to build wasm kernel artifact: ${formatSpawnFailure(buildResult)}`)
  }
  const missingOutputPath = outputPaths.find((outputPath) => !artifactExists(outputPath))
  if (missingOutputPath) {
    throw new Error(`pnpm wasm:build completed but did not create '${missingOutputPath}'.`)
  }
  return artifactPath
}

function wasmKernelOutputsAreFresh(args: {
  readonly outputPaths: readonly string[]
  readonly sourcePaths: readonly string[]
  readonly existsSync: (path: string) => boolean
  readonly statSync: (path: string) => Pick<Stats, 'mtimeMs'>
}): boolean {
  if (args.outputPaths.some((path) => !args.existsSync(path))) {
    return false
  }
  const newestSourceMtime = newestMtime(args.sourcePaths, args.statSync)
  const oldestOutputMtime = oldestMtime(args.outputPaths, args.statSync)
  return oldestOutputMtime >= newestSourceMtime
}

function newestMtime(paths: readonly string[], readStat: (path: string) => Pick<Stats, 'mtimeMs'>): number {
  let newest = 0
  for (const path of paths) {
    newest = Math.max(newest, readStat(path).mtimeMs)
  }
  return newest
}

function oldestMtime(paths: readonly string[], readStat: (path: string) => Pick<Stats, 'mtimeMs'>): number {
  let oldest = Number.POSITIVE_INFINITY
  for (const path of paths) {
    oldest = Math.min(oldest, readStat(path).mtimeMs)
  }
  return oldest
}

function collectWasmKernelSourceFiles(
  path: string,
  readStat: (path: string) => Pick<Stats, 'isDirectory' | 'isFile'>,
  readDirectory: (path: string, options: { withFileTypes: true }) => Dirent[],
): string[] {
  const stats = readStat(path)
  if (stats.isFile()) {
    return isWasmKernelSourceFile(path) ? [path] : []
  }
  if (!stats.isDirectory()) {
    return []
  }
  return readDirectory(path, { withFileTypes: true }).flatMap((entry) =>
    collectWasmKernelSourceFiles(join(path, entry.name), readStat, readDirectory),
  )
}

function isWasmKernelSourceFile(path: string): boolean {
  if (path.endsWith('.test.ts') || path.includes('/__tests__/')) {
    return false
  }
  return path.endsWith('.ts') || path.endsWith('.js') || path.endsWith('.json') || path.endsWith('.wasm') || path.endsWith('package.json')
}
