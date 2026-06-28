import { describe, expect, it, vi } from 'vitest'
import { ensureWasmKernelArtifact, resolveWasmKernelArtifactPath } from '../ensure-wasm-kernel.js'

describe('ensureWasmKernelArtifact', () => {
  it('returns the existing artifact path without spawning a build', () => {
    const spawnSync = vi.fn()
    const artifactPath = ensureWasmKernelArtifact({
      rootDir: '/repo',
      existsSync: (path) => path === '/repo/packages/wasm-kernel/build/release.wasm' || path === '/repo/packages/wasm-kernel/dist/index.js',
      spawnSync,
      sourcePaths: ['/repo/packages/wasm-kernel/src/index.ts'],
      outputPaths: ['/repo/packages/wasm-kernel/build/release.wasm', '/repo/packages/wasm-kernel/dist/index.js'],
      statSync: (path) => ({ mtimeMs: path.endsWith('index.ts') ? 1 : 2 }),
    })

    expect(artifactPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(spawnSync).not.toHaveBeenCalled()
  })

  it('builds the artifact when it is missing before startup', () => {
    const existingPaths = new Set<string>()
    const spawnSync = vi.fn((command: string, args: readonly string[], options: { cwd: string }) => {
      expect(command).toBe('pnpm')
      expect(args).toEqual(['wasm:build'])
      expect(options.cwd).toBe('/repo')
      existingPaths.add('/repo/packages/wasm-kernel/build/release.wasm')
      existingPaths.add('/repo/packages/wasm-kernel/dist/index.js')
      return {
        status: 0,
        stdout: Buffer.alloc(0),
        stderr: Buffer.alloc(0),
      }
    })

    const artifactPath = ensureWasmKernelArtifact({
      rootDir: '/repo',
      existsSync: (path) => existingPaths.has(path),
      spawnSync,
      env: {},
      sourcePaths: ['/repo/packages/wasm-kernel/src/index.ts'],
      outputPaths: ['/repo/packages/wasm-kernel/build/release.wasm', '/repo/packages/wasm-kernel/dist/index.js'],
      statSync: (path) => ({ mtimeMs: path.endsWith('index.ts') ? 2 : 3 }),
    })

    expect(artifactPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(spawnSync).toHaveBeenCalledTimes(1)
  })

  it('rebuilds when ignored wasm kernel output is older than source', () => {
    const spawnSync = vi.fn(() => ({
      status: 0,
      stdout: Buffer.alloc(0),
      stderr: Buffer.alloc(0),
    }))

    const artifactPath = ensureWasmKernelArtifact({
      rootDir: '/repo',
      existsSync: () => true,
      spawnSync,
      env: {},
      sourcePaths: ['/repo/packages/wasm-kernel/src/index.ts'],
      outputPaths: ['/repo/packages/wasm-kernel/build/release.wasm', '/repo/packages/wasm-kernel/dist/index.js'],
      statSync: (path) => ({ mtimeMs: path.endsWith('index.ts') ? 10 : 5 }),
    })

    expect(artifactPath).toBe('/repo/packages/wasm-kernel/build/release.wasm')
    expect(spawnSync).toHaveBeenCalledTimes(1)
  })

  it('throws a clear error when the build fails', () => {
    expect(() =>
      ensureWasmKernelArtifact({
        rootDir: '/repo',
        existsSync: () => false,
        spawnSync: () => ({
          status: 1,
          stdout: Buffer.from(''),
          stderr: Buffer.from('boom'),
        }),
        env: {},
        sourcePaths: ['/repo/packages/wasm-kernel/src/index.ts'],
        outputPaths: ['/repo/packages/wasm-kernel/build/release.wasm', '/repo/packages/wasm-kernel/dist/index.js'],
        statSync: () => ({ mtimeMs: 1 }),
      }),
    ).toThrow('Failed to build wasm kernel artifact: boom')
  })

  it('throws when the build exits successfully without producing the artifact', () => {
    expect(() =>
      ensureWasmKernelArtifact({
        rootDir: '/repo',
        existsSync: () => false,
        spawnSync: () => ({
          status: 0,
          stdout: Buffer.alloc(0),
          stderr: Buffer.alloc(0),
        }),
        env: {},
        sourcePaths: ['/repo/packages/wasm-kernel/src/index.ts'],
        outputPaths: ['/repo/packages/wasm-kernel/build/release.wasm'],
        statSync: () => ({ mtimeMs: 1 }),
      }),
    ).toThrow("did not create '/repo/packages/wasm-kernel/build/release.wasm'")
  })
})

describe('resolveWasmKernelArtifactPath', () => {
  it('derives the release.wasm path from the workspace root', () => {
    expect(resolveWasmKernelArtifactPath('/repo')).toBe('/repo/packages/wasm-kernel/build/release.wasm')
  })
})
