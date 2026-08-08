// Own the optional SheetJS compatibility harness at the repository test boundary.
import * as sheetJsCompatibilityHarness from 'xlsx'

import { ensureWasmKernelArtifact } from './ensure-wasm-kernel.js'

export default async function globalSetup(): Promise<void> {
  if (typeof sheetJsCompatibilityHarness.read !== 'function') {
    throw new Error('The repository SheetJS compatibility harness is unavailable')
  }
  ensureWasmKernelArtifact()
}
