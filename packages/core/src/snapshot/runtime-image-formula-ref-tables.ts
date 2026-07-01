import type { CompiledFormula } from '@bilig/formula'
import type { CellValue } from '@bilig/protocol'
import type { EngineFormulaSourceRef, EngineFormulaSourceRefTable } from '../cell-mutations-at.js'
import type { DeferredInitialFormulaFamilyRun } from '../engine/services/formula-initialization-family-runs.js'
import type { FormulaInstanceSnapshot } from '../formula/formula-instance-table.js'

export interface PreparedRuntimeFormulaRef {
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly source: string
  readonly compiled: CompiledFormula
  readonly templateId?: number
  readonly cellIndex?: number
}

export interface HydratedPreparedRuntimeFormulaRef extends PreparedRuntimeFormulaRef {
  readonly value: CellValue
  readonly preserveCachedValueOnFullRecalc?: boolean
}

interface MutableHydratedPreparedRuntimeFormulaRef {
  sheetId: number
  row: number
  col: number
  source: string
  compiled: CompiledFormula
  templateId?: number
  cellIndex?: number
  value: CellValue
  preserveCachedValueOnFullRecalc?: boolean
}

export class RestoredFormulaSourceRefTable implements EngineFormulaSourceRefTable {
  readonly sheetIds: Uint32Array
  readonly cellIndices: Uint32Array
  readonly rows: Uint32Array
  readonly cols: Uint32Array
  readonly sources: string[]
  readonly reusable: EngineFormulaSourceRef = {
    sheetId: 0,
    cellIndex: 0,
    row: 0,
    col: 0,
    source: '',
  }
  length = 0

  constructor(capacity: number) {
    this.sheetIds = new Uint32Array(capacity)
    this.cellIndices = new Uint32Array(capacity)
    this.rows = new Uint32Array(capacity)
    this.cols = new Uint32Array(capacity)
    this.sources = []
  }

  push(sheetId: number, cellIndex: number, row: number, col: number, source: string): void {
    const index = this.length
    this.sheetIds[index] = sheetId
    this.cellIndices[index] = cellIndex
    this.rows[index] = row
    this.cols[index] = col
    this.sources[index] = source
    this.length = index + 1
  }

  at(index: number): EngineFormulaSourceRef {
    this.reusable.sheetId = this.sheetIds[index]!
    this.reusable.cellIndex = this.cellIndices[index]!
    this.reusable.row = this.rows[index]!
    this.reusable.col = this.cols[index]!
    this.reusable.source = this.sources[index]!
    return this.reusable
  }
}

export class RestoredHydratedPreparedFormulaRefTable implements Iterable<HydratedPreparedRuntimeFormulaRef> {
  readonly sheetIds: Uint32Array
  readonly cellIndices: Uint32Array
  readonly rows: Uint32Array
  readonly cols: Uint32Array
  readonly templateIds: Int32Array
  readonly sources: string[]
  readonly compiled: CompiledFormula[]
  readonly values: CellValue[]
  readonly preserveCachedValueOnFullRecalc: Uint8Array
  freshFormulaInstances: readonly FormulaInstanceSnapshot[] | undefined
  freshFormulaFamilyRuns: readonly DeferredInitialFormulaFamilyRun[] | undefined
  freshFormulaFamilyRunFallbackCount = 0
  private reusable: MutableHydratedPreparedRuntimeFormulaRef | undefined
  length = 0

  constructor(capacity: number, freshFormulaInstances?: readonly FormulaInstanceSnapshot[]) {
    this.sheetIds = new Uint32Array(capacity)
    this.cellIndices = new Uint32Array(capacity)
    this.rows = new Uint32Array(capacity)
    this.cols = new Uint32Array(capacity)
    this.templateIds = new Int32Array(capacity)
    this.templateIds.fill(-1)
    this.sources = []
    this.compiled = []
    this.values = []
    this.preserveCachedValueOnFullRecalc = new Uint8Array(capacity)
    this.freshFormulaInstances = freshFormulaInstances
  }

  push(
    sheetId: number,
    cellIndex: number,
    row: number,
    col: number,
    source: string,
    compiled: CompiledFormula,
    templateId: number | undefined,
    value: CellValue,
    runtimeImageCellIndex: number,
    preserveCachedValueOnFullRecalc = false,
  ): void {
    const index = this.length
    if (this.freshFormulaInstances !== undefined && runtimeImageCellIndex !== cellIndex) {
      this.freshFormulaInstances = undefined
    }
    this.sheetIds[index] = sheetId
    this.cellIndices[index] = cellIndex
    this.rows[index] = row
    this.cols[index] = col
    this.templateIds[index] = templateId ?? -1
    this.sources[index] = source
    this.compiled[index] = compiled
    this.values[index] = value
    this.preserveCachedValueOnFullRecalc[index] = preserveCachedValueOnFullRecalc ? 1 : 0
    this.length = index + 1
  }

  at(index: number): HydratedPreparedRuntimeFormulaRef {
    const reusable =
      this.reusable ??
      (this.reusable = {
        sheetId: 0,
        row: 0,
        col: 0,
        source: '',
        compiled: this.compiled[index]!,
        value: this.values[index]!,
      })
    reusable.sheetId = this.sheetIds[index]!
    reusable.cellIndex = this.cellIndices[index]!
    reusable.row = this.rows[index]!
    reusable.col = this.cols[index]!
    reusable.source = this.sources[index]!
    reusable.compiled = this.compiled[index]!
    const templateId = this.templateIds[index]!
    if (templateId === -1) {
      delete reusable.templateId
    } else {
      reusable.templateId = templateId
    }
    reusable.value = this.values[index]!
    if (this.preserveCachedValueOnFullRecalc[index] === 1) {
      reusable.preserveCachedValueOnFullRecalc = true
    } else {
      delete reusable.preserveCachedValueOnFullRecalc
    }
    return reusable
  }

  *[Symbol.iterator](): IterableIterator<HydratedPreparedRuntimeFormulaRef> {
    for (let index = 0; index < this.length; index += 1) {
      const templateId = this.templateIds[index]!
      yield {
        sheetId: this.sheetIds[index]!,
        cellIndex: this.cellIndices[index]!,
        row: this.rows[index]!,
        col: this.cols[index]!,
        source: this.sources[index]!,
        compiled: this.compiled[index]!,
        ...(templateId === -1 ? {} : { templateId }),
        value: this.values[index]!,
        ...(this.preserveCachedValueOnFullRecalc[index] === 1 ? { preserveCachedValueOnFullRecalc: true } : {}),
      }
    }
  }
}

export function preparedRuntimeFormulaCellIndices(refs: readonly PreparedRuntimeFormulaRef[]): Uint32Array {
  const cellIndices = new Uint32Array(refs.length)
  for (let index = 0; index < refs.length; index += 1) {
    const cellIndex = refs[index]!.cellIndex
    if (cellIndex === undefined) {
      return new Uint32Array(0)
    }
    cellIndices[index] = cellIndex
  }
  return cellIndices
}
