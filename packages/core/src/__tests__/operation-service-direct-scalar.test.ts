import { indexToColumn } from '@bilig/formula'
import { ValueTag } from '@bilig/protocol'
import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'
import { SpreadsheetEngine } from '../engine.js'
import { createBatch } from '../replica-state.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

function withOperationServiceBlocked<T>(engine: SpreadsheetEngine, callback: () => T): T {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const descriptor = Object.getOwnPropertyDescriptor(runtime, 'operations')
  if (!descriptor) {
    throw new TypeError('Expected runtime operations descriptor')
  }
  Object.defineProperty(runtime, 'operations', {
    configurable: true,
    get() {
      throw new Error('operation service should not be initialized')
    },
  })
  try {
    return callback()
  } finally {
    Object.defineProperty(runtime, 'operations', descriptor)
  }
}
describe('EngineOperationService direct scalar and local mutation paths', () => {
  it('rejects trusted direct aggregate numeric mutations when lookup dependents share the input column', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-aggregate-lookup-guard',
      trackReplicaVersions: false,
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A32,0)')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 528 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('returns typed changed cells for trusted existing numeric direct scalar chains', async () => {
    const downstreamCount = 24
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-direct-scalar-chain',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    for (let offset = 1; offset <= downstreamCount; offset += 1) {
      const col = offset
      engine.setCellFormula('Sheet1', `${indexToColumn(col)}1`, `${indexToColumn(col - 1)}1+1`)
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const terminalIndex = engine.workbook.getCellIndex('Sheet1', `${indexToColumn(downstreamCount)}1`)!

    engine.resetPerformanceCounters()
    const result = engine.tryApplyExistingNumericCellMutationAt({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 99,
      emitTracked: false,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
    })

    expect(result?.explicitChangedCount).toBe(1)
    expect(result?.changedCellIndices?.length).toBe(downstreamCount + 1)
    expect(result?.changedCellIndices?.[0]).toBe(inputIndex)
    expect(result?.changedCellIndices?.[downstreamCount]).toBe(terminalIndex)
    expect(engine.getCellValue('Sheet1', `${indexToColumn(downstreamCount)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 99 + downstreamCount,
    })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('applies trusted existing numeric direct scalar chains before operation runtime initialization', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-existing-numeric-direct-scalar-pre-ops',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'A1*4')
    engine.setCellFormula('Sheet1', 'C1', 'B1-2')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!

    engine.resetPerformanceCounters()
    const result = withOperationServiceBlocked(engine, () =>
      engine.tryApplyExistingNumericCellMutationAt({
        sheetId,
        row: 0,
        col: 0,
        cellIndex: inputIndex,
        value: 6,
        emitTracked: false,
        trustedExistingNumericLiteral: true,
        oldNumericValue: 3,
      }),
    )

    expect(result?.explicitChangedCount).toBe(1)
    expect(result?.changedCellIndices).toEqual(
      Uint32Array.from([inputIndex, engine.workbook.getCellIndex('Sheet1', 'B1')!, engine.workbook.getCellIndex('Sheet1', 'C1')!]),
    )
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('applies overlapping sliding aggregate deltas on the single-literal fast path', async () => {
    const rowCount = 96
    const window = 16
    const updateRow = 16
    const engine = new SpreadsheetEngine({ workbookName: 'operation-sliding-aggregate-direct-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      const endRow = Math.min(rowCount, row + window - 1)
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A${row}:A${endRow})`)
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', `A${updateRow}`, 99)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 219 })
    expect(engine.getCellValue('Sheet1', `B${updateRow}`)).toEqual({ tag: ValueTag.Number, value: 459 })
    expect(engine.getCellValue('Sheet1', `B${updateRow + 1}`)).toEqual({ tag: ValueTag.Number, value: 392 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(window)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${updateRow}`)
    expect(inputIndex).toBeDefined()
    const event = tracked.mock.calls.at(-1)?.[0]
    expect(event).toEqual(expect.objectContaining({ explicitChangedCount: 1 }))
    const changedIndices = Array.from(event.changedCellIndices)
    const expectedFormulaIndices = Array.from({ length: window }, (_, index) => {
      const cellIndex = engine.workbook.getCellIndex('Sheet1', `B${index + 1}`)
      expect(cellIndex).toBeDefined()
      return cellIndex!
    })
    expect(changedIndices[0]).toBe(inputIndex)
    expect(changedIndices.slice(1).toSorted((left, right) => left - right)).toEqual(
      expectedFormulaIndices.toSorted((left, right) => left - right),
    )
    unsubscribe()
  })

  it('accumulates direct aggregate deltas across generic batch literal writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-aggregate-batch-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 32; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 10 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 20 },
    ])

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 555 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('counts direct scalar generic batch updates without dirty traversal', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-batch-metrics' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('updates dense same-column affine scalar batches without dirty traversal', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-affine-column-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2`)
    }
    const refs: EngineCellMutationRef[] = Array.from({ length: rowCount }, (_, index) => ({
      sheetId,
      cellIndex: engine.workbook.getCellIndex('Sheet1', `A${index + 1}`)!,
      mutation: { kind: 'setCellValue', row: index, col: 0, value: index * 3 },
    }))
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'local', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount - 1) * 6 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 2)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount)
    unsubscribe()
  })

  it('updates descending dense affine scalar undo batches without dirty traversal', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-affine-column-undo-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row * 3)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2`)
    }
    const refs: EngineCellMutationRef[] = Array.from({ length: rowCount }, (_, index) => {
      const row = rowCount - index - 1
      return {
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1 },
      }
    })
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'undo', 0))

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount * 2 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 2)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount)
    unsubscribe()
  })

  it('updates dense row-pair simple scalar batches without dirty traversal', async () => {
    const rowCount = 48
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-row-pair-batch', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row + 1)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `A${row}*B${row}`)
    }
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 0, value: row * 3 },
      })
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `B${row + 1}`)!,
        mutation: { kind: 'setCellValue', row, col: 1, value: row * 5 },
      })
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    const sheet = engine.workbook.getSheetById(sheetId)!
    const inputColumnAVersionBefore = sheet.columnVersions[0] ?? 0
    const inputColumnBVersionBefore = sheet.columnVersions[1] ?? 0

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, null, 'local', 0))

    expect(engine.getCellValue('Sheet1', `C${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount - 1) * 8 })
    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount * 2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(sheet.columnVersions[0]).toBe(inputColumnAVersionBefore + 1)
    expect(sheet.columnVersions[1]).toBe(inputColumnBVersionBefore + 1)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(changed).toHaveLength(rowCount * 4)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(rowCount * 2)
    unsubscribe()
  })

  it('accumulates cell-by-cell direct scalar deltas across same-row batch writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-cell-product-batch-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'A1*B1')

    const batch = createBatch(getReplicaState(engine), [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 7 },
    ])

    engine.resetPerformanceCounters()
    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 35 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(2)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('propagates simple direct scalar chains with numeric deltas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-chain-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')
    engine.setCellFormula('Sheet1', 'D1', 'C1+1')

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 5)

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(3)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
  })

  it('updates large direct scalar fanout with constant bulk deltas', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-bulk-deltas', trackReplicaVersions: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellFormula('Sheet1', `B${row}`, 'A1+1')
    }
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const outputColumnVersionBefore = engine.workbook.getSheetById(sheetId)?.columnVersions[1] ?? 0
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.applyCellMutationsAtWithOptions(
      [
        {
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', 'A1')!,
          mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 },
        },
      ],
      { captureUndo: true, potentialNewCells: 0, source: 'local', returnUndoOps: false, reuseRefs: true },
    )

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.workbook.getSheetById(sheetId)?.columnVersions[1] ?? 0).toBe(outputColumnVersionBefore)
    const changed = tracked.mock.calls.at(-1)?.[0].changedCellIndices
    expect(changed).toBeInstanceOf(Uint32Array)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSheetId')).toBe(sheetId)
    expect(Reflect.get(changed, '__biligTrackedPhysicalSortedSliceSplit')).toBe(1)
    unsubscribe()
  })

  it('keeps mixed direct scalar and aggregate fanout on constant delta storage', async () => {
    const rowCount = 64
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-scalar-aggregate-mixed-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `=$A$1+${row}`)
      engine.setCellFormula('Sheet1', `C${row}`, `=SUM(A1:A${row})`)
    }
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'A1', 99)

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: 99 + rowCount })
    expect(engine.getCellValue('Sheet1', `C${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2 + 98,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)

    const event = tracked.mock.calls.at(-1)?.[0]
    const changedIndices = Array.from(event.changedCellIndices)
    expect(event).toEqual(expect.objectContaining({ explicitChangedCount: 1 }))
    expect(changedIndices[0]).toBe(engine.workbook.getCellIndex('Sheet1', 'A1'))
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `B${rowCount}`))
    expect(changedIndices).toContain(engine.workbook.getCellIndex('Sheet1', `C${rowCount}`))
    unsubscribe()
  })

  it('updates copied SUMIF formulas from aggregate column writes with direct deltas', async () => {
    const formulaCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'operation-direct-criteria-sum-deltas' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Group')
    engine.setCellValue('Sheet1', 'B1', 'Value')
    engine.setCellValue('Sheet1', 'D1', 'A')
    for (let row = 2; row <= 9; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row % 2 === 0 ? 'A' : 'B')
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    for (let index = 0; index < formulaCount; index += 1) {
      engine.setCellFormula('Sheet1', `${indexToColumn(4 + index)}1`, '=SUMIF(A2:A9,D1,B2:B9)')
    }

    engine.resetPerformanceCounters()
    engine.setCellValue('Sheet1', 'B2', 100)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 118 })
    expect(engine.getCellValue('Sheet1', `${indexToColumn(4 + formulaCount - 1)}1`)).toEqual({
      tag: ValueTag.Number,
      value: 118,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(formulaCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('applies local cell mutation refs through the service', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-local-refs', replicaId: 'a' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 0, col: 0, value: 10 },
      },
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' },
      },
      {
        sheetId,
        mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'SUM(' },
      },
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 3, col: 3 },
      },
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 0, col: 0 },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 3))

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Empty })
  })

  it('rejects local cell mutation refs for unknown sheets', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-local-refs-missing',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const refs: EngineCellMutationRef[] = [
      {
        sheetId: 999,
        mutation: { kind: 'setCellValue', row: 0, col: 0, value: 1 },
      },
    ]
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }])

    expect(() => Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))).toThrow('Unknown sheet id: 999')
  })
})
