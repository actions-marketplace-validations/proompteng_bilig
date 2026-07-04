import { ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import type { EngineCellMutationRef } from '../cell-mutations-at.js'
import { SpreadsheetEngine } from '../engine.js'

function operationHook<Args extends readonly unknown[], Result>(engine: SpreadsheetEngine, name: string): (...args: Args) => Result {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const hooks = Reflect.get(operations, '__testHooks')
  if (typeof hooks !== 'object' || hooks === null) {
    throw new TypeError('Expected operation hooks')
  }
  const hook = Reflect.get(hooks, name)
  if (typeof hook !== 'function') {
    throw new TypeError(`Expected operation hook ${name}`)
  }
  return (...args: Args): Result => Reflect.apply(hook, hooks, args)
}

function runtimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null) {
    throw new TypeError('Expected formulas store')
  }
  const get = Reflect.get(formulas, 'get')
  if (typeof get !== 'function') {
    throw new TypeError('Expected formulas get method')
  }
  return get.call(formulas, cellIndex)
}

describe('operation-service literal and localized fast paths', () => {
  it('applies multi-dependent direct scalar literal mutations without events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-direct-scalar-multiple-no-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'A1+2')

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[{ existingIndex: number; value: number; oldNumber: number; newNumber: number }], boolean>(
      engine,
      'tryApplySingleDirectScalarLiteralMutationWithoutEvents',
    )({
      existingIndex: inputIndex,
      value: 5,
      oldNumber: 2,
      newNumber: 5,
    })

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(2)
  })

  it('emits tracked events for kernel-sync-only literal mutation fast paths', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'kernel-sync-literal-hook-tracked' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    let afterWriteCalled = false

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[{ existingIndex: number; value: number; emitTracked: boolean; afterWrite: () => void }], boolean>(
      engine,
      'tryApplySingleKernelSyncOnlyLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      value: 9,
      emitTracked: true,
      afterWrite: () => {
        afterWriteCalled = true
      },
    })

    expect(applied).toBe(true)
    expect(afterWriteCalled).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies direct lookup operand mutation fast paths with tracked changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-lookup-operand-hook-tracked', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          formulaCellIndex: number
          value: number
          exactLookupValue: number | undefined
          approximateLookupValue: number | undefined
          emitTracked: boolean
          lookupSheetHint: unknown
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectLookupOperandMutationFastPath',
    )({
      existingIndex: inputIndex,
      formulaCellIndex: formulaIndex,
      value: 5,
      exactLookupValue: 5,
      approximateLookupValue: undefined,
      emitTracked: true,
      lookupSheetHint: engine.workbook.getSheet('Sheet1'),
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex, formulaIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies direct aggregate literal mutation fast paths with tracked changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-tracked' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A10)')
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 100,
      delta: 99,
      emitTracked: true,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 154 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex, formulaIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies single existing exact lookup column writes through the direct literal gate', async () => {
    const rowCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-exact-lookup-gate', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: rowCount - 1, col: 0, value: rowCount + 100 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount + 100 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('applies single existing approximate lookup column writes through the direct literal gate', async () => {
    const rowCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-approximate-lookup-gate' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: rowCount - 1, col: 0, value: rowCount + 1 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount + 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('coalesces direct scalar and aggregate dependents through the single literal gate', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-scalar-aggregate-gate' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A10)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: 0, col: 0, value: 100 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 101 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 154 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
  })

  it('patches uniform lookup tail writes for single and multiple dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'uniform-lookup-tail-patch-hook', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', 9)
    engine.setCellValue('Sheet1', 'D2', 9)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A5,0)')
    engine.setCellFormula('Sheet1', 'F1', 'XMATCH(D2,A1:A5,0)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const patch = operationHook<
      [
        {
          sheetId: number
          col: number
          row: number
          oldNumeric: number
          newNumeric: number
          exact: boolean
          sorted: boolean
        },
      ],
      { exact: boolean; sorted: boolean }
    >(
      engine,
      'patchUniformLookupTailWrites',
    )({
      sheetId,
      col: 0,
      row: 4,
      oldNumeric: 5,
      newNumeric: 9,
      exact: true,
      sorted: false,
    })

    expect(patch).toEqual({ exact: true, sorted: false })
    for (const address of ['E1', 'F1']) {
      const formulaIndex = engine.workbook.getCellIndex('Sheet1', address)!
      const formula = runtimeFormula(engine, formulaIndex)
      const directLookup = typeof formula === 'object' && formula !== null ? Reflect.get(formula, 'directLookup') : undefined
      expect(directLookup).toEqual(
        expect.objectContaining({
          tailPatch: expect.objectContaining({ row: 4, oldNumeric: 5, newNumeric: 9 }),
        }),
      )
    }
  })

  it('applies direct lookup current-result branches when exact matches disappear', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-direct-lookup-current-result-no-events', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const applied = operationHook<
      [
        {
          existingIndex: number
          formulaCellIndex: number
          value: number
          oldNumber: number
          newNumber: number
          exactLookupValue: number | undefined
          approximateLookupValue: number | undefined
        },
      ],
      boolean
    >(
      engine,
      'tryApplySingleDirectFormulaLiteralMutationWithoutEvents',
    )({
      existingIndex: inputIndex,
      formulaCellIndex: formulaIndex,
      value: 9,
      oldNumber: 3,
      newNumber: 9,
      exactLookupValue: 9,
      approximateLookupValue: undefined,
    })

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: ValueTag.Error })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
  })

  it('uses aggregate fast path for writes with no applicable range dependent', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-no-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 20,
      col: 0,
      value: 5,
      delta: 4,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: 1,
      dirtyFormulaCount: 0,
    })
  })

  it('applies aggregate delta fast path to a single dependent without tracked events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-single-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A2')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      { changedCellIndices: Uint32Array } | null
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      value: 20,
      delta: 18,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('emits tracked aggregate delta changes for multiple direct dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-multiple-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A2)')

    const tracked = vi.fn()
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      { changedCellIndices: Uint32Array } | null
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 10,
      delta: 9,
      emitTracked: true,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        explicitChangedCount: 1,
        changedCellIndices: expect.any(Uint32Array),
      }),
    )
    unsubscribeTracked()
  })
})
