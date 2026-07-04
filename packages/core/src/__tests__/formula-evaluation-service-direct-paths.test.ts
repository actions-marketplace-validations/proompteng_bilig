import { ErrorCode, ValueTag } from '@bilig/protocol'
import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'
import type { RuntimeFormula } from '../engine/runtime-state.js'
import type { EngineFormulaEvaluationService } from '../engine/services/formula-evaluation-service.js'

function isEngineFormulaEvaluationService(value: unknown): value is EngineFormulaEvaluationService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'evaluateUnsupportedFormula') === 'function' &&
    typeof Reflect.get(value, 'resolveStructuredReference') === 'function' &&
    typeof Reflect.get(value, 'resolveSpillReference') === 'function'
  )
}

function isRuntimeFormulaTable(value: unknown): value is { get(cellIndex: number): RuntimeFormula | undefined } {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'get') === 'function'
}

function getEvaluationService(engine: SpreadsheetEngine): EngineFormulaEvaluationService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const evaluation = Reflect.get(runtime, 'evaluation')
  if (!isEngineFormulaEvaluationService(evaluation)) {
    throw new TypeError('Expected engine formula evaluation service')
  }
  return evaluation
}
function getInternalFormulaStore(engine: SpreadsheetEngine): { get(cellIndex: number): RuntimeFormula | undefined } {
  const formulas = Reflect.get(engine, 'formulas')
  if (!isRuntimeFormulaTable(formulas)) {
    throw new TypeError('Expected internal formulas store')
  }
  return formulas
}

function readRuntimeDirectLookup(engine: SpreadsheetEngine, sheetName: string, address: string): object | undefined {
  const formulas = getInternalFormulaStore(engine)
  const cellIndex = engine.workbook.getCellIndex(sheetName, address)
  if (cellIndex === undefined) {
    throw new Error(`expected runtime formula at ${sheetName}!${address}`)
  }
  const runtimeFormula = formulas.get(cellIndex)
  if (typeof runtimeFormula !== 'object' || runtimeFormula === null) {
    throw new Error(`expected runtime formula at ${sheetName}!${address}`)
  }
  const directLookup = Reflect.get(runtimeFormula, 'directLookup')
  if (typeof directLookup !== 'object' || directLookup === null) {
    return undefined
  }
  return directLookup
}

function readRuntimeDirectLookupKind(engine: SpreadsheetEngine, sheetName: string, address: string): string | undefined {
  const directLookup = readRuntimeDirectLookup(engine, sheetName, address)
  if (directLookup === undefined) {
    return undefined
  }
  const kind = Reflect.get(directLookup, 'kind')
  return typeof kind === 'string' ? kind : undefined
}

function readRuntimeDirectLookupTailPatch(engine: SpreadsheetEngine, sheetName: string, address: string): unknown {
  const directLookup = readRuntimeDirectLookup(engine, sheetName, address)
  if (directLookup === undefined) {
    return undefined
  }
  return Reflect.get(directLookup, 'tailPatch')
}

function overwriteRuntimeNumber(engine: SpreadsheetEngine, sheetName: string, address: string, value: number): number {
  const cellIndex = engine.workbook.getCellIndex(sheetName, address)
  if (cellIndex === undefined) {
    throw new Error(`expected cell at ${sheetName}!${address}`)
  }
  const { cellStore } = engine.workbook
  cellStore.tags[cellIndex] = ValueTag.Number
  cellStore.errors[cellIndex] = ErrorCode.None
  cellStore.stringIds[cellIndex] = 0
  cellStore.numbers[cellIndex] = value
  return cellIndex
}

describe('EngineFormulaEvaluationService: direct lookup and aggregate paths', () => {
  it('evaluates direct approximate lookup formulas across uniform, refreshed, and text columns', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-approx-service' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 1)
    engine.setCellValue('Sheet1', 'C1', 'apple')
    engine.setCellValue('Sheet1', 'C2', 'banana')
    engine.setCellValue('Sheet1', 'C3', 'pear')
    engine.setCellValue('Sheet1', 'D1', true)
    engine.setCellValue('Sheet1', 'D2', 2.5)
    engine.setCellValue('Sheet1', 'D3', 'peach')
    engine.setCellValue('Sheet1', 'D4', 5)

    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D1,A1:A3,1)')
    engine.setCellFormula('Sheet1', 'F2', 'MATCH(D2,B1:B3,-1)')
    engine.setCellFormula('Sheet1', 'F3', 'MATCH(D3,C1:C3,1)')
    engine.setCellFormula('Sheet1', 'F4', 'MATCH(D4,C1:C3,1)')

    const evaluation = getEvaluationService(engine)
    const f1Index = engine.workbook.getCellIndex('Sheet1', 'F1')
    const f2Index = engine.workbook.getCellIndex('Sheet1', 'F2')
    const f3Index = engine.workbook.getCellIndex('Sheet1', 'F3')
    const f4Index = engine.workbook.getCellIndex('Sheet1', 'F4')
    expect(f1Index).toBeDefined()
    expect(f2Index).toBeDefined()
    expect(f3Index).toBeDefined()
    expect(f4Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 1 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!))
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 2 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(f4Index!))
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellValue('Sheet1', 'A3', 5)
    engine.setCellValue('Sheet1', 'D1', 4.5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f1Index!))
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' }, [[6], [5], [4]])
    engine.setCellValue('Sheet1', 'D2', 4.5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f2Index!))
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValue('Sheet1', 'C2', 'blueberry')
    Effect.runSync(evaluation.evaluateDirectLookupFormula(f3Index!))
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('refreshes direct uniform numeric lookup descriptors across exact and approximate branches', async () => {
    const exactEngine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-exact-uniform-refresh',
      useColumnIndex: true,
    })
    await exactEngine.ready()
    exactEngine.createSheet('Sheet1')
    exactEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[1], [2], [3], [4]])
    exactEngine.setCellValue('Sheet1', 'B1', 3)
    exactEngine.setCellFormula('Sheet1', 'C1', 'MATCH(B1,A1:A4,0)')

    const exactEvaluation = getEvaluationService(exactEngine)
    const exactIndex = exactEngine.workbook.getCellIndex('Sheet1', 'C1')
    expect(exactIndex).toBeDefined()
    expect(readRuntimeDirectLookupKind(exactEngine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')

    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })

    exactEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[12], [22], [32], [42]])
    exactEngine.setCellValue('Sheet1', 'B1', 32)
    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(readRuntimeDirectLookupKind(exactEngine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')

    exactEngine.setCellValue('Sheet1', 'B1', '32')
    Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))
    expect(exactEngine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    exactEngine.setCellFormula('Sheet1', 'B1', '1/0')
    expect(Effect.runSync(exactEvaluation.evaluateDirectLookupFormula(exactIndex!))).toBeUndefined()

    const approximateEngine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-approx-uniform-refresh',
    })
    await approximateEngine.ready()
    approximateEngine.createSheet('Sheet1')
    approximateEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' }, [
      [10, 40],
      [20, 30],
      [30, 20],
      [40, 10],
    ])
    approximateEngine.setCellValue('Sheet1', 'D1', 35)
    approximateEngine.setCellValue('Sheet1', 'E1', 25)
    approximateEngine.setCellFormula('Sheet1', 'F1', 'MATCH(D1,A1:A4,1)')
    approximateEngine.setCellFormula('Sheet1', 'F2', 'MATCH(E1,B1:B4,-1)')

    const approximateEvaluation = getEvaluationService(approximateEngine)
    const ascendingIndex = approximateEngine.workbook.getCellIndex('Sheet1', 'F1')
    const descendingIndex = approximateEngine.workbook.getCellIndex('Sheet1', 'F2')
    expect(ascendingIndex).toBeDefined()
    expect(descendingIndex).toBeDefined()
    expect(readRuntimeDirectLookupKind(approximateEngine, 'Sheet1', 'F1')).toBe('approximate-uniform-numeric')
    expect(readRuntimeDirectLookupKind(approximateEngine, 'Sheet1', 'F2')).toBe('approximate-uniform-numeric')

    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', null)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', 50)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })

    approximateEngine.setCellValue('Sheet1', 'D1', 'pear')
    expect(Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))).toBeUndefined()

    approximateEngine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[100], [90], [80], [70]])
    approximateEngine.setCellValue('Sheet1', 'D1', 85)
    expect(Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(ascendingIndex!))).toBeUndefined()

    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    approximateEngine.setCellValue('Sheet1', 'E1', 50)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    approximateEngine.setCellValue('Sheet1', 'E1', 5)
    Effect.runSync(approximateEvaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(approximateEngine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
  })

  it('evaluates direct exact uniform lookup tail patches without rebuilding owners', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'evaluation-direct-exact-uniform-tail-patch',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A4' }, [[1], [2], [3], [4]])
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'MATCH(B1,A1:A4,0)')

    const evaluation = getEvaluationService(engine)
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(formulaIndex).toBeDefined()
    expect(readRuntimeDirectLookupKind(engine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')
    const directLookup = readRuntimeDirectLookup(engine, 'Sheet1', 'C1')
    expect(directLookup).toBeDefined()
    const tailIndex = overwriteRuntimeNumber(engine, 'Sheet1', 'A4', 10)
    engine.workbook.notifyCellValueWritten(tailIndex)
    Reflect.set(directLookup!, 'tailPatch', {
      row: 3,
      oldNumeric: 4,
      newNumeric: 10,
      columnVersion: engine.workbook.getSheet('Sheet1')!.columnVersions[0],
    })
    expect(readRuntimeDirectLookupTailPatch(engine, 'Sheet1', 'C1')).toEqual(
      expect.objectContaining({ row: 3, oldNumeric: 4, newNumeric: 10 }),
    )

    overwriteRuntimeNumber(engine, 'Sheet1', 'B1', 4)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(formulaIndex!))
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    overwriteRuntimeNumber(engine, 'Sheet1', 'B1', 10)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(formulaIndex!))
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(readRuntimeDirectLookupKind(engine, 'Sheet1', 'C1')).toBe('exact-uniform-numeric')
  })

  it('evaluates direct approximate uniform lookup tail patches in both sort directions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-approx-uniform-tail-patch' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B4' }, [
      [1, 4],
      [2, 3],
      [3, 2],
      [4, 1],
    ])
    engine.setCellValue('Sheet1', 'D1', 2.5)
    engine.setCellValue('Sheet1', 'D2', 2.5)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A4,1)')
    engine.setCellFormula('Sheet1', 'E2', 'MATCH(D2,B1:B4,-1)')

    const evaluation = getEvaluationService(engine)
    const ascendingIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const descendingIndex = engine.workbook.getCellIndex('Sheet1', 'E2')
    expect(ascendingIndex).toBeDefined()
    expect(descendingIndex).toBeDefined()
    const ascendingLookup = readRuntimeDirectLookup(engine, 'Sheet1', 'E1')
    const descendingLookup = readRuntimeDirectLookup(engine, 'Sheet1', 'E2')
    expect(ascendingLookup).toBeDefined()
    expect(descendingLookup).toBeDefined()
    const ascendingTailIndex = overwriteRuntimeNumber(engine, 'Sheet1', 'A4', 10)
    const descendingTailIndex = overwriteRuntimeNumber(engine, 'Sheet1', 'B4', -5)
    engine.workbook.notifyCellValueWritten(ascendingTailIndex)
    engine.workbook.notifyCellValueWritten(descendingTailIndex)
    const sheet = engine.workbook.getSheet('Sheet1')!
    Reflect.set(ascendingLookup!, 'tailPatch', {
      row: 3,
      oldNumeric: 4,
      newNumeric: 10,
      columnVersion: sheet.columnVersions[0],
    })
    Reflect.set(descendingLookup!, 'tailPatch', {
      row: 3,
      oldNumeric: 1,
      newNumeric: -5,
      columnVersion: sheet.columnVersions[1],
    })
    expect(readRuntimeDirectLookupTailPatch(engine, 'Sheet1', 'E1')).toEqual(
      expect.objectContaining({ row: 3, oldNumeric: 4, newNumeric: 10 }),
    )
    expect(readRuntimeDirectLookupTailPatch(engine, 'Sheet1', 'E2')).toEqual(
      expect.objectContaining({ row: 3, oldNumeric: 1, newNumeric: -5 }),
    )

    overwriteRuntimeNumber(engine, 'Sheet1', 'D1', 6)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })

    overwriteRuntimeNumber(engine, 'Sheet1', 'D1', 10)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(ascendingIndex!))
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })

    overwriteRuntimeNumber(engine, 'Sheet1', 'D2', 0)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 3 })

    overwriteRuntimeNumber(engine, 'Sheet1', 'D2', -5)
    Effect.runSync(evaluation.evaluateDirectLookupFormula(descendingIndex!))
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 4 })
  })

  it('evaluates direct aggregate formulas with progressive prefixes and coercion rules', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-service' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A4', 'skip')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A4)')
    engine.setCellFormula('Sheet1', 'B2', 'AVERAGE(A1:A4)')
    engine.setCellFormula('Sheet1', 'B3', 'COUNT(A1:A4)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A1)')
    engine.setCellFormula('Sheet1', 'C2', 'SUM(A1:A2)')
    engine.setCellFormula('Sheet1', 'C3', 'SUM(A1:A4)')
    engine.setCellFormula('Sheet1', 'C4', 'SUM(A2:A4)')
    engine.setCellFormula('Sheet1', 'C5', 'AVERAGE(A2:A4)')

    const evaluation = getEvaluationService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')
    const b3Index = engine.workbook.getCellIndex('Sheet1', 'B3')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    const c2Index = engine.workbook.getCellIndex('Sheet1', 'C2')
    const c3Index = engine.workbook.getCellIndex('Sheet1', 'C3')
    const c4Index = engine.workbook.getCellIndex('Sheet1', 'C4')
    const c5Index = engine.workbook.getCellIndex('Sheet1', 'C5')
    expect(b1Index).toBeDefined()
    expect(b2Index).toBeDefined()
    expect(b3Index).toBeDefined()
    expect(c1Index).toBeDefined()
    expect(c2Index).toBeDefined()
    expect(c3Index).toBeDefined()
    expect(c4Index).toBeDefined()
    expect(c5Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(c1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c3Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c4Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(c5Index!))
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'C5')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })

    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 1 })

    engine.setCellFormula('Sheet1', 'A4', 'NA()')
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 1 })
  })

  it('evaluates direct aggregate formulas with formula members from live cell state', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-live' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'A2', 'A1*3')
    engine.setCellFormula('Sheet1', 'A3', 'IF(A1>0,NA(),"ok")')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'B2', 'COUNT(A1:A3)')
    engine.setCellFormula('Sheet1', 'B3', 'MIN(A1:A3)')
    engine.setCellFormula('Sheet1', 'B4', 'MAX(A1:A3)')

    const evaluation = getEvaluationService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const b2Index = engine.workbook.getCellIndex('Sheet1', 'B2')
    const b3Index = engine.workbook.getCellIndex('Sheet1', 'B3')
    const b4Index = engine.workbook.getCellIndex('Sheet1', 'B4')
    expect(b1Index).toBeDefined()
    expect(b2Index).toBeDefined()
    expect(b3Index).toBeDefined()
    expect(b4Index).toBeDefined()

    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b2Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b3Index!))
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b4Index!))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('evaluates short direct aggregate windows without building column owners', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-short-window' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let expected = 0
    for (let row = 1; row <= 16; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      expected += row
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A16)')

    const evaluation = getEvaluationService(engine)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(b1Index).toBeDefined()

    engine.resetPerformanceCounters()
    Effect.runSync(evaluation.evaluateDirectLookupFormula(b1Index!))

    expect(engine.getPerformanceCounters().columnOwnerBuilds).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateScanCells).toBe(16)
    expect(engine.getPerformanceCounters().directAggregatePrefixEvaluations).toBe(0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: expected })
  })

  it('promotes sliding aggregate windows to shared prefix evaluation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'evaluation-direct-aggregate-prefix-window' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 64; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, 1)
      engine.setCellValue('Sheet1', `C${row}`, 'ignored')
      if (row % 2 === 1) {
        engine.setCellValue('Sheet1', `D${row}`, 1)
      }
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A32)')
    engine.setCellFormula('Sheet1', 'B2', 'SUM(A2:A33)')
    engine.setCellFormula('Sheet1', 'B3', 'COUNT(A1:A32)')
    engine.setCellFormula('Sheet1', 'B4', 'AVERAGE(A1:A32)')
    engine.setCellFormula('Sheet1', 'B5', 'AVERAGE(C1:C32)')
    engine.setCellFormula('Sheet1', 'B6', 'AVERAGE(D1:D32)')

    const evaluation = getEvaluationService(engine)
    const formulaIndices = ['B1', 'B2', 'B3', 'B4', 'B5', 'B6'].map((address) => engine.workbook.getCellIndex('Sheet1', address))
    expect(formulaIndices.every((index) => index !== undefined)).toBe(true)

    engine.resetPerformanceCounters()
    for (const formulaIndex of formulaIndices) {
      Effect.runSync(evaluation.evaluateDirectLookupFormula(formulaIndex!))
    }

    const counters = engine.getPerformanceCounters()
    expect(counters.directAggregateScanEvaluations).toBe(0)
    expect(counters.directAggregateScanCells).toBe(0)
    expect(counters.directAggregatePrefixEvaluations).toBe(6)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: 1 })
  })
})
