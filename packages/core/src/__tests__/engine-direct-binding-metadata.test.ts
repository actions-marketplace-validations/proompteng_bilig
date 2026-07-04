import type { EngineOpBatch } from './engine-test-helpers.js'
import {
  ErrorCode,
  FormulaMode,
  SpreadsheetEngine,
  ValueTag,
  afterEach,
  describe,
  expect,
  isRuntimeFormulaWithDependencies,
  isRuntimeFormulaWithDirectLookup,
  it,
  readRuntimeFormula,
  vi,
} from './engine-test-helpers.js'

describe('SpreadsheetEngine direct binding and metadata', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('uses the direct js path for approximate sorted MATCH', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'approx-lookup' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'C1', 1)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellValue('Sheet1', 'C3', 5)
    engine.setCellValue('Sheet1', 'D1', 4)

    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,C1:C3,1)')
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const operandCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(isRuntimeFormulaWithDirectLookup(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.directLookup.kind).toBe('approximate-uniform-numeric')
    expect(runtimeFormula?.directLookup.operandCellIndex).toBe(operandCellIndex)
    expect(runtimeFormula?.directLookup).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 2,
      length: 3,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'C2', 4)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })
  })

  it('uses the compact direct js path for repeated approximate MATCH keys', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'approx-lookup-duplicates' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 1, 2, 2, 3, 3].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    engine.setCellValue('Sheet1', 'D1', 2.5)

    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A6,1)')
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(isRuntimeFormulaWithDirectLookup(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.directLookup).toMatchObject({
      kind: 'approximate-uniform-numeric',
      start: 1,
      step: 1,
      repeatedRunLength: 2,
      length: 6,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'D1', 3)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })
  })

  it('skips dirtying approximate MATCH when an irrelevant high-side tail value changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'approx-lookup-tail' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4, 5].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    engine.setCellValue('Sheet1', 'D1', 2.5)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A5,1)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValue('Sheet1', 'A5', 6)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
  })

  it('uses the direct indexed path for exact MATCH when column indexing is enabled', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'indexed-lookup', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'D1', 2)

    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A3,0)')
    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const operandCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(isRuntimeFormulaWithDirectLookup(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.directLookup.kind).toBe('exact-uniform-numeric')
    expect(runtimeFormula?.directLookup.operandCellIndex).toBe(operandCellIndex)
    expect(runtimeFormula?.directLookup).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 10)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })

    engine.setCellValue('Sheet1', 'A2', 20)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    engine.setCellValue('Sheet1', 'D1', 3)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0, wasmFormulaCount: 0 })
  })

  it('binds exact uniform MATCH without lookup-owner construction when column indexing is disabled', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'uniform-lookup-without-index', useColumnIndex: false })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'D1', 2)

    engine.resetPerformanceCounters()
    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A3,0)')

    const formulaCellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    const runtimeFormula = formulaCellIndex === undefined ? undefined : readRuntimeFormula(engine, formulaCellIndex)
    expect(runtimeFormula?.directLookup?.kind).toBe('exact-uniform-numeric')
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      columnOwnerBuilds: 0,
      exactIndexBuilds: 0,
      lookupOwnerBuilds: 0,
    })
  })

  it('caches exact indexed lookup impact across irrelevant coordinate batch writes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'indexed-lookup-batch-impact', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 0; row < 10; row += 1) {
      engine.setCellValueAt(sheetId, row, 0, row + 1)
    }
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A10,0)')

    engine.applyCellMutationsAtWithOptions(
      Array.from({ length: 4 }, (_, offset) => ({
        sheetId,
        mutation: { kind: 'setCellValue' as const, row: 6 + offset, col: 0, value: 100 + offset },
      })),
      { captureUndo: false, potentialNewCells: 0, returnUndoOps: false },
    )

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
  })

  it('caches exact indexed lookup impact across ordinary transaction batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'indexed-lookup-op-batch-impact', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A10,0)')
    const executeLocalTransaction = Reflect.get(engine, 'executeLocalTransaction')
    if (typeof executeLocalTransaction !== 'function') {
      throw new TypeError('Expected executeLocalTransaction')
    }

    executeLocalTransaction.call(engine, [
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A8', value: 108 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A9', value: 109 },
      { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A10', value: 110 },
    ])

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
  })

  it('marks exact indexed lookup dependents once a coordinate batch touches the operand key', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'indexed-lookup-batch-hit', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 0; row < 10; row += 1) {
      engine.setCellValueAt(sheetId, row, 0, row + 1)
    }
    engine.setCellValue('Sheet1', 'D1', 2)
    engine.setCellFormula('Sheet1', 'E1', '=MATCH(D1,A1:A10,0)')

    engine.applyCellMutationsAtWithOptions(
      [
        { sheetId, mutation: { kind: 'setCellValue', row: 7, col: 0, value: 107 } },
        { sheetId, mutation: { kind: 'setCellValue', row: 1, col: 0, value: 20 } },
      ],
      { captureUndo: false, potentialNewCells: 0, returnUndoOps: false },
    )

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 1,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
  })

  it('uses the direct indexed path for exact string MATCH and reverse XMATCH when column indexing is enabled', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'indexed-string-lookup',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'pear')
    engine.setCellValue('Sheet1', 'A2', 'apple')
    engine.setCellValue('Sheet1', 'A3', 'pear')

    engine.setCellFormula('Sheet1', 'B1', '=MATCH("APPLE",A1:A3,0)')
    engine.setCellFormula('Sheet1', 'B2', '=XMATCH("pear",A1:A3,0,-1)')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'B2').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 3 })

    engine.setCellValue('Sheet1', 'A3', 'banana')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 2, wasmFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity info and date builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 'hello')
    engine.setCellValue('Sheet1', 'A4', 45351)
    engine.setCellValue('Sheet1', 'A5', 45351.75)
    engine.setCellValue('Sheet1', 'A6', 60)
    engine.setCellValue('Sheet1', 'A7', 45322)
    engine.setCellValue('Sheet1', 'A8', 45337)
    engine.setCellValue('Sheet1', 'A9', 'bad')
    engine.setCellValue('Sheet1', 'A10', 0.5208333333333334)
    engine.setCellValue('Sheet1', 'A11', 0.5208449074074074)

    engine.setCellFormula('Sheet1', 'B1', 'ISBLANK(A12)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'ISNUMBER(A1)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'ISTEXT(A3)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'ISBLANK(A1)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'ISNUMBER(A1)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'ISTEXT(A3)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B7', 'DATE(2024,2,29)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 45351 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B8', 'YEAR(B7)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: 2024 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B9', 'MONTH(A5)')
    expect(engine.getCellValue('Sheet1', 'B9')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B10', 'DAY(A6)')
    expect(engine.getCellValue('Sheet1', 'B10')).toEqual({ tag: ValueTag.Number, value: 29 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B11', 'EDATE(A7,1.9)')
    expect(engine.getCellValue('Sheet1', 'B11')).toEqual({ tag: ValueTag.Number, value: 45351 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B12', 'EOMONTH(A8,A2)')
    expect(engine.getCellValue('Sheet1', 'B12')).toEqual({ tag: ValueTag.Number, value: 45382 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B13', 'DATE(A9,2,29)')
    expect(engine.getCellValue('Sheet1', 'B13')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B14', 'EDATE(A9,1)')
    expect(engine.getCellValue('Sheet1', 'B14')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B15', 'EOMONTH(A9,1)')
    expect(engine.getCellValue('Sheet1', 'B15')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B16', 'TIME(12,30,0)')
    expect(engine.getCellValue('Sheet1', 'B16')).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B17', 'HOUR(A10)')
    expect(engine.getCellValue('Sheet1', 'B17')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B18', 'MINUTE(A10)')
    expect(engine.getCellValue('Sheet1', 'B18')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B19', 'SECOND(A11)')
    expect(engine.getCellValue('Sheet1', 'B19')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B20', 'WEEKDAY(DATE(2026,3,15))')
    expect(engine.getCellValue('Sheet1', 'B20')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity information and threshold helpers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 42)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 'alpha')

    engine.setCellFormula('Sheet1', 'B1', 'T(A3)')
    engine.setCellFormula('Sheet1', 'B2', 'N(A2)')
    engine.setCellFormula('Sheet1', 'B3', 'TYPE(A3)')
    engine.setCellFormula('Sheet1', 'B4', 'DELTA(4,4)')
    engine.setCellFormula('Sheet1', 'B5', 'GESTEP(-1)')
    engine.setCellFormula('Sheet1', 'B6', 'GAUSS(0)')
    engine.setCellFormula('Sheet1', 'B7', 'PHI(0)')

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'alpha',
    })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
    expect(engine.getCellValue('Sheet1', 'B7')).toMatchObject({
      tag: ValueTag.Number,
      value: 0.3989422804014327,
    })

    for (const address of ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for literal and dynamic VALUE coercion', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'VALUE("42")')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', '  -17.25e1  ')
    engine.setCellFormula('Sheet1', 'B2', 'VALUE(A1)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: -172.5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A2', 'not-a-number')
    engine.setCellFormula('Sheet1', 'B3', 'VALUE(A2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity LEN builtin', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', null)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 123.4)
    engine.setCellValue('Sheet1', 'A4', 'hello')
    engine.setCellFormula('Sheet1', 'A5', 'A3/0')

    engine.setCellFormula('Sheet1', 'B1', 'LEN(A1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'LEN(A2)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'LEN(A3)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'LEN(A4)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'LEN(A5)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for EXACT text equality', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Alpha')
    engine.setCellValue('Sheet1', 'A2', 'alpha')

    engine.setCellFormula('Sheet1', 'B1', 'EXACT(A1,A1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'EXACT(A1,A2)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for string literals, direct refs, and CONCAT', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'world')

    engine.setCellFormula('Sheet1', 'B1', '"hello"')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'hello',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'A1')
    expect(engine.getCellValue('Sheet1', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'world',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'CONCAT("hi ",A1)')
    expect(engine.getCellValue('Sheet1', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'hi world',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for text slicing, casing, and search builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Alpha')
    engine.setCellValue('Sheet1', 'A2', '  alpha   beta  ')
    engine.setCellValue('Sheet1', 'A3', 'alpha')
    engine.setCellValue('Sheet1', 'A4', 'BETA')
    engine.setCellValue('Sheet1', 'A5', 'alphabet')

    engine.setCellFormula('Sheet1', 'B1', 'LEFT(A1,2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Al',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'RIGHT(A1)')
    expect(engine.getCellValue('Sheet1', 'B2')).toMatchObject({ tag: ValueTag.String, value: 'a' })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B3', 'MID(A1,2,3)')
    expect(engine.getCellValue('Sheet1', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'lph',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B4', 'TRIM(A2)')
    expect(engine.getCellValue('Sheet1', 'B4')).toMatchObject({
      tag: ValueTag.String,
      value: 'alpha beta',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B5', 'UPPER(A3)')
    expect(engine.getCellValue('Sheet1', 'B5')).toMatchObject({
      tag: ValueTag.String,
      value: 'ALPHA',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'LOWER(A4)')
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.String,
      value: 'beta',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B7', 'FIND("ph",A5,3)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B8', 'SEARCH("P*",A5)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('coerces numeric text expressions in wasm SUM without summing referenced text cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', "5'7")
    engine.setCellValue('Sheet1', 'A2', '5')

    engine.setCellFormula('Sheet1', 'B1', 'SUM(LEFT(A1,1),RIGHT(A1,LEN(A1)-2)/12)')
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Number,
      value: 5 + 7 / 12,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B2', 'SUM(A2)')
    expect(engine.getCellValue('Sheet1', 'B2')).toMatchObject({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('coerces comma-grouped numeric cell text in direct arithmetic and JS formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', '61,111')
    engine.setCellValue('Sheet1', 'B1', '72,522')

    engine.setCellFormula('Sheet1', 'C1', 'B1/A1')
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: 72522 / 61111,
    })

    engine.setCellFormula('Sheet1', 'D1', '(POWER(B1/A1,0.3333333333)-1)*100')
    const cagr = engine.getCellValue('Sheet1', 'D1')
    expect(cagr.tag).toBe(ValueTag.Number)
    if (cagr.tag !== ValueTag.Number) {
      throw new Error(`Expected numeric CAGR, received ${JSON.stringify(cagr)}`)
    }
    expect(cagr.value).toBeCloseTo(5.872571270499272, 12)
  })

  it('uses the wasm fast path for exact-parity rounding builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 123.4)
    engine.setCellFormula('Sheet1', 'B1', 'ROUND(A1,-1)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B2', 'FLOOR(TRUE,0.5)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B3', 'CEILING(7,2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B4', 'FLOOR(A1,0)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellValue('Sheet1', 'A2', 'oops')
    engine.setCellFormula('Sheet1', 'B5', 'ROUND(A2,1)')
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B6', 'INT(-3.1)')
    expect(engine.getCellValue('Sheet1', 'B6')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B7', 'ROUNDUP(-3.141,2)')
    expect(engine.getCellValue('Sheet1', 'B7')).toEqual({ tag: ValueTag.Number, value: -3.15 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B8', 'ROUNDDOWN(-3.141,2)')
    expect(engine.getCellValue('Sheet1', 'B8')).toEqual({ tag: ValueTag.Number, value: -3.14 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)
  })

  it('preserves topo order across mixed wasm and js formula runs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 5)
    engine.setCellFormula('Sheet1', 'B2', 'A1+A2')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(2:2)')

    engine.setCellValue('Sheet1', 'A1', 12)

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(2)
    expect(engine.getLastMetrics().jsFormulaCount).toBe(0)
  })

  it('rebinds formulas when a referenced sheet appears later', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'Sheet2!B1*2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.createSheet('Sheet2')
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet2', 'B1', 3)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
  })

  it('rebinds bounded cross-sheet ranges from #REF! back onto the wasm path when a sheet appears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'SUM(Sheet2!A1:A2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.createSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet2', 'A1', 2)
    engine.setCellValue('Sheet2', 'A2', 3)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
  })

  it('rebinds formulas to #REF! when a referenced sheet is deleted', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet2', 'B1', 3)
    engine.setCellFormula('Sheet1', 'A1', 'Sheet2!B1*2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.deleteSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('rebinds bounded cross-sheet ranges to #REF! when a referenced sheet is deleted', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet2', 'A1', 2)
    engine.setCellValue('Sheet2', 'A2', 3)
    engine.setCellFormula('Sheet1', 'A1', 'SUM(Sheet2!A1:A2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.deleteSheet('Sheet2')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('keeps sheet-qualified aggregate templates scoped to their source sheets', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data1')
    engine.createSheet('Data2')
    engine.createSheet('Summary')
    engine.setCellValue('Data1', 'A1', 2)
    engine.setCellValue('Data2', 'A1', 7)
    engine.setCellFormula('Summary', 'A1', 'SUM(Data1!A1:A1)')
    engine.setCellFormula('Summary', 'A2', 'SUM(Data2!A1:A1)')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 7 })
  })

  it('renames sheets without breaking formulas, names, or sheet metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [
      [2, 3],
      [4, 5],
    ])
    engine.setCellFormula('Data', 'C1', 'A1+B1')
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!A1:B2)')
    engine.setDefinedName('AnchorCell', { kind: 'cell-ref', sheetName: 'Data', address: 'A1' })
    engine.setDefinedName('SalesRange', '=Data!A1:B2')
    engine.setFreezePane('Data', 1, 0)
    engine.setFilter('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' })
    engine.setSort('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Q1', 'Q2'],
      headerRow: false,
      totalsRow: false,
    })
    engine.setPivotTable('Summary', 'D2', {
      name: 'SalesPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B2' },
      groupBy: ['Q1'],
      values: [{ sourceColumn: 'Q2', summarizeBy: 'sum' }],
    })

    engine.renameSheet('Data', 'Revenue')

    expect(engine.exportSnapshot().sheets.map((sheet) => sheet.name)).toEqual(['Revenue', 'Summary'])
    expect(engine.getCell('Revenue', 'C1').formula).toBe('A1+B1')
    expect(engine.getCellValue('Revenue', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCell('Summary', 'A1').formula).toBe('SUM(Revenue!A1:B2)')
    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getDefinedName('AnchorCell')).toEqual({
      name: 'AnchorCell',
      value: { kind: 'cell-ref', sheetName: 'Revenue', address: 'A1' },
    })
    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: '=Revenue!A1:B2',
    })
    expect(engine.getFreezePane('Revenue')).toEqual({ sheetName: 'Revenue', rows: 1, cols: 0 })
    expect(engine.getFreezePane('Data')).toBeUndefined()
    expect(engine.getFilters('Revenue')).toEqual([
      {
        sheetName: 'Revenue',
        range: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
      },
    ])
    expect(engine.getSorts('Revenue')).toEqual([
      {
        sheetName: 'Revenue',
        range: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Revenue',
        startAddress: 'A1',
        endAddress: 'B2',
        columnNames: ['Q1', 'Q2'],
        headerRow: false,
        totalsRow: false,
      },
    ])
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesPivot',
        sheetName: 'Summary',
        address: 'D2',
        source: { sheetName: 'Revenue', startAddress: 'A1', endAddress: 'B2' },
        cacheFields: ['2', '3'],
        cachedRecords: [[4, 5]],
        groupBy: ['Q1'],
        values: [{ sourceColumn: 'Q2', summarizeBy: 'sum' }],
        rows: 1,
        cols: 1,
      },
    ])
  })

  it('clears reverse range edges when a range-backed formula is removed', async () => {
    const engine = new SpreadsheetEngine()
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A:A)')

    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toContain('Sheet1!B1')

    engine.clearCell('Sheet1', 'B1')

    expect(engine.getDependencies('Sheet1', 'A1').directDependents).toEqual([])
  })

  it('rebinds column and row range formulas when new cells materialize later', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A:A)')
    engine.setCellFormula('Sheet1', 'B3', 'SUM(2:2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 0 })

    engine.setCellValue('Sheet1', 'A4', 3)
    engine.setCellValue('Sheet1', 'C2', 5)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getDependencies('Sheet1', 'B1').directPrecedents).toEqual(['Sheet1!A1', 'Sheet1!A4'])
    expect(engine.getDependencies('Sheet1', 'B3').directPrecedents).toEqual(['Sheet1!C2'])

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(b1Index).toBeDefined()
    const runtimeFormula = b1Index === undefined ? undefined : readRuntimeFormula(engine, b1Index)
    expect(isRuntimeFormulaWithDependencies(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.dependencyIndices).toBeInstanceOf(Uint32Array)
  })

  it('converges under reordered replicated batches and restores replica state', async () => {
    const engineA = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const engineB = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([engineA.ready(), engineB.ready()])

    const outboundA: EngineOpBatch[] = []
    const outboundB: EngineOpBatch[] = []
    engineA.subscribeBatches((batch) => outboundA.push(batch))
    engineB.subscribeBatches((batch) => outboundB.push(batch))

    engineA.createSheet('Sheet1')
    engineB.createSheet('Sheet1')
    engineA.setCellValue('Sheet1', 'A1', 1)
    engineB.setCellValue('Sheet1', 'A1', 2)

    ;[...outboundB].toReversed().forEach((batch) => engineA.applyRemoteBatch(batch))
    ;[...outboundA].forEach((batch) => engineB.applyRemoteBatch(batch))

    expect(engineA.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engineB.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })

    const restored = new SpreadsheetEngine({ workbookName: 'restored', replicaId: 'b' })
    await restored.ready()
    restored.importSnapshot(engineB.exportSnapshot())
    restored.importReplicaSnapshot(engineB.exportReplicaSnapshot())

    const latestOutboundA = outboundA.at(-1)
    expect(latestOutboundA).toBeDefined()
    restored.applyRemoteBatch(latestOutboundA)
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('ignores duplicate remote batches and stale cell replays behind sheet tombstones', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    const createBatch = outbound.at(-1)
    expect(createBatch).toBeDefined()

    primary.setCellValue('Sheet1', 'A1', 7)
    const valueBatch = outbound.at(-1)
    expect(valueBatch).toBeDefined()

    replica.applyRemoteBatch(createBatch)
    expect(replica.applyRemoteBatch(valueBatch)).toBe(true)
    const versionBeforeDuplicate = replica.explainCell('Sheet1', 'A1').version

    expect(replica.applyRemoteBatch(valueBatch)).toBe(false)
    expect(replica.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(replica.explainCell('Sheet1', 'A1').version).toBe(versionBeforeDuplicate)

    primary.deleteSheet('Sheet1')
    const deleteBatch = outbound.at(-1)
    expect(deleteBatch).toBeDefined()
    replica.applyRemoteBatch(deleteBatch)
    expect(replica.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    const restored = new SpreadsheetEngine({ workbookName: 'restored', replicaId: 'b' })
    await restored.ready()
    restored.importSnapshot(replica.exportSnapshot())
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot())

    restored.applyRemoteBatch(valueBatch)
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
  })

  it('resolves workbook defined names through engine metadata and recalculates dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 100)
    engine.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    const changed: number[][] = []
    const unsubscribe = engine.subscribe((event) => {
      changed.push(Array.from(event.changedCellIndices))
    })

    const a2Index = engine.workbook.getCellIndex('Sheet1', 'A2')
    expect(a2Index).toBeDefined()

    engine.setDefinedName('TaxRate', 0.085)

    expect(engine.getDefinedName('taxrate')).toEqual({ name: 'TaxRate', value: 0.085 })
    expect(engine.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.085 }])
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8.5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })
    expect(changed.at(-1)).toContain(a2Index!)

    engine.setDefinedName('TAXRATE', 0.09)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    expect(engine.deleteDefinedName('taxrate')).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    unsubscribe()
  })

  it('replicates defined-name batches and replays them through transaction history', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    primary.setCellValue('Sheet1', 'A1', 100)
    primary.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')
    outbound.forEach((batch) => replica.applyRemoteBatch(batch))

    primary.setDefinedName('TaxRate', 0.08)
    const defineBatch = outbound.at(-1)
    expect(defineBatch?.ops).toEqual([{ kind: 'upsertDefinedName', name: 'TaxRate', value: 0.08 }])
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    replica.applyRemoteBatch(defineBatch)
    expect(replica.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    expect(primary.undo()).toBe(true)
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    expect(primary.redo()).toBe(true)
    expect(primary.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 8 })

    primary.deleteDefinedName('taxrate')
    const deleteBatch = outbound.at(-1)
    expect(deleteBatch?.ops).toEqual([{ kind: 'deleteDefinedName', name: 'taxrate' }])

    replica.applyRemoteBatch(deleteBatch)
    expect(replica.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
  })

  it('emits lightweight tracked events for ordinary mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const tracked = vi.fn()

    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.setCellValue('Sheet1', 'A1', 7)

    expect(tracked).toHaveBeenCalledTimes(1)
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices: new Uint32Array([0]),
        invalidatedRows: [],
        invalidatedColumns: [],
      }),
    )

    unsubscribe()
  })
})
