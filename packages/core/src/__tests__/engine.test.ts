import { ErrorCode, FormulaMode, SpreadsheetEngine, ValueTag, afterEach, describe, expect, it, vi } from './engine-test-helpers.js'

describe('SpreadsheetEngine core mutations and dynamic arrays', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('recalculates simple formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })

    engine.setCellValue('Sheet1', 'A1', 12)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
  })

  it('recalculateDirty performs incremental recalculation from regions', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'C1+10')

    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(20)
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15)

    // Update A1 and C1 without recalculating immediately.
    // SpreadsheetEngine methods mostly call applyBatch which always recalculates.
    // We'll use the internal workbook store to bypass immediate calc.
    const a1Index = engine.workbook.ensureCell('Sheet1', 'A1')
    const c1Index = engine.workbook.ensureCell('Sheet1', 'C1')

    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 50 })
    engine.workbook.cellStore.setValue(c1Index, { tag: ValueTag.Number, value: 100 })

    // Verify values are updated but NOT recalculated (B1 and D1 should have old results)
    expect(engine.getCellValue('Sheet1', 'A1').value).toBe(50)
    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(20) // Still old result
    expect(engine.getCellValue('Sheet1', 'C1').value).toBe(100)
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15) // Still old result

    // Recalculate only A1's region
    const changed = engine.recalculateDirty([{ sheetName: 'Sheet1', rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }])

    // Should contain B1 (0,1) because it depends on A1 (0,0)
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')!
    expect([...changed]).toContain(b1Index)

    // D1 (0,3) should NOT be in changed because we didn't mark C1 (0,2) as dirty
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')!
    expect([...changed]).not.toContain(d1Index)

    expect(engine.getCellValue('Sheet1', 'B1').value).toBe(100) // 50 * 2
    expect(engine.getCellValue('Sheet1', 'D1').value).toBe(15) // Still old value
  })

  it('evaluates string concatenation and string comparisons on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'hello')
    engine.setCellFormula('Sheet1', 'B1', 'A1&" world"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    engine.setCellFormula('Sheet1', 'C1', 'A1="HELLO"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    engine.setCellFormula('Sheet1', 'D1', '"b">"A"')
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'hello world',
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Boolean, value: true })
  })

  it('reports differential drift when restored recalculation produces different changed-cell sets', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'A1', 12)

    const differential = engine.recalculateDifferential()

    expect(differential.drift).toEqual([])
    expect(differential.wasm).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sheetName: 'Sheet1',
          address: 'B1',
          value: { tag: ValueTag.Number, value: 24 },
        }),
      ]),
    )
    expect(differential.js).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          sheetName: 'Sheet1',
          address: 'B1',
          value: { tag: ValueTag.Number, value: 24 },
        }),
      ]),
    )
  })

  it('relocates relative formulas when copying a range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 5)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
    )

    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*2')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('preserves absolute references when copying formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellFormula('Sheet1', 'B1', '$A1+A$1+$A$1')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'C2', endAddress: 'C2' },
    )

    expect(engine.getCell('Sheet1', 'C2').formula).toBe('$A2+B$1+$A$1')
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('moves a range and clears the source cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B2', 'left')
    engine.setCellValue('Sheet1', 'C2', 'right')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C2' },
      { sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'E4' },
    )

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'left',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'E4')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'right',
      }),
    )
  })

  it('copies, fills, and moves cell presentation without leaving stale target fills', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'range-style-state' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B2', 'styled')
    engine.setCellValue('Sheet1', 'B3', 'plain')
    engine.setCellValue('Sheet1', 'D2', 'target')
    engine.setCellValue('Sheet1', 'D3', 'target')
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, { fill: { backgroundColor: '#34a853' } })
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, '0.00')
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D3' }, { fill: { backgroundColor: '#93c5fd' } })
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D3' }, '$0')

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B3' },
      { sheetName: 'Sheet1', startAddress: 'D2', endAddress: 'D3' },
    )

    expect(engine.getCellStyle(engine.getCell('Sheet1', 'D2').styleId)?.fill?.backgroundColor).toBe('#34a853')
    expect(engine.getCell('Sheet1', 'D2').format).toBe('0.00')
    expect(engine.getCell('Sheet1', 'D3').styleId).toBeUndefined()
    expect(engine.getCell('Sheet1', 'D3').format).toBeUndefined()

    engine.setCellValue('Sheet1', 'F2', 'fill-target')
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'F2', endAddress: 'F3' }, { fill: { backgroundColor: '#93c5fd' } })
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'F2', endAddress: 'F3' }, '$0')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
      { sheetName: 'Sheet1', startAddress: 'F2', endAddress: 'F3' },
    )

    expect(engine.getCellStyle(engine.getCell('Sheet1', 'F2').styleId)?.fill?.backgroundColor).toBe('#34a853')
    expect(engine.getCellStyle(engine.getCell('Sheet1', 'F3').styleId)?.fill?.backgroundColor).toBe('#34a853')
    expect(engine.getCell('Sheet1', 'F2').format).toBe('0.00')
    expect(engine.getCell('Sheet1', 'F3').format).toBe('0.00')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'F2', endAddress: 'F2' },
      { sheetName: 'Sheet1', startAddress: 'H2', endAddress: 'H2' },
    )

    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCell('Sheet1', 'F2').styleId).toBeUndefined()
    expect(engine.getCell('Sheet1', 'F2').format).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'H2')).toMatchObject({ tag: ValueTag.String, value: 'styled' })
    expect(engine.getCellStyle(engine.getCell('Sheet1', 'H2').styleId)?.fill?.backgroundColor).toBe('#34a853')
    expect(engine.getCell('Sheet1', 'H2').format).toBe('0.00')
  })

  it('treats copying empty cells into tracked empty dependencies as a history no-op', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'copy-undo-empty-targets-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    seed.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const snapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'copy-undo-empty-targets' })
    await engine.ready()
    engine.importSnapshot(snapshot)

    const beforeCopy = engine.exportSnapshot()

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(beforeCopy)
    expect(engine.undo()).toBe(false)
    expect(engine.exportSnapshot()).toEqual(beforeCopy)
  })

  it('treats clearing an already-empty tracked dependency cell as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-empty-dependency-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    engine.clearCell('Sheet1', 'D4')

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats sheet-id clear mutations on already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-empty-dependency-noop-by-id' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const before = engine.exportSnapshot()
    engine.clearCellAt(sheetId, 3, 3)

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats copying empty cells into empty targets as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'copy-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const before = engine.exportSnapshot()

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats filling empty cells into empty targets as a no-op', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fill-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const before = engine.exportSnapshot()

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'E6' },
      { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'D4' },
    )

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('undoes formula fills into blank targets without leaving explicit empty cells behind', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'fill-undo-blank-target-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const initialSnapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'fill-undo-blank-target' })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula('Sheet1', 'A1', 'E5+A1')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'E5', endAddress: 'E5' },
    )
    engine.setRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, '0.00')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
    )
    engine.insertColumns('Sheet1', 0, 1)

    let undoCount = 0
    while (engine.undo()) {
      undoCount += 1
      expect(undoCount).toBeLessThanOrEqual(16)
    }
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('undoes formula creation on tracked dependency placeholders without exporting authored blanks', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'formula-undo-placeholder-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    const initialSnapshot = seed.exportSnapshot()

    const engine = new SpreadsheetEngine({ workbookName: 'formula-undo-placeholder' })
    await engine.ready()
    engine.importSnapshot(initialSnapshot)

    engine.setCellFormula('Sheet1', 'A1', 'C3+A1')
    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
    )
    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' })
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' }, [[null]])
    engine.setCellFormula('Sheet1', 'C3', 'A1+A1')

    let undoCount = 0
    while (engine.undo()) {
      undoCount += 1
      expect(undoCount).toBeLessThanOrEqual(16)
    }
    expect(undoCount).toBeGreaterThan(0)
    expect(engine.exportSnapshot()).toEqual(initialSnapshot)
  })

  it('applies cell mutations by sheet id and returns inverse ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cell-mutation-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const undoOps = engine.applyCellMutationsAt([
      {
        sheetId,
        mutation: {
          kind: 'setCellValue',
          row: 0,
          col: 0,
          value: 10,
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 1,
          formula: 'A1*2',
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'clearCell',
          row: 3,
          col: 3,
        },
      },
      {
        sheetId,
        mutation: {
          kind: 'setCellFormula',
          row: 0,
          col: 2,
          formula: 'SUM(',
        },
      },
    ])

    expect(undoOps).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ kind: 'clearCell', address: 'A1' }),
        expect.objectContaining({ kind: 'clearCell', address: 'B1' }),
        expect.objectContaining({ kind: 'clearCell', address: 'D4' }),
      ]),
    )
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCell('Sheet1', 'C1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getLastMetrics().compileMs).toBeGreaterThanOrEqual(0)
  })

  it('supports direct cell mutations by coordinates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-cell-mutations' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    expect(engine.setCellValueAt(sheetId, 1, 1, 5)).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(engine.setCellFormulaAt(sheetId, 1, 2, 'B2*3')).toEqual({
      tag: ValueTag.Number,
      value: 15,
    })

    engine.clearCellAt(sheetId, 1, 1)

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(() => engine.setCellValueAt(999, 0, 0, 1)).toThrow('Unknown sheet id: 999')
  })

  it('moves overlapping ranges without losing cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'first')
    engine.setCellValue('Sheet1', 'B1', 'second')

    engine.moveRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1' },
    )

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'first',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual(
      expect.objectContaining({
        tag: ValueTag.String,
        value: 'second',
      }),
    )
  })

  it('relocates formulas when filling down', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B3' },
    )

    expect(engine.getCell('Sheet1', 'B2').formula).toBe('A2*3')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('A3*3')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('validates bulk range helpers and no-ops empty fills', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    expect(() => engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [[1]])).toThrow(
      'setRangeValues requires a value matrix that exactly matches the target range',
    )
    expect(() => engine.setRangeFormulas({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [['A1']])).toThrow(
      'setRangeFormulas requires a formula matrix that exactly matches the target range',
    )

    engine.fillRange(
      { sheetName: 'Missing', startAddress: 'A1', endAddress: 'A1' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D2' },
    )
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Empty })

    expect(() =>
      engine.copyRange(
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
        { sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'C1' },
      ),
    ).toThrow('copyRange requires source and target dimensions to match exactly')
  })

  it('stores invalid formulas as #VALUE errors instead of throwing', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    expect(() => engine.setCellFormula('Sheet1', 'A1', '1+')).not.toThrow()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('spills sequence formulas through the runtime and recalculates downstream refs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'A3*2')
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 3, cols: 1 }])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(restored.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(restored.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(restored.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 3, cols: 1 }])
  })

  it('keeps spill references live when dynamic arrays shrink to one cell and grow again', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(A1,1,1,1)')
    engine.setCellFormula('Sheet1', 'D1', 'ROWS(B1#)')
    engine.setCellFormula('Sheet1', 'E1', 'SUM(B1#)')
    engine.setCellFormula('Sheet1', 'F1', 'IFERROR(INDEX(B1#,2),"missing")')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValue('Sheet1', 'A1', 1)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({ tag: ValueTag.String, value: 'missing' })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 1, cols: 1 }])

    engine.setCellValue('Sheet1', 'A1', 2)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 2, cols: 1 }])
  })

  it('clears prior sequence spills when the owner becomes a scalar', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'A3*2')
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    engine.setCellValue('Sheet1', 'A1', 7)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toBeUndefined()
  })

  it('blocks sequence spills when target cells are occupied', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A2', 99)
    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Spill,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'A1', rows: 1, cols: 1 }])
  })

  it('blocks and unblocks dynamic arrays when spill children are authored and cleared', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(A1,1,1,1)')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(B1#)')
    engine.setCellFormula('Sheet1', 'E1', 'ROWS(B1#)')
    engine.setCellFormula('Sheet1', 'F1', 'IFERROR(INDEX(B1#,2),"missing")')

    engine.setCellValue('Sheet1', 'B2', 99)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Spill })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Spill })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({ tag: ValueTag.String, value: 'missing' })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 1, cols: 1 }])

    engine.clearCell('Sheet1', 'B2')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 3, cols: 1 }])
  })

  it('rematerializes dynamic arrays when structural row edits cut through spill children', async () => {
    const buildEngine = async (): Promise<SpreadsheetEngine> => {
      const engine = new SpreadsheetEngine({ workbookName: 'spec' })
      await engine.ready()
      engine.createSheet('Sheet1')
      engine.setCellValue('Sheet1', 'A1', 3)
      engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(A1,1,1,1)')
      engine.setCellFormula('Sheet1', 'D1', 'SUM(B1#)')
      engine.setCellFormula('Sheet1', 'E1', 'ROWS(B1#)')
      engine.setCellFormula('Sheet1', 'F1', 'IFERROR(INDEX(B1#,2),"missing")')
      return engine
    }
    const expectRematerialized = (engine: SpreadsheetEngine): void => {
      expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Empty })
      expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 3, cols: 1 }])
    }

    const inserted = await buildEngine()
    inserted.insertRows('Sheet1', 1, 1)
    expectRematerialized(inserted)

    const deleted = await buildEngine()
    deleted.deleteRows('Sheet1', 1, 1)
    expectRematerialized(deleted)
  })

  it('rematerializes horizontal dynamic arrays when structural column edits cut through spill children', async () => {
    const buildEngine = async (): Promise<SpreadsheetEngine> => {
      const engine = new SpreadsheetEngine({ workbookName: 'spec' })
      await engine.ready()
      engine.createSheet('Sheet1')
      engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(1,3,1,1)')
      engine.setCellFormula('Sheet1', 'A3', 'SUM(B1#)')
      engine.setCellFormula('Sheet1', 'A4', 'COLUMNS(B1#)')
      engine.setCellFormula('Sheet1', 'A5', 'IFERROR(INDEX(B1#,1,2),"missing")')
      return engine
    }
    const expectRematerialized = (engine: SpreadsheetEngine): void => {
      expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
      expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Empty })
      expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 6 })
      expect(engine.getCellValue('Sheet1', 'A4')).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(engine.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B1', rows: 1, cols: 3 }])
    }

    const inserted = await buildEngine()
    inserted.insertColumns('Sheet1', 2, 1)
    expectRematerialized(inserted)

    const deleted = await buildEngine()
    deleted.deleteColumns('Sheet1', 2, 1)
    expectRematerialized(deleted)
  })

  it('rewrites spill-reference consumers when moving dynamic-array owner rows', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(A1,1,1,1)')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(B1#)')
    engine.setCellFormula('Sheet1', 'E1', 'ROWS(B1#)')
    engine.setCellFormula('Sheet1', 'F1', 'IFERROR(INDEX(B1#,2),"missing")')

    engine.moveRows('Sheet1', 0, 1, 2)

    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SEQUENCE(A3,1,1,1)')
    expect(engine.getCell('Sheet1', 'D3').formula).toBe('SUM(B3#)')
    expect(engine.getCell('Sheet1', 'E3').formula).toBe('ROWS(B3#)')
    expect(engine.getCell('Sheet1', 'F3').formula).toBe('IFERROR(INDEX(B3#,2),"missing")')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'E3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'B3', rows: 3, cols: 1 }])
  })

  it('rewrites spill-reference consumers when moving dynamic-array owner columns', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'SEQUENCE(1,3,1,1)')
    engine.setCellFormula('Sheet1', 'A3', 'SUM(B1#)')
    engine.setCellFormula('Sheet1', 'A4', 'COLUMNS(B1#)')
    engine.setCellFormula('Sheet1', 'A5', 'IFERROR(INDEX(B1#,1,2),"missing")')

    engine.moveColumns('Sheet1', 1, 1, 2)

    expect(engine.getCell('Sheet1', 'C1').formula).toBe('SEQUENCE(1,3,1,1)')
    expect(engine.getCell('Sheet1', 'A3').formula).toBe('SUM(C1#)')
    expect(engine.getCell('Sheet1', 'A4').formula).toBe('COLUMNS(C1#)')
    expect(engine.getCell('Sheet1', 'A5').formula).toBe('IFERROR(INDEX(C1#,1,2),"missing")')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'A5')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'C1', rows: 1, cols: 3 }])
  })

  it('keeps two-dimensional spill consumers valid when moving the owner column', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B2', 'SEQUENCE(2,3,1,1)')
    engine.setCellFormula('Sheet1', 'G2', 'SUM(B2#)')
    engine.setCellFormula('Sheet1', 'G3', 'ROWS(B2#)')
    engine.setCellFormula('Sheet1', 'G4', 'COLUMNS(B2#)')
    engine.setCellFormula('Sheet1', 'G5', 'IFERROR(INDEX(B2#,2,2),"missing")')

    engine.moveColumns('Sheet1', 1, 1, 3)

    expect(engine.getCell('Sheet1', 'D2').formula).toBe('SEQUENCE(2,3,1,1)')
    expect(engine.getCell('Sheet1', 'G2').formula).toBe('SUM(D2#)')
    expect(engine.getCell('Sheet1', 'G3').formula).toBe('ROWS(D2#)')
    expect(engine.getCell('Sheet1', 'G4').formula).toBe('COLUMNS(D2#)')
    expect(engine.getCell('Sheet1', 'G5').formula).toBe('IFERROR(INDEX(D2#,2,2),"missing")')
    expect(['D2', 'E2', 'F2', 'D3', 'E3', 'F3'].map((address) => engine.getCellValue('Sheet1', address))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 21 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G5')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'D2', rows: 2, cols: 3 }])
  })

  it('evaluates nested sequence aggregates on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'C1', 'AVG(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'D1', 'MIN(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'E1', 'MAX(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'F1', 'COUNT(SEQUENCE(A1,1,1,1))')
    engine.setCellFormula('Sheet1', 'G1', 'COUNTA(SEQUENCE(A1,1,1,1))')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('evaluates unsupported wasm formulas through the JS runtime fallback', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', '4+1')
    engine.setCellFormula('Sheet1', 'B1', 'FORMULATEXT(A1)')

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: '=4+1',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates LET through the wasm fast path after rewrite-based lowering', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'LET(x,2,x+3)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates reference metadata functions through the JS runtime fallback', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Summary')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellValue('Sheet1', 'A2', 'text')
    engine.setCellFormula('Sheet2', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'ROW()')
    engine.setCellFormula('Sheet1', 'D1', 'COLUMN()')
    engine.setCellFormula('Sheet1', 'E1', 'FORMULATEXT(Sheet2!B1)')
    engine.setCellFormula('Sheet1', 'F1', 'SHEET()')
    engine.setCellFormula('Sheet1', 'G1', 'SHEETS()')
    engine.setCellFormula('Sheet1', 'H1', 'CELL("address",B3)')
    engine.setCellFormula('Sheet1', 'I1', 'CELL("contents",A1)')
    engine.setCellFormula('Sheet1', 'J1', 'CELL("type",A2)')
    engine.setCellFormula('Sheet1', 'K1', 'MID(CELL("filename",F4),FIND("]",CELL("filename",F4))+1,99)')
    engine.setCellFormula('Sheet1', 'L1', 'ISFORMULA(Sheet2!B1)')
    engine.setCellFormula('Sheet1', 'M1', 'ISFORMULA(A1)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.String,
      value: '=A1*2',
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.String,
      value: '$B$3',
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.String,
      value: 'l',
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Sheet1',
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.JsOnly)
  })

  it('routes TEXTSPLIT, EXPAND, and TRIMRANGE through wasm while keeping indirection helpers on JS', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 'red,blue|green')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'L2', 1)
    engine.setCellValue('Sheet1', 'M2', 2)
    engine.setCellValue('Sheet1', 'L3', 3)
    engine.setCellFormula('Sheet2', 'A1', 'B1*2')
    engine.setCellFormula('Sheet1', 'C1', 'TEXTSPLIT(A1,",","|")')
    engine.setCellFormula('Sheet1', 'E1', 'EXPAND(B1:B2,3,2,0)')
    engine.setCellFormula('Sheet1', 'G1', 'INDIRECT("B1:B2")')
    engine.setCellFormula('Sheet1', 'H1', 'INDIRECT("B2")')
    engine.setCellFormula('Sheet1', 'I1', 'FORMULA(Sheet2!A1)')
    engine.setCellFormula('Sheet1', 'K6', 'TRIMRANGE(K1:N4)')

    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.String,
      value: 'red',
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.String,
      value: 'blue',
    })
    expect(engine.getCellValue('Sheet1', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: 'green',
    })
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'E3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.String,
      value: '=B1*2',
    })
    expect(engine.getCellValue('Sheet1', 'K6')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L6')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'K7')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'L7')).toEqual({ tag: ValueTag.Empty })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'K6').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes DATEDIF and financial scalar helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'DATEDIF(DATE(2020,1,15),DATE(2021,3,20),"YM")')
    engine.setCellFormula('Sheet1', 'B1', 'DAYS360(DATE(2024,1,29),DATE(2024,3,31))')
    engine.setCellFormula('Sheet1', 'C1', 'DAYS360(DATE(2024,1,29),DATE(2024,3,31),TRUE)')
    engine.setCellFormula('Sheet1', 'D1', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)')
    engine.setCellFormula('Sheet1', 'E1', 'FVSCHEDULE(1000,0.09,0.11,0.1)')
    engine.setCellFormula('Sheet1', 'F1', 'DB(10000,1000,5,1)')
    engine.setCellFormula('Sheet1', 'G1', 'DDB(2400,300,10,2)')
    engine.setCellFormula('Sheet1', 'H1', 'VDB(2400,300,10,1,3)')
    engine.setCellFormula('Sheet1', 'I1', 'SLN(10000,1000,9)')
    engine.setCellFormula('Sheet1', 'J1', 'SYD(10000,1000,9,1)')
    engine.setCellFormula('Sheet1', 'K1', 'DISC(DATE(2023,1,1),DATE(2023,4,1),97,100,2)')
    engine.setCellFormula('Sheet1', 'L1', 'INTRATE(DATE(2023,1,1),DATE(2023,4,1),1000,1030,2)')
    engine.setCellFormula('Sheet1', 'M1', 'RECEIVED(DATE(2023,1,1),DATE(2023,4,1),1000,0.12,2)')
    engine.setCellFormula('Sheet1', 'N1', 'PRICEDISC(DATE(2008,2,16),DATE(2008,3,1),0.0525,100,2)')
    engine.setCellFormula('Sheet1', 'O1', 'YIELDDISC(DATE(2008,2,16),DATE(2008,3,1),99.795,100,2)')
    engine.setCellFormula('Sheet1', 'P1', 'TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.09)')
    engine.setCellFormula('Sheet1', 'Q1', 'TBILLYIELD(DATE(2008,3,31),DATE(2008,6,1),98.45)')
    engine.setCellFormula('Sheet1', 'R1', 'TBILLEQ(DATE(2008,3,31),DATE(2008,6,1),0.0914)')
    engine.setCellFormula('Sheet1', 'S1', 'PRICEMAT(DATE(2008,2,15),DATE(2008,4,13),DATE(2007,11,11),0.061,0.061,0)')
    engine.setCellFormula('Sheet1', 'T1', 'YIELDMAT(DATE(2008,3,15),DATE(2008,11,3),DATE(2007,11,8),0.0625,100.0123,0)')
    engine.setCellFormula(
      'Sheet1',
      'U1',
      'ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)',
    )
    engine.setCellFormula('Sheet1', 'V1', 'ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,84.5,100,2,0)')
    engine.setCellFormula('Sheet1', 'W1', 'ODDLPRICE(DATE(2008,2,7),DATE(2008,6,15),DATE(2007,10,15),0.0375,0.0405,100,2,0)')
    engine.setCellFormula('Sheet1', 'X1', 'ODDLYIELD(DATE(2008,4,20),DATE(2008,6,15),DATE(2007,12,24),0.0375,99.875,100,2,0)')
    engine.setCellFormula('Sheet1', 'Y1', 'EFFECT(0.12,12)')
    engine.setCellFormula('Sheet1', 'Z1', 'NOMINAL(0.12682503013196977,12)')
    engine.setCellFormula('Sheet1', 'AA1', 'PDURATION(0.1,100,121)')
    engine.setCellFormula('Sheet1', 'AB1', 'RRI(2,100,121)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Number,
      value: 62,
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: 61,
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(182 / 365, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1330.89, 12),
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3690, 12),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(384, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(691.2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 1000 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 1800 })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'M1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1030.9278350515465, 12),
    })
    expect(engine.getCellValue('Sheet1', 'N1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.79583333333333, 12),
    })
    expect(engine.getCellValue('Sheet1', 'O1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.05282257198685834, 12),
    })
    expect(engine.getCellValue('Sheet1', 'P1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(98.45, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09141696292534264, 12),
    })
    expect(engine.getCellValue('Sheet1', 'R1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09415149356594302, 12),
    })
    expect(engine.getCellValue('Sheet1', 'S1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.98449887555694, 12),
    })
    expect(engine.getCellValue('Sheet1', 'T1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.060954333691538576, 12),
    })
    expect(engine.getCellValue('Sheet1', 'U1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(113.597717474079, 12),
    })
    expect(engine.getCellValue('Sheet1', 'V1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0772455415972989, 11),
    })
    expect(engine.getCellValue('Sheet1', 'W1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.8782860147213, 12),
    })
    expect(engine.getCellValue('Sheet1', 'X1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0451922356291692, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Y1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12682503013196977, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Z1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(engine.getCellValue('Sheet1', 'AA1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'AB1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'N1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'O1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'P1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Q1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'R1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'S1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'T1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'U1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'V1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'W1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'X1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Y1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'Z1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'AA1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'AB1').mode).toBe(FormulaMode.WasmFastPath)
  })
})
