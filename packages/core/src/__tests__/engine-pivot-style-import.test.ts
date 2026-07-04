import type { EngineEvent } from './engine-test-helpers.js'
import {
  ErrorCode,
  FormulaMode,
  SpreadsheetEngine,
  ValueTag,
  afterEach,
  describe,
  expect,
  it,
  seedPivotSource,
  vi,
} from './engine-test-helpers.js'

describe('SpreadsheetEngine pivots, styles, and import surfaces', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('evaluates GROUPBY, PIVOTBY, and MULTIPLE.OPERATIONS end to end', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D5' }, [
      ['Region', 'Product', 'Sales', 'Include'],
      ['East', 'Widget', 10, true],
      ['West', 'Widget', 7, true],
      ['East', 'Gizmo', 5, true],
      ['West', 'Gizmo', 4, true],
    ])
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)

    engine.setCellFormula('Sheet1', 'F1', 'GROUPBY(A1:A5,C1:C5,SUM,3,1)')
    engine.setCellFormula('Sheet1', 'J1', 'PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)')
    engine.setCellFormula('Sheet1', 'N1', 'MULTIPLE.OPERATIONS(P5,P3,Q4,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: 'Region',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'Sales',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({
      tag: ValueTag.String,
      value: 'East',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 26 })

    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({
      tag: ValueTag.String,
      value: 'Region',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({
      tag: ValueTag.String,
      value: 'Widget',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'J2')).toEqual({
      tag: ValueTag.String,
      value: 'East',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'M4')).toEqual({ tag: ValueTag.Number, value: 26 })

    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 23 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'N1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates row-only MULTIPLE.OPERATIONS substitutions through the JS workbook path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P5,P3,Q4)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('returns literal, empty, and cycle states for MULTIPLE.OPERATIONS target cells', async () => {
    const literalEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-literal' })
    await literalEngine.ready()
    literalEngine.setCellValue('Sheet1', 'B5', 42)
    literalEngine.setCellValue('Sheet1', 'C4', 5)
    literalEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(B5,B3,C4)')
    expect(literalEngine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })

    const missingEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-missing' })
    await missingEngine.ready()
    missingEngine.setCellValue('Sheet1', 'C4', 5)
    missingEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(Z9,B3,C4)')
    expect(missingEngine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    const cycleEngine = new SpreadsheetEngine({ workbookName: 'multiple-ops-cycle' })
    await cycleEngine.ready()
    cycleEngine.setCellFormula('Sheet1', 'P2', 'P3')
    cycleEngine.setCellFormula('Sheet1', 'P3', 'P2')
    cycleEngine.setCellValue('Sheet1', 'Q4', 5)
    cycleEngine.setCellValue('Sheet1', 'R4', 9)
    cycleEngine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P2,Q4,R4)')
    expect(cycleEngine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('applies MULTIPLE.OPERATIONS replacements through ranged formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'multiple-ops-range' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P6', 'SUM(P2:P4)')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P6,P3,Q4,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('evaluates nested MULTIPLE.OPERATIONS formulas through the workbook callback path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'multiple-ops-nested' })
    await engine.ready()
    engine.setCellValue('Sheet1', 'P2', 1)
    engine.setCellValue('Sheet1', 'P3', 2)
    engine.setCellFormula('Sheet1', 'P4', 'P2+P3')
    engine.setCellFormula('Sheet1', 'P5', 'P2*P3+P4')
    engine.setCellValue('Sheet1', 'Q4', 5)
    engine.setCellValue('Sheet1', 'R2', 3)
    engine.setCellFormula('Sheet1', 'P7', 'MULTIPLE.OPERATIONS(P5,P3,Q4)')
    engine.setCellFormula('Sheet1', 'A1', 'MULTIPLE.OPERATIONS(P7,P2,R2)')

    expect(engine.getCellValue('Sheet1', 'P7')).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.JsOnly)
  })

  it('undoes pivot deletion through the transaction log', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D3' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B3')).toEqual({ tag: ValueTag.Empty })

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTables()).toHaveLength(1)
    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
  })

  it('returns #VALUE for pivots whose configured headers are missing', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])

    engine.setPivotTable('Pivot', 'A1', {
      name: 'BrokenPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Missing'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('returns #REF for missing pivot source sheets and rebinds once source cells appear', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Pivot')

    engine.setPivotTable('Pivot', 'A1', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])

    expect(engine.getCellValue('Pivot', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'B1')).toMatchObject({
      tag: ValueTag.String,
      value: 'SUM of Sales',
    })
    expect(engine.getCellValue('Pivot', 'A2')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Pivot', 'A3')).toMatchObject({
      tag: ValueTag.String,
      value: 'West',
    })
    expect(engine.getCellValue('Pivot', 'B3')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('blocks overlapping pivot outputs and deletes pivots when users overwrite pivot cells', async () => {
    const blockedByValue = new SpreadsheetEngine({ workbookName: 'pivot-blocked-value' })
    await blockedByValue.ready()
    seedPivotSource(blockedByValue)
    blockedByValue.setCellValue('Pivot', 'C3', 99)
    blockedByValue.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByValue.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByValue.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 99 })

    const blockedByFormula = new SpreadsheetEngine({ workbookName: 'pivot-blocked-formula' })
    await blockedByFormula.ready()
    seedPivotSource(blockedByFormula)
    blockedByFormula.setCellFormula('Pivot', 'C3', '1+1')
    blockedByFormula.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByFormula.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByFormula.getCellValue('Pivot', 'C3')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    const blockedBySpillChild = new SpreadsheetEngine({ workbookName: 'pivot-blocked-spill' })
    await blockedBySpillChild.ready()
    seedPivotSource(blockedBySpillChild)
    blockedBySpillChild.setCellFormula('Pivot', 'C2', 'SEQUENCE(2,1,1,1)')
    blockedBySpillChild.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedBySpillChild.getCellValue('Pivot', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedBySpillChild.getCellValue('Pivot', 'C3')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })

    const blockedByPivotOwner = new SpreadsheetEngine({ workbookName: 'pivot-blocked-pivot' })
    await blockedByPivotOwner.ready()
    seedPivotSource(blockedByPivotOwner)
    blockedByPivotOwner.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    blockedByPivotOwner.setPivotTable('Pivot', 'A1', {
      name: 'Overlap',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(blockedByPivotOwner.getCellValue('Pivot', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(blockedByPivotOwner.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })

    const overwrittenPivot = new SpreadsheetEngine({ workbookName: 'pivot-overwrite' })
    await overwrittenPivot.ready()
    seedPivotSource(overwrittenPivot)
    overwrittenPivot.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(overwrittenPivot.getPivotTables()).toHaveLength(1)

    overwrittenPivot.setCellValue('Pivot', 'B3', 'manual')

    expect(overwrittenPivot.getPivotTables()).toEqual([])
    expect(overwrittenPivot.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
    expect(overwrittenPivot.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'manual',
    })
    expect(overwrittenPivot.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('explains missing cells and undoes table spill and pivot metadata changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')

    expect(engine.explainCell('Data', 'Z99')).toEqual({
      sheetName: 'Data',
      address: 'Z99',
      value: { tag: ValueTag.Empty },
      flags: 0,
      version: 0,
      inCycle: false,
      directPrecedents: [],
      directDependents: [],
    })

    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 5],
    ])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setSpillRange('Pivot', 'E1', 2, 2)
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: true,
    })
    engine.setSpillRange('Pivot', 'E1', 3, 1)
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: true })
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 3,
      cols: 1,
    })
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.deleteTable('Sales')).toBe(true)
    expect(engine.deleteSpillRange('Pivot', 'E1')).toBe(true)
    expect(engine.deletePivotTable('Pivot', 'B2')).toBe(true)

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Sales', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 3,
      cols: 1,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: true })

    expect(engine.undo()).toBe(true)
    expect(engine.getPivotTable('Pivot', 'B2')).toMatchObject({
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getSpillRanges()).toContainEqual({
      sheetName: 'Pivot',
      address: 'E1',
      rows: 2,
      cols: 2,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getTable('Sales')).toMatchObject({ totalsRow: false })
  })

  it('exports sparse high-row cells without truncating the sheet', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A10002', 7)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toContainEqual({ address: 'A10002', value: 7 })
  })

  it('roundtrips a single sheet through CSV with formulas and quoted strings', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'A2', 'alpha,beta')

    const csv = engine.exportSheetCsv('Sheet1')
    expect(csv).toBe('12,=A1*2\n"alpha,beta",')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSheetCsv('Sheet1', csv)

    expect(restored.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(restored.getCell('Sheet1', 'B1').formula).toBe('A1*2')
    expect(restored.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(restored.getCell('Sheet1', 'A2').value).toEqual({
      tag: ValueTag.String,
      value: 'alpha,beta',
      stringId: 1,
    })
  })

  it('recalculates range formulas over imported formula cells after CSV import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'csv-range-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'SUM(B1:B1)')
    engine.setCellValue('Sheet1', 'B1', 'text:@4yt')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(B2:C4)')
    engine.setCellValue('Sheet1', 'A2', false)
    engine.setCellValue('Sheet1', 'C2', true)
    engine.setCellFormula('Sheet1', 'A3', 'SUM(B1:B1)')
    engine.setCellFormula('Sheet1', 'C3', 'B1+B1')
    engine.setCellFormula('Sheet1', 'A4', 'SUM(B1:B1)')
    engine.setCellValue('Sheet1', 'B4', 'text:"k')
    engine.setCellValue('Sheet1', 'C4', 'text:&Pr!}${')

    const csv = engine.exportSheetCsv('Sheet1')

    const restored = new SpreadsheetEngine({ workbookName: 'csv-range-recalc-restored' })
    await restored.ready()
    restored.importSheetCsv('Sheet1', csv)

    expect(restored.getCellValue('Sheet1', 'C3')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(restored.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('settles formula CSV imports without redundant public recalculation passes', async () => {
    const restored = new SpreadsheetEngine({ workbookName: 'csv-import-single-settle' })
    await restored.ready()
    const recalculateNowSpy = vi.spyOn(restored, 'recalculateNow')

    restored.importSheetCsv('Sheet1', '=B1*2,3\n=B1+1,=A2+A1')

    expect(recalculateNowSpy).not.toHaveBeenCalled()
    expect(restored.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(restored.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('recalculates transitive range dependents when a downstream formula becomes an error', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'transitive-range-recalc' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'C1', 1_835_115_565)
    engine.setCellValue('Sheet1', 'D1', -24)
    engine.setCellFormula('Sheet1', 'A2', 'SUM(B1:E2)')
    engine.setCellFormula('Sheet1', 'B2', 'IF(E2+0>0,"text:yes","text:no")')

    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Number,
      value: 1_835_115_541,
    })

    engine.setCellValue('Sheet1', 'E2', 'text:ooe)ZL#<')

    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('persists cell formats through imperative updates and snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormat('Sheet1', 'A1', 'currency-usd')

    expect(engine.getCell('Sheet1', 'A1').format).toBe('currency-usd')
    expect(engine.explainCell('Sheet1', 'A1').format).toBe('currency-usd')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())

    expect(restored.getCell('Sheet1', 'A1').format).toBe('currency-usd')
    expect(restored.exportSnapshot().sheets[0]?.cells).toContainEqual({
      address: 'A1',
      value: 12,
      format: 'currency-usd',
    })
  })

  it('includes format-only mutations in changed cell events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 12)

    const changed: Array<{ indices: number[]; cells: EngineEvent['changedCells'] }> = []
    const unsubscribe = engine.subscribe((event) => {
      changed.push({
        indices: Array.from(event.changedCellIndices),
        cells: event.changedCells,
      })
    })

    engine.setCellFormat('Sheet1', 'A1', 'currency-usd')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    expect(a1Index).toBeDefined()
    expect(sheetId).toBeDefined()
    expect(changed.at(-1)?.indices).toEqual([a1Index!])
    expect(changed.at(-1)?.cells).toEqual([
      {
        kind: 'cell',
        cellIndex: a1Index!,
        address: { sheet: sheetId!, row: 0, col: 0 },
        sheetName: 'Sheet1',
        a1: 'A1',
        newValue: { tag: ValueTag.Number, value: 12 },
      },
    ])

    unsubscribe()
  })

  it('persists pooled cell styles and style ranges without materializing empty cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
      {
        fill: { backgroundColor: '#ABCDEF' },
        font: { family: 'Fira Sans' },
      },
    )

    expect(engine.workbook.getCellIndex('Sheet1', 'B2')).toBeUndefined()
    const styled = engine.getCell('Sheet1', 'B2')
    expect(styled.styleId).toBeDefined()

    const snapshot = engine.exportSnapshot()
    expect(snapshot.sheets[0]?.cells).toEqual([])
    expect(snapshot.workbook.metadata?.styles).toHaveLength(1)
    expect(snapshot.workbook.metadata?.styles?.[0]).toMatchObject({
      fill: { backgroundColor: '#abcdef' },
      font: { family: 'Fira Sans' },
    })
    expect(snapshot.sheets[0]?.metadata?.styleRanges).toEqual([
      {
        range: { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' },
        styleId: styled.styleId,
      },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getCell('Sheet1', 'C3').styleId).toBe(styled.styleId)
    expect(restored.getCellStyle(styled.styleId)).toMatchObject({
      fill: { backgroundColor: '#abcdef' },
      font: { family: 'Fira Sans' },
    })
  })

  it('emits full invalidation when importing a snapshot', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const events: Array<{
      invalidation: 'cells' | 'full'
      changedCellIndices: number[]
      changedCells: number
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        changedCellIndices: Array.from(event.changedCellIndices),
        changedCells: event.changedCells.length,
      })
    })

    engine.importSnapshot({
      version: 1,
      workbook: { name: 'spec' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'A1', value: 12 }],
        },
      ],
    })

    expect(events.at(-1)).toEqual({
      invalidation: 'full',
      changedCellIndices: [],
      changedCells: 0,
    })
    unsubscribe()
  })

  it('emits tracked full invalidation when importing a raw snapshot directly', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec-tracked-import' })
    await engine.ready()
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    engine.importSnapshot({
      version: 1,
      workbook: { name: 'spec-tracked-import' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [{ address: 'A1', value: 12 }],
        },
      ],
    })

    expect(tracked).toHaveBeenCalledTimes(1)
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
        changedCellIndices: new Uint32Array(),
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        explicitChangedCount: 0,
      }),
    )
    expect(tracked.mock.calls[0]?.[0].patches ?? []).toEqual([])
    expect(tracked.mock.calls[0]?.[0].metrics.batchId).toBe(1)
    unsubscribe()
  })

  it('emits targeted range invalidation for style-only edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const events: Array<{
      invalidation: 'cells' | 'full'
      invalidatedRanges: readonly { sheetName: string; startAddress: string; endAddress: string }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        invalidatedRanges: event.invalidatedRanges,
      })
    })

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }, { fill: { backgroundColor: '#ABCDEF' } })

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }],
    })
    unsubscribe()
  })

  it('interns identical cell styles across ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const patch = {
      fill: { backgroundColor: '#ff0000' },
      font: { family: 'IBM Plex Sans' },
    } as const
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }, patch)
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, patch)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.styles).toHaveLength(1)
    expect(snapshot.sheets[0]?.metadata?.styleRanges).toHaveLength(2)
    expect(snapshot.sheets[0]?.metadata?.styleRanges?.[0]?.styleId).toBe(snapshot.sheets[0]?.metadata?.styleRanges?.[1]?.styleId)
  })

  it('merges and clears style fields independently', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      {
        fill: { backgroundColor: '#ff0000' },
        font: { family: 'Inter' },
      },
    )
    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, { font: { family: 'IBM Plex Sans' } })

    const mergedStyle = engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)
    expect(mergedStyle).toMatchObject({
      fill: { backgroundColor: '#ff0000' },
      font: { family: 'IBM Plex Sans' },
    })

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, ['fontFamily'])
    const clearedStyle = engine.getCellStyle(engine.getCell('Sheet1', 'A1').styleId)
    expect(clearedStyle).toMatchObject({
      fill: { backgroundColor: '#ff0000' },
    })
    expect(clearedStyle?.font).toBeUndefined()
  })

  it('notifies address listeners for style-only edits on empty cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let notified = 0
    const unsubscribe = engine.subscribeCell('Sheet1', 'D4', () => {
      notified += 1
    })

    engine.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'D4' }, { fill: { backgroundColor: '#00ff00' } })

    expect(notified).toBe(1)
    unsubscribe()
  })

  it('persists pooled number formats and sparse format ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeNumberFormat(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B4' },
      { kind: 'accounting', currency: 'USD', decimals: 2, useGrouping: true },
    )

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.formats).toHaveLength(1)
    expect(snapshot.sheets[0]?.metadata?.formatRanges).toHaveLength(1)
    expect(engine.getCell('Sheet1', 'B3').format).toContain('accounting:USD:2')
  })

  it('clears number formats, clears sorts, and tracks existing watched cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)

    let notifications = 0
    const unsubscribe = engine.subscribeCells('Sheet1', ['A1', 'Z9'], () => {
      notifications += 1
    })

    engine.setCellValue('Sheet1', 'A1', 6)
    expect(notifications).toBe(1)
    unsubscribe()

    engine.setRangeNumberFormat(
      { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' },
      { kind: 'currency', currency: 'USD', decimals: 2, useGrouping: true },
    )
    engine.clearRangeNumberFormat({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' })
    expect(engine.getCell('Sheet1', 'B2').format).toBeUndefined()

    const sortRange = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    engine.setSort('Sheet1', sortRange, [{ keyAddress: 'A1', direction: 'asc' }])
    expect(engine.clearSort('Sheet1', sortRange)).toBe(true)
    expect(engine.getSorts('Sheet1')).toEqual([])
    expect(engine.getVolatileContext()).toEqual({ recalcEpoch: 0 })
  })

  it('emits targeted axis invalidation for column metadata edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const events: Array<{
      invalidation: 'cells' | 'full'
      invalidatedColumns: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        invalidatedColumns: event.invalidatedColumns,
      })
    })

    engine.updateColumnMetadata('Sheet1', 2, 2, 120, true)

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      invalidatedColumns: [{ sheetName: 'Sheet1', startIndex: 2, endIndex: 3 }],
    })
    unsubscribe()
  })

  it('emits structural row invalidation without flooding changed cells for row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')

    const events: Array<{
      invalidation: 'cells' | 'full'
      changedCells: number
      invalidatedRows: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        changedCells: event.changedCells.length,
        invalidatedRows: event.invalidatedRows,
      })
    })

    engine.insertRows('Sheet1', 1, 1)

    expect(events.at(-1)).toEqual({
      invalidation: 'cells',
      changedCells: 0,
      invalidatedRows: [{ sheetName: 'Sheet1', startIndex: 1, endIndex: 1 }],
    })
    unsubscribe()
  })

  it('keeps tracked structural no-value-change events lightweight', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tracked-structural-no-value-change' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', '=A1+B1')
    engine.setCellFormula('Sheet1', 'D1', '=C1*2')

    const events: EngineEvent[] = []
    const unsubscribe = engine.events.subscribeTracked((event) => {
      events.push(event)
    })

    engine.insertColumns('Sheet1', 1, 1)

    const event = events.at(-1)
    expect(event?.changedCellIndices).toHaveLength(0)
    expect(event?.invalidatedColumns).toEqual([{ sheetName: 'Sheet1', startIndex: 1, endIndex: 1 }])
    expect(event?.patches).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })
    unsubscribe()
  })

  it('emits large structural column invalidations without flooding changed cell payloads for deletes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-delete-column-event' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 260; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    const events: Array<{
      invalidation: 'cells' | 'full'
      changedCellIndices: number[]
      changedCells: number
      invalidatedColumns: readonly { sheetName: string; startIndex: number; endIndex: number }[]
    }> = []
    const unsubscribe = engine.subscribe((event) => {
      events.push({
        invalidation: event.invalidation,
        changedCellIndices: Array.from(event.changedCellIndices),
        changedCells: event.changedCells.length,
        invalidatedColumns: event.invalidatedColumns,
      })
    })

    engine.deleteColumns('Sheet1', 1, 1)

    expect(events.at(-1)?.invalidation).toBe('cells')
    expect(events.at(-1)?.changedCellIndices.length).toBeGreaterThan(512)
    expect(events.at(-1)?.changedCells).toBe(0)
    expect(events.at(-1)?.invalidatedColumns).toEqual([{ sheetName: 'Sheet1', startIndex: 1, endIndex: 1 }])
    unsubscribe()
  })

  it('merges advanced style fields including borders and font weight', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'C5', endAddress: 'C5' },
      {
        font: { bold: true, color: '#111827', size: 14 },
        alignment: { horizontal: 'right', wrap: true },
        borders: {
          bottom: { style: 'double', weight: 'medium', color: '#111827' },
        },
      },
    )

    const style = engine.getCellStyle(engine.getCell('Sheet1', 'C5').styleId)
    expect(style).toMatchObject({
      font: { bold: true, color: '#111827', size: 14 },
      alignment: { horizontal: 'right', wrap: true },
      borders: {
        bottom: { style: 'double', weight: 'medium', color: '#111827' },
      },
    })
  })

  it('removes style subfields through null patches and clearing all style fields', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' },
      {
        font: { family: 'Inter', bold: true },
        alignment: { horizontal: 'center', wrap: true, indent: 2 },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111111' },
          right: { style: 'double', weight: 'medium', color: '#222222' },
        },
      },
    )
    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' },
      {
        alignment: { horizontal: null, wrap: null },
        borders: {
          top: null,
          right: { style: 'solid', weight: 'thin', color: null },
        },
      },
    )

    const partiallyCleared = engine.getCellStyle(engine.getCell('Sheet1', 'D6').styleId)
    expect(partiallyCleared).toMatchObject({
      font: { family: 'Inter', bold: true },
      alignment: { indent: 2 },
    })
    expect(partiallyCleared?.borders).toBeUndefined()

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'D6', endAddress: 'D6' })
    expect(engine.getCell('Sheet1', 'D6').styleId).toBeUndefined()
  })

  it('preserves sibling style fields when clearing only part of a section', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'E7', endAddress: 'E7' },
      {
        font: { family: 'Inter', bold: true },
        alignment: { horizontal: 'right', wrap: true },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111111' },
          left: { style: 'double', weight: 'medium', color: '#222222' },
        },
      },
    )

    engine.clearRangeStyle({ sheetName: 'Sheet1', startAddress: 'E7', endAddress: 'E7' }, ['fontBold', 'alignmentWrap', 'borderTop'])

    const style = engine.getCellStyle(engine.getCell('Sheet1', 'E7').styleId)
    expect(style).toMatchObject({
      font: { family: 'Inter' },
      alignment: { horizontal: 'right' },
      borders: {
        left: { style: 'double', weight: 'medium', color: '#222222' },
      },
    })
  })

  it('replaces existing sheet contents on CSV import', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'C3', 9)

    engine.importSheetCsv('Sheet1', '7,8')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Empty })
  })

  it('explains formula cells with mode, version, and dependencies', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const explanation = engine.explainCell('Sheet1', 'B1')

    expect(explanation.formula).toBe('A1*2')
    expect(explanation.mode).toBeDefined()
    expect(explanation.version).toBeGreaterThan(0)
    expect(explanation.directPrecedents).toEqual(['Sheet1!A1'])
    expect(explanation.directDependents).toEqual([])
    expect(explanation.inCycle).toBe(false)
  })
})
