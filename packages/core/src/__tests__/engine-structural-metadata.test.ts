import type { EngineOpBatch } from './engine-test-helpers.js'
import {
  ErrorCode,
  FormulaMode,
  SpreadsheetEngine,
  ValueTag,
  afterEach,
  describe,
  expect,
  it,
  readRuntimeDirectScalar,
  readRuntimeFormula,
  readRuntimeTemplateId,
  vi,
} from './engine-test-helpers.js'

describe('SpreadsheetEngine structural metadata and references', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('skips no-op defined-name, sort, and table writes and reports missing clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const outbound: EngineOpBatch[] = []
    const unsubscribeBatches = engine.subscribeBatches((batch) => outbound.push(batch))

    expect(engine.deleteDefinedName('MissingRate')).toBe(false)
    expect(
      engine.clearSort('Sheet1', {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      }),
    ).toBe(false)
    expect(engine.getCellNumberFormat(undefined)).toMatchObject({
      kind: 'general',
      code: 'general',
    })
    expect(engine.getWorkbookMetadata('locale')).toBeUndefined()
    expect(engine.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
    })

    engine.setDefinedName('Rate', 0.1)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'upsertDefinedName', name: 'Rate', value: 0.1 }])
    outbound.splice(0)

    engine.setDefinedName('Rate', 0.1)
    expect(outbound).toEqual([])

    const sortRange = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } as const
    const sortKeys = [{ keyAddress: 'B1', direction: 'asc' as const }]
    engine.setSort('Sheet1', sortRange, sortKeys)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'setSort', sheetName: 'Sheet1', range: sortRange, keys: sortKeys }])
    outbound.splice(0)

    engine.setSort('Sheet1', sortRange, sortKeys)
    expect(outbound).toEqual([])

    engine.setWorkbookMetadata('locale', 'en-US')
    expect(engine.getWorkbookMetadata('locale')).toEqual({ key: 'locale', value: 'en-US' })
    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    outbound.splice(0)

    engine.setWorkbookMetadata('locale', 'en-US')
    expect(outbound).toEqual([])

    engine.setCalculationSettings({ mode: 'automatic' })
    expect(outbound).toEqual([])

    engine.setCalculationSettings({ iterate: true, iterateCount: 32, iterateDelta: '0.01', calcOnSave: true, calcCompleted: false })
    expect(outbound.at(-1)?.ops).toEqual([
      {
        kind: 'setCalculationSettings',
        settings: {
          mode: 'automatic',
          compatibilityMode: 'excel-modern',
          iterate: true,
          iterateCount: 32,
          iterateDelta: '0.01',
          calcOnSave: true,
          calcCompleted: false,
        },
      },
    ])
    expect(engine.getCalculationSettings()).toEqual({
      mode: 'automatic',
      compatibilityMode: 'excel-modern',
      iterate: true,
      iterateCount: 32,
      iterateDelta: '0.01',
      calcOnSave: true,
      calcCompleted: false,
    })
    outbound.splice(0)

    engine.setCalculationSettings({ iterate: true, iterateCount: 32, iterateDelta: '0.01', calcOnSave: true, calcCompleted: false })
    expect(outbound).toEqual([])

    const table = {
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B2',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    } as const
    engine.setTable(table)
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'upsertTable', table }])
    expect(engine.getTable('Sales')).toEqual(table)
    expect(engine.getTables()).toEqual([table])
    outbound.splice(0)

    engine.setTable({ ...table, columnNames: [...table.columnNames] })
    expect(outbound).toEqual([])

    expect(engine.deleteTable('MissingTable')).toBe(false)

    engine.setSpillRange('Sheet1', 'D4', 2, 3)
    expect(engine.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'D4', rows: 2, cols: 3 }])
    outbound.splice(0)

    engine.setSpillRange('Sheet1', 'D4', 2, 3)
    expect(outbound).toEqual([])
    expect(engine.deleteSpillRange('Sheet1', 'Z9')).toBe(false)
    expect(engine.deletePivotTable('Sheet1', 'A1')).toBe(false)

    engine.setPivotTable('Sheet1', 'F1', {
      name: 'SalesPivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })
    expect(engine.getPivotTable('Sheet1', 'F1')).toMatchObject({
      name: 'SalesPivot',
      sheetName: 'Sheet1',
      address: 'F1',
    })
    expect(engine.getPivotTables()).toHaveLength(1)

    outbound.splice(0)
    engine.clearCell('Sheet1', 'A1')
    expect(outbound.at(-1)?.ops).toEqual([{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }])

    outbound.splice(0)
    unsubscribeBatches()
    engine.setWorkbookMetadata('timezone', 'UTC')
    expect(outbound).toEqual([])
  })

  it('reads and clears data validations through the direct engine helpers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'data-validation-helpers' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const range = {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A2',
    } as const

    engine.setDataValidation({
      range,
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: true,
    })

    expect(engine.getDataValidation('Sheet1', range)).toEqual({
      range,
      rule: {
        kind: 'list',
        values: ['Draft', 'Final'],
      },
      allowBlank: true,
    })
    expect(engine.clearDataValidation('Sheet1', range)).toBe(true)
    expect(engine.getDataValidation('Sheet1', range)).toBeUndefined()
    expect(engine.clearDataValidation('Sheet1', range)).toBe(false)
  })

  it('replicates structural workbook metadata through authoritative op batches', async () => {
    const primary = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'a' })
    const replica = new SpreadsheetEngine({ workbookName: 'spec', replicaId: 'b' })
    await Promise.all([primary.ready(), replica.ready()])

    const outbound: EngineOpBatch[] = []
    primary.subscribeBatches((batch) => outbound.push(batch))

    primary.createSheet('Sheet1')
    outbound.splice(0).forEach((batch) => replica.applyRemoteBatch(batch))

    primary.setWorkbookMetadata('locale', 'en-US')
    primary.updateRowMetadata('Sheet1', 2, 3, 24, false)
    primary.updateColumnMetadata('Sheet1', 1, 2, 120, true)
    primary.setFreezePane('Sheet1', 1, 2)
    primary.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' })
    primary.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }, [{ keyAddress: 'B1', direction: 'desc' }])
    primary.mergeCells({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' })
    primary.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C10',
      columnNames: ['Region', 'Product', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    primary.setSpillRange('Sheet1', 'E1', 2, 3)

    expect(outbound.at(0)?.ops).toEqual([{ kind: 'setWorkbookMetadata', key: 'locale', value: 'en-US' }])
    expect(outbound.at(1)?.ops).toEqual([
      {
        kind: 'updateRowMetadata',
        sheetName: 'Sheet1',
        start: 2,
        count: 3,
        size: 24,
        hidden: false,
      },
    ])
    expect(outbound.at(2)?.ops).toEqual([
      {
        kind: 'updateColumnMetadata',
        sheetName: 'Sheet1',
        start: 1,
        count: 2,
        size: 120,
        hidden: true,
      },
    ])
    expect(outbound.at(3)?.ops).toEqual([{ kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 2 }])
    expect(outbound.at(4)?.ops).toEqual([
      {
        kind: 'setFilter',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(outbound.at(5)?.ops).toEqual([
      {
        kind: 'setSort',
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'desc' }],
      },
    ])
    expect(outbound.at(6)?.ops).toEqual([
      {
        kind: 'mergeCells',
        range: { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' },
      },
    ])
    expect(outbound.at(7)?.ops).toEqual([
      {
        kind: 'upsertTable',
        table: {
          name: 'Sales',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'C10',
          columnNames: ['Region', 'Product', 'Sales'],
          headerRow: true,
          totalsRow: false,
        },
      },
    ])
    expect(outbound.at(8)?.ops).toEqual([
      {
        kind: 'upsertSpillRange',
        sheetName: 'Sheet1',
        address: 'E1',
        rows: 2,
        cols: 3,
      },
    ])

    outbound.forEach((batch) => replica.applyRemoteBatch(batch))

    expect(replica.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(replica.getRowMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 2, count: 3, size: 24, hidden: false }])
    expect(replica.getColumnMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 1, count: 2, size: 120, hidden: true }])
    expect(replica.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 2 })
    expect(replica.listMergeRanges('Sheet1')).toEqual([{ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' }])
    expect(replica.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(replica.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'desc' }],
      },
    ])
    expect(replica.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(replica.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 3 }])
  })

  it('undoes and redoes structural metadata through the transaction log', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setWorkbookMetadata('locale', 'en-US')
    engine.setFreezePane('Sheet1', 1, 1)
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B5',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })

    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })
    expect(engine.getTables()).toHaveLength(1)

    expect(engine.undo()).toBe(true)
    expect(engine.getTables()).toEqual([])

    expect(engine.undo()).toBe(true)
    expect(engine.getFreezePane('Sheet1')).toBeUndefined()

    expect(engine.undo()).toBe(true)
    expect(engine.getWorkbookMetadataEntries()).toEqual([])

    expect(engine.redo()).toBe(true)
    expect(engine.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])

    expect(engine.redo()).toBe(true)
    expect(engine.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 1 })

    expect(engine.redo()).toBe(true)
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B5',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
  })

  it('treats no-op structural, metadata, freeze, and filter updates as stable public operations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.updateRowMetadata('Sheet1', 1, 1, 25, false)
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)
    engine.setFreezePane('Sheet1', 1, 2)
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B5' } as const
    engine.setFilter('Sheet1', range)

    const before = engine.exportSnapshot()

    engine.insertRows('Sheet1', 0, 0)
    engine.deleteRows('Sheet1', 0, 0)
    engine.moveRows('Sheet1', 1, 0, 2)
    engine.moveRows('Sheet1', 1, 1, 1)
    engine.insertColumns('Sheet1', 0, 0)
    engine.deleteColumns('Sheet1', 0, 0)
    engine.moveColumns('Sheet1', 0, 0, 1)
    engine.moveColumns('Sheet1', 0, 1, 0)
    engine.updateRowMetadata('Sheet1', 1, 1, 25, false)
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)
    engine.updateRowMetadata('Sheet1', 5, 1, null, null)
    engine.updateColumnMetadata('Sheet1', 5, 1, null, null)
    engine.setFreezePane('Sheet1', 1, 2)
    engine.setFilter('Sheet1', range)

    expect(engine.exportSnapshot()).toEqual(before)
    expect(engine.clearFreezePane('Sheet1')).toBe(true)
    expect(engine.clearFreezePane('Sheet1')).toBe(false)
    expect(engine.clearFilter('Sheet1', range)).toBe(true)
    expect(engine.clearFilter('Sheet1', range)).toBe(false)
    expect(engine.getFreezePane('Sheet1')).toBeUndefined()
    expect(engine.getFilters('Sheet1')).toEqual([])
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('tracks structural row identities and rewrites formulas for row inserts and moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')
    engine.updateRowMetadata('Sheet1', 1, 1, 30, false)

    const before = engine.getRowAxisEntries('Sheet1')
    expect(before).toEqual([{ id: 'row-1', index: 1, size: 30, hidden: false }])

    engine.insertRows('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A3)')
    expect(engine.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-2', index: 1 },
      { id: 'row-1', index: 2, size: 30, hidden: false },
    ])

    engine.moveRows('Sheet1', 2, 1, 0)
    expect(engine.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-1', index: 0, size: 30, hidden: false },
      { id: 'row-2', index: 2 },
    ])
  })

  it('preserves adjacent merge ranges across structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'adjacent-merge-structural-rewrite' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.mergeCells({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'K1' })
    engine.mergeCells({ sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'K2' })
    engine.mergeCells({ sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'K3' })

    engine.insertRows('Sheet1', 0, 1)

    expect(engine.listMergeRanges('Sheet1')).toEqual([
      { sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'K2' },
      { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'K3' },
      { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'K4' },
    ])
  })

  it('leaves unsupported defined-name formulas unchanged across structural row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'unsupported-defined-name-structural-rewrite' })
    await engine.ready()
    engine.createSheet('Data')
    engine.setDefinedName('UnsupportedExcelRef', { kind: 'formula', formula: '={"\'excel\'!$A$1:$I$24"}' })

    engine.insertRows('Data', 0, 1)

    expect(engine.getDefinedName('UnsupportedExcelRef')).toEqual({
      name: 'UnsupportedExcelRef',
      value: { kind: 'formula', formula: '={"\'excel\'!$A$1:$I$24"}' },
    })
  })

  it('keeps repeated direct aggregate families correct across structural row transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-aggregate-rows' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `SUM(A1:A${row})`)
    }

    engine.insertRows('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('SUM(A1:A1)')
    expect(engine.getCell('Sheet1', 'B3').formula).toBe('SUM(A1:A3)')
    expect(engine.getCell('Sheet1', 'B5').formula).toBe('SUM(A1:A5)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B5')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.deleteRows('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `B${row}`).formula).toBe(`SUM(A1:A${row})`)
      expect(engine.getCellValue('Sheet1', `B${row}`)).toEqual({
        tag: ValueTag.Number,
        value: (row * (row + 1)) / 2,
      })
    }
  })

  it('retargets direct aggregate region subscriptions across structural column inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-aggregate-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [3]])
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:A3)')

    engine.insertColumns('Sheet1', 0, 1)

    expect(engine.getCell('Sheet1', 'E1').formula).toBe('SUM(B1:B3)')
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.setCellValue('Sheet1', 'B2', 20)

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 24 })
  })

  it('rewrites metadata-backed ranges, names, freeze panes, and pivot sources across structural row edits', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 7],
      ['East', 5],
    ])
    engine.setDefinedName('SalesRange', '=Data!A1:B4')
    engine.setFreezePane('Data', 1, 0)
    engine.setFilter('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' })
    engine.setSort('Data', { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesPivot',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' },
      groupBy: ['Region'],
      values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
    })

    engine.insertRows('Data', 0, 1)

    expect(engine.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: '=Data!A2:B5',
    })
    expect(engine.getFreezePane('Data')).toEqual({ sheetName: 'Data', rows: 2, cols: 0 })
    expect(engine.getFilters('Data')).toEqual([{ sheetName: 'Data', range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' } }])
    expect(engine.getSorts('Data')).toEqual([
      {
        sheetName: 'Data',
        range: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' },
        keys: [{ keyAddress: 'B2', direction: 'asc' }],
      },
    ])
    expect(engine.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Data',
        startAddress: 'A2',
        endAddress: 'B5',
        columnNames: ['Region', 'Sales'],
        headerRow: true,
        totalsRow: false,
      },
    ])
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesPivot',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A2', endAddress: 'B5' },
        cacheFields: ['Region', 'Sales'],
        cachedRecords: [
          ['East', 10],
          ['West', 7],
          ['East', 5],
        ],
        groupBy: ['Region'],
        values: [{ sourceColumn: 'Sales', summarizeBy: 'sum' }],
        rows: 3,
        cols: 2,
      },
    ])

    engine.deleteRows('Data', 0, 5)

    expect(engine.getDefinedName('SalesRange')).toEqual({ name: 'SalesRange', value: '=#REF!' })
    expect(engine.getFreezePane('Data')).toBeUndefined()
    expect(engine.getFilters('Data')).toEqual([])
    expect(engine.getSorts('Data')).toEqual([])
    expect(engine.getTables()).toEqual([])
    expect(engine.getPivotTables()).toEqual([])
    expect(engine.getCellValue('Pivot', 'B2')).toEqual({ tag: ValueTag.Empty })
  })

  it('recalculates named, table, and direct cross-sheet formulas after structural row deletes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B4' }, [
      ['Region', 'Sales'],
      ['East', 10],
      ['West', 7],
      ['North', 5],
    ])
    engine.setDefinedName('SalesRange', '=Data!B2:B4')
    engine.setTable({
      name: 'Sales',
      sheetName: 'Data',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Sales'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!B2:B4)')
    engine.setCellFormula('Summary', 'A2', 'SUM(Sales[Sales])')
    engine.setCellFormula('Summary', 'A3', 'SUM(SalesRange)')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 22 })

    engine.deleteRows('Data', 2, 1)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('recalculates positional cross-sheet formulas after structural row moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A4' }, [[10], [20], [30], [40]])
    engine.setCellFormula('Summary', 'A1', 'INDEX(Data!A1:A4,1)')
    engine.setCellFormula('Summary', 'A2', 'INDEX(Data!A1:A4,2)')
    engine.setCellFormula('Summary', 'A3', 'SUM(Data!A1:A4)')

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 100 })

    engine.moveRows('Data', 2, 1, 0)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Summary', 'A3')).toEqual({ tag: ValueTag.Number, value: 100 })
  })

  it('undoes structural row deletes without losing cross-sheet formula correctness', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-delete-undo-cross-sheet' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Summary')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'A4' }, [[10], [20], [30], [40]])
    engine.setCellFormula('Summary', 'A1', 'SUM(Data!A1:A4)')
    engine.setCellFormula('Summary', 'A2', 'INDEX(Data!A1:A4,2)')

    engine.deleteRows('Data', 1, 1)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 80 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 30 })

    expect(engine.undo()).toBe(true)

    expect(engine.getCellValue('Summary', 'A1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Summary', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Data', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
  })

  it('rewrites formulas for structural column inserts and roundtrips calc settings metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:B1)')
    engine.setCalculationSettings({ mode: 'manual', compatibilityMode: 'odf-1.4', calcOnSave: true, calcCompleted: false })

    engine.insertColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:C1)')

    engine.recalculateNow()
    expect(engine.exportSnapshot().workbook.metadata?.calculationSettings).toEqual({
      mode: 'manual',
      compatibilityMode: 'odf-1.4',
      calcOnSave: true,
      calcCompleted: false,
    })
    expect(engine.exportSnapshot().workbook.metadata?.volatileContext?.recalcEpoch).toBeGreaterThan(0)

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(engine.exportSnapshot())
    expect(restored.getCalculationSettings()).toEqual({
      mode: 'manual',
      compatibilityMode: 'odf-1.4',
      calcOnSave: true,
      calcCompleted: false,
    })
  })

  it('rewrites formulas and axis identities for structural column deletes and moves', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'E1', 'SUM(A1:B1)')
    engine.updateColumnMetadata('Sheet1', 0, 1, 90, true)

    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 0, size: 90, hidden: true }])

    engine.deleteColumns('Sheet1', 0, 1)
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('SUM(A1:A1)')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([])

    engine.updateColumnMetadata('Sheet1', 1, 1, 110, false)
    engine.setCellFormula('Sheet1', 'D2', 'B1')
    engine.moveColumns('Sheet1', 1, 1, 0)

    expect(engine.getCell('Sheet1', 'D2').formula).toBe('A1')
    expect(engine.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-2', index: 0, size: 110, hidden: false }])
  })

  it('keeps simple cell-reference formula families correct across structural column transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-simple-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'C1*2')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.insertColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'D1').formula).toBe('A1+C1')
    expect(engine.getCell('Sheet1', 'E1').formula).toBe('D1*2')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 10 })

    const restoredAfterInsert = new SpreadsheetEngine({ workbookName: 'structural-simple-columns-restored' })
    await restoredAfterInsert.ready()
    restoredAfterInsert.importSnapshot(engine.exportSnapshot())
    expect(restoredAfterInsert.getCell('Sheet1', 'D1').formula).toBe('A1+C1')
    expect(restoredAfterInsert.getCell('Sheet1', 'E1').formula).toBe('D1*2')

    engine.deleteColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'C1').formula).toBe('A1+B1')
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('C1*2')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })

    engine.resetPerformanceCounters()
    engine.moveColumns('Sheet1', 1, 1, 0)

    expect(engine.getCell('Sheet1', 'C1').formula).toBe('B1+A1')
    expect(engine.getCell('Sheet1', 'D1').formula).toBe('C1*2')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('defers deleted-column ref errors without rebinding simple formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-column-ref-error-deferral' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', true)
    engine.setCellValue('Sheet1', 'B3', 'restore-me')
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    engine.setCellFormula('Sheet1', 'D1', 'C1*2')

    engine.resetPerformanceCounters()
    engine.deleteColumns('Sheet1', 1, 1)

    expect(engine.getCell('Sheet1', 'B1').formula).toBe('A1+#REF!')
    expect(engine.getCell('Sheet1', 'C1').formula).toBe('B1*2')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(engine.getPerformanceCounters()).toMatchObject({
      structuralFormulaImpactCandidates: 0,
      structuralFormulaRebindInputs: 0,
      structuralUndoCapturedCells: 0,
    })

    expect(engine.undo()).toBe(true)
    expect(engine.getPerformanceCounters().structuralUndoCapturedCells).toBe(3)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.String, value: 'restore-me', stringId: expect.any(Number) })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('keeps repeated row-shifted formula families correct across structural column transforms', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'structural-column-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 2)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `C${row}*2`)
    }

    const planIdsBeforeInsert = new Map<string, number>()
    const templateIdsBeforeInsert = new Map<string, number | undefined>()
    const directScalarsBeforeInsert = new Map<string, unknown>()
    for (let row = 1; row <= 4; row += 1) {
      const cIndex = engine.workbook.getCellIndex('Sheet1', `C${row}`)
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      planIdsBeforeInsert.set(`C${row}`, readRuntimeFormula(engine, cIndex!)!.planId)
      planIdsBeforeInsert.set(`D${row}`, readRuntimeFormula(engine, dIndex!)!.planId)
      templateIdsBeforeInsert.set(`C${row}`, readRuntimeTemplateId(engine, cIndex!))
      templateIdsBeforeInsert.set(`D${row}`, readRuntimeTemplateId(engine, dIndex!))
      directScalarsBeforeInsert.set(`C${row}`, readRuntimeDirectScalar(engine, cIndex!))
      directScalarsBeforeInsert.set(`D${row}`, readRuntimeDirectScalar(engine, dIndex!))
    }

    engine.resetPerformanceCounters()
    engine.insertColumns('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `D${row}`).formula).toBe(`A${row}+C${row}`)
      expect(engine.getCell('Sheet1', `E${row}`).formula).toBe(`D${row}*2`)
      expect(engine.getCellValue('Sheet1', `D${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 3,
      })
      expect(engine.getCellValue('Sheet1', `E${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 6,
      })
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      const eIndex = engine.workbook.getCellIndex('Sheet1', `E${row}`)
      expect(readRuntimeFormula(engine, dIndex!)?.planId).toBe(planIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeFormula(engine, eIndex!)?.planId).toBe(planIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeTemplateId(engine, dIndex!)).toBe(templateIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeTemplateId(engine, eIndex!)).toBe(templateIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeDirectScalar(engine, dIndex!)).toBe(directScalarsBeforeInsert.get(`C${row}`))
      expect(readRuntimeDirectScalar(engine, eIndex!)).toBe(directScalarsBeforeInsert.get(`D${row}`))
    }

    engine.resetPerformanceCounters()
    engine.deleteColumns('Sheet1', 1, 1)

    for (let row = 1; row <= 4; row += 1) {
      expect(engine.getCell('Sheet1', `C${row}`).formula).toBe(`A${row}+B${row}`)
      expect(engine.getCell('Sheet1', `D${row}`).formula).toBe(`C${row}*2`)
      expect(engine.getCellValue('Sheet1', `C${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 3,
      })
      expect(engine.getCellValue('Sheet1', `D${row}`)).toEqual({
        tag: ValueTag.Number,
        value: row * 6,
      })
      const cIndex = engine.workbook.getCellIndex('Sheet1', `C${row}`)
      const dIndex = engine.workbook.getCellIndex('Sheet1', `D${row}`)
      expect(readRuntimeFormula(engine, cIndex!)?.planId).toBe(planIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeFormula(engine, dIndex!)?.planId).toBe(planIdsBeforeInsert.get(`D${row}`))
      expect(readRuntimeTemplateId(engine, cIndex!)).toBe(templateIdsBeforeInsert.get(`C${row}`))
      expect(readRuntimeTemplateId(engine, dIndex!)).toBe(templateIdsBeforeInsert.get(`D${row}`))
    }
  })

  it('rebinds multi-name scalar formulas once names exist', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'TaxRate+FeeRate')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    engine.setDefinedName('TaxRate', 0.085)
    engine.setDefinedName('FeeRate', 0.015)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.1 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })
  })

  it('resolves named range formulas through workbook metadata and rebinds dependencies', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 12)
    engine.setCellValue('Sheet1', 'A3', 15)
    engine.setDefinedName('SalesRange', '=Sheet1!A1:A3')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(SalesRange)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 37 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellValue('Sheet1', 'A2', 20)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 45 })

    engine.setDefinedName('SalesRange', '=Sheet1!A1:A2')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 30 })
  })

  it('binds structured table references through table metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Region')
    engine.setCellValue('Sheet1', 'B1', 'Amount')
    engine.setCellValue('Sheet1', 'A2', 'North')
    engine.setCellValue('Sheet1', 'B2', 10)
    engine.setCellValue('Sheet1', 'A3', 'South')
    engine.setCellValue('Sheet1', 'B3', 12)
    engine.setCellValue('Sheet1', 'A4', 'West')
    engine.setCellValue('Sheet1', 'B4', 15)
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B4',
      columnNames: ['Region', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })

    engine.setCellFormula('Sheet1', 'C1', 'SUM(Sales[Amount])')
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 37 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellValue('Sheet1', 'B3', 20)
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 45 })

    engine.setCellFormula('Sheet1', 'D1', 'Sales[Amount]')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('rebinds spill-shape formulas when owner ranges appear and resize', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1#)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(3,1,1,1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.setCellFormula('Sheet1', 'A1', 'SEQUENCE(2,1,1,1)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('materializes pivot tables, refreshes aggregates, and roundtrips snapshot metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
      ['East', 'priority', 'Gizmo', 5],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    expect(engine.getCellValue('Pivot', 'B2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Region',
    })
    expect(engine.getCellValue('Pivot', 'C2')).toMatchObject({
      tag: ValueTag.String,
      value: 'SUM of Sales',
    })
    expect(engine.getCellValue('Pivot', 'D2')).toMatchObject({
      tag: ValueTag.String,
      value: 'Rows',
    })
    expect(engine.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(engine.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Pivot', 'D3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Pivot', 'B4')).toMatchObject({
      tag: ValueTag.String,
      value: 'West',
    })
    expect(engine.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Pivot', 'D4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getPivotTables()).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
        cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
        cachedRecords: [
          ['East', 'priority', 'Widget', 10],
          ['West', 'priority', 'Widget', 7],
          ['East', 'priority', 'Gizmo', 5],
        ],
        groupBy: ['Region'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
        ],
        rows: 3,
        cols: 3,
      },
    ])

    engine.setCellValue('Data', 'D3', 9)

    expect(engine.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 9 })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.pivots).toEqual([
      {
        name: 'SalesByRegion',
        sheetName: 'Pivot',
        address: 'B2',
        source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
        cacheFields: ['Region', 'Notes', 'Product', 'Sales'],
        cachedRecords: [
          ['East', 'priority', 'Widget', 10],
          ['West', 'priority', 'Widget', 9],
          ['East', 'priority', 'Gizmo', 5],
        ],
        groupBy: ['Region'],
        values: [
          { sourceColumn: 'Sales', summarizeBy: 'sum' },
          { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
        ],
        rows: 3,
        cols: 3,
      },
    ])
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Pivot')?.cells).toEqual([])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getCellValue('Pivot', 'B3')).toMatchObject({
      tag: ValueTag.String,
      value: 'East',
    })
    expect(restored.getCellValue('Pivot', 'C3')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(restored.getCellValue('Pivot', 'C4')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(restored.exportSnapshot().workbook.metadata?.pivots).toEqual(snapshot.workbook.metadata?.pivots)
  })

  it('evaluates GETPIVOTDATA against workbook pivot metadata', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Data')
    engine.createSheet('Pivot')
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' }, [
      ['Region', 'Notes', 'Product', 'Sales'],
      ['East', 'priority', 'Widget', 10],
      ['West', 'priority', 'Widget', 7],
      ['East', 'priority', 'Gizmo', 5],
    ])

    engine.setPivotTable('Pivot', 'B2', {
      name: 'SalesByRegion',
      source: { sheetName: 'Data', startAddress: 'A1', endAddress: 'D4' },
      groupBy: ['Region'],
      values: [
        { sourceColumn: 'Sales', summarizeBy: 'sum' },
        { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
      ],
    })

    engine.setCellFormula('Sheet1', 'A1', 'GETPIVOTDATA("Sales",Pivot!B2)')
    engine.setCellFormula('Sheet1', 'A2', 'GETPIVOTDATA("Sales",Pivot!B2,"Region","East")')
    engine.setCellFormula('Sheet1', 'A3', 'GETPIVOTDATA("Rows",Pivot!B2,"Region","West")')
    engine.setCellFormula('Sheet1', 'A4', 'GETPIVOTDATA("Missing",Pivot!B2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(engine.explainCell('Sheet1', 'A2').mode).toBe(FormulaMode.JsOnly)
  })
})
