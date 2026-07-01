import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../index.js'

describe('SpreadsheetEngine snapshot roundtrips', () => {
  it('persists workbook defined names through snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.setDefinedName('TaxRate', 0.085)
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 100)
    engine.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    const snapshot = engine.exportSnapshot()

    expect(snapshot.workbook.metadata?.definedNames).toEqual([{ name: 'TaxRate', value: 0.085 }])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.085 }])
    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Number,
      value: 8.5,
    })
    expect(restored.exportSnapshot().workbook.metadata?.definedNames).toEqual([{ name: 'TaxRate', value: 0.085 }])
  })

  it('supports explicit range-ref and formula defined-name metadata values', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [3]])
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'A3',
    })
    engine.setDefinedName('TaxExpr', { kind: 'formula', formula: '=B1*0.1' })
    engine.setCellFormula('Sheet1', 'C1', 'SUM(SalesRange)+TaxExpr')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'SalesRange',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A3',
        },
      },
      { name: 'TaxExpr', value: { kind: 'formula', formula: '=B1*0.1' } },
    ])

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)
    expect(restored.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
  })

  it('treats invalid and cyclic formula-backed defined names as workbook metadata errors', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'defined-name-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setDefinedName('BrokenExpr', { kind: 'formula', formula: '=1+' })
    engine.setDefinedName('LoopA', { kind: 'formula', formula: '=LoopB' })
    engine.setDefinedName('LoopB', { kind: 'formula', formula: '=LoopA' })

    engine.setCellFormula('Sheet1', 'A1', 'BrokenExpr')
    engine.setCellFormula('Sheet1', 'A2', 'LoopA')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
  })

  it('persists expanded workbook metadata through snapshot roundtrip', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setWorkbookMetadata('locale', 'en-US')
    engine.updateRowMetadata('Sheet1', 2, 2, 24, false)
    engine.updateColumnMetadata('Sheet1', 1, 1, 140, null)
    engine.setFreezePane('Sheet1', 1, 2)
    engine.setFilter('Sheet1', {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C10',
    })
    engine.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }, [{ keyAddress: 'B1', direction: 'asc' }])
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'C10',
      columnNames: ['Region', 'Product', 'Sales'],
      headerRow: true,
      totalsRow: true,
    })
    engine.setSpillRange('Sheet1', 'E1', 2, 2)

    const snapshot = engine.exportSnapshot()

    expect(snapshot.workbook.metadata?.properties).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(snapshot.workbook.metadata?.tables).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: true,
      },
    ])
    expect(snapshot.workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 2 }])
    expect(snapshot.sheets.find((sheet) => sheet.name === 'Sheet1')?.metadata).toEqual({
      rows: [
        { id: 'row-1', index: 2, size: 24, hidden: false },
        { id: 'row-2', index: 3, size: 24, hidden: false },
      ],
      columns: [{ id: 'column-1', index: 1, size: 140 }],
      rowMetadata: [{ start: 2, count: 2, size: 24, hidden: false }],
      columnMetadata: [{ start: 1, count: 1, size: 140 }],
      freezePane: { rows: 1, cols: 2 },
      filters: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' }],
      sorts: [
        {
          range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
          keys: [{ keyAddress: 'B1', direction: 'asc' }],
        },
      ],
    })

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getWorkbookMetadataEntries()).toEqual([{ key: 'locale', value: 'en-US' }])
    expect(restored.getRowAxisEntries('Sheet1')).toEqual([
      { id: 'row-1', index: 2, size: 24, hidden: false },
      { id: 'row-2', index: 3, size: 24, hidden: false },
    ])
    expect(restored.getColumnAxisEntries('Sheet1')).toEqual([{ id: 'column-1', index: 1, size: 140 }])
    expect(restored.getRowMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 2, count: 2, size: 24, hidden: false }])
    expect(restored.getColumnMetadata('Sheet1')).toEqual([{ sheetName: 'Sheet1', start: 1, count: 1, size: 140, hidden: null }])
    expect(restored.getFreezePane('Sheet1')).toEqual({ sheetName: 'Sheet1', rows: 1, cols: 2 })
    expect(restored.getFilters('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
      },
    ])
    expect(restored.getSorts('Sheet1')).toEqual([
      {
        sheetName: 'Sheet1',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'C10' },
        keys: [{ keyAddress: 'B1', direction: 'asc' }],
      },
    ])
    expect(restored.getTables()).toEqual([
      {
        name: 'Sales',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'C10',
        columnNames: ['Region', 'Product', 'Sales'],
        headerRow: true,
        totalsRow: true,
      },
    ])
    expect(restored.getSpillRanges()).toEqual([{ sheetName: 'Sheet1', address: 'E1', rows: 2, cols: 2 }])
    expect(restored.exportSnapshot()).toEqual(snapshot)
  })

  it('roundtrips structurally shifted range-backed defined names through snapshots', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'snapshot-structural-defined-range',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [
      ['Qty', 'Amount'],
      [1, 10],
      [2, 20],
    ])
    engine.setDefinedName('SalesRange', {
      kind: 'range-ref',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
    })
    engine.setTable({
      name: 'Sales',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Qty', 'Amount'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setCellFormula('Sheet1', 'C1', 'SUM(SalesRange)')
    engine.insertColumns('Sheet1', 0, 1)

    const snapshot = engine.exportSnapshot()
    expect(snapshot.workbook.metadata?.definedNames).toEqual([
      {
        name: 'SalesRange',
        value: {
          kind: 'range-ref',
          sheetName: 'Sheet1',
          startAddress: 'B1',
          endAddress: 'C3',
        },
      },
    ])

    const restored = new SpreadsheetEngine({
      workbookName: 'snapshot-structural-defined-range-restored',
    })
    await restored.ready()
    restored.importSnapshot(snapshot)

    expect(restored.getDefinedName('SalesRange')).toEqual({
      name: 'SalesRange',
      value: {
        kind: 'range-ref',
        sheetName: 'Sheet1',
        startAddress: 'B1',
        endAddress: 'C3',
      },
    })
    expect(restored.getCell('Sheet1', 'D1').formula).toBe('SUM(SalesRange)')
    expect(restored.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.Number,
      value: 33,
    })
    expect(restored.exportSnapshot()).toEqual(snapshot)
  })
})
