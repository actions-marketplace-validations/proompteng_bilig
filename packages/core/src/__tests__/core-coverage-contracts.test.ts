import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { WorkbookActionPlan } from '@bilig/workbook'
import { SpreadsheetEngine } from '../engine.js'

function expectNumber(engine: SpreadsheetEngine, address: string, value: number): void {
  expect(engine.getCellValue('Model', address)).toEqual({ tag: ValueTag.Number, value })
}

function seedModelWorkbook(workbookName: string): SpreadsheetEngine {
  const engine = new SpreadsheetEngine({ workbookName, useColumnIndex: true })
  engine.createSheet('Model')
  engine.createSheet('Lookup')
  engine.setCellValue('Model', 'A1', 'Region')
  engine.setCellValue('Model', 'B1', 'Units')
  engine.setCellValue('Model', 'C1', 'Price')
  engine.setCellValue('Model', 'D1', 'Revenue')
  const rows = [
    ['West', 4, 10],
    ['East', 8, 12],
    ['West', 3, 9],
    ['North', 5, 11],
  ] as const
  rows.forEach(([region, units, price], index) => {
    const row = index + 2
    engine.setCellValue('Model', `A${row}`, region)
    engine.setCellValue('Model', `B${row}`, units)
    engine.setCellValue('Model', `C${row}`, price)
    engine.setCellFormula('Model', `D${row}`, `B${row}*C${row}`)
  })
  engine.setTable({
    name: 'Sales',
    sheetName: 'Model',
    startAddress: 'A1',
    endAddress: 'D5',
    columnNames: ['Region', 'Units', 'Price', 'Revenue'],
    columns: [{ name: 'Region' }, { name: 'Units' }, { name: 'Price' }, { name: 'Revenue' }],
    headerRow: true,
    totalsRow: false,
  })
  engine.setCellValue('Lookup', 'A1', 'West')
  engine.setCellValue('Lookup', 'A2', 'East')
  engine.setCellValue('Lookup', 'A3', 'North')
  engine.setCellValue('Lookup', 'B1', 100)
  engine.setCellValue('Lookup', 'B2', 200)
  engine.setCellValue('Lookup', 'B3', 300)
  return engine
}

describe('core coverage contracts', () => {
  it('keeps direct scalar, aggregate, lookup, and inline leaf fast paths semantically aligned', async () => {
    const engine = seedModelWorkbook('coverage-fast-path-workload')
    await engine.ready()

    engine.setCellFormula('Model', 'F1', 'SUM(D2:D5)')
    engine.setCellFormula('Model', 'F2', 'AVERAGE(B2:B5)')
    engine.setCellFormula('Model', 'F3', 'MIN(C2:C5)+MAX(C2:C5)')
    engine.setCellFormula('Model', 'F4', 'MATCH(A2,Lookup!A1:A3,0)')
    engine.setCellFormula('Model', 'F5', 'XLOOKUP(A3,Lookup!A1:A3,Lookup!B1:B3)')
    engine.setCellFormula('Model', 'F6', 'LEN(A2)+LEN(A3)')
    engine.setCellFormula('Model', 'F7', 'CONCATENATE(A2,"-",A3)')
    engine.setCellFormula('Model', 'F8', 'ROUND(SQRT(B3),2)')
    engine.setCellFormula('Model', 'F9', 'IF(B2>3,"large","small")')
    engine.setCellFormula('Model', 'F10', 'PMT(C2/12,B2,D2)')

    expectNumber(engine, 'F1', 218)
    expectNumber(engine, 'F2', 5)
    expectNumber(engine, 'F3', 21)
    expectNumber(engine, 'F4', 1)
    expectNumber(engine, 'F5', 200)
    expect(engine.getCellValue('Model', 'F6')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Model', 'F7')).toEqual({ tag: ValueTag.String, value: 'West-East', stringId: expect.any(Number) })
    expectNumber(engine, 'F8', 2.83)
    expect(engine.getCellValue('Model', 'F9')).toEqual({ tag: ValueTag.String, value: 'large', stringId: expect.any(Number) })

    engine.setCellValue('Model', 'B2', 6)
    engine.setCellValue('Lookup', 'B2', 250)

    expectNumber(engine, 'D2', 60)
    expectNumber(engine, 'F1', 238)
    expectNumber(engine, 'F2', 5.5)
    expectNumber(engine, 'F5', 250)
  })

  it('keeps table metadata and structured formula rewrites stable through structural edits', async () => {
    const engine = seedModelWorkbook('coverage-table-structure-workload')
    await engine.ready()

    engine.setCellFormula('Model', 'G1', 'SUM(Sales[Revenue])')
    engine.setCellFormula('Model', 'G2', 'SUMIF(Sales[Region],"West",Sales[Units])')
    engine.setDataValidation({
      range: { sheetName: 'Model', startAddress: 'B2', endAddress: 'B5' },
      rule: { kind: 'whole', operator: 'greaterThan', values: [0] },
    })

    engine.insertRows('Model', 2, 1)
    engine.setCellValue('Model', 'A3', 'West')
    engine.setCellValue('Model', 'B3', 7)
    engine.setCellValue('Model', 'C3', 13)
    engine.setCellFormula('Model', 'D3', 'B3*C3')
    engine.insertColumns('Model', 4, 1)
    engine.setCellValue('Model', 'E1', 'Revenue')
    engine.setCellFormula('Model', 'G4', 'SUM(Sales[Revenue])')

    const table = engine.getTable('Sales')
    expect(table).toMatchObject({
      name: 'Sales',
      sheetName: 'Model',
      startAddress: 'A1',
      endAddress: 'D6',
      columnNames: ['Region', 'Units', 'Price', 'Revenue'],
    })
    expect(engine.getDataValidations()).toEqual([])
    expect(engine.getCellValue('Model', 'G4').tag).toBe(ValueTag.Number)

    engine.deleteRows('Model', 3, 1)
    engine.deleteColumns('Model', 3, 1)
    expect(engine.getTable('Sales')).toMatchObject({ startAddress: 'A1', endAddress: 'C5' })
  })

  it('surfaces public run-adapter revision proof failures without throwing', async () => {
    const engine = seedModelWorkbook('coverage-run-adapter-workload')
    await engine.ready()
    const { createWorkbookRunAdapter } = await import('../workbook-run-adapter.js')
    const plan: WorkbookActionPlan = {
      modelName: 'coverage',
      actionName: 'badRevision',
      refsUsed: [],
      changed: [],
      checks: [],
      commands: [],
      ops: [],
    }

    const failed = createWorkbookRunAdapter(engine, { baseRevision: -1 }).apply(plan)
    expect(failed).toMatchObject({ status: 'failed', errors: [{ code: 'apply_failed' }] })
  })

  it('keeps formula errors observable while structural cleanup continues', async () => {
    const engine = seedModelWorkbook('coverage-error-structure-workload')
    await engine.ready()
    engine.setCellFormula('Model', 'H1', '1/0')
    engine.setCellFormula('Model', 'H2', 'H1+1')

    expect(engine.getCellValue('Model', 'H1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(engine.getCellValue('Model', 'H2')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })

    engine.clearCell('Model', 'H1')
    expect(engine.getCellValue('Model', 'H2')).toEqual({ tag: ValueTag.Number, value: 1 })
  })
})
