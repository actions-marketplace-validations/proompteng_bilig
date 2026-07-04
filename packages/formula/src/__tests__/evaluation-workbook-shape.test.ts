import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { evaluateAst, evaluateAstResult } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'

describe('formula evaluator: workbook shape hooks and array results', () => {
  it('evaluates GETPIVOTDATA through the pivot-data context hook', () => {
    const result = evaluateAst(parseFormula('GETPIVOTDATA("Sales",B2,"Region","East")'), {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
      resolvePivotData: ({ dataField, sheetName, address, filters }) =>
        dataField === 'Sales' &&
        sheetName === 'Sheet1' &&
        address === 'B2' &&
        filters.length === 1 &&
        filters[0]?.field === 'Region' &&
        filters[0]?.item.tag === ValueTag.String &&
        filters[0].item.value === 'East'
          ? { tag: ValueTag.Number, value: 15 }
          : { tag: ValueTag.Error, code: ErrorCode.Ref },
    })

    expect(result).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('evaluates GROUPBY, PIVOTBY, and MULTIPLE.OPERATIONS workbook-shape formulas', () => {
    const matrixContext = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue =>
        address === 'B5' ? { tag: ValueTag.Number, value: 0 } : { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A5') {
          return [
            { tag: ValueTag.String, value: 'Region', stringId: 0 },
            { tag: ValueTag.String, value: 'East', stringId: 0 },
            { tag: ValueTag.String, value: 'West', stringId: 0 },
            { tag: ValueTag.String, value: 'East', stringId: 0 },
            { tag: ValueTag.String, value: 'West', stringId: 0 },
          ]
        }
        if (start === 'B1' && end === 'B5') {
          return [
            { tag: ValueTag.String, value: 'Product', stringId: 0 },
            { tag: ValueTag.String, value: 'Widget', stringId: 0 },
            { tag: ValueTag.String, value: 'Widget', stringId: 0 },
            { tag: ValueTag.String, value: 'Gizmo', stringId: 0 },
            { tag: ValueTag.String, value: 'Gizmo', stringId: 0 },
          ]
        }
        if (start === 'C1' && end === 'C5') {
          return [
            { tag: ValueTag.String, value: 'Sales', stringId: 0 },
            { tag: ValueTag.Number, value: 10 },
            { tag: ValueTag.Number, value: 7 },
            { tag: ValueTag.Number, value: 5 },
            { tag: ValueTag.Number, value: 4 },
          ]
        }
        return []
      },
      resolveMultipleOperations: ({
        formulaSheetName,
        formulaAddress,
        rowCellAddress,
        rowReplacementAddress,
        columnCellAddress,
        columnReplacementAddress,
      }: {
        formulaSheetName: string
        formulaAddress: string
        rowCellSheetName: string
        rowCellAddress: string
        rowReplacementSheetName: string
        rowReplacementAddress: string
        columnCellSheetName?: string
        columnCellAddress?: string
        columnReplacementSheetName?: string
        columnReplacementAddress?: string
      }) =>
        formulaSheetName === 'Sheet1' &&
        formulaAddress === 'B5' &&
        rowCellAddress === 'B3' &&
        rowReplacementAddress === 'C4' &&
        columnCellAddress === 'B2' &&
        columnReplacementAddress === 'D2'
          ? { tag: ValueTag.Number, value: 5 }
          : { tag: ValueTag.Error, code: ErrorCode.Ref },
    }

    const groupBy = evaluateAstResult(parseFormula('GROUPBY(A1:A5,C1:C5,SUM,3,1)'), matrixContext)
    expect(groupBy).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'Region', stringId: 0 },
        { tag: ValueTag.String, value: 'Sales', stringId: 0 },
        { tag: ValueTag.String, value: 'East', stringId: 0 },
        { tag: ValueTag.Number, value: 15 },
        { tag: ValueTag.String, value: 'West', stringId: 0 },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.String, value: 'Total', stringId: 0 },
        { tag: ValueTag.Number, value: 26 },
      ],
    })

    const pivotBy = evaluateAstResult(parseFormula('PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3,1,0,1)'), matrixContext)
    expect(pivotBy).toEqual({
      kind: 'array',
      rows: 4,
      cols: 4,
      values: [
        { tag: ValueTag.String, value: 'Region', stringId: 0 },
        { tag: ValueTag.String, value: 'Widget', stringId: 0 },
        { tag: ValueTag.String, value: 'Gizmo', stringId: 0 },
        { tag: ValueTag.String, value: 'Total', stringId: 0 },
        { tag: ValueTag.String, value: 'East', stringId: 0 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 15 },
        { tag: ValueTag.String, value: 'West', stringId: 0 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.String, value: 'Total', stringId: 0 },
        { tag: ValueTag.Number, value: 17 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 26 },
      ],
    })

    expect(evaluateAst(parseFormula('MULTIPLE.OPERATIONS(B5,B3,C4,B2,D2)'), matrixContext)).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('resolves scalar defined names through the JS evaluator', () => {
    const value = evaluateAst(parseFormula('TaxRate*A1'), {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue =>
        address === 'A1' ? { tag: ValueTag.Number, value: 100 } : { tag: ValueTag.Empty },
      resolveRange: (): CellValue[] => [],
      resolveName: (name: string): CellValue =>
        name.toUpperCase() === 'TAXRATE' ? { tag: ValueTag.Number, value: 0.085 } : { tag: ValueTag.Error, code: ErrorCode.Name },
    })

    expect(value).toEqual({ tag: ValueTag.Number, value: 8.5 })
  })

  it('preserves sequence array results while flattening them for scalar consumers', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAstResult(parseFormula('SEQUENCE(3,1,1,1)'), context)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ],
    })
    expect(evaluateAst(parseFormula('SUM(SEQUENCE(3,1,1,1))'), context)).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })

    expect(
      evaluateAstResult(parseFormula('FILTER(A1:A4,A1:A4>2)'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('UNIQUE(A1:A4)'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.String, value: 'A', stringId: 0 },
          { tag: ValueTag.String, value: 'B', stringId: 0 },
          { tag: ValueTag.String, value: 'A', stringId: 0 },
          { tag: ValueTag.String, value: 'C', stringId: 0 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'A', stringId: 0 },
        { tag: ValueTag.String, value: 'B', stringId: 0 },
        { tag: ValueTag.String, value: 'C', stringId: 0 },
      ],
    })

    expect(evaluateAstResult(parseFormula('REGEXEXTRACT("a1 b2 c3","[a-z][0-9]",1)'), context)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'a1', stringId: 0 },
        { tag: ValueTag.String, value: 'b2', stringId: 0 },
        { tag: ValueTag.String, value: 'c3', stringId: 0 },
      ],
    })

    expect(evaluateAstResult(parseFormula('REGEXEXTRACT("abc-123","([a-z]+)-([0-9]+)",2)'), context)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'abc', stringId: 0 },
        { tag: ValueTag.String, value: '123', stringId: 0 },
      ],
    })

    expect(evaluateAstResult(parseFormula('TEXTSPLIT("red,blue|green",",","|")'), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'red', stringId: 0 },
        { tag: ValueTag.String, value: 'blue', stringId: 0 },
        { tag: ValueTag.String, value: 'green', stringId: 0 },
        { tag: ValueTag.Error, code: ErrorCode.NA },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('EXPAND(A1:A3,4,2,0)'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 2 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('TRIMRANGE(A1:D4)'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
          { tag: ValueTag.Empty },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Empty },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('INDIRECT("A1:A3")'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 7 },
          { tag: ValueTag.Number, value: 8 },
          { tag: ValueTag.Number, value: 9 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 9 },
      ],
    })

    expect(evaluateAstResult(parseFormula('MAKEARRAY(2,2,LAMBDA(r,c,r+c))'), context)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ],
    })

    expect(evaluateAstResult(parseFormula('MUNIT(3)'), context)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('MMULT(A1:B2,C1:D2)'), {
        ...context,
        resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
          if (start === 'A1' && end === 'B2') {
            return [
              { tag: ValueTag.Number, value: 1 },
              { tag: ValueTag.Number, value: 2 },
              { tag: ValueTag.Number, value: 3 },
              { tag: ValueTag.Number, value: 4 },
            ]
          }
          return [
            { tag: ValueTag.Number, value: 5 },
            { tag: ValueTag.Number, value: 6 },
            { tag: ValueTag.Number, value: 7 },
            { tag: ValueTag.Number, value: 8 },
          ]
        },
      }),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 19 },
        { tag: ValueTag.Number, value: 22 },
        { tag: ValueTag.Number, value: 43 },
        { tag: ValueTag.Number, value: 50 },
      ],
    })

    expect(
      evaluateAst(parseFormula('SUMXMY2(A1:A2,B1:B2)'), {
        ...context,
        resolveRange: (_sheetName: string, start: string, _end: string): CellValue[] =>
          start === 'A1'
            ? [
                { tag: ValueTag.Number, value: 1 },
                { tag: ValueTag.Number, value: 2 },
              ]
            : [
                { tag: ValueTag.Number, value: 3 },
                { tag: ValueTag.Number, value: 4 },
              ],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 8 })

    const randomGrid = evaluateAstResult(parseFormula('RANDARRAY(2,2,3,7,TRUE())'), context)
    expect(randomGrid).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    if (randomGrid.kind !== 'array') {
      throw new Error('expected RANDARRAY to return an array')
    }
    for (const value of randomGrid.values) {
      expect(value.tag).toBe(ValueTag.Number)
      expect(value.value).toBeGreaterThanOrEqual(3)
      expect(value.value).toBeLessThanOrEqual(7)
    }

    expect(
      evaluateAstResult(parseFormula('MAP(A1:A3,LAMBDA(x,x*2))'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ],
    })

    expect(
      evaluateAst(parseFormula('REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })

    expect(
      evaluateAstResult(parseFormula('SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 6 },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('BYROW(A1:B2,LAMBDA(r,SUM(r)))'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 7 },
      ],
    })

    expect(
      evaluateAstResult(parseFormula('BYCOL(A1:B2,LAMBDA(c,SUM(c)))'), {
        ...context,
        resolveRange: (): CellValue[] => [
          { tag: ValueTag.Number, value: 1 },
          { tag: ValueTag.Number, value: 2 },
          { tag: ValueTag.Number, value: 3 },
          { tag: ValueTag.Number, value: 4 },
        ],
      }),
    ).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 6 },
      ],
    })
  })
})
