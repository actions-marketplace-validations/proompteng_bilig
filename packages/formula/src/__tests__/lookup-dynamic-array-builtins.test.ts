import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { getLookupBuiltin, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const empty = (): CellValue => ({ tag: ValueTag.Empty })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}
describe('lookup builtins: dynamic arrays, matrix helpers, and lookup edge cases', () => {
  it('covers remaining SUMPRODUCT validation branches', () => {
    const SUMPRODUCT = getLookupBuiltin('SUMPRODUCT')!

    expect(SUMPRODUCT()).toEqual(err(ErrorCode.Value))
    expect(SUMPRODUCT(num(1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMPRODUCT(cellRange([num(1), num(2)], 2, 1), cellRange([num(3)], 1, 1))).toEqual(err(ErrorCode.Value))
  })

  it('covers COUNTIFS and SUMIFS error branches', () => {
    const COUNTIFS = getLookupBuiltin('COUNTIFS')!
    const SUMIFS = getLookupBuiltin('SUMIFS')!

    expect(COUNTIFS(cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(COUNTIFS(cellRange([num(1), num(2)], 2, 1), text('>1'), cellRange([num(2)], 1, 1), text('2'))).toEqual(err(ErrorCode.Value))
    expect(
      SUMIFS(cellRange([num(1), num(2)], 2, 1), cellRange([num(1), num(2)], 2, 1), text('>1'), cellRange([num(3)], 1, 1), text('3')),
    ).toEqual(err(ErrorCode.Value))
  })

  it('covers remaining database passthrough and SUMPRODUCT scalar validation branches', () => {
    const DSTDEVP = getLookupBuiltin('DSTDEVP')!
    const DSUM = getLookupBuiltin('DSUM')!
    const SUMPRODUCT = getLookupBuiltin('SUMPRODUCT')!

    const database = cellRange([text('Age'), text('Yield'), num(10), num(5), num(12), num(7)], 3, 2)

    expect(DSTDEVP(database, text('Yield'), cellRange([text('Age'), err(ErrorCode.Ref)], 2, 1))).toEqual(err(ErrorCode.Ref))
    expect(DSUM(database, text('Yield'), cellRange([text('Age'), err(ErrorCode.Name)], 2, 1))).toEqual(err(ErrorCode.Name))
    expect(SUMPRODUCT(num(1), num(2))).toEqual(err(ErrorCode.Value))
  })

  it('ignores blocked database criteria rows with missing or blank headers', () => {
    const DCOUNT = getLookupBuiltin('DCOUNT')!
    const DSUM = getLookupBuiltin('DSUM')!

    const database = cellRange([text('Age'), text('Yield'), num(10), num(5), num(12), num(7), num(12), num(9)], 4, 2)

    expect(DCOUNT(database, text('Yield'), cellRange([text('Age'), text('Missing'), num(12), num(1)], 2, 2))).toEqual(num(0))
    expect(DSUM(database, text('Yield'), cellRange([{ tag: ValueTag.Empty }, text('Yield'), num(1), num(7)], 2, 2))).toEqual(num(0))
  })

  it('covers UNIQUE by-column/row modes with duplicate and error branches', () => {
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    expect(UNIQUE(cellRange([num(1), num(1), num(2), num(1), num(1), num(2)], 2, 3), bool(true), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(2), num(2)],
    })

    expect(UNIQUE(cellRange([num(1), num(2), num(3), num(4)], 2, 2), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(2), num(3), num(4)],
    })

    expect(UNIQUE(cellRange([num(1), err(ErrorCode.Ref), num(3), num(4)], 2, 2), bool(true))).toEqual(err(ErrorCode.Value))

    expect(UNIQUE(cellRange([num(1), num(2), err(ErrorCode.Name), num(4)], 2, 2))).toEqual(err(ErrorCode.Value))
  })

  it('covers criteria matching with error values and invalid operand types', () => {
    const COUNTIF = getLookupBuiltin('COUNTIF')!

    expect(COUNTIF(cellRange([err(ErrorCode.Ref), num(1), num(2)], 3, 1), text('>0'))).toEqual(num(2))
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text('>=a'))).toEqual(num(0))
    expect(COUNTIF(cellRange([num(3), num(1), num(4)], 3, 1), text('<2'))).toEqual(num(1))
  })

  it('covers boundary behavior for lookup reshaping helpers', () => {
    const OFFSET = getLookupBuiltin('OFFSET')!
    const TAKE = getLookupBuiltin('TAKE')!
    const DROP = getLookupBuiltin('DROP')!
    const CHOOSECOLS = getLookupBuiltin('CHOOSECOLS')!
    const CHOOSEROWS = getLookupBuiltin('CHOOSEROWS')!

    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2)
    const column = cellRange([num(1), num(2), num(3)], 3, 1)

    expect(OFFSET(matrix, num(0), num(0))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: matrix.values,
    })
    expect(OFFSET(matrix, num(0), num(0), num(1), num(1), num(2))).toEqual(err(ErrorCode.Value))

    expect(TAKE(column)).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: column.values,
    })
    expect(TAKE(column, num(0))).toEqual(err(ErrorCode.Value))

    expect(DROP(matrix, num(0), num(0))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: matrix.values,
    })
    expect(DROP(matrix, num(5))).toEqual(err(ErrorCode.Value))

    expect(CHOOSECOLS(matrix, num(3))).toEqual(err(ErrorCode.Value))
    expect(CHOOSEROWS(matrix, num(3))).toEqual(err(ErrorCode.Value))
  })

  it('supports FILTER and UNIQUE dynamic-array results', () => {
    const FILTER = getLookupBuiltin('FILTER')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    const filtered = FILTER(
      cellRange([num(1), num(3), num(2), num(4)], 4, 1),
      cellRange([bool(false), bool(true), bool(false), bool(true)], 4, 1),
    )
    expect(filtered).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(3), num(4)],
    })

    const unique = UNIQUE(cellRange([text('A'), text('B'), text('A'), text('C')], 4, 1))
    expect(unique).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [text('A'), text('B'), text('C')],
    })
  })

  it('covers FILTER fallbacks and UNIQUE row and column modes', () => {
    const FILTER = getLookupBuiltin('FILTER')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    expect(
      FILTER(cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3), cellRange([bool(true), bool(false), bool(true)], 1, 3)),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(3), num(4), num(6)],
    })

    expect(FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([bool(false), bool(false)], 2, 1), text('empty'))).toEqual(text('empty'))

    expect(FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([text('bad'), bool(true)], 2, 1))).toEqual(err(ErrorCode.Value))

    expect(FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([bool(true), bool(false), bool(true)], 3, 1))).toEqual(err(ErrorCode.Value))

    expect(FILTER(cellRange([num(1), num(2)], 2, 1), cellRange([err(ErrorCode.Ref), bool(true)], 2, 1))).toEqual(err(ErrorCode.Ref))

    expect(UNIQUE(cellRange([text('A'), text('b'), text('a'), text('C'), text('c')], 5, 1), bool(false), bool(true))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [text('b')],
    })

    expect(UNIQUE(cellRange([text('A'), text('B'), text('A'), text('C'), num(1), num(2), num(1), num(3)], 2, 4), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [text('A'), text('B'), text('C'), num(1), num(2), num(3)],
    })

    expect(
      UNIQUE(cellRange([text('A'), text('B'), text('A'), text('C'), num(1), num(2), num(1), num(3)], 2, 4), bool(true), bool(true)),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('B'), text('C'), num(2), num(3)],
    })

    expect(UNIQUE(cellRange([text('A'), num(1), text('a'), num(1), text('B'), num(2)], 3, 2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('A'), num(1), text('B'), num(2)],
    })

    expect(UNIQUE(cellRange([err(ErrorCode.Name)], 1, 1))).toEqual(err(ErrorCode.Name))
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), text('bad'))).toEqual(err(ErrorCode.Value))
  })

  it('covers FILTER horizontal validation and UNIQUE argument validation branches', () => {
    const FILTER = getLookupBuiltin('FILTER')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    expect(FILTER(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([err(ErrorCode.Ref), bool(true)], 1, 2))).toEqual(
      err(ErrorCode.Ref),
    )
    expect(FILTER(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([text('bad'), bool(true)], 1, 2))).toEqual(
      err(ErrorCode.Value),
    )
    expect(
      FILTER(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([bool(false), bool(false)], 1, 2), cellRange([num(0)], 1, 1)),
    ).toEqual(err(ErrorCode.Value))

    expect(UNIQUE(err(ErrorCode.Name))).toEqual(err(ErrorCode.Value))
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), cellRange([bool(true)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(UNIQUE(cellRange([num(1), num(2)], 2, 1), bool(false), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(UNIQUE(cellRange([text('A'), num(1), text('A'), num(1), text('B'), num(2)], 3, 2), bool(false), bool(true))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [text('B'), num(2)],
    })
  })

  it('supports matrix and extended numeric lookup builtins', () => {
    const SUMX2MY2 = getLookupBuiltin('SUMX2MY2')!
    const SUMX2PY2 = getLookupBuiltin('SUMX2PY2')!
    const SUMXMY2 = getLookupBuiltin('SUMXMY2')!
    const MDETERM = getLookupBuiltin('MDETERM')!
    const MINVERSE = getLookupBuiltin('MINVERSE')!
    const MMULT = getLookupBuiltin('MMULT')!
    const PERCENTOF = getLookupBuiltin('PERCENTOF')!

    expect(SUMX2MY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(num(11))
    expect(SUMX2PY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(num(15))
    expect(SUMXMY2(cellRange([num(2), num(3)], 2, 1), cellRange([num(1), num(1)], 2, 1))).toEqual(num(5))

    expect(MDETERM(cellRange([num(1), num(2), num(3), num(4)], 2, 2))).toEqual(num(-2))
    expect(MDETERM(cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(err(ErrorCode.Value))

    const inverse = MINVERSE(cellRange([num(4), num(7), num(2), num(6)], 2, 2))
    expect(inverse).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    if (!(inverse && 'kind' in inverse && inverse.kind === 'array')) {
      throw new Error('expected MINVERSE to return an array')
    }
    expect(inverse.values.map((value) => value.tag)).toEqual([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number])
    expect(inverse.values.map((value) => value.value)).toEqual([
      expect.closeTo(0.6, 12),
      expect.closeTo(-0.7, 12),
      expect.closeTo(-0.2, 12),
      expect.closeTo(0.4, 12),
    ])
    expect(MINVERSE(cellRange([num(1), num(2), num(2), num(4)], 2, 2))).toEqual(err(ErrorCode.Value))

    expect(MMULT(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(5), num(6), num(7), num(8)], 2, 2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(19), num(22), num(43), num(50)],
    })
    expect(MMULT(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(5), num(6), num(7)], 3, 1))).toEqual(
      err(ErrorCode.Value),
    )

    expect(PERCENTOF(cellRange([num(2), num(3)], 2, 1), cellRange([num(10), num(10)], 2, 1))).toEqual(num(0.25))
    expect(PERCENTOF(cellRange([num(2)], 1, 1), cellRange([num(0)], 1, 1))).toEqual(err(ErrorCode.Div0))
    expect(SUMXMY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
  })

  it('covers matrix helper validation and percent-of error branches', () => {
    const SUMX2PY2 = getLookupBuiltin('SUMX2PY2')!
    const MDETERM = getLookupBuiltin('MDETERM')!
    const MINVERSE = getLookupBuiltin('MINVERSE')!
    const MMULT = getLookupBuiltin('MMULT')!
    const PERCENTOF = getLookupBuiltin('PERCENTOF')!

    expect(MINVERSE(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value))
    expect(MINVERSE(cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3))).toEqual(err(ErrorCode.Value))

    expect(MMULT(err(ErrorCode.Name), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(MMULT(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value))
    expect(MMULT(cellRange([num(1), num(2)], 1, 2), cellRange([num(1), num(2)], 1, 2))).toEqual(err(ErrorCode.Value))

    expect(PERCENTOF(err(ErrorCode.Name), cellRange([num(10)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(PERCENTOF(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value))
    expect(SUMX2PY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMX2PY2(cellRange([num(1)], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value))
    expect(SUMX2PY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(MDETERM(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value))
  })

  it('matches Excel empty-string criteria semantics for conditional aggregates', () => {
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const COUNTIFS = getLookupBuiltin('COUNTIFS')!
    const SUMIF = getLookupBuiltin('SUMIF')!
    const SUMIFS = getLookupBuiltin('SUMIFS')!

    const criteriaRange = cellRange([text('North'), empty(), text(''), text(' '), text('South')], 5, 1)
    const sumRange = cellRange([num(10), num(20), num(30), num(40), num(50)], 5, 1)

    expect(COUNTIF(criteriaRange, text(''))).toEqual(num(2))
    expect(COUNTIF(criteriaRange, text('='))).toEqual(num(2))
    expect(COUNTIF(criteriaRange, text('<>'))).toEqual(num(4))
    expect(COUNTIFS(criteriaRange, text('<>'))).toEqual(num(4))
    expect(SUMIF(criteriaRange, text(''), sumRange)).toEqual(num(50))
    expect(SUMIF(criteriaRange, text('<>'), sumRange)).toEqual(num(130))
    expect(SUMIFS(sumRange, criteriaRange, text('<>'))).toEqual(num(130))
  })

  it('covers conditional criteria parsing variants', () => {
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const SUMIF = getLookupBuiltin('SUMIF')!

    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), num(2))).toEqual(num(1))
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text('<>2'))).toEqual(num(2))
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text('>=2'))).toEqual(num(2))
    expect(COUNTIF(cellRange([num(1), num(2), num(3)], 3, 1), text('<=2'))).toEqual(num(2))
    expect(COUNTIF(cellRange([bool(true), bool(false), bool(true)], 3, 1), text('=TRUE'))).toEqual(num(2))
    expect(COUNTIF(cellRange([text(''), text('x'), text('')], 3, 1), text('='))).toEqual(num(2))
    expect(SUMIF(cellRange([text('a'), text('b'), text('c')], 3, 1), text('<>b'), cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(
      num(4),
    )
  })

  it('covers FILTER column selection and UNIQUE row and column de-duplication branches', () => {
    const FILTER = getLookupBuiltin('FILTER')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!

    expect(FILTER(cellRange([text('A'), text('B'), text('C'), text('D')], 2, 2), cellRange([bool(true), bool(false)], 1, 2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [text('A'), text('C')],
    })
    expect(FILTER(cellRange([text('A'), text('B'), text('C'), text('D')], 2, 2), cellRange([text('bad'), bool(true)], 1, 2))).toEqual(
      err(ErrorCode.Value),
    )
    expect(
      FILTER(cellRange([text('A'), text('B'), text('C'), text('D')], 2, 2), cellRange([bool(false), bool(false)], 1, 2), text('empty')),
    ).toEqual(text('empty'))
    expect(
      FILTER(
        cellRange([text('A'), text('B'), text('C'), text('D')], 2, 2),
        cellRange([bool(false), bool(false)], 1, 2),
        cellRange([text('x')], 1, 1),
      ),
    ).toEqual(err(ErrorCode.Value))
    expect(FILTER(err(ErrorCode.Ref), cellRange([bool(true)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(FILTER(cellRange([text('A')], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value))

    expect(
      UNIQUE(cellRange([text('A'), text('A'), text('B'), text('C'), num(1), num(1), num(2), num(3)], 2, 4), bool(true), bool(true)),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('B'), text('C'), num(2), num(3)],
    })
    expect(UNIQUE(cellRange([text('A'), num(1), text('A'), num(1), text('B'), num(2)], 3, 2), bool(false), bool(false))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [text('A'), num(1), text('B'), num(2)],
    })
  })

  it('covers SUMPRODUCT and the remaining matrix helper validation paths', () => {
    const SUMPRODUCT = getLookupBuiltin('SUMPRODUCT')!
    const SUMX2MY2 = getLookupBuiltin('SUMX2MY2')!
    const SUMXMY2 = getLookupBuiltin('SUMXMY2')!
    const MDETERM = getLookupBuiltin('MDETERM')!

    expect(SUMPRODUCT()).toEqual(err(ErrorCode.Value))
    expect(SUMPRODUCT(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMPRODUCT(cellRange([num(2), num(3)], 2, 1), cellRange([num(4), num(5)], 2, 1))).toEqual(num(23))
    expect(SUMPRODUCT(cellRange([num(2), num(3)], 2, 1), cellRange([num(4), num(5), num(6)], 3, 1))).toEqual(err(ErrorCode.Value))

    expect(SUMX2MY2(err(ErrorCode.Ref), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMX2MY2(cellRange([num(1)], 1, 1), err(ErrorCode.Name))).toEqual(err(ErrorCode.Value))
    expect(SUMX2MY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMXMY2(cellRange([num(1)], 1, 1), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Value))
    expect(SUMXMY2(cellRange([num(1)], 1, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))

    expect(MDETERM(cellRange([], 0, 0))).toEqual(err(ErrorCode.Value))
  })

  it('covers AVERAGEIFS, MINIFS, and MAXIFS validation and zero-match branches', () => {
    const AVERAGEIFS = getLookupBuiltin('AVERAGEIFS')!
    const MINIFS = getLookupBuiltin('MINIFS')!
    const MAXIFS = getLookupBuiltin('MAXIFS')!

    expect(AVERAGEIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(AVERAGEIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('match')], 1, 1), text('match'))).toEqual(err(ErrorCode.Ref))
    expect(AVERAGEIFS(cellRange([text('skip')], 1, 1), cellRange([num(1)], 1, 1), text('>0'))).toEqual(err(ErrorCode.Div0))
    expect(MINIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(MINIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('match')], 1, 1), text('match'))).toEqual(err(ErrorCode.Ref))
    expect(MINIFS(cellRange([num(1)], 1, 1), err(ErrorCode.Name), text('>0'))).toEqual(err(ErrorCode.Name))
    expect(MAXIFS(err(ErrorCode.Ref), cellRange([num(1)], 1, 1), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(MAXIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('match')], 1, 1), text('match'))).toEqual(err(ErrorCode.Ref))
    expect(MAXIFS(cellRange([num(1)], 1, 1), err(ErrorCode.Name), text('>0'))).toEqual(err(ErrorCode.Name))
  })

  it('covers remaining matrix, sort, and criteria edge cases', () => {
    const MINVERSE = getLookupBuiltin('MINVERSE')!
    const SORT = getLookupBuiltin('SORT')!
    const UNIQUE = getLookupBuiltin('UNIQUE')!
    const HSTACK = getLookupBuiltin('HSTACK')!
    const VSTACK = getLookupBuiltin('VSTACK')!
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const TOCOL = getLookupBuiltin('TOCOL')!
    const TOROW = getLookupBuiltin('TOROW')!

    // Singular matrix for MINVERSE
    expect(MINVERSE(cellRange([num(1), num(2), num(2), num(4)], 2, 2))).toEqual(err(ErrorCode.Value))

    // SORT by column
    const matrix = cellRange([num(3), num(1), num(4), num(2)], 2, 2)
    expect(SORT(matrix, num(1), num(1), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(3), num(2), num(4)],
    })

    // UNIQUE exactlyOnce on 2D
    const matrix2 = cellRange([num(1), num(2), num(1), num(2), num(3), num(4)], 3, 2)
    expect(UNIQUE(matrix2, bool(false), bool(true))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(3), num(4)],
    })

    // HSTACK/VSTACK pad ragged array edges with #N/A instead of broadcasting.
    expect(HSTACK(cellRange([num(1)], 1, 1), cellRange([num(2), num(3)], 2, 1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), num(2), err(ErrorCode.NA), num(3)],
    })
    expect(VSTACK(cellRange([num(1)], 1, 1), cellRange([num(2), num(3)], 1, 2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(1), err(ErrorCode.NA), num(2), num(3)],
    })
    expect(VSTACK()).toEqual(err(ErrorCode.Value))
    expect(
      VSTACK({
        kind: 'range',
        refKind: 'rows',
        rows: 1,
        cols: 1,
        values: [num(1)],
      }),
    ).toEqual(err(ErrorCode.Value))
    expect(VSTACK(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(5), num(6), num(7)], 1, 3))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 3,
      values: [num(1), num(2), err(ErrorCode.NA), num(3), num(4), err(ErrorCode.NA), num(5), num(6), num(7)],
    })

    // matchesCriteria operators
    const range = cellRange([num(1), num(2), num(3), num(4)], 4, 1)
    expect(COUNTIF(range, text('<>2'))).toEqual(num(3))
    expect(COUNTIF(range, text('<=2'))).toEqual(num(2))
    expect(COUNTIF(range, text('>=3'))).toEqual(num(2))
    expect(COUNTIF(range, text('<3'))).toEqual(num(2))

    // TOCOL/TOROW ignoreEmpty
    const sparse = cellRange([num(1), { tag: ValueTag.Empty }, num(2)], 3, 1)
    expect(TOCOL(sparse, num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(1), num(2)],
    })
    expect(TOROW(sparse, num(1))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(1), num(2)],
    })
  })

  it('covers remaining lookup, database, and cash-flow validation branches', () => {
    const AREAS = getLookupBuiltin('AREAS')!
    const ARRAYTOTEXT = getLookupBuiltin('ARRAYTOTEXT')!
    const ROWS = getLookupBuiltin('ROWS')!
    const COLUMNS = getLookupBuiltin('COLUMNS')!
    const DCOUNT = getLookupBuiltin('DCOUNT')!
    const MATCH = getLookupBuiltin('MATCH')!
    const LOOKUP = getLookupBuiltin('LOOKUP')!
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!
    const XLOOKUP = getLookupBuiltin('XLOOKUP')!
    const XMATCH = getLookupBuiltin('XMATCH')!

    const database = cellRange(
      [text('Age'), text('Height'), text('Yield'), num(10), num(100), num(5), num(12), num(110), num(7), num(12), num(120), num(9)],
      4,
      3,
    )
    const ageCriteria = cellRange([text('Age'), num(12)], 2, 1)
    const ageCriteriaWithBlankClause = cellRange([text('Age'), text('Yield'), num(12), { tag: ValueTag.Empty }], 2, 2)

    expect(DCOUNT(num(1), text('Yield'), ageCriteria)).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, text('Yield'), num(1))).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(cellRange([], 0, 0), text('Yield'), ageCriteria)).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, text('Yield'), ageCriteriaWithBlankClause)).toEqual(num(2))

    expect(AREAS(num(1))).toEqual(err(ErrorCode.Value))
    expect(ROWS(num(1))).toEqual(num(1))
    expect(ROWS(err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(COLUMNS(num(1))).toEqual(num(1))
    expect(COLUMNS(err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(ARRAYTOTEXT(cellRange([err(ErrorCode.Ref)], 1, 1))).toEqual(err(ErrorCode.Value))

    const duplicateLookup = cellRange([text('pear'), text('apple'), text('pear')], 3, 1)
    const duplicateReturn = cellRange([num(10), num(20), num(30)], 3, 1)

    expect(MATCH(cellRange([text('pear')], 1, 1), duplicateLookup, num(0))).toEqual(err(ErrorCode.Value))
    expect(MATCH(err(ErrorCode.Ref), duplicateLookup, num(0))).toEqual(err(ErrorCode.Ref))
    expect(MATCH(text('pear'), duplicateLookup, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(MATCH(text('pear'), duplicateLookup, num(2))).toEqual(err(ErrorCode.Value))
    expect(MATCH(num(2), cellRange([text('x'), num(3)], 2, 1), num(1))).toEqual(err(ErrorCode.NA))

    expect(LOOKUP(cellRange([num(1)], 1, 1), duplicateLookup)).toEqual(err(ErrorCode.Value))
    expect(Reflect.apply(LOOKUP, undefined, [])).toEqual(err(ErrorCode.Value))
    expect(LOOKUP(err(ErrorCode.Ref), duplicateLookup)).toEqual(err(ErrorCode.Ref))
    expect(LOOKUP(text('pear'), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(LOOKUP(text('pear'), cellRange([text('pear')], 1, 1), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(LOOKUP(text('pear'), cellRange([text('pear'), text('apple'), text('plum'), text('berry')], 2, 2))).toEqual(err(ErrorCode.Value))
    expect(LOOKUP(text('pear'), duplicateLookup, cellRange([num(1), num(2)], 2, 1))).toEqual(num(1))

    const verticalTable = cellRange([text('apple'), num(10), text('pear'), num(20), text('plum'), num(30)], 3, 2)
    const horizontalTable = cellRange([text('apple'), text('pear'), text('plum'), num(10), num(20), num(30)], 2, 3)
    expect(VLOOKUP(cellRange([text('pear')], 1, 1), verticalTable, num(2))).toEqual(err(ErrorCode.Value))
    expect(VLOOKUP(err(ErrorCode.Ref), verticalTable, num(2))).toEqual(err(ErrorCode.Ref))
    expect(VLOOKUP(text('pear'), verticalTable, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(VLOOKUP(text('pear'), verticalTable, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(VLOOKUP(text('pear'), verticalTable, num(0))).toEqual(err(ErrorCode.Value))
    expect(VLOOKUP(text('pear'), verticalTable, num(2), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(VLOOKUP(num(2), cellRange([text('x'), num(1), num(10), num(20)], 2, 2), num(2))).toEqual(err(ErrorCode.Value))

    expect(HLOOKUP(cellRange([text('pear')], 1, 1), horizontalTable, num(2))).toEqual(err(ErrorCode.Value))
    expect(HLOOKUP(err(ErrorCode.Ref), horizontalTable, num(2))).toEqual(err(ErrorCode.Ref))
    expect(HLOOKUP(text('pear'), horizontalTable, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(HLOOKUP(text('pear'), horizontalTable, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(HLOOKUP(text('pear'), horizontalTable, num(0))).toEqual(err(ErrorCode.Value))
    expect(HLOOKUP(text('pear'), horizontalTable, num(2), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(HLOOKUP(num(2), cellRange([text('x'), num(1), num(10), num(20)], 2, 2), num(2))).toEqual(err(ErrorCode.Value))

    expect(XLOOKUP(text('pear'), duplicateLookup, duplicateReturn, text('fallback'), num(0), num(-1))).toEqual(num(30))
    expect(XLOOKUP(text('pear'), duplicateLookup, duplicateReturn, text('fallback'), num(1), num(1))).toEqual(num(10))
    expect(XLOOKUP(cellRange([text('pear')], 1, 1), duplicateLookup, duplicateReturn)).toEqual(num(10))
    expect(XLOOKUP(text('pear'), duplicateLookup, cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(XLOOKUP(err(ErrorCode.Ref), duplicateLookup, duplicateReturn)).toEqual(err(ErrorCode.Ref))
    expect(XLOOKUP(text('pear'), duplicateLookup, duplicateReturn, text('fallback'), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(XLOOKUP(text('pear'), duplicateLookup, duplicateReturn, text('fallback'), num(0), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(XMATCH(text('pear'), duplicateLookup, num(0), num(-1))).toEqual(num(3))
    expect(XMATCH(text('pear'), duplicateLookup, num(2), num(1))).toEqual(num(1))
    expect(XMATCH(text('pear'), duplicateLookup, num(3), num(1))).toEqual(err(ErrorCode.Value))
    expect(XMATCH(cellRange([text('pear')], 1, 1), duplicateLookup)).toEqual(err(ErrorCode.Value))
    expect(XMATCH(err(ErrorCode.Ref), duplicateLookup)).toEqual(err(ErrorCode.Ref))
    expect(XMATCH(text('pear'), duplicateLookup, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(XMATCH(text('pear'), duplicateLookup, num(0), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(XMATCH(text('missing'), duplicateLookup)).toEqual(err(ErrorCode.NA))
  })
})
