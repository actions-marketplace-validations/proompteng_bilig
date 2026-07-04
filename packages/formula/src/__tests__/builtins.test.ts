import { BuiltinId, ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it } from 'vitest'
import { getBuiltin, getBuiltinId } from '../builtins.js'
import { clearExternalFunctionAdapters } from '../external-function-adapter.js'
import type { ArrayValue } from '../runtime-values.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: core scalar and aggregate builtins', () => {
  it('supports CHOOSE, COUNTBLANK, and bitwise builtins', () => {
    const CHOOSE = getBuiltin('CHOOSE')!
    const COUNTBLANK = getBuiltin('COUNTBLANK')!
    const BITAND = getBuiltin('BITAND')!
    const BITOR = getBuiltin('BITOR')!
    const BITXOR = getBuiltin('BITXOR')!
    const BITLSHIFT = getBuiltin('BITLSHIFT')!
    const BITRSHIFT = getBuiltin('BITRSHIFT')!

    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: 'zero', stringId: 1 },
        {
          tag: ValueTag.Number,
          value: 10,
        },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(
      CHOOSE({ tag: ValueTag.String, value: '1', stringId: 2 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 20 }),
    ).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(
      CHOOSE(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
        {
          tag: ValueTag.Number,
          value: 20,
        },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(CHOOSE({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CHOOSE({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 20 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      COUNTBLANK(
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'x', stringId: 1 },
        { tag: ValueTag.String, value: '', stringId: 2 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    expect(BITAND({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(BITOR({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(BITXOR({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(BITRSHIFT({ tag: ValueTag.Number, value: 16 }, { tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(BITLSHIFT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad' })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('rejects bitwise calls that do not have exactly two arguments', () => {
    const threeArgs = [
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 1 },
    ] as const

    for (const name of ['BITAND', 'BITOR', 'BITXOR', 'BITLSHIFT', 'BITRSHIFT']) {
      const builtin = getBuiltin(name)!
      expect(builtin({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
      expect(builtin(...threeArgs)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    }
  })

  it('rejects excess fixed-arity scalar builtin arguments through direct lookup', () => {
    const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
    const number = { tag: ValueTag.Number, value: 1 } as const
    const text = { tag: ValueTag.String, value: 'abc', stringId: 1 } as const
    const format = { tag: ValueTag.String, value: '0', stringId: 2 } as const

    for (const [name, args] of [
      ['ABS', [number, number]],
      ['INT', [number, number]],
      ['SIN', [number, number]],
      ['COS', [number, number]],
      ['TAN', [number, number]],
      ['ASIN', [number, number]],
      ['ACOS', [number, number]],
      ['ATAN', [number, number]],
      ['DEGREES', [number, number]],
      ['RADIANS', [number, number]],
      ['EXP', [number, number]],
      ['LN', [number, number]],
      ['LOG10', [number, number]],
      ['SQRT', [number, number]],
      ['SIGN', [number, number]],
      ['EVEN', [number, number]],
      ['ODD', [number, number]],
      ['FACT', [number, number]],
      ['FACTDOUBLE', [number, number]],
      ['ROUND', [number, number, number]],
      ['ROUNDUP', [number, number, number]],
      ['ROUNDDOWN', [number, number, number]],
      ['LOG', [number, number, number]],
      ['POWER', [number, number, number]],
      ['LEFT', [text, number, number]],
      ['RIGHT', [text, number, number]],
      ['LEN', [text, number]],
      ['LENB', [text, number]],
      ['UPPER', [text, number]],
      ['LOWER', [text, number]],
      ['TRIM', [text, number]],
      ['VALUE', [text, number]],
      ['CHAR', [number, number]],
      ['CODE', [text, number]],
      ['UNICODE', [text, number]],
      ['UNICHAR', [number, number]],
      ['CLEAN', [text, number]],
      ['REPT', [text, number, number]],
      ['TEXT', [number, format, number]],
      ['EXACT', [text, text, number]],
      ['MID', [text, number, number, number]],
      ['FIND', [text, text, number, number]],
      ['SEARCH', [text, text, number, number]],
      ['REPLACE', [text, number, number, text, number]],
      ['SUBSTITUTE', [text, text, text, number, number]],
      ['YEAR', [number, number]],
      ['MONTH', [number, number]],
      ['DAY', [number, number]],
      ['HOUR', [number, number]],
      ['MINUTE', [number, number]],
      ['SECOND', [number, number]],
      ['EDATE', [number, number, number]],
      ['EOMONTH', [number, number, number]],
      ['TIMEVALUE', [text, number]],
      ['SINH', [number, number]],
      ['COSH', [number, number]],
      ['TANH', [number, number]],
      ['ASINH', [number, number]],
      ['ACOSH', [number, number]],
      ['ATANH', [number, number]],
      ['ACOT', [number, number]],
      ['ACOTH', [number, number]],
      ['COT', [number, number]],
      ['COTH', [number, number]],
      ['CSC', [number, number]],
      ['CSCH', [number, number]],
      ['SEC', [number, number]],
      ['SECH', [number, number]],
      ['DELTA', [number, number, number]],
      ['GESTEP', [number, number, number]],
      ['DOLLARDE', [number, number, number]],
      ['DOLLARFR', [number, number, number]],
      ['COMBIN', [number, number, number]],
      ['COMBINA', [number, number, number]],
      ['QUOTIENT', [number, number, number]],
      ['BASE', [number, number, number, number]],
      ['DECIMAL', [text, number, number]],
      ['BIN2DEC', [text, number]],
      ['BIN2HEX', [text, number, number]],
      ['ADDRESS', [number, number, number, number, text, number]],
      ['CONVERT', [number, text, text, number]],
    ] as const) {
      expect(getBuiltin(name)?.(...args)).toEqual(valueError)
    }
  })

  it('matches Excel direct empty text aggregate argument semantics', () => {
    const SUM = getBuiltin('SUM')!
    const AVERAGE = getBuiltin('AVERAGE')!
    const COUNT = getBuiltin('COUNT')!
    const MIN = getBuiltin('MIN')!
    const MAX = getBuiltin('MAX')!
    const PRODUCT = getBuiltin('PRODUCT')!
    const SUMSQ = getBuiltin('SUMSQ')!
    const GEOMEAN = getBuiltin('GEOMEAN')!
    const HARMEAN = getBuiltin('HARMEAN')!

    const emptyText = { tag: ValueTag.String, value: '', stringId: 1 } as const

    expect(SUM(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(AVERAGE(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(COUNT(emptyText)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(MIN(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(MAX(emptyText, { tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(PRODUCT(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(SUMSQ(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(GEOMEAN(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(HARMEAN(emptyText)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('supports numeric aggregates and error propagation', () => {
    const sum = getBuiltin('SUM')
    const avg = getBuiltin('AVG')
    const mod = getBuiltin('MOD')

    expect(sum?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    expect(
      avg?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.String, value: 'skip', stringId: 1 }, { tag: ValueTag.Empty }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(sum?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(mod?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
  })

  it('matches Excel log error semantics', () => {
    const LN = getBuiltin('LN')!
    const LOG10 = getBuiltin('LOG10')!
    const LOG = getBuiltin('LOG')!
    const number = (value: number): CellValue => ({ tag: ValueTag.Number, value })
    const error = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

    expect(LN(number(-1))).toEqual(error(ErrorCode.Num))
    expect(LN(number(0))).toEqual(error(ErrorCode.Num))
    expect(LOG10(number(0))).toEqual(error(ErrorCode.Num))
    expect(LOG(number(-1))).toEqual(error(ErrorCode.Num))
    expect(LOG(number(10), number(1))).toEqual(error(ErrorCode.Div0))
    expect(LOG(error(ErrorCode.Name))).toEqual(error(ErrorCode.Name))
  })

  it('supports SEQUENCE spill generation and validation', () => {
    const SEQUENCE = getBuiltin('SEQUENCE')!

    expect(
      SEQUENCE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 12 },
        { tag: ValueTag.Number, value: 14 },
        { tag: ValueTag.Number, value: 16 },
        { tag: ValueTag.Number, value: 18 },
        { tag: ValueTag.Number, value: 20 },
      ],
    })
    expect(SEQUENCE({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      SEQUENCE(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 400 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(SEQUENCE({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('supports Bessel engineering builtins', () => {
    const BESSELI = getBuiltin('BESSELI')!
    const BESSELJ = getBuiltin('BESSELJ')!
    const BESSELK = getBuiltin('BESSELK')!
    const BESSELY = getBuiltin('BESSELY')!
    const expectNumeric = (value: CellValue, expected: number) => {
      expect(value).toMatchObject({ tag: ValueTag.Number })
      if (value.tag !== ValueTag.Number) {
        throw new Error('Expected numeric builtin result')
      }
      expect(value.value).toBeCloseTo(expected, 7)
    }

    const besseli = BESSELI({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(besseli, 0.981666428)
    const besselj = BESSELJ({ tag: ValueTag.Number, value: 1.9 }, { tag: ValueTag.Number, value: 2 })
    expectNumeric(besselj, 0.329925728)
    const besselk = BESSELK({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(besselk, 0.277387804)
    const bessely = BESSELY({ tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Number, value: 1 })
    expectNumeric(bessely, 0.145918138)
    expect(BESSELK({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('supports CONVERT and EUROCONVERT', () => {
    const CONVERT = getBuiltin('CONVERT')!
    const EUROCONVERT = getBuiltin('EUROCONVERT')!

    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.String, value: 'mi', stringId: 1 },
        { tag: ValueTag.String, value: 'km', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 9.656064,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 68 },
        { tag: ValueTag.String, value: 'F', stringId: 3 },
        { tag: ValueTag.String, value: 'C', stringId: 4 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 20,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 2.5 },
        { tag: ValueTag.String, value: 'ft', stringId: 5 },
        { tag: ValueTag.String, value: 'sec', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1.2 },
        { tag: ValueTag.String, value: 'DEM', stringId: 7 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0.61,
    })
    const triangulated = EUROCONVERT(
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'FRF', stringId: 9 },
      { tag: ValueTag.String, value: 'DEM', stringId: 7 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Number, value: 3 },
    )
    expect(triangulated).toMatchObject({ tag: ValueTag.Number })
    if (triangulated.tag !== ValueTag.Number) {
      throw new Error('Expected EUROCONVERT result to be numeric')
    }
    expect(triangulated.value).toBeCloseTo(0.29728616, 12)
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'BAD', stringId: 10 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 32 },
        { tag: ValueTag.String, value: 'F', stringId: 11 },
        { tag: ValueTag.String, value: 'K', stringId: 12 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 273.15,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'F', stringId: 11 },
        { tag: ValueTag.String, value: 'm', stringId: 13 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      CONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: '??', stringId: 14 },
        { tag: ValueTag.String, value: 'm', stringId: 13 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 3.5 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
        { tag: ValueTag.String, value: 'EUR', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 3.5,
    })
    expect(
      EUROCONVERT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'FRF', stringId: 9 },
        { tag: ValueTag.String, value: 'DEM', stringId: 7 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports boolean and string builtins and builtin ids', () => {
    expect(getBuiltin('AND')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })

    expect(getBuiltin('OR')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })

    expect(getBuiltin('NOT')?.({ tag: ValueTag.Boolean, value: false })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })

    expect(getBuiltin('LEN')?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })

    expect(
      getBuiltin('CONCAT')?.(
        { tag: ValueTag.String, value: 'alpha', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'alpha2', stringId: 0 })

    expect(
      getBuiltin('EXACT')?.({ tag: ValueTag.String, value: 'Alpha', stringId: 1 }, { tag: ValueTag.String, value: 'Alpha', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Boolean, value: true })

    expect(
      getBuiltin('EXACT')?.({ tag: ValueTag.String, value: 'Alpha', stringId: 1 }, { tag: ValueTag.String, value: 'alpha', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Boolean, value: false })

    expect(getBuiltin('LEFT')?.({ tag: ValueTag.String, value: 'alpha', stringId: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: 'alp',
      stringId: 0,
    })

    expect(
      getBuiltin('TEXTBEFORE')?.(
        { tag: ValueTag.String, value: 'alpha-beta', stringId: 1 },
        { tag: ValueTag.String, value: '-', stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'alpha', stringId: 0 })

    expect(
      getBuiltin('IFERROR')?.({ tag: ValueTag.Error, code: ErrorCode.Div0 }, { tag: ValueTag.String, value: 'fallback', stringId: 1 }),
    ).toEqual({ tag: ValueTag.String, value: 'fallback', stringId: 1 })

    expect(
      getBuiltin('DATE')?.({ tag: ValueTag.Number, value: 2026 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 15 }),
    ).toEqual({ tag: ValueTag.Number, value: 46096 })

    expect(
      getBuiltin('AVERAGE')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 6 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      getBuiltin('AVERAGE')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltinId('sum')).toBe(BuiltinId.Sum)
    expect(getBuiltinId('concat')).toBe(BuiltinId.Concat)
    expect(getBuiltinId('choose')).toBe(BuiltinId.Choose)
    expect(getBuiltinId('countblank')).toBe(BuiltinId.Countblank)
    expect(getBuiltinId('lenb')).toBe(BuiltinId.Lenb)
    expect(getBuiltinId('leftb')).toBe(BuiltinId.Leftb)
    expect(getBuiltinId('midb')).toBe(BuiltinId.Midb)
    expect(getBuiltinId('rightb')).toBe(BuiltinId.Rightb)
    expect(getBuiltinId('findb')).toBe(BuiltinId.Findb)
    expect(getBuiltinId('searchb')).toBe(BuiltinId.Searchb)
    expect(getBuiltinId('replaceb')).toBe(BuiltinId.Replaceb)
    expect(getBuiltinId('address')).toBe(BuiltinId.Address)
    expect(getBuiltinId('days360')).toBe(BuiltinId.Days360)
    expect(getBuiltinId('dollar')).toBe(BuiltinId.Dollar)
    expect(getBuiltinId('dollarde')).toBe(BuiltinId.Dollarde)
    expect(getBuiltinId('dollarfr')).toBe(BuiltinId.Dollarfr)
    expect(getBuiltinId('yearfrac')).toBe(BuiltinId.Yearfrac)
    expect(getBuiltinId('isoweeknum')).toBe(BuiltinId.Isoweeknum)
    expect(getBuiltinId('timevalue')).toBe(BuiltinId.Timevalue)
    expect(getBuiltinId('textbefore')).toBe(BuiltinId.Textbefore)
    expect(getBuiltinId('textafter')).toBe(BuiltinId.Textafter)
    expect(getBuiltinId('textjoin')).toBe(BuiltinId.Textjoin)
    expect(getBuiltinId('textsplit')).toBe(BuiltinId.Textsplit)
    expect(getBuiltinId('correl')).toBe(BuiltinId.Correl)
    expect(getBuiltinId('covar')).toBe(BuiltinId.Covar)
    expect(getBuiltinId('pearson')).toBe(BuiltinId.Pearson)
    expect(getBuiltinId('covariance.p')).toBe(BuiltinId.CovarianceP)
    expect(getBuiltinId('covariance.s')).toBe(BuiltinId.CovarianceS)
    expect(getBuiltinId('forecast')).toBe(BuiltinId.Forecast)
    expect(getBuiltinId('intercept')).toBe(BuiltinId.Intercept)
    expect(getBuiltinId('median')).toBe(BuiltinId.Median)
    expect(getBuiltinId('small')).toBe(BuiltinId.Small)
    expect(getBuiltinId('large')).toBe(BuiltinId.Large)
    expect(getBuiltinId('percentile')).toBe(BuiltinId.Percentile)
    expect(getBuiltinId('percentile.inc')).toBe(BuiltinId.PercentileInc)
    expect(getBuiltinId('percentile.exc')).toBe(BuiltinId.PercentileExc)
    expect(getBuiltinId('percentrank')).toBe(BuiltinId.Percentrank)
    expect(getBuiltinId('percentrank.inc')).toBe(BuiltinId.PercentrankInc)
    expect(getBuiltinId('percentrank.exc')).toBe(BuiltinId.PercentrankExc)
    expect(getBuiltinId('quartile')).toBe(BuiltinId.Quartile)
    expect(getBuiltinId('quartile.inc')).toBe(BuiltinId.QuartileInc)
    expect(getBuiltinId('quartile.exc')).toBe(BuiltinId.QuartileExc)
    expect(getBuiltinId('mode.mult')).toBe(BuiltinId.ModeMult)
    expect(getBuiltinId('frequency')).toBe(BuiltinId.Frequency)
    expect(getBuiltinId('rank')).toBe(BuiltinId.Rank)
    expect(getBuiltinId('rank.eq')).toBe(BuiltinId.RankEq)
    expect(getBuiltinId('rank.avg')).toBe(BuiltinId.RankAvg)
    expect(getBuiltinId('rsq')).toBe(BuiltinId.Rsq)
    expect(getBuiltinId('slope')).toBe(BuiltinId.Slope)
    expect(getBuiltinId('steyx')).toBe(BuiltinId.Steyx)
    expect(getBuiltinId('disc')).toBe(BuiltinId.Disc)
    expect(getBuiltinId('intrate')).toBe(BuiltinId.Intrate)
    expect(getBuiltinId('received')).toBe(BuiltinId.Received)
    expect(getBuiltinId('irr')).toBe(BuiltinId.Irr)
    expect(getBuiltinId('mirr')).toBe(BuiltinId.Mirr)
    expect(getBuiltinId('xnpv')).toBe(BuiltinId.Xnpv)
    expect(getBuiltinId('xirr')).toBe(BuiltinId.Xirr)
    expect(getBuiltinId('base')).toBe(BuiltinId.Base)
    expect(getBuiltinId('decimal')).toBe(BuiltinId.Decimal)
    expect(getBuiltinId('convert')).toBe(BuiltinId.Convert)
    expect(getBuiltinId('euroconvert')).toBe(BuiltinId.Euroconvert)
    expect(getBuiltinId('bitand')).toBe(BuiltinId.Bitand)
    expect(getBuiltinId('bitor')).toBe(BuiltinId.Bitor)
    expect(getBuiltinId('bitxor')).toBe(BuiltinId.Bitxor)
    expect(getBuiltinId('bitlshift')).toBe(BuiltinId.Bitlshift)
    expect(getBuiltinId('bitrshift')).toBe(BuiltinId.Bitrshift)
    expect(getBuiltinId('besseli')).toBe(BuiltinId.Besseli)
    expect(getBuiltinId('besselj')).toBe(BuiltinId.Besselj)
    expect(getBuiltinId('besselk')).toBe(BuiltinId.Besselk)
    expect(getBuiltinId('bessely')).toBe(BuiltinId.Bessely)
    expect(getBuiltinId('use.the.countif')).toBe(BuiltinId.Countif)
    expect(getBuiltinId('')).toBeUndefined()
    expect(getBuiltin('missing')).toBeUndefined()
  })

  it('supports the remaining scalar numeric builtins and conditional defaults', () => {
    expect(getBuiltin('MIN')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })

    expect(
      getBuiltin('MAX')?.(
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('MAX')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    expect(
      getBuiltin('COUNT')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(
      getBuiltin('COUNTA')?.(
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'x', stringId: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(getBuiltin('ABS')?.({ tag: ValueTag.Number, value: -3.4 })).toEqual({
      tag: ValueTag.Number,
      value: 3.4,
    })
    expect(getBuiltin('ABS')?.({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('ABS')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('INT')?.({ tag: ValueTag.Number, value: -3.1 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('INT')?.({ tag: ValueTag.String, value: '0008', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(getBuiltin('ROUND')?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3.15,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 0.07 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0.07,
    })
    expect(getBuiltin('ROUNDDOWN')?.({ tag: ValueTag.Number, value: -3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -3.14,
    })
    expect(getBuiltin('ROUNDDOWN')?.({ tag: ValueTag.Number, value: 0.29 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0.29,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: 16.4 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 16.4,
    })
    expect(getBuiltin('ROUND')?.({ tag: ValueTag.Number, value: 3.145 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3.15,
    })
    const refError = { tag: ValueTag.Error, code: ErrorCode.Ref } as const
    expect(getBuiltin('ROUND')?.(refError, { tag: ValueTag.Number, value: 0 })).toEqual(refError)
    expect(getBuiltin('ROUND')?.({ tag: ValueTag.Number, value: 3.145 }, refError)).toEqual(refError)
    expect(getBuiltin('ROUNDUP')?.(refError, { tag: ValueTag.Number, value: 0 })).toEqual(refError)
    expect(getBuiltin('ROUNDDOWN')?.(refError, { tag: ValueTag.Number, value: 0 })).toEqual(refError)
    expect(getBuiltin('TRUNC')?.(refError, { tag: ValueTag.Number, value: 0 })).toEqual(refError)
    expect(getBuiltin('MROUND')?.(refError, { tag: ValueTag.Number, value: 2 })).toEqual(refError)
    expect(getBuiltin('FLOOR')?.({ tag: ValueTag.Number, value: 3.6 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('FLOOR')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('CEILING')?.({ tag: ValueTag.Number, value: 3.1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('CEILING')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })

    expect(getBuiltin('IF')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'truthy', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'truthy',
      stringId: 1,
    })

    expect(getBuiltin('IF')?.({ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Boolean, value: false })
  })

  it('supports radix conversion and complex engineering builtins', () => {
    expect(getBuiltin('BIN2DEC')?.({ tag: ValueTag.String, value: '1111111111', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(getBuiltin('DEC2BIN')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: 0,
    })
    expect(getBuiltin('BIN2HEX')?.({ tag: ValueTag.String, value: '1111111111', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'FFFFFFFFFF',
      stringId: 0,
    })
    expect(getBuiltin('HEX2BIN')?.({ tag: ValueTag.String, value: 'A', stringId: 1 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: 0,
    })
    expect(getBuiltin('HEX2DEC')?.({ tag: ValueTag.String, value: 'FFFFFFFFFF', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(getBuiltin('OCT2HEX')?.({ tag: ValueTag.String, value: '17', stringId: 1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.String,
      value: '000F',
      stringId: 0,
    })

    expect(
      getBuiltin('COMPLEX')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: -4 },
        { tag: ValueTag.String, value: 'j', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: '3-4j', stringId: 0 })
    expect(getBuiltin('IMREAL')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('IMAGINARY')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(getBuiltin('IMABS')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(getBuiltin('IMARGUMENT')?.({ tag: ValueTag.String, value: 'i', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 2,
    })
    expect(getBuiltin('IMCONJUGATE')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '3-4i',
      stringId: 0,
    })
    expect(
      getBuiltin('IMSUM')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 }, { tag: ValueTag.String, value: '-1+2i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '2+6i', stringId: 0 })
    expect(
      getBuiltin('IMPRODUCT')?.({ tag: ValueTag.String, value: '1+i', stringId: 1 }, { tag: ValueTag.String, value: '1-i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '2', stringId: 0 })
    expect(
      getBuiltin('IMDIV')?.({ tag: ValueTag.String, value: '3+4i', stringId: 1 }, { tag: ValueTag.String, value: '1-i', stringId: 2 }),
    ).toEqual({ tag: ValueTag.String, value: '-0.5+3.5i', stringId: 0 })
    expect(getBuiltin('IMSQRT')?.({ tag: ValueTag.String, value: '-4', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '2i',
      stringId: 0,
    })
    expect(getBuiltin('IMSIN')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMCOS')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMSECH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
  })

  it('covers the remaining complex engineering and value-classification builtins', () => {
    expect(getBuiltin('IMEXP')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMLN')?.({ tag: ValueTag.String, value: '1', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMLOG10')?.({ tag: ValueTag.String, value: '10', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMLOG2')?.({ tag: ValueTag.String, value: '8', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '3',
      stringId: 0,
    })
    expect(getBuiltin('IMPOWER')?.({ tag: ValueTag.String, value: '2', stringId: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: '8',
      stringId: 0,
    })
    expect(getBuiltin('IMTAN')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMSINH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '0',
      stringId: 0,
    })
    expect(getBuiltin('IMCOSH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMSEC')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: '1',
      stringId: 0,
    })
    expect(getBuiltin('IMCSC')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMCOT')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMCSCH')?.({ tag: ValueTag.String, value: '0', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('IMLOG10')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('IMPOWER')?.({ tag: ValueTag.String, value: '1+i', stringId: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 1999 })).toEqual({
      tag: ValueTag.String,
      value: 'MCMXCIX',
      stringId: 0,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'MCMXCIX', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 1999,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('T')?.({ tag: ValueTag.String, value: 'text', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'text',
      stringId: 1,
    })
    expect(getBuiltin('T')?.({ tag: ValueTag.Number, value: 7 })).toEqual({
      tag: ValueTag.String,
      value: '',
      stringId: 0,
    })
    expect(getBuiltin('T')?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.String,
      value: '',
      stringId: 0,
    })
    expect(getBuiltin('T')?.({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(getBuiltin('ISOMITTED')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(getBuiltin('ISOMITTED')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('N')?.({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('N')?.({ tag: ValueTag.String, value: 'text', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('N')?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    const matrix: ArrayValue = {
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ],
    }
    expect(getBuiltin('TYPE')?.({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    expect(Reflect.apply(getBuiltin('TYPE')!, undefined, [matrix])).toEqual({
      tag: ValueTag.Number,
      value: 64,
    })
    expect(getBuiltin('DELTA')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('GESTEP')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('supports expanded math and numeric utility builtins', () => {
    expect(getBuiltin('SIN')?.({ tag: ValueTag.Number, value: Math.PI / 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('COS')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: -32 }, { tag: ValueTag.Number, value: 1 / 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.Number, value: -9 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: -3.98 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: -3.9,
    })
    expect(getBuiltin('CEILING.MATH')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: 5 })).toEqual({
      tag: ValueTag.Number,
      value: 120,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(getBuiltin('GCD')?.({ tag: ValueTag.Number, value: 18 }, { tag: ValueTag.Number, value: 24 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('LCM')?.({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 8 })).toEqual({
      tag: ValueTag.Number,
      value: 24,
    })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 9,
    })
    expect(
      getBuiltin('PRODUCT')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('SUMSQ')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 13,
    })
    expect(
      getBuiltin('SERIESSUM')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 18 })
    expect(
      getBuiltin('BASE')?.({ tag: ValueTag.Number, value: 31 }, { tag: ValueTag.Number, value: 16 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.String, value: '001F', stringId: 0 })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '1F', stringId: 1 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 31,
    })
    expect(getBuiltin('BASE')?.({ tag: ValueTag.Number, value: 15 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.String,
      value: '1111',
      stringId: 0,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '  1111  ', stringId: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 15,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '', stringId: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.String,
      value: 'XIV',
      stringId: 0,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'XIV', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 14,
    })

    expect(getBuiltin('MUNIT')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    })

    const randArray = getBuiltin('RANDARRAY')?.(
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Boolean, value: true },
    )
    expect(randArray).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    if (!(randArray && 'kind' in randArray && randArray.kind === 'array')) {
      throw new Error('expected RANDARRAY to return an array')
    }
    for (const value of randArray.values) {
      expect(value.tag).toBe(ValueTag.Number)
      expect(value.value).toBeGreaterThanOrEqual(3)
      expect(value.value).toBeLessThanOrEqual(7)
    }
  })

  it('covers math builtin edge cases and aggregate variants', () => {
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('EVEN')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ODD')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('LN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('SQRT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('FACTDOUBLE')?.({ tag: ValueTag.Number, value: 6 })).toEqual({
      tag: ValueTag.Number,
      value: 48,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('GCD')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('LCM')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: -10 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('MULTINOMIAL')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('RANDBETWEEN')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('RANDBETWEEN')?.({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('BASE')?.({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.String, value: '1Z', stringId: 1 }, { tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('DECIMAL')?.({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('ROMAN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'IIII', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.String, value: 'VX', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('RANDARRAY')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('RANDARRAY')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({ kind: 'array', rows: 2, cols: 2 })
    expect(getBuiltin('MUNIT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('MUNIT')?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('RANDARRAY')?.({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      getBuiltin('SERIESSUM')?.(
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('SQRTPI')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('SQRTPI')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 9 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(getBuiltin('SUBTOTAL')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 99 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Error, code: ErrorCode.Div0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Error, code: ErrorCode.Div0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(getBuiltin('ARABIC')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers address, formatting, and mean helper builtins', () => {
    expect(
      getBuiltin('MAXA')?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      getBuiltin('MINA')?.(
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.String, value: 'skip', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: "O'Brien", stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.String, value: "'O''Brien'!$AB2", stringId: 0 })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.String, value: 'R2C[28]', stringId: 0 })
    expect(
      getBuiltin('ADDRESS')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 28 }, { tag: ValueTag.Number, value: 5 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 28 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: -1234.567 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.6',
      stringId: 0,
    })
    expect(
      getBuiltin('DOLLAR')?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('FIXED')?.(
        { tag: ValueTag.Number, value: 1234.567 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: '1234.6', stringId: 0 })
    expect(
      getBuiltin('FIXED')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 127 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.String, value: `1.${'0'.repeat(127)}`, stringId: 0 })
    expect(
      getBuiltin('FIXED')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 128 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('FIXED')?.({ tag: ValueTag.Number, value: 1234.567 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.08 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 1.5,
    })
    expect(getBuiltin('DOLLARFR')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 1.08,
    })
    expect(getBuiltin('DOLLARFR')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(getBuiltin('GEOMEAN')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(getBuiltin('GEOMEAN')?.({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 3 / 1.75 })
    expect(getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('covers extended trigonometric and precise rounding builtins', () => {
    expect(getBuiltin('SINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COSH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('TANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASINH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ACOSH')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ATANH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ACOT')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 4,
    })
    expect(getBuiltin('ACOT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: Math.PI / 2,
    })
    expect(getBuiltin('ACOTH')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5 * Math.log(3),
    })
    expect(getBuiltin('COTH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSCH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('SEC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SECH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('FLOOR.MATH')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: 5.1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 5.1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })
})
