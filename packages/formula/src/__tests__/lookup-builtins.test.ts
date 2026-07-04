import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { compileCriteriaMatcher, getLookupBuiltin, matchesCompiledCriteria, type RangeBuiltinArgument } from '../builtins/lookup.js'

const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const err = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })
const empty = (): CellValue => ({ tag: ValueTag.Empty })

function cellRange(values: CellValue[], rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind: 'cells', values, rows, cols }
}

function axisRange(values: CellValue[], refKind: 'rows' | 'cols', rows: number, cols: number): RangeBuiltinArgument {
  return { kind: 'range', refKind, values, rows, cols }
}

describe('lookup builtins: lookup, database, and conditional aggregate builtins', () => {
  it('supports MATCH across one-dimensional cell ranges', () => {
    const MATCH = getLookupBuiltin('MATCH')!
    expect(MATCH(text('b'), cellRange([text('a'), text('b'), text('c')], 3, 1), num(0))).toEqual(num(2))
    expect(MATCH(text('b?t'), cellRange([text('bat'), text('bot'), text('cat')], 3, 1), num(0))).toEqual(num(1))
    expect(MATCH(text('b~?t'), cellRange([text('bat'), text('b?t'), text('bot')], 3, 1), num(0))).toEqual(num(2))
    expect(MATCH(num(4), cellRange([num(1), num(3), num(5)], 3, 1), num(1))).toEqual(num(2))
    expect(MATCH(num(3), cellRange([num(5), num(3), num(1)], 3, 1), num(-1))).toEqual(num(2))
    expect(MATCH(text('z'), cellRange([text('a'), text('b')], 2, 1), num(0))).toEqual(err(ErrorCode.NA))
    expect(MATCH(text('x'), cellRange([text('a'), text('b'), text('c'), text('d')], 2, 2), num(0))).toEqual(err(ErrorCode.NA))
  })

  it('supports AREAS, ARRAYTOTEXT, ROWS, COLUMNS, and CORREL', () => {
    const AREAS = getLookupBuiltin('AREAS')!
    const ARRAYTOTEXT = getLookupBuiltin('ARRAYTOTEXT')!
    const ROWS = getLookupBuiltin('ROWS')!
    const COLUMNS = getLookupBuiltin('COLUMNS')!
    const CORREL = getLookupBuiltin('CORREL')!

    const matrix = cellRange([num(1), text('x'), num(3), num(4)], 2, 2)

    expect(AREAS(matrix)).toEqual(num(1))
    expect(ROWS(matrix)).toEqual(num(2))
    expect(COLUMNS(matrix)).toEqual(num(2))
    expect(ARRAYTOTEXT(matrix)).toEqual(text('1\tx;3\t4'))
    expect(ARRAYTOTEXT(matrix, num(1))).toEqual(text('{1, "x";3, 4}'))
    expect(CORREL(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(num(1))
    expect(CORREL(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Value))
  })

  it('maps USE.THE.COUNTIF to the COUNTIF lookup implementation', () => {
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const alias = getLookupBuiltin('USE.THE.COUNTIF')!
    const sample = cellRange([num(1), num(0), num(3)], 3, 1)

    expect(alias).toBe(COUNTIF)
    expect(alias(sample, text('>0'))).toEqual(num(2))
  })

  it('compiles and matches direct criteria helpers across wildcard, escaped, and scalar comparisons', () => {
    const wildcard = compileCriteriaMatcher(text('ap*'))
    expect(matchesCompiledCriteria(text('apple'), wildcard)).toBe(true)
    expect(matchesCompiledCriteria(text('pear'), wildcard)).toBe(false)
    expect(matchesCompiledCriteria(empty(), wildcard)).toBe(false)
    expect(matchesCompiledCriteria(num(123), wildcard)).toBe(false)

    const negatedWildcard = compileCriteriaMatcher(text('<>ap*'))
    expect(matchesCompiledCriteria(text('apple'), negatedWildcard)).toBe(false)
    expect(matchesCompiledCriteria(text('pear'), negatedWildcard)).toBe(true)
    expect(matchesCompiledCriteria(empty(), negatedWildcard)).toBe(true)
    expect(matchesCompiledCriteria(num(123), negatedWildcard)).toBe(true)

    const escapedPattern = compileCriteriaMatcher(text('a~*b'))
    expect(escapedPattern.operand).toMatchObject({ tag: ValueTag.String, value: 'a*b' })
    expect(matchesCompiledCriteria(text('a*b'), escapedPattern)).toBe(true)

    const greaterThan = compileCriteriaMatcher(text('>2'))
    expect(matchesCompiledCriteria(num(3), greaterThan)).toBe(true)
    expect(matchesCompiledCriteria(num(2), greaterThan)).toBe(false)
    expect(matchesCompiledCriteria(text('3'), greaterThan)).toBe(false)
    expect(matchesCompiledCriteria(err(ErrorCode.Div0), greaterThan)).toBe(false)

    const groupedThousands = compileCriteriaMatcher(text('>=200,000'))
    expect(groupedThousands.operand).toMatchObject({ tag: ValueTag.Number, value: 200_000 })
    expect(matchesCompiledCriteria(num(250_000), groupedThousands)).toBe(true)
    expect(matchesCompiledCriteria(num(50_000), groupedThousands)).toBe(false)

    const invalidGrouping = compileCriteriaMatcher(text('>=20,00'))
    expect(invalidGrouping.operand).toMatchObject({ tag: ValueTag.String, value: '20,00' })
    expect(matchesCompiledCriteria(num(20_000), invalidGrouping)).toBe(false)

    const scalarEquality = compileCriteriaMatcher(bool(true))
    expect(matchesCompiledCriteria(bool(true), scalarEquality)).toBe(true)
    expect(matchesCompiledCriteria(bool(false), scalarEquality)).toBe(false)
  })

  it('supports database aggregation builtins on matching records', () => {
    const DAVERAGE = getLookupBuiltin('DAVERAGE')!
    const DCOUNT = getLookupBuiltin('DCOUNT')!
    const DCOUNTA = getLookupBuiltin('DCOUNTA')!
    const DGET = getLookupBuiltin('DGET')!
    const DMAX = getLookupBuiltin('DMAX')!
    const DMIN = getLookupBuiltin('DMIN')!
    const DPRODUCT = getLookupBuiltin('DPRODUCT')!
    const DSTDEV = getLookupBuiltin('DSTDEV')!
    const DSTDEVP = getLookupBuiltin('DSTDEVP')!
    const DSUM = getLookupBuiltin('DSUM')!
    const DVAR = getLookupBuiltin('DVAR')!
    const DVARP = getLookupBuiltin('DVARP')!

    const database = cellRange(
      [
        text('Age'),
        text('Height'),
        text('Yield'),
        num(10),
        num(100),
        num(5),
        num(12),
        num(110),
        num(7),
        num(12),
        num(120),
        num(9),
        num(15),
        num(130),
        num(11),
      ],
      5,
      3,
    )
    const ageIsTwelve = cellRange([text('Age'), num(12)], 2, 1)
    const ageIsFifteen = cellRange([text('Age'), num(15)], 2, 1)

    expect(DAVERAGE(database, text('Yield'), ageIsTwelve)).toEqual(num(8))
    expect(DCOUNT(database, text('Yield'), ageIsTwelve)).toEqual(num(2))
    expect(DCOUNT(database, { tag: ValueTag.Empty }, ageIsTwelve)).toEqual(num(2))
    expect(DCOUNTA(database, text('Height'), ageIsTwelve)).toEqual(num(2))
    expect(DCOUNTA(database, { tag: ValueTag.Empty }, ageIsTwelve)).toEqual(num(2))
    expect(DGET(database, text('Height'), ageIsFifteen)).toEqual(num(130))
    expect(DMAX(database, text('Yield'), ageIsTwelve)).toEqual(num(9))
    expect(DMIN(database, text('Yield'), ageIsTwelve)).toEqual(num(7))
    expect(DPRODUCT(database, text('Yield'), ageIsTwelve)).toEqual(num(63))
    expect(DSUM(database, text('Yield'), ageIsTwelve)).toEqual(num(16))
    expect(DVAR(database, text('Yield'), ageIsTwelve)).toEqual(num(2))
    expect(DVARP(database, text('Yield'), ageIsTwelve)).toEqual(num(1))

    const dstdev = DSTDEV(database, text('Yield'), ageIsTwelve)
    if (dstdev.tag !== ValueTag.Number) {
      throw new Error('DSTDEV should return a number')
    }
    expect(dstdev.value).toBeCloseTo(Math.SQRT2, 12)

    const dstdevp = DSTDEVP(database, text('Yield'), ageIsTwelve)
    if (dstdevp.tag !== ValueTag.Number) {
      throw new Error('DSTDEVP should return a number')
    }
    expect(dstdevp.value).toBeCloseTo(1, 12)

    expect(DGET(database, text('Yield'), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DAVERAGE(database, text('Missing'), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, text('Yield'), cellRange([text('Age'), err(ErrorCode.Ref)], 2, 1))).toEqual(err(ErrorCode.Ref))
  })

  it('covers database builtin validation and empty-match edge cases', () => {
    const DAVERAGE = getLookupBuiltin('DAVERAGE')!
    const DCOUNT = getLookupBuiltin('DCOUNT')!
    const DGET = getLookupBuiltin('DGET')!
    const DMAX = getLookupBuiltin('DMAX')!
    const DPRODUCT = getLookupBuiltin('DPRODUCT')!
    const DSTDEVP = getLookupBuiltin('DSTDEVP')!
    const DVAR = getLookupBuiltin('DVAR')!
    const DVARP = getLookupBuiltin('DVARP')!

    const database = cellRange(
      [
        text('Age'),
        text('Height'),
        text('Yield'),
        num(10),
        num(100),
        num(5),
        num(12),
        num(110),
        num(7),
        num(12),
        num(120),
        num(9),
        num(15),
        num(130),
        num(11),
      ],
      5,
      3,
    )
    const ageIsTwelve = cellRange([text('Age'), num(12)], 2, 1)
    const ageMissing = cellRange([text('Age'), num(99)], 2, 1)

    expect(DGET(database, cellRange([text('Height')], 1, 1), cellRange([text('Age'), num(15)], 2, 1))).toEqual(num(130))
    expect(DCOUNT(database, text(''), ageIsTwelve)).toEqual(num(2))
    expect(DAVERAGE(database, text(''), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, cellRange([text('Yield'), text('Height')], 1, 2), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DMAX(database, bool(true), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DMAX(database, num(0), ageIsTwelve)).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, text('Yield'), cellRange([text('Age')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(DCOUNT(database, text('Yield'), cellRange([err(ErrorCode.Ref), num(12)], 2, 1))).toEqual(err(ErrorCode.Ref))
    expect(DCOUNT(database, text('Yield'), cellRange([text('Age'), err(ErrorCode.Name)], 2, 1))).toEqual(err(ErrorCode.Name))
    expect(DCOUNT(database, text('Yield'), cellRange([{ tag: ValueTag.Empty }, num(12)], 2, 1))).toEqual(num(0))
    expect(DCOUNT(database, text('Yield'), cellRange([text('Missing'), num(12)], 2, 1))).toEqual(num(0))
    expect(DAVERAGE(database, text('Yield'), ageMissing)).toEqual(err(ErrorCode.Div0))
    expect(DMAX(database, text('Yield'), ageMissing)).toEqual(num(0))
    expect(DPRODUCT(database, text('Yield'), ageMissing)).toEqual(num(0))
    expect(DSTDEVP(database, text('Yield'), ageMissing)).toEqual(err(ErrorCode.Div0))
    expect(DVAR(database, text('Yield'), cellRange([text('Age'), err(ErrorCode.Ref)], 2, 1))).toEqual(err(ErrorCode.Ref))
    expect(DVARP(database, text('Missing'), ageIsTwelve)).toEqual(err(ErrorCode.Value))
  })

  it('supports case-insensitive database headers with OR criteria rows', () => {
    const DCOUNT = getLookupBuiltin('DCOUNT')!
    const DSUM = getLookupBuiltin('DSUM')!

    const database = cellRange(
      [
        text('Tree'),
        text('Yield'),
        text('Height'),
        text('Apple'),
        num(10),
        num(18),
        text('Pear'),
        num(14),
        num(20),
        text('Apple'),
        num(16),
        num(22),
        text('Plum'),
        num(8),
        num(16),
      ],
      5,
      3,
    )
    const criteria = cellRange([text('tree'), text('yield'), text('apple'), text('>12'), text('pear'), { tag: ValueTag.Empty }], 3, 2)

    expect(DCOUNT(database, text('Yield'), criteria)).toEqual(num(2))
    expect(DSUM(database, text('Yield'), criteria)).toEqual(num(30))
  })

  it('supports COVAR, COVARIANCE.P, COVARIANCE.S, AVEDEV, and DEVSQ', () => {
    const COVAR = getLookupBuiltin('COVAR')!
    const COVARP = getLookupBuiltin('COVARIANCE.P')!
    const COVARS = getLookupBuiltin('COVARIANCE.S')!
    const AVEDEV = getLookupBuiltin('AVEDEV')!
    const DEVSQ = getLookupBuiltin('DEVSQ')!

    const first = cellRange([num(1), num(2), num(3)], 3, 1)
    const second = cellRange([num(4), num(5), num(6)], 3, 1)
    expect(COVAR(first, second)).toEqual(num(2 / 3))
    expect(COVARP(first, second)).toEqual(num(2 / 3))
    expect(COVARS(first, second)).toEqual(num(1))

    expect(COVAR(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1), num(2), num(3)], 3, 1))).toEqual(
      err(ErrorCode.Value),
    )

    expect(COVARS(cellRange([num(2)], 1, 1), cellRange([num(4)], 1, 1))).toEqual(err(ErrorCode.Div0))

    expect(AVEDEV(num(1), num(2), num(3))).toEqual(num(2 / 3))
    expect(DEVSQ(num(1), num(2), num(3))).toEqual(num(2))
    expect(AVEDEV(cellRange([text('bad')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(DEVSQ(cellRange([text('bad')], 1, 1))).toEqual(err(ErrorCode.Value))
  })

  it('supports CHISQ.TEST and legacy aliases across actual and expected matrices', () => {
    const CHISQ_TEST = getLookupBuiltin('CHISQ.TEST')!
    const CHITEST = getLookupBuiltin('CHITEST')!
    const LEGACY_CHITEST = getLookupBuiltin('LEGACY.CHITEST')!

    const actual = cellRange([num(58), num(35), num(11), num(25), num(10), num(23)], 3, 2)
    const expected = cellRange([num(45.35), num(47.65), num(17.56), num(18.44), num(16.09), num(16.91)], 3, 2)

    const chisqTest = CHISQ_TEST(actual, expected)
    if (chisqTest.tag !== ValueTag.Number) {
      throw new Error('CHISQ.TEST should return a number')
    }
    expect(chisqTest.value).toBeCloseTo(0.0003082, 7)
    const chiTest = CHITEST(actual, expected)
    if (chiTest.tag !== ValueTag.Number) {
      throw new Error('CHITEST should return a number')
    }
    expect(chiTest.value).toBeCloseTo(0.0003082, 7)
    const legacyChiTest = LEGACY_CHITEST(actual, expected)
    if (legacyChiTest.tag !== ValueTag.Number) {
      throw new Error('LEGACY.CHITEST should return a number')
    }
    expect(legacyChiTest.value).toBeCloseTo(0.0003082, 7)

    expect(CHISQ_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.NA))
    expect(CHISQ_TEST(num(1), num(1))).toEqual(err(ErrorCode.NA))
    expect(CHISQ_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1), num(0)], 2, 1))).toEqual(err(ErrorCode.Div0))
  })

  it('supports F.TEST and Z.TEST legacy aliases on numeric samples', () => {
    const F_TEST = getLookupBuiltin('F.TEST')!
    const FTEST = getLookupBuiltin('FTEST')!
    const Z_TEST = getLookupBuiltin('Z.TEST')!
    const ZTEST = getLookupBuiltin('ZTEST')!

    const first = cellRange([num(6), num(7), num(9), num(15), num(21)], 5, 1)
    const second = cellRange([num(20), num(28), num(31), num(38), num(40)], 5, 1)
    const zSample = cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1)

    const fTest = F_TEST(first, second)
    if (fTest.tag !== ValueTag.Number) {
      throw new Error('F.TEST should return a number')
    }
    expect(fTest.value).toBeCloseTo(0.648317846786175, 12)

    const legacyFTest = FTEST(first, second)
    if (legacyFTest.tag !== ValueTag.Number) {
      throw new Error('FTEST should return a number')
    }
    expect(legacyFTest.value).toBeCloseTo(0.648317846786175, 12)

    const zTest = Z_TEST(zSample, num(2), num(1))
    if (zTest.tag !== ValueTag.Number) {
      throw new Error('Z.TEST should return a number')
    }
    expect(zTest.value).toBeCloseTo(0.012673659338733989, 12)

    const legacyZTest = ZTEST(zSample, num(2), num(1))
    if (legacyZTest.tag !== ValueTag.Number) {
      throw new Error('ZTEST should return a number')
    }
    expect(legacyZTest.value).toBeCloseTo(0.012673659338733989, 12)

    expect(F_TEST(cellRange([num(1), text('x')], 2, 1), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Div0))
    expect(F_TEST(cellRange([num(3), num(3)], 2, 1), cellRange([num(1), num(2)], 2, 1))).toEqual(err(ErrorCode.Div0))
    expect(Z_TEST(zSample, num(2), num(0))).toEqual(err(ErrorCode.Div0))
  })

  it('supports T.TEST across paired, equal-variance, and Welch modes', () => {
    const T_TEST = getLookupBuiltin('T.TEST')!
    const TTEST = getLookupBuiltin('TTEST')!

    const pairedFirst = cellRange([num(1), num(2), num(4)], 3, 1)
    const pairedSecond = cellRange([num(1), num(3), num(3)], 3, 1)
    expect(T_TEST(pairedFirst, pairedSecond, num(2), num(1))).toEqual(num(1))
    expect(TTEST(pairedFirst, pairedSecond, num(2), num(1))).toEqual(num(1))

    const independentFirst = cellRange([num(6), num(7), num(9), num(15), num(21)], 5, 1)
    const independentSecond = cellRange([num(20), num(28), num(31), num(38), num(40)], 5, 1)
    const equalVariance = T_TEST(independentFirst, independentSecond, num(2), num(2))
    if (equalVariance.tag !== ValueTag.Number) {
      throw new Error('T.TEST equal-variance mode should return a number')
    }
    expect(equalVariance.value).toBeCloseTo(0.0025154774780675737, 12)

    const welch = T_TEST(independentFirst, independentSecond, num(2), num(3))
    if (welch.tag !== ValueTag.Number) {
      throw new Error('T.TEST Welch mode should return a number')
    }
    expect(welch.value).toBeGreaterThan(equalVariance.value)
    expect(welch.value).toBeLessThan(0.01)

    expect(T_TEST(independentFirst, independentSecond, num(3), num(2))).toEqual(err(ErrorCode.Value))
    expect(T_TEST(independentFirst, independentSecond, num(2), num(4))).toEqual(err(ErrorCode.Value))
    expect(T_TEST(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1), num(2), num(1))).toEqual(err(ErrorCode.NA))
  })

  it('supports LOOKUP, TRANSPOSE, HSTACK, VSTACK, and PEARSON', () => {
    const LOOKUP = getLookupBuiltin('LOOKUP')!
    const TRANSPOSE = getLookupBuiltin('TRANSPOSE')!
    const HSTACK = getLookupBuiltin('HSTACK')!
    const VSTACK = getLookupBuiltin('VSTACK')!
    const PEARSON = getLookupBuiltin('PEARSON')!

    const lookupValues = cellRange([num(1), num(2), num(3)], 3, 1)
    const resultValues = cellRange([num(10), num(20), num(30)], 3, 1)
    const xlsxFixtureCorpusThresholds = cellRange(
      [
        num(0),
        num(0.199),
        num(0.299),
        num(0.599),
        num(1.1),
        num(2.1),
        num(4.1),
        num(8.1),
        num(16.1),
        num(32.1),
        num(64.1),
        num(128.1),
        num(256.1),
        num(512.1),
        num(1024.1),
        num(2048),
      ],
      16,
      1,
    )
    const xlsxFixtureCorpusScores = cellRange(
      [num(1), num(2), num(3), num(4), num(5), num(6), num(7), num(8), num(9), num(10), num(11), num(12), num(13), num(14), num(15)],
      15,
      1,
    )

    expect(LOOKUP(num(2), lookupValues, resultValues)).toEqual(num(20))
    expect(LOOKUP(num(4), lookupValues, resultValues)).toEqual(num(30))
    expect(
      LOOKUP(
        text('office'),
        axisRange([text('agency'), text('office'), text('port')], 'cols', 3, 1),
        axisRange([text('AGY'), text('AUD'), text('PRT')], 'cols', 3, 1),
      ),
    ).toEqual(text('AUD'))
    expect(LOOKUP(num(0.6666666666666666), xlsxFixtureCorpusThresholds, xlsxFixtureCorpusScores)).toEqual(num(4))
    expect(LOOKUP(num(2048), xlsxFixtureCorpusThresholds, xlsxFixtureCorpusScores)).toEqual(err(ErrorCode.NA))
    expect(
      LOOKUP(text('not-found'), cellRange([text('a'), text('b'), text('c')], 3, 1), cellRange([num(1), num(2), num(3)], 3, 1)),
    ).toEqual(err(ErrorCode.NA))

    expect(TRANSPOSE(cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 2, 3))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [num(1), num(4), num(2), num(5), num(3), num(6)],
    })
    expect(TRANSPOSE(num(7))).toEqual(num(7))

    expect(HSTACK(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([text('a'), text('b')], 1, 2), num(99))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 4,
      values: [
        num(1),
        text('a'),
        text('b'),
        num(99),
        num(2),
        err(ErrorCode.NA),
        err(ErrorCode.NA),
        err(ErrorCode.NA),
        num(3),
        err(ErrorCode.NA),
        err(ErrorCode.NA),
        err(ErrorCode.NA),
      ],
    })

    expect(VSTACK(cellRange([text('x'), text('y')], 1, 2), cellRange([num(3), num(4), num(5), num(6)], 2, 2), num(7))).toEqual({
      kind: 'array',
      rows: 4,
      cols: 2,
      values: [text('x'), text('y'), num(3), num(4), num(5), num(6), num(7), err(ErrorCode.NA)],
    })

    expect(PEARSON(lookupValues, resultValues)).toEqual(num(1))
    expect(PEARSON(cellRange([num(1)], 1, 1), cellRange([num(2)], 1, 1))).toEqual(err(ErrorCode.Div0))
  })

  it('validates TOCOL and TOROW control arguments', () => {
    const TOCOL = getLookupBuiltin('TOCOL')!
    const TOROW = getLookupBuiltin('TOROW')!
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2)

    expect(TOCOL(matrix, num(4))).toEqual(err(ErrorCode.Value))
    expect(TOCOL(matrix, num(0), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(TOROW(matrix, num(4))).toEqual(err(ErrorCode.Value))
    expect(TOROW(matrix, num(0), text('bad'))).toEqual(err(ErrorCode.Value))
  })

  it('supports INDEX over cell ranges', () => {
    const INDEX = getLookupBuiltin('INDEX')!
    const matrix = cellRange([num(10), num(11), num(20), num(21)], 2, 2)
    expect(INDEX(matrix, num(2), num(1))).toEqual(num(20))
    expect(INDEX(cellRange([text('a'), text('b'), text('c')], 1, 3), num(2))).toEqual(text('b'))
    expect(INDEX(matrix, num(3), num(1))).toEqual(err(ErrorCode.Ref))
    expect(INDEX(matrix, text('oops'))).toEqual(err(ErrorCode.Value))
  })

  it('supports exact and approximate VLOOKUP', () => {
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const table = cellRange([text('a'), num(10), text('b'), num(20), text('c'), num(30)], 3, 2)

    expect(VLOOKUP(text('b'), table, num(2), bool(false))).toEqual(num(20))
    expect(VLOOKUP(text('bb'), table, num(2), bool(true))).toEqual(num(20))
    expect(VLOOKUP(text('z'), table, num(2), bool(false))).toEqual(err(ErrorCode.NA))
    expect(VLOOKUP(text('a'), table, num(3), bool(false))).toEqual(err(ErrorCode.Value))
  })

  it('prefers exact matches before approximate VLOOKUP and HLOOKUP fallbacks', () => {
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!

    expect(VLOOKUP(text('b'), cellRange([text('a'), num(10), text('c'), num(30), text('b'), num(20)], 3, 2), num(2))).toEqual(num(20))
    expect(HLOOKUP(text('b'), cellRange([text('a'), text('c'), text('b'), num(10), num(30), num(20)], 2, 3), num(2))).toEqual(num(20))
  })

  it('skips incomparable keys during exact VLOOKUP and HLOOKUP scans', () => {
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!

    expect(VLOOKUP(text('b'), cellRange([num(1), text('ignored'), text('b'), num(20)], 2, 2), num(2), bool(false))).toEqual(num(20))
    expect(HLOOKUP(text('b'), cellRange([num(1), text('b'), text('ignored'), num(20)], 2, 2), num(2), bool(false))).toEqual(num(20))
  })

  it('skips blank keys during approximate VLOOKUP and HLOOKUP scans', () => {
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!

    expect(
      VLOOKUP(
        text('United States of America'),
        cellRange(
          [
            text('Country'),
            num(0),
            { tag: ValueTag.Empty },
            num(0),
            text('United States of America'),
            num(0.0446),
            text('Zimbabwe'),
            num(0.1589),
          ],
          4,
          2,
        ),
        num(2),
      ),
    ).toEqual(num(0.0446))
    expect(
      HLOOKUP(
        text('United States of America'),
        cellRange(
          [
            text('Country'),
            { tag: ValueTag.Empty },
            text('United States of America'),
            text('Zimbabwe'),
            num(0),
            num(0),
            num(0.0446),
            num(0.1589),
          ],
          2,
          4,
        ),
        num(2),
      ),
    ).toEqual(num(0.0446))
  })

  it('coerces blank lookup return cells to zero', () => {
    const LOOKUP = getLookupBuiltin('LOOKUP')!
    const VLOOKUP = getLookupBuiltin('VLOOKUP')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!
    const XLOOKUP = getLookupBuiltin('XLOOKUP')!
    const blank: CellValue = { tag: ValueTag.Empty }

    expect(LOOKUP(text('b'), cellRange([text('a'), text('b'), text('c')], 3, 1), cellRange([num(10), blank, num(30)], 3, 1))).toEqual(
      num(0),
    )
    expect(VLOOKUP(text('b'), cellRange([text('a'), num(10), text('b'), blank, text('c'), num(30)], 3, 2), num(2), bool(false))).toEqual(
      num(0),
    )
    expect(HLOOKUP(text('b'), cellRange([text('a'), text('b'), text('c'), num(10), blank, num(30)], 2, 3), num(2), bool(false))).toEqual(
      num(0),
    )
    expect(XLOOKUP(text('b'), cellRange([text('a'), text('b'), text('c')], 3, 1), cellRange([num(10), blank, num(30)], 3, 1))).toEqual(
      num(0),
    )
  })

  it('supports exact XLOOKUP and conditional aggregates', () => {
    const XLOOKUP = getLookupBuiltin('XLOOKUP')!
    const XMATCH = getLookupBuiltin('XMATCH')!
    const HLOOKUP = getLookupBuiltin('HLOOKUP')!
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const COUNTIFS = getLookupBuiltin('COUNTIFS')!
    const SUMIF = getLookupBuiltin('SUMIF')!
    const SUMIFS = getLookupBuiltin('SUMIFS')!
    const AVERAGEIF = getLookupBuiltin('AVERAGEIF')!
    const AVERAGEIFS = getLookupBuiltin('AVERAGEIFS')!
    const MINIFS = getLookupBuiltin('MINIFS')!
    const MAXIFS = getLookupBuiltin('MAXIFS')!
    const SUMPRODUCT = getLookupBuiltin('SUMPRODUCT')!

    expect(
      XLOOKUP(text('pear'), cellRange([text('apple'), text('pear'), text('plum')], 3, 1), cellRange([num(10), num(20), num(30)], 3, 1)),
    ).toEqual(num(20))
    expect(
      XLOOKUP(
        text('pear'),
        axisRange([text('apple'), text('pear'), text('plum')], 'cols', 3, 1),
        axisRange([num(10), num(20)], 'cols', 2, 1),
      ),
    ).toEqual(num(20))
    expect(
      XLOOKUP(
        text('plum'),
        axisRange([text('apple'), text('pear'), text('plum')], 'cols', 3, 1),
        axisRange([num(10), num(20)], 'cols', 2, 1),
      ),
    ).toEqual(num(0))

    expect(
      XLOOKUP(
        text('missing'),
        cellRange([text('apple'), text('pear'), text('plum')], 3, 1),
        cellRange([num(10), num(20), num(30)], 3, 1),
        text('fallback'),
      ),
    ).toEqual(text('fallback'))

    expect(COUNTIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text('>0'))).toEqual(num(3))
    expect(
      COUNTIF(cellRange([num(446_968), { tag: ValueTag.Empty }, num(25_399), num(200_000), num(50_000)], 5, 1), text('>=200,000')),
    ).toEqual(num(2))
    expect(
      COUNTIF(cellRange([num(446_968), { tag: ValueTag.Empty }, num(25_399), num(200_000), num(50_000)], 5, 1), text('<=50,000')),
    ).toEqual(num(2))

    expect(
      COUNTIFS(
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text('>0'),
        cellRange([text('a'), text('a'), text('b'), text('a')], 4, 1),
        text('a'),
      ),
    ).toEqual(num(3))

    expect(
      COUNTIFS(
        cellRange([text('North'), text('south'), text('Northeast'), text('west'), text('Northwest')], 5, 1),
        text('north*'),
        cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1),
        text('>1'),
      ),
    ).toEqual(num(2))
    expect(COUNTIF(cellRange([text('north*'), text('northwest'), text('south')], 3, 1), text('north~*'))).toEqual(num(1))

    expect(SUMIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text('>0'))).toEqual(num(12))

    expect(
      SUMIFS(
        cellRange([num(10), num(20), num(30), num(40)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text('>0'),
        cellRange([text('a'), text('a'), text('b'), text('a')], 4, 1),
        text('a'),
      ),
    ).toEqual(num(70))

    expect(
      SUMIFS(
        cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1),
        cellRange([text('North'), text('south'), text('Northeast'), text('west'), text('Northwest')], 5, 1),
        text('north*'),
        cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1),
        text('>1'),
      ),
    ).toEqual(num(8))
    expect(
      SUMIF(
        cellRange([empty(), text('North'), num(42), text('South')], 4, 1),
        text('*'),
        cellRange([num(10), num(20), num(30), num(40)], 4, 1),
      ),
    ).toEqual(num(60))
    expect(COUNTIF(cellRange([empty(), text('North'), num(42), text('South')], 4, 1), text('*'))).toEqual(num(2))

    expect(AVERAGEIF(cellRange([num(2), num(4), num(-1), num(6)], 4, 1), text('>0'))).toEqual(num(4))

    expect(
      AVERAGEIFS(
        cellRange([num(10), num(20), num(30), num(40)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text('>0'),
        cellRange([text('a'), text('a'), text('b'), text('a')], 4, 1),
        text('a'),
      ),
    ).toEqual(num((10 + 20 + 40) / 3))

    expect(
      MINIFS(
        cellRange([num(10), { tag: ValueTag.Empty }, num(30), num(5)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text('>0'),
        cellRange([text('a'), text('a'), text('b'), text('a')], 4, 1),
        text('a'),
      ),
    ).toEqual(num(5))

    expect(
      MAXIFS(
        cellRange([num(10), text('skip'), num(30), num(5)], 4, 1),
        cellRange([num(2), num(4), num(-1), num(6)], 4, 1),
        text('>0'),
        cellRange([text('a'), text('a'), text('b'), text('a')], 4, 1),
        text('a'),
      ),
    ).toEqual(num(10))

    expect(MINIFS(cellRange([{ tag: ValueTag.Empty }, text('skip')], 2, 1), cellRange([text('a'), text('a')], 2, 1), text('a'))).toEqual(
      num(0),
    )

    expect(MAXIFS(cellRange([num(1), num(2)], 2, 1), cellRange([num(1)], 1, 1), text('>0'))).toEqual(err(ErrorCode.Value))

    expect(SUMPRODUCT(cellRange([num(1), num(2), num(3)], 3, 1), cellRange([num(4), num(5), num(6)], 3, 1))).toEqual(num(32))

    expect(XMATCH(text('pear'), cellRange([text('apple'), text('pear'), text('plum')], 3, 1))).toEqual(num(2))

    expect(
      HLOOKUP(text('pear'), cellRange([text('apple'), text('pear'), text('plum'), num(10), num(20), num(30)], 2, 3), num(2), bool(false)),
    ).toEqual(num(20))
  })

  it('supports advanced XLOOKUP match and spill shapes', () => {
    const XLOOKUP = getLookupBuiltin('XLOOKUP')!

    expect(
      XLOOKUP(
        num(72),
        cellRange([num(50), num(60), num(70), num(80), num(90)], 5, 1),
        cellRange([text('D'), text('C'), text('B'), text('A'), text('S')], 5, 1),
        undefined,
        num(-1),
      ),
    ).toEqual(text('B'))
    expect(
      XLOOKUP(
        num(72),
        cellRange([num(50), num(60), num(70), num(80), num(90)], 5, 1),
        cellRange([text('D'), text('C'), text('B'), text('A'), text('S')], 5, 1),
        undefined,
        num(1),
      ),
    ).toEqual(text('A'))

    expect(
      XLOOKUP(
        text('ID2'),
        cellRange([text('ID1'), text('ID2'), text('ID3')], 3, 1),
        cellRange([text('Alex'), text('North'), num(10), text('James'), text('South'), num(20), text('Mina'), text('West'), num(30)], 3, 3),
      ),
    ).toEqual({ kind: 'array', rows: 1, cols: 3, values: [text('James'), text('South'), num(20)] })

    expect(
      XLOOKUP(
        cellRange([text('Q2'), text('Q4'), text('missing')], 3, 1),
        cellRange([text('Q1'), text('Q2'), text('Q3'), text('Q4')], 1, 4),
        cellRange([text('Keyboard'), text('Printer'), text('Monitor'), text('Dock')], 1, 4),
        text('fallback'),
      ),
    ).toEqual({ kind: 'array', rows: 3, cols: 1, values: [text('Printer'), text('Dock'), text('fallback')] })
    expect(
      XLOOKUP(
        cellRange([text('Q2'), err(ErrorCode.Ref)], 2, 1),
        cellRange([text('Q1'), text('Q2'), text('Q3'), text('Q4')], 1, 4),
        cellRange([text('Keyboard'), text('Printer'), text('Monitor'), text('Dock')], 1, 4),
      ),
    ).toEqual({ kind: 'array', rows: 2, cols: 1, values: [text('Printer'), err(ErrorCode.Ref)] })
  })

  it('supports XLOOKUP and XMATCH wildcard and binary search modes', () => {
    const XLOOKUP = getLookupBuiltin('XLOOKUP')!
    const XMATCH = getLookupBuiltin('XMATCH')!

    const names = cellRange([text('North-1'), text('South-1'), text('North-2'), text('North-3')], 4, 1)
    const amounts = cellRange([num(10), num(20), num(30), num(40)], 4, 1)

    expect(XLOOKUP(text('North-*'), names, amounts, text('missing'), num(2))).toEqual(num(10))
    expect(XLOOKUP(text('North-*'), names, amounts, text('missing'), num(2), num(-1))).toEqual(num(40))
    expect(XMATCH(text('North-?'), names, num(2))).toEqual(num(1))
    expect(XMATCH(text('North-?'), names, num(2), num(-1))).toEqual(num(4))

    const ascendingKeys = cellRange([num(10), num(20), num(30), num(40), num(50)], 5, 1)
    const ascendingLabels = cellRange([text('ten'), text('twenty'), text('thirty'), text('forty'), text('fifty')], 5, 1)
    expect(XLOOKUP(num(40), ascendingKeys, ascendingLabels, text('missing'), num(0), num(2))).toEqual(text('forty'))
    expect(XLOOKUP(num(35), ascendingKeys, ascendingLabels, text('missing'), num(-1), num(2))).toEqual(text('thirty'))
    expect(XLOOKUP(num(35), ascendingKeys, ascendingLabels, text('missing'), num(1), num(2))).toEqual(text('forty'))
    expect(XMATCH(num(35), ascendingKeys, num(-1), num(2))).toEqual(num(3))

    const descendingKeys = cellRange([num(50), num(40), num(30), num(20), num(10)], 5, 1)
    const descendingLabels = cellRange([text('fifty'), text('forty'), text('thirty'), text('twenty'), text('ten')], 5, 1)
    expect(XLOOKUP(num(40), descendingKeys, descendingLabels, text('missing'), num(0), num(-2))).toEqual(text('forty'))
    expect(XLOOKUP(num(35), descendingKeys, descendingLabels, text('missing'), num(-1), num(-2))).toEqual(text('thirty'))
    expect(XLOOKUP(num(35), descendingKeys, descendingLabels, text('missing'), num(1), num(-2))).toEqual(text('forty'))
    expect(XMATCH(num(35), descendingKeys, num(1), num(-2))).toEqual(num(2))

    expect(XLOOKUP(text('N*'), names, amounts, text('missing'), num(2), num(2))).toEqual(err(ErrorCode.Value))
  })

  it('covers conditional aggregate validation and error branches', () => {
    const COUNTIF = getLookupBuiltin('COUNTIF')!
    const COUNTIFS = getLookupBuiltin('COUNTIFS')!
    const SUMIF = getLookupBuiltin('SUMIF')!
    const SUMIFS = getLookupBuiltin('SUMIFS')!
    const AVERAGEIF = getLookupBuiltin('AVERAGEIF')!
    const AVERAGEIFS = getLookupBuiltin('AVERAGEIFS')!

    const values = cellRange([num(2), num(4), num(-1)], 3, 1)
    const otherValues = cellRange([num(10), text('skip'), num(30)], 3, 1)

    expect(COUNTIF(num(2), text('>0'))).toEqual(err(ErrorCode.Value))
    expect(COUNTIF(values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(COUNTIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(COUNTIF(err(ErrorCode.Ref), text('>0'))).toEqual(err(ErrorCode.Ref))

    expect(COUNTIFS()).toEqual(err(ErrorCode.Value))
    expect(COUNTIFS(values)).toEqual(err(ErrorCode.Value))
    expect(COUNTIFS(num(1), text('>0'))).toEqual(err(ErrorCode.Value))
    expect(COUNTIFS(values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(COUNTIFS(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(COUNTIFS(err(ErrorCode.Ref), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(COUNTIFS(values, text('>0'), cellRange([text('a'), text('b')], 2, 1), text('a'))).toEqual(err(ErrorCode.Value))

    expect(SUMIF(num(2), text('>0'))).toEqual(err(ErrorCode.Value))
    expect(SUMIF(values, text('>0'), num(2))).toEqual(err(ErrorCode.Value))
    expect(SUMIF(values, text('>0'), cellRange([num(10), num(20)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMIF(values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(SUMIF(err(ErrorCode.Ref), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(SUMIF(values, text('>0'), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(SUMIF(cellRange([text('match')], 1, 1), text('match'), cellRange([err(ErrorCode.Ref)], 1, 1))).toEqual(err(ErrorCode.Ref))

    expect(SUMIFS(num(10))).toEqual(err(ErrorCode.Value))
    expect(SUMIFS(values)).toEqual(err(ErrorCode.Value))
    expect(SUMIFS(num(10), values, text('>0'))).toEqual(err(ErrorCode.Value))
    expect(SUMIFS(values, values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(SUMIFS(values, values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(SUMIFS(err(ErrorCode.Ref), values, text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(SUMIFS(values, err(ErrorCode.Ref), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(SUMIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('match')], 1, 1), text('match'))).toEqual(err(ErrorCode.Ref))
    expect(SUMIFS(cellRange([err(ErrorCode.Ref)], 1, 1), cellRange([text('skip')], 1, 1), text('match'))).toEqual(num(0))
    expect(SUMIFS(values, cellRange([text('a'), text('b')], 2, 1), text('a'))).toEqual(err(ErrorCode.Value))

    expect(AVERAGEIF(num(2), text('>0'))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIF(values, text('>0'), num(2))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIF(values, text('>0'), cellRange([num(10), num(20)], 2, 1))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIF(values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIF(values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(AVERAGEIF(err(ErrorCode.Ref), text('>0'))).toEqual(err(ErrorCode.Ref))
    expect(AVERAGEIF(values, text('>0'), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(AVERAGEIF(cellRange([text('match')], 1, 1), text('match'), cellRange([err(ErrorCode.Ref)], 1, 1))).toEqual(err(ErrorCode.Ref))
    expect(AVERAGEIF(values, text('<-100'))).toEqual(err(ErrorCode.Div0))
    expect(AVERAGEIF(values, text('>0'), otherValues)).toEqual(num(10))

    expect(AVERAGEIFS(num(10))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIFS(values)).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIFS(num(10), values, text('>0'))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIFS(values, values, cellRange([text('>0')], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIFS(values, values, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(AVERAGEIFS(values, cellRange([text('a'), text('b')], 2, 1), text('a'))).toEqual(err(ErrorCode.Value))
    expect(AVERAGEIFS(values, values, text('<-100'))).toEqual(err(ErrorCode.Div0))
  })

  it('supports OFFSET, TAKE, and DROP shape transformations', () => {
    const OFFSET = getLookupBuiltin('OFFSET')!
    const TAKE = getLookupBuiltin('TAKE')!
    const DROP = getLookupBuiltin('DROP')!

    const matrix = cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 3, 2)

    expect(OFFSET(matrix, num(1), num(0), num(1), num(1))).toEqual(num(3))
    expect(OFFSET(matrix, num(-1), num(-1), num(1), num(1))).toEqual(num(6))
    expect(TAKE(matrix, num(2), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(1), num(3)],
    })
    expect(TAKE(matrix, num(-2), num(2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(3), num(4), num(5), num(6)],
    })
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 4, 1), num(2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(3), num(4)],
    })
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 1, 4), num(0), num(2))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(3), num(4)],
    })
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 1, 4), num(4))).toEqual(err(ErrorCode.Value))
    expect(DROP(cellRange([num(1), num(2), num(3), num(4)], 4, 1), num(4))).toEqual(err(ErrorCode.Value))
  })

  it('supports CHOOSECOLS and CHOOSEROWS extraction', () => {
    const CHOOSECOLS = getLookupBuiltin('CHOOSECOLS')!
    const CHOOSEROWS = getLookupBuiltin('CHOOSEROWS')!
    const matrix = cellRange([num(1), num(2), num(3), num(4), num(5), num(6)], 3, 2)

    expect(CHOOSECOLS(matrix, num(2))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [num(2), num(4), num(6)],
    })
    expect(CHOOSECOLS(matrix, num(2), num(1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [num(2), num(1), num(4), num(3), num(6), num(5)],
    })
    expect(CHOOSEROWS(matrix, num(3), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(5), num(6), num(1), num(2)],
    })
  })

  it('supports SORT and SORTBY ordering', () => {
    const SORT = getLookupBuiltin('SORT')!
    const SORTBY = getLookupBuiltin('SORTBY')!

    expect(SORT(cellRange([num(3), num(1), num(2)], 3, 1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [num(1), num(2), num(3)],
    })
    expect(SORT(cellRange([num(3), num(1), num(2), num(4)], 2, 2), num(1))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [num(2), num(4), num(3), num(1)],
    })

    expect(SORT(cellRange([num(9), num(2), num(8), num(1), num(5), num(7)], 2, 3), num(2), num(1), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [num(9), num(2), num(8), num(1), num(5), num(7)],
    })

    expect(SORTBY(cellRange([text('pear'), text('apple'), text('plum')], 3, 1), cellRange([num(2), num(1), num(3)], 3, 1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [text('apple'), text('pear'), text('plum')],
    })
    expect(SORTBY(cellRange([num(2), num(1), num(3)], 3, 1), cellRange([num(5), num(1), num(3)], 3, 1), num(-1))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [num(2), num(3), num(1)],
    })
  })

  it('supports TOCOL and TOROW flattening modes', () => {
    const TOCOL = getLookupBuiltin('TOCOL')!
    const TOROW = getLookupBuiltin('TOROW')!
    const matrix = cellRange([num(1), num(2), num(3), num(4)], 2, 2)

    expect(TOCOL(matrix)).toEqual({
      kind: 'array',
      rows: 4,
      cols: 1,
      values: [num(1), num(2), num(3), num(4)],
    })
    expect(TOROW(matrix)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 4,
      values: [num(1), num(2), num(3), num(4)],
    })
  })

  it('covers TOCOL and TOROW argument edge cases', () => {
    const TOCOL = getLookupBuiltin('TOCOL')!
    const TOROW = getLookupBuiltin('TOROW')!
    const rowRefRange: RangeBuiltinArgument = {
      kind: 'range',
      refKind: 'rows',
      rows: 1,
      cols: 1,
      values: [num(1)],
    }

    expect(TOCOL(num(1))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [num(1)],
    })
    expect(TOCOL(rowRefRange)).toEqual(err(ErrorCode.Value))
    expect(TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(TOCOL(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(TOCOL(cellRange([num(1), empty(), err(ErrorCode.Div0), num(4)], 2, 2), num(2))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 1,
      values: [num(1), empty(), num(4)],
    })
    expect(TOCOL(cellRange([num(1), empty(), err(ErrorCode.Div0), num(4)], 2, 2), num(3))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [num(1), num(4)],
    })
    expect(TOROW(num(1))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [num(1)],
    })
    expect(TOROW(rowRefRange)).toEqual(err(ErrorCode.Value))
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(TOROW(cellRange([num(1), empty(), err(ErrorCode.Div0), num(4)], 2, 2), num(2))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 3,
      values: [num(1), empty(), num(4)],
    })
    expect(TOROW(cellRange([num(1), empty(), err(ErrorCode.Div0), num(4)], 2, 2), num(3))).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [num(1), num(4)],
    })
    expect(TOROW(cellRange([num(1), num(2), num(3), num(4)], 2, 2), num(0), text('bad'))).toEqual(err(ErrorCode.Value))
  })

  it('supports WRAPROWS and WRAPCOLS packing', () => {
    const WRAPROWS = getLookupBuiltin('WRAPROWS')!
    const WRAPCOLS = getLookupBuiltin('WRAPCOLS')!
    const vector = cellRange([num(1), num(2), num(3), num(4), num(5)], 5, 1)
    const rowRefRange: RangeBuiltinArgument = {
      kind: 'range',
      refKind: 'rows',
      rows: 1,
      cols: 1,
      values: [num(1)],
    }

    expect(WRAPROWS(vector, num(2))).toEqual({
      kind: 'array',
      rows: 3,
      cols: 2,
      values: [num(1), num(2), num(3), num(4), num(5), err(ErrorCode.NA)],
    })
    expect(WRAPCOLS(vector, num(2))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [num(1), num(3), num(5), num(2), num(4), err(ErrorCode.NA)],
    })
    expect(WRAPCOLS(vector, num(2), text('pad'), bool(true))).toEqual({
      kind: 'array',
      rows: 2,
      cols: 3,
      values: [num(1), num(3), num(5), num(2), num(4), text('pad')],
    })

    expect(WRAPROWS(vector, num(0))).toEqual(err(ErrorCode.Value))
    expect(WRAPROWS(vector, num(2), cellRange([num(1)], 1, 1))).toEqual(err(ErrorCode.Value))
    expect(WRAPROWS(vector, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(WRAPROWS(vector, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(WRAPROWS(vector, num(2), text('pad'), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(WRAPROWS(rowRefRange, num(2))).toEqual(err(ErrorCode.Value))
    expect(WRAPROWS(vector, num(2), text('pad'), text('bad'))).toEqual(err(ErrorCode.Value))
    expect(WRAPCOLS(vector, err(ErrorCode.Name))).toEqual(err(ErrorCode.Name))
    expect(WRAPCOLS(vector, num(2), err(ErrorCode.NA))).toEqual(err(ErrorCode.NA))
    expect(WRAPCOLS(vector, num(2), text('pad'), err(ErrorCode.Ref))).toEqual(err(ErrorCode.Ref))
    expect(WRAPCOLS(vector, num(0))).toEqual(err(ErrorCode.Value))
    expect(WRAPCOLS(rowRefRange, num(2))).toEqual(err(ErrorCode.Value))
    expect(WRAPCOLS(vector, cellRange([num(1)], 1, 1), text('pad'))).toEqual(err(ErrorCode.Value))
    expect(WRAPCOLS(vector, num(2), text('pad'), text('bad'))).toEqual(err(ErrorCode.Value))
  })
})
