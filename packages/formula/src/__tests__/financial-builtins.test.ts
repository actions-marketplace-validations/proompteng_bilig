import { ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it } from 'vitest'
import { getBuiltin } from '../builtins.js'
import { clearExternalFunctionAdapters } from '../external-function-adapter.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: legacy financial and validation builtins', () => {
  it('supports ACCRINT, ACCRINTM, AMORDEGRC, and AMORLINC', () => {
    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstInterest = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 30 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    )
    const cost = { tag: ValueTag.Number, value: 2000 }
    const salvage = { tag: ValueTag.Number, value: 10 }
    const period = { tag: ValueTag.Number, value: 4 }
    const rate = { tag: ValueTag.Number, value: 0.1 }
    const basis = { tag: ValueTag.Number, value: 0 }

    expect(issue?.tag).toBe(ValueTag.Number)
    expect(firstInterest?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)

    const firstAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    )
    expect(firstAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(firstAccrual?.tag === ValueTag.Number ? firstAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    const omittedBasisAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
    )
    expect(omittedBasisAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(omittedBasisAccrual?.tag === ValueTag.Number ? omittedBasisAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    const fullAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
    )
    const shortAccrual = getBuiltin('ACCRINT')?.(
      issue,
      firstInterest,
      settlement,
      rate,
      { tag: ValueTag.Number, value: 1000 },
      { tag: ValueTag.Number, value: 2 },
      basis,
      { tag: ValueTag.Boolean, value: false },
    )
    expect(fullAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(shortAccrual).toMatchObject({ tag: ValueTag.Number })
    const shortAccrualValue = shortAccrual?.tag === ValueTag.Number ? shortAccrual.value : Number.NaN
    const fullAccrualValue = fullAccrual?.tag === ValueTag.Number ? fullAccrual.value : Number.NaN
    expect(shortAccrualValue).toBeLessThan(fullAccrualValue)

    const maturityAccrual = getBuiltin('ACCRINTM')?.(issue, settlement, rate, undefined, basis)
    expect(maturityAccrual).toMatchObject({ tag: ValueTag.Number })
    expect(maturityAccrual?.tag === ValueTag.Number ? maturityAccrual.value : Number.NaN).toBeCloseTo(91.66666666666667, 12)

    expect(getBuiltin('AMORLINC')?.(cost, issue, settlement, salvage, period, rate, basis)).toEqual({ tag: ValueTag.Number, value: 200 })

    expect(getBuiltin('AMORDEGRC')?.(cost, issue, settlement, salvage, period, rate, basis)).toEqual({ tag: ValueTag.Number, value: 163 })

    expect(
      getBuiltin('ACCRINT')?.(issue, settlement, issue, rate, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      getBuiltin('ACCRINT')?.(
        issue,
        settlement,
        settlement,
        rate,
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers ACCRINT and ACCRINTM basis variants and invalid argument branches', () => {
    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstInterest = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 1 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const rate = { tag: ValueTag.Number, value: 0.08 }
    const par = { tag: ValueTag.Number, value: 1000 }
    const frequency = { tag: ValueTag.Number, value: 2 }

    expect(issue?.tag).toBe(ValueTag.Number)
    expect(firstInterest?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)

    for (const basis of [0, 1, 2, 3, 4]) {
      expect(
        getBuiltin('ACCRINT')?.(issue, firstInterest, settlement, rate, par, frequency, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number })
      expect(
        getBuiltin('ACCRINTM')?.(issue, settlement, rate, par, {
          tag: ValueTag.Number,
          value: basis,
        }),
      ).toMatchObject({ tag: ValueTag.Number })
    }

    expect(
      getBuiltin('ACCRINT')?.(issue, firstInterest, settlement, rate, par, frequency, {
        tag: ValueTag.Number,
        value: 5,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(getBuiltin('ACCRINTM')?.(issue, settlement, rate, par, { tag: ValueTag.Number, value: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers AMORLINC and AMORDEGRC branch-heavy scenarios', () => {
    const cost = { tag: ValueTag.Number, value: 1000 }
    const datePurchased = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const firstPeriod = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2021 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const basis = { tag: ValueTag.Number, value: 0 }

    expect(datePurchased?.tag).toBe(ValueTag.Number)
    expect(firstPeriod?.tag).toBe(ValueTag.Number)

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 150 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 75 })

    expect(
      getBuiltin('AMORLINC')?.(
        cost,
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 25 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: 0.15 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.2 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 240 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.3 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 247 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.5 },
        basis,
      ),
    ).toMatchObject({ tag: ValueTag.Number, value: 250 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 1000 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1.2 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.Number, value: 100 },
        datePurchased,
        firstPeriod,
        { tag: ValueTag.Number, value: 200 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        basis,
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers combinatorics, product, quotient, and financial validation edge branches', () => {
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      getBuiltin('GCD')?.({ tag: ValueTag.Number, value: 54 }, { tag: ValueTag.Number, value: 24.9 }, { tag: ValueTag.Number, value: 6 }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(
      getBuiltin('LCM')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 3.8 }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('MROUND')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(
      getBuiltin('MULTINOMIAL')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(getBuiltin('PRODUCT')?.()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(getBuiltin('PRODUCT')?.({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('QUOTIENT')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    const issue = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 1 },
    )
    const settlement = getBuiltin('DATE')?.(
      { tag: ValueTag.Number, value: 2020 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 31 },
    )
    expect(issue?.tag).toBe(ValueTag.Number)
    expect(settlement?.tag).toBe(ValueTag.Number)
    expect(
      getBuiltin('AMORDEGRC')?.(
        { tag: ValueTag.String, value: 'bad', stringId: 2 },
        issue,
        settlement,
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('AMORLINC')?.(
        { tag: ValueTag.Number, value: 1000 },
        issue,
        settlement,
        { tag: ValueTag.String, value: 'bad', stringId: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers bitwise, integer, and rounding validation branches', () => {
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITXOR')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.String, value: 'bad', stringId: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITLSHIFT')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITLSHIFT')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 33 })).toEqual({
      tag: ValueTag.Number,
      value: 2 ** 33,
    })
    expect(getBuiltin('BITRSHIFT')?.({ tag: ValueTag.String, value: 'bad', stringId: 4 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('INT')?.({ tag: ValueTag.String, value: 'bad', stringId: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ROUNDUP')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 6 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('ROUNDDOWN')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 7 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: 12.34 }, { tag: ValueTag.String, value: 'bad', stringId: 8 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('TRUNC')?.({ tag: ValueTag.Number, value: -12.34 })).toEqual({
      tag: ValueTag.Number,
      value: -12,
    })
  })

  it('covers ceiling, floor, parity, factorial, and combinatoric branches', () => {
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(
      getBuiltin('FLOOR.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -6,
    })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(
      getBuiltin('CEILING.MATH')?.(
        { tag: ValueTag.Number, value: -5.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: -5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(
      getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.String, value: 'bad', stringId: 9 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('ISO.CEILING')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })

    expect(getBuiltin('BITAND')?.({ tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITAND')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.String, value: 'bad', stringId: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 11 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(getBuiltin('EVEN')?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Number,
      value: -4,
    })
    expect(getBuiltin('ODD')?.({ tag: ValueTag.Number, value: -2 })).toEqual({
      tag: ValueTag.Number,
      value: -3,
    })
    expect(getBuiltin('FACT')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('FACTDOUBLE')?.({ tag: ValueTag.Number, value: -3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COMBIN')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COMBINA')?.({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('GCD')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('LCM')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers logarithmic, hyperbolic, and sign-related math edge branches', () => {
    expect(getBuiltin('HARMEAN')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('HARMEAN')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })

    expect(getBuiltin('LOG10')?.({ tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('LOG10')?.({ tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('LOG')?.({ tag: ValueTag.Number, value: 100 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(
      getBuiltin('ACOT')?.({
        tag: ValueTag.Number,
        value: 0,
      }),
    ).toEqual({ tag: ValueTag.Number, value: Math.PI / 2 })
    expect(getBuiltin('ACOTH')?.({ tag: ValueTag.Number, value: 0.5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('COT')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('COTH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSC')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('CSCH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('SECH')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.Number, value: -42 })).toEqual({
      tag: ValueTag.Number,
      value: -1,
    })
    expect(getBuiltin('SIGN')?.({ tag: ValueTag.String, value: 'bad', stringId: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers direct trig, exponential, and positive rounding builtin paths', () => {
    expect(getBuiltin('SIN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('COS')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(getBuiltin('TAN')?.({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASIN')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 2, 12),
    })
    expect(getBuiltin('ACOS')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('ASIN')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('ACOS')?.({ tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('ATAN')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    })
    expect(getBuiltin('ATAN2')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI / 4, 12),
    })
    expect(getBuiltin('DEGREES')?.({ tag: ValueTag.Number, value: Math.PI })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(180, 12),
    })
    expect(getBuiltin('RADIANS')?.({ tag: ValueTag.Number, value: 180 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    })
    expect(getBuiltin('EXP')?.({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.E, 12),
    })
    expect(getBuiltin('LN')?.({ tag: ValueTag.Number, value: Math.E })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(getBuiltin('POWER')?.({ tag: ValueTag.Number, value: -32 }, { tag: ValueTag.Number, value: 1 / 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('SQRT')?.({ tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(getBuiltin('PI')?.()).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.PI, 12),
    })
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
    expect(getBuiltin('ACOSH')?.({ tag: ValueTag.Number, value: 0.5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(getBuiltin('ATANH')?.({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })

    expect(getBuiltin('FLOOR.MATH')?.({ tag: ValueTag.Number, value: 5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(getBuiltin('CEILING.MATH')?.({ tag: ValueTag.Number, value: 5.5 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(getBuiltin('FLOOR.PRECISE')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('CEILING.PRECISE')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(getBuiltin('BITAND')?.({ tag: ValueTag.String, value: 'bad', stringId: 13 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('BITOR')?.({ tag: ValueTag.String, value: 'bad', stringId: 14 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers AVERAGEA and SUBTOTAL aggregate dispatch branches', () => {
    expect(
      getBuiltin('AVERAGEA')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.String, value: 'skip', stringId: 15 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0.75 })

    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      getBuiltin('SUBTOTAL')?.(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 16 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Empty },
        { tag: ValueTag.String, value: 'skip', stringId: 17 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({ tag: ValueTag.Number, value: expect.closeTo(Math.sqrt(2), 12) })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      getBuiltin('SUBTOTAL')?.({ tag: ValueTag.Number, value: 11 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
  })

  it('covers aggregate aliases and formatting validation branches', () => {
    expect(getBuiltin('AVERAGEA')?.({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('AVERAGE')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(getBuiltin('AVG')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(getBuiltin('AVERAGEA')?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(
      getBuiltin('AGGREGATE')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Error, code: ErrorCode.NA },
        { tag: ValueTag.String, value: 'skip', stringId: 17 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(getBuiltin('AVERAGE')?.({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('AVG')?.({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(getBuiltin('MAXA')?.({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(getBuiltin('MINA')?.({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(
      getBuiltin('ADDRESS')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 5 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('ADDRESS')?.({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      getBuiltin('ADDRESS')?.(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Empty },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('FIXED')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1.5 })).toEqual({
      tag: ValueTag.String,
      value: '$10.0',
      stringId: 0,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Error, code: ErrorCode.NA }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('DOLLARFR')?.({ tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(getBuiltin('DOLLARDE')?.({ tag: ValueTag.Number, value: 1.6 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
  })

  it('supports ADDRESS and DOLLAR formatting edge cases', () => {
    const ADDRESS = getBuiltin('ADDRESS')!
    expect(ADDRESS({ tag: ValueTag.Number, value: 12 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.String,
      value: '$C$12',
      stringId: 0,
    })
    expect(ADDRESS({ tag: ValueTag.Number, value: 7 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.String,
      value: 'B$7',
      stringId: 0,
    })
    expect(
      ADDRESS(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.String,
      value: 'R4C5',
      stringId: 0,
    })

    expect(getBuiltin('DOLLAR')?.({ tag: ValueTag.Number, value: -1234.5 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.5',
      stringId: 0,
    })
    expect(
      getBuiltin('DOLLAR')?.(
        { tag: ValueTag.Number, value: 1234.56 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })
})
