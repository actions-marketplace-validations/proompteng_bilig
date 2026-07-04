import { BuiltinId, ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it } from 'vitest'
import { getBuiltin, getBuiltinId } from '../builtins.js'
import { getLookupBuiltin } from '../builtins/lookup.js'
import { clearExternalFunctionAdapters } from '../external-function-adapter.js'
import type { ArrayValue } from '../runtime-values.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: statistical and distribution builtins', () => {
  it('covers the new type, statistical, distribution, and combinatoric builtins', () => {
    const T = getBuiltin('T')!
    const N = getBuiltin('N')!
    const TYPE = getBuiltin('TYPE')!
    const DELTA = getBuiltin('DELTA')!
    const GESTEP = getBuiltin('GESTEP')!
    const GAUSS = getBuiltin('GAUSS')!
    const PHI = getBuiltin('PHI')!
    const STANDARDIZE = getBuiltin('STANDARDIZE')!
    const MODE = getBuiltin('MODE')!
    const MODE_SNGL = getBuiltin('MODE.SNGL')!
    const STDEV = getBuiltin('STDEV')!
    const STDEV_S = getBuiltin('STDEV.S')!
    const STDEVP = getBuiltin('STDEVP')!
    const STDEV_P = getBuiltin('STDEV.P')!
    const STDEVA = getBuiltin('STDEVA')!
    const STDEVPA = getBuiltin('STDEVPA')!
    const VAR = getBuiltin('VAR')!
    const VAR_S = getBuiltin('VAR.S')!
    const VARP = getBuiltin('VARP')!
    const VAR_P = getBuiltin('VAR.P')!
    const VARA = getBuiltin('VARA')!
    const VARPA = getBuiltin('VARPA')!
    const SKEW = getBuiltin('SKEW')!
    const SKEW_P = getBuiltin('SKEW.P')!
    const KURT = getBuiltin('KURT')!
    const NORMDIST = getBuiltin('NORMDIST')!
    const NORM_DIST = getBuiltin('NORM.DIST')!
    const NORMINV = getBuiltin('NORMINV')!
    const NORM_INV = getBuiltin('NORM.INV')!
    const NORMSDIST = getBuiltin('NORMSDIST')!
    const NORM_S_DIST = getBuiltin('NORM.S.DIST')!
    const NORMSINV = getBuiltin('NORMSINV')!
    const NORM_S_INV = getBuiltin('NORM.S.INV')!
    const LOGINV = getBuiltin('LOGINV')!
    const LOGNORMDIST = getBuiltin('LOGNORMDIST')!
    const LOGNORM_DIST = getBuiltin('LOGNORM.DIST')!
    const LOGNORM_INV = getBuiltin('LOGNORM.INV')!
    const CONFIDENCE_NORM = getBuiltin('CONFIDENCE.NORM')!
    const PERMUT = getBuiltin('PERMUT')!
    const PERMUTATIONA = getBuiltin('PERMUTATIONA')!

    expect(T({ tag: ValueTag.String, value: 'alpha', stringId: 1 })).toEqual({
      tag: ValueTag.String,
      value: 'alpha',
      stringId: 1,
    })
    expect(T({ tag: ValueTag.Number, value: 42 })).toEqual({
      tag: ValueTag.String,
      value: '',
      stringId: 0,
    })
    expect(T({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(N({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(N({ tag: ValueTag.String, value: 'alpha', stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(N({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })

    expect(TYPE({ tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(TYPE({ tag: ValueTag.String, value: 'alpha', stringId: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(TYPE({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(TYPE({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Number,
      value: 16,
    })
    const arrayValue: ArrayValue = {
      kind: 'array',
      rows: 1,
      cols: 1,
      values: [{ tag: ValueTag.Number, value: 1 }],
    }
    expect(TYPE(arrayValue)).toEqual({
      tag: ValueTag.Number,
      value: 64,
    })

    expect(DELTA({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(DELTA({ tag: ValueTag.Number, value: 4 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(DELTA({ tag: ValueTag.String, value: 'bad', stringId: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(GESTEP({ tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(GESTEP({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(GESTEP({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.String, value: 'bad', stringId: 5 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(GAUSS({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 8),
    })
    expect(PHI({ tag: ValueTag.Number, value: 0 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    })
    expect(
      STANDARDIZE({ tag: ValueTag.Number, value: 42 }, { tag: ValueTag.Number, value: 40 }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      STANDARDIZE({ tag: ValueTag.Number, value: 42 }, { tag: ValueTag.Number, value: 40 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })

    expect(MODE({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(
      MODE(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(MODE_SNGL({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(MODE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(MODE_SNGL({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      STDEV(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    })
    expect(
      STDEV_S(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      STDEVP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(1.25), 12),
    })
    expect(STDEV_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      STDEVA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 7 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      STDEVPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 8 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(2 / 3), 12),
    })
    expect(STDEV({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(STDEV_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(STDEVP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(STDEV_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(STDEVA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(STDEVPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(STDEV({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(STDEV_S({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(STDEVP({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      VAR(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    })
    expect(VAR_S({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      VARP(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.25, 12),
    })
    expect(VAR_P({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      VARA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 9 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(
      VARPA(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.String, value: 'skip', stringId: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    })
    expect(VAR({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(VAR_S({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(VARP({ tag: ValueTag.Error, code: ErrorCode.Value })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(VAR_P({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(VARA({ tag: ValueTag.Error, code: ErrorCode.Num })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(VARPA({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(VAR({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(VAR_S({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(VARP({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      SKEW(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      SKEW_P(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 6 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      KURT(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    })
    expect(KURT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(SKEW({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(SKEW_P({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(KURT({ tag: ValueTag.Error, code: ErrorCode.Div0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })

    expect(
      NORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(
      NORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    })
    expect(
      NORMINV({ tag: ValueTag.Number, value: 0.8413447460685429 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(
      NORM_INV({ tag: ValueTag.Number, value: 0.8413447460685429 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(NORMSDIST({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(NORM_S_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: false })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.24197072451914337, 12),
    })
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.001 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    })
    expect(NORM_S_INV({ tag: ValueTag.Number, value: 0.999 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3.090232306167813, 8),
    })
    expect(NORMSINV({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(
      NORMINV({ tag: ValueTag.String, value: 'bad', stringId: 11 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NORM_S_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NORMSINV({ tag: ValueTag.String, value: 'bad', stringId: 13 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(LOGINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      LOGINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(
      LOGNORMDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    })
    expect(
      LOGNORM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1 / Math.sqrt(2 * Math.PI), 12),
    })
    expect(
      LOGNORM_INV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })
    expect(
      CONFIDENCE_NORM({ tag: ValueTag.Number, value: 0.05 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    })
    expect(
      CONFIDENCE_NORM({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(LOGNORMDIST({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 })).toEqual(
      {
        tag: ValueTag.Error,
        code: ErrorCode.Num,
      },
    )

    expect(PERMUT({ tag: ValueTag.Number, value: 5 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 60,
    })
    expect(PERMUTATIONA({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 8,
    })
    expect(PERMUTATIONA({ tag: ValueTag.String, value: 'bad', stringId: 14 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(PERMUT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })

    expect(getBuiltinId('norm.dist')).toBe(BuiltinId.NormDist)
    expect(getBuiltinId('norm.s.inv')).toBe(BuiltinId.NormSInv)
    expect(getBuiltinId('confidence.norm')).toBe(BuiltinId.ConfidenceNorm)
    expect(getBuiltinId('confidence.t')).toBe(BuiltinId.ConfidenceT)
    expect(getBuiltinId('gamma.inv')).toBe(BuiltinId.GammaInv)
    expect(getBuiltinId('gammainv')).toBe(BuiltinId.Gammainv)
    expect(getBuiltinId('permutationa')).toBe(BuiltinId.Permutationa)
    expect(getBuiltinId('chisq.test')).toBe(BuiltinId.ChisqTest)
    expect(getBuiltinId('chitest')).toBe(BuiltinId.Chitest)
    expect(getBuiltinId('legacy.chitest')).toBe(BuiltinId.LegacyChitest)
    expect(getBuiltinId('f.test')).toBe(BuiltinId.FTest)
    expect(getBuiltinId('ftest')).toBe(BuiltinId.Ftest)
    expect(getBuiltinId('z.test')).toBe(BuiltinId.ZTest)
    expect(getBuiltinId('ztest')).toBe(BuiltinId.Ztest)
    expect(getBuiltinId('workday.intl')).toBe(BuiltinId.WorkdayIntl)
    expect(getBuiltinId('networkdays.intl')).toBe(BuiltinId.NetworkdaysIntl)
    expect(getBuiltinId('numbervalue')).toBe(BuiltinId.Numbervalue)
    expect(getBuiltinId('valuetotext')).toBe(BuiltinId.Valuetotext)
    expect(getBuiltinId('asc')).toBe(BuiltinId.Asc)
    expect(getBuiltinId('jis')).toBe(BuiltinId.Jis)
    expect(getBuiltinId('dbcs')).toBe(BuiltinId.Dbcs)
    expect(getBuiltinId('daverage')).toBe(BuiltinId.Daverage)
    expect(getBuiltinId('dcount')).toBe(BuiltinId.Dcount)
    expect(getBuiltinId('dcounta')).toBe(BuiltinId.Dcounta)
    expect(getBuiltinId('dget')).toBe(BuiltinId.Dget)
    expect(getBuiltinId('dmax')).toBe(BuiltinId.Dmax)
    expect(getBuiltinId('dmin')).toBe(BuiltinId.Dmin)
    expect(getBuiltinId('dproduct')).toBe(BuiltinId.Dproduct)
    expect(getBuiltinId('dstdev')).toBe(BuiltinId.Dstdev)
    expect(getBuiltinId('dstdevp')).toBe(BuiltinId.Dstdevp)
    expect(getBuiltinId('dsum')).toBe(BuiltinId.Dsum)
    expect(getBuiltinId('dvar')).toBe(BuiltinId.Dvar)
    expect(getBuiltinId('dvarp')).toBe(BuiltinId.Dvarp)
    expect(getBuiltinId('prob')).toBe(BuiltinId.Prob)
    expect(getBuiltinId('trimmean')).toBe(BuiltinId.Trimmean)
    expect(getBuiltinId('growth')).toBe(BuiltinId.Growth)
    expect(getBuiltinId('trend')).toBe(BuiltinId.Trend)
    expect(getBuiltinId('forecast.linear')).toBe(BuiltinId.Forecast)
  })

  it('supports the new statistical distribution builtins and aliases', () => {
    const ERF = getBuiltin('ERF')!
    const ERF_PRECISE = getBuiltin('ERF.PRECISE')!
    const ERFC = getBuiltin('ERFC')!
    const ERFC_PRECISE = getBuiltin('ERFC.PRECISE')!
    const FISHER = getBuiltin('FISHER')!
    const FISHERINV = getBuiltin('FISHERINV')!
    const GAMMALN = getBuiltin('GAMMALN')!
    const GAMMALN_PRECISE = getBuiltin('GAMMALN.PRECISE')!
    const GAMMA = getBuiltin('GAMMA')!
    const CONFIDENCE = getBuiltin('CONFIDENCE')!
    const EXPONDIST = getBuiltin('EXPONDIST')!
    const EXPON_DIST = getBuiltin('EXPON.DIST')!
    const POISSON = getBuiltin('POISSON')!
    const POISSON_DIST = getBuiltin('POISSON.DIST')!
    const WEIBULL = getBuiltin('WEIBULL')!
    const WEIBULL_DIST = getBuiltin('WEIBULL.DIST')!
    const GAMMADIST = getBuiltin('GAMMADIST')!
    const GAMMA_DIST = getBuiltin('GAMMA.DIST')!
    const GAMMA_INV = getBuiltin('GAMMA.INV')!
    const GAMMAINV = getBuiltin('GAMMAINV')!
    const CHIDIST = getBuiltin('CHIDIST')!
    const CHISQ_DIST_RT = getBuiltin('CHISQ.DIST.RT')!
    const CHISQ_DIST = getBuiltin('CHISQ.DIST')!
    const BETA_DIST = getBuiltin('BETA.DIST')!
    const BETA_INV = getBuiltin('BETA.INV')!
    const BETADIST = getBuiltin('BETADIST')!
    const BETAINV = getBuiltin('BETAINV')!
    const F_DIST = getBuiltin('F.DIST')!
    const F_DIST_RT = getBuiltin('F.DIST.RT')!
    const F_INV = getBuiltin('F.INV')!
    const F_INV_RT = getBuiltin('F.INV.RT')!
    const FDIST = getBuiltin('FDIST')!
    const FINV = getBuiltin('FINV')!
    const LEGACY_FDIST = getBuiltin('LEGACY.FDIST')!
    const LEGACY_FINV = getBuiltin('LEGACY.FINV')!
    const T_DIST = getBuiltin('T.DIST')!
    const T_DIST_RT = getBuiltin('T.DIST.RT')!
    const T_DIST_2T = getBuiltin('T.DIST.2T')!
    const TDIST = getBuiltin('TDIST')!
    const T_INV = getBuiltin('T.INV')!
    const T_INV_2T = getBuiltin('T.INV.2T')!
    const TINV = getBuiltin('TINV')!
    const T_TEST = getLookupBuiltin('T.TEST')!
    const TTEST = getLookupBuiltin('TTEST')!
    const CONFIDENCE_T = getBuiltin('CONFIDENCE.T')!
    const BINOMDIST = getBuiltin('BINOMDIST')!
    const BINOM_DIST = getBuiltin('BINOM.DIST')!
    const BINOM_DIST_RANGE = getBuiltin('BINOM.DIST.RANGE')!
    const CRITBINOM = getBuiltin('CRITBINOM')!
    const BINOM_INV = getBuiltin('BINOM.INV')!
    const HYPGEOMDIST = getBuiltin('HYPGEOMDIST')!
    const HYPGEOM_DIST = getBuiltin('HYPGEOM.DIST')!
    const NEGBINOMDIST = getBuiltin('NEGBINOMDIST')!
    const NEGBINOM_DIST = getBuiltin('NEGBINOM.DIST')!

    expect(ERF({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERF({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERF_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8427006897475899, 7),
    })
    expect(ERFC({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    })
    expect(ERFC_PRECISE({ tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.15729931025241006, 7),
    })
    expect(ERF({ tag: ValueTag.String, value: 'bad', stringId: 15 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FISHER({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5493061443340549, 12),
    })
    expect(FISHERINV({ tag: ValueTag.Number, value: 0.5493061443340549 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(FISHERINV({ tag: ValueTag.String, value: 'bad', stringId: 16 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(GAMMALN({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    })
    expect(GAMMALN_PRECISE({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.log(24), 12),
    })
    expect(GAMMA({ tag: ValueTag.Number, value: 5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(24, 10),
    })
    expect(
      CONFIDENCE({ tag: ValueTag.Number, value: 0.05 }, { tag: ValueTag.Number, value: 1.5 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2939945976810081, 9),
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.764892328404345, 12),
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CONFIDENCE_T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(
      EXPONDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: false }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2706705664732254, 12),
    })
    expect(
      EXPON_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8646647167633873, 12),
    })
    expect(
      EXPONDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Boolean, value: false }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      EXPONDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: false }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      EXPON_DIST(
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      POISSON({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Boolean, value: false }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.21376301724973648, 12),
    })
    expect(
      POISSON_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.7575761331330662, 12),
    })
    expect(
      POISSON({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Boolean, value: false }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      POISSON_DIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 2.5 }, { tag: ValueTag.Boolean, value: true }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      POISSON(
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.2596002610238016, 12),
    })
    expect(
      WEIBULL_DIST(
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.22119921692859512, 12),
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      WEIBULL_DIST(
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      WEIBULL(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09196986029286061, 12),
    })
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08030139707139418, 12),
    })
    expect(
      GAMMA_INV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(GAMMA_INV({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      GAMMA_INV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      GAMMAINV(
        { tag: ValueTag.Number, value: 0.08030139707139418 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    })
    expect(CHIDIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    })
    expect(CHISQ_DIST_RT({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5578254003710748, 12),
    })
    expect(
      CHISQ_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.4421745996289252, 12),
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6854705810117458, 10),
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: false },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.4837646, 7),
    })
    expect(
      BETADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6854705810117458, 10),
    })
    expect(
      BETA_INV(
        { tag: ValueTag.Number, value: 0.6854705810117458 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      BETAINV(
        { tag: ValueTag.Number, value: 0.6854705810117458 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 10),
    })
    expect(
      F_DIST(
        { tag: ValueTag.Number, value: 15.2068649 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.99, 9),
    })
    expect(
      F_DIST(
        { tag: ValueTag.Number, value: 15.2068649 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0012238, 9),
    })
    expect(
      F_DIST_RT({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      FDIST({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      LEGACY_FDIST({ tag: ValueTag.Number, value: 15.2068649 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.01, 9),
    })
    expect(
      F_INV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.10930991466299911, 8),
    })
    expect(
      F_INV_RT({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      FINV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      LEGACY_FINV({ tag: ValueTag.Number, value: 0.01 }, { tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 4 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(15.206864870947697, 7),
    })
    expect(
      T_DIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: true }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.75, 12),
    })
    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.25, 12),
    })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      T_DIST(
        { tag: ValueTag.String, value: 'bad', stringId: 23 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.String, value: 'bad', stringId: 24 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(TDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toMatchObject(
      {
        tag: ValueTag.Number,
        value: expect.closeTo(0.25, 12),
      },
    )
    expect(T_INV({ tag: ValueTag.Number, value: 0.75 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(T_INV_2T({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(TINV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 9),
    })
    expect(
      T_TEST(
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 2 },
            { tag: ValueTag.Number, value: 4 },
          ],
        },
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 3 },
            { tag: ValueTag.Number, value: 3 },
          ],
        },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      TTEST(
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 2 },
            { tag: ValueTag.Number, value: 4 },
          ],
        },
        {
          kind: 'range',
          refKind: 'cells',
          rows: 3,
          cols: 1,
          values: [
            { tag: ValueTag.Number, value: 1 },
            { tag: ValueTag.Number, value: 3 },
            { tag: ValueTag.Number, value: 3 },
          ],
        },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(T_DIST_2T({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.375, 12),
    })
    expect(
      BINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6875, 12),
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.78125, 12),
    })
    expect(
      CRITBINOM({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.7 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      CRITBINOM({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.999999999999 }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      BINOM_INV({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 0.7 }),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2 / 3, 12),
    })
    expect(
      NEGBINOMDIST({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 0.5 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1875, 12),
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 12),
    })

    expect(FISHER({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: -1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      GAMMADIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      GAMMA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(GAMMA({ tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(CHIDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CHISQ_DIST({ tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Boolean, value: true }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      BETA_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 8 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Boolean, value: true },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(F_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      F_INV({ tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.String, value: 'bad', stringId: 71 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      F_INV_RT(
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.String, value: 'bad', stringId: 72 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Boolean, value: false },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(CRITBINOM({ tag: ValueTag.Number, value: 6 }, { tag: ValueTag.Number, value: 0.5 }, { tag: ValueTag.Number, value: 1 })).toEqual(
      {
        tag: ValueTag.Error,
        code: ErrorCode.Num,
      },
    )
    expect(
      HYPGEOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      HYPGEOM_DIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NEGBINOMDIST({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 1.5 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      NEGBINOM_DIST(
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(getBuiltinId('gamma.dist')).toBe(BuiltinId.GammaDist)
    expect(getBuiltinId('negbinom.dist')).toBe(BuiltinId.NegbinomDist)
    expect(getBuiltinId('binom.inv')).toBe(BuiltinId.BinomInv)
  })

  it('covers remaining Student t and binomial validation guards', () => {
    const T_DIST_RT = getBuiltin('T.DIST.RT')!
    const T_INV = getBuiltin('T.INV')!
    const T_INV_2T = getBuiltin('T.INV.2T')!
    const TDIST = getBuiltin('TDIST')!
    const BINOMDIST = getBuiltin('BINOMDIST')!
    const BINOM_DIST_RANGE = getBuiltin('BINOM.DIST.RANGE')!

    expect(T_DIST_RT({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(T_INV({ tag: ValueTag.String, value: 'bad', stringId: 401 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(T_INV_2T({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(TDIST({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(TDIST({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      BINOMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1.5 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      BINOM_DIST_RANGE(
        { tag: ValueTag.Number, value: 6 },
        { tag: ValueTag.Number, value: 0.5 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })
})
