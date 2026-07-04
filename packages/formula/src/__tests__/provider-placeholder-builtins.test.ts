import { BuiltinId, ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { getBuiltin, getBuiltinId } from '../builtins.js'
import { getLookupBuiltin } from '../builtins/lookup.js'
import { placeholderBuiltinNames, protocolPlaceholderBuiltinNames } from '../builtins/placeholder.js'
import { clearExternalFunctionAdapters, installExternalFunctionAdapter } from '../external-function-adapter.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: provider and placeholder builtins', () => {
  it('supports legacy statistical aliases', () => {
    const legacyNormsDist = getBuiltin('LEGACY.NORMSDIST')?.({ tag: ValueTag.Number, value: 0 })
    expect(legacyNormsDist).toMatchObject({ tag: ValueTag.Number })
    if (legacyNormsDist?.tag !== ValueTag.Number) {
      throw new Error('LEGACY.NORMSDIST should return a number')
    }
    expect(legacyNormsDist.value).toBeCloseTo(0.5, 8)
    expect(getBuiltin('LEGACY.NORMSINV')?.({ tag: ValueTag.Number, value: 0.5 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0, 12),
    })
    expect(getBuiltin('LEGACY.CHIDIST')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.exp(-0.5), 12),
    })
    expect(
      getBuiltin('SKEWP')?.({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 3 }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
  })

  it('covers coupon-date and periodic bond helpers with accelerated semantics', () => {
    const DATE = getBuiltin('DATE')!
    const COUPDAYBS = getBuiltin('COUPDAYBS')!
    const COUPDAYS = getBuiltin('COUPDAYS')!
    const COUPDAYSNC = getBuiltin('COUPDAYSNC')!
    const COUPNCD = getBuiltin('COUPNCD')!
    const COUPNUM = getBuiltin('COUPNUM')!
    const COUPPCD = getBuiltin('COUPPCD')!
    const PRICE = getBuiltin('PRICE')!
    const YIELD = getBuiltin('YIELD')!
    const DURATION = getBuiltin('DURATION')!
    const MDURATION = getBuiltin('MDURATION')!

    const couponSettlement = DATE(
      { tag: ValueTag.Number, value: 2007 },
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 25 },
    )
    const couponMaturity = DATE(
      { tag: ValueTag.Number, value: 2009 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.Number, value: 15 },
    )

    expect(COUPDAYBS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 70,
    })
    expect(COUPDAYS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 180,
    })
    expect(COUPDAYSNC(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 110,
    })
    expect(COUPNCD(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 39217,
    })
    expect(COUPNUM(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
    expect(COUPPCD(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 39036,
    })

    expect(
      PRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2017 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 0.065 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(94.63436162132213, 12),
    })
    expect(
      YIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2016 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 95.04287 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.065, 7),
    })
    expect(
      DURATION(
        DATE(
          { tag: ValueTag.Number, value: 2018 },
          { tag: ValueTag.Number, value: 7 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2048 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(10.919145281591925, 12),
    })
    expect(
      MDURATION(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2016 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5.735669813918838, 12),
    })

    expect(COUPDAYBS(couponSettlement, couponMaturity, { tag: ValueTag.Number, value: 3 }, { tag: ValueTag.Number, value: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PRICE(
        couponSettlement,
        couponMaturity,
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: -0.01 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      YIELD(
        couponSettlement,
        couponMaturity,
        { tag: ValueTag.Number, value: 0.05 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      DURATION(
        couponMaturity,
        couponSettlement,
        { tag: ValueTag.Number, value: 0.08 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers remaining complex and distribution error branches', () => {
    const IMPRODUCT = getBuiltin('IMPRODUCT')!
    const IMSUB = getBuiltin('IMSUB')!
    const IMTAN = getBuiltin('IMTAN')!
    const IMSEC = getBuiltin('IMSEC')!
    const IMCSC = getBuiltin('IMCSC')!
    const IMCOT = getBuiltin('IMCOT')!
    const NORMDIST = getBuiltin('NORMDIST')!
    const LOGNORM_DOT_DIST = getBuiltin('LOGNORM.DIST')!
    const LOGNORMDIST = getBuiltin('LOGNORMDIST')!
    const NPV = getBuiltin('NPV')!

    expect(IMPRODUCT()).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMSUB({ tag: ValueTag.String, value: 'bad', stringId: 89 }, { tag: ValueTag.String, value: '1+i', stringId: 94 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMTAN({ tag: ValueTag.String, value: 'bad', stringId: 90 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMSEC({ tag: ValueTag.String, value: 'bad', stringId: 92 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMCSC({ tag: ValueTag.String, value: 'bad', stringId: 93 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(IMCOT({ tag: ValueTag.String, value: 'bad', stringId: 91 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NORMDIST(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      LOGNORMDIST(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      LOGNORM_DOT_DIST(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      NPV({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.String, value: 'bad', stringId: 95 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('registers protocol-declared placeholder builtins as blocked', () => {
    for (const name of placeholderBuiltinNames) {
      expect(getBuiltin(name)?.()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    }

    for (const name of protocolPlaceholderBuiltinNames) {
      expect(getBuiltinId(name.toLowerCase())).toBeDefined()
    }

    expect(getBuiltinId('sin')).toBe(BuiltinId.Sin)
    expect(getBuiltinId('weeknum')).toBe(BuiltinId.Weeknum)
    expect(getBuiltinId('rept')).toBe(BuiltinId.Rept)
    expect(getBuiltinId('filter')).toBe(BuiltinId.Filter)
    expect(getBuiltinId('let')).toBe(BuiltinId.Let)
    expect(getBuiltinId('multiple.operations')).toBe(BuiltinId.MultipleOperations)
    expect(getBuiltinId('textjoin')).toBe(BuiltinId.Textjoin)
  })

  it('routes provider-backed formulas through external adapters and blocks when none are installed', () => {
    const TRANSLATE = getBuiltin('TRANSLATE')!
    const COPILOT = getBuiltin('COPILOT')!
    const HYPERLINK = getBuiltin('HYPERLINK')!
    const DDE = getBuiltin('DDE')!
    const INFO = getBuiltin('INFO')!
    const PY = getBuiltin('PY')!
    const REGISTER_ID = getBuiltin('REGISTER.ID')!
    const FILTERXML = getLookupBuiltin('FILTERXML')!
    const GOOGLEFINANCE = getLookupBuiltin('GOOGLEFINANCE')!
    const IMPORTDATA = getLookupBuiltin('IMPORTDATA')!
    const IMPORTFEED = getLookupBuiltin('IMPORTFEED')!
    const IMPORTHTML = getLookupBuiltin('IMPORTHTML')!
    const IMPORTRANGE = getLookupBuiltin('IMPORTRANGE')!
    const IMPORTXML = getLookupBuiltin('IMPORTXML')!
    const STOCKHISTORY = getLookupBuiltin('STOCKHISTORY')!

    const hello = { tag: ValueTag.String, value: 'hello', stringId: 1 } as const
    const sourceLang = { tag: ValueTag.String, value: 'en', stringId: 2 } as const
    const targetLang = { tag: ValueTag.String, value: 'es', stringId: 3 } as const

    expect(placeholderBuiltinNames).not.toContain('TRANSLATE')
    expect(placeholderBuiltinNames).not.toContain('COPILOT')
    expect(placeholderBuiltinNames).not.toContain('FILTERXML')
    expect(placeholderBuiltinNames).not.toContain('GOOGLEFINANCE')
    expect(placeholderBuiltinNames).not.toContain('IMPORTDATA')
    expect(placeholderBuiltinNames).not.toContain('IMPORTFEED')
    expect(placeholderBuiltinNames).not.toContain('IMPORTHTML')
    expect(placeholderBuiltinNames).not.toContain('IMPORTRANGE')
    expect(placeholderBuiltinNames).not.toContain('IMPORTXML')
    expect(placeholderBuiltinNames).not.toContain('STOCKHISTORY')

    expect(TRANSLATE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(COPILOT(hello)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(HYPERLINK(hello, sourceLang)).toBe(sourceLang)
    expect(DDE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(INFO(hello)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    expect(PY(hello)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    expect(REGISTER_ID(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(FILTERXML(hello, sourceLang)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Blocked })
    expect(GOOGLEFINANCE(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(IMPORTDATA(hello)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(IMPORTFEED(hello)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(IMPORTHTML(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(IMPORTRANGE(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(IMPORTXML(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })
    expect(STOCKHISTORY(hello, sourceLang)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Blocked,
    })

    const translateImpl = vi.fn(() => ({
      tag: ValueTag.String,
      value: 'hola',
      stringId: 0,
    }))
    const copilotImpl = vi.fn(() => ({
      tag: ValueTag.String,
      value: 'summary',
      stringId: 0,
    }))
    const hyperlinkImpl = vi.fn(() => ({
      tag: ValueTag.String,
      value: 'linked',
      stringId: 0,
    }))
    const ddeImpl = vi.fn(() => ({ tag: ValueTag.Number, value: 42 }))
    const infoImpl = vi.fn(() => ({ tag: ValueTag.String, value: 'mac', stringId: 0 }))
    const pyImpl = vi.fn(() => ({ tag: ValueTag.String, value: 'dataframe', stringId: 0 }))
    const registerIdImpl = vi.fn(() => ({ tag: ValueTag.Number, value: 17 }))
    const filterXmlImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'one', stringId: 0 },
        { tag: ValueTag.String, value: 'two', stringId: 0 },
      ],
    }))
    const googleFinanceImpl = vi.fn(() => ({ tag: ValueTag.Number, value: 123.45 }))
    const importDataImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'symbol', stringId: 0 },
        { tag: ValueTag.String, value: 'GOOG', stringId: 0 },
      ],
    }))
    const importFeedImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'title', stringId: 0 },
        { tag: ValueTag.String, value: 'link', stringId: 0 },
      ],
    }))
    const importHtmlImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'rank', stringId: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    }))
    const importRangeImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'region', stringId: 0 },
        { tag: ValueTag.String, value: 'arr', stringId: 0 },
        { tag: ValueTag.String, value: 'west', stringId: 0 },
        { tag: ValueTag.Number, value: 96_000 },
      ],
    }))
    const importXmlImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'h1', stringId: 0 },
        { tag: ValueTag.String, value: 'h2', stringId: 0 },
      ],
    }))
    const stockHistoryImpl = vi.fn(() => ({
      kind: 'array' as const,
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 11 },
      ],
    }))

    installExternalFunctionAdapter({
      surface: 'web',
      resolveFunction(name) {
        if (name === 'TRANSLATE') {
          return { kind: 'scalar', implementation: translateImpl }
        }
        if (name === 'COPILOT') {
          return { kind: 'scalar', implementation: copilotImpl }
        }
        if (name === 'HYPERLINK') {
          return { kind: 'scalar', implementation: hyperlinkImpl }
        }
        if (name === 'INFO') {
          return { kind: 'scalar', implementation: infoImpl }
        }
        if (name === 'FILTERXML') {
          return { kind: 'lookup', implementation: filterXmlImpl }
        }
        if (name === 'GOOGLEFINANCE') {
          return { kind: 'lookup', implementation: googleFinanceImpl }
        }
        if (name === 'IMPORTDATA') {
          return { kind: 'lookup', implementation: importDataImpl }
        }
        if (name === 'IMPORTFEED') {
          return { kind: 'lookup', implementation: importFeedImpl }
        }
        if (name === 'IMPORTHTML') {
          return { kind: 'lookup', implementation: importHtmlImpl }
        }
        if (name === 'IMPORTRANGE') {
          return { kind: 'lookup', implementation: importRangeImpl }
        }
        if (name === 'IMPORTXML') {
          return { kind: 'lookup', implementation: importXmlImpl }
        }
        if (name === 'STOCKHISTORY') {
          return { kind: 'lookup', implementation: stockHistoryImpl }
        }
        return undefined
      },
    })
    installExternalFunctionAdapter({
      surface: 'external-data',
      resolveFunction(name) {
        if (name === 'DDE') {
          return { kind: 'scalar', implementation: ddeImpl }
        }
        if (name === 'REGISTER.ID') {
          return { kind: 'scalar', implementation: registerIdImpl }
        }
        return undefined
      },
    })
    installExternalFunctionAdapter({
      surface: 'python',
      resolveFunction(name) {
        if (name === 'PY') {
          return { kind: 'scalar', implementation: pyImpl }
        }
        return undefined
      },
    })

    expect(TRANSLATE(hello, sourceLang, targetLang)).toEqual({
      tag: ValueTag.String,
      value: 'hola',
      stringId: 0,
    })
    expect(COPILOT(hello)).toEqual({
      tag: ValueTag.String,
      value: 'summary',
      stringId: 0,
    })
    expect(HYPERLINK(hello, sourceLang)).toEqual({
      tag: ValueTag.String,
      value: 'linked',
      stringId: 0,
    })
    expect(DDE(hello, sourceLang, targetLang)).toEqual({ tag: ValueTag.Number, value: 42 })
    expect(INFO(hello)).toEqual({
      tag: ValueTag.String,
      value: 'mac',
      stringId: 0,
    })
    expect(PY(hello)).toEqual({
      tag: ValueTag.String,
      value: 'dataframe',
      stringId: 0,
    })
    expect(REGISTER_ID(hello, sourceLang)).toEqual({ tag: ValueTag.Number, value: 17 })
    expect(FILTERXML(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'one', stringId: 0 },
        { tag: ValueTag.String, value: 'two', stringId: 0 },
      ],
    })
    expect(GOOGLEFINANCE(hello, sourceLang)).toEqual({ tag: ValueTag.Number, value: 123.45 })
    expect(IMPORTDATA(hello)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'symbol', stringId: 0 },
        { tag: ValueTag.String, value: 'GOOG', stringId: 0 },
      ],
    })
    expect(IMPORTFEED(hello)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'title', stringId: 0 },
        { tag: ValueTag.String, value: 'link', stringId: 0 },
      ],
    })
    expect(IMPORTHTML(hello, sourceLang, targetLang)).toEqual({
      kind: 'array',
      rows: 1,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'rank', stringId: 0 },
        { tag: ValueTag.Number, value: 1 },
      ],
    })
    expect(IMPORTRANGE(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.String, value: 'region', stringId: 0 },
        { tag: ValueTag.String, value: 'arr', stringId: 0 },
        { tag: ValueTag.String, value: 'west', stringId: 0 },
        { tag: ValueTag.Number, value: 96_000 },
      ],
    })
    expect(IMPORTXML(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 1,
      values: [
        { tag: ValueTag.String, value: 'h1', stringId: 0 },
        { tag: ValueTag.String, value: 'h2', stringId: 0 },
      ],
    })
    expect(STOCKHISTORY(hello, sourceLang)).toEqual({
      kind: 'array',
      rows: 2,
      cols: 2,
      values: [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 11 },
      ],
    })

    expect(translateImpl).toHaveBeenCalledWith(hello, sourceLang, targetLang)
    expect(copilotImpl).toHaveBeenCalledWith(hello)
    expect(pyImpl).toHaveBeenCalledWith(hello)
    expect(filterXmlImpl).toHaveBeenCalledWith(hello, sourceLang)
    expect(googleFinanceImpl).toHaveBeenCalledWith(hello, sourceLang)
    expect(importDataImpl).toHaveBeenCalledWith(hello)
    expect(importFeedImpl).toHaveBeenCalledWith(hello)
    expect(importHtmlImpl).toHaveBeenCalledWith(hello, sourceLang, targetLang)
    expect(importRangeImpl).toHaveBeenCalledWith(hello, sourceLang)
    expect(importXmlImpl).toHaveBeenCalledWith(hello, sourceLang)
    expect(stockHistoryImpl).toHaveBeenCalledWith(hello, sourceLang)
  })
})
