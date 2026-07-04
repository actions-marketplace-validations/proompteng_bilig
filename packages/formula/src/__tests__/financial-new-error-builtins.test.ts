import { ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it } from 'vitest'
import { getBuiltin } from '../builtins.js'
import { clearExternalFunctionAdapters } from '../external-function-adapter.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: new financial builtin error branches', () => {
  it('covers the new financial builtin error branches', () => {
    const EFFECT = getBuiltin('EFFECT')!
    const NOMINAL = getBuiltin('NOMINAL')!
    const PDURATION = getBuiltin('PDURATION')!
    const RRI = getBuiltin('RRI')!
    const FV = getBuiltin('FV')!
    const PV = getBuiltin('PV')!
    const PMT = getBuiltin('PMT')!
    const NPER = getBuiltin('NPER')!
    const RATE = getBuiltin('RATE')!
    const NPV = getBuiltin('NPV')!
    const IPMT = getBuiltin('IPMT')!
    const PPMT = getBuiltin('PPMT')!
    const ISPMT = getBuiltin('ISPMT')!
    const CUMIPMT = getBuiltin('CUMIPMT')!
    const CUMPRINC = getBuiltin('CUMPRINC')!
    const DATE = getBuiltin('DATE')!
    const FVSCHEDULE = getBuiltin('FVSCHEDULE')!
    const DB = getBuiltin('DB')!
    const DDB = getBuiltin('DDB')!
    const VDB = getBuiltin('VDB')!
    const SLN = getBuiltin('SLN')!
    const SYD = getBuiltin('SYD')!
    const DISC = getBuiltin('DISC')!
    const INTRATE = getBuiltin('INTRATE')!
    const RECEIVED = getBuiltin('RECEIVED')!
    const PRICEDISC = getBuiltin('PRICEDISC')!
    const YIELDDISC = getBuiltin('YIELDDISC')!
    const PRICEMAT = getBuiltin('PRICEMAT')!
    const YIELDMAT = getBuiltin('YIELDMAT')!
    const ODDFPRICE = getBuiltin('ODDFPRICE')!
    const ODDFYIELD = getBuiltin('ODDFYIELD')!
    const ODDLPRICE = getBuiltin('ODDLPRICE')!
    const ODDLYIELD = getBuiltin('ODDLYIELD')!
    const TBILLPRICE = getBuiltin('TBILLPRICE')!
    const TBILLYIELD = getBuiltin('TBILLYIELD')!
    const TBILLEQ = getBuiltin('TBILLEQ')!

    expect(EFFECT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(EFFECT({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(EFFECT({ tag: ValueTag.String, value: 'bad', stringId: 106 }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.String, value: 'bad', stringId: 107 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PDURATION({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      PDURATION({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: -100 }, { tag: ValueTag.Number, value: 121 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      PDURATION(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.String, value: 'bad', stringId: 108 },
        { tag: ValueTag.Number, value: 121 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(RRI({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(RRI({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 121 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(RRI({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: -121 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      RRI({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.String, value: 'bad', stringId: 109 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(PMT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PV({ tag: ValueTag.String, value: 'bad', stringId: 39 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PMT({ tag: ValueTag.String, value: 'bad', stringId: 40 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1000 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(RATE({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      RATE(
        { tag: ValueTag.String, value: 'bad', stringId: 41 },
        { tag: ValueTag.Number, value: -200 },
        { tag: ValueTag.Number, value: 8000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NPER({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      NPER(
        { tag: ValueTag.String, value: 'bad', stringId: 42 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NPV({ tag: ValueTag.String, value: 'bad', stringId: 11 }, { tag: ValueTag.Number, value: 100 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      FV({ tag: ValueTag.String, value: 'bad', stringId: 17 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: -100 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FVSCHEDULE({ tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.String, value: 'bad', stringId: 18 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(FVSCHEDULE({ tag: ValueTag.String, value: 'bad', stringId: 43 }, { tag: ValueTag.Number, value: 0.09 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DB(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      DB(
        { tag: ValueTag.String, value: 'bad', stringId: 44 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      DDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      DDB(
        { tag: ValueTag.String, value: 'bad', stringId: 45 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      VDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      VDB(
        { tag: ValueTag.String, value: 'bad', stringId: 46 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.String, value: 'bad', stringId: 47 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.String, value: 'bad', stringId: 48 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(SLN({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(SLN({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: -1 })).toEqual(
      {
        tag: ValueTag.Number,
        value: -9000,
      },
    )
    expect(
      SYD(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 10 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      DISC(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      INTRATE(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1030 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      INTRATE(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 1030 },
        { tag: ValueTag.Number, value: 5 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      RECEIVED(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 4 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      PRICEDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        { tag: ValueTag.Number, value: 0.0525 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      YIELDDISC(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 16,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      PRICEMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 13,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      YIELDMAT(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 3,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 8,
          },
        ),
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDFPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.0785 },
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDFYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2021 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: -0.0575 },
        { tag: ValueTag.Number, value: 84.5 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 7,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 10 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        { tag: ValueTag.Number, value: 0.0375 },
        { tag: ValueTag.Number, value: 0.0405 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      ODDLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 4 },
          {
            tag: ValueTag.Number,
            value: 20,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 15,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 12 },
          {
            tag: ValueTag.Number,
            value: 24,
          },
        ),
        { tag: ValueTag.Number, value: -0.0375 },
        { tag: ValueTag.Number, value: 99.875 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      TBILLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2009 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.09 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      TBILLYIELD(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      TBILLEQ(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 3 },
          {
            tag: ValueTag.Number,
            value: 31,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      RECEIVED(
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        DATE(
          { tag: ValueTag.Number, value: 2023 },
          { tag: ValueTag.Number, value: 1 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0.12 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })
})
