import { ErrorCode, ValueTag } from '@bilig/protocol'
import { afterEach, describe, expect, it } from 'vitest'
import { getBuiltin } from '../builtins.js'
import { clearExternalFunctionAdapters } from '../external-function-adapter.js'

afterEach(() => {
  clearExternalFunctionAdapters()
})

describe('formula builtins: new financial builtins', () => {
  it('covers the new financial builtins', () => {
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

    expect(EFFECT({ tag: ValueTag.Number, value: 0.12 }, { tag: ValueTag.Number, value: 12 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12682503013196977, 12),
    })
    expect(NOMINAL({ tag: ValueTag.Number, value: 0.12682503013196977 }, { tag: ValueTag.Number, value: 12 })).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
    expect(
      PDURATION({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(
      RRI({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 121 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1, 12),
    })

    expect(
      FV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: -1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 12),
    })
    expect(
      FVSCHEDULE(
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0.09 },
        { tag: ValueTag.Number, value: 0.11 },
        { tag: ValueTag.Number, value: 0.1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1330.89, 12),
    })
    expect(
      DB(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 5 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(3690, 12),
    })
    expect(
      DDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(384, 12),
    })
    expect(
      VDB(
        { tag: ValueTag.Number, value: 2400 },
        { tag: ValueTag.Number, value: 300 },
        { tag: ValueTag.Number, value: 10 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(691.2, 12),
    })
    expect(
      PV(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -100 },
        { tag: ValueTag.Number, value: 1420 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1000, 12),
    })
    expect(PV({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 })).toEqual({
      tag: ValueTag.Number,
      value: 1000,
    })
    expect(
      PMT({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 1000 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904761, 12),
    })
    expect(PMT({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: -100,
    })
    expect(
      RATE({ tag: ValueTag.Number, value: 48 }, { tag: ValueTag.Number, value: -200 }, { tag: ValueTag.Number, value: 8000 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(RATE({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.Number, value: -100 }, { tag: ValueTag.Number, value: 1000 })).toEqual(
      {
        tag: ValueTag.Number,
        value: 0,
      },
    )
    expect(
      RATE(
        { tag: ValueTag.Number, value: 48 },
        { tag: ValueTag.Number, value: -200 },
        { tag: ValueTag.Number, value: 8000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(
      NPER(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: -576.1904761904761 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(NPER({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: -100 }, { tag: ValueTag.Number, value: 1000 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
    expect(
      NPV({ tag: ValueTag.Number, value: 0.1 }, { tag: ValueTag.Number, value: 100 }, { tag: ValueTag.Number, value: 100 }),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(173.55371900826447, 12),
    })
    expect(
      IPMT(
        { tag: ValueTag.String, value: 'bad', stringId: 17 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -100,
    })
    expect(
      IPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      PPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.19047619047615, 12),
    })
    expect(
      ISPMT(
        { tag: ValueTag.Number, value: 0.1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1000 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: -50,
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-11135.232130750845, 12),
    })
    expect(
      CUMIPMT(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 13 },
        { tag: ValueTag.Number, value: 24 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 12),
    })
    expect(
      CUMPRINC(
        { tag: ValueTag.Number, value: 0.09 / 12 },
        { tag: ValueTag.Number, value: 30 * 12 },
        { tag: ValueTag.Number, value: 125000 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-998.2910880208206, 12),
    })
    expect(SLN({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1000 }, { tag: ValueTag.Number, value: 9 })).toEqual({
      tag: ValueTag.Number,
      value: 1000,
    })
    expect(
      SYD(
        { tag: ValueTag.Number, value: 10000 },
        { tag: ValueTag.Number, value: 1000 },
        { tag: ValueTag.Number, value: 9 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1800,
    })
    expect(
      DISC(
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
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
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
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
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
        { tag: ValueTag.Number, value: 0.12 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1030.9278350515465, 12),
    })
    expect(
      PRICEDISC(
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
        { tag: ValueTag.Number, value: 0.0525 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.79583333333333, 12),
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
        { tag: ValueTag.Number, value: 99.795 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.05282257198685834, 12),
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
          { tag: ValueTag.Number, value: 2007 },
          { tag: ValueTag.Number, value: 11 },
          {
            tag: ValueTag.Number,
            value: 11,
          },
        ),
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0.061 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.98449887555694, 12),
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
        { tag: ValueTag.Number, value: 100.0123 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.060954333691538576, 12),
    })
    expect(
      ODDFPRICE(
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
        { tag: ValueTag.Number, value: 0.0785 },
        { tag: ValueTag.Number, value: 0.0625 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(113.597717474079, 12),
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
        { tag: ValueTag.Number, value: 0.0575 },
        { tag: ValueTag.Number, value: 84.5 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0772455415972989, 11),
    })
    expect(
      ODDLPRICE(
        DATE(
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 2 },
          {
            tag: ValueTag.Number,
            value: 7,
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
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(99.8782860147213, 12),
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
        { tag: ValueTag.Number, value: 0.0375 },
        { tag: ValueTag.Number, value: 99.875 },
        { tag: ValueTag.Number, value: 100 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.0451922356291692, 12),
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
          { tag: ValueTag.Number, value: 2008 },
          { tag: ValueTag.Number, value: 6 },
          {
            tag: ValueTag.Number,
            value: 1,
          },
        ),
        { tag: ValueTag.Number, value: 0.09 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(98.45, 12),
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
        { tag: ValueTag.Number, value: 98.45 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09141696292534264, 12),
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
        { tag: ValueTag.Number, value: 0.0914 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.09415149356594302, 12),
    })
    expect(
      DISC(
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
        { tag: ValueTag.Number, value: 97 },
        { tag: ValueTag.Number, value: 100 },
      ),
    ).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.12, 12),
    })
  })
})
