import { ErrorCode, FormulaMode, SpreadsheetEngine, ValueTag, afterEach, describe, expect, it, vi } from './engine-test-helpers.js'

describe('SpreadsheetEngine wasm financial and statistical helpers', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('routes cash-flow rate helpers through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    ;[-70000, 12000, 15000, 18000, 21000, 26000].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[-120000, 39000, 30000, 21000, 37000, 46000].forEach((value, index) => {
      engine.setCellValue('Sheet1', `C${index + 1}`, value)
    })
    ;[-10000, 2750, 4250, 3250, 2750].forEach((value, index) => {
      engine.setCellValue('Sheet1', `E${index + 1}`, value)
    })
    ;[39448, 39508, 39751, 39859, 39904].forEach((value, index) => {
      engine.setCellValue('Sheet1', `F${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'H1', 'IRR(A1:A6)')
    engine.setCellFormula('Sheet1', 'I1', 'MIRR(C1:C6,10%,12%)')
    engine.setCellFormula('Sheet1', 'J1', 'XNPV(0.09,E1:E5,F1:F5)')
    engine.setCellFormula('Sheet1', 'K1', 'XIRR(E1:E5,F1:F5)')

    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.08663094803653162, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1260941303659051, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2086.647602031535, 9),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.37336253351883136, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes covariance and regression helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)

    engine.setCellFormula('Sheet1', 'C1', 'CORREL(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'D1', 'COVAR(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'E1', 'COVARIANCE.P(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'F1', 'COVARIANCE.S(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'G1', 'PEARSON(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'H1', 'INTERCEPT(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'I1', 'SLOPE(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'J1', 'RSQ(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'K1', 'STEYX(A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'L1', 'FORECAST(4,A1:A3,B1:B3)')
    engine.setCellFormula('Sheet1', 'M1', 'FORECAST.LINEAR(4,A1:A3,B1:B3)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 14 })

    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills TREND and GROWTH through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)
    engine.setCellValue('Sheet1', 'D1', 4)
    engine.setCellValue('Sheet1', 'D2', 5)
    engine.setCellFormula('Sheet1', 'F1', 'TREND(A1:A3,B1:B3,D1:D2)')

    engine.setCellValue('Sheet1', 'H1', 2)
    engine.setCellValue('Sheet1', 'H2', 4)
    engine.setCellValue('Sheet1', 'H3', 8)
    engine.setCellValue('Sheet1', 'I1', 1)
    engine.setCellValue('Sheet1', 'I2', 2)
    engine.setCellValue('Sheet1', 'I3', 3)
    engine.setCellValue('Sheet1', 'K1', 4)
    engine.setCellValue('Sheet1', 'K2', 5)
    engine.setCellFormula('Sheet1', 'M1', 'GROWTH(H1:H3,I1:I3,K1:K2)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 17 })
    expect(engine.getCellValue('Sheet1', 'M1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(16, 12),
    })
    expect(engine.getCellValue('Sheet1', 'M2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(32, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills LINEST and LOGEST through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 5)
    engine.setCellValue('Sheet1', 'A2', 8)
    engine.setCellValue('Sheet1', 'A3', 11)
    engine.setCellValue('Sheet1', 'B1', 1)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'B3', 3)
    engine.setCellFormula('Sheet1', 'D1', 'LINEST(A1:A3,B1:B3)')

    engine.setCellValue('Sheet1', 'G1', 2)
    engine.setCellValue('Sheet1', 'G2', 4)
    engine.setCellValue('Sheet1', 'G3', 8)
    engine.setCellValue('Sheet1', 'H1', 1)
    engine.setCellValue('Sheet1', 'H2', 2)
    engine.setCellValue('Sheet1', 'H3', 3)
    engine.setCellFormula('Sheet1', 'J1', 'LOGEST(G1:G3,H1:H3)')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 12),
    })

    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes promoted scalar and reducer math helpers through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 4)
    engine.setCellFormula('Sheet1', 'B1', 'ACOSH(1)')
    engine.setCellFormula('Sheet1', 'C1', 'COT(1)')
    engine.setCellFormula('Sheet1', 'D1', 'SECH(0)')
    engine.setCellFormula('Sheet1', 'E1', 'EVEN(-3)')
    engine.setCellFormula('Sheet1', 'F1', 'ODD(-2)')
    engine.setCellFormula('Sheet1', 'G1', 'FACT(5)')
    engine.setCellFormula('Sheet1', 'H1', 'FACTDOUBLE(6)')
    engine.setCellFormula('Sheet1', 'I1', 'COMBIN(8,3)')
    engine.setCellFormula('Sheet1', 'J1', 'COMBINA(3,2)')
    engine.setCellFormula('Sheet1', 'K1', 'GCD(A1:A3)')
    engine.setCellFormula('Sheet1', 'L1', 'LCM(A1:A3)')
    engine.setCellFormula('Sheet1', 'M1', 'PRODUCT(A1:A3)')
    engine.setCellFormula('Sheet1', 'N1', 'QUOTIENT(7,3)')
    engine.setCellFormula('Sheet1', 'O1', 'GEOMEAN(A1:A3)')
    engine.setCellFormula('Sheet1', 'P1', 'HARMEAN(A1:A3)')
    engine.setCellFormula('Sheet1', 'Q1', 'SUMSQ(A1:A3)')
    engine.setCellFormula('Sheet1', 'R1', 'TRUNC(-3.98,1)')
    engine.setCellFormula('Sheet1', 'S1', 'FLOOR.MATH(-5.5,2)')
    engine.setCellFormula('Sheet1', 'T1', 'FLOOR.PRECISE(-5.5,2)')
    engine.setCellFormula('Sheet1', 'U1', 'CEILING.MATH(-5.5,2)')
    engine.setCellFormula('Sheet1', 'V1', 'CEILING.PRECISE(-5.5,2)')
    engine.setCellFormula('Sheet1', 'W1', 'ISO.CEILING(-5.5,2)')
    engine.setCellFormula('Sheet1', 'X1', 'MROUND(10,4)')
    engine.setCellFormula('Sheet1', 'Y1', 'SQRTPI(2)')
    engine.setCellFormula('Sheet1', 'Z1', 'PERMUT(5,3)')
    engine.setCellFormula('Sheet1', 'AA1', 'PERMUTATIONA(2,3)')
    engine.setCellFormula('Sheet1', 'AB1', 'SERIESSUM(2,1,2,1,2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.6420926159343306, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: -3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 48 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 56 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'O1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.8844991406148166, 12),
    })
    expect(engine.getCellValue('Sheet1', 'P1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.769230769230769, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({ tag: ValueTag.Number, value: 29 })
    expect(engine.getCellValue('Sheet1', 'R1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.9, 12),
    })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(engine.getCellValue('Sheet1', 'T1')).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(engine.getCellValue('Sheet1', 'U1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'V1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'W1')).toEqual({ tag: ValueTag.Number, value: -4 })
    expect(engine.getCellValue('Sheet1', 'X1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'Y1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2.5066282746310002, 12),
    })
    expect(engine.getCellValue('Sheet1', 'Z1')).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(engine.getCellValue('Sheet1', 'AA1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'AB1')).toEqual({ tag: ValueTag.Number, value: 18 })

    for (const address of [
      'B1',
      'C1',
      'D1',
      'E1',
      'F1',
      'G1',
      'H1',
      'I1',
      'J1',
      'K1',
      'L1',
      'M1',
      'N1',
      'O1',
      'P1',
      'Q1',
      'R1',
      'S1',
      'T1',
      'U1',
      'V1',
      'W1',
      'X1',
      'Y1',
      'Z1',
      'AA1',
      'AB1',
    ]) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('routes coupon-date and bond pricing helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'COUPDAYBS(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'B1', 'COUPDAYS(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'C1', 'COUPDAYSNC(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'D1', 'COUPNCD(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'E1', 'COUPNUM(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'F1', 'COUPPCD(DATE(2007,1,25),DATE(2009,11,15),2,4)')
    engine.setCellFormula('Sheet1', 'G1', 'PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)')
    engine.setCellFormula('Sheet1', 'H1', 'YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)')
    engine.setCellFormula('Sheet1', 'I1', 'DURATION(DATE(2018,7,1),DATE(2048,1,1),0.08,0.09,2,1)')
    engine.setCellFormula('Sheet1', 'J1', 'MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 180 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 110 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 39217 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 39036 })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(94.63436162132213, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.065, 7),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(10.919145281591925, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5.735669813918838, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes annuity and cumulative loan helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'PV(0.1,2,-576.1904761904761)')
    engine.setCellFormula('Sheet1', 'B1', 'PMT(0.1,2,1000)')
    engine.setCellFormula('Sheet1', 'C1', 'NPER(0.1,-576.1904761904761,1000)')
    engine.setCellFormula('Sheet1', 'D1', 'RATE(48,-200,8000)')
    engine.setCellFormula('Sheet1', 'E1', 'IPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'F1', 'PPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'G1', 'ISPMT(0.1,1,2,1000)')
    engine.setCellFormula('Sheet1', 'H1', 'CUMIPMT(9%/12,30*12,125000,13,24,0)')
    engine.setCellFormula('Sheet1', 'I1', 'CUMPRINC(9%/12,30*12,125000,13,24,0)')
    engine.setCellFormula('Sheet1', 'J1', 'FV(0.1,2,-100,-1000)')
    engine.setCellFormula('Sheet1', 'K1', 'NPV(0.1,100,200,300)')

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1000.0000000000006, 12),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-576.1904761904758, 12),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1.9999999999999982, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.007701472488246008, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: -100 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-476.1904761904758, 12),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: -50 })
    expect(engine.getCellValue('Sheet1', 'H1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-11135.232130750845, 12),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-934.1071234208765, 12),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1420, 12),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(481.5927873779113, 12),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes rank helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 20)
    engine.setCellValue('Sheet1', 'A4', 30)

    engine.setCellFormula('Sheet1', 'B1', 'RANK(20,A1:A4)')
    engine.setCellFormula('Sheet1', 'C1', 'RANK.EQ(20,A1:A4)')
    engine.setCellFormula('Sheet1', 'D1', 'RANK.AVG(20,A1:A4)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2.5 })

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes order-statistics helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'B1', 'MEDIAN(A1:A8)')
    engine.setCellFormula('Sheet1', 'C1', 'SMALL(A1:A8,3)')
    engine.setCellFormula('Sheet1', 'D1', 'LARGE(A1:A8,2)')
    engine.setCellFormula('Sheet1', 'E1', 'PERCENTILE(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'F1', 'PERCENTILE.INC(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'G1', 'PERCENTILE.EXC(A1:A8,0.25)')
    engine.setCellFormula('Sheet1', 'H1', 'QUARTILE(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'I1', 'QUARTILE.INC(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'J1', 'QUARTILE.EXC(A1:A8,1)')
    engine.setCellFormula('Sheet1', 'K1', 'PERCENTRANK(A1:A8,8)')
    engine.setCellFormula('Sheet1', 'L1', 'PERCENTRANK.INC(A1:A8,8)')
    engine.setCellFormula('Sheet1', 'M1', 'PERCENTRANK.EXC(A1:A8,8)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7.5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 0.571 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 0.571 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 0.555 })

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills MODE.MULT and FREQUENCY through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 2, 3, 3, 4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[79, 85, 78, 85, 50, 81].forEach((value, index) => {
      engine.setCellValue('Sheet1', `C${index + 1}`, value)
    })
    ;[60, 80, 90].forEach((value, index) => {
      engine.setCellValue('Sheet1', `D${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'F1', 'MODE.MULT(A1:A6)')
    engine.setCellFormula('Sheet1', 'G1', 'FREQUENCY(C1:C6,D1:D3)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G4')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates database aggregation formulas through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[
      ['Age', 'Height', 'Yield'],
      [10, 100, 5],
      [12, 110, 7],
      [12, 120, 9],
      [15, 130, 11],
    ].forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`, value)
      })
    })
    engine.setCellValue('Sheet1', 'E1', 'Age')
    engine.setCellValue('Sheet1', 'E2', 12)
    engine.setCellValue('Sheet1', 'F1', 'Age')
    engine.setCellValue('Sheet1', 'F2', 15)

    engine.setCellFormula('Sheet1', 'H1', 'DAVERAGE(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H2', 'DCOUNT(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H3', 'DCOUNTA(A1:C5,"Height",E1:E2)')
    engine.setCellFormula('Sheet1', 'H4', 'DGET(A1:C5,"Height",F1:F2)')
    engine.setCellFormula('Sheet1', 'H5', 'DPRODUCT(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H6', 'DSTDEV(A1:C5,"Yield",E1:E2)')
    engine.setCellFormula('Sheet1', 'H7', 'DVARP(A1:C5,"Yield",E1:E2)')

    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'H4')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'H5')).toEqual({ tag: ValueTag.Number, value: 63 })
    expect(engine.getCellValue('Sheet1', 'H6')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.SQRT2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'H7')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H6').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes PROB and TRIMMEAN through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })
    ;[0.1, 0.2, 0.3, 0.4].forEach((value, index) => {
      engine.setCellValue('Sheet1', `B${index + 1}`, value)
    })
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      engine.setCellValue('Sheet1', `D${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'F1', 'PROB(A1:A4,B1:B4,2,3)')
    engine.setCellFormula('Sheet1', 'G1', 'TRIMMEAN(D1:D8,0.25)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0.5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(40 / 6, 12),
    })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes COUNTBLANK, ISOWEEKNUM, and TIMEVALUE through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 8)
    engine.setCellValue('Sheet1', 'A2', null)
    engine.setCellValue('Sheet1', 'B1', 'x')
    engine.setCellValue('Sheet1', 'B2', null)
    engine.setCellFormula('Sheet1', 'D1', 'COUNTBLANK(A1:B2)')
    engine.setCellFormula('Sheet1', 'E1', 'ISOWEEKNUM(DATE(2024,1,1))')
    engine.setCellFormula('Sheet1', 'F1', 'TIMEVALUE("1:30 PM")')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5625, 12),
    })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes chi-square inverse functions and compatibility aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'CHIDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'B1', 'LEGACY.CHIDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'C1', 'CHISQDIST(18.307,10)')
    engine.setCellFormula('Sheet1', 'D1', 'CHIINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'E1', 'CHISQ.INV.RT(0.050001,10)')
    engine.setCellFormula('Sheet1', 'F1', 'CHISQINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'G1', 'LEGACY.CHIINV(0.050001,10)')
    engine.setCellFormula('Sheet1', 'H1', 'CHISQ.INV(0.93,1)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(0.05000058909139826, 12)

    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(0.05000058909139826, 12)

    const c1 = engine.getCellValue('Sheet1', 'C1')
    expect(c1).toMatchObject({ tag: ValueTag.Number })
    expect(c1.value).toBeCloseTo(0.05000058909139826, 12)

    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(18.30697345696106, 12)

    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(18.30697345696106, 12)

    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(18.30697345696106, 12)

    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(18.30697345696106, 12)

    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(3.2830202867594993, 12)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes chi-square test functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' }, [
      [58, 35],
      [11, 25],
      [10, 23],
    ])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E3' }, [
      [45.35, 47.65],
      [17.56, 18.44],
      [16.09, 16.91],
    ])
    engine.setCellFormula('Sheet1', 'G1', 'CHISQ.TEST(A1:B3,D1:E3)')
    engine.setCellFormula('Sheet1', 'H1', 'CHITEST(A1:B3,D1:E3)')
    engine.setCellFormula('Sheet1', 'I1', 'LEGACY.CHITEST(A1:B3,D1:E3)')

    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.0003082, 7)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.0003082, 7)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.0003082, 7)

    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes beta and f distribution functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'BETA.DIST(2,8,10,TRUE,1,3)')
    engine.setCellFormula('Sheet1', 'B1', 'BETADIST(2,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'C1', 'BETA.INV(0.6854705810117458,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'D1', 'BETAINV(0.6854705810117458,8,10,1,3)')
    engine.setCellFormula('Sheet1', 'E1', 'F.DIST(15.2068649,6,4,TRUE)')
    engine.setCellFormula('Sheet1', 'F1', 'F.DIST.RT(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'G1', 'FDIST(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'H1', 'LEGACY.FDIST(15.2068649,6,4)')
    engine.setCellFormula('Sheet1', 'I1', 'F.INV(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'J1', 'F.INV.RT(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'K1', 'FINV(0.01,6,4)')
    engine.setCellFormula('Sheet1', 'L1', 'LEGACY.FINV(0.01,6,4)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(0.6854705810117458, 10)
    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(0.6854705810117458, 10)
    const c1 = engine.getCellValue('Sheet1', 'C1')
    expect(c1).toMatchObject({ tag: ValueTag.Number })
    expect(c1.value).toBeCloseTo(2, 10)
    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(2, 10)
    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(0.99, 9)
    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.01, 9)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.01, 9)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.01, 9)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.10930991466299911, 8)
    const j1 = engine.getCellValue('Sheet1', 'J1')
    expect(j1).toMatchObject({ tag: ValueTag.Number })
    expect(j1.value).toBeCloseTo(15.206864870947697, 7)
    const k1 = engine.getCellValue('Sheet1', 'K1')
    expect(k1).toMatchObject({ tag: ValueTag.Number })
    expect(k1.value).toBeCloseTo(15.206864870947697, 7)
    const l1 = engine.getCellValue('Sheet1', 'L1')
    expect(l1).toMatchObject({ tag: ValueTag.Number })
    expect(l1.value).toBeCloseTo(15.206864870947697, 7)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes student-t distribution functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A3' }, [[1], [2], [4]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B3' }, [[1], [3], [3]])
    engine.setCellFormula('Sheet1', 'D1', 'T.DIST(1,1,TRUE)')
    engine.setCellFormula('Sheet1', 'E1', 'T.DIST.RT(1,1)')
    engine.setCellFormula('Sheet1', 'F1', 'T.DIST.2T(1,1)')
    engine.setCellFormula('Sheet1', 'G1', 'TDIST(1,1,2)')
    engine.setCellFormula('Sheet1', 'H1', 'T.INV(0.75,1)')
    engine.setCellFormula('Sheet1', 'I1', 'T.INV.2T(0.5,1)')
    engine.setCellFormula('Sheet1', 'J1', 'TINV(0.5,1)')
    engine.setCellFormula('Sheet1', 'K1', 'CONFIDENCE.T(0.5,2,4)')
    engine.setCellFormula('Sheet1', 'L1', 'T.TEST(A1:A3,B1:B3,2,1)')
    engine.setCellFormula('Sheet1', 'M1', 'TTEST(A1:A3,B1:B3,2,1)')

    const d1 = engine.getCellValue('Sheet1', 'D1')
    expect(d1).toMatchObject({ tag: ValueTag.Number })
    expect(d1.value).toBeCloseTo(0.75, 12)
    const e1 = engine.getCellValue('Sheet1', 'E1')
    expect(e1).toMatchObject({ tag: ValueTag.Number })
    expect(e1.value).toBeCloseTo(0.25, 12)
    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.5, 12)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.5, 12)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(1, 9)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(1, 9)
    const j1 = engine.getCellValue('Sheet1', 'J1')
    expect(j1).toMatchObject({ tag: ValueTag.Number })
    expect(j1.value).toBeCloseTo(1, 9)
    expect(engine.getCellValue('Sheet1', 'K1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.764892328404345, 12),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 1 })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'J1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes statistical scalar and dispersion builtins through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    ;[1, 2, 3, 4, 5].forEach((value, index) => {
      engine.setCellValue('Sheet1', `A${index + 1}`, value)
    })

    engine.setCellFormula('Sheet1', 'C1', 'STANDARDIZE(1,0,1)')
    engine.setCellFormula('Sheet1', 'C2', 'STDEV(A1:A4)')
    engine.setCellFormula('Sheet1', 'C3', 'STDEVA(2,TRUE(),"skip")')
    engine.setCellFormula('Sheet1', 'C4', 'VAR(A1:A4)')
    engine.setCellFormula('Sheet1', 'C5', 'VARA(2,TRUE(),"skip")')
    engine.setCellFormula('Sheet1', 'C6', 'MODE(1,2,2,3)')
    engine.setCellFormula('Sheet1', 'C7', 'MODE.SNGL(1,2,2,3)')
    engine.setCellFormula('Sheet1', 'D1', 'SKEW(A1:A5)')
    engine.setCellFormula('Sheet1', 'D2', 'KURT(A1:A5)')
    engine.setCellFormula('Sheet1', 'D3', 'NORMDIST(1,0,1,TRUE)')
    engine.setCellFormula('Sheet1', 'D4', 'NORMINV(0.8413447460685429,0,1)')
    engine.setCellFormula('Sheet1', 'D5', 'NORMSDIST(1)')
    engine.setCellFormula('Sheet1', 'E1', 'CONFIDENCE.NORM(0.05,1,100)')
    engine.setCellFormula('Sheet1', 'E2', 'NORMSINV(0.001)')
    engine.setCellFormula('Sheet1', 'E3', 'LOGINV(0.5,0,1)')
    engine.setCellFormula('Sheet1', 'E4', 'LOGNORMDIST(1,0,1)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(5 / 3), 12),
    })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(5 / 3, 12),
    })
    expect(engine.getCellValue('Sheet1', 'C5')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C6')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C7')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'D2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-1.2, 12),
    })
    expect(engine.getCellValue('Sheet1', 'D3')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(engine.getCellValue('Sheet1', 'D4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(1, 8),
    })
    expect(engine.getCellValue('Sheet1', 'D5')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.8413447460685429, 7),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.1959963986120195, 12),
    })
    expect(engine.getCellValue('Sheet1', 'E2')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(-3.090232306167813, 8),
    })
    expect(engine.getCellValue('Sheet1', 'E3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E4')).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(0.5, 8),
    })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C5').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C6').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C7').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D4').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D5').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E4').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('keeps JS fallback ABS from swallowing arithmetic errors before wasm is ready', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'js-fallback-error-propagation' })
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'J12', 100)
    engine.setCellValue('Sheet1', 'M12', 78)
    engine.setCellFormula('Sheet1', 'Q12', 'SQRT(ABS((1/J12)+(1/K12)+(1/M12)+(1/N12)))')

    expect(engine.getCellValue('Sheet1', 'Q12')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
  })

  it('routes gamma inverse functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'GAMMA.INV(0.08030139707139418,3,2)')
    engine.setCellFormula('Sheet1', 'B1', 'GAMMAINV(0.08030139707139418,3,2)')

    const a1 = engine.getCellValue('Sheet1', 'A1')
    expect(a1).toMatchObject({ tag: ValueTag.Number })
    expect(a1.value).toBeCloseTo(2, 10)
    const b1 = engine.getCellValue('Sheet1', 'B1')
    expect(b1).toMatchObject({ tag: ValueTag.Number })
    expect(b1.value).toBeCloseTo(2, 10)

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes f-test and z-test functions and aliases through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' }, [[6], [7], [9], [15], [21]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B5' }, [[20], [28], [31], [38], [40]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'D5' }, [[1], [2], [3], [4], [5]])
    engine.setCellFormula('Sheet1', 'F1', 'F.TEST(A1:A5,B1:B5)')
    engine.setCellFormula('Sheet1', 'G1', 'FTEST(A1:A5,B1:B5)')
    engine.setCellFormula('Sheet1', 'H1', 'Z.TEST(D1:D5,2,1)')
    engine.setCellFormula('Sheet1', 'I1', 'ZTEST(D1:D5,2,1)')

    const f1 = engine.getCellValue('Sheet1', 'F1')
    expect(f1).toMatchObject({ tag: ValueTag.Number })
    expect(f1.value).toBeCloseTo(0.648317846786175, 12)
    const g1 = engine.getCellValue('Sheet1', 'G1')
    expect(g1).toMatchObject({ tag: ValueTag.Number })
    expect(g1.value).toBeCloseTo(0.648317846786175, 12)
    const h1 = engine.getCellValue('Sheet1', 'H1')
    expect(h1).toMatchObject({ tag: ValueTag.Number })
    expect(h1.value).toBeCloseTo(0.012673659338733989, 12)
    const i1 = engine.getCellValue('Sheet1', 'I1')
    expect(i1).toMatchObject({ tag: ValueTag.Number })
    expect(i1.value).toBeCloseTo(0.012673659338733989, 12)

    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes WORKDAY.INTL and NETWORKDAYS.INTL through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 46094)
    engine.setCellValue('Sheet1', 'A2', 46098)
    engine.setCellValue('Sheet1', 'B1', 46096)
    engine.setCellFormula('Sheet1', 'D1', 'WORKDAY.INTL(A1,1,7)')
    engine.setCellFormula('Sheet1', 'E1', 'WORKDAY.INTL(A1,2,7,B1)')
    engine.setCellFormula('Sheet1', 'F1', 'NETWORKDAYS.INTL(A1,A2,7)')
    engine.setCellFormula('Sheet1', 'G1', 'NETWORKDAYS.INTL(A1,A2,7,B1)')

    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 46096 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 46098 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2 })

    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes NUMBERVALUE and VALUETOTEXT through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'NUMBERVALUE("2.500,27",",",".")')
    engine.setCellFormula('Sheet1', 'B1', 'VALUETOTEXT("alpha",1)')
    engine.setCellFormula('Sheet1', 'C1', 'TEXT(1234.567,"#,##0.00")')
    engine.setCellFormula('Sheet1', 'D1', 'TEXT(DATE(2024,3,5),"yyyy-mm-dd")')

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({ tag: ValueTag.Number })
    expect(engine.getCellValue('Sheet1', 'A1').value).toBeCloseTo(2500.27, 12)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: '"alpha"',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: '1,234.57',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: '2024-03-05',
      stringId: expect.any(Number),
    })

    expect(engine.explainCell('Sheet1', 'A1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('routes scalar text conversion helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'CHAR(65)')
    engine.setCellFormula('Sheet1', 'B1', 'CODE("A")')
    engine.setCellFormula('Sheet1', 'C1', 'UNICODE("A")')
    engine.setCellFormula('Sheet1', 'D1', 'UNICHAR(66)')
    engine.setCellFormula('Sheet1', 'E1', 'CLEAN(CHAR(97)&CHAR(1)&CHAR(98))')
    engine.setCellFormula('Sheet1', 'F1', 'ASC("ＡＢＣ　１２３")')
    engine.setCellFormula('Sheet1', 'G1', 'JIS("ABC 123")')
    engine.setCellFormula('Sheet1', 'H1', 'DBCS("ｶﾞｷﾞｸﾞｹﾞｺﾞ")')
    engine.setCellFormula('Sheet1', 'I1', 'BAHTTEXT(1234)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'A',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 65 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 65 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: 'B',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.String,
      value: 'ab',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: 'ABC 123',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'ＡＢＣ　１２３',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({
      tag: ValueTag.String,
      value: 'ガギグゲゴ',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({
      tag: ValueTag.String,
      value: 'หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน',
      stringId: expect.any(Number),
    })

    for (const address of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes PHONETIC through the wasm path and reads the top-left range member', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'カタカナ')
    engine.setCellFormula('Sheet1', 'B1', '1/0')
    engine.setCellFormula('Sheet1', 'C1', 'PHONETIC(A1:B1)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: 'カタカナ',
      stringId: expect.any(Number),
    })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })
})
