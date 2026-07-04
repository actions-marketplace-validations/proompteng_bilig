import {
  ErrorCode,
  FormulaMode,
  SpreadsheetEngine,
  ValueTag,
  afterEach,
  describe,
  expect,
  isRuntimeFormulaWithDependencies,
  isRuntimeFormulaWithDirectAggregate,
  isRuntimeFormulaWithDirectCriteria,
  isRuntimeFormulaWithRanges,
  it,
  readRuntimeFormula,
  utcDateToExcelSerial,
  vi,
} from './engine-test-helpers.js'

describe('SpreadsheetEngine wasm text and lookup helpers', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('routes CHOOSE, TEXTBEFORE, TEXTAFTER, and TEXTJOIN through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'C1', 20)
    engine.setCellValue('Sheet1', 'B2', 30)
    engine.setCellValue('Sheet1', 'C2', 40)
    engine.setCellValue('Sheet1', 'D1', 100)
    engine.setCellValue('Sheet1', 'E1', 200)
    engine.setCellValue('Sheet1', 'D2', 300)
    engine.setCellValue('Sheet1', 'E2', 400)
    engine.setCellValue('Sheet1', 'G1', 'alpha')
    engine.setCellValue('Sheet1', 'G2', null)
    engine.setCellValue('Sheet1', 'G3', 'beta')
    engine.setCellFormula('Sheet1', 'A1', 'CHOOSE(2,"red","blue","green")')
    engine.setCellFormula('Sheet1', 'A2', 'TEXTBEFORE("alpha-beta","-")')
    engine.setCellFormula('Sheet1', 'A3', 'TEXTAFTER("alpha-beta","-")')
    engine.setCellFormula('Sheet1', 'A4', 'TEXTJOIN("-",TRUE,G1:G3)')
    engine.setCellFormula('Sheet1', 'H1', 'CHOOSE(1,B1:C2,D1:E2)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: 'blue',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({
      tag: ValueTag.String,
      value: 'alpha',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A3')).toEqual({
      tag: ValueTag.String,
      value: 'beta',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({
      tag: ValueTag.String,
      value: 'alpha-beta',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Sheet1', 'I2')).toEqual({ tag: ValueTag.Number, value: 40 })

    for (const address of ['A1', 'A2', 'A3', 'A4', 'H1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes byte-oriented text builtins through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'LENB("é")')
    engine.setCellFormula('Sheet1', 'B1', 'LEFTB("abcdef",2)')
    engine.setCellFormula('Sheet1', 'C1', 'MIDB("abcdef",3,2)')
    engine.setCellFormula('Sheet1', 'D1', 'RIGHTB("abcdef",3)')
    engine.setCellFormula('Sheet1', 'E1', 'FINDB("d","abcdef",3)')
    engine.setCellFormula('Sheet1', 'F1', 'SEARCHB("ph","alphabet")')
    engine.setCellFormula('Sheet1', 'G1', 'REPLACEB("alphabet",3,2,"Z")')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: 'ab',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({
      tag: ValueTag.String,
      value: 'cd',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({
      tag: ValueTag.String,
      value: 'def',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({
      tag: ValueTag.String,
      value: 'alZabet',
      stringId: expect.any(Number),
    })

    for (const address of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes ADDRESS and dollar-format helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'ADDRESS(12,3)')
    engine.setCellFormula('Sheet1', 'B1', 'DOLLAR(-1234.5,1)')
    engine.setCellFormula('Sheet1', 'C1', 'DOLLARDE(1.08,16)')
    engine.setCellFormula('Sheet1', 'D1', 'DOLLARFR(1.5,16)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.String,
      value: '$C$12',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.String,
      value: '-$1,234.5',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1.5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1.08 })

    for (const address of ['A1', 'B1', 'C1', 'D1'] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes bitwise and base-conversion helpers through the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'BITAND(6,3)')
    engine.setCellFormula('Sheet1', 'B1', 'BITOR(6,3)')
    engine.setCellFormula('Sheet1', 'C1', 'BITXOR(6,3)')
    engine.setCellFormula('Sheet1', 'D1', 'BITLSHIFT(1,4)')
    engine.setCellFormula('Sheet1', 'E1', 'BITRSHIFT(16,4)')
    engine.setCellFormula('Sheet1', 'F1', 'BASE(255,16,4)')
    engine.setCellFormula('Sheet1', 'G1', 'DECIMAL("00FF",16)')
    engine.setCellFormula('Sheet1', 'H1', 'BIN2DEC("1111111111")')
    engine.setCellFormula('Sheet1', 'I1', 'BIN2HEX("1111111111")')
    engine.setCellFormula('Sheet1', 'J1', 'BIN2OCT("1111111111")')
    engine.setCellFormula('Sheet1', 'K1', 'DEC2BIN(10,8)')
    engine.setCellFormula('Sheet1', 'L1', 'DEC2HEX(255,4)')
    engine.setCellFormula('Sheet1', 'M1', 'DEC2OCT(15,4)')
    engine.setCellFormula('Sheet1', 'N1', 'HEX2BIN("A",8)')
    engine.setCellFormula('Sheet1', 'O1', 'HEX2DEC("FFFFFFFFFF")')
    engine.setCellFormula('Sheet1', 'P1', 'HEX2OCT("F",4)')
    engine.setCellFormula('Sheet1', 'Q1', 'OCT2BIN("12",8)')
    engine.setCellFormula('Sheet1', 'R1', 'OCT2DEC("17")')
    engine.setCellFormula('Sheet1', 'S1', 'OCT2HEX("17",4)')
    engine.setCellFormula('Sheet1', 'T1', 'BESSELI(1.5,1)')
    engine.setCellFormula('Sheet1', 'U1', 'BESSELJ(1.9,2)')
    engine.setCellFormula('Sheet1', 'V1', 'BESSELK(1.5,1)')
    engine.setCellFormula('Sheet1', 'W1', 'BESSELY(2.5,1)')
    engine.setCellFormula('Sheet1', 'X1', 'CONVERT(6,"mi","km")')
    engine.setCellFormula('Sheet1', 'Y1', 'EUROCONVERT(1.2,"DEM","EUR")')
    engine.setCellFormula('Sheet1', 'Z1', 'EUROCONVERT(1,"FRF","DEM",TRUE,3)')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 16 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({
      tag: ValueTag.String,
      value: '00FF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 255 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({
      tag: ValueTag.String,
      value: 'FFFFFFFFFF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({
      tag: ValueTag.String,
      value: '7777777777',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({
      tag: ValueTag.String,
      value: '00FF',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({
      tag: ValueTag.String,
      value: '0017',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'O1')).toEqual({ tag: ValueTag.Number, value: -1 })
    expect(engine.getCellValue('Sheet1', 'P1')).toEqual({
      tag: ValueTag.String,
      value: '0017',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({
      tag: ValueTag.String,
      value: '00001010',
      stringId: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'R1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({
      tag: ValueTag.String,
      value: '000F',
      stringId: expect.any(Number),
    })
    const besseli = engine.getCellValue('Sheet1', 'T1')
    expect(besseli).toMatchObject({ tag: ValueTag.Number })
    if (besseli.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELI cell to be numeric')
    }
    expect(besseli.value).toBeCloseTo(0.981666428, 7)
    const besselj = engine.getCellValue('Sheet1', 'U1')
    expect(besselj).toMatchObject({ tag: ValueTag.Number })
    if (besselj.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELJ cell to be numeric')
    }
    expect(besselj.value).toBeCloseTo(0.329925728, 7)
    const besselk = engine.getCellValue('Sheet1', 'V1')
    expect(besselk).toMatchObject({ tag: ValueTag.Number })
    if (besselk.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELK cell to be numeric')
    }
    expect(besselk.value).toBeCloseTo(0.277387804, 7)
    const bessely = engine.getCellValue('Sheet1', 'W1')
    expect(bessely).toMatchObject({ tag: ValueTag.Number })
    if (bessely.tag !== ValueTag.Number) {
      throw new Error('Expected BESSELY cell to be numeric')
    }
    expect(bessely.value).toBeCloseTo(0.145918138, 7)
    expect(engine.getCellValue('Sheet1', 'X1')).toEqual({ tag: ValueTag.Number, value: 9.656064 })
    expect(engine.getCellValue('Sheet1', 'Y1')).toEqual({ tag: ValueTag.Number, value: 0.61 })
    const z1 = engine.getCellValue('Sheet1', 'Z1')
    expect(z1).toMatchObject({ tag: ValueTag.Number })
    if (z1.tag !== ValueTag.Number) {
      throw new Error('Expected EUROCONVERT triangulation cell to be numeric')
    }
    expect(z1.value).toBeCloseTo(0.29728616, 12)

    for (const address of [
      'A1',
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
    ] as const) {
      expect(engine.explainCell('Sheet1', address).mode).toBe(FormulaMode.WasmFastPath)
    }
  })

  it('routes USE.THE.COUNTIF through the wasm path as a COUNTIF alias', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', -2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'USE.THE.COUNTIF(A1:A3,">0")')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills FILTER with a computed comparison mask and UNIQUE through the wasm fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 2)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellFormula('Sheet1', 'B1', 'FILTER(A1:A4,A1:A4>2)')
    engine.setCellFormula('Sheet1', 'C1', 'UNIQUE(A1:A4)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C3')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('spills FILTER through the wasm fast path when the include mask is a range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellValue('Sheet1', 'A3', 2)
    engine.setCellValue('Sheet1', 'A4', 4)
    engine.setCellValue('Sheet1', 'B1', false)
    engine.setCellValue('Sheet1', 'B2', true)
    engine.setCellValue('Sheet1', 'B3', false)
    engine.setCellValue('Sheet1', 'B4', true)
    engine.setCellFormula('Sheet1', 'C1', 'FILTER(A1:A4,B1:B4)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates missing logical functions and lambda arrays', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'B1', 4)
    engine.setCellValue('Sheet1', 'B2', 5)
    engine.setCellValue('Sheet1', 'B3', 6)
    engine.setCellFormula('Sheet1', 'C1', 'IFS(A1>1,"big",TRUE(),"small")')
    engine.setCellFormula('Sheet1', 'D1', 'SWITCH(A1,1,"one","other")')
    engine.setCellFormula('Sheet1', 'E1', 'XOR(TRUE(),FALSE(),TRUE())')
    engine.setCellFormula('Sheet1', 'F1', 'LAMBDA(x,x+1)(4)')
    engine.setCellFormula('Sheet1', 'G1', 'MAP(A1:A3,LAMBDA(x,x*2))')
    engine.setCellFormula('Sheet1', 'H1', 'BYROW(A1:B3,LAMBDA(r,SUM(r)))')
    engine.setCellFormula('Sheet1', 'I1', 'BYCOL(A1:B3,LAMBDA(c,SUM(c)))')
    engine.setCellFormula('Sheet1', 'K1', 'REDUCE(0,A1:A3,LAMBDA(acc,x,acc+x))')
    engine.setCellFormula('Sheet1', 'L1', 'SCAN(0,A1:A3,LAMBDA(acc,x,acc+x))')
    engine.setCellFormula('Sheet1', 'M1', 'MAKEARRAY(2,2,LAMBDA(r,c,r+c))')
    engine.setCellFormula('Sheet1', 'O1', 'BYROW(A1:B3,LAMBDA(r,AVERAGE(r)))')
    engine.setCellFormula('Sheet1', 'P1', 'BYCOL(A1:B3,LAMBDA(c,COUNTA(c)))')
    engine.setCellFormula('Sheet1', 'R1', 'REDUCE(1,A1:A3,LAMBDA(acc,x,acc*x))')
    engine.setCellFormula('Sheet1', 'S1', 'SCAN(1,A1:A3,LAMBDA(acc,x,acc*x))')

    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.String,
      value: 'small',
    })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({
      tag: ValueTag.String,
      value: 'one',
    })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'G3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'L1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'L2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'L3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'M1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'N1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'M2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'N2')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'O1')).toEqual({ tag: ValueTag.Number, value: 2.5 })
    expect(engine.getCellValue('Sheet1', 'O2')).toEqual({ tag: ValueTag.Number, value: 3.5 })
    expect(engine.getCellValue('Sheet1', 'O3')).toEqual({ tag: ValueTag.Number, value: 4.5 })
    expect(engine.getCellValue('Sheet1', 'P1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'Q1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'R1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'S1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'S2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'S3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.explainCell('Sheet1', 'C1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'D1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'H1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'I1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'K1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'L1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'M1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'O1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'P1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'R1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'S1').mode).toBe(FormulaMode.WasmFastPath)
  })

  it('evaluates accelerated math builtins and JS matrix spills', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'B2', 4)
    engine.setCellValue('Sheet1', 'C1', 4)
    engine.setCellValue('Sheet1', 'C2', 5)
    engine.setCellValue('Sheet1', 'D1', 6)
    engine.setCellValue('Sheet1', 'D2', 7)
    engine.setCellValue('Sheet1', 'K1', Math.PI / 2)
    engine.setCellValue('Sheet1', 'K2', -3.98)
    engine.setCellFormula('Sheet1', 'E1', 'SIN(K1)')
    engine.setCellFormula('Sheet1', 'F1', 'TRUNC(K2,1)')
    engine.setCellFormula('Sheet1', 'G1', 'MUNIT(2)')
    engine.setCellFormula('Sheet1', 'I1', 'MMULT(A1:B2,C1:D2)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: -3.9 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'G2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.explainCell('Sheet1', 'G1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Number, value: 19 })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 27 })
    expect(engine.getCellValue('Sheet1', 'I2')).toEqual({ tag: ValueTag.Number, value: 28 })
    expect(engine.getCellValue('Sheet1', 'J2')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('supports cross-sheet references', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Sheet2')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellFormula('Sheet2', 'B2', 'Sheet1!A1*3')
    expect(engine.getCellValue('Sheet2', 'B2')).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('recalculates TODAY and NOW on the wasm path for each recalc-triggering batch', async () => {
    vi.useFakeTimers()
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    vi.setSystemTime(new Date('2026-03-19T15:45:30.000Z'))
    engine.setCellFormula('Sheet1', 'A1', 'TODAY()')
    engine.setCellFormula('Sheet1', 'B1', 'NOW()')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(utcDateToExcelSerial(new Date('2026-03-19T15:45:30.000Z'))),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date('2026-03-19T15:45:30.000Z')),
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })

    vi.setSystemTime(new Date('2026-03-20T01:02:03.000Z'))
    engine.setCellValue('Sheet1', 'C1', 1)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: Math.floor(utcDateToExcelSerial(new Date('2026-03-20T01:02:03.000Z'))),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date('2026-03-20T01:02:03.000Z')),
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 0 })
  })

  it('recalculates RAND on the wasm path for each recalc-triggering batch', async () => {
    const randomSpy = vi.spyOn(Math, 'random')
    randomSpy.mockReturnValueOnce(0.125).mockReturnValueOnce(0.875)

    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'RAND()')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.125 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'B1', 1)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 0.875 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('evaluates row and column aggregate ranges on the wasm path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A3', 5)
    engine.setCellValue('Sheet1', 'B3', 7)
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A:A)')
    engine.setCellFormula('Sheet1', 'C2', 'SUM(3:3)')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for supported aggregate formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)+ABS(A1/2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)
  })

  it('deduplicates overlapped precedents when scalar refs and ranges touch the same cell', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)+A1')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getDependencies('Sheet1', 'B1').directPrecedents).toEqual(['Sheet1!A1', 'Sheet1!A2'])
  })

  it('uses the wasm fast path for IF branch formulas once comparison parity exists', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    engine.setCellValue('Sheet1', 'A2', 9)
    engine.setCellFormula('Sheet1', 'B1', 'IF(A1>0,A1*2,A2-1)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for exact-parity logical builtins', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 0)
    engine.setCellFormula('Sheet1', 'B1', 'AND(A1,TRUE)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B2', 'OR(A2,FALSE)')
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B3', 'NOT(A2)')
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B11', 'AND(TRUE,A4)')
    expect(engine.getCellValue('Sheet1', 'B11')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B12', 'OR(A4,TRUE)')
    expect(engine.getCellValue('Sheet1', 'B12')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B13', 'NOT(2)')
    expect(engine.getCellValue('Sheet1', 'B13')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B14', 'IF(1/0,1,2)')
    expect(engine.getCellValue('Sheet1', 'B14')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B15', 'AND(FALSE(),NA())')
    expect(engine.getCellValue('Sheet1', 'B15')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B16', 'OR(TRUE(),1/0)')
    expect(engine.getCellValue('Sheet1', 'B16')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A3', 'hello')
    engine.setCellFormula('Sheet1', 'B4', 'AND(A3,TRUE)')
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(engine.getLastMetrics().wasmFormulaCount).toBe(1)

    engine.setCellFormula('Sheet1', 'B5', 'IFERROR(A1/0,"fallback")')
    expect(engine.getCellValue('Sheet1', 'B5')).toMatchObject({
      tag: ValueTag.String,
      value: 'fallback',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'B6', 'IFNA(NA(),"missing")')
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({
      tag: ValueTag.String,
      value: 'missing',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'C1', 2)
    engine.setCellValue('Sheet1', 'C2', 'ignored')
    engine.setCellValue('Sheet1', 'C3', true)
    engine.setCellFormula('Sheet1', 'D1', 'AND(C1:C4)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D2', 'OR(C2:C4)')
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D3', 'XOR(C1:C4)')
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D4', 'AND(C2:C4)')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D5', 'OR(C2:C2)')
    expect(engine.getCellValue('Sheet1', 'D5')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'C5', false)
    engine.setCellFormula('Sheet1', 'C6', 'NA()')

    engine.setCellFormula('Sheet1', 'D6', 'AND(C5:C6)')
    expect(engine.getCellValue('Sheet1', 'D6')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D7', 'OR(C3:C6)')
    expect(engine.getCellValue('Sheet1', 'D7')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })
  })

  it('uses the wasm fast path for INDEX VLOOKUP and HLOOKUP', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'pear')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'A2', 'apple')
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'D1', 'Q1')
    engine.setCellValue('Sheet1', 'E1', 'Q2')
    engine.setCellValue('Sheet1', 'F1', 'Q3')
    engine.setCellValue('Sheet1', 'D2', 100)
    engine.setCellValue('Sheet1', 'E2', 200)
    engine.setCellValue('Sheet1', 'F2', 300)
    engine.setCellFormula('Sheet1', 'H1', 'INDEX(A1:B2,2,2)')
    engine.setCellFormula('Sheet1', 'H2', 'VLOOKUP("apple",A1:B2,2,FALSE)')
    engine.setCellFormula('Sheet1', 'H3', 'HLOOKUP("Q3",D1:F2,2,FALSE)')

    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'H3')).toEqual({ tag: ValueTag.Number, value: 300 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })

  it('uses the direct criteria path for conditional aggregates and keeps SUMPRODUCT on wasm', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 4)
    engine.setCellValue('Sheet1', 'A3', -1)
    engine.setCellValue('Sheet1', 'A4', 6)
    engine.setCellValue('Sheet1', 'B1', 'x')
    engine.setCellValue('Sheet1', 'B2', 'x')
    engine.setCellValue('Sheet1', 'B3', 'y')
    engine.setCellValue('Sheet1', 'B4', 'x')
    engine.setCellValue('Sheet1', 'C1', 10)
    engine.setCellValue('Sheet1', 'C2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)
    engine.setCellValue('Sheet1', 'C4', 40)
    engine.setCellValue('Sheet1', 'D1', 1)
    engine.setCellValue('Sheet1', 'D2', 2)
    engine.setCellValue('Sheet1', 'D3', 3)
    engine.setCellValue('Sheet1', 'E1', 4)
    engine.setCellValue('Sheet1', 'E2', 5)
    engine.setCellValue('Sheet1', 'E3', 6)

    engine.setCellFormula('Sheet1', 'F1', 'COUNTIF(A1:A4,">0")')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F2', 'COUNTIFS(A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F2').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F3', 'SUMIF(A1:A4,">0",C1:C4)')
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F3').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F4', 'SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F4').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F5', 'AVERAGEIF(A1:A4,">0")')
    expect(engine.getCellValue('Sheet1', 'F5')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F5').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F6', 'AVERAGEIFS(C1:C4,A1:A4,">0",B1:B4,"x")')
    expect(engine.getCellValue('Sheet1', 'F6')).toEqual({
      tag: ValueTag.Number,
      value: (10 + 20 + 40) / 3,
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.explainCell('Sheet1', 'F6').mode).toBe(FormulaMode.WasmFastPath)

    engine.setCellFormula('Sheet1', 'F7', 'SUMPRODUCT(D1:D3,E1:E3)')
    expect(engine.getCellValue('Sheet1', 'F7')).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A3', 8)
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Sheet1', 'F5')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'F6')).toEqual({
      tag: ValueTag.Number,
      value: (10 + 20 + 40) / 3,
    })
    expect(engine.getLastMetrics().jsFormulaCount).toBe(0)
  })

  it('binds direct criteria descriptors for cell-driven criteria and handles min and max aggregates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'criteria-cell-driven' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'x')
    engine.setCellValue('Sheet1', 'A2', 'y')
    engine.setCellValue('Sheet1', 'A3', 'x')
    engine.setCellValue('Sheet1', 'A4', 'y')
    engine.setCellValue('Sheet1', 'C1', 10)
    engine.setCellValue('Sheet1', 'C2', 20)
    engine.setCellValue('Sheet1', 'C3', 30)
    engine.setCellValue('Sheet1', 'C4', 40)
    engine.setCellValue('Sheet1', 'D1', 'x')
    engine.setCellFormula('Sheet1', 'F1', 'COUNTIF(A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F2', 'MINIFS(C1:C4,A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F3', 'MAXIFS(C1:C4,A1:A4,D1)')
    engine.setCellFormula('Sheet1', 'F4', 'AVERAGEIFS(C1:C4,A1:A4,D1)')

    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F2').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F3').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'F4').mode).toBe(FormulaMode.WasmFastPath)

    const countIndex = engine.workbook.getCellIndex('Sheet1', 'F1')
    const minIndex = engine.workbook.getCellIndex('Sheet1', 'F2')
    const maxIndex = engine.workbook.getCellIndex('Sheet1', 'F3')
    const averageIndex = engine.workbook.getCellIndex('Sheet1', 'F4')
    if (countIndex === undefined || minIndex === undefined || maxIndex === undefined || averageIndex === undefined) {
      throw new Error('expected direct criteria formulas to exist')
    }

    const countFormula = readRuntimeFormula(engine, countIndex)
    if (!isRuntimeFormulaWithDirectCriteria(countFormula)) {
      throw new Error('expected COUNTIF runtime formula to expose direct criteria metadata')
    }
    expect(countFormula.directCriteria.aggregateKind).toBe('count')
    expect(countFormula.directCriteria.aggregateRange).toBeUndefined()
    expect(countFormula.directCriteria.criteriaPairs).toHaveLength(1)
    expect(countFormula.directCriteria.criteriaPairs[0]?.criterion).toMatchObject({
      kind: 'cell',
    })

    const minFormula = readRuntimeFormula(engine, minIndex)
    if (!isRuntimeFormulaWithDirectCriteria(minFormula)) {
      throw new Error('expected MINIFS runtime formula to expose direct criteria metadata')
    }
    expect(minFormula.directCriteria.aggregateKind).toBe('min')
    expect(minFormula.directCriteria.aggregateRange).toMatchObject({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 2,
      length: 4,
    })

    const maxFormula = readRuntimeFormula(engine, maxIndex)
    if (!isRuntimeFormulaWithDirectCriteria(maxFormula)) {
      throw new Error('expected MAXIFS runtime formula to expose direct criteria metadata')
    }
    expect(maxFormula.directCriteria.aggregateKind).toBe('max')

    const averageFormula = readRuntimeFormula(engine, averageIndex)
    if (!isRuntimeFormulaWithDirectCriteria(averageFormula)) {
      throw new Error('expected AVERAGEIFS runtime formula to expose direct criteria metadata')
    }
    expect(averageFormula.directCriteria.aggregateKind).toBe('average')

    engine.setCellValue('Sheet1', 'D1', 'y')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({ tag: ValueTag.Number, value: 30 })

    engine.setCellValue('Sheet1', 'D1', 'z')
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F3')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'F4')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
  })

  it('binds direct aggregate descriptors for bounded single-column SUM, AVERAGE, and COUNT', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'B2', 'AVERAGE(A1:A3)')
    engine.setCellFormula('Sheet1', 'B3', 'COUNT(A1:A3)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.explainCell('Sheet1', 'B2').mode).toBe(FormulaMode.JsOnly)
    expect(engine.explainCell('Sheet1', 'B3').mode).toBe(FormulaMode.WasmFastPath)

    const sumIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    const averageIndex = engine.workbook.getCellIndex('Sheet1', 'B2')
    const countIndex = engine.workbook.getCellIndex('Sheet1', 'B3')
    if (sumIndex === undefined || averageIndex === undefined || countIndex === undefined) {
      throw new Error('expected direct aggregate formulas to exist')
    }

    const sumFormula = readRuntimeFormula(engine, sumIndex)
    const averageFormula = readRuntimeFormula(engine, averageIndex)
    const countFormula = readRuntimeFormula(engine, countIndex)
    if (!isRuntimeFormulaWithDirectAggregate(sumFormula)) {
      throw new Error('expected SUM runtime formula to expose direct aggregate metadata')
    }
    if (!isRuntimeFormulaWithDirectAggregate(averageFormula)) {
      throw new Error('expected AVERAGE runtime formula to expose direct aggregate metadata')
    }
    if (!isRuntimeFormulaWithDirectAggregate(countFormula)) {
      throw new Error('expected COUNT runtime formula to expose direct aggregate metadata')
    }
    expect(isRuntimeFormulaWithRanges(sumFormula)).toBe(true)
    expect(sumFormula.rangeDependencies).toHaveLength(0)
    expect(isRuntimeFormulaWithDependencies(sumFormula)).toBe(true)
    expect(sumFormula.dependencyIndices).toEqual(new Uint32Array())

    expect(sumFormula.directAggregate).toMatchObject({
      aggregateKind: 'sum',
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
    })
    expect(averageFormula.directAggregate.aggregateKind).toBe('average')
    expect(countFormula.directAggregate.aggregateKind).toBe('count')

    engine.setCellValue('Sheet1', 'A1', 5)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 1 })
  })

  it('rebinds direct aggregate formulas when a formula appears inside a previously literal aggregate range', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-rebind-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'A2', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A2)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.setCellFormula('Sheet1', 'A2', 'A1*3')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })

    engine.setCellValue('Sheet1', 'A1', 4)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('keeps direct aggregate dependents current through coordinate mutation APIs', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-coordinate-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellValueAt(sheetId, 1, 0, 3)
    engine.setCellFormulaAt(sheetId, 0, 1, 'SUM(A1:A2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.setCellFormulaAt(sheetId, 1, 0, 'A1*3')
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })

    engine.setCellValueAt(sheetId, 0, 0, 4)
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 16 })
  })

  it('keeps direct aggregate formulas current through coordinate clears', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-clear-at-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellValueAt(sheetId, 1, 0, 3)
    engine.setCellFormulaAt(sheetId, 0, 1, 'SUM(A1:A2)')
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })

    engine.clearCellAt(sheetId, 1, 0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('keeps lookup formulas current through coordinate literal writes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-coordinate-write-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 10)
    engine.setCellValueAt(sheetId, 1, 0, 20)
    engine.setCellValueAt(sheetId, 2, 0, 30)
    engine.setCellValueAt(sheetId, 0, 3, 20)
    engine.setCellValueAt(sheetId, 1, 3, 25)
    engine.setCellFormulaAt(sheetId, 0, 4, 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormulaAt(sheetId, 0, 5, 'MATCH(D2,A1:A3,1)')

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValueAt(sheetId, 1, 0, 25)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })

    engine.setCellValueAt(sheetId, 1, 0, 20)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('keeps lookup formulas current through coordinate clears', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-coordinate-clear-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')?.id
    if (sheetId === undefined) {
      throw new Error('expected Sheet1 to exist')
    }

    engine.setCellValueAt(sheetId, 0, 0, 10)
    engine.setCellValueAt(sheetId, 1, 0, 20)
    engine.setCellValueAt(sheetId, 2, 0, 30)
    engine.setCellValueAt(sheetId, 0, 3, 20)
    engine.setCellValueAt(sheetId, 1, 3, 25)
    engine.setCellFormulaAt(sheetId, 0, 4, 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormulaAt(sheetId, 0, 5, 'MATCH(D2,A1:A3,1)')

    engine.clearCellAt(sheetId, 1, 0)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('uses the direct js path for exact MATCH and XMATCH while keeping XLOOKUP on wasm', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'apple')
    engine.setCellValue('Sheet1', 'A2', 'pear')
    engine.setCellValue('Sheet1', 'A3', 'pear')
    engine.setCellValue('Sheet1', 'A4', 'plum')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'B2', 20)
    engine.setCellValue('Sheet1', 'B3', 30)
    engine.setCellValue('Sheet1', 'B4', 40)
    engine.setCellValue('Sheet1', 'C1', 1)
    engine.setCellValue('Sheet1', 'C2', 3)
    engine.setCellValue('Sheet1', 'C3', 5)
    engine.setCellValue('Sheet1', 'O1', 'ID1')
    engine.setCellValue('Sheet1', 'O2', 'ID2')
    engine.setCellValue('Sheet1', 'O3', 'ID3')
    engine.setCellValue('Sheet1', 'P1', 'Alex')
    engine.setCellValue('Sheet1', 'Q1', 'North')
    engine.setCellValue('Sheet1', 'R1', 10)
    engine.setCellValue('Sheet1', 'P2', 'James')
    engine.setCellValue('Sheet1', 'Q2', 'South')
    engine.setCellValue('Sheet1', 'R2', 20)
    engine.setCellValue('Sheet1', 'P3', 'Mina')
    engine.setCellValue('Sheet1', 'Q3', 'West')
    engine.setCellValue('Sheet1', 'R3', 30)

    engine.setCellFormula('Sheet1', 'D1', 'MATCH("pear",A1:A4,0)')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellFormula('Sheet1', 'D3', 'XMATCH("pear",A1:A4,0,-1)')
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 0, jsFormulaCount: 1 })

    engine.setCellFormula('Sheet1', 'D4', 'XLOOKUP("pear",A1:A4,B1:B4)')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D5', 'XLOOKUP("missing",A1:A4,B1:B4,"fallback")')
    expect(engine.getCellValue('Sheet1', 'D5')).toMatchObject({
      tag: ValueTag.String,
      value: 'fallback',
    })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D6', 'XLOOKUP(4,C1:C3,B1:B3,"fallback",-1)')
    expect(engine.getCellValue('Sheet1', 'D6')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.explainCell('Sheet1', 'D6').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'D7', 'XLOOKUP(4,C1:C3,B1:B3,"fallback",1)')
    expect(engine.getCellValue('Sheet1', 'D7')).toEqual({ tag: ValueTag.Number, value: 30 })
    expect(engine.explainCell('Sheet1', 'D7').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellFormula('Sheet1', 'F1', 'XLOOKUP("ID2",O1:O3,P1:R3)')
    expect(engine.getCellValue('Sheet1', 'F1')).toMatchObject({ tag: ValueTag.String, value: 'James' })
    expect(engine.getCellValue('Sheet1', 'G1')).toMatchObject({ tag: ValueTag.String, value: 'South' })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.exportSnapshot().workbook.metadata?.spills).toEqual([{ sheetName: 'Sheet1', address: 'F1', rows: 1, cols: 3 }])
    expect(engine.explainCell('Sheet1', 'F1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 1, jsFormulaCount: 0 })

    engine.setCellValue('Sheet1', 'A1', 'pear')
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getLastMetrics()).toMatchObject({ wasmFormulaCount: 2, jsFormulaCount: 2 })
  })
})
