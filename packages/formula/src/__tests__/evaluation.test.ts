import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { getBuiltin, getBuiltinId } from '../builtins.js'
import { evaluateAst, evaluateAstResult } from '../js-evaluator.js'
import { parseFormula } from '../parser.js'

describe('formula evaluator: scalar coercion and builtin dispatch', () => {
  it('returns value errors for missing required scalar builtin arguments instead of throwing', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }
    const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

    for (const formula of ['ABS()', 'SIN()', 'POWER(2)', 'ATAN2(1)', 'DOLLAR()', 'SQRTPI()']) {
      expect(evaluateAst(parseFormula(formula), context)).toEqual(valueError)
    }
  })

  it('returns value errors for documented scalar arity violations', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }
    const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

    for (const formula of [
      'ISBLANK()',
      'ISNUMBER()',
      'ISTEXT()',
      'ISERROR()',
      'ISERR()',
      'ISLOGICAL()',
      'ISNONTEXT()',
      'ISEVEN()',
      'ISODD()',
      'ISNA()',
      'ISREF()',
      'ERROR.TYPE()',
      'T()',
      'N()',
      'TYPE()',
      'PI(1)',
      'FLOOR(2.5)',
      'CEILING(2.5)',
      'DOLLAR(1234.5,1,TRUE())',
      'YEAR(1,2)',
      'MONTH(1,2)',
      'DAY(1,2)',
      'HOUR(1,2)',
      'MINUTE(1,2)',
      'SECOND(1,2)',
      'EDATE(1,1,1)',
      'EOMONTH(1,1,1)',
      'TIMEVALUE("12:00",1)',
      'WORKDAY(46094,1,46097,46098)',
      'NETWORKDAYS(46094,46101,46097,46101)',
      'WORKDAY.INTL(46094,1,1,46097,46098)',
      'NETWORKDAYS.INTL(46094,46101,1,46097,46101)',
      'VLOOKUP(2,A1:B3,2,FALSE(),1)',
      'MATCH(2,A1:A3,0,1)',
      'XLOOKUP(2,A1:A3,B1:B3,"missing",0,1,1)',
      'COUNTIF(A1:A3,">1",1)',
      'SUMIF(A1:A3,">1",B1:B3,1)',
      'AVERAGEIF(A1:A3,">1",B1:B3,1)',
      'IF(TRUE(),1,2,3)',
      'IFERROR(1,2,3)',
      'IFNA(1,2,3)',
      'FILTER(A1:A3,A1:A3>1,"none",1)',
      'SUBTOTAL(9)',
      'AGGREGATE(9,0)',
      'CORREL(A1:A3,B1:B3,1)',
      'NORM.DIST(1,0,1,TRUE(),1)',
      'NORM.S.DIST(1,TRUE(),1)',
      'BINOM.DIST(1,2,0.5,TRUE(),1)',
      'BETA.DIST(0.5,1,1,TRUE(),0,1,1)',
      'T.TEST(A1:A3,B1:B3,1,1,1)',
      'PMT(0.1,10,100,0,0,1)',
      'FV(0.1,10,100,0,0,1)',
      'PV(0.1,10,100,0,0,1)',
      'NPER(0.1,100,1000,0,0,1)',
      'IPMT(0.1,1,10,1000,0,0,1)',
      'PPMT(0.1,1,10,1000,0,0,1)',
      'SQRTPI(1,2)',
      'MOD(3,2,1)',
      'MROUND(2,1,1)',
      'TRUNC(1,0,1)',
      'RANDBETWEEN(1,10,1)',
      'RANDARRAY(1,1,0,1,TRUE(),1)',
      'MUNIT(2,1)',
      'ROMAN(10,0,1)',
      'ARABIC("X",1)',
      'COMPLEX(1)',
      'COMPLEX(1,2,"i",1)',
      'IMABS("1+i",1)',
      'IMDIV("1+i","1-i",1)',
      'IMPOWER("1+i",2,1)',
      'IMSQRT("1+i",1)',
      'SUMX2MY2(A1:A3,B1:B3,1)',
      'SUMX2PY2(A1:A3,B1:B3,1)',
      'SUMXMY2(A1:A3,B1:B3,1)',
      'MDETERM(A1:B2,1)',
      'MINVERSE(A1:B2,1)',
      'MMULT(A1:B2,B1:C2,1)',
      'ABS(1,2)',
      'INT(1,2)',
      'SIN(1,2)',
      'COS(1,2)',
      'TAN(1,2)',
      'ASIN(0.5,2)',
      'ACOS(0.5,2)',
      'ATAN(1,2)',
      'DEGREES(1,2)',
      'RADIANS(1,2)',
      'EXP(1,2)',
      'LN(1,2)',
      'LOG10(100,2)',
      'SQRT(4,2)',
      'SIGN(1,2)',
      'EVEN(1,2)',
      'ODD(1,2)',
      'FACT(3,2)',
      'FACTDOUBLE(5,2)',
      'SINH(1,2)',
      'COSH(1,2)',
      'TANH(1,2)',
      'ASINH(1,2)',
      'ACOSH(2,2)',
      'ATANH(0.5,2)',
      'ACOT(1,2)',
      'ACOTH(2,2)',
      'COT(1,2)',
      'COTH(1,2)',
      'CSC(1,2)',
      'CSCH(1,2)',
      'SEC(1,2)',
      'SECH(1,2)',
      'ROUND(1,2,3)',
      'ROUNDUP(1,2,3)',
      'ROUNDDOWN(1,2,3)',
      'LOG(100,10,2)',
      'POWER(2,3,4)',
      'DELTA(1,0,1)',
      'GESTEP(1,0,1)',
      'DOLLARDE(1,8,1)',
      'DOLLARFR(1,8,1)',
      'COMBIN(3,2,1)',
      'COMBINA(3,2,1)',
      'QUOTIENT(5,2,1)',
      'BASE(10,2,2,1)',
      'DECIMAL("10",2,1)',
      'BIN2DEC("10",1)',
      'BIN2HEX("10",1,1)',
      'ADDRESS(1,1,1,TRUE(),"s",1)',
      'CONVERT(1,"m","ft",1)',
      'LEFT("abc",1,2)',
      'RIGHT("abc",1,2)',
      'LEN("abc",1)',
      'LENB("abc",1)',
      'UPPER("a",1)',
      'LOWER("A",1)',
      'TRIM(" a ",1)',
      'VALUE("1",2)',
      'CHAR(65,2)',
      'CODE("A",2)',
      'UNICODE("A",2)',
      'UNICHAR(65,2)',
      'CLEAN("a",1)',
      'REPT("a",2,3)',
      'TEXT(1,"0",3)',
      'EXACT("a","a",1)',
      'MID("abc",1,1,2)',
      'FIND("a","abc",1,2)',
      'SEARCH("a","abc",1,2)',
      'REPLACE("abc",1,1,"x",2)',
      'SUBSTITUTE("abc","a","x",1,2)',
    ]) {
      expect(evaluateAst(parseFormula(formula), context)).toEqual(valueError)
    }
  })

  it('matches Excel logical empty and invalid text coercion edges', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('IF("",1,2)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('NOT("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('IFS("",1,TRUE(),2)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('AND("bad",TRUE())'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('OR("bad",FALSE())'), context)).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(evaluateAst(parseFormula('XOR("bad",TRUE())'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('AND("bad")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('OR("bad")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('XOR("bad")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('covers scalar builtins and builtin id lookup', () => {
    const SUM = getBuiltin('sum')!
    const AVG = getBuiltin('AVG')!
    const MOD = getBuiltin('MOD')!
    const LEN = getBuiltin('LEN')!
    const CONCAT = getBuiltin('CONCAT')!
    const IF = getBuiltin('IF')!
    const AND = getBuiltin('AND')!
    const OR = getBuiltin('OR')!
    const NOT = getBuiltin('NOT')!

    expect(getBuiltinId('sum')).toBeDefined()
    expect(getBuiltinId('')).toBeUndefined()
    expect(getBuiltin('missing')).toBeUndefined()

    expect(SUM({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(
      AVG({ tag: ValueTag.Number, value: 2 }, { tag: ValueTag.String, value: 'ignored', stringId: 0 }, { tag: ValueTag.Number, value: 4 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(MOD({ tag: ValueTag.Number, value: 8 }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(LEN({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(
      CONCAT({ tag: ValueTag.String, value: 'hello', stringId: 0 }, { tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 7 }),
    ).toEqual({ tag: ValueTag.String, value: 'hello7', stringId: 0 })
    expect(IF({ tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(AND({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(OR({ tag: ValueTag.Empty }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(NOT({ tag: ValueTag.Empty })).toEqual({ tag: ValueTag.Boolean, value: true })
  })

  it('coerces numeric text produced by expressions without summing referenced text cells', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return { tag: ValueTag.String, value: '5-7', stringId: 1 }
        }
        if (address === 'A2') {
          return { tag: ValueTag.String, value: '5', stringId: 2 }
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('SUM(LEFT(A1,1),RIGHT(A1,LEN(A1)-2)/12)'), context)).toEqual({
      tag: ValueTag.Number,
      value: 5 + 7 / 12,
    })
    expect(evaluateAst(parseFormula('SUM(A2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('coerces xlsx-fixture-corpus date and time text for arithmetic operators', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.String, value: '22/07/2008', stringId: 1 },
      A2: { tag: ValueTag.String, value: '31/12/2019', stringId: 2 },
      B1: { tag: ValueTag.Number, value: 41_153 },
      B2: { tag: ValueTag.String, value: '30/09/2021', stringId: 3 },
      C1: { tag: ValueTag.String, value: '00:00:00', stringId: 4 },
      C2: { tag: ValueTag.String, value: '00:00:00', stringId: 5 },
    }
    const context = {
      sheetName: 'Sheet1',
      dateSystem: '1900' as const,
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('(A2-A1)/365'), context)).toEqual({
      tag: ValueTag.Number,
      value: 11.449315068493151,
    })
    expect(evaluateAst(parseFormula('(B2-B1)/365'), context)).toEqual({
      tag: ValueTag.Number,
      value: 9.084931506849315,
    })
    expect(evaluateAst(parseFormula('(C2-C1)/365'), context)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
  })

  it('identifies references passed to ISREF before scalar dereferencing', () => {
    const context = {
      sheetName: 'Sheet1',
      currentAddress: 'C3',
      resolveCell: (): CellValue => ({ tag: ValueTag.Number, value: 1 }),
      resolveRange: (): CellValue[] => [{ tag: ValueTag.Number, value: 1 }],
    }

    expect(evaluateAst(parseFormula('ISREF(A1)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF(A1:B2)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF(A:A)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF(1:1)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF(OFFSET(A1,0,0))'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF(INDEX(A1:B2,1,1))'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: true,
    })
    expect(evaluateAst(parseFormula('ISREF("A1")'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(evaluateAst(parseFormula('ISREF(A1+1)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(evaluateAst(parseFormula('ISREF({1,2})'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(evaluateAst(parseFormula('ISREF(#REF!)'), context)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
  })

  it('propagates original error inputs for ISEVEN and ISODD', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('ISEVEN(NA())'), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(evaluateAst(parseFormula('ISODD(1/0)'), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Div0,
    })
    expect(evaluateAst(parseFormula('ISEVEN("bad")'), context)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('coerces direct numeric text for eligible aggregates without coercing referenced text cells', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return { tag: ValueTag.String, value: '2', stringId: 1 }
        }
        if (address === 'A2') {
          return { tag: ValueTag.String, value: 'bad', stringId: 2 }
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A2') {
          return [
            { tag: ValueTag.String, value: '2', stringId: 1 },
            { tag: ValueTag.String, value: 'bad', stringId: 2 },
          ]
        }
        return []
      },
    }

    expect(evaluateAst(parseFormula('PRODUCT("2","3")'), context)).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(evaluateAst(parseFormula('MIN("2","3")'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MAX("2","3")'), context)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(evaluateAst(parseFormula('COUNT("2","bad")'), context)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(evaluateAst(parseFormula('AVERAGE("2","4")'), context)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(evaluateAst(parseFormula('AVG("2","4")'), context)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(evaluateAst(parseFormula('SUM("",4)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('AVERAGE("",4)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('COUNT("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('MIN("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('MAX("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('PRODUCT("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('SUMSQ("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('GEOMEAN("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('HARMEAN("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('AVERAGE("bad",4)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(evaluateAst(parseFormula('PRODUCT("2",A1)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('COUNT(A1:A2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('AVERAGE(A1:A2)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(evaluateAst(parseFormula('MIN(A1:A2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('MAX(A1:A2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('returns #DIV/0! for average families when references provide no numeric values', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Empty },
      A2: { tag: ValueTag.String, value: 'skip', stringId: 1 },
      A3: { tag: ValueTag.String, value: '', stringId: 2 },
      B1: { tag: ValueTag.Empty },
      B2: { tag: ValueTag.Empty },
      C1: { tag: ValueTag.Boolean, value: true },
      C2: { tag: ValueTag.Boolean, value: false },
    }
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A3') {
          return ['A1', 'A2', 'A3'].map((address) => valuesByAddress[address])
        }
        if (start === 'B1' && end === 'B2') {
          return ['B1', 'B2'].map((address) => valuesByAddress[address])
        }
        if (start === 'C1' && end === 'C2') {
          return ['C1', 'C2'].map((address) => valuesByAddress[address])
        }
        return []
      },
    }

    const div0Error = { tag: ValueTag.Error, code: ErrorCode.Div0 } as const
    expect(evaluateAst(parseFormula('AVERAGE(A1:A3)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('AVG(A1:A3)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('AVERAGEA(B1:B2)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('SUBTOTAL(1,A1:A3)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('SUBTOTAL(101,A1:A3)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('AGGREGATE(1,6,A1:A3)'), context)).toEqual(div0Error)
    expect(evaluateAst(parseFormula('SUBTOTAL(1,C1:C2)'), context)).toEqual(div0Error)
  })

  it('requires COUNTBLANK to evaluate one referenced range or cell', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Empty },
      A2: { tag: ValueTag.String, value: '', stringId: 1 },
      A3: { tag: ValueTag.Number, value: 2 },
    }
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A3') {
          return ['A1', 'A2', 'A3'].map((address) => valuesByAddress[address])
        }
        return []
      },
    }
    const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const

    expect(evaluateAst(parseFormula('COUNTBLANK(A1:A3)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('COUNTBLANK(A1)'), context)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(evaluateAst(parseFormula('COUNTBLANK()'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('COUNTBLANK("")'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('COUNTBLANK(A1,A2)'), context)).toEqual(valueError)
  })

  it('ignores referenced non-numeric aggregate values without changing direct literal coercion', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Number, value: 2 },
      A2: { tag: ValueTag.Number, value: 4 },
      A3: { tag: ValueTag.String, value: '5', stringId: 1 },
      A4: { tag: ValueTag.String, value: 'bad', stringId: 2 },
      A5: { tag: ValueTag.Boolean, value: true },
      A6: { tag: ValueTag.Empty },
    }
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        if (start === 'A1' && end === 'A6') {
          return ['A1', 'A2', 'A3', 'A4', 'A5', 'A6'].map((address) => valuesByAddress[address])
        }
        return []
      },
    }

    expect(evaluateAst(parseFormula('SUM(A1:A6)'), context)).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(evaluateAst(parseFormula('SUM(A5)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('PRODUCT(A1:A6)'), context)).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(evaluateAst(parseFormula('MIN(A1:A6)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MIN(A3)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('MIN(A5,2)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MAX(A3)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('MAX(A5,0)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('COUNT(A1:A6)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('COUNT("2",A3,A5,A6)'), context)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(evaluateAst(parseFormula('AVERAGE(A1:A6)'), context)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(evaluateAst(parseFormula('AVERAGE(2,A3,A5,A6)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('applies reference-specific logical filtering for AND, OR, and XOR', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Number, value: 2 },
      A2: { tag: ValueTag.Number, value: 4 },
      A3: { tag: ValueTag.String, value: 'ignored', stringId: 1 },
      A4: { tag: ValueTag.String, value: 'also ignored', stringId: 2 },
      A5: { tag: ValueTag.Boolean, value: true },
      A6: { tag: ValueTag.Empty },
      A7: { tag: ValueTag.Number, value: 0 },
      A8: { tag: ValueTag.Boolean, value: false },
      B1: { tag: ValueTag.Error, code: ErrorCode.NA },
    }
    const addresses = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'B1']
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        const startIndex = addresses.indexOf(start)
        const endIndex = addresses.indexOf(end)
        if (startIndex >= 0 && endIndex >= startIndex) {
          return addresses.slice(startIndex, endIndex + 1).map((address) => valuesByAddress[address] ?? { tag: ValueTag.Empty })
        }
        return []
      },
    }
    const valueError = { tag: ValueTag.Error, code: ErrorCode.Value } as const
    const naError = { tag: ValueTag.Error, code: ErrorCode.NA } as const

    expect(evaluateAst(parseFormula('AND(A1:A6)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('AND(A1:A8)'), context)).toEqual({ tag: ValueTag.Boolean, value: false })
    expect(evaluateAst(parseFormula('AND(A3:A6)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('AND(A3:A4)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('AND(A6)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('AND(TRUE,A6)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })

    expect(evaluateAst(parseFormula('OR(A3:A6)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('OR(A3:A4)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('OR(A6)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('OR(A7:A8)'), context)).toEqual({ tag: ValueTag.Boolean, value: false })

    expect(evaluateAst(parseFormula('XOR(A1:A6)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('XOR(A1:A8)'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('XOR(A3:A4)'), context)).toEqual(valueError)
    expect(evaluateAst(parseFormula('XOR(A6)'), context)).toEqual(valueError)

    expect(evaluateAst(parseFormula('AND(2,4,"bad",TRUE())'), context)).toEqual({ tag: ValueTag.Boolean, value: true })
    expect(evaluateAst(parseFormula('AND(B1,A1)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('AND(FALSE(),B1)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('AND(A7:B1)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('OR(B1,A7)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('OR(TRUE(),B1)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('OR(A5:B1)'), context)).toEqual(naError)
    expect(evaluateAst(parseFormula('XOR(B1,A1)'), context)).toEqual(naError)
  })

  it('applies direct-versus-reference numeric rules for statistical summaries', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Number, value: 2 },
      A2: { tag: ValueTag.Number, value: 4 },
      A3: { tag: ValueTag.String, value: '5', stringId: 1 },
      A4: { tag: ValueTag.String, value: 'bad', stringId: 2 },
      A5: { tag: ValueTag.Boolean, value: true },
      A6: { tag: ValueTag.Empty },
      A7: { tag: ValueTag.Boolean, value: false },
      A8: { tag: ValueTag.Number, value: 0 },
      B1: { tag: ValueTag.Number, value: 2 },
      B2: { tag: ValueTag.Empty },
      C1: { tag: ValueTag.Number, value: -2 },
      C2: { tag: ValueTag.Empty },
    }
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        const addresses = ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'B1', 'B2', 'C1', 'C2']
        const startIndex = addresses.indexOf(start)
        const endIndex = addresses.indexOf(end)
        if (startIndex >= 0 && endIndex >= startIndex) {
          return addresses.slice(startIndex, endIndex + 1).map((address) => valuesByAddress[address])
        }
        return []
      },
    }

    expect(evaluateAst(parseFormula('STDEV(A1:A8)'), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(2, 12),
    })
    expect(evaluateAst(parseFormula('STDEVP(A1:A8)'), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(8 / 3), 12),
    })
    expect(evaluateAst(parseFormula('VAR(A1:A8)'), context)).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(evaluateAst(parseFormula('VARP(A1:A8)'), context)).toEqual({ tag: ValueTag.Number, value: 8 / 3 })
    expect(evaluateAst(parseFormula('STDEV("2","4")'), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(2), 12),
    })
    expect(evaluateAst(parseFormula('STDEV("",1)'), context)).toMatchObject({
      tag: ValueTag.Number,
      value: expect.closeTo(Math.sqrt(0.5), 12),
    })
    expect(evaluateAst(parseFormula('STDEV("")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(evaluateAst(parseFormula('STDEVP("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('VAR("2","4")'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('VAR("",1)'), context)).toEqual({ tag: ValueTag.Number, value: 0.5 })
    expect(evaluateAst(parseFormula('VARP("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('SKEW("",1,2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('SKEW.P("",1,2)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(evaluateAst(parseFormula('MEDIAN(A1:A8)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('LARGE(A1:A8,1)'), context)).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(evaluateAst(parseFormula('SMALL(A1:A8,1)'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('LARGE(TRUE(),1)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(evaluateAst(parseFormula('SMALL(TRUE(),1)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(evaluateAst(parseFormula('LARGE("2",1)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(evaluateAst(parseFormula('SMALL("2",1)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(evaluateAst(parseFormula('MEDIAN("2","4")'), context)).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(evaluateAst(parseFormula('MEDIAN("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('AVEDEV("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(evaluateAst(parseFormula('DEVSQ("")'), context)).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(evaluateAst(parseFormula('AVERAGEA(A1:A8)'), context)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(evaluateAst(parseFormula('MINA(B1:B2)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MAXA(C1:C2)'), context)).toEqual({ tag: ValueTag.Number, value: -2 })
  })

  it('applies direct-versus-reference numeric rules for MODE.SNGL', () => {
    const valuesByAddress: Record<string, CellValue> = {
      A1: { tag: ValueTag.Number, value: 2 },
      A2: { tag: ValueTag.Number, value: 4 },
      A3: { tag: ValueTag.String, value: '2', stringId: 1 },
      A4: { tag: ValueTag.Boolean, value: true },
      A5: { tag: ValueTag.Empty },
      B1: { tag: ValueTag.Number, value: 2 },
      B2: { tag: ValueTag.Number, value: 4 },
      B3: { tag: ValueTag.String, value: 'ignored', stringId: 2 },
      B4: { tag: ValueTag.Empty },
    }
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => valuesByAddress[address] ?? { tag: ValueTag.Empty },
      resolveRange: (_sheetName: string, start: string, end: string): CellValue[] => {
        const addresses = ['A1', 'A2', 'A3', 'A4', 'A5', 'B1', 'B2', 'B3', 'B4']
        const startIndex = addresses.indexOf(start)
        const endIndex = addresses.indexOf(end)
        if (startIndex >= 0 && endIndex >= startIndex) {
          return addresses.slice(startIndex, endIndex + 1).map((address) => valuesByAddress[address])
        }
        return []
      },
    }

    expect(evaluateAst(parseFormula('MODE.SNGL("2","4","2")'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MODE("2","bad")'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('MODE.SNGL(A1:A5)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA })
    expect(evaluateAst(parseFormula('MODE.SNGL(A1:B2)'), context)).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(evaluateAst(parseFormula('MODE(B1:B4)'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA })
  })

  it('preserves incoming formula errors before helper argument validation', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
      resolveRange: (): CellValue[] => [],
    }

    const cases = [
      'ADDRESS(NA(),1)',
      'ADDRESS(1,NA())',
      'DOLLAR(NA(),2)',
      'DOLLAR(1,NA())',
      'FIXED(NA(),2)',
      'FIXED(1,NA())',
      'DOLLARDE(NA(),16)',
      'DOLLARDE(1.08,NA())',
      'DOLLARFR(NA(),16)',
      'DOLLARFR(1.5,NA())',
      'RANDBETWEEN(NA(),1)',
      'RANDBETWEEN(1,NA())',
      'SEQUENCE(NA())',
      'SEQUENCE(1,NA())',
      'MUNIT(NA())',
      'RANDARRAY(NA(),1)',
      'RANDARRAY(1,NA())',
      'CHOOSE(NA(),10,20)',
    ]

    for (const formula of cases) {
      expect(evaluateAstResult(parseFormula(formula), context)).toEqual({
        tag: ValueTag.Error,
        code: ErrorCode.NA,
      })
    }
  })

  it('coerces comma-grouped numeric text in arithmetic expressions', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return { tag: ValueTag.String, value: '61,111', stringId: 1 }
        }
        if (address === 'B1') {
          return { tag: ValueTag.String, value: '72,522', stringId: 2 }
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('B1/A1'), context)).toEqual({
      tag: ValueTag.Number,
      value: 72522 / 61111,
    })
    const cagr = evaluateAst(parseFormula('(POWER(B1/A1,0.3333333333)-1)*100'), context)
    expect(cagr.tag).toBe(ValueTag.Number)
    if (cagr.tag === ValueTag.Number) {
      expect(cagr.value).toBeCloseTo(5.872571270499272, 12)
    }
  })

  it('does not coerce whitespace-only text to zero in arithmetic expressions', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return { tag: ValueTag.String, value: '  ', stringId: 1 }
        }
        if (address === 'A2') {
          return { tag: ValueTag.String, value: '\u3000', stringId: 2 }
        }
        if (address === 'A3') {
          return { tag: ValueTag.String, value: '', stringId: 3 }
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('A1*2'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('IFERROR(A2*2,"fallback")'), context)).toEqual({
      tag: ValueTag.String,
      value: 'fallback',
      stringId: 0,
    })
    expect(evaluateAst(parseFormula('A3*2'), context)).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('validates the left arithmetic operand before propagating right operand errors', () => {
    const context = {
      sheetName: 'Sheet1',
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') {
          return { tag: ValueTag.String, value: '523a', stringId: 1 }
        }
        if (address === 'A2') {
          return { tag: ValueTag.String, value: '42', stringId: 2 }
        }
        return { tag: ValueTag.Empty }
      },
      resolveRange: (): CellValue[] => [],
    }

    expect(evaluateAst(parseFormula('A1+#REF!'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(evaluateAst(parseFormula('#REF!+A1'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(evaluateAst(parseFormula('A2+#REF!'), context)).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('treats omitted SUM arguments as empty values instead of errors', () => {
    expect(
      evaluateAst(parseFormula('SUM(2,3,)'), {
        sheetName: 'Sheet1',
        resolveCell: (): CellValue => ({ tag: ValueTag.Empty }),
        resolveRange: (): CellValue[] => [],
      }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
  })

  it('evaluates the Excel implicit-intersection SINGLE wrapper around lookup results and ranges', () => {
    const num = (value: number): CellValue => ({ tag: ValueTag.Number, value })
    const valueError = (): CellValue => ({ tag: ValueTag.Error, code: ErrorCode.Value })
    const rangeContext = (currentAddress: string) => ({
      sheetName: 'Sheet1',
      currentAddress,
      resolveCell: (_sheetName: string, address: string): CellValue => {
        if (address === 'A1') return num(10)
        if (address === 'A2') return num(20)
        if (address === 'A3') return num(30)
        return { tag: ValueTag.Empty }
      },
      resolveRange: (_sheetName: string, start: string, end: string, refKind: 'cells' | 'rows' | 'cols'): CellValue[] => {
        if (refKind === 'cells' && start === 'A1' && end === 'A3') {
          return [num(10), num(20), num(30)]
        }
        return []
      },
    })

    expect(evaluateAst(parseFormula('SINGLE(INDEX(A1:A3,2))'), rangeContext('B2'))).toEqual(num(20))
    expect(evaluateAst(parseFormula('SINGLE(A1:A3)'), rangeContext('B2'))).toEqual(num(20))
    expect(evaluateAst(parseFormula('A1:A3'), rangeContext('B2'))).toEqual(num(20))
    expect(evaluateAst(parseFormula('A1:A3'), rangeContext('B5'))).toEqual(valueError())
  })
})
