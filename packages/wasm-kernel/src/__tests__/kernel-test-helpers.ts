import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
import { expect } from 'vitest'
export { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
export { describe, expect, it } from 'vitest'
export { createKernel } from '../index.js'

export const BUILTIN = {
  CONCAT: BuiltinId.Concat,
  LEN: BuiltinId.Len,
  ISBLANK: BuiltinId.IsBlank,
  ISNUMBER: BuiltinId.IsNumber,
  ISTEXT: BuiltinId.IsText,
  DATE: BuiltinId.Date,
  COUNT: BuiltinId.Count,
  MIN: BuiltinId.Min,
  MAX: BuiltinId.Max,
  YEAR: BuiltinId.Year,
  MONTH: BuiltinId.Month,
  DAY: BuiltinId.Day,
  EDATE: BuiltinId.Edate,
  EOMONTH: BuiltinId.Eomonth,
  EXACT: BuiltinId.Exact,
  INT: BuiltinId.Int,
  ROUNDUP: BuiltinId.RoundUp,
  ROUNDDOWN: BuiltinId.RoundDown,
  TIME: BuiltinId.Time,
  HOUR: BuiltinId.Hour,
  MINUTE: BuiltinId.Minute,
  SECOND: BuiltinId.Second,
  WEEKDAY: BuiltinId.Weekday,
  DAYS: BuiltinId.Days,
  DAYS360: BuiltinId.Days360,
  WORKDAY: BuiltinId.Workday,
  NETWORKDAYS: BuiltinId.Networkdays,
  WEEKNUM: BuiltinId.Weeknum,
  ISOWEEKNUM: BuiltinId.Isoweeknum,
  TIMEVALUE: BuiltinId.Timevalue,
  TODAY: BuiltinId.Today,
  NOW: BuiltinId.Now,
  RAND: BuiltinId.Rand,
  WORKDAY_INTL: BuiltinId.WorkdayIntl,
  NETWORKDAYS_INTL: BuiltinId.NetworkdaysIntl,
  FILTER: BuiltinId.Filter,
  UNIQUE: BuiltinId.Unique,
  BYROW_SUM: BuiltinId.ByrowSum,
  BYCOL_SUM: BuiltinId.BycolSum,
  REDUCE_SUM: BuiltinId.ReduceSum,
  SCAN_SUM: BuiltinId.ScanSum,
  REDUCE_PRODUCT: BuiltinId.ReduceProduct,
  SCAN_PRODUCT: BuiltinId.ScanProduct,
  MAKEARRAY_SUM: BuiltinId.MakearraySum,
  BYROW_AGGREGATE: BuiltinId.ByrowAggregate,
  BYCOL_AGGREGATE: BuiltinId.BycolAggregate,
  LEFT: BuiltinId.Left,
  RIGHT: BuiltinId.Right,
  MID: BuiltinId.Mid,
  TRIM: BuiltinId.Trim,
  UPPER: BuiltinId.Upper,
  LOWER: BuiltinId.Lower,
  REPLACE: BuiltinId.Replace,
  SUBSTITUTE: BuiltinId.Substitute,
  REPT: BuiltinId.Rept,
  FIND: BuiltinId.Find,
  SEARCH: BuiltinId.Search,
  VALUE: BuiltinId.Value,
  NUMBERVALUE: BuiltinId.Numbervalue,
  VALUETOTEXT: BuiltinId.Valuetotext,
  LEFTB: BuiltinId.Leftb,
  MIDB: BuiltinId.Midb,
  RIGHTB: BuiltinId.Rightb,
  FINDB: BuiltinId.Findb,
  LENB: BuiltinId.Lenb,
  SEARCHB: BuiltinId.Searchb,
  REPLACEB: BuiltinId.Replaceb,
  ADDRESS: BuiltinId.Address,
  DOLLAR: BuiltinId.Dollar,
  DOLLARDE: BuiltinId.Dollarde,
  DOLLARFR: BuiltinId.Dollarfr,
  BASE: BuiltinId.Base,
  DECIMAL: BuiltinId.Decimal,
  BIN2DEC: BuiltinId.Bin2dec,
  BIN2HEX: BuiltinId.Bin2hex,
  BIN2OCT: BuiltinId.Bin2oct,
  DEC2BIN: BuiltinId.Dec2bin,
  DEC2HEX: BuiltinId.Dec2hex,
  DEC2OCT: BuiltinId.Dec2oct,
  HEX2BIN: BuiltinId.Hex2bin,
  HEX2DEC: BuiltinId.Hex2dec,
  HEX2OCT: BuiltinId.Hex2oct,
  OCT2BIN: BuiltinId.Oct2bin,
  OCT2DEC: BuiltinId.Oct2dec,
  OCT2HEX: BuiltinId.Oct2hex,
  CHAR: BuiltinId.Char,
  CODE: BuiltinId.Code,
  UNICODE: BuiltinId.Unicode,
  UNICHAR: BuiltinId.Unichar,
  CLEAN: BuiltinId.Clean,
  ASC: BuiltinId.Asc,
  JIS: BuiltinId.Jis,
  DBCS: BuiltinId.Dbcs,
  BAHTTEXT: BuiltinId.Bahttext,
  TEXT: BuiltinId.Text,
  PHONETIC: BuiltinId.Phonetic,
  DAVERAGE: BuiltinId.Daverage,
  DCOUNT: BuiltinId.Dcount,
  DCOUNTA: BuiltinId.Dcounta,
  DGET: BuiltinId.Dget,
  DMAX: BuiltinId.Dmax,
  DMIN: BuiltinId.Dmin,
  DPRODUCT: BuiltinId.Dproduct,
  DSTDEV: BuiltinId.Dstdev,
  DSTDEVP: BuiltinId.Dstdevp,
  DSUM: BuiltinId.Dsum,
  DVAR: BuiltinId.Dvar,
  DVARP: BuiltinId.Dvarp,
  CHOOSE: BuiltinId.Choose,
  TEXTBEFORE: BuiltinId.Textbefore,
  TEXTAFTER: BuiltinId.Textafter,
  TEXTJOIN: BuiltinId.Textjoin,
  TEXTSPLIT: BuiltinId.Textsplit,
  T: BuiltinId.T,
  N: BuiltinId.N,
  TYPE: BuiltinId.Type,
  DELTA: BuiltinId.Delta,
  GESTEP: BuiltinId.Gestep,
  GAUSS: BuiltinId.Gauss,
  PHI: BuiltinId.Phi,
  STANDARDIZE: BuiltinId.Standardize,
  STDEV: BuiltinId.Stdev,
  STDEV_P: BuiltinId.StdevP,
  STDEV_S: BuiltinId.StdevS,
  STDEVA: BuiltinId.Stdeva,
  STDEVP: BuiltinId.Stdevp,
  STDEVPA: BuiltinId.Stdevpa,
  VAR: BuiltinId.Var,
  VAR_P: BuiltinId.VarP,
  VAR_S: BuiltinId.VarS,
  VARA: BuiltinId.Vara,
  VARP: BuiltinId.Varp,
  VARPA: BuiltinId.Varpa,
  SKEW: BuiltinId.Skew,
  SKEW_P: BuiltinId.SkewP,
  KURT: BuiltinId.Kurt,
  NORMDIST: BuiltinId.Normdist,
  NORM_DIST: BuiltinId.NormDist,
  NORMINV: BuiltinId.Norminv,
  NORM_INV: BuiltinId.NormInv,
  NORMSDIST: BuiltinId.Normsdist,
  NORM_S_DIST: BuiltinId.NormSDist,
  NORMSINV: BuiltinId.Normsinv,
  NORM_S_INV: BuiltinId.NormSInv,
  LOGINV: BuiltinId.Loginv,
  LOGNORMDIST: BuiltinId.Lognormdist,
  LOGNORM_DIST: BuiltinId.LognormDist,
  LOGNORM_INV: BuiltinId.LognormInv,
  BITAND: BuiltinId.Bitand,
  BITOR: BuiltinId.Bitor,
  BITXOR: BuiltinId.Bitxor,
  BITLSHIFT: BuiltinId.Bitlshift,
  BITRSHIFT: BuiltinId.Bitrshift,
  CONVERT: BuiltinId.Convert,
  EUROCONVERT: BuiltinId.Euroconvert,
  SINH: BuiltinId.Sinh,
  COSH: BuiltinId.Cosh,
  TANH: BuiltinId.Tanh,
  ASINH: BuiltinId.Asinh,
  ACOSH: BuiltinId.Acosh,
  ATANH: BuiltinId.Atanh,
  ACOT: BuiltinId.Acot,
  ACOTH: BuiltinId.Acoth,
  COT: BuiltinId.Cot,
  COTH: BuiltinId.Coth,
  CSC: BuiltinId.Csc,
  CSCH: BuiltinId.Csch,
  SEC: BuiltinId.Sec,
  SECH: BuiltinId.Sech,
  SIGN: BuiltinId.Sign,
  EVEN: BuiltinId.Even,
  ODD: BuiltinId.Odd,
  FACT: BuiltinId.Fact,
  FACTDOUBLE: BuiltinId.Factdouble,
  COMBIN: BuiltinId.Combin,
  COMBINA: BuiltinId.Combina,
  PERMUT: BuiltinId.Permut,
  PERMUTATIONA: BuiltinId.Permutationa,
  GCD: BuiltinId.Gcd,
  LCM: BuiltinId.Lcm,
  PRODUCT: BuiltinId.Product,
  QUOTIENT: BuiltinId.Quotient,
  MROUND: BuiltinId.Mround,
  GEOMEAN: BuiltinId.Geomean,
  HARMEAN: BuiltinId.Harmean,
  SUMSQ: BuiltinId.Sumsq,
  FLOOR_MATH: BuiltinId.FloorMath,
  FLOOR_PRECISE: BuiltinId.FloorPrecise,
  CEILING_MATH: BuiltinId.CeilingMath,
  CEILING_PRECISE: BuiltinId.CeilingPrecise,
  ISO_CEILING: BuiltinId.IsoCeiling,
  TRUNC: BuiltinId.Trunc,
  SQRTPI: BuiltinId.Sqrtpi,
  SERIESSUM: BuiltinId.Seriessum,
  BESSELI: BuiltinId.Besseli,
  BESSELJ: BuiltinId.Besselj,
  BESSELK: BuiltinId.Besselk,
  BESSELY: BuiltinId.Bessely,
  NA: BuiltinId.Na,
  IFS: BuiltinId.Ifs,
  IFERROR: BuiltinId.Iferror,
  IFNA: BuiltinId.Ifna,
  SWITCH: BuiltinId.Switch,
  XOR: BuiltinId.Xor,
  COUNTIF: BuiltinId.Countif,
  COUNTIFS: BuiltinId.Countifs,
  SUMIF: BuiltinId.Sumif,
  SUMIFS: BuiltinId.Sumifs,
  AVERAGEIF: BuiltinId.Averageif,
  AVERAGEIFS: BuiltinId.Averageifs,
  SUMPRODUCT: BuiltinId.Sumproduct,
  MATCH: BuiltinId.Match,
  LOOKUP: BuiltinId.Lookup,
  AREAS: BuiltinId.Areas,
  ARRAYTOTEXT: BuiltinId.Arraytotext,
  COLUMNS: BuiltinId.Columns,
  ROWS: BuiltinId.Rows,
  TRANSPOSE: BuiltinId.Transpose,
  HSTACK: BuiltinId.Hstack,
  VSTACK: BuiltinId.Vstack,
  MINIFS: BuiltinId.Minifs,
  MAXIFS: BuiltinId.Maxifs,
  INDEX: BuiltinId.Index,
  VLOOKUP: BuiltinId.Vlookup,
  HLOOKUP: BuiltinId.Hlookup,
  XMATCH: BuiltinId.Xmatch,
  XLOOKUP: BuiltinId.Xlookup,
  OFFSET: BuiltinId.Offset,
  TAKE: BuiltinId.Take,
  DROP: BuiltinId.Drop,
  EXPAND: BuiltinId.Expand,
  TRIMRANGE: BuiltinId.Trimrange,
  DATEDIF: BuiltinId.Datedif,
  FV: BuiltinId.Fv,
  FVSCHEDULE: BuiltinId.Fvschedule,
  PV: BuiltinId.Pv,
  PMT: BuiltinId.Pmt,
  NPER: BuiltinId.Nper,
  NPV: BuiltinId.Npv,
  RATE: BuiltinId.Rate,
  IPMT: BuiltinId.Ipmt,
  PPMT: BuiltinId.Ppmt,
  ISPMT: BuiltinId.Ispmt,
  CUMIPMT: BuiltinId.Cumipmt,
  CUMPRINC: BuiltinId.Cumprinc,
  DB: BuiltinId.Db,
  DDB: BuiltinId.Ddb,
  VDB: BuiltinId.Vdb,
  SLN: BuiltinId.Sln,
  SYD: BuiltinId.Syd,
  DISC: BuiltinId.Disc,
  INTRATE: BuiltinId.Intrate,
  RECEIVED: BuiltinId.Received,
  COUPDAYBS: BuiltinId.Coupdaybs,
  COUPDAYS: BuiltinId.Coupdays,
  COUPDAYSNC: BuiltinId.Coupdaysnc,
  COUPNCD: BuiltinId.Coupncd,
  COUPNUM: BuiltinId.Coupnum,
  COUPPCD: BuiltinId.Couppcd,
  PRICEDISC: BuiltinId.Pricedisc,
  YIELDDISC: BuiltinId.Yielddisc,
  PRICEMAT: BuiltinId.Pricemat,
  YIELDMAT: BuiltinId.Yieldmat,
  ODDFPRICE: BuiltinId.Oddfprice,
  ODDFYIELD: BuiltinId.Oddfyield,
  ODDLPRICE: BuiltinId.Oddlprice,
  ODDLYIELD: BuiltinId.Oddlyield,
  PRICE: BuiltinId.Price,
  YIELD: BuiltinId.Yield,
  DURATION: BuiltinId.Duration,
  MDURATION: BuiltinId.Mduration,
  TBILLPRICE: BuiltinId.Tbillprice,
  TBILLYIELD: BuiltinId.Tbillyield,
  TBILLEQ: BuiltinId.Tbilleq,
  EFFECT: BuiltinId.Effect,
  NOMINAL: BuiltinId.Nominal,
  PDURATION: BuiltinId.Pduration,
  RRI: BuiltinId.Rri,
  IRR: BuiltinId.Irr,
  MIRR: BuiltinId.Mirr,
  XNPV: BuiltinId.Xnpv,
  XIRR: BuiltinId.Xirr,
  YEARFRAC: BuiltinId.Yearfrac,
  COUNTBLANK: BuiltinId.Countblank,
  CHOOSECOLS: BuiltinId.Choosecols,
  CHOOSEROWS: BuiltinId.Chooserows,
  CORREL: BuiltinId.Correl,
  COVAR: BuiltinId.Covar,
  PEARSON: BuiltinId.Pearson,
  COVARIANCE_P: BuiltinId.CovarianceP,
  COVARIANCE_S: BuiltinId.CovarianceS,
  FORECAST: BuiltinId.Forecast,
  GROWTH: BuiltinId.Growth,
  INTERCEPT: BuiltinId.Intercept,
  LINEST: BuiltinId.Linest,
  LOGEST: BuiltinId.Logest,
  MEDIAN: BuiltinId.Median,
  MODE: BuiltinId.Mode,
  MODE_SNGL: BuiltinId.ModeSngl,
  MODE_MULT: BuiltinId.ModeMult,
  FREQUENCY: BuiltinId.Frequency,
  SMALL: BuiltinId.Small,
  LARGE: BuiltinId.Large,
  PERCENTILE: BuiltinId.Percentile,
  PERCENTILE_INC: BuiltinId.PercentileInc,
  PERCENTILE_EXC: BuiltinId.PercentileExc,
  PERCENTRANK: BuiltinId.Percentrank,
  PERCENTRANK_INC: BuiltinId.PercentrankInc,
  PERCENTRANK_EXC: BuiltinId.PercentrankExc,
  PROB: BuiltinId.Prob,
  QUARTILE: BuiltinId.Quartile,
  QUARTILE_INC: BuiltinId.QuartileInc,
  QUARTILE_EXC: BuiltinId.QuartileExc,
  RANK: BuiltinId.Rank,
  RANK_EQ: BuiltinId.RankEq,
  RANK_AVG: BuiltinId.RankAvg,
  RSQ: BuiltinId.Rsq,
  SLOPE: BuiltinId.Slope,
  STEYX: BuiltinId.Steyx,
  TREND: BuiltinId.Trend,
  TRIMMEAN: BuiltinId.Trimmean,
  SORT: BuiltinId.Sort,
  SORTBY: BuiltinId.Sortby,
  TOCOL: BuiltinId.Tocol,
  TOROW: BuiltinId.Torow,
  WRAPROWS: BuiltinId.Wraprows,
  WRAPCOLS: BuiltinId.Wrapcols,
  ERF: BuiltinId.Erf,
  ERFC: BuiltinId.Erfc,
  FISHER: BuiltinId.Fisher,
  FISHERINV: BuiltinId.Fisherinv,
  GAMMALN: BuiltinId.Gammaln,
  GAMMA: BuiltinId.Gamma,
  GAMMA_INV: BuiltinId.GammaInv,
  GAMMAINV: BuiltinId.Gammainv,
  CONFIDENCE_NORM: BuiltinId.ConfidenceNorm,
  CONFIDENCE: BuiltinId.Confidence,
  CONFIDENCE_T: BuiltinId.ConfidenceT,
  EXPONDIST: BuiltinId.Expondist,
  POISSON: BuiltinId.Poisson,
  WEIBULL: BuiltinId.Weibull,
  GAMMADIST: BuiltinId.Gammadist,
  CHIDIST: BuiltinId.Chidist,
  LEGACY_CHIDIST: BuiltinId.LegacyChidist,
  CHIINV: BuiltinId.Chiinv,
  CHISQ_INV_RT: BuiltinId.ChisqInvRt,
  CHISQ_INV: BuiltinId.ChisqInv,
  CHISQ_DIST: BuiltinId.ChisqDist,
  CHISQDIST: BuiltinId.Chisqdist,
  CHISQINV: BuiltinId.Chisqinv,
  LEGACY_CHIINV: BuiltinId.LegacyChiinv,
  CHISQ_TEST: BuiltinId.ChisqTest,
  CHITEST: BuiltinId.Chitest,
  LEGACY_CHITEST: BuiltinId.LegacyChitest,
  F_TEST: BuiltinId.FTest,
  FTEST: BuiltinId.Ftest,
  Z_TEST: BuiltinId.ZTest,
  ZTEST: BuiltinId.Ztest,
  BETA_DIST: BuiltinId.BetaDist,
  BETA_INV: BuiltinId.BetaInv,
  BETADIST: BuiltinId.Betadist,
  BETAINV: BuiltinId.Betainv,
  F_DIST: BuiltinId.FDist,
  F_DIST_RT: BuiltinId.FDistRt,
  F_INV: BuiltinId.FInv,
  F_INV_RT: BuiltinId.FInvRt,
  FDIST: BuiltinId.Fdist,
  FINV: BuiltinId.Finv,
  LEGACY_FDIST: BuiltinId.LegacyFdist,
  LEGACY_FINV: BuiltinId.LegacyFinv,
  T_DIST: BuiltinId.TDist,
  T_DIST_RT: BuiltinId.TDistRt,
  T_DIST_2T: BuiltinId.TDist2T,
  T_INV: BuiltinId.TInv,
  T_INV_2T: BuiltinId.TInv2T,
  TDIST: BuiltinId.Tdist,
  TINV: BuiltinId.Tinv,
  T_TEST: BuiltinId.TTest,
  TTEST: BuiltinId.Ttest,
  BINOMDIST: BuiltinId.Binomdist,
  BINOM_DIST_RANGE: BuiltinId.BinomDistRange,
  CRITBINOM: BuiltinId.Critbinom,
  HYPGEOMDIST: BuiltinId.Hypgeomdist,
  NEGBINOMDIST: BuiltinId.Negbinomdist,
} as const

export const OUTPUT_STRING_BASE = 2147483648

export function asciiCodes(text: string): Uint16Array {
  return Uint16Array.from(Array.from(text, (char) => char.charCodeAt(0)))
}

export function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

export function encodePushCell(cellOffset: number): number {
  return (Opcode.PushCell << 24) | cellOffset
}

export function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex
}

export function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

export function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

export function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0)
}

export function encodePushError(code: ErrorCode): number {
  return (Opcode.PushError << 24) | code
}

export function encodeBinary(opcode: Opcode): number {
  return opcode << 24
}

export function encodeRet(): number {
  return Opcode.Ret << 24
}

export function packPrograms(programs: number[][]): {
  programs: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const program of programs) {
    offsets.push(offset)
    lengths.push(program.length)
    flat.push(...program)
    offset += program.length
  }

  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

export function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0

  for (const constants of constantsByProgram) {
    offsets.push(offset)
    lengths.push(constants.length)
    flat.push(...constants)
    offset += constants.length
  }

  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

export function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

export type KernelInstance = Awaited<ReturnType<typeof createKernel>>

export function decodeValueTag(rawTag: number): ValueTag {
  switch (rawTag) {
    case 0:
      return ValueTag.Empty
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    default:
      throw new Error(`Unexpected spill tag: ${rawTag}`)
  }
}

export function decodeErrorCode(rawCode: number): ErrorCode {
  switch (rawCode) {
    case 0:
      return ErrorCode.None
    case 1:
      return ErrorCode.Div0
    case 2:
      return ErrorCode.Ref
    case 3:
      return ErrorCode.Value
    case 4:
      return ErrorCode.Name
    case 5:
      return ErrorCode.NA
    case 6:
      return ErrorCode.Cycle
    case 7:
      return ErrorCode.Spill
    case 8:
      return ErrorCode.Blocked
    default:
      throw new Error(`Unexpected error code: ${rawCode}`)
  }
}

export function readSpillValues(kernel: KernelInstance, ownerCellIndex: number, pooledStrings: readonly string[]): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  const outputStrings = kernel.readOutputStrings()
  return Array.from({ length }, (_, index) => {
    const tag = decodeValueTag(tags[offset + index] ?? ValueTag.Empty)
    const rawValue = values[offset + index] ?? 0
    switch (tag) {
      case ValueTag.Number:
        return { tag, value: rawValue }
      case ValueTag.Boolean:
        return { tag, value: rawValue !== 0 }
      case ValueTag.Empty:
        return { tag }
      case ValueTag.Error:
        return { tag, code: decodeErrorCode(rawValue) }
      case ValueTag.String: {
        const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1
        return {
          tag,
          value: outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : (pooledStrings[rawValue] ?? ''),
          stringId: 0,
        }
      }
    }
    throw new Error('Unexpected decoded spill tag')
  })
}

export function expectNumberCell(kernel: KernelInstance, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

export function expectErrorCell(kernel: KernelInstance, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}
