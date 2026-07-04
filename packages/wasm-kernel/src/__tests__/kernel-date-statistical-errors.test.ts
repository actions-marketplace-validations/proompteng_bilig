import {
  BUILTIN,
  BuiltinId,
  ErrorCode,
  ValueTag,
  asciiCodes,
  cellIndex,
  createKernel,
  describe,
  encodeCall,
  encodePushBoolean,
  encodePushCell,
  encodePushError,
  encodePushNumber,
  encodePushRange,
  encodePushString,
  encodeRet,
  expect,
  expectErrorCell,
  expectNumberCell,
  it,
  packConstants,
  packPrograms,
} from './kernel-test-helpers.js'

describe('wasm kernel date and statistical error helpers', () => {
  it('evaluates DAYS, WEEKNUM, WORKDAY, and NETWORKDAYS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(30, 8, 1, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([46097, 46101, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(30),
      new Uint16Array(30),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DAYS, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.WEEKNUM, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WEEKNUM, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushCell(0), encodeCall(BUILTIN.WORKDAY, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NETWORKDAYS, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushCell(0), encodeCall(BUILTIN.NETWORKDAYS, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )
    kernel.uploadConstants(
      new Float64Array([46101, 46094, 46096, 46096, 2, 46094, 1, 46094, 1, 46094, 46101, 46094, 46101]),
      new Uint32Array([0, 2, 3, 5, 7, 9, 11]),
      new Uint32Array([2, 1, 2, 2, 2, 2, 2]),
    )
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    )

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(7)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(11)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(46097)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(46098)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(6)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(5)
  })

  it('evaluates WORKDAY.INTL and NETWORKDAYS.INTL on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 10, 4, 4)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.WORKDAY_INTL, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.WORKDAY_INTL, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NETWORKDAYS_INTL, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.NETWORKDAYS_INTL, 4),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [46094, 1, 7],
      [46094, 2, 7, 46096],
      [46094, 46098, 7],
      [46094, 46098, 7, 46096],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(46096)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(46098)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2)
  })

  it('rejects out-of-range workday serials and invalid INTL weekend codes on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 4, 0, 1, 1)
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.WORKDAY, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NETWORKDAYS, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NETWORKDAYS, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY_INTL, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY_INTL, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.WORKDAY_INTL, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NETWORKDAYS_INTL, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.NETWORKDAYS_INTL, 4),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [-1, 1],
      [1, -10],
      [1, 1, -1],
      [-1, 1],
      [1, 2, -1],
      [-1, 1],
      [1, -10],
      [1, 1, 0],
      [1, 2, 0],
      [1, 2, 1, -1],
    ])
    const targetCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(1, 8, width),
      cellIndex(1, 9, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targetCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targetCells)

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 5, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 6, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 7, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 8, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 9, width), ErrorCode.Num)
  })

  it('evaluates logical and rounding builtins with parity-safe scalar semantics', async () => {
    const kernel = await createKernel()
    kernel.init(8, 8, 4, 4, 4)
    kernel.writeCells(
      new Uint8Array([1, 1, 4, 0, 0, 0, 0, 0]),
      new Float64Array([123.4, 1, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(8),
      new Uint16Array([0, 0, ErrorCode.Value, 0, 0, 0, 0, 0]),
    )
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (1 << 24) | 0,
        (20 << 24) | (8 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (2 << 24) | 1,
        (20 << 24) | (9 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (20 << 24) | (15 << 8) | 1,
        255 << 24,

        (3 << 24) | 2,
        (2 << 24) | 2,
        (20 << 24) | (13 << 8) | 2,
        255 << 24,
      ]),
      new Uint32Array([0, 4, 8, 11]),
      new Uint32Array([4, 4, 3, 4]),
      new Uint32Array([3, 4, 5, 6]),
    )
    kernel.uploadConstants(new Float64Array([-1, 0.5, 1]), new Uint32Array([0, 0, 0, 0]), new Uint32Array([2, 2, 0, 1]))

    kernel.evalBatch(new Uint32Array([3, 4, 5, 6]))

    expect(kernel.readNumbers()[3]).toBe(120)
    expect(kernel.readNumbers()[4]).toBe(1)
    expect(kernel.readTags()[5]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[5]).toBe(0)
    expect(kernel.readTags()[6]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[6]).toBe(ErrorCode.Value)
  })

  it('evaluates statistical special functions on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 10, 11, 1, 1)
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.ERF, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ERF, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.ErfPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.ERFC, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.ErfcPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FISHER, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FISHERINV, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMALN, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.GammalnPrecise, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMA, 1), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[1], [0, 1], [1], [1], [1], [0.5], [0.5493061443340549], [5], [5], [5]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.8427006897475899, 7)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.15729931025241006, 7)
    expectNumberCell(kernel, cellIndex(1, 4, width), 0.15729931025241006, 7)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.5493061443340549, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 0.5, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), Math.log(24), 12)
    expectNumberCell(kernel, cellIndex(1, 8, width), Math.log(24), 12)
    expectNumberCell(kernel, cellIndex(1, 9, width), 24, 10)
  })

  it('evaluates statistical scalar and dispersion builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 14
    kernel.init(32, 12, 16, 2, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 4]), Uint32Array.from([0, 4, 4]), asciiCodes('skip'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1, 2, 3, 0, 1, 2, 3, 4]), new Uint32Array([0, 4]), new Uint32Array([4, 5]))
    kernel.uploadRangeShapes(new Uint32Array([4, 5]), new Uint32Array([1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.STANDARDIZE, 3), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.STDEV, 1), encodeRet()],
      [encodePushNumber(0), encodePushBoolean(true), encodePushCell(5), encodeCall(BUILTIN.STDEVA, 3), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.VAR, 1), encodeRet()],
      [encodePushNumber(0), encodePushBoolean(true), encodePushCell(5), encodeCall(BUILTIN.VARA, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.SKEW, 1), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.KURT, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BUILTIN.NORMDIST, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NORMINV, 3), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.NORMSDIST, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.NORMSINV, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LOGINV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LOGNORMDIST, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [1, 0, 1],
      [],
      [2],
      [],
      [2],
      [],
      [],
      [1, 0, 1],
      [0.8413447460685429, 0, 1],
      [1],
      [0.001],
      [0.5, 0, 1],
      [1, 0, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 1)
    expectNumberCell(kernel, cellIndex(1, 1, width), Math.sqrt(5 / 3), 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1)
    expectNumberCell(kernel, cellIndex(1, 3, width), 5 / 3, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), -1.2, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 0.8413447460685429, 7)
    expectNumberCell(kernel, cellIndex(1, 8, width), 1, 8)
    expectNumberCell(kernel, cellIndex(1, 9, width), 0.8413447460685429, 7)
    expectNumberCell(kernel, cellIndex(1, 10, width), -3.090232306167813, 8)
    expectNumberCell(kernel, cellIndex(1, 11, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 12, width), 0.5, 8)
  })

  it('uses Excel direct-versus-reference numeric rules for statistical summaries on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(48, 10, 8, 2, 8)
    kernel.uploadStrings(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([1, 1, 1, 3]), asciiCodes('254bad'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Empty,
        ValueTag.Boolean,
        ValueTag.Number,
      ]),
      new Float64Array([2, 4, 0, 0, 1, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 3, 0, 0, 0, 0]),
      new Uint16Array(48),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7]), Uint32Array.from([0]), Uint32Array.from([8]))
    kernel.uploadRangeShapes(Uint32Array.from([8]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.STDEV, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.STDEVP, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.VAR, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.VARP, 1), encodeRet()],
      [encodePushString(0), encodePushString(2), encodeCall(BUILTIN.STDEV, 2), encodeRet()],
      [encodePushString(0), encodePushString(2), encodeCall(BUILTIN.VAR, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.MEDIAN, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.LARGE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.SMALL, 2), encodeRet()],
      [encodePushString(0), encodePushString(2), encodeCall(BUILTIN.MEDIAN, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[], [], [], [], [], [], [], [1], [1], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 10 }, (_, index) => cellIndex(1, index, width))))

    expectNumberCell(kernel, cellIndex(1, 0, width), 2, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), Math.sqrt(8 / 3), 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 4, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 8 / 3, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), Math.sqrt(2), 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 2, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 2, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 4, 12)
    expectNumberCell(kernel, cellIndex(1, 8, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 9, width), 3, 12)
  })

  it('evaluates statistical distribution builtins and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(48, 22, 64, 1, 1)
    kernel.writeCells(new Uint8Array(48), new Float64Array(48), new Uint32Array(48), new Uint16Array(48))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(false), encodeCall(BUILTIN.EXPONDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BuiltinId.ExponDist, 3), encodeRet()],
      [encodePushNumber(3), encodePushNumber(4), encodePushBoolean(false), encodeCall(BUILTIN.POISSON, 3), encodeRet()],
      [encodePushNumber(3), encodePushNumber(4), encodePushBoolean(true), encodeCall(BuiltinId.PoissonDist, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.WEIBULL, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.WeibullDist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.GAMMADIST, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.GammaDist, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqDistRt, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.CHISQ_DIST, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(false),
        encodeCall(BUILTIN.BINOMDIST, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.BinomDist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BUILTIN.BINOM_DIST_RANGE, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.BINOM_DIST_RANGE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CRITBINOM, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.BinomInv, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.HYPGEOMDIST, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushBoolean(true),
        encodeCall(BuiltinId.HypgeomDist, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NEGBINOMDIST, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.NegbinomDist, 4),
        encodeRet(),
      ],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1 + Math.floor(index / width), index % width, width))),
    )
    const constants = packConstants([
      [0.05, 1.5, 100],
      [1, 2],
      [1, 2],
      [3, 2.5],
      [1.5, 2, 3],
      [1.5, 2, 3],
      [2, 3, 2],
      [2, 3, 2],
      [3, 4],
      [3, 4],
      [3, 4],
      [2, 4, 0.5],
      [2, 4, 0.5],
      [6, 0.5, 2, 4],
      [6, 0.5, 2],
      [6, 0.5, 0.7],
      [6, 0.5, 0.7],
      [1, 4, 3, 10],
      [1, 4, 3, 10],
      [2, 3, 0.5],
      [2, 3, 0.5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    const targetCells = Uint32Array.from(
      Array.from({ length: 22 }, (_, index) => cellIndex(1 + Math.floor(index / width), index % width, width)),
    )
    kernel.evalBatch(targetCells)

    expectNumberCell(kernel, targetCells[0], 0.2939945976810081, 9)
    expectNumberCell(kernel, targetCells[1], 0.2706705664732254, 12)
    expectNumberCell(kernel, targetCells[2], 0.8646647167633873, 12)

    expect(kernel.readTags()[targetCells[3]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[4]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[5]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[6]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[7]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[8]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[9]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[10]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[11]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[12]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[13]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[14]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[15]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[16]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[17]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[18]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[19]]).toBe(ValueTag.Error)
    expect(kernel.readTags()[targetCells[20]]).toBe(ValueTag.Number)
    expect(kernel.readTags()[targetCells[21]]).toBe(ValueTag.Error)
  })

  it('returns statistical value errors on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 5, 5, 1, 2)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Empty, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([1, 0.5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(new Uint32Array([2]), new Uint32Array([1]))

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.FISHER, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.GAMMA, 1), encodeRet()],
      [
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BINOM_DIST_RANGE, 4),
        encodeRet(),
      ],
      [encodePushRange(0), encodeCall(BUILTIN.ERF, 1), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodePushNumber(0), encodePushBoolean(false), encodeCall(BUILTIN.POISSON, 3), encodeRet()],
    ])

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[0], [0], [4, 0.5, 3, 2], [], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 5 }, (_, index) => cellIndex(1, index, width))))

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 2, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 4, width), ErrorCode.Ref)
  })

  it('evaluates student-t scalar distribution builtins and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 9
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.T_DIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_DIST_RT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_DIST_2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_INV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_INV_2T, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TINV, 2), encodeRet()],
    ])
    const constants = packConstants([
      [1, 1],
      [1, 1],
      [1, 1],
      [1, 1, 2],
      [0.75, 1],
      [0.5, 1],
      [0.5, 1],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 0.75, 12)
    expectNumberCell(kernel, outputCells[1], 0.25, 12)
    expectNumberCell(kernel, outputCells[2], 0.5, 12)
    expectNumberCell(kernel, outputCells[3], 0.5, 12)
    expectNumberCell(kernel, outputCells[4], 1, 9)
    expectNumberCell(kernel, outputCells[5], 1, 9)
    expectNumberCell(kernel, outputCells[6], 1, 9)
  })

  it('evaluates CONFIDENCE.T on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 2
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE_T, 3), encodeRet()],
    ])
    const constants = packConstants([[0.5, 2, 4]])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 0.764892328404345, 12)
  })

  it('evaluates GAMMA.INV and legacy alias on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 3
    kernel.init(width, 2, width, 1, 1)
    kernel.writeCells(new Uint8Array(width * 2), new Float64Array(width * 2), new Uint32Array(width * 2), new Uint16Array(width * 2))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.GAMMA_INV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.GAMMAINV, 3), encodeRet()],
    ])
    const constants = packConstants([
      [0.08030139707139418, 3, 2],
      [0.08030139707139418, 3, 2],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 2, 10)
    expectNumberCell(kernel, outputCells[1], 2, 10)
  })

  it('evaluates T.TEST and legacy alias on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(width, 4, 2, 2, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 1, 0, 0, 0, 0, 2, 3, 0, 0, 0, 0, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(width * 4),
      new Uint16Array(width * 4),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
      ]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.T_TEST, 4), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TTEST, 4), encodeRet()],
    ])
    const constants = packConstants([
      [2, 1],
      [2, 1],
    ])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 1, 12)
    expectNumberCell(kernel, outputCells[1], 1, 12)
  })

  it('materializes pivots using the actual source width', async () => {
    const kernel = await createKernel()
    kernel.init(16, 1, 1, 1, 16)

    const strings = ['', 'Region', 'Notes', 'Product', 'Sales', 'East', 'Widget', 'West', 'Gizmo', 'priority']
    const offsets = new Uint32Array(strings.length)
    const lengths = new Uint32Array(strings.length)
    const data: number[] = []
    let offset = 0
    strings.forEach((text, index) => {
      offsets[index] = offset
      lengths[index] = text.length
      for (const char of text) {
        data.push(char.charCodeAt(0))
      }
      offset += text.length
    })
    kernel.uploadStrings(offsets, lengths, Uint16Array.from(data))

    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
      ]),
      new Float64Array([0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 7, 0, 0, 0, 5]),
      new Uint32Array([1, 2, 3, 4, 5, 9, 6, 0, 7, 9, 6, 0, 5, 9, 8, 0]),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from(Array.from({ length: 16 }, (_, index) => index)),
      new Uint32Array([0]),
      new Uint32Array([16]),
    )
    kernel.uploadRangeShapes(new Uint32Array([4]), new Uint32Array([4]))

    const materialized = kernel.materializePivotTable(0, 4, Uint32Array.from([0]), Uint32Array.from([3, 2]), Uint8Array.from([1, 2]))

    expect(materialized.rows).toBe(3)
    expect(materialized.cols).toBe(3)
    expect(materialized.tags[0]).toBe(ValueTag.String)
    expect(materialized.stringIds[0]).toBe(1)
    expect(materialized.tags[1]).toBe(ValueTag.String)
    expect(materialized.stringIds[1]).toBe(4)
    expect(materialized.tags[2]).toBe(ValueTag.String)
    expect(materialized.stringIds[2]).toBe(3)
    expect(materialized.tags[3]).toBe(ValueTag.String)
    expect(materialized.stringIds[3]).toBe(5)
    expect(materialized.tags[4]).toBe(ValueTag.Number)
    expect(materialized.numbers[4]).toBe(15)
    expect(materialized.tags[5]).toBe(ValueTag.Number)
    expect(materialized.numbers[5]).toBe(2)
    expect(materialized.tags[6]).toBe(ValueTag.String)
    expect(materialized.stringIds[6]).toBe(7)
    expect(materialized.tags[7]).toBe(ValueTag.Number)
    expect(materialized.numbers[7]).toBe(7)
    expect(materialized.tags[8]).toBe(ValueTag.Number)
    expect(materialized.numbers[8]).toBe(1)
  })
})
