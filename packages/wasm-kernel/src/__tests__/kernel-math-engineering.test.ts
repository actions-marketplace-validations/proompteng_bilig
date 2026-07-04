import {
  BUILTIN,
  ErrorCode,
  Opcode,
  ValueTag,
  asciiCodes,
  cellIndex,
  createKernel,
  describe,
  encodeBinary,
  encodeCall,
  encodePushCell,
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
  readSpillValues,
} from './kernel-test-helpers.js'

describe('wasm kernel math and engineering helpers', () => {
  it('evaluates promoted scalar and reducer math helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 26, 0, 4, 4)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
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
      new Float64Array([
        2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array(40),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1, 2]), new Uint32Array([0]), new Uint32Array([3]))
    kernel.uploadRangeShapes(new Uint32Array([3]), new Uint32Array([1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.ACOSH, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.COT, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SECH, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.FACT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.COMBIN, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.GCD, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.LCM, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.PRODUCT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.QUOTIENT, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.GEOMEAN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.HARMEAN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.SUMSQ, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.EVEN, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.ODD, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SIGN, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TRUNC, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.FLOOR_MATH, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.FLOOR_PRECISE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CEILING_MATH, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CEILING_PRECISE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ISO_CEILING, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MROUND, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.SQRTPI, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PERMUT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PERMUTATIONA, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.SERIESSUM, 5),
        encodeRet(),
      ],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(2, 0, width),
      cellIndex(2, 1, width),
      cellIndex(2, 2, width),
      cellIndex(2, 3, width),
      cellIndex(2, 4, width),
      cellIndex(2, 5, width),
      cellIndex(2, 6, width),
      cellIndex(2, 7, width),
      cellIndex(3, 0, width),
      cellIndex(3, 1, width),
      cellIndex(3, 2, width),
      cellIndex(3, 3, width),
      cellIndex(3, 4, width),
      cellIndex(3, 5, width),
      cellIndex(3, 6, width),
      cellIndex(3, 7, width),
      cellIndex(4, 0, width),
      cellIndex(4, 1, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([
      [1],
      [1],
      [0],
      [5],
      [8, 3],
      [],
      [],
      [],
      [7, 3],
      [],
      [],
      [],
      [-3],
      [-2],
      [-42],
      [-3.98, 1],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [-5.5, 2],
      [10, 4],
      [2],
      [5, 3],
      [2, 3],
      [2, 1, 2, 1, 2],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 0, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.6420926159343306, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 120, 12)
    expectNumberCell(kernel, cellIndex(1, 4, width), 56, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 12, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 24, 12)
    expectNumberCell(kernel, cellIndex(2, 0, width), 2, 12)
    expectNumberCell(kernel, cellIndex(2, 1, width), 2.8844991406148166, 12)
    expectNumberCell(kernel, cellIndex(2, 2, width), 2.769230769230769, 12)
    expectNumberCell(kernel, cellIndex(2, 3, width), 29, 12)
    expectNumberCell(kernel, cellIndex(2, 4, width), -4, 12)
    expectNumberCell(kernel, cellIndex(2, 5, width), -3, 12)
    expectNumberCell(kernel, cellIndex(2, 6, width), -1, 12)
    expectNumberCell(kernel, cellIndex(2, 7, width), -3.9, 12)
    expectNumberCell(kernel, cellIndex(3, 0, width), -6, 12)
    expectNumberCell(kernel, cellIndex(3, 1, width), -6, 12)
    expectNumberCell(kernel, cellIndex(3, 2, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 3, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 4, width), -4, 12)
    expectNumberCell(kernel, cellIndex(3, 5, width), 12, 12)
    expectNumberCell(kernel, cellIndex(3, 6, width), 2.5066282746310002, 12)
    expectNumberCell(kernel, cellIndex(3, 7, width), 60, 12)
    expectNumberCell(kernel, cellIndex(4, 0, width), 8, 12)
    expectNumberCell(kernel, cellIndex(4, 1, width), 18, 12)
  })

  it('matches Excel direct empty text aggregate argument semantics on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 7, 1, 0, 0)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([0]), asciiCodes(''))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BUILTIN.COUNT, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.MIN, 1), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.MAX, 2), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.PRODUCT, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.SUMSQ, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.GEOMEAN, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.HARMEAN, 1), encodeRet()],
    ])
    const constants = packConstants([[], [], [-1], [], [], [], []])
    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(0, 0, width), 0, 12)
    expectErrorCell(kernel, cellIndex(0, 1, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(0, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(0, 3, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(0, 4, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(0, 5, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(0, 6, width), ErrorCode.Value)
  })

  it('returns #NUM for GEOMEAN and HARMEAN ranges without numeric values on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 2, 0, 1, 2)

    const cellTags = new Uint8Array(16)
    cellTags[0] = ValueTag.String
    cellTags[1] = ValueTag.Boolean
    kernel.writeCells(cellTags, new Float64Array(16), new Uint32Array(16), new Uint16Array(16))
    kernel.uploadRangeMembers(Uint32Array.from([0, 1]), Uint32Array.from([0]), Uint32Array.from([2]))
    kernel.uploadRangeShapes(Uint32Array.from([1]), Uint32Array.from([2]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.GEOMEAN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.HARMEAN, 1), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(0, 0, width), cellIndex(0, 1, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(Float64Array.from([]), Uint32Array.from([0, 0]), Uint32Array.from([0, 0]))
    kernel.evalBatch(outputCells)

    expectErrorCell(kernel, cellIndex(0, 0, width), ErrorCode.Num)
    expectErrorCell(kernel, cellIndex(0, 1, width), ErrorCode.Num)
  })

  it('evaluates exact-parity information and threshold helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 0, 6, 6)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Boolean, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([42, 1, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(0),
      new Uint16Array(0),
    )

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.T, 1), encodeRet()],
      [encodeCall(BUILTIN.N, 0), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.N, 1), encodeRet()],
      [encodeCall(BUILTIN.TYPE, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.TYPE, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(0), encodeCall(BUILTIN.DELTA, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.GESTEP, 2), encodeRet()],
      [encodePushNumber(2), encodeCall(BUILTIN.GAUSS, 1), encodeRet()],
      [encodePushNumber(2), encodeCall(BUILTIN.PHI, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.T, 1), encodeRet()],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(1, 0, width),
      cellIndex(1, 1, width),
      cellIndex(1, 2, width),
      cellIndex(1, 3, width),
      cellIndex(1, 4, width),
      cellIndex(1, 5, width),
      cellIndex(1, 6, width),
      cellIndex(1, 7, width),
      cellIndex(2, 0, width),
      cellIndex(2, 1, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([[4], [2], [0]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectNumberCell(kernel, cellIndex(1, 2, width), 1, 12)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.Value)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 5, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 6, width), 1, 12)
    expectNumberCell(kernel, cellIndex(1, 7, width), 0, 8)
    expectNumberCell(kernel, cellIndex(2, 0, width), 0.3989422804014327, 12)
    expect(kernel.readTags()[cellIndex(2, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['', ''])
  })

  it('evaluates IFERROR, IFNA, and NA on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 8]), Uint32Array.from([0, 8, 7]), asciiCodes('fallbackmissing'))
    kernel.writeCells(
      new Uint8Array([ValueTag.Error, ValueTag.Error, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([ErrorCode.Div0, ErrorCode.Ref, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array([ErrorCode.Div0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
      [encodeCall(BUILTIN.NA, 0), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(1), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(2), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(7)
  })

  it('returns #VALUE for whitespace-only text in wasm arithmetic', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2, 3, 4]), Uint32Array.from([0, 2, 1, 1, 8]), asciiCodes('  \u30005fallback'))
    kernel.writeCells(
      new Uint8Array([ValueTag.String, ValueTag.String, ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(24),
      new Uint32Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodeBinary(Opcode.Mul), encodeRet()],
      [encodePushCell(1), encodePushNumber(0), encodeBinary(Opcode.Mul), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeBinary(Opcode.Mul), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeBinary(Opcode.Mul), encodeRet()],
      [encodePushCell(1), encodePushNumber(0), encodeBinary(Opcode.Mul), encodePushString(4), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    )
    const constants = packConstants([[2], [2], [2], [2], [2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    )

    expectErrorCell(kernel, cellIndex(1, 0, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(1, 1, width), ErrorCode.Value)
    expectNumberCell(kernel, cellIndex(1, 2, width), 10)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 4, width)]).toBe(4)
  })

  it('evaluates conditional aggregates and SUMPRODUCT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(32, 8, 5, 1, 2)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2, 3]), Uint32Array.from([0, 2, 1, 1]), asciiCodes('>0xy'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
      ]),
      new Float64Array([2, 4, -1, 6, 0, 0, 0, 0, 10, 20, 30, 40, 1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 2, 2, 3, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(
      new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]),
      Uint32Array.from([0, 4, 8, 12, 15]),
      Uint32Array.from([4, 4, 4, 3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 4, 3, 3]), Uint32Array.from([1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.COUNTIF, 2), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(1), encodePushString(2), encodeCall(BUILTIN.COUNTIFS, 4), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(2), encodeCall(BUILTIN.SUMIF, 3), encodeRet()],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.SUMIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.AVERAGEIF, 2), encodeRet()],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.AVERAGEIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(3), encodePushRange(4), encodeCall(BUILTIN.SUMPRODUCT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 2, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(70)
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(70)
    expect(kernel.readTags()[cellIndex(3, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 5, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(3, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 6, width)]).toBeCloseTo((10 + 20 + 40) / 3)
    expect(kernel.readTags()[cellIndex(3, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 7, width)]).toBe(32)
  })

  it('evaluates criteria aggregates over dynamic arrays on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 8, 3, 1, 12)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2, 3]), Uint32Array.from([0, 2, 1, 1]), asciiCodes('>2xy'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
        0,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ...Array.from({ length: 16 }, () => ValueTag.Empty),
      ]),
      new Float64Array([
        2,
        4,
        -1,
        6,
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
        10,
        20,
        30,
        40,
        ...Array.from({ length: 16 }, () => 0),
      ]),
      new Uint32Array([0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 3, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...Array.from({ length: 16 }, () => 0)]),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(
      new Uint32Array([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(0, 3, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(2, 3, width),
      ]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([1, 1, 1]), Uint32Array.from([4, 4, 4]))

    const takeNumbers = [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3)]
    const takeLabels = [encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3)]
    const takeAmounts = [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3)]
    const packed = packPrograms([
      [...takeNumbers, encodePushString(1), encodeCall(BUILTIN.COUNTIF, 2), encodeRet()],
      [...takeNumbers, encodePushString(1), ...takeLabels, encodePushString(2), encodeCall(BUILTIN.COUNTIFS, 4), encodeRet()],
      [...takeNumbers, encodePushString(1), ...takeAmounts, encodeCall(BUILTIN.SUMIF, 3), encodeRet()],
      [...takeAmounts, ...takeNumbers, encodePushString(1), ...takeLabels, encodePushString(2), encodeCall(BUILTIN.SUMIFS, 5), encodeRet()],
      [...takeNumbers, encodePushString(1), ...takeAmounts, encodeCall(BUILTIN.AVERAGEIF, 3), encodeRet()],
      [
        ...takeAmounts,
        ...takeNumbers,
        encodePushString(1),
        ...takeLabels,
        encodePushString(2),
        encodeCall(BUILTIN.AVERAGEIFS, 5),
        encodeRet(),
      ],
      [...takeAmounts, ...takeNumbers, encodePushString(1), ...takeLabels, encodePushString(2), encodeCall(BUILTIN.MINIFS, 5), encodeRet()],
      [...takeAmounts, ...takeNumbers, encodePushString(1), ...takeLabels, encodePushString(2), encodeCall(BUILTIN.MAXIFS, 5), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )
    const constants = packConstants([
      [3, 4],
      [3, 4, 3, 4],
      [3, 4, 3, 4],
      [3, 4, 3, 4, 3, 4],
      [3, 4, 3, 4],
      [3, 4, 3, 4, 3, 4],
      [3, 4, 3, 4, 3, 4],
      [3, 4, 3, 4, 3, 4],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 0, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 2, width)]).toBe(60)
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(60)
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(30)
    expect(kernel.readTags()[cellIndex(3, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 5, width)]).toBe(30)
    expect(kernel.readTags()[cellIndex(3, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 6, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(3, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 7, width)]).toBe(40)
  })

  it('evaluates INDEX, VLOOKUP, and HLOOKUP on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 6, 3, 3, 12)
    kernel.uploadStrings(Uint32Array.from([0, 4, 9, 11, 13]), Uint32Array.from([4, 5, 2, 2, 2]), asciiCodes('pearappleQ1Q2Q3'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
        ValueTag.Number,
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
      ]),
      new Float64Array([0, 10, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 100, 200, 300, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 1, 0, 0, 0, 2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 4, 5, 0, 1, 4, 5, 8, 9, 10, 12, 13, 14]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 6]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 2]), Uint32Array.from([2, 2, 3]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.INDEX, 3), encodeRet()],
      [encodePushString(1), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.VLOOKUP, 4), encodeRet()],
      [encodePushString(4), encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.HLOOKUP, 4), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 2, 2, 0, 2, 0]), new Uint32Array([0, 2, 4]), new Uint32Array([2, 2, 2]))

    kernel.evalBatch(Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]))

    expect(kernel.readNumbers()[cellIndex(2, 0, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(2, 1, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(2, 2, width)]).toBe(300)
  })

  it('evaluates table lookup builtins over dynamic arrays on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 8, 1, 5, 20)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => ValueTag.Empty),
      ]),
      new Float64Array([
        1,
        10,
        100,
        0,
        0,
        0,
        0,
        0,
        2,
        20,
        200,
        0,
        0,
        0,
        0,
        0,
        3,
        30,
        300,
        0,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => 0),
      ]),
      new Uint32Array(40),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
      ]),
      Uint32Array.from([0]),
      Uint32Array.from([9]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([3]))

    const takeTable = [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3)]
    const packed = packPrograms([
      [...takeTable, encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.INDEX, 3), encodeRet()],
      [...takeTable, encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.INDEX, 3), encodeRet()],
      [...takeTable, encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.INDEX, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushRange(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.TAKE, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.VLOOKUP, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.TAKE, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.HLOOKUP, 4),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 3, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
      ]),
    )
    const constants = packConstants([
      [2, 3, 2, 3],
      [2, 3, 0, 2],
      [2, 3, 2, 0],
      [2, 2, 3, 3, 0],
      [100, 2, 3, 2, 0],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 3, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 0, width)]).toBe(200)
    expect(readSpillValues(kernel, cellIndex(3, 1, width), [])).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 20 },
    ])
    expect(kernel.readSpillRows()[cellIndex(3, 1, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(3, 1, width)]).toBe(1)
    expect(readSpillValues(kernel, cellIndex(3, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 200 },
    ])
    expect(kernel.readSpillRows()[cellIndex(3, 3, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(3, 3, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(3, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 5, width)]).toBe(200)
    expect(kernel.readTags()[cellIndex(3, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 6, width)]).toBe(200)
  })

  it('evaluates database aggregation builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    const height = 12
    const cellCount = width * height
    kernel.init(cellCount, 3, 0, 12, 3)
    kernel.uploadStrings(Uint32Array.from([0, 3, 9]), Uint32Array.from([3, 6, 5]), asciiCodes('AgeHeightYield'))

    const tags = new Uint8Array(cellCount)
    const numbers = new Float64Array(cellCount)
    const stringIds = new Uint32Array(cellCount)
    const errors = new Uint16Array(cellCount)
    const setNumber = (row: number, col: number, value: number) => {
      const index = cellIndex(row, col, width)
      tags[index] = ValueTag.Number
      numbers[index] = value
    }
    const setString = (row: number, col: number, stringId: number) => {
      const index = cellIndex(row, col, width)
      tags[index] = ValueTag.String
      stringIds[index] = stringId
    }

    setString(0, 0, 0)
    setString(0, 1, 1)
    setString(0, 2, 2)
    setNumber(1, 0, 10)
    setNumber(1, 1, 100)
    setNumber(1, 2, 5)
    setNumber(2, 0, 12)
    setNumber(2, 1, 110)
    setNumber(2, 2, 7)
    setNumber(3, 0, 12)
    setNumber(3, 1, 120)
    setNumber(3, 2, 9)
    setNumber(4, 0, 15)
    setNumber(4, 1, 130)
    setNumber(4, 2, 11)
    setString(0, 4, 0)
    setNumber(1, 4, 12)
    setString(0, 5, 0)
    setNumber(1, 5, 15)
    kernel.writeCells(tags, numbers, stringIds, errors)

    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(4, 0, width),
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(0, 4, width),
        cellIndex(1, 4, width),
        cellIndex(0, 5, width),
        cellIndex(1, 5, width),
      ]),
      Uint32Array.from([0, 15, 17]),
      Uint32Array.from([15, 2, 2]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([5, 2, 2]), Uint32Array.from([3, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DAVERAGE, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DCOUNT, 3), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(1), encodeCall(BUILTIN.DCOUNTA, 3), encodeRet()],
      [encodePushRange(0), encodePushString(1), encodePushRange(2), encodeCall(BUILTIN.DGET, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DMAX, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DMIN, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DPRODUCT, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSTDEV, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSTDEVP, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DSUM, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DVAR, 3), encodeRet()],
      [encodePushRange(0), encodePushString(2), encodePushRange(1), encodeCall(BUILTIN.DVARP, 3), encodeRet()],
    ])
    const outputCells = Uint32Array.from(Array.from({ length: 12 }, (_, index) => cellIndex(index, 7, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(outputCells)

    const outputTags = kernel.readTags()
    outputCells.forEach((index) => expect(outputTags[index]).toBe(ValueTag.Number))
    expectNumberCell(kernel, outputCells[0], 8)
    expectNumberCell(kernel, outputCells[1], 2)
    expectNumberCell(kernel, outputCells[2], 2)
    expectNumberCell(kernel, outputCells[3], 130)
    expectNumberCell(kernel, outputCells[4], 9)
    expectNumberCell(kernel, outputCells[5], 7)
    expectNumberCell(kernel, outputCells[6], 63)
    expectNumberCell(kernel, outputCells[7], Math.SQRT2, 12)
    expectNumberCell(kernel, outputCells[8], 1, 12)
    expectNumberCell(kernel, outputCells[9], 16)
    expectNumberCell(kernel, outputCells[10], 2)
    expectNumberCell(kernel, outputCells[11], 1)
  })

  it('evaluates vector lookup builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 8, 6, 1, 2)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 9, 14, 18, 26]),
      Uint32Array.from([0, 5, 4, 5, 4, 8, 8]),
      asciiCodes('applepearpearplumfallbacknotfound'),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
      ]),
      new Float64Array([
        0, 0, 0, 0, 10, 20, 30, 40, 1, 3, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array([
        1, 2, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(
      new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]),
      Uint32Array.from([0, 4, 8, 4]),
      Uint32Array.from([4, 4, 3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 3, 3]), Uint32Array.from([1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushString(2), encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.MATCH, 3), encodeRet()],
      [encodePushNumber(1), encodePushRange(2), encodePushNumber(2), encodeCall(BUILTIN.MATCH, 3), encodeRet()],
      [encodePushString(2), encodePushRange(0), encodePushNumber(0), encodePushNumber(3), encodeCall(BUILTIN.XMATCH, 4), encodeRet()],
      [encodePushString(2), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.XLOOKUP, 3), encodeRet()],
      [encodePushString(6), encodePushRange(0), encodePushRange(1), encodePushString(5), encodeCall(BUILTIN.XLOOKUP, 4), encodeRet()],
      [
        encodePushNumber(1),
        encodePushRange(2),
        encodePushRange(3),
        encodePushString(5),
        encodePushNumber(3),
        encodeCall(BUILTIN.XLOOKUP, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(1),
        encodePushRange(2),
        encodePushRange(3),
        encodePushString(5),
        encodePushNumber(2),
        encodeCall(BUILTIN.XLOOKUP, 5),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
        cellIndex(4, 6, width),
        cellIndex(4, 7, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([0, 4, 1, -1]), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([1, 2, 2, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
        cellIndex(4, 6, width),
        cellIndex(4, 7, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(4, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(4, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(4, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 3, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(4, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 4, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(4, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(4, 5, width)]).toBe(5)
    expect(kernel.readTags()[cellIndex(4, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 6, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(4, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(4, 7, width)]).toBe(30)
  })

  it('evaluates lookup functions over dynamic arrays and multi-cell XLOOKUP returns on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 8, 4, 4, 20)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => ValueTag.Empty),
      ]),
      new Float64Array([
        1,
        10,
        100,
        0,
        0,
        0,
        0,
        0,
        2,
        20,
        200,
        0,
        0,
        0,
        0,
        0,
        3,
        30,
        300,
        0,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => 0),
      ]),
      new Uint32Array(40),
      new Uint16Array(40),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(2, 0, width),
        cellIndex(2, 1, width),
        cellIndex(2, 2, width),
      ]),
      Uint32Array.from([0, 3, 9, 12]),
      Uint32Array.from([3, 6, 3, 6]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 1, 2]), Uint32Array.from([1, 2, 3, 3]))

    const takeLookupVector = [encodePushRange(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 2)]
    const takeReturnTable = [encodePushRange(1), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 2)]
    const packed = packPrograms([
      [encodePushNumber(0), ...takeLookupVector, encodePushNumber(2), encodeCall(BUILTIN.MATCH, 3), encodeRet()],
      [encodePushNumber(0), ...takeLookupVector, encodePushNumber(2), encodeCall(BUILTIN.XMATCH, 3), encodeRet()],
      [encodePushNumber(0), ...takeLookupVector, ...takeReturnTable, encodeCall(BUILTIN.XLOOKUP, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(2), encodePushRange(3), encodeCall(BUILTIN.XLOOKUP, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 5, width)]),
    )
    const constants = packConstants([[2, 3, 0], [3, 3, 0], [2, 3, 3], [10]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 5, width)]))

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 0, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(3)
    expect(readSpillValues(kernel, cellIndex(3, 2, width), [])).toEqual([
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 200 },
    ])
    expect(kernel.readSpillRows()[cellIndex(3, 2, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(3, 2, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(3, 5, width), [])).toEqual([
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 30 },
    ])
    expect(kernel.readSpillRows()[cellIndex(3, 5, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(3, 5, width)]).toBe(1)
  })

  it('evaluates LOOKUP, AREAS, ARRAYTOTEXT, COLUMNS, and ROWS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 8, 2, 3, 8)
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 1]), asciiCodes('z'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
      ]),
      new Float64Array([1, 0, 3, 4, 10, 20, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(32),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 2, 3, 4, 5, 6, 4, 5]), Uint32Array.from([0, 3, 6]), Uint32Array.from([3, 3, 2]))
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 1]), Uint32Array.from([1, 1, 2]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LOOKUP, 3), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.AREAS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.COLUMNS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.ROWS, 1), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.ARRAYTOTEXT, 1), encodeRet()],
      [encodePushRange(2), encodePushNumber(1), encodeCall(BUILTIN.ARRAYTOTEXT, 2), encodeRet()],
      [encodePushString(1), encodePushString(1), encodeCall(BUILTIN.LOOKUP, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([3.5, 1]), new Uint32Array([0, 0, 0, 0, 0, 0, 0]), new Uint32Array([2, 2, 2, 2, 2, 2, 2]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(20)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readOutputStrings()).toEqual(['10\t20', '{10, 20}'])
  })
})
