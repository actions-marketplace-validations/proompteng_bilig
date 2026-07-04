import {
  BUILTIN,
  ErrorCode,
  OUTPUT_STRING_BASE,
  Opcode,
  ValueTag,
  asciiCodes,
  cellIndex,
  createKernel,
  describe,
  encodeBinary,
  encodeCall,
  encodePushBoolean,
  encodePushCell,
  encodePushError,
  encodePushNumber,
  encodePushRange,
  encodePushString,
  encodeRet,
  expect,
  expectNumberCell,
  it,
  packConstants,
  packPrograms,
  readSpillValues,
} from './kernel-test-helpers.js'

describe('wasm kernel scalar, aggregate, and text basics', () => {
  it('evaluates a simple program batch', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 4, 4)
    kernel.writeCells(new Uint8Array([1, 0, 0, 0]), new Float64Array([10, 0, 0, 0]), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (1 << 24) | 0, 7 << 24, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([4]),
      new Uint32Array([1]),
    )
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0]), new Uint32Array([1]))
    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(20)
    expect(kernel.readConstantOffsets()[0]).toBe(0)
    expect(kernel.readConstantLengths()[0]).toBe(1)
    expect(kernel.readConstants()[0]).toBe(2)
  })

  it('evaluates aggregate and numeric builtins', async () => {
    const kernel = await createKernel()
    kernel.init(6, 6, 2, 6, 6)
    kernel.writeCells(new Uint8Array([1, 1, 0, 0, 0, 0]), new Float64Array([2, 3, 0, 0, 0, 0]), new Uint32Array(6), new Uint16Array(6))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (3 << 24) | 1, (20 << 24) | (1 << 8) | 2, (1 << 24) | 0, 5 << 24, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([2]),
    )
    kernel.uploadConstants(new Float64Array([4]), new Uint32Array([0]), new Uint32Array([1]))

    kernel.evalBatch(new Uint32Array([2]))
    expect(kernel.readNumbers()[2]).toBe(9)
  })

  it('evaluates branch programs with jump opcodes', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 4, 4)
    kernel.writeCells(new Uint8Array([2, 0, 0, 0]), new Float64Array([1, 0, 0, 0]), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (19 << 24) | 4, (1 << 24) | 0, (18 << 24) | 5, (1 << 24) | 1, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([1]),
    )
    kernel.uploadConstants(new Float64Array([10, 20]), new Uint32Array([0]), new Uint32Array([2]))

    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(10)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    const errors = kernel.readErrors()
    kernel.writeCells(
      new Uint8Array([2, tags[1], 0, 0]),
      new Float64Array([0, numbers[1], 0, 0]),
      new Uint32Array(4),
      new Uint16Array([0, errors[1], 0, 0]),
    )
    kernel.evalBatch(new Uint32Array([1]))
    expect(kernel.readNumbers()[1]).toBe(20)
  })

  it('evaluates aggregate builtins through uploaded range members', async () => {
    const kernel = await createKernel()
    kernel.init(6, 6, 1, 4, 4)
    kernel.writeCells(new Uint8Array([1, 1, 0, 0, 0, 0]), new Float64Array([2, 3, 0, 0, 0, 0]), new Uint32Array(6), new Uint16Array(6))
    kernel.uploadPrograms(
      new Uint32Array([(4 << 24) | 0, (20 << 24) | (1 << 8) | 1, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([3]),
      new Uint32Array([2]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))
    kernel.uploadRangeShapes(new Uint32Array([2]), new Uint32Array([1]))

    kernel.evalBatch(new Uint32Array([2]))

    expect(kernel.readNumbers()[2]).toBe(5)
    expect(kernel.readRangeLengths()[0]).toBe(2)
    expect(kernel.readRangeMembers()[1]).toBe(1)
  })

  it('evaluates exact-safe logical info builtins with zero-arg, scalar, and range cases', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 2, 2, 2)
    kernel.writeCells(
      new Uint8Array([0, 1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 42, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))

    const packed = packPrograms([
      [encodeCall(BUILTIN.ISBLANK, 0), encodeRet()],
      [encodeCall(BUILTIN.ISNUMBER, 0), encodeRet()],
      [encodeCall(BUILTIN.ISTEXT, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.ISBLANK, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.ISTEXT, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
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
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
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

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value)
  })

  it('evaluates LEN with scalar coercion and range rejection', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 1, 1, 2)
    kernel.uploadStringLengths(Uint32Array.from([0, 0, 0, 0, 0, 0, 0, 0, 0, 5]))
    kernel.writeCells(
      new Uint8Array([0, 2, 1, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 123.45, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, 0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]))

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
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
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(6)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(5)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Ref)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 6, width)]).toBe(ErrorCode.Value)
  })

  it('evaluates EXACT and numeric rounding builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 1, 2)
    kernel.uploadStrings(Uint32Array.from([0, 0, 5, 10]), Uint32Array.from([0, 5, 5, 5]), asciiCodes('AlphaAlphaalpha'))
    kernel.writeCells(
      new Uint8Array([3, 3, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 0, -3.145, 0.07, 0.29, 16.4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeCall(BUILTIN.EXACT, 2), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.INT, 1), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDUP, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDDOWN, 2), encodeRet()],
      [encodePushCell(3), encodePushNumber(0), encodeCall(BUILTIN.ROUNDUP, 2), encodeRet()],
      [encodePushCell(4), encodePushNumber(0), encodeCall(BUILTIN.ROUNDDOWN, 2), encodeRet()],
      [encodePushCell(5), encodePushNumber(0), encodeCall(BUILTIN.TRUNC, 2), encodeRet()],
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
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0, 0, 0, 0, 0, 0, 0]), new Uint32Array([0, 0, 1, 1, 1, 1, 1]))
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

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(-3.15)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-3.14)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(0.07)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(0.29)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(16.4)
  })

  it('evaluates string literals and CONCAT through the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 2]), Uint32Array.from([0, 2, 3]), asciiCodes('xyfoo'))
    kernel.writeCells(
      new Uint8Array([ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(16),
      new Uint32Array([2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16),
    )
    const packed = packPrograms([
      [encodePushString(1), encodeRet()],
      [encodePushString(1), encodePushCell(0), encodeCall(BUILTIN.CONCAT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['xyfoo'])
  })

  it('evaluates binary text comparison and concat operators on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 5, 10, 11]), Uint32Array.from([0, 5, 5, 1, 1]), asciiCodes('helloHELLObA'))
    kernel.writeCells(
      new Uint8Array([ValueTag.String, ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(24),
      new Uint32Array([1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeBinary(Opcode.Eq), encodeRet()],
      [encodePushString(3), encodePushString(4), encodeBinary(Opcode.Gt), encodeRet()],
      [encodePushCell(0), encodePushString(4), encodeBinary(Opcode.Concat), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['helloA'])
  })

  it('sorts text after numbers in wasm comparisons while keeping zero distinct from empty text', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 1]), Uint32Array.from([0, 1, 1]), asciiCodes(' 1'))
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0]),
      new Float64Array([46023, 12, 999, 0, 0, 0, 0, 0]),
      new Uint32Array(8),
      new Uint16Array(8),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushString(0), encodeBinary(Opcode.Eq), encodeRet()],
      [encodePushCell(1), encodePushString(0), encodeBinary(Opcode.Eq), encodeRet()],
      [encodePushCell(1), encodePushString(1), encodeBinary(Opcode.Lt), encodeRet()],
      [encodePushCell(2), encodePushString(2), encodeBinary(Opcode.Lt), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(1)
  })

  it('evaluates text slicing, casing, and search builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 8, 2, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 21, 26, 30, 38, 40]),
      Uint32Array.from([0, 5, 16, 5, 4, 8, 2, 2]),
      asciiCodes('Alpha  alpha   beta  alphaBETAalphabetphP*'),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
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
      new Float64Array(40),
      new Uint32Array([
        1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodeCall(BUILTIN.LEFT, 2), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.RIGHT, 1), encodeRet()],
      [encodePushCell(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MID, 3), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.TRIM, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.UPPER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LOWER, 1), encodeRet()],
      [encodePushString(6), encodePushCell(4), encodeCall(BUILTIN.FIND, 2), encodeRet()],
      [encodePushString(7), encodePushCell(4), encodeCall(BUILTIN.SEARCH, 2), encodeRet()],
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
        cellIndex(1, 8, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([2, 2, 3]), new Uint32Array([0, 1, 1, 3, 3, 3, 3]), new Uint32Array([1, 0, 2, 0, 0, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Number)
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(3)
    expect(kernel.readOutputStrings()).toEqual(['Al', 'a', 'lph', 'alpha beta', 'ALPHA', 'beta'])
  })

  it('evaluates REPLACE, SUBSTITUTE, and REPT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 8, 9, 15, 17, 19]),
      Uint32Array.from([0, 8, 1, 6, 2, 2, 2]),
      asciiCodes('alphabetZbananaanooxo'),
    )
    kernel.writeCells(
      new Uint8Array([ValueTag.String, ValueTag.String, ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(24),
      new Uint32Array([1, 3, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodePushNumber(1), encodePushString(2), encodeCall(BUILTIN.REPLACE, 4), encodeRet()],
      [encodePushCell(1), encodePushString(4), encodePushString(5), encodeCall(BUILTIN.SUBSTITUTE, 3), encodeRet()],
      [encodePushCell(1), encodePushString(4), encodePushString(5), encodePushNumber(0), encodeCall(BUILTIN.SUBSTITUTE, 4), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.REPT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([3, 2, 2, 3]), new Uint32Array([0, 0, 2, 3]), new Uint32Array([2, 0, 1, 1]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['alZabet', 'booooa', 'banooa', 'xoxoxo'])
  })

  it('evaluates CHOOSE, TEXTBEFORE, TEXTAFTER, and TEXTJOIN on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(32, 7, 2, 1, 1)
    const pooledStrings = ['alpha-beta', '-', 'alpha', 'beta'] as const
    kernel.uploadStrings(Uint32Array.from([0, 0, 10, 11, 16]), Uint32Array.from([0, 10, 1, 5, 4]), asciiCodes('alpha-beta-alphabeta'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Empty,
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
      ]),
      new Float64Array([10, 20, 30, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 3, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(32),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6]), Uint32Array.from([0, 4]), Uint32Array.from([4, 3]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 3]), Uint32Array.from([2, 1]))

    const packed = packPrograms([
      [encodePushString(1), encodePushString(2), encodeCall(BUILTIN.TEXTBEFORE, 2), encodeRet()],
      [encodePushString(1), encodePushString(2), encodeCall(BUILTIN.TEXTAFTER, 2), encodeRet()],
      [encodePushString(2), encodePushBoolean(true), encodePushRange(1), encodeCall(BUILTIN.TEXTJOIN, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHOOSE, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    const constants = packConstants([[], [], [], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readOutputStrings()).toEqual(['alpha', 'beta', 'alpha-beta'])
    expect(kernel.readSpillRows()[cellIndex(1, 4, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 4, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(1, 4, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 40 },
    ])
  })

  it('evaluates TEXTSPLIT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 14, 15]), Uint32Array.from([0, 14, 1, 1]), asciiCodes('red,blue|green,|'))
    const packed = packPrograms([
      [encodePushString(1), encodePushString(2), encodePushString(3), encodeCall(BUILTIN.TEXTSPLIT, 3), encodeRet()],
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, Uint32Array.from([cellIndex(1, 1, width)]))
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width)]))

    expect(kernel.readSpillRows()[cellIndex(1, 1, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 1, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(1, 1, width), [])).toEqual([
      { tag: ValueTag.String, value: 'red', stringId: 0 },
      { tag: ValueTag.String, value: 'blue', stringId: 0 },
      { tag: ValueTag.String, value: 'green', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ])
  })

  it('evaluates VALUE for dynamic scalar inputs on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 0, 4, 16]), Uint32Array.from([0, 4, 12, 3]), asciiCodes('42.5  -17.25e1  not'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Empty,
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
      new Float64Array([0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
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
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0, 0]))
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(42.5)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-172.5)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 4, width)]).toBe(ErrorCode.Value)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Value)
  })

  it('evaluates NUMBERVALUE and VALUETOTEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 8, 9, 10]), Uint32Array.from([8, 1, 1, 5]), asciiCodes('2.500,27,.alpha'))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))
    const packed = packPrograms([
      [encodePushString(0), encodePushString(1), encodePushString(2), encodeCall(BUILTIN.NUMBERVALUE, 3), encodeRet()],
      [encodePushString(3), encodePushNumber(0), encodeCall(BUILTIN.VALUETOTEXT, 2), encodeRet()],
      [encodePushError(ErrorCode.Ref), encodeCall(BUILTIN.VALUETOTEXT, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    const constants = packConstants([[], [1], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(2500.27, 12)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['"alpha"', '#REF!'])
  })

  it('evaluates TEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 8, 18, 28, 36]),
      Uint32Array.from([8, 10, 10, 8, 5]),
      asciiCodes('#,##0.00yyyy-mm-ddh:mm AM/PMprefix @alpha'),
    )
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushString(0), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(1), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(2), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
      [encodePushString(4), encodePushString(3), encodeCall(BUILTIN.TEXT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]),
    )
    const constants = packConstants([[1234.567], [45356], [0.5], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width), cellIndex(1, 4, width)]))

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['1,234.57', '2024-03-05', '12:00 PM', 'prefix alpha'])
  })

  it('evaluates scalar text conversion builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 6, 1, 1, 1)
    kernel.uploadStrings(Uint32Array.from([0, 1, 6]), Uint32Array.from([1, 5, 2]), Uint16Array.from([65, 97, 1, 98, 127, 99, 54, 54]))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.CHAR, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.CODE, 1), encodeRet()],
      [encodePushString(0), encodeCall(BUILTIN.UNICODE, 1), encodeRet()],
      [encodePushString(2), encodeCall(BUILTIN.UNICHAR, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.CLEAN, 1), encodeRet()],
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
    const constants = packConstants([[65], [], [], [], []])
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

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(1, 1, width), 65)
    expectNumberCell(kernel, cellIndex(1, 2, width), 65)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['A', 'B', 'ab\u007fc'])
  })

  it('evaluates ASC, JIS, and DBCS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 1, 1, 1)
    const strings = ['ＡＢＣ　１２３', 'ABC 123', 'ｶﾞｷﾞｸﾞｹﾞｺﾞ']
    const offsets = []
    const lengths = []
    const data = []
    let cursor = 0
    for (const value of strings) {
      const codes = Array.from(value, (char) => char.charCodeAt(0))
      offsets.push(cursor)
      lengths.push(codes.length)
      data.push(...codes)
      cursor += codes.length
    }
    kernel.uploadStrings(Uint32Array.from(offsets), Uint32Array.from(lengths), Uint16Array.from(data))
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BUILTIN.ASC, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.JIS, 1), encodeRet()],
      [encodePushString(2), encodeCall(BUILTIN.DBCS, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    )
    const constants = packConstants([[], [], []])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)]))
    const outputStrings = kernel.readOutputStrings()
    const readStringCell = (index: number): string => {
      expect(kernel.readTags()[index]).toBe(ValueTag.String)
      const raw = kernel.readStringIds()[index] ?? 0
      const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
      return outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : (strings[raw] ?? '')
    }

    expect(readStringCell(cellIndex(1, 0, width))).toBe('ABC 123')
    expect(readStringCell(cellIndex(1, 1, width))).toBe('ＡＢＣ　１２３')
    expect(readStringCell(cellIndex(1, 2, width))).toBe('ガギグゲゴ')
  })

  it('evaluates BAHTTEXT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 0, 0, 0)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BUILTIN.BAHTTEXT, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.BAHTTEXT, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([[1234], [21.25]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน', 'ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์'])
  })

  it('evaluates PHONETIC on the wasm path for scalar text values', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 2, 0, 0, 0)
    kernel.uploadStrings(Uint32Array.from([0]), Uint32Array.from([4]), Uint16Array.from([65, 66, 67, 68]))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array(8),
      new Uint32Array([0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(8),
    )

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BUILTIN.PHONETIC, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.PHONETIC, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    )
    const constants = packConstants([[], [42]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]))

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['ABCD', '42'])
  })

  it('evaluates byte-oriented text builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 8, 4, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 6, 14, 16, 17, 18]),
      Uint32Array.from([6, 8, 2, 1, 1, 1]),
      Uint16Array.from(Array.from('abcdefalphabetphdZé', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushString(5), encodeCall(BUILTIN.LENB, 1), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.LEFTB, 2), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MIDB, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.RIGHTB, 2), encodeRet()],
      [encodePushString(3), encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.FINDB, 3), encodeRet()],
      [encodePushString(2), encodePushString(1), encodeCall(BUILTIN.SEARCHB, 2), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodePushNumber(1), encodePushString(4), encodeCall(BUILTIN.REPLACEB, 4), encodeRet()],
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
    const constants = packConstants([[], [2], [3, 2], [3], [3], [], [3, 2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
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

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(4)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['ab', 'cd', 'def', 'alZabet'])
  })

  it('evaluates ADDRESS and dollar-format helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 6, 1, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0]),
      Uint32Array.from([7]),
      Uint16Array.from(Array.from("O'Brien", (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ADDRESS, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushString(0),
        encodeCall(BUILTIN.ADDRESS, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLAR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLARDE, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DOLLARFR, 2), encodeRet()],
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
      ]),
    )
    const constants = packConstants([
      [12, 3],
      [2, 28, 3, 1],
      [-1234.5, 1],
      [1.08, 16],
      [1.5, 16],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1.5)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1.08)
    expect(kernel.readOutputStrings()).toEqual(['$C$12', "'O''Brien'!$AB2", '-$1,234.5'])
  })

  it('evaluates bitwise and base-conversion helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 8, 1, 1, 1)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITAND, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITOR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITXOR, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITLSHIFT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BITRSHIFT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.BASE, 3), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BUILTIN.DECIMAL, 2), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2DEC, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2HEX, 1), encodeRet()],
      [encodePushString(1), encodeCall(BUILTIN.BIN2OCT, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2BIN, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2HEX, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DEC2OCT, 2), encodeRet()],
      [encodePushString(2), encodePushNumber(0), encodeCall(BUILTIN.HEX2BIN, 2), encodeRet()],
      [encodePushString(3), encodeCall(BUILTIN.HEX2DEC, 1), encodeRet()],
      [encodePushString(4), encodePushNumber(0), encodeCall(BUILTIN.HEX2OCT, 2), encodeRet()],
      [encodePushString(5), encodePushNumber(0), encodeCall(BUILTIN.OCT2BIN, 2), encodeRet()],
      [encodePushString(6), encodeCall(BUILTIN.OCT2DEC, 1), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodeCall(BUILTIN.OCT2HEX, 2), encodeRet()],
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
      ]),
    )
    const constants = packConstants([
      [6, 3],
      [6, 3],
      [6, 3],
      [1, 4],
      [16, 4],
      [255, 16, 4],
      [16],
      [],
      [],
      [],
      [10, 8],
      [255, 4],
      [15, 4],
      [8],
      [],
      [4],
      [8],
      [],
      [4],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.uploadStrings(
      Uint32Array.from([0, 4, 14, 15, 25, 26, 28]),
      Uint32Array.from([4, 10, 1, 10, 1, 2, 2]),
      asciiCodes('00FF1111111111AFFFFFFFFFFF1217'),
    )
    kernel.evalBatch(
      Uint32Array.from([
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
      ]),
    )

    expectNumberCell(kernel, cellIndex(1, 0, width), 2)
    expectNumberCell(kernel, cellIndex(1, 1, width), 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 5)
    expectNumberCell(kernel, cellIndex(1, 3, width), 16)
    expectNumberCell(kernel, cellIndex(1, 4, width), 1)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual([
      '00FF',
      'FFFFFFFFFF',
      '7777777777',
      '00001010',
      '00FF',
      '0017',
      '00001010',
      '0017',
      '00001010',
      '000F',
    ])
    expectNumberCell(kernel, cellIndex(1, 6, width), 255)
    expectNumberCell(kernel, cellIndex(1, 7, width), -1)
    expect(kernel.readTags()[cellIndex(2, 0, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 1, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 4, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(2, 5, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(2, 6, width), -1)
    expect(kernel.readTags()[cellIndex(2, 7, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.String)
    expectNumberCell(kernel, cellIndex(3, 1, width), 15)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.String)
  })

  it('evaluates Bessel engineering helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 4, 0, 1, 1)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELI, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELJ, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELK, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.BESSELY, 2), encodeRet()],
    ])
    const constants = packConstants([
      [1.5, 1],
      [1.9, 2],
      [1.5, 1],
      [2.5, 1],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.981666428, 7)
    expectNumberCell(kernel, cellIndex(1, 1, width), 0.329925728, 7)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.277387804, 7)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.145918138, 7)
  })

  it('evaluates CONVERT and EUROCONVERT on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 4, 5, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 2, 4, 5, 6, 9, 12]),
      Uint32Array.from([0, 2, 2, 1, 1, 3, 3, 3]),
      asciiCodes('mikmFCDEMEURFRF'),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushString(1), encodePushString(2), encodeCall(BUILTIN.CONVERT, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.CONVERT, 3), encodeRet()],
      [encodePushNumber(0), encodePushString(5), encodePushString(6), encodeCall(BUILTIN.EUROCONVERT, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushString(7),
        encodePushString(5),
        encodePushBoolean(true),
        encodePushNumber(1),
        encodeCall(BUILTIN.EUROCONVERT, 5),
        encodeRet(),
      ],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([[6], [68], [1.2], [1, 3]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 0, width), 9.656064, 12)
    expectNumberCell(kernel, cellIndex(1, 1, width), 20, 12)
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.61, 12)
    expectNumberCell(kernel, cellIndex(1, 3, width), 0.29728616, 12)
  })
})
