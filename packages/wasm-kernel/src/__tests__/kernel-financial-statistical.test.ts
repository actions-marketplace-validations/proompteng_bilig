import {
  BUILTIN,
  BuiltinId,
  ErrorCode,
  Opcode,
  ValueTag,
  asciiCodes,
  cellIndex,
  createKernel,
  decodeValueTag,
  describe,
  encodeBinary,
  encodeCall,
  encodePushBoolean,
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

describe('wasm kernel financial and statistical helpers', () => {
  it('evaluates LINEST and LOGEST spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 4, 0, 1, 3)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        0,
      ]),
      new Float64Array([5, 1, 0, 0, 2, 0, 8, 2, 0, 0, 4, 0, 11, 3, 0, 0, 8, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(0, 4, width),
        cellIndex(1, 4, width),
        cellIndex(2, 4, width),
      ]),
      Uint32Array.from([0, 3, 6]),
      Uint32Array.from([3, 3, 3]),
      Uint32Array.from([3, 3, 3]),
      Uint32Array.from([1, 1, 1]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 3]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LINEST, 2), encodeRet()],
      [encodePushRange(2), encodePushRange(1), encodeCall(BUILTIN.LOGEST, 2), encodeRet()],
    ])
    const constants = packConstants([[], []])
    const outputCells = Uint32Array.from([cellIndex(0, 2, width), cellIndex(0, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readSpillRows()[cellIndex(0, 2, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(0, 2, width)]).toBe(2)
    expect(
      readSpillValues(kernel, cellIndex(0, 2, width), []).map((value) => (value.tag === ValueTag.Number ? value.value : value)),
    ).toEqual([3, 2])

    expect(kernel.readSpillRows()[cellIndex(0, 3, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(0, 3, width)]).toBe(2)
    const logestValues = readSpillValues(kernel, cellIndex(0, 3, width), []).map((value) =>
      value.tag === ValueTag.Number ? value.value : value,
    )
    expect(logestValues[0]).toBeCloseTo(2, 12)
    expect(logestValues[1]).toBeCloseTo(1, 12)
  })

  it('evaluates rank helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 3, 0, 1, 4)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([10, 20, 20, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(20),
      new Uint16Array(20),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([0]), Uint32Array.from([4]))
    kernel.uploadRangeShapes(Uint32Array.from([1]), Uint32Array.from([4]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK_EQ, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.RANK_AVG, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodeCall(BUILTIN.RANK_EQ, 3), encodeRet()],
    ])
    const constants = packConstants([[20], [20], [20], [30, 1]])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(tags[cell]).toBe(ValueTag.Number))
    expect(numbers[cellIndex(1, 0, width)]).toBe(2)
    expect(numbers[cellIndex(1, 1, width)]).toBe(2)
    expect(numbers[cellIndex(1, 2, width)]).toBe(2.5)
    expect(numbers[cellIndex(1, 3, width)]).toBe(4)
  })

  it('evaluates order-statistics helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(24, 12, 0, 1, 8)
    kernel.writeCells(
      new Uint8Array([
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
      ]),
      new Float64Array([1, 2, 4, 7, 8, 9, 10, 12, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7]), Uint32Array.from([0]), Uint32Array.from([8]))
    kernel.uploadRangeShapes(Uint32Array.from([8]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MEDIAN, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.SMALL, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.LARGE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTILE_EXC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.PERCENTRANK_EXC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE_INC, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.QUARTILE_EXC, 2), encodeRet()],
    ])
    const constants = packConstants([[], [3], [2], [0.25], [0.25], [0.25], [8], [8], [8], [1], [1], [1]])
    const outputCells = Uint32Array.from([
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
      cellIndex(1, 10, width),
      cellIndex(1, 11, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(tags[cell]).toBe(ValueTag.Number))
    expect(numbers[cellIndex(1, 0, width)]).toBe(7.5)
    expect(numbers[cellIndex(1, 1, width)]).toBe(4)
    expect(numbers[cellIndex(1, 2, width)]).toBe(10)
    expect(numbers[cellIndex(1, 3, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 4, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 5, width)]).toBe(2.5)
    expect(numbers[cellIndex(1, 6, width)]).toBe(0.571)
    expect(numbers[cellIndex(1, 7, width)]).toBe(0.571)
    expect(numbers[cellIndex(1, 8, width)]).toBe(0.555)
    expect(numbers[cellIndex(1, 9, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 10, width)]).toBe(3.5)
    expect(numbers[cellIndex(1, 11, width)]).toBe(2.5)
  })

  it('evaluates MODE.MULT and FREQUENCY spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(30, 2, 0, 2, 8)
    kernel.writeCells(
      new Uint8Array([
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
      new Float64Array([1, 2, 2, 3, 3, 4, 79, 85, 78, 85, 50, 81, 60, 80, 90, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(30),
      new Uint16Array(30),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]),
      Uint32Array.from([0, 6, 12]),
      Uint32Array.from([6, 6, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 3]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MODE_MULT, 1), encodeRet()],
      [encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.FREQUENCY, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 2, width)]))

    expect(kernel.readSpillRows()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(1, 0, width)]).toBe(1)
    expect(readSpillValues(kernel, cellIndex(1, 0, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
    ])
    expect(kernel.readSpillRows()[cellIndex(1, 2, width)]).toBe(4)
    expect(kernel.readSpillCols()[cellIndex(1, 2, width)]).toBe(1)
    expect(readSpillValues(kernel, cellIndex(1, 2, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 0 },
    ])
  })

  it('evaluates MODE, CONFIDENCE.NORM, IFS, SWITCH, and XOR on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 6, 2, 12)
    const pooledStrings = ['', 'big', 'small', 'one', 'other', '2'] as const
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 3, 8, 11, 16]),
      Uint32Array.from(pooledStrings.map((value) => value.length)),
      asciiCodes(pooledStrings.join('')),
    )
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Boolean,
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
      new Float64Array([1, 2, 2, 3, 3, 3, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0, 0, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 0, 1, 6, 7, 8]), Uint32Array.from([0, 6]), Uint32Array.from([6, 5]))
    kernel.uploadRangeShapes(Uint32Array.from([6, 5]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.MODE, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.MODE_SNGL, 1), encodeRet()],
      [encodePushString(5), encodePushNumber(0), encodePushString(5), encodeCall(BUILTIN.MODE_SNGL, 3), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.MODE_SNGL, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.CONFIDENCE_NORM, 3), encodeRet()],
      [
        encodePushBoolean(false),
        encodePushString(1),
        encodePushBoolean(true),
        encodePushString(2),
        encodeCall(BUILTIN.IFS, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.SWITCH, 4), encodeRet()],
      [encodePushBoolean(true), encodePushBoolean(false), encodePushBoolean(true), encodeCall(BUILTIN.XOR, 3), encodeRet()],
      [encodePushNumber(0), encodePushBoolean(true), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.SWITCH, 4), encodeRet()],
      [encodePushNumber(0), encodePushString(5), encodePushString(3), encodePushString(4), encodeCall(BUILTIN.SWITCH, 4), encodeRet()],
    ])
    const constants = packConstants([[], [], [4], [], [0.05, 1, 100], [], [3, 3], [], [1], [2]])
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
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(3)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(2)
    expectErrorCell(kernel, cellIndex(1, 3, width), ErrorCode.NA)
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBeCloseTo(0.1959963986120195, 12)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(1, 5, width)]] ?? '').toBe('small')
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(1, 6, width)]] ?? '').toBe('one')
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(2, 0, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(2, 0, width)]] ?? '').toBe('other')
    expect(kernel.readTags()[cellIndex(2, 1, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(2, 1, width)]] ?? '').toBe('other')
  })

  it('evaluates PROB and TRIMMEAN on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(32, 2, 0, 3, 16)
    const cellTags = new Uint8Array(32)
    const cellNumbers = new Float64Array(32)
    ;[1, 2, 3, 4].forEach((value, index) => {
      cellTags[cellIndex(index, 0, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 0, width)] = value
    })
    ;[0.1, 0.2, 0.3, 0.4].forEach((value, index) => {
      cellTags[cellIndex(index, 1, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 1, width)] = value
    })
    ;[1, 2, 4, 7, 8, 9, 10, 12].forEach((value, index) => {
      cellTags[cellIndex(index, 2, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 2, width)] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(32), new Uint16Array(32))
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(3, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(3, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 2, width),
        cellIndex(2, 2, width),
        cellIndex(3, 2, width),
        cellIndex(4, 2, width),
        cellIndex(5, 2, width),
        cellIndex(6, 2, width),
        cellIndex(7, 2, width),
      ]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 8]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 8]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.PROB, 4), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodeCall(BUILTIN.TRIMMEAN, 2), encodeRet()],
    ])
    const constants = packConstants([[2, 3], [0.25]])
    const outputCells = Uint32Array.from([cellIndex(0, 3, width), cellIndex(1, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 3, width)]).toBe(ValueTag.Number)
    expect(tags[cellIndex(1, 3, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(0.5, 12)
    expect(numbers[cellIndex(1, 3, width)]).toBeCloseTo(40 / 6, 12)
  })

  it('evaluates cash-flow rate helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(40, 4, 3, 4, 22)

    const tags = new Uint8Array(40)
    tags.fill(0)
    for (let index = 0; index < 22; index += 1) {
      tags[index] = ValueTag.Number
    }
    const numbers = new Float64Array(40)
    numbers.set(
      [
        -70000, 12000, 15000, 18000, 21000, 26000, -120000, 39000, 30000, 21000, 37000, 46000, -10000, 2750, 4250, 3250, 2750, 39448, 39508,
        39751, 39859, 39904,
      ],
      0,
    )
    kernel.writeCells(tags, numbers, new Uint32Array(40), new Uint16Array(40))
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]),
      Uint32Array.from([0, 6, 12, 17]),
      Uint32Array.from([6, 6, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 5, 5]), Uint32Array.from([1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.IRR, 1), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MIRR, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(2), encodePushRange(3), encodeCall(BUILTIN.XNPV, 3), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodeCall(BUILTIN.XIRR, 2), encodeRet()],
    ])
    const constants = packConstants([[], [0.1, 0.12], [0.09], []])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const resultTags = kernel.readTags()
    const resultNumbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(resultTags[cell]).toBe(ValueTag.Number))
    expect(resultNumbers[cellIndex(3, 0, width)]).toBeCloseTo(0.08663094803653162, 12)
    expect(resultNumbers[cellIndex(3, 1, width)]).toBeCloseTo(0.1260941303659051, 12)
    expect(resultNumbers[cellIndex(3, 2, width)]).toBeCloseTo(2086.647602031535, 9)
    expect(resultNumbers[cellIndex(3, 3, width)]).toBeCloseTo(0.37336253351883136, 12)
  })

  it('evaluates DAYS360 and YEARFRAC on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 0, 3, 8)
    kernel.writeCells(new Uint8Array(8), new Float64Array(8), new Uint32Array(8), new Uint16Array(8))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DAYS360, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushBoolean(true), encodeCall(BUILTIN.DAYS360, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.YEARFRAC, 3), encodeRet()],
    ])
    const constants = packConstants([
      [45320, 45382],
      [45320, 45382],
      [45292, 45474, 3],
    ])
    const outputCells = Uint32Array.from([cellIndex(0, 0, width), cellIndex(0, 1, width), cellIndex(0, 2, width)])

    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(62)
    expect(numbers[cellIndex(0, 1, width)]).toBe(61)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(182 / 365, 12)
  })

  it('evaluates COUNTBLANK, ISOWEEKNUM, and TIMEVALUE on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(12, 1, 3, 1, 4)
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 7]), asciiCodes('1:30 PM'))
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Empty, ValueTag.String, ValueTag.Empty, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(12),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([0]), Uint32Array.from([4]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([2]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.COUNTBLANK, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeCall(BUILTIN.ISOWEEKNUM, 1),
        encodeRet(),
      ],
      [encodePushString(1), encodeCall(BUILTIN.TIMEVALUE, 1), encodeRet()],
    ])
    const constants = packConstants([[], [2024, 1, 1], []])
    const outputCells = Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width)])

    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(0.5625, 12)
  })

  it('returns numeric spill descriptors for SEQUENCE on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 4, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Sequence, 4),
        encodeRet(),
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([0]),
    )
    kernel.uploadConstants(new Float64Array([3, 1, 1, 1]), new Uint32Array([0]), new Uint32Array([4]))

    kernel.evalBatch(new Uint32Array([0]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(1)
    expect(kernel.readSpillRows()[0]).toBe(3)
    expect(kernel.readSpillCols()[0]).toBe(1)
    expect(kernel.readSpillOffsets()[0]).toBe(0)
    expect(kernel.readSpillLengths()[0]).toBe(3)
    expect(Array.from(kernel.readSpillTags().slice(0, kernel.getSpillValueCount()))).toEqual([
      ValueTag.Number,
      ValueTag.Number,
      ValueTag.Number,
    ])
    expect(Array.from(kernel.readSpillNumbers().slice(0, kernel.getSpillValueCount()))).toEqual([1, 2, 3])
  })

  it('evaluates chi-square inverse functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(width, 1, 8, 8, 1)
    kernel.writeCells(new Uint8Array(width), new Float64Array(width), new Uint32Array(width), new Uint16Array(width))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.LEGACY_CHIDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQDIST, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHIINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQ_INV_RT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.LEGACY_CHIINV, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.CHISQ_INV, 2), encodeRet()],
    ])
    const constants = packConstants([
      [18.307, 10],
      [18.307, 10],
      [18.307, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.93, 1],
    ])
    const outputCells = Uint32Array.from([
      cellIndex(0, 0, width),
      cellIndex(0, 1, width),
      cellIndex(0, 2, width),
      cellIndex(0, 3, width),
      cellIndex(0, 4, width),
      cellIndex(0, 5, width),
      cellIndex(0, 6, width),
      cellIndex(0, 7, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(0.0500006, 6)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(18.306973, 6)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(3.2830202867594993, 12)
  })

  it('evaluates chi-square test functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 2, 2, 12)
    kernel.writeCells(
      new Uint8Array([
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
      ]),
      new Float64Array([58, 35, 11, 25, 10, 23, 45.35, 47.65, 17.56, 18.44, 16.09, 16.91, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]), Uint32Array.from([0, 6]), Uint32Array.from([6, 6]))
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([2, 2]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHISQ_TEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CHITEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.LEGACY_CHITEST, 2), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(1, 4, width), cellIndex(1, 5, width), cellIndex(1, 6, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 0]))

    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(1, 4, width), 0.0003082, 7)
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.0003082, 7)
    expectNumberCell(kernel, cellIndex(1, 6, width), 0.0003082, 7)
  })

  it('evaluates f-test and z-test functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 4, 3, 15)
    kernel.writeCells(
      new Uint8Array([
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
      new Float64Array([6, 7, 9, 15, 21, 20, 28, 31, 38, 40, 1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]),
      Uint32Array.from([0, 5, 10]),
      Uint32Array.from([5, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([5, 5, 5]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.F_TEST, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FTEST, 2), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.Z_TEST, 3), encodeRet()],
      [encodePushRange(2), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.ZTEST, 3), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width), cellIndex(2, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array([2, 1, 2, 1]), new Uint32Array([0, 0, 0, 2]), new Uint32Array([0, 0, 2, 2]))

    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, cellIndex(2, 0, width), 0.648317846786175, 12)
    expectNumberCell(kernel, cellIndex(2, 1, width), 0.648317846786175, 12)
    expectNumberCell(kernel, cellIndex(2, 2, width), 0.012673659338733989, 12)
    expectNumberCell(kernel, cellIndex(2, 3, width), 0.012673659338733989, 12)
  })

  it('evaluates beta and f distribution functions and aliases on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(width, 1, 12, 12, 1)
    kernel.writeCells(new Uint8Array(width), new Float64Array(width), new Uint32Array(width), new Uint16Array(width))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETA_DIST, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETADIST, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETA_INV, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.BETAINV, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushBoolean(true), encodeCall(BUILTIN.F_DIST, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_DIST_RT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.FDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LEGACY_FDIST, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_INV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.F_INV_RT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.FINV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.LEGACY_FINV, 3), encodeRet()],
    ])
    const constants = packConstants([
      [2, 8, 10, 1, 3],
      [2, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
    ])
    const outputCells = Uint32Array.from(Array.from({ length: width }, (_, index) => cellIndex(0, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    for (let index = 0; index < width; index += 1) {
      expect(tags[cellIndex(0, index, width)]).toBe(ValueTag.Number)
    }
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(0.6854705810117458, 9)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(0.6854705810117458, 9)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(2, 9)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(2, 9)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(0.99, 9)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.01, 9)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(0.10930991466299911, 8)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(15.206864870947697, 7)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(15.206864870947697, 7)
    expect(numbers[cellIndex(0, 11, width)]).toBeCloseTo(15.206864870947697, 7)
  })

  it('evaluates FILTER and UNIQUE spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    const pooledStrings = ['A', 'a', 'B', 'C']
    kernel.init(18, 6, 4, 3, 12)
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(Uint32Array.from([0, 1, 2, 3]), Uint32Array.from([1, 1, 1, 1]), asciiCodes(pooledStrings.join('')))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
        ValueTag.Boolean,
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
      ]),
      new Float64Array([1, 3, 2, 4, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 2, 3, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 4]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FILTER, 2), encodeRet()],
      [encodePushRange(2), encodeCall(BUILTIN.UNIQUE, 1), encodeRet()],
      [encodePushRange(0), encodePushRange(0), encodePushNumber(0), encodeBinary(Opcode.Gt), encodeCall(BUILTIN.FILTER, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    )
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0, 0, 0]), new Uint32Array([0, 0, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]))

    expect(readSpillValues(kernel, cellIndex(2, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 1, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'A', stringId: 0 },
      { tag: ValueTag.String, value: 'B', stringId: 0 },
      { tag: ValueTag.String, value: 'C', stringId: 0 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
  })

  it('evaluates internal BYROW and BYCOL sum spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 0, 1, 6)
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
      ]),
      new Float64Array([1, 2, 0, 0, 0, 0, 3, 4, 0, 0, 0, 0, 5, 6, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 6, 7, 12, 13]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([2]))
    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.BYROW_SUM, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.BYCOL_SUM, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array(0), new Uint32Array([0, 0]), new Uint32Array([0, 0]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 11 },
    ])
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 12 },
    ])
  })

  it('evaluates internal REDUCE and SCAN sum builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 1, 1, 3)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2]), Uint32Array.from([0]), Uint32Array.from([3]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.REDUCE_SUM, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.SCAN_SUM, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([0]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(decodeValueTag(kernel.readTags()[cellIndex(0, 3, width)] ?? ValueTag.Empty)).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(0, 3, width)]).toBe(6)
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 6 },
    ])
  })

  it('evaluates internal REDUCE and SCAN product builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 1, 1, 3)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2]), Uint32Array.from([0]), Uint32Array.from([3]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.REDUCE_PRODUCT, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.SCAN_PRODUCT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([1, 1]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(decodeValueTag(kernel.readTags()[cellIndex(0, 3, width)] ?? ValueTag.Empty)).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(0, 3, width)]).toBe(24)
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 24 },
    ])
  })

  it('evaluates internal MAKEARRAY sum spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 4
    kernel.init(8, 1, 2, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array(8),
      new Uint32Array(8),
      new Uint16Array(8),
    )
    kernel.uploadPrograms(
      Uint32Array.from([encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.MAKEARRAY_SUM, 2), encodeRet()]),
      Uint32Array.from([0]),
      Uint32Array.from([4]),
      Uint32Array.from([cellIndex(0, 0, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 2]), new Uint32Array([0]), new Uint32Array([2]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 0, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 0, width), [])).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
    ])
  })

  it('evaluates internal BYROW and BYCOL aggregate spill builtins on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 2, 2, 1, 2)
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
      ]),
      new Float64Array([1, 2, 0, 0, 0, 0, 3, 4, 0, 0, 0, 0, 5, 6, 0, 0, 0, 0]),
      new Uint32Array(18),
      new Uint16Array(18),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 6, 7, 12, 13]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([2]))
    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.BYROW_AGGREGATE, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BUILTIN.BYCOL_AGGREGATE, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]),
    )
    kernel.uploadConstants(new Float64Array([2, 6]), new Uint32Array([0, 1]), new Uint32Array([1, 1]))

    kernel.evalBatch(Uint32Array.from([cellIndex(0, 3, width), cellIndex(0, 4, width)]))

    expect(readSpillValues(kernel, cellIndex(0, 3, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1.5 },
      { tag: ValueTag.Number, value: 3.5 },
      { tag: ValueTag.Number, value: 5.5 },
    ])
    expect(readSpillValues(kernel, cellIndex(0, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 3 },
    ])
  })

  it('evaluates numeric aggregate builtins over native SEQUENCE arrays on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(12, 4, 24, 1, 1)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array([3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(12),
      new Uint16Array(12),
    )
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Sum, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Avg, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Min, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Max, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Count, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.CountA, 1),
        encodeRet(),
      ]),
      new Uint32Array([0, 7, 14, 21, 28, 35]),
      new Uint32Array([7, 7, 7, 7, 7, 7]),
      new Uint32Array([1, 2, 3, 4, 5, 6]),
    )
    kernel.uploadConstants(
      new Float64Array([1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]),
      new Uint32Array([0, 3, 6, 9, 12, 15]),
      new Uint32Array([3, 3, 3, 3, 3, 3]),
    )

    kernel.evalBatch(new Uint32Array([1, 2, 3, 4, 5, 6]))

    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBe(6)
    expect(kernel.readNumbers()[2]).toBe(2)
    expect(kernel.readNumbers()[3]).toBe(1)
    expect(kernel.readNumbers()[4]).toBe(3)
    expect(kernel.readNumbers()[5]).toBe(3)
    expect(kernel.readNumbers()[6]).toBe(3)
  })

  it('evaluates TODAY and NOW from the uploaded recalc timestamp on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 1, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([encodeCall(BUILTIN.TODAY, 0), encodeRet(), encodeCall(BUILTIN.NOW, 0), encodeRet()]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 2]),
      new Uint32Array([0, 1]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.uploadVolatileNowSerial(46100.65659722222)

    kernel.evalBatch(new Uint32Array([0, 1]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(46100)
    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBeCloseTo(46100.65659722222, 12)
  })

  it('evaluates TIME, HOUR, MINUTE, SECOND, and WEEKDAY on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(24, 8, 5, 1, 1)
    kernel.writeCells(
      new Uint8Array([1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0.5208333333333334, 0.5208449074074074, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TIME, 3), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.HOUR, 1), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.MINUTE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.SECOND, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeCall(BUILTIN.WEEKDAY, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodePushNumber(3),
        encodeCall(BUILTIN.WEEKDAY, 2),
        encodeRet(),
      ],
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
    kernel.uploadConstants(
      new Float64Array([12, 30, 0, 2026, 3, 15, 2]),
      new Uint32Array([0, 0, 0, 0, 3, 3]),
      new Uint32Array([3, 0, 0, 0, 3, 4]),
    )
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

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0.5208333333333334)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(30)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(7)
  })
})
