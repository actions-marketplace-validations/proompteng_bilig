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
  encodePushBoolean,
  encodePushCell,
  encodePushNumber,
  encodePushRange,
  encodePushString,
  encodeRet,
  expect,
  it,
  packConstants,
  packPrograms,
  readSpillValues,
} from './kernel-test-helpers.js'

describe('wasm kernel lookup and dynamic array helpers', () => {
  it('evaluates TRANSPOSE, HSTACK, VSTACK, MINIFS, and MAXIFS on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    const pooledStrings = ['', 'x', 'a', 'b', 'c', '>0']
    kernel.init(40, 8, 1, 8, 24)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ...Array.from({ length: 16 }, () => ValueTag.Empty),
      ]),
      new Float64Array([
        1,
        0,
        1,
        4,
        10,
        20,
        0,
        0,
        30,
        0,
        40,
        50,
        10,
        0,
        30,
        5,
        2,
        4,
        -1,
        6,
        0,
        0,
        0,
        0,
        ...Array.from({ length: 16 }, () => 0),
      ]),
      new Uint32Array([0, 1, 0, 0, 0, 0, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 2, 3, 2, ...Array.from({ length: 16 }, () => 0)]),
      new Uint16Array(40),
    )
    kernel.uploadStringLengths(Uint32Array.from(pooledStrings.map((value) => value.length)))
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 1, 2, 3, 4]),
      Uint32Array.from(pooledStrings.map((value) => value.length)),
      asciiCodes(pooledStrings.join('')),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]),
      Uint32Array.from([0, 4, 6, 8, 12, 16, 20]),
      Uint32Array.from([4, 2, 2, 4, 4, 4, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 1, 2, 4, 4, 4]), Uint32Array.from([2, 1, 2, 2, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BUILTIN.TRANSPOSE, 1), encodeRet()],
      [encodePushRange(1), encodePushRange(2), encodePushString(4), encodeCall(BUILTIN.HSTACK, 3), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodePushString(4), encodeCall(BUILTIN.VSTACK, 3), encodeRet()],
      [
        encodePushRange(4),
        encodePushRange(5),
        encodePushString(5),
        encodePushRange(6),
        encodePushString(2),
        encodeCall(BUILTIN.MINIFS, 5),
        encodeRet(),
      ],
      [
        encodePushRange(4),
        encodePushRange(5),
        encodePushString(5),
        encodePushRange(6),
        encodePushString(2),
        encodeCall(BUILTIN.MAXIFS, 5),
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
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0, 0]))

    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(readSpillValues(kernel, cellIndex(3, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 4 },
    ])

    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(readSpillValues(kernel, cellIndex(3, 1, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.String, value: 'a', stringId: 0 },
      { tag: ValueTag.String, value: 'b', stringId: 0 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ])

    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.String)
    expect(readSpillValues(kernel, cellIndex(3, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'a', stringId: 0 },
      { tag: ValueTag.String, value: 'b', stringId: 0 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Boolean, value: false },
      { tag: ValueTag.Number, value: 40 },
      { tag: ValueTag.Number, value: 50 },
      { tag: ValueTag.String, value: 'c', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ])

    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(5)
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(10)
  })

  it('evaluates exact-safe date builtins with Excel coercion and errors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 10, 5, 2, 2)
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 3]), asciiCodes('bad'))
    kernel.writeCells(
      new Uint8Array([3, 2, 4, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 0, 45351, 45351.75, 60, 45322, 45337, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    )

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.YEAR, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.MONTH, 1), encodeRet()],
      [encodePushCell(5), encodeCall(BUILTIN.DAY, 1), encodeRet()],
      [encodePushCell(6), encodePushNumber(3), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(0), encodePushNumber(4), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(7), encodePushCell(1), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(4), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
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
        cellIndex(1, 9, width),
      ]),
    )
    kernel.uploadConstants(new Float64Array([2024, 2, 29, 1.9, 1]), new Uint32Array(9), new Uint32Array(9).fill(5))
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
        cellIndex(1, 9, width),
      ]),
    )

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(45351)
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2024)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(29)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(45351)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(45382)
    expect(kernel.readTags()[cellIndex(1, 9, width)]).toBe(ValueTag.Error)
    expect(kernel.readErrors()[cellIndex(1, 9, width)]).toBe(ErrorCode.Ref)
  })

  it('evaluates numeric-only dynamic-array builtins on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(28, 13, 1, 1, 1)
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
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 4, 3, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(28),
      new Uint16Array(28),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]), Uint32Array.from([0, 6]), Uint32Array.from([6, 4]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([3, 1]))

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BUILTIN.OFFSET, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.TAKE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(1), encodeCall(BUILTIN.DROP, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSECOLS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSEROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodePushNumber(4), encodeCall(BUILTIN.SORT, 3), encodeRet()],
      [encodePushRange(1), encodePushRange(1), encodeCall(BUILTIN.SORTBY, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOCOL, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOROW, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPCOLS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(4), encodeCall(BUILTIN.CHOOSECOLS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(4), encodeCall(BUILTIN.CHOOSEROWS, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]),
    )
    kernel.uploadConstants(new Float64Array([0, 1, 2, 3, -1]), new Uint32Array([0, 0, 0, 0, 0]), new Uint32Array([1, 1, 1, 1, 1]))
    kernel.evalBatch(Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]))

    for (const index of [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]) {
      expect(kernel.readTags()[index]).toBe(ValueTag.Number)
    }

    expect(kernel.readSpillRows()[12]).toBe(2)
    expect(kernel.readSpillCols()[12]).toBe(1)
    expect(kernel.readSpillLengths()[12]).toBe(2)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[12], kernel.readSpillOffsets()[12] + 2))).toEqual([1, 4])

    expect(kernel.readSpillRows()[13]).toBe(2)
    expect(kernel.readSpillCols()[13]).toBe(3)
    expect(kernel.readSpillLengths()[13]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[13], kernel.readSpillOffsets()[13] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[14]).toBe(1)
    expect(kernel.readSpillCols()[14]).toBe(3)
    expect(kernel.readSpillLengths()[14]).toBe(3)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[14], kernel.readSpillOffsets()[14] + 3))).toEqual([4, 5, 6])
    expect(kernel.readSpillRows()[15]).toBe(2)
    expect(kernel.readSpillCols()[15]).toBe(1)
    expect(kernel.readSpillLengths()[15]).toBe(2)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[15], kernel.readSpillOffsets()[15] + 2))).toEqual([2, 5])
    expect(kernel.readSpillRows()[16]).toBe(1)
    expect(kernel.readSpillCols()[16]).toBe(3)
    expect(kernel.readSpillLengths()[16]).toBe(3)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[16], kernel.readSpillOffsets()[16] + 3))).toEqual([4, 5, 6])
    expect(kernel.readSpillRows()[17]).toBe(2)
    expect(kernel.readSpillCols()[17]).toBe(3)
    expect(kernel.readSpillLengths()[17]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[17], kernel.readSpillOffsets()[17] + 6))).toEqual([
      4, 5, 6, 1, 2, 3,
    ])
    expect(kernel.readSpillRows()[18]).toBe(4)
    expect(kernel.readSpillCols()[18]).toBe(1)
    expect(kernel.readSpillLengths()[18]).toBe(4)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[18], kernel.readSpillOffsets()[18] + 4))).toEqual([
      1, 2, 3, 4,
    ])
    expect(kernel.readSpillRows()[19]).toBe(6)
    expect(kernel.readSpillCols()[19]).toBe(1)
    expect(kernel.readSpillLengths()[19]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[19], kernel.readSpillOffsets()[19] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[20]).toBe(1)
    expect(kernel.readSpillCols()[20]).toBe(6)
    expect(kernel.readSpillLengths()[20]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[20], kernel.readSpillOffsets()[20] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[21]).toBe(3)
    expect(kernel.readSpillCols()[21]).toBe(2)
    expect(kernel.readSpillLengths()[21]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[21], kernel.readSpillOffsets()[21] + 6))).toEqual([
      1, 2, 3, 4, 5, 6,
    ])
    expect(kernel.readSpillRows()[22]).toBe(2)
    expect(kernel.readSpillCols()[22]).toBe(3)
    expect(kernel.readSpillLengths()[22]).toBe(6)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[22], kernel.readSpillOffsets()[22] + 6))).toEqual([
      1, 3, 5, 2, 4, 6,
    ])
    expect(kernel.readSpillRows()[23]).toBe(2)
    expect(kernel.readSpillCols()[23]).toBe(1)
    expect(kernel.readSpillLengths()[23]).toBe(2)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[23], kernel.readSpillOffsets()[23] + 2))).toEqual([3, 6])
    expect(kernel.readSpillRows()[24]).toBe(1)
    expect(kernel.readSpillCols()[24]).toBe(3)
    expect(kernel.readSpillLengths()[24]).toBe(3)
    expect(Array.from(kernel.readSpillNumbers().slice(kernel.readSpillOffsets()[24], kernel.readSpillOffsets()[24] + 3))).toEqual([4, 5, 6])
  })

  it('preserves general cell values and ignore modes when flattening on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    const pooledStrings = ['', 'x']
    kernel.init(18, 6, 1, 3, 16)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Error,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ...Array.from({ length: 12 }, () => ValueTag.Empty),
      ]),
      new Float64Array([1, 0, 0, 0, 5, 6, ...Array.from({ length: 12 }, () => 0)]),
      new Uint32Array([0, 0, 0, 1, 0, 0, ...Array.from({ length: 12 }, () => 0)]),
      new Uint16Array([0, 0, ErrorCode.NA, 0, 0, 0, ...Array.from({ length: 12 }, () => 0)]),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([3]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.TOCOL, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.TOCOL, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushBoolean(true), encodeCall(BUILTIN.TOROW, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    )
    const constants = packConstants([[3], [2], [1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]))

    expect(readSpillValues(kernel, cellIndex(2, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 1, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Empty },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(2, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Number, value: 6 },
    ])
  })

  it('preserves general cell values across windowing and wrapping on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    const pooledStrings = ['', 'x', 'pad']
    kernel.init(40, 8, 1, 6, 24)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Error,
        ValueTag.Boolean,
        ValueTag.Empty,
        ValueTag.Number,
        ...Array.from({ length: 34 }, () => ValueTag.Empty),
      ]),
      new Float64Array([1, 0, 0, 1, 0, 6, ...Array.from({ length: 34 }, () => 0)]),
      new Uint32Array([0, 1, 0, 0, 0, 0, ...Array.from({ length: 34 }, () => 0)]),
      new Uint16Array([0, 0, ErrorCode.NA, 0, 0, 0, ...Array.from({ length: 34 }, () => 0)]),
    )
    kernel.uploadStrings(Uint32Array.from([0, 0, 1]), Uint32Array.from([0, 1, 3]), asciiCodes('xpad'))
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([3]))

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(2),
        encodeCall(BUILTIN.OFFSET, 5),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(1),
        encodePushNumber(1),
        encodeCall(BUILTIN.OFFSET, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DROP, 3), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BUILTIN.WRAPROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodePushString(2), encodeCall(BUILTIN.WRAPCOLS, 3), encodeRet()],
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
      ]),
    )
    const constants = packConstants([[0, 1, 2], [0, 1], [-1, 2], [0, 1], [4], [4]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
      ]),
    )

    expect(readSpillValues(kernel, cellIndex(3, 0, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.String)
    expect(pooledStrings[kernel.readStringIds()[cellIndex(3, 1, width)]] ?? '').toBe('x')
    expect(readSpillValues(kernel, cellIndex(3, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Empty },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 3, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 4, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.Empty },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.Error, code: ErrorCode.NA },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 5, width), pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Empty },
      { tag: ValueTag.String, value: 'x', stringId: 0 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Error, code: ErrorCode.NA },
      { tag: ValueTag.String, value: 'pad', stringId: 0 },
      { tag: ValueTag.Boolean, value: true },
      { tag: ValueTag.String, value: 'pad', stringId: 0 },
    ])
  })

  it('evaluates metadata and ARRAYTOTEXT over dynamic array results on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 8, 1, 4, 16)
    kernel.uploadStrings(Uint32Array.from([0, 0, 1]), Uint32Array.from([0, 1, 1]), asciiCodes('xy'))
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Boolean,
        ValueTag.Empty,
        ValueTag.String,
        ...Array.from({ length: 18 }, () => ValueTag.Empty),
      ]),
      new Float64Array([1, 0, 3, 1, 0, 0, ...Array.from({ length: 18 }, () => 0)]),
      new Uint32Array([0, 1, 0, 0, 0, 2, ...Array.from({ length: 18 }, () => 0)]),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5]), Uint32Array.from([0]), Uint32Array.from([6]))
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([3]))

    const takeArrayProgram = [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.TAKE, 3)]
    const packed = packPrograms([
      [...takeArrayProgram, encodeCall(BUILTIN.ROWS, 1), encodeRet()],
      [...takeArrayProgram, encodeCall(BUILTIN.COLUMNS, 1), encodeRet()],
      [...takeArrayProgram, encodeCall(BUILTIN.ARRAYTOTEXT, 1), encodeRet()],
      [...takeArrayProgram, encodePushNumber(2), encodeCall(BUILTIN.ARRAYTOTEXT, 2), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 3, width)]),
    )
    const constants = packConstants([
      [2, 2],
      [2, 2],
      [2, 2],
      [2, 2, 1],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 3, width)]))

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 0, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.String)
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.String)
    expect(kernel.readOutputStrings()).toEqual(['1\tx;TRUE\t', '{1, "x";TRUE, }'])
  })

  it('evaluates RAND from the uploaded recalc random sequence on the wasm path', async () => {
    const kernel = await createKernel()
    kernel.init(4, 4, 1, 1, 1)
    kernel.writeCells(new Uint8Array(4), new Float64Array(4), new Uint32Array(4), new Uint16Array(4))
    kernel.uploadPrograms(
      new Uint32Array([
        encodeCall(BUILTIN.RAND, 0),
        encodeRet(),
        encodeCall(BUILTIN.RAND, 0),
        encodeCall(BUILTIN.RAND, 0),
        encodeBinary(Opcode.Add),
        encodeRet(),
      ]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 4]),
      new Uint32Array([0, 1]),
    )
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]))
    kernel.uploadVolatileRandomValues(new Float64Array([0.625, 0.125, 0.875]))

    kernel.evalBatch(new Uint32Array([0, 1]))

    expect(kernel.readTags()[0]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[0]).toBe(0.625)
    expect(kernel.readTags()[1]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[1]).toBeCloseTo(1, 12)
  })

  it('evaluates EXPAND and TRIMRANGE on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(24, 4, 0, 4, 32)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Empty,
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
      ]),
      new Float64Array([0, 0, 0, 0, 10, 20, 0, 1, 2, 0, 30, 40, 0, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 4, width),
        cellIndex(0, 5, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
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
        cellIndex(3, 0, width),
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
      ]),
      Uint32Array.from([0, 4]),
      Uint32Array.from([4, 16]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([2, 4]))

    const packedPrograms = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.EXPAND, 4), encodeRet()],
      [encodePushRange(1), encodeCall(BUILTIN.TRIMRANGE, 1), encodeRet()],
    ])
    const packedConstants = packConstants([[3, 3, 0], []])

    kernel.uploadPrograms(
      packedPrograms.programs,
      packedPrograms.offsets,
      packedPrograms.lengths,
      Uint32Array.from([cellIndex(3, 4, width), cellIndex(3, 5, width)]),
    )
    kernel.uploadConstants(packedConstants.constants, packedConstants.offsets, packedConstants.lengths)

    kernel.evalBatch(Uint32Array.from([cellIndex(3, 4, width), cellIndex(3, 5, width)]))

    expect(readSpillValues(kernel, cellIndex(3, 4, width), [])).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 20 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 30 },
      { tag: ValueTag.Number, value: 40 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
      { tag: ValueTag.Number, value: 0 },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 5, width), [])).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Empty },
    ])
  })

  it('evaluates DATEDIF and financial scalar helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 25
    kernel.init(50, 1, 0, 25, 50)
    kernel.writeCells(new Uint8Array(50), new Float64Array(50), new Uint32Array(50), new Uint16Array(50))
    kernel.uploadStrings(Uint32Array.from([0, 0]), Uint32Array.from([0, 2]), asciiCodes('YM'))

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.DATE, 3),
        encodePushString(1),
        encodeCall(BUILTIN.DATEDIF, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.FVSCHEDULE, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.DB, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.DDB, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.VDB, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.SLN, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.SYD, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.DISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.INTRATE, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.RECEIVED, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.PRICEDISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BUILTIN.YIELDDISC, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.PRICEMAT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.YIELDMAT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodeCall(BUILTIN.ODDFPRICE, 9),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodePushNumber(8),
        encodeCall(BUILTIN.ODDFYIELD, 9),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BUILTIN.ODDLPRICE, 8),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BUILTIN.ODDLYIELD, 8),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLPRICE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLYIELD, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.TBILLEQ, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.EFFECT, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NOMINAL, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PDURATION, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.RRI, 3), encodeRet()],
    ])
    const constants = packConstants([
      [2020, 1, 15, 2021, 3, 20],
      [1000, 0.09, 0.11, 0.1],
      [10000, 1000, 5, 1],
      [2400, 300, 10, 2],
      [2400, 300, 10, 1, 3],
      [10000, 1000, 9],
      [10000, 1000, 9, 1],
      [44927, 45017, 97, 100, 2],
      [44927, 45017, 1000, 1030, 2],
      [44927, 45017, 1000, 0.12, 2],
      [39494, 39508, 0.0525, 100, 2],
      [39494, 39508, 99.795, 100, 2],
      [39493, 39551, 39397, 0.061, 0.061, 0],
      [39522, 39755, 39394, 0.0625, 100.0123, 0],
      [39763, 44256, 39736, 39873, 0.0785, 0.0625, 100, 2, 1],
      [39763, 44256, 39736, 39873, 0.0575, 84.5, 100, 2, 0],
      [39485, 39614, 39370, 0.0375, 0.0405, 100, 2, 0],
      [39558, 39614, 39440, 0.0375, 99.875, 100, 2, 0],
      [39538, 39600, 0.09],
      [39538, 39600, 98.45],
      [39538, 39600, 0.0914],
      [0.12, 12],
      [0.12682503013196977, 12],
      [0.1, 100, 121],
      [2, 100, 121],
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
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
      cellIndex(0, 10, width),
      cellIndex(0, 11, width),
      cellIndex(0, 12, width),
      cellIndex(0, 13, width),
      cellIndex(0, 14, width),
      cellIndex(0, 15, width),
      cellIndex(0, 16, width),
      cellIndex(0, 17, width),
      cellIndex(0, 18, width),
      cellIndex(0, 19, width),
      cellIndex(0, 20, width),
      cellIndex(0, 21, width),
      cellIndex(0, 22, width),
      cellIndex(0, 23, width),
      cellIndex(0, 24, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(2)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(1330.89, 12)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(3690, 12)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(384, 12)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(691.2, 12)
    expect(numbers[cellIndex(0, 5, width)]).toBe(1000)
    expect(numbers[cellIndex(0, 6, width)]).toBe(1800)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(1030.9278350515465, 12)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(99.79583333333333, 12)
    expect(numbers[cellIndex(0, 11, width)]).toBeCloseTo(0.05282257198685834, 12)
    expect(numbers[cellIndex(0, 12, width)]).toBeCloseTo(99.98449887555694, 12)
    expect(numbers[cellIndex(0, 13, width)]).toBeCloseTo(0.060954333691538576, 12)
    expect(numbers[cellIndex(0, 14, width)]).toBeCloseTo(113.597717474079, 12)
    expect(numbers[cellIndex(0, 15, width)]).toBeCloseTo(0.0772455415972989, 11)
    expect(numbers[cellIndex(0, 16, width)]).toBeCloseTo(99.8782860147213, 12)
    expect(numbers[cellIndex(0, 17, width)]).toBeCloseTo(0.0451922356291692, 12)
    expect(numbers[cellIndex(0, 18, width)]).toBeCloseTo(98.45, 12)
    expect(numbers[cellIndex(0, 19, width)]).toBeCloseTo(0.09141696292534264, 12)
    expect(numbers[cellIndex(0, 20, width)]).toBeCloseTo(0.09415149356594302, 12)
    expect(numbers[cellIndex(0, 21, width)]).toBeCloseTo(0.12682503013196977, 12)
    expect(numbers[cellIndex(0, 22, width)]).toBeCloseTo(0.12, 12)
    expect(numbers[cellIndex(0, 23, width)]).toBeCloseTo(2, 12)
    expect(numbers[cellIndex(0, 24, width)]).toBeCloseTo(0.1, 12)
  })

  it('evaluates annuity and cumulative loan helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 11
    kernel.init(22, 1, 0, 11, 28)
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PV, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.PMT, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.NPER, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.RATE, 3), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.IPMT, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.PPMT, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.ISPMT, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.CUMIPMT, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.CUMPRINC, 6),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.FV, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.NPV, 4), encodeRet()],
    ])
    const constants = packConstants([
      [0.1, 2, -576.1904761904761],
      [0.1, 2, 1000],
      [0.1, -576.1904761904761, 1000],
      [48, -200, 8000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.1, 2, -100, -1000],
      [0.1, 100, 200, 300],
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
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
      cellIndex(0, 10, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBeCloseTo(1000.0000000000006, 12)
    expect(numbers[cellIndex(0, 1, width)]).toBeCloseTo(-576.1904761904758, 12)
    expect(numbers[cellIndex(0, 2, width)]).toBeCloseTo(1.9999999999999982, 12)
    expect(numbers[cellIndex(0, 3, width)]).toBeCloseTo(0.007701472488246008, 12)
    expect(numbers[cellIndex(0, 4, width)]).toBeCloseTo(-100, 12)
    expect(numbers[cellIndex(0, 5, width)]).toBeCloseTo(-476.1904761904758, 12)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(-50, 12)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(-11135.232130750845, 9)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(-934.1071234208765, 9)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(1420, 12)
    expect(numbers[cellIndex(0, 10, width)]).toBeCloseTo(481.5927873779113, 12)
  })

  it('evaluates coupon-date and periodic bond helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(20, 1, 0, 10, 20)
    kernel.writeCells(new Uint8Array(20), new Float64Array(20), new Uint32Array(20), new Uint16Array(20))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYBS, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYS, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPDAYSNC, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPNCD, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPNUM, 4), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BUILTIN.COUPPCD, 4), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BUILTIN.PRICE, 7),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BUILTIN.YIELD, 7),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.DURATION, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BUILTIN.MDURATION, 6),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [39493, 43054, 0.0575, 0.065, 100, 2, 0],
      [39493, 42689, 0.0575, 95.04287, 100, 2, 0],
      [43282, 54058, 0.08, 0.09, 2, 1],
      [39448, 42370, 0.08, 0.09, 2, 1],
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
      cellIndex(0, 8, width),
      cellIndex(0, 9, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)

    kernel.evalBatch(outputCells)

    const tags = kernel.readTags()
    const numbers = kernel.readNumbers()
    expect(tags[cellIndex(0, 0, width)]).toBe(ValueTag.Number)
    expect(numbers[cellIndex(0, 0, width)]).toBe(70)
    expect(numbers[cellIndex(0, 1, width)]).toBe(180)
    expect(numbers[cellIndex(0, 2, width)]).toBe(110)
    expect(numbers[cellIndex(0, 3, width)]).toBe(39217)
    expect(numbers[cellIndex(0, 4, width)]).toBe(6)
    expect(numbers[cellIndex(0, 5, width)]).toBe(39036)
    expect(numbers[cellIndex(0, 6, width)]).toBeCloseTo(94.63436162132213, 12)
    expect(numbers[cellIndex(0, 7, width)]).toBeCloseTo(0.065, 7)
    expect(numbers[cellIndex(0, 8, width)]).toBeCloseTo(10.919145281591925, 12)
    expect(numbers[cellIndex(0, 9, width)]).toBeCloseTo(5.735669813918838, 12)
  })

  it('evaluates covariance and regression helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(64, 10, 1, 2, 6)
    const cellTags = new Uint8Array(64)
    const cellNumbers = new Float64Array(64)
    cellTags[cellIndex(0, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(0, 0, width)] = 5
    cellTags[cellIndex(0, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(0, 1, width)] = 1
    cellTags[cellIndex(1, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(1, 0, width)] = 8
    cellTags[cellIndex(1, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(1, 1, width)] = 2
    cellTags[cellIndex(2, 0, width)] = ValueTag.Number
    cellNumbers[cellIndex(2, 0, width)] = 11
    cellTags[cellIndex(2, 1, width)] = ValueTag.Number
    cellNumbers[cellIndex(2, 1, width)] = 3
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(64), new Uint16Array(64))
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
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.CORREL, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVAR, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVARIANCE_P, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.COVARIANCE_S, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.PEARSON, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.INTERCEPT, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.SLOPE, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.RSQ, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.STEYX, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BUILTIN.FORECAST, 3), encodeRet()],
    ])
    const constants = packConstants([[], [], [], [], [], [], [], [], [], [4]])
    const outputCells = Uint32Array.from([
      cellIndex(3, 0, width),
      cellIndex(3, 1, width),
      cellIndex(3, 2, width),
      cellIndex(3, 3, width),
      cellIndex(3, 4, width),
      cellIndex(3, 5, width),
      cellIndex(3, 6, width),
      cellIndex(3, 7, width),
      cellIndex(3, 8, width),
      cellIndex(3, 9, width),
    ])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    const resultTags = kernel.readTags()
    const resultNumbers = kernel.readNumbers()
    outputCells.forEach((cell) => expect(resultTags[cell]).toBe(ValueTag.Number))
    expect(resultNumbers[cellIndex(3, 0, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 1, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 2, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 3, width)]).toBe(3)
    expect(resultNumbers[cellIndex(3, 4, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 5, width)]).toBe(2)
    expect(resultNumbers[cellIndex(3, 6, width)]).toBe(3)
    expect(resultNumbers[cellIndex(3, 7, width)]).toBe(1)
    expect(resultNumbers[cellIndex(3, 8, width)]).toBe(0)
    expect(resultNumbers[cellIndex(3, 9, width)]).toBe(14)
  })

  it('evaluates TREND and GROWTH spill helpers on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(18, 4, 0, 1, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        ValueTag.Number,
        0,
        0,
      ]),
      new Float64Array([5, 1, 4, 2, 0, 0, 8, 2, 5, 4, 0, 0, 11, 3, 0, 8, 0, 0]),
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
        cellIndex(0, 2, width),
        cellIndex(1, 2, width),
        cellIndex(0, 3, width),
        cellIndex(1, 3, width),
        cellIndex(2, 3, width),
      ]),
      Uint32Array.from([0, 3, 6, 8]),
      Uint32Array.from([3, 3, 2, 3]),
      Uint32Array.from([3, 3, 2, 3]),
      Uint32Array.from([1, 1, 1, 1]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 2, 3]), Uint32Array.from([1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.TREND, 3), encodeRet()],
      [encodePushRange(3), encodePushRange(1), encodePushRange(2), encodeCall(BUILTIN.GROWTH, 3), encodeRet()],
    ])
    const constants = packConstants([[], []])
    const outputCells = Uint32Array.from([cellIndex(0, 4, width), cellIndex(0, 5, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expect(kernel.readSpillRows()[cellIndex(0, 4, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(0, 4, width)]).toBe(1)
    expect(
      readSpillValues(kernel, cellIndex(0, 4, width), []).map((value) => (value.tag === ValueTag.Number ? value.value : value)),
    ).toEqual([14, 17])

    expect(kernel.readSpillRows()[cellIndex(0, 5, width)]).toBe(2)
    expect(kernel.readSpillCols()[cellIndex(0, 5, width)]).toBe(1)
    const growthValues = readSpillValues(kernel, cellIndex(0, 5, width), []).map((value) =>
      value.tag === ValueTag.Number ? value.value : value,
    )
    expect(growthValues[0]).toBeCloseTo(16, 12)
    expect(growthValues[1]).toBeCloseTo(32, 12)
  })
})
