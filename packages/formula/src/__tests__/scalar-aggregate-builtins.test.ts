import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createScalarAggregateBuiltins } from '../builtins/scalar-aggregate-builtins.js'
import { toDirectAggregateNumber, toNumber } from '../builtins/scalar-coercion.js'

const number = (value: number): CellValue => ({ tag: ValueTag.Number, value })
const text = (value: string): CellValue => ({ tag: ValueTag.String, value, stringId: 0 })
const bool = (value: boolean): CellValue => ({ tag: ValueTag.Boolean, value })
const empty = (): CellValue => ({ tag: ValueTag.Empty })
const error = (code: ErrorCode): CellValue => ({ tag: ValueTag.Error, code })

const builtins = createScalarAggregateBuiltins({
  toDirectAggregateNumber,
  toNumber,
  firstError: (args) => args.find((arg): arg is CellValue => arg?.tag === ValueTag.Error),
  numberResult: (value) => ({ tag: ValueTag.Number, value }),
  valueError: () => ({ tag: ValueTag.Error, code: ErrorCode.Value }),
  div0Error: () => ({ tag: ValueTag.Error, code: ErrorCode.Div0 }),
  numError: () => ({ tag: ValueTag.Error, code: ErrorCode.Num }),
})

describe('scalar aggregate builtins', () => {
  it('keeps direct aggregate coercion and error propagation together', () => {
    expect(builtins.SUM(number(2), text('3'), bool(true))).toEqual(number(6))
    expect(builtins.AVERAGE(number(2), text('4'))).toEqual(number(3))
    expect(builtins.MIN(number(2), text('4'))).toEqual(number(2))
    expect(builtins.COUNT(number(2), text('4'), text('bad'), empty())).toEqual(number(2))
    expect(builtins.SUM(number(2), text('bad'))).toEqual(error(ErrorCode.Value))
    expect(builtins.SUBTOTAL(number(9), number(2), error(ErrorCode.Ref))).toEqual(error(ErrorCode.Ref))
    expect(builtins.AGGREGATE(number(9), number(6), number(2), error(ErrorCode.Ref))).toEqual(number(2))
  })
})
