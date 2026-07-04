import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { excelDateTimeFixtureSuite } from '../../../excel-fixtures/src/datetime-fixtures.js'
import {
  addMonthsToExcelDate,
  datetimeBuiltins,
  endOfMonthExcelDate,
  excelDatePartsToSerial,
  excelSerialToDateParts,
} from '../builtins/datetime.js'

describe('datetime builtins: calendar edge cases and fixture coverage', () => {
  it('supports EOMONTH end-of-month lookups', () => {
    expect(endOfMonthExcelDate(45337, 0)).toBe(45351)
    expect(endOfMonthExcelDate(45337, 1)).toBe(45382)
    expect(endOfMonthExcelDate(1, -1)).toBeUndefined()

    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45337 }, { tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 45382,
    })

    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45337 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('supports TIMEVALUE parsing and extended week-number variants', () => {
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:30 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:00 AM', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '1:02:03 pm', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: (13 * 3600 + 2 * 60 + 3) / 86_400,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '24:00', stringId: 1 })).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '13:00 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:60', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 13 / 24,
    })
    expect(datetimeBuiltins.TIMEVALUE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(
      datetimeBuiltins.ISOWEEKNUM({
        tag: ValueTag.Number,
        value: excelDatePartsToSerial(2026, 1, 1)!,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 12 }),
    ).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 16 }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 21 }),
    ).toEqual({ tag: ValueTag.Number, value: 11 })
  })

  it('handles reverse ranges, weekend starts, and invalid helper inputs', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const saturdaySerial = excelDatePartsToSerial(2026, 3, 14)!
    const fridayNextWeek = excelDatePartsToSerial(2026, 3, 20)!

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: saturdaySerial }, { tag: ValueTag.Number, value: 0 })).toEqual({
      tag: ValueTag.Number,
      value: saturdaySerial,
    })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2026, 3, 12)!,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: fridayNextWeek }, { tag: ValueTag.Number, value: fridaySerial }),
    ).toEqual({ tag: ValueTag.Number, value: -6 })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: fridayNextWeek },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(excelSerialToDateParts(Number.NaN)).toBeUndefined()
    expect(excelDatePartsToSerial(10_000, 1, 1)).toBeUndefined()
    expect(addMonthsToExcelDate(Number.NaN, 1)).toBeUndefined()
    expect(endOfMonthExcelDate(Number.NaN, 1)).toBeUndefined()
  })

  it('covers remaining datetime coercion and calendar validation branches', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const feb28 = excelDatePartsToSerial(2023, 2, 28)!
    const feb29 = excelDatePartsToSerial(2024, 2, 29)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(excelDatePartsToSerial(Number.NaN, 1, 1)).toBeUndefined()
    expect(excelDatePartsToSerial(-1, 1, 1)).toBeUndefined()
    expect(excelSerialToDateParts(Number.POSITIVE_INFINITY)).toBeUndefined()
    expect(addMonthsToExcelDate(jan31, Number.NaN)).toBeUndefined()
    expect(endOfMonthExcelDate(jan31, Number.NaN)).toBeUndefined()

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: Number.NaN },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 33_000 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Boolean, value: true }, { tag: ValueTag.Boolean, value: false }, { tag: ValueTag.Empty }),
    ).toEqual({
      tag: ValueTag.Number,
      value: 1 / 24,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Number, value: 0.5 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.Boolean, value: true })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: 'bad:00', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '-1:00', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '23:59:60', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '32768:00', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: Number.NaN }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Number,
      value: 60,
    })

    const leapSpan = datetimeBuiltins.YEARFRAC(
      { tag: ValueTag.Number, value: feb28 },
      { tag: ValueTag.Number, value: feb29 },
      { tag: ValueTag.Number, value: 1 },
    )
    expect(leapSpan.tag).toBe(ValueTag.Number)
    if (leapSpan.tag === ValueTag.Number) {
      expect(leapSpan.value).toBeCloseTo(366 / 365.5, 12)
    }
    const multiYearSpan = datetimeBuiltins.YEARFRAC(
      { tag: ValueTag.Number, value: excelDatePartsToSerial(2020, 1, 1)! },
      { tag: ValueTag.Number, value: excelDatePartsToSerial(2023, 7, 1)! },
      { tag: ValueTag.Number, value: 1 },
    )
    expect(multiYearSpan.tag).toBe(ValueTag.Number)
    if (multiYearSpan.tag === ValueTag.Number) {
      expect(multiYearSpan.value).toBeCloseTo(1277 / 365.25, 12)
    }

    expect(datetimeBuiltins.ISOWEEKNUM()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Number, value: Number.NaN })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: '000000', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers additional datetime edge validation paths', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.String, value: ' ', stringId: 2 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.HOUR()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.WEEKDAY()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.ISOWEEKNUM({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: Number.NaN })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEDIF({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
  })

  it('ships a focused datetime fixture suite for later aggregation', () => {
    expect(excelDateTimeFixtureSuite.id).toBe('datetime-serial-1900')
    expect(excelDateTimeFixtureSuite.sheets).toEqual([{ name: 'Sheet1' }])
    expect(excelDateTimeFixtureSuite.cases).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 'date-time:serial-addition',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 45299 } }],
        }),
        expect.objectContaining({
          id: 'date-time:date-empty-year-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:date-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:date-constructor-leap-day',
          formula: '=DATE(2024,2,29)',
        }),
        expect.objectContaining({
          id: 'date-time:year-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1900 } }],
        }),
        expect.objectContaining({
          id: 'date-time:month-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:day-leap-bug-serial',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 29 } }],
        }),
        expect.objectContaining({
          id: 'date-time:time-basic',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 0.5208333333333334 } }],
        }),
        expect.objectContaining({
          id: 'date-time:hour-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 12 } }],
        }),
        expect.objectContaining({
          id: 'date-time:minute-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 30 } }],
        }),
        expect.objectContaining({
          id: 'date-time:second-basic',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:weekday-basic',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 1 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-month-shift',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 45351 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 32 } }],
        }),
        expect.objectContaining({
          id: 'date-time:edate-text-error',
          outputs: [
            {
              address: 'A2',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'date-time:eomonth-boolean-coercion',
          outputs: [{ address: 'A2', expected: { kind: 'number', value: 60 } }],
        }),
        expect.objectContaining({
          id: 'date-time:eomonth-text-error',
          outputs: [
            {
              address: 'A3',
              expected: { kind: 'error', code: ErrorCode.Value, display: '#VALUE!' },
            },
          ],
        }),
        expect.objectContaining({
          id: 'volatile:today-captured-utc',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 46100 } }],
        }),
        expect.objectContaining({
          id: 'volatile:rand-captured',
          outputs: [{ address: 'A1', expected: { kind: 'number', value: 0.625 } }],
        }),
        expect.objectContaining({
          id: 'date-time:yearfrac-invalid-basis-num',
          outputs: [
            {
              address: 'A1',
              expected: { kind: 'error', code: ErrorCode.Num, display: '#NUM!' },
            },
          ],
        }),
      ]),
    )
  })

  it('covers yearFracByBasis complex branches and edge cases', () => {
    const d1 = excelDatePartsToSerial(2023, 1, 1)!
    const d2 = excelDatePartsToSerial(2024, 1, 1)!
    const d3 = excelDatePartsToSerial(2025, 1, 1)!

    // Basis 1: Actual/actual
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d1 },
        { tag: ValueTag.Number, value: d2 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 365 / 365 })

    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d2 },
        { tag: ValueTag.Number, value: d3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 366 / 366 })

    // Multi-year crossing
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: d1 },
        { tag: ValueTag.Number, value: d3 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2.0009124087591244 })
  })

  it('covers more TIMEVALUE parsing scenarios', () => {
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '  12:30 PM  ', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5208333333333334,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '12:00:00', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5,
    })
    expect(datetimeBuiltins.TIMEVALUE({ tag: ValueTag.String, value: '11:59:59 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: (23 * 3600 + 59 * 60 + 59) / 86400,
    })
  })

  it('covers date boundary limits and normalization', () => {
    expect(excelDatePartsToSerial(-1, 1, 1)).toBeUndefined()
    expect(excelDatePartsToSerial(10000, 1, 1)).toBeUndefined()
    expect(addMonthsToExcelDate(1, 120000)).toBeUndefined()
    expect(endOfMonthExcelDate(1, 120000)).toBeUndefined()
  })

  it('covers remaining datetime public validation and edge branches', () => {
    const jan31 = excelDatePartsToSerial(2024, 1, 31)!
    const feb29 = excelDatePartsToSerial(2024, 2, 29)!
    const mar31 = excelDatePartsToSerial(2024, 3, 31)!

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 2 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 10000 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.DATE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Empty })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Number, value: Number.POSITIVE_INFINITY })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: '   ', stringId: 3 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.String, value: 'bad', stringId: 4 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: Number.NEGATIVE_INFINITY })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Error, code: ErrorCode.NA },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.String, value: 'bad', stringId: 5 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.String, value: 'bad', stringId: 6 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: 32768 }, { tag: ValueTag.Number, value: 0 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins.TIME()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.HOUR()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: -0.25 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: -1 })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: jan31 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Number, value: 17 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Error, code: ErrorCode.Ref },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: mar31 }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref }),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 11 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 52,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: jan31 }, { tag: ValueTag.Error, code: ErrorCode.Ref })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.String, value: 'bad', stringId: 7 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.String, value: 'D', stringId: 8 },
      ),
    ).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: feb29 },
        { tag: ValueTag.String, value: 'YD', stringId: 9 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: jan31 },
        { tag: ValueTag.Number, value: feb29 },
        { tag: ValueTag.String, value: 'MD', stringId: 10 },
      ),
    ).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })
  })
})
