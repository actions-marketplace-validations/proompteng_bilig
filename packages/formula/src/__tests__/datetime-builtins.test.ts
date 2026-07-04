import { ErrorCode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import {
  addMonthsToExcelDate,
  createNowBuiltin,
  createRandBuiltin,
  createTodayBuiltin,
  datetimeBuiltins,
  excelDatePartsToSerial,
  excelSerialToDateParts,
  utcDateToExcelSerial,
} from '../builtins/datetime.js'

describe('datetime builtins: serials, workdays, and random date functions', () => {
  it('converts between Excel serials and date parts in the 1900 system', () => {
    expect(excelDatePartsToSerial(1900, 1, 1)).toBe(1)
    expect(excelDatePartsToSerial(1900, 2, 29)).toBe(60)
    expect(excelDatePartsToSerial(1900, 3, 1)).toBe(61)
    expect(excelDatePartsToSerial(2024, 2, 29)).toBe(45351)

    expect(excelSerialToDateParts(60)).toEqual({ year: 1900, month: 2, day: 29 })
    expect(excelSerialToDateParts(61)).toEqual({ year: 1900, month: 3, day: 1 })
    expect(excelSerialToDateParts(45351)).toEqual({ year: 2024, month: 2, day: 29 })
  })

  it('supports DATE with Excel-style year and month/day normalization', () => {
    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 2024 }, { tag: ValueTag.Number, value: 2 }, { tag: ValueTag.Number, value: 29 }),
    ).toEqual({ tag: ValueTag.Number, value: 45351 })

    expect(
      datetimeBuiltins.DATE({ tag: ValueTag.Number, value: 24 }, { tag: ValueTag.Number, value: 14 }, { tag: ValueTag.Number, value: 1 }),
    ).toEqual({ tag: ValueTag.Number, value: 9164 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.String, value: '2024', stringId: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 29 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 45351 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.String, value: '2024', stringId: 1 },
        { tag: ValueTag.String, value: '1', stringId: 2 },
        { tag: ValueTag.String, value: '2', stringId: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 45293 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.String, value: '', stringId: 4 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 45261 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Number, value: 2024 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: '', stringId: 5 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 45291 })

    expect(
      datetimeBuiltins.DATE(
        { tag: ValueTag.Error, code: ErrorCode.Ref },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 29 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref })
  })

  it('extracts YEAR, MONTH, and DAY from serial inputs including the leap-year bug date', () => {
    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Number, value: 45351 })).toEqual({
      tag: ValueTag.Number,
      value: 2024,
    })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Number, value: 45351.75 })).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(datetimeBuiltins.DAY({ tag: ValueTag.Number, value: 60 })).toEqual({
      tag: ValueTag.Number,
      value: 29,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.String, value: '45351', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.YEAR({ tag: ValueTag.Empty })).toEqual({
      tag: ValueTag.Number,
      value: 1899,
    })
    expect(datetimeBuiltins.MONTH({ tag: ValueTag.Boolean, value: true })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
  })

  it('treats recognized date text as DATEVALUE dates for serial date consumers', () => {
    const jan1Text = { tag: ValueTag.String, value: '1/1/31', stringId: 1 } as const
    const feb2Text = { tag: ValueTag.String, value: '2/2/31', stringId: 2 } as const
    const jan1Serial = excelDatePartsToSerial(1931, 1, 1)!
    const feb1Serial = excelDatePartsToSerial(1931, 2, 1)!
    const feb2Serial = excelDatePartsToSerial(1931, 2, 2)!
    const febEndSerial = excelDatePartsToSerial(1931, 2, 28)!

    expect(datetimeBuiltins.YEAR(jan1Text)).toEqual({ tag: ValueTag.Number, value: 1931 })
    expect(datetimeBuiltins.MONTH(jan1Text)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.DAY(jan1Text)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.WEEKDAY(jan1Text)).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(datetimeBuiltins.WEEKDAY(jan1Text, { tag: ValueTag.Number, value: 2 })).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(datetimeBuiltins.WEEKNUM(jan1Text)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.WEEKNUM(jan1Text, { tag: ValueTag.Number, value: 2 })).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.ISOWEEKNUM(jan1Text)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(datetimeBuiltins.EDATE(jan1Text, { tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Number, value: feb1Serial })
    expect(datetimeBuiltins.EOMONTH(jan1Text, { tag: ValueTag.Number, value: 1 })).toEqual({ tag: ValueTag.Number, value: febEndSerial })
    expect(datetimeBuiltins.DAYS360(jan1Text, feb2Text)).toEqual({ tag: ValueTag.Number, value: 31 })
    expect(datetimeBuiltins.YEARFRAC(jan1Text, feb2Text)).toEqual({ tag: ValueTag.Number, value: 31 / 360 })

    expect(datetimeBuiltins.DAYS(feb2Text, jan1Text)).toEqual({ tag: ValueTag.Number, value: feb2Serial - jan1Serial })
  })

  it('supports TIME plus HOUR, MINUTE, SECOND, and WEEKDAY extraction', () => {
    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!

    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: 12 }, { tag: ValueTag.Number, value: 30 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Number, value: 0.5208333333333334 })

    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.String, value: '1', stringId: 1 },
        { tag: ValueTag.String, value: '2', stringId: 2 },
        { tag: ValueTag.String, value: '3', stringId: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3723 / 86400 })

    expect(
      datetimeBuiltins.TIME(
        { tag: ValueTag.String, value: '', stringId: 4 },
        { tag: ValueTag.Number, value: 0 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0 })

    expect(datetimeBuiltins.HOUR({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.Number, value: 0.5208333333333334 })).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: 0.5208449074074074 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.HOUR({ tag: ValueTag.String, value: '6:45 PM', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 18,
    })
    expect(datetimeBuiltins.MINUTE({ tag: ValueTag.String, value: '6:45 PM', stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 45,
    })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.String, value: '4:48:18 PM', stringId: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 18,
    })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.Number, value: 0.50001 })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.SECOND({ tag: ValueTag.String, value: '12:00:59.6', stringId: 4 })).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial })).toEqual({
      tag: ValueTag.Number,
      value: 1,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
    expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('returns Excel errors for unsupported time-part coercions and weekday return types', () => {
    expect(
      datetimeBuiltins.TIME({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 30 }, { tag: ValueTag.Number, value: 0 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(datetimeBuiltins.HOUR({ tag: ValueTag.String, value: 'not-a-time', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 99 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })

    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!
    const weekdayTypes = [1, 2, 11, 12, 13, 14, 15, 16, 17]
    const expectedForSunday = [1, 7, 7, 6, 5, 4, 3, 2, 1]
    weekdayTypes.forEach((type, i) => {
      expect(datetimeBuiltins.WEEKDAY({ tag: ValueTag.Number, value: sundaySerial }, { tag: ValueTag.Number, value: type })).toEqual({
        tag: ValueTag.Number,
        value: expectedForSunday[i],
      })
    })
  })

  it('supports DAYS, WEEKNUM, WORKDAY, and NETWORKDAYS', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const mondayHoliday = excelDatePartsToSerial(2026, 3, 16)!
    const fridayNextWeek = excelDatePartsToSerial(2026, 3, 20)!
    const jan1Holiday = excelDatePartsToSerial(2026, 1, 1)!
    const jan2Holiday = excelDatePartsToSerial(2026, 1, 2)!
    const jan7Serial = excelDatePartsToSerial(2026, 1, 7)!

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: fridayNextWeek }, { tag: ValueTag.Number, value: fridaySerial })).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })

    expect(
      datetimeBuiltins.WEEKNUM({
        tag: ValueTag.Number,
        value: excelDatePartsToSerial(2026, 3, 15)!,
      }),
    ).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: excelDatePartsToSerial(2026, 3, 15)! }, { tag: ValueTag.Number, value: 2 }),
    ).toEqual({ tag: ValueTag.Number, value: 11 })

    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: mondayHoliday,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: mondayHoliday },
      ),
    ).toEqual({ tag: ValueTag.Number, value: mondayHoliday + 1 })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: jan1Holiday },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: jan1Holiday },
        { tag: ValueTag.Number, value: jan2Holiday },
      ),
    ).toEqual({ tag: ValueTag.Number, value: jan7Serial })

    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: fridayNextWeek }),
    ).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: fridayNextWeek },
        { tag: ValueTag.Number, value: mondayHoliday },
        { tag: ValueTag.Number, value: fridayNextWeek },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 4 })
  })

  it('supports WORKDAY.INTL and NETWORKDAYS.INTL weekend masks', () => {
    const fridaySerial = excelDatePartsToSerial(2026, 3, 13)!
    const sundaySerial = excelDatePartsToSerial(2026, 3, 15)!
    const mondaySerial = excelDatePartsToSerial(2026, 3, 16)!
    const tuesdaySerial = excelDatePartsToSerial(2026, 3, 17)!
    const jan1Holiday = excelDatePartsToSerial(2026, 1, 1)!
    const jan2Holiday = excelDatePartsToSerial(2026, 1, 2)!
    const jan7Serial = excelDatePartsToSerial(2026, 1, 7)!

    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Number,
      value: mondaySerial,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: sundaySerial })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: '0000011', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: mondaySerial })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: sundaySerial },
      ),
    ).toEqual({ tag: ValueTag.Number, value: tuesdaySerial })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: jan1Holiday },
        { tag: ValueTag.Number, value: 3 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: jan1Holiday },
        { tag: ValueTag.Number, value: jan2Holiday },
      ),
    ).toEqual({ tag: ValueTag.Number, value: jan7Serial })

    expect(
      datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: fridaySerial }, { tag: ValueTag.Number, value: tuesdaySerial }),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.Number, value: 7 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.String, value: '1000001', stringId: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: sundaySerial },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 7 },
        { tag: ValueTag.Number, value: sundaySerial },
      ),
    ).toEqual({ tag: ValueTag.Number, value: -2 })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.String, value: ' 0000110 ', stringId: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 3 })

    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 99 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: fridaySerial },
        { tag: ValueTag.Number, value: tuesdaySerial },
        { tag: ValueTag.String, value: '1111111', stringId: 4 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('rejects out-of-range workday date serials with documented error classes', () => {
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins['WORKDAY.INTL']({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins['WORKDAY.INTL'](
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })

    expect(datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: -1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(datetimeBuiltins['NETWORKDAYS.INTL']({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins['NETWORKDAYS.INTL'](
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: 2 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
  })

  it('supports DATEDIF units', () => {
    const start = excelDatePartsToSerial(2020, 1, 15)!
    const end = excelDatePartsToSerial(2021, 3, 20)!

    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: end - start })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'M', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'Y', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'YM', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'YD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 64 })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'MD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 5 })

    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.String, value: 'D', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins.DATEDIF(
        { tag: ValueTag.Number, value: start },
        { tag: ValueTag.Number, value: end },
        { tag: ValueTag.String, value: 'BAD', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })

  it('creates deterministic TODAY and NOW builtins from injected UTC dates', () => {
    const fixedNow = new Date('2026-03-19T15:45:30.000Z')
    const TODAY = createTodayBuiltin(() => fixedNow)
    const NOW = createNowBuiltin(() => fixedNow)

    expect(TODAY()).toEqual({ tag: ValueTag.Number, value: 46100 })
    expect(NOW()).toEqual({ tag: ValueTag.Number, value: 46100.65659722222 })
    expect(utcDateToExcelSerial(fixedNow)).toBe(46100.65659722222)

    expect(TODAY({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(NOW({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('supports RAND with Excel-style numeric bounds and injectable randomness', () => {
    const RAND = createRandBuiltin(() => 0.625)
    const highRAND = createRandBuiltin(() => 2)
    const lowRAND = createRandBuiltin(() => -0.5)
    const invalidRAND = createRandBuiltin(() => Number.NaN)

    expect(RAND()).toEqual({ tag: ValueTag.Number, value: 0.625 })
    expect(highRAND()).toEqual({ tag: ValueTag.Number, value: 1 - Number.EPSILON })
    expect(lowRAND()).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(invalidRAND()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.RAND({ tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('returns explicit errors for WORKDAY, NETWORKDAYS, TODAY, NOW, RAND, EDATE, and EOMONTH edge inputs', () => {
    expect(datetimeBuiltins.WORKDAY()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.WORKDAY({ tag: ValueTag.Number, value: 46094 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WORKDAY(
        { tag: ValueTag.Number, value: 46094 },
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.NETWORKDAYS()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Error, code: ErrorCode.Ref }, { tag: ValueTag.Number, value: 46095 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Ref,
    })
    expect(
      datetimeBuiltins.NETWORKDAYS({ tag: ValueTag.Number, value: 46094 }, { tag: ValueTag.String, value: 'bad', stringId: 1 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(
      datetimeBuiltins.NETWORKDAYS(
        { tag: ValueTag.Number, value: 46094 },
        { tag: ValueTag.Number, value: 46095 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(
      createTodayBuiltin(() => new Date('2026-03-19T00:00:00.000Z'))({
        tag: ValueTag.Error,
        code: ErrorCode.Name,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
    expect(
      createNowBuiltin(() => new Date('2026-03-19T00:00:00.000Z'))({
        tag: ValueTag.Number,
        value: 1,
      }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(createRandBuiltin(() => 0.5)({ tag: ValueTag.Error, code: ErrorCode.NA })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.EOMONTH({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.String, value: 'bad', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('covers DAYS and WEEKNUM validation and alternate return-type branches', () => {
    const sampleDate = excelDatePartsToSerial(2026, 3, 15)!

    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: 10 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DAYS({ tag: ValueTag.Number, value: 10 }, { tag: ValueTag.String, value: 'bad', stringId: 2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })

    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 12 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 13 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 14 })).toEqual({
      tag: ValueTag.Number,
      value: 11,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 15 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.Number, value: 16 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.String, value: 'bad', stringId: 3 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(
      datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: sampleDate }, { tag: ValueTag.String, value: 'bad', stringId: 4 }),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.WEEKNUM({ tag: ValueTag.Number, value: Number.NaN }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports EDATE month shifting with end-of-month clamping', () => {
    expect(addMonthsToExcelDate(45322, 1)).toBe(45351)
    expect(addMonthsToExcelDate(45351, -1)).toBe(45320)
    expect(addMonthsToExcelDate(1, -1)).toBeUndefined()

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: 45322 }, { tag: ValueTag.Number, value: 1.9 })).toEqual({
      tag: ValueTag.Number,
      value: 45351,
    })

    expect(datetimeBuiltins.EDATE({ tag: ValueTag.String, value: 'bad', stringId: 1 }, { tag: ValueTag.Number, value: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.EDATE({ tag: ValueTag.Number, value: 1 }, { tag: ValueTag.Number, value: -1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Num,
    })
  })

  it('supports DATEVALUE for text dates and rejects numeric serials', () => {
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Number, value: 1.2 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: '2024-02-29', stringId: 1 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2024, 2, 29)!,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: '1-FEB-2021 11:59 PM', stringId: 2 })).toEqual({
      tag: ValueTag.Number,
      value: excelDatePartsToSerial(2021, 2, 1)!,
    })
    expect(datetimeBuiltins.DATEVALUE()).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.Error, code: ErrorCode.Name })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Name,
    })
    expect(datetimeBuiltins.DATEVALUE({ tag: ValueTag.String, value: 'not-a-date', stringId: 1 })).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    })
  })

  it('supports DAYS360 and YEARFRAC across basis modes', () => {
    const jan1 = excelDatePartsToSerial(2024, 1, 1)!
    const jul1 = excelDatePartsToSerial(2024, 7, 1)!
    const feb28 = excelDatePartsToSerial(2023, 2, 28)!
    const mar31 = excelDatePartsToSerial(2023, 3, 31)!

    expect(datetimeBuiltins.DAYS360({ tag: ValueTag.Number, value: feb28 }, { tag: ValueTag.Number, value: mar31 })).toEqual({
      tag: ValueTag.Number,
      value: 31,
    })
    expect(
      datetimeBuiltins.DAYS360(
        { tag: ValueTag.Number, value: feb28 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Boolean, value: true },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 32 })
    expect(
      datetimeBuiltins.DAYS360(
        { tag: ValueTag.Number, value: feb28 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })

    expect(datetimeBuiltins.YEARFRAC({ tag: ValueTag.Number, value: jan1 }, { tag: ValueTag.Number, value: jul1 })).toEqual({
      tag: ValueTag.Number,
      value: 0.5,
    })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 366 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 2 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 360 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: feb28 },
        { tag: ValueTag.Number, value: mar31 },
        { tag: ValueTag.Number, value: 0 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 31 / 360 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 3 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 365 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 4 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 0.5 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: 1 },
      ),
    ).toEqual({ tag: ValueTag.Number, value: 182 / 366 })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: 9 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.Number, value: -1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Num })
    expect(
      datetimeBuiltins.YEARFRAC(
        { tag: ValueTag.Number, value: jan1 },
        { tag: ValueTag.Number, value: jul1 },
        { tag: ValueTag.String, value: 'bad', stringId: 1 },
      ),
    ).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
  })
})
