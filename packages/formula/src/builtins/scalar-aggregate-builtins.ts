import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EvaluationResult } from '../runtime-values.js'
import { collectNumericArgs } from './numeric.js'
import { populationVariance, sampleVariance } from './statistics.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

export interface ScalarAggregateBuiltinHelpers {
  readonly toDirectAggregateNumber: (value: CellValue) => number | undefined
  readonly toNumber: (value: CellValue) => number | undefined
  readonly firstError: (args: readonly (CellValue | undefined)[]) => CellValue | undefined
  readonly numberResult: (value: number) => CellValue
  readonly valueError: () => CellValue
  readonly div0Error: () => CellValue
  readonly numError: () => CellValue
}

function aggregateOptionIgnoresErrors(option: number): boolean {
  return option === 2 || option === 3 || option === 6 || option === 7
}

export function createScalarAggregateBuiltins(helpers: ScalarAggregateBuiltinHelpers): Record<string, Builtin> {
  const { toDirectAggregateNumber, toNumber, firstError, numberResult, valueError, div0Error, numError } = helpers

  const toZeroNumericValue = (value: CellValue): number | undefined => {
    if (value.tag === ValueTag.String) {
      return 0
    }
    return toNumber(value)
  }

  const collectDirectAggregateNumbers = (args: readonly CellValue[]): number[] | CellValue => {
    const values: number[] = []
    for (const arg of args) {
      const numeric = toDirectAggregateNumber(arg)
      if (numeric === undefined) {
        return valueError()
      }
      values.push(numeric)
    }
    return values
  }

  const aggregateByCode = (functionNum: number, values: CellValue[], options: { readonly propagateErrors?: boolean } = {}): CellValue => {
    if (options.propagateErrors) {
      const error = firstError(values)
      if (error) {
        return error
      }
    }
    const normalized = functionNum > 100 ? functionNum - 100 : functionNum
    const numericValues = collectNumericArgs(values, toNumber)
    switch (normalized) {
      case 1:
        return numericValues.length === 0
          ? div0Error()
          : numberResult(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length)
      case 2:
        return numberResult(values.filter((value) => value.tag === ValueTag.Number || value.tag === ValueTag.Boolean).length)
      case 3:
        return numberResult(values.filter((value) => value.tag !== ValueTag.Empty).length)
      case 4:
        return numberResult(numericValues.length === 0 ? 0 : Math.max(...numericValues))
      case 5:
        return numberResult(numericValues.length === 0 ? 0 : Math.min(...numericValues))
      case 6:
        return numberResult(numericValues.length === 0 ? 0 : numericValues.reduce((product, value) => product * value, 1))
      case 7:
        return numberResult(Math.sqrt(sampleVariance(numericValues)))
      case 8:
        return numberResult(Math.sqrt(populationVariance(numericValues)))
      case 9:
        return numberResult(numericValues.reduce((sum, value) => sum + value, 0))
      case 10:
        return numberResult(sampleVariance(numericValues))
      case 11:
        return numberResult(populationVariance(numericValues))
      default:
        return valueError()
    }
  }

  const positiveDirectNumbers = (args: readonly CellValue[]): number[] | CellValue => {
    const error = firstError(args)
    if (error) {
      return error
    }
    const numbers = collectDirectAggregateNumbers(args)
    if (!Array.isArray(numbers)) {
      return numbers
    }
    if (numbers.length === 0) {
      return valueError()
    }
    if (numbers.some((value) => value <= 0)) {
      return numError()
    }
    return numbers
  }

  return {
    SUM: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = collectDirectAggregateNumbers(args)
      if (!Array.isArray(values)) return values
      return numberResult(values.reduce((sum, value) => sum + value, 0))
    },
    AVERAGEA: (...args) => {
      const error = firstError(args)
      if (error) return error
      const numbers = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
      return numbers.length === 0 ? div0Error() : numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
    },
    AVERAGE: (...args) => {
      const error = firstError(args)
      if (error) return error
      const numbers = collectDirectAggregateNumbers(args)
      if (!Array.isArray(numbers)) return numbers
      if (numbers.length === 0) return div0Error()
      return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
    },
    AVG: (...args) => {
      const error = firstError(args)
      if (error) return error
      const numbers = collectDirectAggregateNumbers(args)
      if (!Array.isArray(numbers)) return numbers
      if (numbers.length === 0) return div0Error()
      return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length)
    },
    MIN: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = collectDirectAggregateNumbers(args)
      if (!Array.isArray(values)) return values
      return values.length === 0 ? numberResult(0) : numberResult(Math.min(...values))
    },
    MAX: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = collectDirectAggregateNumbers(args)
      if (!Array.isArray(values)) return values
      return values.length === 0 ? numberResult(0) : numberResult(Math.max(...values))
    },
    MAXA: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
      return values.length === 0 ? numberResult(0) : numberResult(Math.max(...values))
    },
    MINA: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = args.map((arg) => toZeroNumericValue(arg)).filter((value): value is number => value !== undefined)
      return values.length === 0 ? numberResult(0) : numberResult(Math.min(...values))
    },
    COUNT: (...args) =>
      numberResult(
        args.filter(
          (value) =>
            value.tag === ValueTag.Number ||
            value.tag === ValueTag.Boolean ||
            (value.tag === ValueTag.String && toDirectAggregateNumber(value) !== undefined),
        ).length,
      ),
    COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
    COUNTBLANK: (...args) => {
      let blanks = 0
      for (const arg of args) {
        if (arg.tag === ValueTag.Empty || (arg.tag === ValueTag.String && arg.value === '')) {
          blanks += 1
        }
      }
      return numberResult(blanks)
    },
    GEOMEAN: (...args) => {
      const numbers = positiveDirectNumbers(args)
      if (!Array.isArray(numbers)) {
        return numbers
      }
      const logSum = numbers.reduce((sum, value) => sum + Math.log(value), 0)
      return numberResult(Math.exp(logSum / numbers.length))
    },
    HARMEAN: (...args) => {
      const numbers = positiveDirectNumbers(args)
      if (!Array.isArray(numbers)) {
        return numbers
      }
      return numberResult(numbers.length / numbers.reduce((sum, value) => sum + 1 / value, 0))
    },
    SUMSQ: (...args) => {
      const error = firstError(args)
      if (error) return error
      const values = collectDirectAggregateNumbers(args)
      if (!Array.isArray(values)) return values
      return numberResult(values.reduce((sum, value) => sum + value ** 2, 0))
    },
    SUBTOTAL: (functionNumArg, ...args) => {
      const functionNum = integerValue(functionNumArg, helpers.toNumber)
      return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args, { propagateErrors: true })
    },
    AGGREGATE: (functionNumArg, optionsArg, ...args) => {
      const functionNum = integerValue(functionNumArg, helpers.toNumber)
      const options = integerValue(optionsArg, helpers.toNumber)
      if (functionNum === undefined || options === undefined || options < 0 || options > 7) {
        return valueError()
      }
      const ignoreErrors = aggregateOptionIgnoresErrors(options)
      const values = ignoreErrors ? args.filter((value) => value.tag !== ValueTag.Error) : args
      return aggregateByCode(functionNum, values, { propagateErrors: !ignoreErrors })
    },
  }
}

function integerValue(value: CellValue | undefined, toNumber: (value: CellValue) => number | undefined): number | undefined {
  if (value === undefined) {
    return undefined
  }
  const numeric = toNumber(value)
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined
  }
  return Math.trunc(numeric)
}
