export { utcDateToExcelSerial } from '@bilig/formula'
export { ErrorCode, FormulaMode, Opcode, ValueTag, type EngineEvent } from '@bilig/protocol'
export type { EngineOpBatch } from '@bilig/workbook'
export { afterEach, describe, expect, it, vi } from 'vitest'
export { SpreadsheetEngine, type EngineSyncClient } from '../index.js'

export type RuntimeFormulaWithDependencies = {
  dependencyIndices: Uint32Array
}

export type RuntimeFormulaWithCompiled = {
  formulaSlotId: number
  planId: number
  compiled: {
    mode: number
    deps: string[]
    jsPlan: unknown[]
  }
  plan: {
    id: number
    source: string
    compiled: {
      mode: number
      deps: string[]
      jsPlan: unknown[]
    }
  }
  dependencyEntities: { ptr: number; len: number }
  runtimeProgram: Uint32Array
}

export type RuntimeFormulaWithRanges = {
  rangeDependencies: Uint32Array
  runtimeProgram: Uint32Array
}

export type RuntimeFormulaWithDirectLookup = {
  directLookup:
    | {
        kind: 'exact'
        operandCellIndex: number
        prepared: { sheetName: string; rowStart: number; rowEnd: number; col: number }
        searchMode: 1 | -1
      }
    | {
        kind: 'exact-uniform-numeric'
        operandCellIndex: number
        sheetName: string
        sheetId: number
        rowStart: number
        rowEnd: number
        col: number
        length: number
        searchMode: 1 | -1
      }
    | {
        kind: 'approximate'
        operandCellIndex: number
        prepared: { sheetName: string; rowStart: number; rowEnd: number; col: number }
        matchMode: 1 | -1
      }
    | {
        kind: 'approximate-uniform-numeric'
        operandCellIndex: number
        sheetName: string
        sheetId: number
        rowStart: number
        rowEnd: number
        col: number
        length: number
        matchMode: 1 | -1
      }
}

export type RuntimeFormulaWithDirectCriteria = {
  directCriteria: {
    aggregateKind: 'count' | 'sum' | 'average' | 'min' | 'max'
    aggregateRange:
      | {
          sheetName: string
          rowStart: number
          rowEnd: number
          col: number
          length: number
        }
      | undefined
    criteriaPairs: Array<{
      range: {
        sheetName: string
        rowStart: number
        rowEnd: number
        col: number
        length: number
      }
      criterion:
        | {
            kind: 'literal'
            value: unknown
          }
        | {
            kind: 'cell'
            cellIndex: number
          }
    }>
  }
}

export type RuntimeFormulaWithDirectAggregate = {
  directAggregate: {
    aggregateKind: 'sum' | 'average' | 'count' | 'min' | 'max'
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
    length: number
  }
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function hasFormulaStore(value: unknown): value is { formulas: { get(cellIndex: number): unknown } } {
  return (
    isRecord(value) &&
    'formulas' in value &&
    isRecord(value.formulas) &&
    'get' in value.formulas &&
    typeof value.formulas.get === 'function'
  )
}

export function readRuntimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  if (!hasFormulaStore(engine)) {
    throw new Error('SpreadsheetEngine test expected an internal formulas store')
  }
  return engine.formulas.get(cellIndex)
}

export function readRuntimeTemplateId(engine: SpreadsheetEngine, cellIndex: number): number | undefined {
  const runtimeFormula = readRuntimeFormula(engine, cellIndex)
  if (!isRecord(runtimeFormula)) {
    return undefined
  }
  return typeof runtimeFormula.templateId === 'number' ? runtimeFormula.templateId : undefined
}

export function readRuntimeDirectScalar(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const runtimeFormula = readRuntimeFormula(engine, cellIndex)
  return isRecord(runtimeFormula) ? runtimeFormula.directScalar : undefined
}

export function isRuntimeFormulaWithDependencies(value: unknown): value is RuntimeFormulaWithDependencies {
  return isRecord(value) && value.dependencyIndices instanceof Uint32Array
}

export function isRuntimeFormulaWithCompiled(value: unknown): value is RuntimeFormulaWithCompiled {
  return (
    isRecord(value) &&
    typeof value.formulaSlotId === 'number' &&
    typeof value.planId === 'number' &&
    isRecord(value.compiled) &&
    typeof value.compiled.mode === 'number' &&
    Array.isArray(value.compiled.deps) &&
    Array.isArray(value.compiled.jsPlan) &&
    isRecord(value.plan) &&
    typeof value.plan.id === 'number' &&
    typeof value.plan.source === 'string' &&
    isRecord(value.plan.compiled) &&
    typeof value.plan.compiled.mode === 'number' &&
    Array.isArray(value.plan.compiled.deps) &&
    Array.isArray(value.plan.compiled.jsPlan) &&
    isRecord(value.dependencyEntities) &&
    typeof value.dependencyEntities.ptr === 'number' &&
    typeof value.dependencyEntities.len === 'number' &&
    value.runtimeProgram instanceof Uint32Array
  )
}

export function isRuntimeFormulaWithRanges(value: unknown): value is RuntimeFormulaWithRanges {
  return isRecord(value) && value.rangeDependencies instanceof Uint32Array && value.runtimeProgram instanceof Uint32Array
}

export function isRuntimeFormulaWithDirectLookup(value: unknown): value is RuntimeFormulaWithDirectLookup {
  if (!isRecord(value) || !('directLookup' in value) || !isRecord(value.directLookup)) {
    return false
  }
  const directLookup = value.directLookup
  const prepared = directLookup.prepared
  if (typeof directLookup.operandCellIndex !== 'number') {
    return false
  }
  if (directLookup.kind === 'exact') {
    return (
      isRecord(prepared) &&
      typeof prepared.sheetName === 'string' &&
      typeof prepared.rowStart === 'number' &&
      typeof prepared.rowEnd === 'number' &&
      typeof prepared.col === 'number' &&
      (directLookup.searchMode === 1 || directLookup.searchMode === -1)
    )
  }
  if (directLookup.kind === 'exact-uniform-numeric') {
    return (
      typeof directLookup.sheetName === 'string' &&
      typeof directLookup.sheetId === 'number' &&
      typeof directLookup.rowStart === 'number' &&
      typeof directLookup.rowEnd === 'number' &&
      typeof directLookup.col === 'number' &&
      typeof directLookup.length === 'number' &&
      (directLookup.searchMode === 1 || directLookup.searchMode === -1)
    )
  }
  if (directLookup.kind === 'approximate') {
    return (
      isRecord(prepared) &&
      typeof prepared.sheetName === 'string' &&
      typeof prepared.rowStart === 'number' &&
      typeof prepared.rowEnd === 'number' &&
      typeof prepared.col === 'number' &&
      (directLookup.matchMode === 1 || directLookup.matchMode === -1)
    )
  }
  if (directLookup.kind === 'approximate-uniform-numeric') {
    return (
      typeof directLookup.sheetName === 'string' &&
      typeof directLookup.sheetId === 'number' &&
      typeof directLookup.rowStart === 'number' &&
      typeof directLookup.rowEnd === 'number' &&
      typeof directLookup.col === 'number' &&
      typeof directLookup.length === 'number' &&
      (directLookup.matchMode === 1 || directLookup.matchMode === -1)
    )
  }
  return false
}

export function isRuntimeFormulaWithDirectCriteria(value: unknown): value is RuntimeFormulaWithDirectCriteria {
  if (!isRecord(value) || !('directCriteria' in value) || !isRecord(value.directCriteria)) {
    return false
  }
  const directCriteria = value.directCriteria
  return (
    (directCriteria.aggregateKind === 'count' ||
      directCriteria.aggregateKind === 'sum' ||
      directCriteria.aggregateKind === 'average' ||
      directCriteria.aggregateKind === 'min' ||
      directCriteria.aggregateKind === 'max') &&
    Array.isArray(directCriteria.criteriaPairs)
  )
}

export function isRuntimeFormulaWithDirectAggregate(value: unknown): value is RuntimeFormulaWithDirectAggregate {
  if (!isRecord(value) || !('directAggregate' in value) || !isRecord(value.directAggregate)) {
    return false
  }
  const directAggregate = value.directAggregate
  return (
    (directAggregate.aggregateKind === 'sum' ||
      directAggregate.aggregateKind === 'average' ||
      directAggregate.aggregateKind === 'count' ||
      directAggregate.aggregateKind === 'min' ||
      directAggregate.aggregateKind === 'max') &&
    typeof directAggregate.sheetName === 'string' &&
    typeof directAggregate.rowStart === 'number' &&
    typeof directAggregate.rowEnd === 'number' &&
    typeof directAggregate.col === 'number' &&
    typeof directAggregate.length === 'number'
  )
}

export function seedPivotSource(engine: SpreadsheetEngine): void {
  engine.createSheet('Data')
  engine.createSheet('Pivot')
  engine.setRangeValues({ sheetName: 'Data', startAddress: 'A1', endAddress: 'B3' }, [
    ['Region', 'Sales'],
    ['East', 10],
    ['West', 7],
  ])
}
