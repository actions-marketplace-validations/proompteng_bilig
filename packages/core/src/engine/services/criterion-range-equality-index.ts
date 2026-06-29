import { ValueTag } from '@bilig/protocol'
import { normalizeExactLookupNumber } from '@bilig/formula'
import { sliceAbsoluteRowsToRangeView, type CriterionRowSetView } from './criterion-rowset-view.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import type { SliceFastPredicate } from './criterion-range-predicate.js'
import { decodeValueTag, normalizeSliceString } from './criterion-range-values.js'

interface CriterionEqualityRowIndex {
  readonly numbers: ReadonlyMap<number, Uint32Array>
  readonly strings: ReadonlyMap<string, Uint32Array>
}

export type CriterionEqualityIndexKey =
  | {
      readonly kind: 'number'
      readonly value: number
    }
  | {
      readonly kind: 'string'
      readonly value: string
    }

const equalityRowIndexes = new WeakMap<RuntimeColumnView['owner'], CriterionEqualityRowIndex>()

export function equalityIndexKeyForPredicate(predicate: SliceFastPredicate): CriterionEqualityIndexKey | undefined {
  if (predicate.kind === 'eq-number' && !predicate.negate && predicate.value !== 0) {
    return { kind: 'number', value: predicate.value }
  }
  if (predicate.kind === 'eq-bool' && !predicate.negate && predicate.value) {
    return { kind: 'number', value: 1 }
  }
  if (predicate.kind === 'eq-string' && !predicate.negate && predicate.value !== '') {
    return { kind: 'string', value: predicate.value }
  }
  return undefined
}

export function equalityIndexKeyForStoredValue(
  tag: ValueTag,
  number: number,
  stringId: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionEqualityIndexKey | undefined {
  if (tag === ValueTag.Number) {
    const value = normalizeExactLookupNumber(number)
    return value === 0 ? undefined : { kind: 'number', value }
  }
  if (tag === ValueTag.Boolean) {
    return number === 0 ? undefined : { kind: 'number', value: 1 }
  }
  if (tag === ValueTag.String) {
    const value = normalizeSliceString(runtimeColumnStore, stringId)
    return value === '' ? undefined : { kind: 'string', value }
  }
  return undefined
}

function sortedUint32Rows(rows: number[]): Uint32Array {
  if (rows.length === 0) {
    return new Uint32Array(0)
  }
  rows.sort((left, right) => left - right)
  return Uint32Array.from(rows)
}

function getOrBuildEqualityRowIndex(
  view: RuntimeColumnView,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionEqualityRowIndex {
  const cached = equalityRowIndexes.get(view.owner)
  if (cached !== undefined) {
    return cached
  }

  const numbers = new Map<number, number[]>()
  const strings = new Map<string, number[]>()
  view.owner.pages.forEach((page) => {
    if (page.nonEmptyCount === 0) {
      return
    }
    const rowEnd = page.rowStart + page.tags.length - 1
    for (let row = page.rowStart; row <= rowEnd; row += 1) {
      const localRow = row - page.rowStart
      const tag = decodeValueTag(page.tags[localRow])
      const key = equalityIndexKeyForStoredValue(tag, page.numbers[localRow] ?? 0, page.stringIds[localRow] ?? 0, runtimeColumnStore)
      if (key?.kind === 'number') {
        let rows = numbers.get(key.value)
        if (rows === undefined) {
          rows = []
          numbers.set(key.value, rows)
        }
        rows.push(row)
        continue
      }
      if (key?.kind === 'string') {
        let rows = strings.get(key.value)
        if (rows === undefined) {
          rows = []
          strings.set(key.value, rows)
        }
        rows.push(row)
      }
    }
  })

  const index: CriterionEqualityRowIndex = {
    numbers: new Map([...numbers].map(([value, rows]) => [value, sortedUint32Rows(rows)])),
    strings: new Map([...strings].map(([value, rows]) => [value, sortedUint32Rows(rows)])),
  }
  equalityRowIndexes.set(view.owner, index)
  return index
}

export function readIndexedEqualityRows(
  view: RuntimeColumnView,
  predicate: SliceFastPredicate,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): CriterionRowSetView | undefined {
  const key = equalityIndexKeyForPredicate(predicate)
  if (key === undefined) {
    return undefined
  }
  const index = getOrBuildEqualityRowIndex(view, runtimeColumnStore)
  const rows = key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value)
  return sliceAbsoluteRowsToRangeView(rows ?? new Uint32Array(0), view.rowStart, view.rowEnd)
}
