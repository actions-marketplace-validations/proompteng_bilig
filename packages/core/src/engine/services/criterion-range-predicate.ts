import { ValueTag } from '@bilig/protocol'
import { matchesCompiledCriteria, normalizeExactLookupNumber, type CompiledCriteriaMatcher, type CriteriaOperator } from '@bilig/formula'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import { decodeValueTag, materializeSliceValue, normalizeSliceString } from './criterion-range-values.js'

export type SliceFastPredicate =
  | {
      kind: 'eq-empty'
      negate: boolean
    }
  | {
      kind: 'eq-bool'
      negate: boolean
      value: boolean
    }
  | {
      kind: 'eq-number'
      negate: boolean
      value: number
    }
  | {
      kind: 'eq-string'
      negate: boolean
      value: string
    }
  | {
      kind: 'cmp-number'
      operator: Exclude<CriteriaOperator, '=' | '<>'>
      value: number
    }
  | {
      kind: 'generic'
      compiled: CompiledCriteriaMatcher
    }

export function buildSlicePredicate(compiled: CompiledCriteriaMatcher): SliceFastPredicate {
  const { operator, operand, wildcardPattern } = compiled
  if (wildcardPattern) {
    return { kind: 'generic', compiled }
  }
  if (operator === '=' || operator === '<>') {
    const negate = operator === '<>'
    switch (operand.tag) {
      case ValueTag.Empty:
        return { kind: 'eq-empty', negate }
      case ValueTag.Boolean:
        return { kind: 'eq-bool', negate, value: operand.value }
      case ValueTag.Number:
        return {
          kind: 'eq-number',
          negate,
          value: normalizeExactLookupNumber(operand.value),
        }
      case ValueTag.String:
        return { kind: 'eq-string', negate, value: operand.value.toUpperCase() }
      case ValueTag.Error:
        return { kind: 'generic', compiled }
    }
  }
  if (operand.tag === ValueTag.Number) {
    return {
      kind: 'cmp-number',
      operator,
      value: normalizeExactLookupNumber(operand.value),
    }
  }
  return { kind: 'generic', compiled }
}

export function slicePredicateMatches(
  predicate: SliceFastPredicate,
  view: RuntimeColumnView,
  offset: number,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): boolean {
  switch (predicate.kind) {
    case 'eq-empty': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const matches =
        tag === ValueTag.Empty || (tag === ValueTag.String && normalizeSliceString(runtimeColumnStore, view.readStringIdAt(offset)) === '')
      return predicate.negate ? !matches : matches
    }
    case 'eq-bool': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const numeric = tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty ? view.readNumberAt(offset) : undefined
      const matches = numeric !== undefined && (Object.is(numeric, -0) ? 0 : numeric) === (predicate.value ? 1 : 0)
      return predicate.negate ? !matches : matches
    }
    case 'eq-number': {
      const tag = decodeValueTag(view.readTagAt(offset))
      const numeric = normalizeExactLookupNumber(view.readNumberAt(offset))
      const matches = (tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.Empty) && numeric === predicate.value
      return predicate.negate ? !matches : matches
    }
    case 'eq-string': {
      const tag = decodeValueTag(view.readTagAt(offset))
      if (predicate.negate && predicate.value === '') {
        return tag !== ValueTag.Empty && tag !== ValueTag.Error
      }
      const matches =
        (tag === ValueTag.String || tag === ValueTag.Empty) &&
        (tag === ValueTag.Empty ? '' : normalizeSliceString(runtimeColumnStore, view.readStringIdAt(offset))) === predicate.value
      return predicate.negate ? !matches : matches
    }
    case 'cmp-number': {
      const tag = decodeValueTag(view.readTagAt(offset))
      if (tag !== ValueTag.Number && tag !== ValueTag.Boolean) {
        return false
      }
      const numeric = normalizeExactLookupNumber(view.readNumberAt(offset))
      switch (predicate.operator) {
        case '>':
          return numeric > predicate.value
        case '>=':
          return numeric >= predicate.value
        case '<':
          return numeric < predicate.value
        case '<=':
          return numeric <= predicate.value
        default:
          return false
      }
    }
    case 'generic':
      return matchesCompiledCriteria(materializeSliceValue(view, offset, runtimeColumnStore), predicate.compiled)
  }
}

export function slicePredicateMatchesEmpty(predicate: SliceFastPredicate): boolean {
  switch (predicate.kind) {
    case 'eq-empty':
      return !predicate.negate
    case 'eq-bool':
      return predicate.negate ? predicate.value : !predicate.value
    case 'eq-number':
      return predicate.negate ? predicate.value !== 0 : predicate.value === 0
    case 'eq-string':
      return predicate.negate ? predicate.value !== '' : predicate.value === ''
    case 'cmp-number':
      return false
    case 'generic':
      return matchesCompiledCriteria({ tag: ValueTag.Empty }, predicate.compiled)
  }
}
