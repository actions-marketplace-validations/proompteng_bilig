import { normalizeExactLookupNumber } from '@bilig/formula'
import { ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeColumnView } from './runtime-column-store-service.js'
import type { CriterionExactAggregateRequest } from './criterion-range-cache-service.js'
import type { CriterionEqualityIndexKey } from './criterion-range-equality-index.js'

export function criteriaCacheKey(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${normalizeExactLookupNumber(value.value)}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${value.value}`
    case ValueTag.Error:
      return `r:${value.code}`
  }
}

export function exactAggregateIndexCacheKey(request: {
  readonly criteriaRegionId: number
  readonly criteriaView: RuntimeColumnView
  readonly aggregateRegionId?: number
  readonly aggregateView?: RuntimeColumnView
}): string {
  return [
    request.criteriaRegionId,
    request.criteriaView.columnVersion,
    request.criteriaView.structureVersion,
    request.aggregateRegionId ?? -1,
    request.aggregateView?.columnVersion ?? -1,
    request.aggregateView?.structureVersion ?? -1,
  ].join('\u0001')
}

export function exactKeyCachePart(key: CriterionEqualityIndexKey): string {
  return key.kind === 'number' ? `n:${key.value}` : `s:${key.value}`
}

export function compoundExactTupleKey(keys: readonly CriterionEqualityIndexKey[]): string {
  return keys.map(exactKeyCachePart).join('\u0002')
}

export function compoundExactAggregateCacheKey(request: {
  readonly aggregateKind: CriterionExactAggregateRequest['aggregateKind']
  readonly criteriaPairs: readonly {
    readonly regionId: number
    readonly view: RuntimeColumnView
    readonly key: CriterionEqualityIndexKey
  }[]
  readonly aggregateRegionId?: number
  readonly aggregateView?: RuntimeColumnView
}): string {
  return [
    request.aggregateKind,
    request.criteriaPairs
      .map(({ regionId, view, key }) => `${regionId}:${view.columnVersion}:${view.structureVersion}:${exactKeyCachePart(key)}`)
      .join('\u0002'),
    request.aggregateRegionId ?? -1,
    request.aggregateView?.columnVersion ?? -1,
    request.aggregateView?.structureVersion ?? -1,
  ].join('\u0001')
}

export function compoundExactAggregateIndexCacheKey(request: {
  readonly criteriaPairs: readonly {
    readonly regionId: number
    readonly view: RuntimeColumnView
  }[]
  readonly aggregateRegionId?: number
  readonly aggregateView?: RuntimeColumnView
}): string {
  return [
    request.criteriaPairs.map(({ regionId, view }) => `${regionId}:${view.columnVersion}:${view.structureVersion}`).join('\u0002'),
    request.aggregateRegionId ?? -1,
    request.aggregateView?.columnVersion ?? -1,
    request.aggregateView?.structureVersion ?? -1,
  ].join('\u0001')
}
