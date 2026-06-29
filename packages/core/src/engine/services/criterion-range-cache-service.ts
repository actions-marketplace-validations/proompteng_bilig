import { ErrorCode, type CellValue } from '@bilig/protocol'
import { compileCriteriaMatcher } from '@bilig/formula'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import type { DepPatternStore } from '../../deps/dep-pattern-store.js'
import type { RegionGraph } from '../../deps/region-graph.js'
import { getOrBuildIndexedPredicateAggregateFromColumnViews } from './criterion-range-indexed-predicate-aggregate.js'
import {
  addExactAggregateBucket,
  addExactAggregateMatch,
  createExactAggregateBucket,
  exactAggregateBucketValue,
  exactAggregateValueFromBucket,
  type CriterionExactAggregateBucket,
  type CriterionExactAggregateIndex,
} from './criterion-range-aggregate.js'
import {
  compoundExactAggregateCacheKey,
  compoundExactAggregateIndexCacheKey,
  compoundExactTupleKey,
  criteriaCacheKey,
  exactAggregateIndexCacheKey,
} from './criterion-range-cache-keys.js'
import {
  equalityIndexKeyForPredicate,
  equalityIndexKeyForStoredValue,
  readIndexedEqualityRows,
  type CriterionEqualityIndexKey,
} from './criterion-range-equality-index.js'
import { buildSlicePredicate, slicePredicateMatches, slicePredicateMatchesEmpty } from './criterion-range-predicate.js'
import { countNonEmptyRowsInView, forEachNonEmptyRowOffsetInView } from './criterion-range-row-iteration.js'
import { decodeValueTag, errorValue, materializeSliceValue } from './criterion-range-values.js'

export type { CriterionExactAggregateBucket } from './criterion-range-aggregate.js'
export type { SliceFastPredicate } from './criterion-range-predicate.js'

export interface CriterionRangeDescriptor {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
}

export interface CriterionRangePair {
  readonly range: CriterionRangeDescriptor
  readonly criteria: CellValue
}

export interface CriterionRangeMatch {
  readonly rows: Uint32Array
  readonly length: number
}

export interface CriterionRangeCacheService {
  readonly getOrBuildMatchingRows: (request: { criteriaPairs: readonly CriterionRangePair[] }) => CriterionRangeMatch | CellValue
  readonly getOrBuildExactCriteriaAggregate: (request: CriterionCompoundExactAggregateRequest) => CellValue | undefined
  readonly getOrBuildExactAggregate: (request: CriterionExactAggregateRequest) => CellValue | undefined
  readonly getOrBuildCompoundExactAggregate: (request: CriterionCompoundExactAggregateRequest) => CellValue | undefined
  readonly getOrBuildIndexedPredicateAggregate: (request: CriterionCompoundExactAggregateRequest) => CellValue | undefined
}

export interface CriterionExactAggregateRequest {
  readonly criteriaPair: CriterionRangePair
  readonly aggregateRange?: CriterionRangeDescriptor
  readonly aggregateKind: 'count' | 'sum' | 'average' | 'min' | 'max'
}

export interface CriterionCompoundExactAggregateRequest {
  readonly criteriaPairs: readonly CriterionRangePair[]
  readonly aggregateRange?: CriterionRangeDescriptor
  readonly aggregateKind: CriterionExactAggregateRequest['aggregateKind']
  readonly useCompoundBucketIndex?: boolean
}

interface CriterionCompoundExactAggregateIndex {
  readonly buckets: ReadonlyMap<string, CriterionExactAggregateBucket>
}

const COMPOUND_EXACT_AGGREGATE_BUCKET_LIMIT = 16_384

export function createCriterionRangeCacheService(args: {
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly regionGraph: Pick<RegionGraph, 'internSingleColumnRegion'>
  readonly depPatternStore: DepPatternStore
}): CriterionRangeCacheService {
  const exactAggregateIndexes = new Map<string, CriterionExactAggregateIndex>()
  const compoundExactAggregateResults = new Map<string, CellValue>()
  const compoundExactAggregateIndexes = new Map<string, CriterionCompoundExactAggregateIndex>()

  const getColumnView = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnView => {
    const direct = Reflect.get(args.runtimeColumnStore, 'getColumnView')
    if (typeof direct === 'function') {
      return direct.call(args.runtimeColumnStore, request)
    }
    const slice = args.runtimeColumnStore.getColumnSlice(request)
    return {
      owner: {
        sheetName: slice.sheetName,
        col: slice.col,
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
        sheetColumnVersions: slice.sheetColumnVersions,
        pages: new Map(),
      },
      sheetName: slice.sheetName,
      rowStart: slice.rowStart,
      rowEnd: slice.rowEnd,
      col: slice.col,
      length: slice.length,
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      sheetColumnVersions: slice.sheetColumnVersions,
      readTagAt(offset) {
        return slice.tags[offset] ?? 0
      },
      readNumberAt(offset) {
        return slice.numbers[offset] ?? 0
      },
      readStringIdAt(offset) {
        return slice.stringIds[offset] ?? 0
      },
      readErrorAt(offset) {
        return slice.errors[offset] ?? ErrorCode.None
      },
      readCellValueAt(offset) {
        return materializeSliceValue(this, offset, args.runtimeColumnStore)
      },
    }
  }

  const rangeRegionId = (range: CriterionRangeDescriptor): number =>
    args.regionGraph.internSingleColumnRegion({
      sheetName: range.sheetName,
      rowStart: range.rowStart,
      rowEnd: range.rowEnd,
      col: range.col,
    })

  const buildExactAggregateIndex = (
    criteriaView: RuntimeColumnView,
    aggregateView: RuntimeColumnView | undefined,
  ): CriterionExactAggregateIndex => {
    const buckets = {
      numbers: new Map<number, CriterionExactAggregateBucket>(),
      strings: new Map<string, CriterionExactAggregateBucket>(),
    }
    criteriaView.owner.pages.forEach((page) => {
      const rowStart = Math.max(criteriaView.rowStart, page.rowStart)
      const rowEnd = Math.min(criteriaView.rowEnd, page.rowStart + page.tags.length - 1)
      if (rowStart > rowEnd || page.nonEmptyCount === 0) {
        return
      }
      for (let row = rowStart; row <= rowEnd; row += 1) {
        const localRow = row - page.rowStart
        const key = equalityIndexKeyForStoredValue(
          decodeValueTag(page.tags[localRow]),
          page.numbers[localRow] ?? 0,
          page.stringIds[localRow] ?? 0,
          args.runtimeColumnStore,
        )
        if (key === undefined) {
          continue
        }
        addExactAggregateBucket(buckets, key, aggregateView, row - criteriaView.rowStart)
      }
    })
    return buckets
  }

  const equalityKeyAtOffset = (view: RuntimeColumnView, offset: number): CriterionEqualityIndexKey | undefined =>
    equalityIndexKeyForStoredValue(
      decodeValueTag(view.readTagAt(offset)),
      view.readNumberAt(offset),
      view.readStringIdAt(offset),
      args.runtimeColumnStore,
    )

  const buildCompoundExactAggregateIndex = (
    resolvedPairs: readonly {
      readonly view: RuntimeColumnView
    }[],
    aggregateView: RuntimeColumnView | undefined,
    length: number,
  ): CriterionCompoundExactAggregateIndex | undefined => {
    const buckets = new Map<string, CriterionExactAggregateBucket>()
    for (let rowOffset = 0; rowOffset < length; rowOffset += 1) {
      const rowKeys: CriterionEqualityIndexKey[] = []
      for (const pair of resolvedPairs) {
        const key = equalityKeyAtOffset(pair.view, rowOffset)
        if (key === undefined) {
          rowKeys.length = 0
          break
        }
        rowKeys.push(key)
      }
      if (rowKeys.length === 0) {
        continue
      }
      const tupleKey = compoundExactTupleKey(rowKeys)
      let bucket = buckets.get(tupleKey)
      if (bucket === undefined) {
        if (buckets.size >= COMPOUND_EXACT_AGGREGATE_BUCKET_LIMIT) {
          return undefined
        }
        bucket = createExactAggregateBucket()
        buckets.set(tupleKey, bucket)
      }
      addExactAggregateMatch(bucket, aggregateView, rowOffset)
    }
    return { buckets }
  }

  const getOrBuildExactAggregate = (request: CriterionExactAggregateRequest): CellValue | undefined => {
    const criteriaPredicate = buildSlicePredicate(compileCriteriaMatcher(request.criteriaPair.criteria))
    const criteriaKey = equalityIndexKeyForPredicate(criteriaPredicate)
    if (criteriaKey === undefined) {
      return undefined
    }
    if (request.aggregateKind !== 'count' && request.aggregateRange === undefined) {
      return undefined
    }
    if (request.aggregateRange !== undefined && request.aggregateRange.length !== request.criteriaPair.range.length) {
      return undefined
    }
    const criteriaRegionId = rangeRegionId(request.criteriaPair.range)
    const criteriaView = getColumnView(request.criteriaPair.range)
    const aggregateRegionId = request.aggregateRange === undefined ? undefined : rangeRegionId(request.aggregateRange)
    const aggregateView = request.aggregateRange === undefined ? undefined : getColumnView(request.aggregateRange)
    const cacheKey = exactAggregateIndexCacheKey({
      criteriaRegionId,
      criteriaView,
      ...(aggregateRegionId === undefined ? {} : { aggregateRegionId }),
      ...(aggregateView === undefined ? {} : { aggregateView }),
    })
    let index = exactAggregateIndexes.get(cacheKey)
    if (index === undefined) {
      index = buildExactAggregateIndex(criteriaView, aggregateView)
      exactAggregateIndexes.set(cacheKey, index)
    }
    return exactAggregateBucketValue(index, criteriaKey, request.aggregateKind)
  }

  const getOrBuildCompoundExactAggregate = (request: CriterionCompoundExactAggregateRequest): CellValue | undefined => {
    const { criteriaPairs } = request
    if (criteriaPairs.length < 2) {
      return undefined
    }
    const expectedLength = criteriaPairs[0]!.range.length
    if (criteriaPairs.some((pair) => pair.range.length !== expectedLength)) {
      return undefined
    }
    if (request.aggregateKind !== 'count' && request.aggregateRange === undefined) {
      return undefined
    }
    if (request.aggregateRange !== undefined && request.aggregateRange.length !== expectedLength) {
      return undefined
    }

    const predicates = criteriaPairs.map((pair) => buildSlicePredicate(compileCriteriaMatcher(pair.criteria)))
    const keys: CriterionEqualityIndexKey[] = []
    for (const predicate of predicates) {
      const key = equalityIndexKeyForPredicate(predicate)
      if (key === undefined) {
        return undefined
      }
      keys.push(key)
    }

    const resolvedPairs = criteriaPairs.map((pair, index) => ({
      regionId: rangeRegionId(pair.range),
      view: getColumnView(pair.range),
      predicate: predicates[index]!,
      key: keys[index]!,
    }))
    const aggregateRegionId = request.aggregateRange === undefined ? undefined : rangeRegionId(request.aggregateRange)
    const aggregateView = request.aggregateRange === undefined ? undefined : getColumnView(request.aggregateRange)
    const cacheKey = compoundExactAggregateCacheKey({
      aggregateKind: request.aggregateKind,
      criteriaPairs: resolvedPairs,
      ...(aggregateRegionId === undefined ? {} : { aggregateRegionId }),
      ...(aggregateView === undefined ? {} : { aggregateView }),
    })
    const cached = compoundExactAggregateResults.get(cacheKey)
    if (cached !== undefined) {
      return cached
    }

    if (request.useCompoundBucketIndex) {
      const indexCacheKey = compoundExactAggregateIndexCacheKey({
        criteriaPairs: resolvedPairs,
        ...(aggregateRegionId === undefined ? {} : { aggregateRegionId }),
        ...(aggregateView === undefined ? {} : { aggregateView }),
      })
      let index = compoundExactAggregateIndexes.get(indexCacheKey)
      if (index === undefined) {
        index = buildCompoundExactAggregateIndex(resolvedPairs, aggregateView, expectedLength)
        if (index !== undefined) {
          compoundExactAggregateIndexes.set(indexCacheKey, index)
        }
      }
      if (index !== undefined) {
        const result = exactAggregateValueFromBucket(index.buckets.get(compoundExactTupleKey(keys)), request.aggregateKind)
        compoundExactAggregateResults.set(cacheKey, result)
        return result
      }
    }

    const rowsets = resolvedPairs.map(({ view, predicate }) => readIndexedEqualityRows(view, predicate, args.runtimeColumnStore))
    if (rowsets.some((rowset) => rowset === undefined)) {
      return undefined
    }

    let limitingPairIndex = 0
    let limitingRows = rowsets[0]!
    for (let index = 1; index < rowsets.length; index += 1) {
      const rowset = rowsets[index]!
      if (rowset.cardinality < limitingRows.cardinality) {
        limitingPairIndex = index
        limitingRows = rowset
      }
    }

    const bucket = createExactAggregateBucket()
    limitingRows.forEachOffset((rowOffset) => {
      for (let pairIndex = 0; pairIndex < resolvedPairs.length; pairIndex += 1) {
        if (pairIndex === limitingPairIndex) {
          continue
        }
        const pair = resolvedPairs[pairIndex]!
        if (!slicePredicateMatches(pair.predicate, pair.view, rowOffset, args.runtimeColumnStore)) {
          return
        }
      }
      addExactAggregateMatch(bucket, aggregateView, rowOffset)
    })

    const result = exactAggregateValueFromBucket(bucket.count === 0 ? undefined : bucket, request.aggregateKind)
    compoundExactAggregateResults.set(cacheKey, result)
    return result
  }

  const getOrBuildExactCriteriaAggregate = (request: CriterionCompoundExactAggregateRequest): CellValue | undefined => {
    return request.criteriaPairs.length === 1
      ? getOrBuildExactAggregate({
          criteriaPair: request.criteriaPairs[0]!,
          ...(request.aggregateRange === undefined ? {} : { aggregateRange: request.aggregateRange }),
          aggregateKind: request.aggregateKind,
        })
      : getOrBuildCompoundExactAggregate(request)
  }

  const getOrBuildIndexedPredicateAggregate = (request: CriterionCompoundExactAggregateRequest): CellValue | undefined => {
    return getOrBuildIndexedPredicateAggregateFromColumnViews({
      request,
      runtimeColumnStore: args.runtimeColumnStore,
      getColumnView,
      buildSlicePredicate,
      readIndexedEqualityRows,
      slicePredicateMatches,
      createExactAggregateBucket,
      addExactAggregateMatch,
      exactAggregateValueFromBucket,
    })
  }

  const getOrBuildMatchingRows = (request: { criteriaPairs: readonly CriterionRangePair[] }): CriterionRangeMatch | CellValue => {
    const { criteriaPairs } = request
    if (criteriaPairs.length === 0) {
      return errorValue(ErrorCode.Value)
    }
    const expectedLength = criteriaPairs[0]!.range.length
    if (criteriaPairs.some((pair) => pair.range.length !== expectedLength)) {
      return errorValue(ErrorCode.Value)
    }

    const resolvedPairs = criteriaPairs.map((pair) => ({
      regionId: args.regionGraph.internSingleColumnRegion({
        sheetName: pair.range.sheetName,
        rowStart: pair.range.rowStart,
        rowEnd: pair.range.rowEnd,
        col: pair.range.col,
      }),
      view: getColumnView({
        sheetName: pair.range.sheetName,
        rowStart: pair.range.rowStart,
        rowEnd: pair.range.rowEnd,
        col: pair.range.col,
      }),
      criteria: pair.criteria,
    }))
    const versionStamp = resolvedPairs
      .map(({ regionId, view }) => `${regionId}:${view.columnVersion}:${view.structureVersion}`)
      .join('\u0001')
    const existing = args.depPatternStore.getCriteriaPattern({
      regionIds: resolvedPairs.map(({ regionId }) => regionId),
      criteriaKeys: resolvedPairs.map(({ criteria }) => criteriaCacheKey(criteria)),
      versionStamp,
    })
    if (existing) {
      return existing
    }

    const predicates = criteriaPairs.map((pair) => buildSlicePredicate(compileCriteriaMatcher(pair.criteria)))
    const matchingRows: number[] = []
    let limitingPairIndex: number | undefined
    let limitingIndexedRows: ReturnType<typeof readIndexedEqualityRows> | undefined
    let limitingPairNonEmptyRows = Number.POSITIVE_INFINITY
    for (let pairIndex = 0; pairIndex < predicates.length; pairIndex += 1) {
      const indexedRows = readIndexedEqualityRows(resolvedPairs[pairIndex]!.view, predicates[pairIndex]!, args.runtimeColumnStore)
      if (indexedRows !== undefined) {
        if (indexedRows.cardinality < limitingPairNonEmptyRows) {
          limitingPairIndex = undefined
          limitingIndexedRows = indexedRows
          limitingPairNonEmptyRows = indexedRows.cardinality
        }
        continue
      }
      if (slicePredicateMatchesEmpty(predicates[pairIndex]!)) {
        continue
      }
      const view = resolvedPairs[pairIndex]!.view
      if (view.owner.pages.size === 0) {
        continue
      }
      const nonEmptyRows = countNonEmptyRowsInView(view)
      if (nonEmptyRows < limitingPairNonEmptyRows) {
        limitingPairIndex = pairIndex
        limitingIndexedRows = undefined
        limitingPairNonEmptyRows = nonEmptyRows
      }
    }
    const visitCandidate = (rowOffset: number): void => {
      let matches = true
      for (let pairIndex = 0; pairIndex < predicates.length; pairIndex += 1) {
        if (!slicePredicateMatches(predicates[pairIndex]!, resolvedPairs[pairIndex]!.view, rowOffset, args.runtimeColumnStore)) {
          matches = false
          break
        }
      }
      if (matches) {
        matchingRows.push(rowOffset)
      }
    }
    if (limitingIndexedRows !== undefined) {
      limitingIndexedRows.forEachOffset(visitCandidate)
    } else if (limitingPairIndex === undefined) {
      for (let rowOffset = 0; rowOffset < expectedLength; rowOffset += 1) {
        visitCandidate(rowOffset)
      }
    } else {
      forEachNonEmptyRowOffsetInView(resolvedPairs[limitingPairIndex]!.view, visitCandidate)
    }

    return args.depPatternStore.setCriteriaPattern({
      regionIds: resolvedPairs.map(({ regionId }) => regionId),
      criteriaKeys: resolvedPairs.map(({ criteria }) => criteriaCacheKey(criteria)),
      versionStamp,
      rows: Uint32Array.from(matchingRows),
      length: matchingRows.length,
    })
  }

  return {
    getOrBuildCompoundExactAggregate,
    getOrBuildExactCriteriaAggregate,
    getOrBuildExactAggregate,
    getOrBuildIndexedPredicateAggregate,
    getOrBuildMatchingRows,
  }
}
