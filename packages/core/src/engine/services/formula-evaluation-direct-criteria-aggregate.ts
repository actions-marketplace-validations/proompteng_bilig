import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeFormula } from '../runtime-state.js'
import type { CriterionRangeCacheService, CriterionRangeMatch } from './criterion-range-cache-service.js'
import type { ExactColumnIndexService } from './exact-column-index-service.js'
import { directCriteriaRangeVersionKey, rememberDirectCriteriaResult } from './formula-evaluation-direct-criteria-cache.js'
import {
  tryEvaluateNativeDirectCriteriaMatchedAggregate,
  tryEvaluateNativeDirectCriteriaPredicateAggregate,
  type NativeDirectCriteriaPredicateLayoutCache,
} from './formula-evaluation-direct-criteria-native.js'
import { createDirectCriteriaSharingContext } from './formula-evaluation-direct-criteria-sharing.js'
import {
  applyDirectCriteriaResultTransforms,
  numericLikeValueInView,
  strictNumericAggregateCandidateInView,
  tryEvaluateDirectCriteriaTransformShortCircuit,
} from './formula-evaluation-direct-criteria-transforms.js'
import { tryEvaluateDirectIndexExactMatch, tryEvaluateDirectIndexOffset } from './formula-evaluation-direct-index.js'
import { directErrorResult, directNumberResult } from './formula-evaluation-helpers.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'

const DIRECT_CRITERIA_MATCH_CACHE_LIMIT = 16_384

export type { NativeDirectCriteriaPredicateLayoutCache }

export function createDirectCriteriaAggregateEvaluator(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas' | 'counters' | 'wasm'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly criterionCache: CriterionRangeCacheService
  readonly exactLookup: Pick<ExactColumnIndexService, 'prepareVectorLookup' | 'findPreparedVectorMatch'>
  readonly directCriteriaAggregateCache: Map<string, CellValue>
  readonly directCriteriaMatchCache: Map<string, CriterionRangeMatch>
  readonly nativeDirectCriteriaPredicateLayoutCache: NativeDirectCriteriaPredicateLayoutCache
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}): (formula: RuntimeFormula, cellIndex: number) => CellValue | undefined {
  const directCriteriaSharing = createDirectCriteriaSharingContext({
    state: args.state,
    readCellValueByIndex: args.readCellValueByIndex,
  })

  const rememberDirectCriteriaMatch = (key: string, value: CriterionRangeMatch): CriterionRangeMatch => {
    if (args.directCriteriaMatchCache.size >= DIRECT_CRITERIA_MATCH_CACHE_LIMIT) {
      const firstKey = args.directCriteriaMatchCache.keys().next().value
      if (firstKey !== undefined) {
        args.directCriteriaMatchCache.delete(firstKey)
      }
    }
    args.directCriteriaMatchCache.set(key, value)
    return value
  }

  const firstMatchedAggregateError = (
    view: ReturnType<EngineRuntimeColumnStoreService['getColumnView']>,
    rows: ArrayLike<number>,
    length: number,
  ): CellValue | undefined => {
    for (let index = 0; index < length; index += 1) {
      const row = rows[index]!
      if ((view.readTagAt(row) as ValueTag) === ValueTag.Error) {
        return directErrorResult(view.readErrorAt(row) as ErrorCode)
      }
    }
    return undefined
  }

  return (formula, cellIndex) => {
    const directCriteria = formula.directCriteria
    if (!directCriteria) return undefined
    const transformShortCircuit = tryEvaluateDirectCriteriaTransformShortCircuit(args.readCellValueByIndex, formula)
    if (transformShortCircuit) {
      return transformShortCircuit
    }
    const directIndexOffsetResult = tryEvaluateDirectIndexOffset({
      directCriteria,
      runtimeColumnStore: args.runtimeColumnStore,
      readCellValueByIndex: args.readCellValueByIndex,
      ownerRow: args.state.workbook.cellStore.rows[cellIndex],
    })
    if (directIndexOffsetResult !== undefined) {
      return applyDirectCriteriaResultTransforms(args.readCellValueByIndex, formula, directIndexOffsetResult)
    }
    const directCriteriaPairs = directCriteriaSharing.resolveDirectCriteriaPairs(formula)
    if (directCriteriaPairs?.error !== undefined) {
      return directCriteriaPairs.error
    }
    const resolvedPairs = directCriteriaPairs?.pairs ?? []
    const aggregateRange = directCriteria.aggregateRange
    const criteriaVersionKey = directCriteriaSharing.directCriteriaVersionKeyForPairs(resolvedPairs)
    const aggregateCacheKey =
      aggregateRange === undefined
        ? directCriteria.aggregateKind === 'count'
          ? [directCriteria.aggregateKind, criteriaVersionKey].join('\u0000')
          : undefined
        : [directCriteria.aggregateKind, directCriteriaRangeVersionKey(args.state, aggregateRange), criteriaVersionKey].join('\u0000')
    const cachedAggregate = aggregateCacheKey === undefined ? undefined : args.directCriteriaAggregateCache.get(aggregateCacheKey)
    if (cachedAggregate) {
      addEngineCounter(args.state.counters, 'directCriteriaAggregateCacheHits')
      return applyDirectCriteriaResultTransforms(args.readCellValueByIndex, formula, cachedAggregate)
    }
    const applyCachedAggregateResult = (value: CellValue): CellValue =>
      applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        aggregateCacheKey === undefined ? value : rememberDirectCriteriaResult(args.directCriteriaAggregateCache, aggregateCacheKey, value),
      )
    const directIndexExactMatchResult =
      resolvedPairs.length === 1
        ? tryEvaluateDirectIndexExactMatch({
            directCriteria,
            exactLookup: args.exactLookup,
            runtimeColumnStore: args.runtimeColumnStore,
            lookupValue: resolvedPairs[0]!.criteria,
          })
        : undefined
    if (directIndexExactMatchResult !== undefined) {
      return applyDirectCriteriaResultTransforms(args.readCellValueByIndex, formula, directIndexExactMatchResult)
    }
    const exactCriteriaAggregateResult =
      directCriteria.aggregateKind !== 'first'
        ? args.criterionCache.getOrBuildExactCriteriaAggregate({
            criteriaPairs: resolvedPairs,
            ...(aggregateRange === undefined ? {} : { aggregateRange }),
            aggregateKind: directCriteria.aggregateKind,
            useCompoundBucketIndex: directCriteriaSharing.shouldUseCompoundExactBucketIndex(resolvedPairs),
          })
        : undefined
    if (exactCriteriaAggregateResult !== undefined) {
      return applyCachedAggregateResult(exactCriteriaAggregateResult)
    }

    const cachedMatches = args.directCriteriaMatchCache.get(criteriaVersionKey)
    if (cachedMatches !== undefined) {
      addEngineCounter(args.state.counters, 'directCriteriaMatchCacheHits')
    }
    const shouldShareCriteriaMatches =
      directCriteria.aggregateKind !== 'first' &&
      cachedMatches === undefined &&
      directCriteriaSharing.directCriteriaShareCount(criteriaVersionKey) > 1
    let matches =
      cachedMatches ??
      (shouldShareCriteriaMatches
        ? args.criterionCache.getOrBuildMatchingRows({
            criteriaPairs: resolvedPairs,
          })
        : undefined)
    if (matches !== undefined && !('tag' in matches) && cachedMatches === undefined) {
      rememberDirectCriteriaMatch(criteriaVersionKey, matches)
    }
    if (matches === undefined) {
      const indexedPredicateAggregateResult =
        directCriteria.aggregateKind !== 'first'
          ? args.criterionCache.getOrBuildIndexedPredicateAggregate({
              criteriaPairs: resolvedPairs,
              ...(aggregateRange === undefined ? {} : { aggregateRange }),
              aggregateKind: directCriteria.aggregateKind,
            })
          : undefined
      if (indexedPredicateAggregateResult !== undefined) {
        return applyCachedAggregateResult(indexedPredicateAggregateResult)
      }

      const nativePredicateAggregateResult = tryEvaluateNativeDirectCriteriaPredicateAggregate(
        {
          state: args.state,
          runtimeColumnStore: args.runtimeColumnStore,
        },
        {
          aggregateKind: directCriteria.aggregateKind,
          aggregateRange,
          criteriaPairs: resolvedPairs,
          criteriaLayoutCache: args.nativeDirectCriteriaPredicateLayoutCache,
          criteriaLayoutCacheKey: criteriaVersionKey,
          shouldUseSharedCriteriaCache: () => directCriteriaSharing.directCriteriaShareCount(criteriaVersionKey) > 1,
        },
      )
      if (nativePredicateAggregateResult !== undefined) {
        return applyCachedAggregateResult(nativePredicateAggregateResult)
      }

      matches = args.criterionCache.getOrBuildMatchingRows({
        criteriaPairs: resolvedPairs,
      })
      if (!('tag' in matches)) {
        rememberDirectCriteriaMatch(criteriaVersionKey, matches)
      }
    }
    if ('tag' in matches) {
      return matches
    }

    if (directCriteria.aggregateKind === 'count') {
      return applyCachedAggregateResult(directNumberResult(matches.length))
    }

    if (!aggregateRange) {
      return undefined
    }
    const concreteAggregateCacheKey = aggregateCacheKey
    if (concreteAggregateCacheKey === undefined) {
      return undefined
    }

    if (directCriteria.aggregateKind === 'first') {
      const firstOffset = matches.rows[0]
      const result =
        firstOffset === undefined
          ? directErrorResult(ErrorCode.NA)
          : args.runtimeColumnStore
              .getColumnView({
                sheetName: aggregateRange.sheetName,
                rowStart: aggregateRange.rowStart,
                rowEnd: aggregateRange.rowEnd,
                col: aggregateRange.col,
              })
              .readCellValueAt(firstOffset)
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(args.directCriteriaAggregateCache, concreteAggregateCacheKey, result),
      )
    }

    const nativeAggregateResult = tryEvaluateNativeDirectCriteriaMatchedAggregate(
      {
        state: args.state,
        runtimeColumnStore: args.runtimeColumnStore,
      },
      {
        aggregateKind: directCriteria.aggregateKind,
        aggregateRange,
        matches,
      },
    )
    if (nativeAggregateResult !== undefined) {
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(args.directCriteriaAggregateCache, concreteAggregateCacheKey, nativeAggregateResult),
      )
    }

    const aggregateView = args.runtimeColumnStore.getColumnView({
      sheetName: aggregateRange.sheetName,
      rowStart: aggregateRange.rowStart,
      rowEnd: aggregateRange.rowEnd,
      col: aggregateRange.col,
    })
    const matchedAggregateError = firstMatchedAggregateError(aggregateView, matches.rows, matches.length)
    if (matchedAggregateError) {
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(args.directCriteriaAggregateCache, concreteAggregateCacheKey, matchedAggregateError),
      )
    }

    if (directCriteria.aggregateKind === 'sum') {
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        sum += numericLikeValueInView(aggregateView, matches.rows[index]!) ?? 0
      }
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(args.directCriteriaAggregateCache, concreteAggregateCacheKey, directNumberResult(sum)),
      )
    }

    if (directCriteria.aggregateKind === 'average') {
      let count = 0
      let sum = 0
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = numericLikeValueInView(aggregateView, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(
          args.directCriteriaAggregateCache,
          concreteAggregateCacheKey,
          count === 0 ? directErrorResult(ErrorCode.Div0) : directNumberResult(sum / count),
        ),
      )
    }

    if (directCriteria.aggregateKind === 'min') {
      let minimum = Number.POSITIVE_INFINITY
      for (let index = 0; index < matches.length; index += 1) {
        const numeric = strictNumericAggregateCandidateInView(aggregateView, matches.rows[index]!)
        if (numeric === undefined) {
          continue
        }
        minimum = Math.min(minimum, numeric)
      }
      return applyDirectCriteriaResultTransforms(
        args.readCellValueByIndex,
        formula,
        rememberDirectCriteriaResult(
          args.directCriteriaAggregateCache,
          concreteAggregateCacheKey,
          directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum),
        ),
      )
    }

    let maximum = Number.NEGATIVE_INFINITY
    for (let index = 0; index < matches.length; index += 1) {
      const numeric = strictNumericAggregateCandidateInView(aggregateView, matches.rows[index]!)
      if (numeric === undefined) {
        continue
      }
      maximum = Math.max(maximum, numeric)
    }
    return applyDirectCriteriaResultTransforms(
      args.readCellValueByIndex,
      formula,
      rememberDirectCriteriaResult(
        args.directCriteriaAggregateCache,
        concreteAggregateCacheKey,
        directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum),
      ),
    )
  }
}
