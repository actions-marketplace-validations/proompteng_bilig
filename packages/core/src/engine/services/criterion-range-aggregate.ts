import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { RuntimeColumnView } from './runtime-column-store-service.js'
import type { CriterionExactAggregateRequest } from './criterion-range-cache-service.js'
import type { CriterionEqualityIndexKey } from './criterion-range-equality-index.js'
import { decodeValueTag, errorValue, numberValue } from './criterion-range-values.js'

export interface CriterionExactAggregateBucket {
  count: number
  sum: number
  numericCount: number
  minimum: number
  maximum: number
  firstError?: CellValue
}

export interface CriterionExactAggregateIndex {
  readonly numbers: ReadonlyMap<number, CriterionExactAggregateBucket>
  readonly strings: ReadonlyMap<string, CriterionExactAggregateBucket>
}

export function createExactAggregateBucket(): CriterionExactAggregateBucket {
  return {
    count: 0,
    sum: 0,
    numericCount: 0,
    minimum: Number.POSITIVE_INFINITY,
    maximum: Number.NEGATIVE_INFINITY,
  }
}

export function addExactAggregateMatch(
  bucket: CriterionExactAggregateBucket,
  aggregateView: RuntimeColumnView | undefined,
  offset: number,
): void {
  bucket.count += 1
  if (aggregateView === undefined) {
    return
  }
  const tag = decodeValueTag(aggregateView.readTagAt(offset))
  if (tag === ValueTag.Error) {
    bucket.firstError ??= { tag: ValueTag.Error, code: aggregateView.readErrorAt(offset) ?? ErrorCode.None }
    return
  }
  if (tag === ValueTag.Number) {
    const numeric = aggregateView.readNumberAt(offset)
    bucket.sum += numeric
    bucket.numericCount += 1
    bucket.minimum = Math.min(bucket.minimum, numeric)
    bucket.maximum = Math.max(bucket.maximum, numeric)
    return
  }
  if (tag === ValueTag.Boolean) {
    bucket.sum += aggregateView.readNumberAt(offset) !== 0 ? 1 : 0
    bucket.numericCount += 1
    return
  }
  if (tag === ValueTag.Empty) {
    bucket.numericCount += 1
  }
}

export function addExactAggregateBucket(
  index: { numbers: Map<number, CriterionExactAggregateBucket>; strings: Map<string, CriterionExactAggregateBucket> },
  key: CriterionEqualityIndexKey,
  aggregateView: RuntimeColumnView | undefined,
  offset: number,
): void {
  let bucket = key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value)
  if (bucket === undefined) {
    bucket = createExactAggregateBucket()
    if (key.kind === 'number') {
      index.numbers.set(key.value, bucket)
    } else {
      index.strings.set(key.value, bucket)
    }
  }
  addExactAggregateMatch(bucket, aggregateView, offset)
}

export function exactAggregateValueFromBucket(
  bucket: CriterionExactAggregateBucket | undefined,
  aggregateKind: CriterionExactAggregateRequest['aggregateKind'],
): CellValue {
  if (bucket === undefined) {
    return aggregateKind === 'average' ? errorValue(ErrorCode.Div0) : numberValue(0)
  }
  if (aggregateKind === 'count') {
    return numberValue(bucket.count)
  }
  if (bucket.firstError) {
    return bucket.firstError
  }
  if (aggregateKind === 'sum') {
    return numberValue(bucket.sum)
  }
  if (aggregateKind === 'average') {
    return bucket.numericCount === 0 ? errorValue(ErrorCode.Div0) : numberValue(bucket.sum / bucket.numericCount)
  }
  if (aggregateKind === 'min') {
    return numberValue(bucket.minimum === Number.POSITIVE_INFINITY ? 0 : bucket.minimum)
  }
  return numberValue(bucket.maximum === Number.NEGATIVE_INFINITY ? 0 : bucket.maximum)
}

export function exactAggregateBucketValue(
  index: CriterionExactAggregateIndex,
  key: CriterionEqualityIndexKey,
  aggregateKind: CriterionExactAggregateRequest['aggregateKind'],
): CellValue {
  return exactAggregateValueFromBucket(key.kind === 'number' ? index.numbers.get(key.value) : index.strings.get(key.value), aggregateKind)
}
