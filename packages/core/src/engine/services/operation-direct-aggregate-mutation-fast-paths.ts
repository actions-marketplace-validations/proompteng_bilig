import type { EngineExistingNumericCellMutationResult } from '../../cell-mutations-at.js'
import {
  tryApplySingleDirectAggregateLiteralMutationFastPath,
  tryApplyTrustedColumnDirectAggregateExistingNumericMutation,
  tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation,
  type OperationDirectAggregateLiteralFastPathArgs,
  type OperationDirectAggregateLiteralMutationRequest,
  type OperationTrustedColumnDirectAggregateExistingNumericMutationRequest,
  type OperationTrustedRangeDirectAggregateExistingNumericMutationRequest,
} from './operation-direct-aggregate-literal-fast-path.js'

export type OperationDirectAggregateMutationFastPathArgs = OperationDirectAggregateLiteralFastPathArgs

export interface OperationDirectAggregateMutationFastPaths {
  readonly tryApplySingleDirectAggregateLiteralMutationFastPath: (
    request: OperationDirectAggregateLiteralMutationRequest,
  ) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation: (
    request: OperationTrustedRangeDirectAggregateExistingNumericMutationRequest,
  ) => EngineExistingNumericCellMutationResult | null
  readonly tryApplyTrustedColumnDirectAggregateExistingNumericMutation: (
    request: OperationTrustedColumnDirectAggregateExistingNumericMutationRequest,
  ) => EngineExistingNumericCellMutationResult | null
}

export function createOperationDirectAggregateMutationFastPaths(
  args: OperationDirectAggregateMutationFastPathArgs,
): OperationDirectAggregateMutationFastPaths {
  return {
    tryApplySingleDirectAggregateLiteralMutationFastPath: (request) => tryApplySingleDirectAggregateLiteralMutationFastPath(args, request),
    tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation: (request) =>
      tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation(args, request),
    tryApplyTrustedColumnDirectAggregateExistingNumericMutation: (request) =>
      tryApplyTrustedColumnDirectAggregateExistingNumericMutation(args, request),
  }
}
