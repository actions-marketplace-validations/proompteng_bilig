import { describe, expect, it } from 'vitest'
import type { RecalcMetrics } from '@bilig/protocol'
import type { SheetRecord } from '../workbook-store.js'
import {
  createOperationDirectAggregateMutationFastPaths,
  type OperationDirectAggregateMutationFastPathArgs,
} from '../engine/services/operation-direct-aggregate-mutation-fast-paths.js'

function recalcMetrics(): RecalcMetrics {
  return {
    batchId: 0,
    changedInputCount: 0,
    dirtyFormulaCount: 0,
    wasmFormulaCount: 0,
    jsFormulaCount: 0,
    rangeNodeVisits: 0,
    recalcMs: 0,
    compileMs: 0,
  }
}

function makeArgs(): {
  readonly args: OperationDirectAggregateMutationFastPathArgs
  readonly writes: number[]
  readonly deferred: number[]
} {
  const writes: number[] = []
  const deferred: number[] = []
  return {
    writes,
    deferred,
    args: {
      state: {
        workbook: {
          cellStore: {
            rows: [],
            cols: [],
            sheetIds: [],
          },
        },
        counters: {
          directAggregateDeltaApplications: 0,
          directAggregateDeltaOnlyRecalcSkips: 0,
        },
        events: {
          emitBatch: () => {},
        },
        setLastMetrics: () => {},
      },
      directRangePostRecalcLimit: 16_384,
      getSingleEntityDependent: () => 4,
      collectAffectedDirectRangeDependents: () => [],
      collectSingleApplicableDirectAggregateDependent: () => 4,
      canApplyDirectAggregateLiteralDeltaForRequest: () => true,
      canApplyDirectAggregateLiteralDelta: () => true,
      writeFastPathLiteralToExistingCell: (cellIndex) => {
        writes.push(cellIndex)
      },
      writeTrustedExistingNumericLiteralToCell: (cellIndex) => {
        writes.push(cellIndex)
      },
      applyTerminalDirectFormulaNumericDeltaAndReturn: () => 12,
      applyDirectFormulaNumericDelta: () => true,
      applyDirectFormulaNumericDeltaBatch: () => true,
      cellsShareVersionColumn: () => false,
      withOptionalColumnVersionBatch: (_enabled, apply) => apply(),
      deferSingleCellKernelSync: (cellIndex) => {
        deferred.push(cellIndex)
      },
      makeSingleLiteralSkipMetrics: recalcMetrics,
    },
  }
}

describe('createOperationDirectAggregateMutationFastPaths', () => {
  it('centralizes direct aggregate fast-path wiring for trusted range and column mutations', () => {
    const sheet: SheetRecord = { id: 1, name: 'Sheet1', order: 0 }
    const { args, writes, deferred } = makeArgs()
    const fastPaths = createOperationDirectAggregateMutationFastPaths(args)

    expect(
      fastPaths.tryApplyTrustedSingleRangeDirectAggregateExistingNumericMutation({
        existingIndex: 2,
        rangeEntityDependent: 99,
        sheet,
        sheetId: 1,
        col: 1,
        value: 7,
        delta: 5,
        hasExactLookupDependents: false,
        hasSortedLookupDependents: false,
      }),
    ).toEqual({
      firstChangedCellIndex: 2,
      secondChangedCellIndex: 4,
      changedCellCount: 2,
      secondChangedNumericValue: 12,
      secondChangedRow: 0,
      secondChangedCol: 0,
      explicitChangedCount: 1,
    })

    expect(
      fastPaths.tryApplyTrustedColumnDirectAggregateExistingNumericMutation({
        existingIndex: 3,
        sheet,
        sheetId: 1,
        sheetName: 'Sheet1',
        row: 1,
        col: 1,
        value: 8,
        delta: 6,
        hasExactLookupDependents: false,
        hasSortedLookupDependents: false,
      }),
    ).toEqual({
      firstChangedCellIndex: 3,
      secondChangedCellIndex: 4,
      changedCellCount: 2,
      secondChangedNumericValue: 12,
      secondChangedRow: 0,
      secondChangedCol: 0,
      explicitChangedCount: 1,
    })

    expect(writes).toEqual([2, 3])
    expect(deferred).toEqual([2, 3])
  })
})
