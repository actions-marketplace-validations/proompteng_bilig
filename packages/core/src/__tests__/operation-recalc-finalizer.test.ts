import { describe, expect, it } from 'vitest'
import { CellFlags } from '../cell-store.js'
import type { RuntimeDirectAggregateDescriptor } from '../engine/runtime-state.js'
import { createCachedCycleDependencyLookup } from '../engine/services/operation-recalc-finalizer.js'

function directAggregate(overrides: Partial<RuntimeDirectAggregateDescriptor> = {}): RuntimeDirectAggregateDescriptor {
  return {
    regionId: 1,
    aggregateKind: 'sum',
    sheetName: 'Sheet1',
    rowStart: 0,
    rowEnd: 1023,
    col: 0,
    colEnd: 0,
    length: 1024,
    ...overrides,
  }
}

describe('operation recalc finalizer', () => {
  it('caches direct aggregate cycle dependency scans by range', () => {
    const formulas = new Map([
      [20, { dependencyIndices: new Uint32Array(), directAggregate: directAggregate() }],
      [21, { dependencyIndices: new Uint32Array(), directAggregate: directAggregate({ regionId: 99 }) }],
    ])
    const flags: Array<number | undefined> = []
    flags[2] = CellFlags.InCycle
    let dependencyWalks = 0
    const hasCycleDependency = createCachedCycleDependencyLookup({
      formulas: { get: (cellIndex) => formulas.get(cellIndex) },
      flags,
      forEachFormulaDependencyCell: (_cellIndex, fn) => {
        dependencyWalks += 1
        fn(1)
        fn(2)
      },
    })

    expect(hasCycleDependency(20)).toBe(true)
    expect(hasCycleDependency(21)).toBe(true)
    expect(dependencyWalks).toBe(1)
  })

  it('does not cache formulas with extra direct dependencies under the aggregate range key', () => {
    const formulas = new Map([
      [20, { dependencyIndices: Uint32Array.of(7), directAggregate: directAggregate() }],
      [21, { dependencyIndices: Uint32Array.of(8), directAggregate: directAggregate() }],
    ])
    let dependencyWalks = 0
    const hasCycleDependency = createCachedCycleDependencyLookup({
      formulas: { get: (cellIndex) => formulas.get(cellIndex) },
      flags: [],
      forEachFormulaDependencyCell: () => {
        dependencyWalks += 1
      },
    })

    expect(hasCycleDependency(20)).toBe(false)
    expect(hasCycleDependency(21)).toBe(false)
    expect(dependencyWalks).toBe(2)
  })
})
