import type { CompiledFormula } from '@bilig/formula'
import type { RuntimeFormula } from '../runtime-state.js'
import type { ParsedCompiledFormula } from './formula-binding-direct-descriptors.js'
import { createFormulaBindingFamilyIndexController } from './formula-binding-family-index-controller.js'
import type { FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'
import { createFormulaBindingInstanceTableRebuildController } from './formula-binding-instance-table-rebuild.js'
import { createFormulaBindingInstanceTracker } from './formula-binding-instance-tracker.js'
import { createFormulaBindingMemberCounts } from './formula-binding-member-counts.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'
import { createFormulaBindingSheetIndex } from './formula-binding-sheet-index.js'
import { updateFormulaBindingVolatileIndex } from './formula-binding-volatile-index.js'
import { rebuildDeferredFormulaFamilyIndex } from './formula-family-index-rebuild.js'

export function createFormulaBindingBookkeeping(args: CreateEngineFormulaBindingServiceArgs) {
  const resolvedCompiledCache = new Map<string, ParsedCompiledFormula>()
  const formulaMemberCounts = createFormulaBindingMemberCounts()
  const formulaSheetIndex = createFormulaBindingSheetIndex()
  const formulaFamilyShapeKeyCache: FormulaBindingFamilyShapeKeyCache = new Map()
  const {
    rebuildFormulaInstancesNow,
    recordFormulaInstanceNow,
    registerFormulaFamilyNow: registerFormulaFamilyInStoreNow,
  } = createFormulaBindingInstanceTracker({
    serviceArgs: args,
    formulaFamilyShapeKeyCache,
  })
  const formulaInstanceTableRebuild = createFormulaBindingInstanceTableRebuildController({
    formulaInstances: args.formulaInstances,
    rebuildFormulaInstancesNow,
    recordFormulaInstanceNow,
  })
  const formulaFamilyIndex = createFormulaBindingFamilyIndexController({
    formulaFamilies: args.formulaFamilies,
    formulaFamilyShapeKeyCache,
    registerFormulaFamilyInStoreNow,
    countFormulaSheetMembersNow: (sheetId) => formulaMemberCounts.countSheetMembers(sheetId),
    rebuildFormulaFamilyIndexNow: () =>
      rebuildDeferredFormulaFamilyIndex({
        state: args.state,
        store: args.formulaFamilies,
        shapeKeyCache: formulaFamilyShapeKeyCache,
      }),
  })

  return {
    resolvedCompiledCache,
    formulaMemberCounts,
    formulaSheetIndex,
    formulaFamilyIndex,
    formulaInstanceTableRebuild,
    recordFormulaInstanceNow,
    registerFormulaFamilyNow: formulaFamilyIndex.registerFormulaFamilyNow,
    clearFormulaBookkeepingNow: (): void => {
      resolvedCompiledCache.clear()
      formulaMemberCounts.clear()
      formulaSheetIndex.clear()
      formulaFamilyIndex.clearNow()
      formulaInstanceTableRebuild.clearRebuildNow()
    },
    updateVolatileFormulaIndex: (cellIndex: number, formula: RuntimeFormula | undefined): void => {
      updateFormulaBindingVolatileIndex(args.volatileFormulaCells, cellIndex, formula)
    },
    trackFormulaSheetIndexes: (cellIndex: number, ownerSheetName: string, compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'>): void => {
      formulaSheetIndex.trackFormula(cellIndex, ownerSheetName, compiled)
    },
    trackFormulaSheetOwnerRun: (ownerSheetName: string, cellIndices: readonly number[] | Uint32Array): void => {
      formulaSheetIndex.trackFormulaOwnerRun(ownerSheetName, cellIndices)
    },
    untrackFormulaSheetIndexes: (
      cellIndex: number,
      ownerSheetName: string | undefined,
      compiled: Pick<CompiledFormula, 'deps' | 'parsedDeps'> | undefined,
    ): void => {
      formulaSheetIndex.untrackFormula(cellIndex, ownerSheetName, compiled)
    },
  }
}
