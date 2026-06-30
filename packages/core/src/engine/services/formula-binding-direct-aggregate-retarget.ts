import type { StructuralAxisTransform } from '@bilig/formula'
import { rewriteDirectAggregateDescriptorForStructuralTransform } from './formula-binding-direct-descriptors.js'
import type { CreateEngineFormulaBindingServiceArgs } from './formula-binding-service-types.js'

export interface DirectAggregateStructuralRetargetInput {
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly preservesValue: boolean
}

export function retargetDirectAggregateFormulaForStructuralTransform(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
  readonly preservesValue: boolean
}): boolean {
  const existing = args.serviceArgs.state.formulas.get(args.cellIndex)
  if (!existing?.directAggregate) {
    return false
  }
  const previousDirectAggregate = existing.directAggregate
  const nextDirectAggregate = rewriteDirectAggregateDescriptorForStructuralTransform({
    descriptor: previousDirectAggregate,
    targetSheetName: args.targetSheetName,
    transform: args.transform,
    regionGraph: args.serviceArgs.regionGraph,
  })
  if (!nextDirectAggregate) {
    return false
  }
  existing.directAggregate = nextDirectAggregate
  existing.structuralSourceTransform = {
    ownerSheetName: args.ownerSheetName,
    targetSheetName: args.targetSheetName,
    transform: args.transform,
    preservesValue: args.preservesValue,
  }
  args.serviceArgs.regionGraph.replaceSingleFormulaSubscription(
    args.cellIndex,
    previousDirectAggregate.regionId,
    nextDirectAggregate.regionId,
  )
  return true
}

export function retargetDirectAggregateFormulasForStructuralTransform(args: {
  readonly serviceArgs: CreateEngineFormulaBindingServiceArgs
  readonly inputs: readonly DirectAggregateStructuralRetargetInput[]
  readonly targetSheetName: string
  readonly transform: StructuralAxisTransform
}): readonly number[] {
  if (args.inputs.length === 0) {
    return []
  }
  const retargetedCellIndices: number[] = []
  const replacements: Array<{ formulaCellIndex: number; previousRegionId: number; nextRegionId: number }> = []
  for (let index = 0; index < args.inputs.length; index += 1) {
    const { cellIndex, ownerSheetName, preservesValue } = args.inputs[index]!
    const existing = args.serviceArgs.state.formulas.get(cellIndex)
    if (!existing?.directAggregate) {
      continue
    }
    const previousDirectAggregate = existing.directAggregate
    const nextDirectAggregate = rewriteDirectAggregateDescriptorForStructuralTransform({
      descriptor: previousDirectAggregate,
      targetSheetName: args.targetSheetName,
      transform: args.transform,
      regionGraph: args.serviceArgs.regionGraph,
    })
    if (!nextDirectAggregate) {
      continue
    }
    existing.directAggregate = nextDirectAggregate
    existing.structuralSourceTransform = {
      ownerSheetName,
      targetSheetName: args.targetSheetName,
      transform: args.transform,
      preservesValue,
    }
    retargetedCellIndices.push(cellIndex)
    replacements.push({
      formulaCellIndex: cellIndex,
      previousRegionId: previousDirectAggregate.regionId,
      nextRegionId: nextDirectAggregate.regionId,
    })
  }
  args.serviceArgs.regionGraph.replaceSingleFormulaSubscriptions(replacements)
  return retargetedCellIndices
}
