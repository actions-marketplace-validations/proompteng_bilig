import { describe, expect, it } from 'vitest'
import {
  defineModel,
  formula,
  prepareWorkbookAction,
  workbookActionCommandDigest,
  workbookPlanId,
  type EngineOp,
  type WorkbookActionPlan,
  type WorkbookRunAdapter,
  type WorkbookRunApplyCommandReceipt,
} from '../index.js'
import { assertWorkbookRunAdapter, checkWorkbookRunAdapter } from '../testing.js'

function rangeRef(label: string, startAddress: string, endAddress = startAddress) {
  return {
    kind: 'range' as const,
    id: `range_${label}`,
    label,
    range: {
      sheetName: 'Resolved',
      startAddress,
      endAddress,
    },
  }
}

function labelSource(labelName: string): string {
  if (labelName === 'input') {
    return 'Resolved!A1'
  }
  if (labelName === 'factor') {
    return 'Resolved!B1'
  }
  return labelName
}

function model() {
  return defineModel({
    name: 'testing-adapter-model',
    find(workbook) {
      return {
        input: workbook.findName('input'),
        factor: workbook.findName('factor'),
        result: workbook.findName('result'),
      }
    },
    checks({ refs, workbook }) {
      const expected = formula.multiply(refs.input, refs.factor)
      return [workbook.check.formulaEquals(refs.result, expected)]
    },
    actions: {
      calculate({ refs, workbook }) {
        workbook.writeFormula(refs.result, formula.multiply(refs.input, refs.factor))
      },
    },
  })
}

function receipt<Refs>(plan: WorkbookActionPlan<Refs>, commandIndex: number, op: EngineOp): WorkbookRunApplyCommandReceipt {
  const command = plan.commands[commandIndex]
  if (command?.kind !== 'writeFormula') {
    throw new Error('expected formula command')
  }
  return {
    commandIndex,
    commandKind: command.kind,
    commandDigest: workbookActionCommandDigest(command),
    previewOps: [op],
    appliedOps: [op],
    resolvedRefs: {
      target: rangeRef('Resolved!C1', 'C1'),
      inputs: [rangeRef('Resolved!A1', 'A1'), rangeRef('Resolved!B1', 'B1')],
    },
    formulaLabels: command.labels.map((label) => ({ name: label.name, source: labelSource(label.name) })),
  }
}

function passingAdapter(): WorkbookRunAdapter<{ readonly refsUsed: ReturnType<typeof prepare>['plan']['refsUsed'] }> {
  return {
    apply(plan) {
      const op: EngineOp = {
        kind: 'setCellFormula',
        sheetName: 'Resolved',
        address: 'C1',
        formula: 'Resolved!A1*Resolved!B1',
      }
      return {
        status: 'applied',
        planId: workbookPlanId(plan),
        baseRevision: 4,
        revision: 5,
        previewOps: [op],
        appliedOps: [op],
        commandReceipts: [receipt(plan, 0, op)],
      }
    },
    read(targets) {
      return targets.map((target) => ({
        target,
        formula: 'Resolved!A1*Resolved!B1',
        formulaLabels: [
          { name: 'input', source: 'Resolved!A1' },
          { name: 'factor', source: 'Resolved!B1' },
        ],
      }))
    },
  }
}

function prepare() {
  const prepared = prepareWorkbookAction(model(), 'calculate')
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared fixture')
  }
  return prepared
}

function prepareNumberFormatClearCommand() {
  const prepared = prepareWorkbookAction(
    defineModel({
      name: 'testing-adapter-format-clear-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },
      actions: {
        format({ refs, workbook }) {
          workbook.format(refs.result, { numberFormat: null })
        },
      },
    }),
    'format',
  )
  if (prepared.status !== 'prepared') {
    throw new Error('expected prepared number format clear fixture')
  }
  return prepared
}
describe('@bilig/workbook testing api clear format receipts', () => {
  it('rejects number format clear receipts without a proven general format record', async () => {
    const prepared = prepareNumberFormatClearCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        formatId: 'format_unknown',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('failed')
    if (check.status !== 'failed') {
      throw new Error('adapter unexpectedly passed')
    }
    expect(check.issues).toEqual([
      {
        code: 'runtime_rejected',
        path: 'result',
        message:
          'Workbook action testing-adapter-format-clear-model.format returned invalid command receipts: commandReceipts[0].previewOps do not match the planned command',
      },
    ])
  })

  it('accepts number format clear receipts with an explicit general format record', async () => {
    const prepared = prepareNumberFormatClearCommand()
    const command = prepared.plan.commands[0]
    if (command?.kind !== 'format') {
      throw new Error('expected format command')
    }
    const ops: readonly EngineOp[] = [
      {
        kind: 'upsertCellNumberFormat',
        format: { id: 'format-0', code: 'general', kind: 'general' },
      },
      {
        kind: 'setFormatRange',
        range: {
          sheetName: 'Resolved',
          startAddress: 'C1',
          endAddress: 'C2',
        },
        formatId: 'format-0',
      },
    ]

    const check = await checkWorkbookRunAdapter(prepared.planData, {
      apply(plan) {
        return {
          status: 'applied',
          planId: workbookPlanId(plan),
          baseRevision: 4,
          revision: 5,
          previewOps: ops,
          appliedOps: ops,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: command.kind,
              commandDigest: workbookActionCommandDigest(command),
              previewOps: ops,
              appliedOps: ops,
              resolvedRefs: {
                target: rangeRef('Resolved!C1:C2', 'C1', 'C2'),
              },
            },
          ],
        }
      },
      verifyChecks(checks) {
        return checks.map((entry) => ({
          ...entry,
          status: 'passed' as const,
          proof: { source: 'adapter' },
        }))
      },
    })

    expect(check.status).toBe('passed')
  })

  it('throws from the assertion helper with the first issue message', async () => {
    const prepared = prepare()
    const adapter = passingAdapter()
    await expect(assertWorkbookRunAdapter(prepared.planData, { apply: (plan) => adapter.apply?.(plan) })).rejects.toThrow(
      'Adapter is missing read for read',
    )

    await expect(assertWorkbookRunAdapter(prepared.planData, passingAdapter())).resolves.toMatchObject({
      status: 'done',
    })
  })
})
