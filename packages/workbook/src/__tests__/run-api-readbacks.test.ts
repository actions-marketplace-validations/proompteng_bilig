import { describe, expect, it, vi } from 'vitest'
import {
  defineModel,
  describeRunResult,
  findRange,
  formula,
  runWorkbookAction,
  runWorkbookPlan,
  verifyWorkbookReadbacks,
  type WorkbookActionPlan,
  type WorkbookModel,
  type WorkbookRunAdapter,
  type WorkbookRunApplyResult,
} from '../index.js'

function valueModel(): WorkbookModel<{ readonly output: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'run-value-model',

    find(workbook) {
      return {
        output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
      }
    },

    actions: {
      write({ refs, workbook }) {
        workbook.writeValue(refs.output, 12)
        workbook.check.valueEquals(refs.output, 12)
      },
    },
  })
}

function first<T>(values: readonly T[]): T {
  const [value] = values
  if (value === undefined) {
    throw new Error('expected at least one value')
  }
  return value
}

function applied<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRunApplyResult {
  return {
    status: 'applied',
    previewOps: plan.ops,
    appliedOps: plan.ops,
  }
}
describe('@bilig/workbook run api readback and async boundaries', () => {
  it('rejects malformed generic check verifier output', async () => {
    const model = defineModel({
      name: 'run-malformed-check-proof-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },

      actions: {
        inspect({ refs }) {
          void refs.result
        },
      },
    })

    await expect(
      runWorkbookAction(model, 'inspect', {
        apply: applied,
        verifyChecks: () => [],
      }),
    ).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned 0 checks for 1 planned checks',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })

    await expect(
      runWorkbookAction(model, 'inspect', {
        apply: applied,
        verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, message: 'Changed message' })),
      }),
    ).resolves.toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier changed the check contract at index 0 for exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('rejects in-place generic check verifier contract mutations', async () => {
    const model = defineModel({
      name: 'run-mutating-check-proof-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },

      actions: {
        inspect({ refs }) {
          void refs.result
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks(checks) {
        Object.defineProperty(first(checks), 'message', { value: 'Changed message' })
        return checks
      },
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier changed the check contract at index 0 for exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('returns failed when the generic check verifier throws', async () => {
    const model = defineModel({
      name: 'run-throwing-check-proof-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.result)]
      },

      actions: {
        inspect({ refs }) {
          void refs.result
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks() {
        throw new Error('check backend unavailable')
      },
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_verification_failed',
          message: 'check backend unavailable',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('does not apply when action planning fails', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'missing', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model run-value-model does not define action missing',
        },
      ],
      checks: [],
    })
  })

  it('does not apply when static plan verification fails', async () => {
    const hidden = findRange({ sheetName: 'Sheet1', address: 'Z9' })
    const model = defineModel({
      name: 'invalid-run-model',

      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          workbook.writeFormula(refs.output, formula.raw('Sheet1!Z9', { inputs: [hidden] }))
        },
      },
    })
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookAction(model, 'calculate', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'formula_input_not_resolved',
          message: 'Sheet1!Z9 is used as a formula input but is missing from refsUsed',
        },
      ],
      checks: [],
    })
  })

  it('does not apply when a plan contains already-proved checks', async () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({ status: 'applied' }))

    const result = await runWorkbookPlan(
      {
        modelName: 'pre-proved-run-plan',
        actionName: 'inspect',
        refs: { target },
        refsUsed: [target],
        commands: [],
        ops: [],
        changed: [],
        checks: [
          {
            status: 'passed',
            kind: 'exists',
            target,
            message: 'Sheet1!C2 exists',
          },
        ],
      },
      { apply },
    )

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_status_not_planned',
          message: 'Sheet1!C2 check exists must start planned before runtime proof',
        },
      ],
      checks: [
        {
          status: 'passed',
          kind: 'exists',
          target,
          message: 'Sheet1!C2 exists',
        },
      ],
    })
  })

  it('returns failed when the adapter apply step fails', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected the plan',
            path: 'adapter.apply',
            issueCode: 'runtime_refusal',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected the plan',
          path: 'adapter.apply',
          issueCode: 'runtime_refusal',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('does not report changed when failed apply proof says no ops were applied', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'failed',
        previewOps: plan.ops,
        appliedOps: [],
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected before writing',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected before writing',
        },
      ],
      apply: {
        matched: false,
        appliedOps: [],
      },
      changed: [],
    })
  })

  it('preserves changed proof when failed apply reports actual applied ops', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'failed',
        previewOps: plan.ops,
        appliedOps: plan.ops,
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime failed after writing',
          },
        ],
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime failed after writing',
        },
      ],
      apply: {
        matched: true,
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
    })
  })

  it('describes failed run results as JSON-safe errors and checks', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'failed',
        errors: [
          {
            code: 'runtime_rejected',
            message: 'runtime rejected the plan',
            path: 'adapter.apply',
            issueCode: 'runtime_refusal',
          },
        ],
      }),
      read: () => [],
    })
    const described = describeRunResult(result)

    expect(Object.isFrozen(result)).toBe(true)
    if (result.status !== 'failed') {
      throw new Error('expected failed result')
    }
    expect(Object.isFrozen(result.errors)).toBe(true)
    expect(Object.isFrozen(result.errors[0])).toBe(true)
    expect(Object.isFrozen(result.apply)).toBe(true)
    expect(Object.isFrozen(result.changed)).toBe(true)
    expect(Object.isFrozen(result.checks)).toBe(true)
    expect(Object.isFrozen(result.checks[0])).toBe(true)
    expect(Object.isFrozen(result.checks[0]?.target)).toBe(true)
    expect(described).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'runtime rejected the plan',
          path: 'adapter.apply',
          issueCode: 'runtime_refusal',
        },
      ],
      apply: {
        matched: null,
      },
      changed: [],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: {
            kind: 'range',
            id: 'range_Sheet1_B2_B2',
            label: 'Sheet1!B2',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'B2',
              endAddress: 'B2',
            },
          },
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
    expect(Object.isFrozen(described)).toBe(true)
    expect(Object.isFrozen(described.errors)).toBe(true)
    expect(Object.isFrozen(described.errors[0])).toBe(true)
    expect(Object.isFrozen(described.apply)).toBe(true)
    expect(Object.isFrozen(described.checks)).toBe(true)
    expect(Object.isFrozen(described.checks[0])).toBe(true)
    expect(Object.isFrozen(described.checks[0]?.target)).toBe(true)
  })

  it('fails before apply when expected readback requires a missing reader', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>(applied)

    const result = await runWorkbookAction(model, 'write', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'adapter_missing_capability',
          message: 'Adapter is missing read for read',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
    })
  })

  it('returns failed when value readback mismatches', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        ...applied(plan),
        undo: { id: 'undo-mismatch' },
      }),
      read: (targets) => [
        {
          target: first(targets),
          value: 13,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'value_mismatch',
          message: 'Sheet1!B2 expected value 12 but read 13',
        },
      ],
      apply: {
        matched: true,
      },
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        {
          status: 'failed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      undo: { id: 'undo-mismatch' },
    })
  })

  it('returns failed when formula readback mismatches', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
            labels: [],
          },
        },
      ],
      [{ target, formula: '=A2+B3' }],
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [
        {
          status: 'failed',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
            labels: [],
          },
        },
      ],
      issues: [
        {
          code: 'formula_mismatch',
          check: expect.objectContaining({ kind: 'formulaEquals' }),
          target,
          expected: 'A2+B2',
          actual: 'A2+B3',
          message: 'Sheet1!C2 expected formula A2+B2 but read A2+B3',
        },
      ],
    })
    expect(Object.isFrozen(verification)).toBe(true)
    expect(Object.isFrozen(verification.checks)).toBe(true)
    expect(Object.isFrozen(verification.checks[0])).toBe(true)
    expect(Object.isFrozen(verification.checks[0]?.expectation)).toBe(true)
    expect(Object.isFrozen(verification.issues)).toBe(true)
    expect(Object.isFrozen(verification.issues[0])).toBe(true)
    expect(Object.isFrozen(verification.issues[0]?.target)).toBe(true)
  })

  it('returns failed when formula readback proof is not parseable', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'formulaEquals',
          target,
          message: 'Sheet1!C2 formula equals A2+B2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'A2+B2',
            inputs: [],
            labels: [],
          },
        },
      ],
      [{ target, formula: '=' }],
    )

    expect(verification.status).toBe('failed')
    expect(verification.checks).toEqual([
      {
        status: 'planned',
        kind: 'formulaEquals',
        target,
        message: 'Sheet1!C2 formula equals A2+B2',
        expectation: {
          kind: 'formulaEquals',
          formula: 'A2+B2',
          inputs: [],
          labels: [],
        },
      },
    ])
    expect(verification.issues).toEqual([
      {
        code: 'readback_invalid',
        message: 'Workbook formula proof at readbacks[0].formula cannot be empty',
      },
    ])
  })

  it('supports async apply and async read adapters', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: async (plan) => applied(plan),
      read: async (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(planned.status).toBe('done')
  })

  it('runs an already planned action without model access', async () => {
    const model = valueModel()
    const planned = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'failed' }),
    })

    expect(planned.status).toBe('failed')
    const planResult = await runWorkbookPlan(
      {
        modelName: 'empty-run-plan',
        actionName: 'noop',
        refs: {},
        refsUsed: [],
        commands: [],
        ops: [],
        changed: [],
        checks: [],
      },
      {
        apply: applied,
      },
    )

    expect(planResult).toMatchObject({
      status: 'done',
      changed: [],
      checks: [],
    })
  })
})
