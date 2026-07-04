import { describe, expect, it, vi } from 'vitest'
import {
  checkWorkbookRunResultDescription,
  defineModel,
  describePlanResult,
  describeRunResult,
  findRange,
  formula,
  isWorkbookRunErrorCode,
  isWorkbookRunResultDescription,
  runWorkbookAction,
  verifyWorkbookReadbacks,
  workbookRunErrorCodes,
  type WorkbookActionPlan,
  type WorkbookCheckResult,
  type WorkbookModel,
  type WorkbookRunAdapter,
  type WorkbookRunApplyResult,
  type WorkbookRunOptions,
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

function accessorArray<T>(getter: () => T): readonly T[] {
  const values: T[] = []
  Object.defineProperty(values, '0', {
    configurable: true,
    enumerable: true,
    get: getter,
  })
  return values
}

function runDescriptionWithNoopEffect(commandKind: string, effect: Record<string, unknown>): Record<string, unknown> {
  return {
    status: 'done',
    apply: {
      matched: true,
      commandReceipts: [
        {
          commandIndex: 0,
          commandKind,
          commandDigest: `bilig-command-v1:${commandKind}`,
          previewOps: [],
          appliedOps: [],
          noop: {
            reason: 'already_satisfied',
            proof: {
              source: 'test',
              evidence: 'adapter_zero_ops',
              commandKind,
              commandDigest: `bilig-command-v1:${commandKind}`,
              opCount: 0,
              effect,
            },
          },
        },
      ],
    },
    changed: [],
    checks: [],
  }
}

function applied<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRunApplyResult {
  return {
    status: 'applied',
    previewOps: plan.ops,
    appliedOps: plan.ops,
  }
}

function inheritedApplyResult<Refs>(plan: WorkbookActionPlan<Refs>): WorkbookRunApplyResult {
  const result: WorkbookRunApplyResult = applied(plan)
  Object.setPrototypeOf(result, {
    status: result.status,
    previewOps: result.previewOps,
    appliedOps: result.appliedOps,
  })
  Reflect.deleteProperty(result, 'status')
  Reflect.deleteProperty(result, 'previewOps')
  Reflect.deleteProperty(result, 'appliedOps')
  return result
}

function inheritedVerifiedCheck(checkResult: WorkbookCheckResult): WorkbookCheckResult {
  const result: WorkbookCheckResult = {
    ...checkResult,
    status: 'passed',
  }
  Object.setPrototypeOf(result, { ...result })
  Reflect.deleteProperty(result, 'status')
  Reflect.deleteProperty(result, 'kind')
  Reflect.deleteProperty(result, 'target')
  Reflect.deleteProperty(result, 'refs')
  Reflect.deleteProperty(result, 'message')
  Reflect.deleteProperty(result, 'expectation')
  Reflect.deleteProperty(result, 'proof')
  return result
}

describe('@bilig/workbook run api execution and adapter boundaries', () => {
  it('exports stable inspectable run error codes', () => {
    expect(Object.isFrozen(workbookRunErrorCodes)).toBe(true)
    expect(workbookRunErrorCodes).toContain('invalid_model')
    expect(workbookRunErrorCodes).toContain('invalid_plan')
    expect(workbookRunErrorCodes).toContain('invalid_plan_data')
    expect(workbookRunErrorCodes).toContain('invalid_action_name')
    expect(workbookRunErrorCodes).toContain('action_not_found')
    expect(workbookRunErrorCodes).toContain('invalid_action_input')
    expect(workbookRunErrorCodes).toContain('ref_not_in_refs')
    expect(workbookRunErrorCodes).toContain('formula_input_not_resolved')
    expect(workbookRunErrorCodes).toContain('apply_not_verified')
    expect(workbookRunErrorCodes).toContain('apply_mismatch')
    expect(workbookRunErrorCodes).toContain('readback_invalid')
    expect(workbookRunErrorCodes).toContain('readback_duplicate')
    expect(workbookRunErrorCodes).toContain('readback_missing')
    expect(workbookRunErrorCodes).toContain('adapter_missing_capability')
    expect(workbookRunErrorCodes).toContain('runtime_rejected')
    expect(new Set(workbookRunErrorCodes).size).toBe(workbookRunErrorCodes.length)
    expect(isWorkbookRunErrorCode('check_not_verified')).toBe(true)
    expect(isWorkbookRunErrorCode('invalid_run_options')).toBe(true)
    expect(isWorkbookRunErrorCode('custom_runtime_error')).toBe(false)
  })

  it('fails closed for accessor-backed run options without invoking getters', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>(applied)
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      { target: first(targets), value: 12 },
    ])
    let strictGetterInvoked = false
    const options: WorkbookRunOptions = {}
    Object.defineProperty(options, 'strict', {
      enumerable: true,
      get() {
        strictGetterInvoked = true
        throw new Error('strict getter must not run')
      },
    })

    const result = await runWorkbookAction(model, 'write', { apply, read }, undefined, options)

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run option strict must be a data property',
          path: 'options.strict',
          issueCode: 'invalid_run_options',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
    expect(strictGetterInvoked).toBe(false)
    expect(apply).not.toHaveBeenCalled()
    expect(read).not.toHaveBeenCalled()
  })

  it('fails closed for non-plain or unknown run options before runtime work starts', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>(applied)
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      { target: first(targets), value: 12 },
    ])

    class RunOptions {
      strict = true
    }

    const classResult = await Reflect.apply(runWorkbookAction, undefined, [model, 'write', { apply, read }, undefined, new RunOptions()])
    expect(classResult).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run options must be a plain object',
          path: 'options',
          issueCode: 'invalid_run_options',
        },
      ],
    })

    const typoResult = await Reflect.apply(runWorkbookAction, undefined, [
      model,
      'write',
      { apply, read },
      undefined,
      {
        requireApplyproof: true,
      },
    ])
    expect(typoResult).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run option requireApplyproof is unknown',
          path: 'options.requireApplyproof',
          issueCode: 'invalid_run_options',
        },
      ],
    })
    expect(apply).not.toHaveBeenCalled()
    expect(read).not.toHaveBeenCalled()
  })

  it('fails closed for symbol-keyed run options before runtime work starts', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>(applied)
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      { target: first(targets), value: 12 },
    ])
    const symbolKey = Symbol('hidden')
    const options: Record<string | symbol, unknown> = {
      strict: true,
      [symbolKey]: true,
    }

    const result = await Reflect.apply(runWorkbookAction, undefined, [model, 'write', { apply, read }, undefined, options])

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_run_options',
          message: 'Workbook run option Symbol(hidden) is unknown',
          path: 'options.Symbol(hidden)',
          issueCode: 'invalid_run_options',
        },
      ],
    })
    expect(apply).not.toHaveBeenCalled()
    expect(read).not.toHaveBeenCalled()
  })

  it('fails closed for accessor-backed runtime adapter methods without invoking getters', async () => {
    const model = valueModel()
    let applyGetterInvoked = false
    const adapter: WorkbookRunAdapter<{ readonly output: ReturnType<typeof findRange> }> = {
      read: () => [],
    }
    Object.defineProperty(adapter, 'apply', {
      enumerable: true,
      get() {
        applyGetterInvoked = true
        throw new Error('apply getter must not run')
      },
    })

    const result = await runWorkbookAction(model, 'write', adapter)

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'adapter_missing_capability',
          message: 'Adapter is missing apply for writeValue',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
    expect(applyGetterInvoked).toBe(false)
  })

  it('plans, verifies, applies, reads back, and returns done for value checks', async () => {
    const model = valueModel()
    const apply = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['apply']>((plan) => ({
      status: 'applied',
      previewOps: plan.ops,
      appliedOps: plan.ops,
      undo: { id: 'undo-1' },
    }))
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>((targets) => [
      {
        target: first(targets),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(model, 'write', { apply, read })

    expect(apply).toHaveBeenCalledTimes(1)
    expect(read).toHaveBeenCalledTimes(1)
    expect(read.mock.calls[0]?.[0]).toEqual([expect.objectContaining({ label: 'Sheet1!B2' })])
    expect(result).toMatchObject({
      status: 'done',
      apply: {
        matched: true,
        previewOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 12,
          },
        ],
        appliedOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 12,
          },
        ],
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
          status: 'passed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
          proof: {
            source: 'readback',
            value: 12,
          },
        },
      ],
      undo: { id: 'undo-1' },
    })
    expect(describeRunResult(result).checks[0]?.proof).toEqual({
      source: 'readback',
      value: 12,
    })
    expect(Object.isFrozen(result.checks[0]?.proof)).toBe(true)
  })

  it('reports unverified apply proof when an adapter does not return preview and applied ops', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({ status: 'applied' }),
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'done',
      apply: {
        matched: null,
      },
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
        },
        {
          kind: 'apply',
          message: 'Adapter did not return commandReceipts, so planned commands are not bound to materialized ops',
        },
      ],
    })
  })

  it('rejects inherited adapter apply result fields', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: inheritedApplyResult,
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned an invalid apply result',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
  })

  it('rejects whitespace-padded undo ids before exposing run proof', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: () => ({
        status: 'applied',
        undo: { id: ' undo-1 ' },
      }),
      read: () => [],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'runtime_rejected',
          message: 'Workbook action run-value-model.write returned invalid undo metadata',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'valueEquals' })],
    })
    expect(result).not.toHaveProperty('undo')
  })

  it('can require apply proof before readback and check verification', async () => {
    const model = valueModel()
    const read = vi.fn<Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read']>(() => [
      {
        target: findRange({ sheetName: 'Sheet1', address: 'B2' }),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(
      model,
      'write',
      {
        apply: () => ({ status: 'applied' }),
        read,
      },
      undefined,
      { requireApplyProof: true },
    )

    expect(read).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'apply_not_verified',
          message: 'Adapter did not return both previewOps and appliedOps',
        },
      ],
      apply: {
        matched: null,
      },
      unverified: [
        {
          kind: 'apply',
          message: 'Adapter did not return both previewOps and appliedOps, so apply match is unverified',
        },
        {
          kind: 'apply',
          message: 'Adapter did not return commandReceipts, so planned commands are not bound to materialized ops',
        },
      ],
    })
  })

  it('fails when an adapter applies ops that do not match its preview', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: (plan) => ({
        status: 'applied',
        previewOps: plan.ops,
        appliedOps: [
          {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 13,
          },
        ],
      }),
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'apply_mismatch',
          message: 'Adapter applied ops do not match its preview ops',
        },
      ],
      apply: {
        matched: false,
      },
    })
  })

  it('describes successful run results without leaking ref helper functions', async () => {
    const model = defineModel({
      name: 'run-description-model',

      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs', headers: ['Amount'] }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },

      actions: {
        inspect({ refs }) {
          void refs.table
        },
      },
    })

    const apply = vi.fn<Required<WorkbookRunAdapter>['apply']>(() => ({
      status: 'applied',
      undo: {
        id: 'undo-1',
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }],
      },
    }))

    const result = await runWorkbookAction(model, 'inspect', {
      apply,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'passed' })),
    })
    const described = describeRunResult(result)

    expect(apply).not.toHaveBeenCalled()
    expect(described).toEqual({
      status: 'done',
      changed: [],
      checks: [
        {
          status: 'passed',
          kind: 'exists',
          target: {
            kind: 'table',
            id: 'table_Inputs_Amount',
            label: 'Inputs',
            name: 'Inputs',
            headers: ['Amount'],
          },
          message: 'Inputs exists',
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
    expect(isWorkbookRunResultDescription(described)).toBe(true)
    expect(isWorkbookRunResultDescription({ ...described, status: 'done', errors: [] })).toBe(false)
    expect(isWorkbookRunResultDescription({ ...described, status: 'failed', errors: [() => undefined] })).toBe(false)
    expect(Object.isFrozen(described)).toBe(true)
    expect(Object.isFrozen(described.changed)).toBe(true)
    expect(Object.isFrozen(described.checks)).toBe(true)
    expect(Object.isFrozen(described.checks[0])).toBe(true)
    expect(Object.isFrozen(described.checks[0]?.target)).toBe(true)
  })

  it('checks run result descriptions with structured frozen issues', async () => {
    const model = valueModel()
    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })
    const described = describeRunResult(result)
    const valid = checkWorkbookRunResultDescription(described)

    expect(valid).toEqual({
      status: 'valid',
      description: described,
      issues: [],
    })
    expect(Object.isFrozen(valid)).toBe(true)
    if (valid.status !== 'valid') {
      throw new Error('expected valid description')
    }
    expect(Object.isFrozen(valid.description)).toBe(true)
    expect(Object.isFrozen(valid.issues)).toBe(true)

    const invalid = checkWorkbookRunResultDescription({ ...described, errors: [] })

    expect(invalid).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'unexpected_field',
          path: 'errors',
          message: 'Workbook run result description must not include errors',
        },
      ],
    })
    expect(Object.isFrozen(invalid)).toBe(true)
    expect(Object.isFrozen(invalid.issues)).toBe(true)
    expect(Object.isFrozen(invalid.issues[0])).toBe(true)

    expect(
      checkWorkbookRunResultDescription({
        status: 'done',
        changed: [],
        checks: [{}],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'checks[0].status',
          message: 'Workbook run result description checks[0].status must be planned, passed, or failed',
        },
        {
          code: 'missing_field',
          path: 'checks[0].kind',
          message: 'Workbook run result description checks[0].kind is required',
        },
        {
          code: 'missing_field',
          path: 'checks[0].message',
          message: 'Workbook run result description checks[0].message is required',
        },
      ],
    })
  })

  it('rejects unsupported description result statuses instead of relabeling them', () => {
    expect(() => Reflect.apply(describePlanResult, undefined, [{ status: 'unknown' }])).toThrowError(
      'Unsupported workbook plan result status: unknown',
    )
    expect(() => Reflect.apply(describeRunResult, undefined, [{ status: 'unknown' }])).toThrowError(
      'Unsupported workbook run result status: unknown',
    )
  })

  it('rejects hidden behavior in run result descriptions without invoking getters', async () => {
    const model = valueModel()
    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: (targets) => [
        {
          target: first(targets),
          value: 12,
        },
      ],
    })
    if (result.status !== 'done') {
      throw new Error('expected done result')
    }
    let getterInvoked = false
    const badResult: typeof result = {
      ...result,
      checks: accessorArray<(typeof result.checks)[number]>(() => {
        getterInvoked = true
        throw new Error('checks getter should not run')
      }),
    }

    expect(() => describeRunResult(badResult)).toThrowError('Workbook description result.checks[0] must be a data property')
    expect(getterInvoked).toBe(false)
  })

  it('reports nested accessor paths in run result description checks without invoking getters', () => {
    let getterInvoked = false
    const error = { code: 'apply_failed', message: 'apply failed' }
    Object.defineProperty(error, 'message', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('message getter should not run')
      },
    })

    const checked = checkWorkbookRunResultDescription({
      status: 'failed',
      errors: [error],
      changed: [],
      checks: [],
    })

    expect(checked).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'errors[0].message',
          message: 'Workbook run result description errors[0].message must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('requires run-result description revisions and command indexes to be non-negative safe integers', () => {
    expect(
      checkWorkbookRunResultDescription({
        status: 'done',
        apply: {
          matched: true,
          baseRevision: 1.5,
          revision: -1,
          commandReceipts: [
            {
              commandIndex: 0.5,
              commandKind: 'writeFormula',
              commandDigest: 'bilig-command-v1:test',
              previewOps: [],
              appliedOps: [],
            },
          ],
        },
        changed: [],
        checks: [],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'apply.baseRevision',
          message: 'Workbook run result description apply.baseRevision must be a non-negative safe integer',
        },
        {
          code: 'invalid_field',
          path: 'apply.revision',
          message: 'Workbook run result description apply.revision must be a non-negative safe integer',
        },
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].commandIndex',
          message: 'Workbook run result description apply.commandReceipts[0].commandIndex must be a non-negative safe integer',
        },
      ],
    })
  })

  it('requires run-result description no-op receipts to preserve proof', () => {
    expect(
      checkWorkbookRunResultDescription({
        status: 'done',
        apply: {
          matched: true,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: 'writeValue',
              commandDigest: 'bilig-command-v1:test',
              previewOps: [],
              appliedOps: [],
              noop: {
                reason: 'already_satisfied',
              },
            },
          ],
        },
        changed: [],
        checks: [],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_field',
          path: 'apply.commandReceipts[0].noop.proof',
          message: 'Workbook run result description apply.commandReceipts[0].noop.proof is required',
        },
      ],
    })
  })

  it('requires run-result description no-op proof to stay bound to the receipt', () => {
    const mutationOp = { kind: 'setCellValue', sheetId: 'sheet-1', row: 0, col: 0, value: 12 }

    expect(
      checkWorkbookRunResultDescription({
        status: 'done',
        apply: {
          matched: true,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: 'writeValue',
              commandDigest: 'bilig-command-v1:actual',
              previewOps: [{ ...mutationOp }],
              appliedOps: [{ ...mutationOp }],
              noop: {
                reason: 'already_satisfied',
                proof: {
                  source: 'test',
                  evidence: 'adapter_zero_ops',
                  commandKind: 'format',
                  commandDigest: 'bilig-command-v1:wrong',
                  opCount: 0,
                  effect: {
                    kind: 'format',
                    numberFormat: '0.00',
                  },
                },
              },
            },
          ],
        },
        changed: [],
        checks: [],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].noop.proof.commandKind',
          message: 'Workbook run result description apply.commandReceipts[0].noop.proof.commandKind must match receipt commandKind',
        },
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].noop.proof.commandDigest',
          message: 'Workbook run result description apply.commandReceipts[0].noop.proof.commandDigest must match receipt commandDigest',
        },
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].noop.proof.effect.kind',
          message: 'Workbook run result description apply.commandReceipts[0].noop.proof.effect.kind must match receipt commandKind',
        },
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].previewOps',
          message: 'Workbook run result description apply.commandReceipts[0].previewOps must be empty when noop is present',
        },
        {
          code: 'invalid_field',
          path: 'apply.commandReceipts[0].appliedOps',
          message: 'Workbook run result description apply.commandReceipts[0].appliedOps must be empty when noop is present',
        },
      ],
    })
  })

  it('requires low-level no-op descriptions to preserve full op effect data', () => {
    expect(
      checkWorkbookRunResultDescription({
        status: 'done',
        apply: {
          matched: true,
          commandReceipts: [
            {
              commandIndex: 0,
              commandKind: 'op',
              commandDigest: 'bilig-command-v1:op',
              previewOps: [],
              appliedOps: [],
              noop: {
                reason: 'already_satisfied',
                proof: {
                  source: 'test',
                  evidence: 'adapter_zero_ops',
                  commandKind: 'op',
                  commandDigest: 'bilig-command-v1:op',
                  opCount: 0,
                  effect: {
                    kind: 'op',
                    opKind: 'setCellValue',
                  },
                },
              },
            },
          ],
        },
        changed: [],
        checks: [],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'missing_field',
          path: 'apply.commandReceipts[0].noop.proof.effect.op',
          message: 'Workbook run result description apply.commandReceipts[0].noop.proof.effect.op is required',
        },
      ],
    })
  })

  it('requires no-op descriptions to preserve command-specific effect fields', () => {
    const effectPath = 'apply.commandReceipts[0].noop.proof.effect'
    const cases = [
      {
        commandKind: 'writeValue',
        effect: { kind: 'writeValue' },
        code: 'missing_field',
        path: `${effectPath}.value`,
        message: `Workbook run result description ${effectPath}.value is required`,
      },
      {
        commandKind: 'writeFormula',
        effect: { kind: 'writeFormula' },
        code: 'missing_field',
        path: `${effectPath}.formula`,
        message: `Workbook run result description ${effectPath}.formula is required`,
      },
      {
        commandKind: 'clear',
        effect: { kind: 'clear' },
        code: 'missing_field',
        path: `${effectPath}.cleared`,
        message: `Workbook run result description ${effectPath}.cleared is required`,
      },
      {
        commandKind: 'format',
        effect: { kind: 'format' },
        code: 'missing_field',
        path: effectPath,
        message: `Workbook run result description ${effectPath} must include style or numberFormat`,
      },
      {
        commandKind: 'format',
        effect: { kind: 'format', style: {} },
        code: 'invalid_field',
        path: `${effectPath}.style`,
        message: `Workbook run result description ${effectPath}.style must request at least one style field`,
      },
      {
        commandKind: 'format',
        effect: { kind: 'format', style: { font: {} } },
        code: 'invalid_field',
        path: `${effectPath}.style`,
        message: `Workbook run result description ${effectPath}.style must request at least one style field`,
      },
      {
        commandKind: 'format',
        effect: { kind: 'format', style: { font: { bold: 'yes' } } },
        code: 'invalid_field',
        path: `${effectPath}.style`,
        message: `Workbook run result description ${effectPath}.style must be a valid cell style patch`,
      },
      {
        commandKind: 'op',
        effect: { kind: 'op', opKind: 'setCellValue', op: { kind: 'setCellValue' } },
        code: 'invalid_field',
        path: `${effectPath}.op`,
        message: `Workbook run result description ${effectPath}.op must be a valid WorkbookOp`,
      },
      {
        commandKind: 'op',
        effect: {
          kind: 'op',
          opKind: 'setCellFormula',
          op: { kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 12 },
        },
        code: 'invalid_field',
        path: `${effectPath}.opKind`,
        message: `Workbook run result description ${effectPath}.opKind must match effect.op.kind`,
      },
    ]

    for (const entry of cases) {
      expect(checkWorkbookRunResultDescription(runDescriptionWithNoopEffect(entry.commandKind, entry.effect))).toMatchObject({
        status: 'invalid',
        issues: expect.arrayContaining([
          {
            code: entry.code,
            path: entry.path,
            message: entry.message,
          },
        ]),
      })
    }
  })

  it('runs readback-only plans without requiring an apply adapter', async () => {
    const model = defineModel({
      name: 'run-readback-only-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      actions: {
        inspect({ refs, workbook }) {
          workbook.check.valueEquals(refs.result, 12)
        },
      },
    })
    const read = vi.fn<Required<WorkbookRunAdapter<{ result: ReturnType<typeof findRange> }>>['read']>((targets) => [
      {
        target: first(targets),
        value: 12,
      },
    ])

    const result = await runWorkbookAction(model, 'inspect', { read })

    expect(read).toHaveBeenCalledTimes(1)
    expect(result).toMatchObject({
      status: 'done',
      changed: [],
      checks: [
        {
          status: 'passed',
          kind: 'valueEquals',
          target: expect.objectContaining({ label: 'Sheet1!C2' }),
          message: 'Sheet1!C2 equals 12',
          proof: {
            source: 'readback',
            value: 12,
          },
        },
      ],
    })
    expect(result).not.toHaveProperty('apply')
    expect(Object.isFrozen(result)).toBe(true)
    expect(Object.isFrozen(result.changed)).toBe(true)
    expect(Object.isFrozen(result.checks)).toBe(true)
    expect(Object.isFrozen(result.checks[0])).toBe(true)
    expect(Object.isFrozen(result.checks[0]?.target)).toBe(true)
    expect(Object.isFrozen(result.checks[0]?.proof)).toBe(true)
  })

  it('passes formula readback checks with canonical formula proof', async () => {
    const model = defineModel({
      name: 'run-formula-model',

      find(workbook) {
        return {
          amount: workbook.findRange({ sheetName: 'Sheet1', address: 'A2' }),
          rate: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      actions: {
        calculate({ refs, workbook }) {
          const expected = formula.multiply(refs.amount, refs.rate)
          workbook.writeFormula(refs.result, expected)
          workbook.check.formulaEquals(refs.result, expected)
        },
      },
    })
    const plannedSource = '(Sheet1!A2)*(Sheet1!B2)'
    const canonicalSource = 'Sheet1!A2*Sheet1!B2'

    const result = await runWorkbookAction(model, 'calculate', {
      apply: applied,
      read: (targets) => [
        {
          target: first(targets),
          formula: '= (Sheet1!A2) * (Sheet1!B2)',
        },
      ],
    })

    expect(result.status).toBe('done')
    expect(result.checks).toEqual([
      {
        status: 'passed',
        kind: 'formulaEquals',
        target: expect.objectContaining({ label: 'Sheet1!C2' }),
        message: `Sheet1!C2 formula equals ${plannedSource}`,
        expectation: {
          kind: 'formulaEquals',
          formula: canonicalSource,
          inputs: [expect.objectContaining({ label: 'Sheet1!A2' }), expect.objectContaining({ label: 'Sheet1!B2' })],
          labels: [
            { name: 'Sheet1!A2', ref: expect.objectContaining({ label: 'Sheet1!A2' }) },
            { name: 'Sheet1!B2', ref: expect.objectContaining({ label: 'Sheet1!B2' }) },
          ],
        },
        proof: {
          source: 'readback',
          formula: canonicalSource,
          expectedFormula: canonicalSource,
          materializedFormula: canonicalSource,
        },
      },
    ])
  })

  it('passes formula readback checks with generic label materialization proof', () => {
    const amount = findRange({ sheetName: 'Sheet1', address: 'A2' })
    const result = findRange({ sheetName: 'Sheet1', address: 'C2' })
    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'formulaEquals',
          target: result,
          message: 'Sheet1!C2 formula equals amount_token*2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'amount_token*2',
            inputs: [amount],
            labels: [{ name: 'amount_token', ref: amount }],
          },
        },
      ],
      [
        {
          target: result,
          formula: 'Sheet1!A2*2',
          formulaLabels: [{ name: 'amount_token', source: 'Sheet1!A2' }],
        },
      ],
    )

    expect(verification).toEqual({
      status: 'passed',
      checks: [
        {
          status: 'passed',
          kind: 'formulaEquals',
          target: result,
          message: 'Sheet1!C2 formula equals amount_token*2',
          expectation: {
            kind: 'formulaEquals',
            formula: 'amount_token*2',
            inputs: [amount],
            labels: [{ name: 'amount_token', ref: amount }],
          },
          proof: {
            source: 'readback',
            formula: 'Sheet1!A2*2',
            expectedFormula: 'amount_token*2',
            materializedFormula: 'Sheet1!A2*2',
            formulaLabels: [{ name: 'amount_token', source: 'Sheet1!A2' }],
          },
        },
      ],
      issues: [],
    })
  })

  it('fails before apply when non-readback checks require a missing verifier', async () => {
    const model = defineModel({
      name: 'run-legacy-check-model',

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

    const apply = vi.fn<Required<WorkbookRunAdapter<{ result: ReturnType<typeof findRange> }>>['apply']>(applied)

    const result = await runWorkbookAction(model, 'inspect', { apply })

    expect(apply).not.toHaveBeenCalled()
    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'adapter_missing_capability',
          message: 'Adapter is missing verifyChecks for verifyCheck',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('does not let apply results drop planned checks', async () => {
    const model = defineModel({
      name: 'run-apply-proof-boundary-model',

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
      apply: () => ({
        status: 'applied',
        checks: [],
      }),
      verifyChecks: (checks) => checks,
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Sheet1!C2 did not verify check exists: Sheet1!C2 exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('fails when the generic check verifier leaves checks planned', async () => {
    const model = defineModel({
      name: 'run-unverified-check-model',

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
      verifyChecks: (checks) => checks,
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_not_verified',
          message: 'Sheet1!C2 did not verify check exists: Sheet1!C2 exists',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })

  it('passes non-readback and custom checks through the generic check verifier', async () => {
    const model = defineModel({
      name: 'run-check-proof-model',

      find(workbook) {
        const table = workbook.findTable({ name: 'Inputs', headers: ['Amount', 'Rate'] })
        return {
          table,
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [
          workbook.check.exists(refs.table),
          workbook.check.noFormulaErrors(refs.result),
          workbook.check.custom({
            kind: 'consumerInvariant',
            target: refs.result,
            refs: [refs.table],
            message: 'Consumer invariant holds',
          }),
        ]
      },

      actions: {
        inspect({ refs }) {
          void refs.table
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'passed' })),
    })

    expect(result).toMatchObject({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({ status: 'passed', kind: 'exists', message: 'Inputs exists' }),
        expect.objectContaining({ status: 'passed', kind: 'noFormulaErrors', message: 'Sheet1!C2 has no formula errors' }),
        expect.objectContaining({ status: 'passed', kind: 'consumerInvariant', message: 'Consumer invariant holds' }),
      ],
    })
  })

  it('returns failed when the generic check verifier marks a check failed', async () => {
    const model = defineModel({
      name: 'run-check-failure-model',

      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'C2' }),
        }
      },

      checks({ refs, workbook }) {
        return [
          workbook.check.custom({
            kind: 'consumerInvariant',
            target: refs.result,
            message: 'Consumer invariant holds',
          }),
        ]
      },

      actions: {
        inspect({ refs }) {
          void refs.result
        },
      },
    })

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map((checkResult) => ({ ...checkResult, status: 'failed' })),
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'check_failed',
          message: 'Sheet1!C2 failed check consumerInvariant: Consumer invariant holds',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'failed',
          kind: 'consumerInvariant',
          message: 'Consumer invariant holds',
        }),
      ],
    })
  })

  it('rejects inherited generic check verifier fields', async () => {
    const model = defineModel({
      name: 'run-inherited-check-proof-model',

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
      verifyChecks: (checks) => checks.map(inheritedVerifiedCheck),
    })

    expect(result).toMatchObject({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned an invalid check at index 0',
        },
      ],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists', message: 'Sheet1!C2 exists' })],
    })
  })
})
