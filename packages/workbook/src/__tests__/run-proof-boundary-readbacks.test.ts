import { describe, expect, it } from 'vitest'
import {
  defineModel,
  describeRunResult,
  findRange,
  runWorkbookAction,
  toWorkbookRefData,
  verifyWorkbookReadbacks,
  workbookActionCommandDigest,
  type WorkbookActionPlan,
  type WorkbookCheckResult,
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

function proofModel(): WorkbookModel<{ readonly result: ReturnType<typeof findRange> }> {
  return defineModel({
    name: 'run-proof-model',

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
    commandReceipts: plan.commands.map((command, commandIndex) => {
      const resolvedRefs = resolvedRefsForCommand(command)
      return {
        commandIndex,
        commandKind: command.kind,
        commandDigest: workbookActionCommandDigest(command),
        previewOps: plan.ops,
        appliedOps: plan.ops,
        ...(resolvedRefs !== undefined ? { resolvedRefs } : {}),
      }
    }),
  }
}

function resolvedRefsForCommand(command: WorkbookActionPlan['commands'][number]) {
  const refs: Record<string, unknown> = {}
  if (command.target !== undefined) {
    refs['target'] = toWorkbookRefData(command.target)
  }
  if (command.kind === 'writeFormula' && command.inputs.length > 0) {
    refs['inputs'] = command.inputs.map((input) => toWorkbookRefData(input))
  }
  return Object.keys(refs).length === 0 ? undefined : refs
}

function invalidProofCheck(checkResult: WorkbookCheckResult): WorkbookCheckResult {
  const verified = {
    ...checkResult,
    status: 'passed' as const,
  }
  Object.defineProperty(verified, 'proof', {
    enumerable: true,
    value: { when: new Date(0) },
  })
  return verified
}

function withUnsupportedField(checkResult: WorkbookCheckResult): WorkbookCheckResult {
  const verified = {
    ...checkResult,
    status: 'passed' as const,
    proof: { source: 'adapter' },
  }
  Object.defineProperty(verified, 'adapterObject', {
    enumerable: true,
    value: { leaked: true },
  })
  return verified
}

function sparseArray(length = 1): unknown[] {
  const value: unknown[] = []
  value.length = length
  return value
}

function accessorArray(get: () => unknown): unknown[] {
  const value = sparseArray()
  Object.defineProperty(value, '0', {
    enumerable: true,
    get,
  })
  return value
}

function arrayBackedVerifiedCheck(checkResult: WorkbookCheckResult): unknown[] {
  const verified: unknown[] = []
  Object.entries({
    ...checkResult,
    status: 'passed' as const,
  }).forEach(([key, value]) => {
    Object.defineProperty(verified, key, {
      enumerable: true,
      value,
    })
  })
  return verified
}

function arrayBackedReadback(fields: Record<string, unknown>): unknown[] {
  const readback: unknown[] = []
  Object.entries(fields).forEach(([key, value]) => {
    Object.defineProperty(readback, key, {
      enumerable: true,
      value,
    })
  })
  return readback
}

function arrayBackedCheck(checkResult: WorkbookCheckResult): unknown[] {
  const check: unknown[] = []
  Object.entries(checkResult).forEach(([key, value]) => {
    Object.defineProperty(check, key, {
      enumerable: true,
      value,
    })
  })
  return check
}

describe('@bilig/workbook run proof boundary readback checks', () => {
  it('allows generic check verifiers to add JSON-safe proof', async () => {
    const model = proofModel()
    const proof = {
      scannedCells: 24,
      errorCells: [],
      source: 'adapter',
    }

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) =>
        checks.map((checkResult) => ({
          ...checkResult,
          status: 'passed',
          proof,
        })),
    })

    expect(result).toEqual({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'exists',
          proof,
        }),
      ],
    })
    expect(describeRunResult(result)).toEqual({
      status: 'done',
      changed: [],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'exists',
          proof,
        }),
      ],
    })
  })

  it('rejects verifier proof that is not JSON-safe', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) => checks.map(invalidProofCheck),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned invalid proof at index 0: Action input at input.when must be a plain JSON object, not Date',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('rejects accessor-backed verifier proof without invoking getters', async () => {
    const model = proofModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: (checks) =>
        checks.map((checkResult) => {
          const verified = {
            ...checkResult,
            status: 'passed' as const,
          }
          Object.defineProperty(verified, 'proof', {
            enumerable: true,
            get() {
              getterInvoked = true
              throw new Error('getter must not run')
            },
          })
          return verified
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned invalid proof at index 0: Workbook check result proof must be a data property',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed verifier check arrays without invoking getters', async () => {
    const model = proofModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'inspect', {
      verifyChecks: () =>
        // @ts-expect-error js-caller boundary: exercising JS adapters that bypass the check array type
        accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned an invalid check at index 0',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects array-backed verifier checks as uninspectable runtime proof', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      // @ts-expect-error js-caller boundary: exercising JS adapters that bypass the check result type
      verifyChecks: (checks) => checks.map(arrayBackedVerifiedCheck),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'invalid_check_verification',
          message: 'Check verifier returned an invalid check at index 0',
        },
      ],
      changed: [],
      checks: [expect.objectContaining({ status: 'planned', kind: 'exists' })],
    })
  })

  it('strips unsupported verifier fields from check results', async () => {
    const model = proofModel()

    const result = await runWorkbookAction(model, 'inspect', {
      apply: applied,
      verifyChecks: (checks) => checks.map(withUnsupportedField),
    })

    expect(result.status).toBe('done')
    const check = first(result.checks)
    expect(check).toEqual(
      expect.objectContaining({
        status: 'passed',
        kind: 'exists',
        proof: { source: 'adapter' },
      }),
    )
    expect(Object.hasOwn(check, 'adapterObject')).toBe(false)
  })

  it('rejects readbacks that were not requested by checks', async () => {
    const model = valueModel()
    const unexpected = findRange({ sheetName: 'Sheet1', address: 'D2' })
    const read: Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read'] = (targets) => [
      {
        target: first(targets),
        value: 12,
      },
      {
        target: unexpected,
        value: 99,
      },
    ]

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read,
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_unexpected',
          message: 'Sheet1!D2 was returned by readback but was not requested',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects duplicate readbacks for the same target', async () => {
    const model = valueModel()
    const read: Required<WorkbookRunAdapter<{ output: ReturnType<typeof findRange> }>>['read'] = (targets) => {
      const target = first(targets)
      return [
        {
          target,
          value: 12,
        },
        {
          target,
          value: 12,
        },
      ]
    }

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read,
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_duplicate',
          message: 'Sheet1!B2 was returned by readback more than once',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'passed',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects accessor-backed readback arrays without invoking getters', async () => {
    const model = valueModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: () =>
        // @ts-expect-error js-caller boundary: exercising JS adapters that bypass the readback array type
        accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        }),
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0] must be a data property',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed readback values without invoking getters', async () => {
    const model = valueModel()
    let getterInvoked = false

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: (targets) => {
        const readback = {
          target: first(targets),
        }
        Object.defineProperty(readback, 'value', {
          enumerable: true,
          get() {
            getterInvoked = true
            throw new Error('getter must not run')
          },
        })
        return [readback]
      },
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0].value must be a data property',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects array-backed readback objects as uninspectable runtime proof', async () => {
    const model = valueModel()

    const result = await runWorkbookAction(model, 'write', {
      apply: applied,
      read: (targets) => [
        // @ts-expect-error js-caller boundary: exercising JS adapters that bypass the readback object type
        arrayBackedReadback({
          target: first(targets),
          value: 12,
        }),
      ],
    })

    expect(result).toEqual({
      status: 'failed',
      errors: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback at readbacks[0] is invalid',
        },
      ],
      apply: expect.objectContaining({ matched: true }),
      changed: [
        {
          kind: 'writeValue',
          target: expect.objectContaining({ label: 'Sheet1!B2' }),
          message: 'Write value to Sheet1!B2',
        },
      ],
      checks: [
        expect.objectContaining({
          status: 'planned',
          kind: 'valueEquals',
          message: 'Sheet1!B2 equals 12',
        }),
      ],
    })
  })

  it('rejects accessor-backed public readback arrays without invoking getters', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'B2' })
    let getterInvoked = false

    const verification = verifyWorkbookReadbacks(
      [
        {
          status: 'planned',
          kind: 'valueEquals',
          target,
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      // @ts-expect-error exercising JS callers that bypass the readback array type
      accessorArray(() => {
        getterInvoked = true
        throw new Error('getter must not run')
      }),
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [
        {
          status: 'planned',
          kind: 'valueEquals',
          target,
          message: 'Sheet1!B2 equals 12',
          expectation: {
            kind: 'valueEquals',
            value: 12,
          },
        },
      ],
      issues: [
        {
          code: 'readback_invalid',
          message: 'Workbook readback proof at readbacks[0] must be a data property',
        },
      ],
    })
    expect(Object.isFrozen(verification)).toBe(true)
    expect(Object.isFrozen(verification.checks)).toBe(true)
    expect(Object.isFrozen(verification.checks[0])).toBe(true)
    expect(Object.isFrozen(verification.issues)).toBe(true)
    expect(Object.isFrozen(verification.issues[0])).toBe(true)
    expect(getterInvoked).toBe(false)
  })

  it('rejects array-backed public check objects as uninspectable readback proof', () => {
    const target = findRange({ sheetName: 'Sheet1', address: 'B2' })
    const check: WorkbookCheckResult = {
      status: 'planned',
      kind: 'valueEquals',
      target,
      message: 'Sheet1!B2 equals 12',
      expectation: {
        kind: 'valueEquals',
        value: 12,
      },
    }

    const verification = verifyWorkbookReadbacks(
      // @ts-expect-error exercising JS callers that bypass the check object type
      [arrayBackedCheck(check)],
      [{ target, value: 12 }],
    )

    expect(verification).toEqual({
      status: 'failed',
      checks: [],
      issues: [
        {
          code: 'readback_invalid',
          message: 'Workbook check at checks[0] is invalid',
        },
      ],
    })
  })
})
