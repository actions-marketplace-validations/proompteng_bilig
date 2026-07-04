import { describe, expect, it } from 'vitest'
import { defineModel, describeModel, describePlanResult, planWorkbookAction } from '../index.js'

function accessorArray<T>(getter: () => T): readonly T[] {
  const values: T[] = []
  Object.defineProperty(values, '0', {
    configurable: true,
    enumerable: true,
    get: getter,
  })
  return values
}
function arrayBackedRecord(fields: Record<string, unknown>): unknown[] {
  const value: unknown[] = []
  for (const [key, entry] of Object.entries(fields)) {
    Object.defineProperty(value, key, {
      enumerable: true,
      value: entry,
    })
  }
  return value
}

function customPrototypeRecord(fields: Record<string, unknown>): Record<string, unknown> {
  const value: Record<string, unknown> = {}
  Object.setPrototypeOf(value, { inherited: true })
  for (const [key, entry] of Object.entries(fields)) {
    Object.defineProperty(value, key, {
      enumerable: true,
      value: entry,
    })
  }
  return value
}

describe('@bilig/workbook model api plan result boundaries', () => {
  it('describes models as JSON-safe manifests without running model code', () => {
    const model = defineModel({
      name: 'described-model-manifest',
      find() {
        throw new Error('find should not run during model description')
      },
      checks() {
        throw new Error('checks should not run during model description')
      },
      actions: {
        calculate() {
          throw new Error('action should not run during model description')
        },
        reset() {
          throw new Error('action should not run during model description')
        },
      },
    })

    const description = describeModel(model)

    expect(description).toEqual({
      name: 'described-model-manifest',
      actions: ['calculate', 'reset'],
      actionDetails: [{ name: 'calculate' }, { name: 'reset' }],
      hasChecks: true,
    })
    expect(JSON.parse(JSON.stringify(description))).toEqual(description)
    expect(Object.isFrozen(description)).toBe(true)
    expect(Object.isFrozen(description.actions)).toBe(true)
    expect(Object.isFrozen(description.actionDetails)).toBe(true)
    expect(Object.isFrozen(description.actionDetails[0])).toBe(true)
  })

  it('returns structured planning failures instead of forcing agents to catch exceptions', () => {
    const model = defineModel({
      name: 'failing-model',
      find(workbook) {
        return {
          result: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: {
        calculate() {
          throw new Error('formula target was not resolved')
        },
      },
    })

    const missing = planWorkbookAction(model, 'missing')
    const failed = planWorkbookAction(model, 'calculate')

    expect(missing).toEqual({
      status: 'failed',
      modelName: 'failing-model',
      actionName: 'missing',
      checks: [],
      errors: [
        {
          code: 'action_not_found',
          message: 'Workbook model failing-model does not define action missing',
        },
      ],
    })
    expect(Object.isFrozen(missing)).toBe(true)
    expect(missing.status).toBe('failed')
    if (missing.status !== 'failed') {
      throw new Error('expected missing action failure')
    }
    expect(Object.isFrozen(missing.checks)).toBe(true)
    expect(Object.isFrozen(missing.errors)).toBe(true)
    expect(Object.isFrozen(missing.errors[0])).toBe(true)

    expect(failed).toEqual({
      status: 'failed',
      modelName: 'failing-model',
      actionName: 'calculate',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'formula target was not resolved',
        },
      ],
    })
    expect(Object.isFrozen(failed)).toBe(true)
    expect(failed.status).toBe('failed')
    if (failed.status !== 'failed') {
      throw new Error('expected action failure')
    }
    expect(Object.isFrozen(failed.checks)).toBe(true)
    expect(Object.isFrozen(failed.errors)).toBe(true)
    expect(Object.isFrozen(failed.errors[0])).toBe(true)
  })

  it('rejects non-data action names without coercing user objects', () => {
    const model = defineModel({
      name: 'action-name-boundary-model',
      find(workbook) {
        return {
          result: workbook.findName('result'),
        }
      },
      actions: {
        calculate({ refs, workbook }) {
          workbook.writeValue(refs.result, 1)
        },
      },
    })

    let coerced = false
    const actionName = {
      toString() {
        coerced = true
        throw new Error('action key coercion should not run')
      },
      valueOf() {
        coerced = true
        throw new Error('action key coercion should not run')
      },
    }

    const result = Reflect.apply(planWorkbookAction, undefined, [model, actionName])

    expect(coerced).toBe(false)
    expect(result).toEqual({
      status: 'failed',
      modelName: 'unknown-model',
      actionName: '<invalid-action-name>',
      checks: [],
      errors: [
        {
          code: 'invalid_action_name',
          message: 'Workbook action name must be a string',
          path: 'actionName',
          issueCode: 'invalid_action_name',
        },
      ],
    })
    expect(Object.isFrozen(result)).toBe(true)
    expect(result.status).toBe('failed')
    if (result.status !== 'failed') {
      throw new Error('expected invalid action-name failure')
    }
    expect(Object.isFrozen(result.errors)).toBe(true)
    expect(Object.isFrozen(result.errors[0])).toBe(true)
  })

  it('rejects malformed action helper calls before returning a plan', () => {
    let opGetterInvoked = false
    let stylePrototypeGetterInvoked = false
    const opWithGetter: Record<string, unknown> = {
      kind: 'setCellValue',
      sheetName: 'Sheet1',
      address: 'A1',
      value: 1,
    }
    Object.defineProperty(opWithGetter, 'extra', {
      enumerable: true,
      get() {
        opGetterInvoked = true
        throw new Error('op getter must not run')
      },
    })
    const styleWithPrototype = Object.create({
      get inherited() {
        stylePrototypeGetterInvoked = true
        throw new Error('style prototype getter must not run')
      },
    })
    Object.defineProperty(styleWithPrototype, 'font', {
      enumerable: true,
      value: { bold: true },
    })

    const model = defineModel({
      name: 'helper-validation-model',
      find(workbook) {
        return {
          target: workbook.findRange({ sheetName: 'Sheet1', address: 'A1' }),
        }
      },
      actions: {
        badTarget({ workbook }) {
          Reflect.apply(workbook.writeValue, undefined, [{ kind: 'range', id: 'missing-range', label: 'missing range' }, 1])
        },
        badValue({ refs, workbook }) {
          Reflect.apply(workbook.writeValue, undefined, [refs.target, Number.NaN])
        },
        badFormat({ refs, workbook }) {
          Reflect.apply(workbook.format, undefined, [refs.target, { numberFormat: 12 }])
        },
        emptyFormatOptions({ refs, workbook }) {
          workbook.format(refs.target, {})
        },
        emptyStyle({ refs, workbook }) {
          workbook.format(refs.target, { style: { font: {} } })
        },
        formatOptionsPrototype({ refs, workbook }) {
          Reflect.apply(workbook.format, undefined, [
            refs.target,
            customPrototypeRecord({
              numberFormat: '0.00',
            }),
          ])
        },
        badOp({ workbook }) {
          Reflect.apply(workbook.addOp, undefined, [opWithGetter])
        },
        addOpOptionsPrototype({ refs, workbook }) {
          Reflect.apply(workbook.addOp, undefined, [
            {
              kind: 'setCellValue',
              sheetName: 'Sheet1',
              address: 'A1',
              value: 1,
            },
            customPrototypeRecord({
              target: refs.target,
            }),
          ])
        },
        stylePrototype({ refs, workbook }) {
          workbook.format(refs.target, { style: styleWithPrototype })
        },
      },
    })

    expect(planWorkbookAction(model, 'badTarget')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'badTarget',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action writeValue target must be a workbook ref',
        },
      ],
    })
    expect(planWorkbookAction(model, 'badValue')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'badValue',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action writeValue value must be a finite JSON literal',
        },
      ],
    })
    expect(planWorkbookAction(model, 'badFormat')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'badFormat',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action format numberFormat must be a string, null, or undefined',
        },
      ],
    })
    expect(planWorkbookAction(model, 'emptyFormatOptions')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'emptyFormatOptions',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action format options must include style or numberFormat',
        },
      ],
    })
    expect(planWorkbookAction(model, 'emptyStyle')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'emptyStyle',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action format style must request at least one style field',
        },
      ],
    })
    expect(planWorkbookAction(model, 'formatOptionsPrototype')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'formatOptionsPrototype',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action format options must be an object',
        },
      ],
    })
    expect(planWorkbookAction(model, 'badOp')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'badOp',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action op.extra must be a data property',
        },
      ],
    })
    expect(planWorkbookAction(model, 'addOpOptionsPrototype')).toEqual({
      status: 'failed',
      modelName: 'helper-validation-model',
      actionName: 'addOpOptionsPrototype',
      checks: [],
      errors: [
        {
          code: 'action_failed',
          message: 'Workbook action addOp options must be an object',
        },
      ],
    })
    expect(opGetterInvoked).toBe(false)
    const stylePlan = planWorkbookAction(model, 'stylePrototype')
    expect(stylePlan.status).toBe('planned')
    if (stylePlan.status === 'planned') {
      const command = stylePlan.plan.commands[0]
      if (command?.kind !== 'format') {
        throw new Error('expected format command')
      }
      expect(command.style).toEqual({ font: { bold: true } })
      expect(Object.getPrototypeOf(command.style)).toBe(Object.prototype)
    }
    expect(stylePrototypeGetterInvoked).toBe(false)
  })

  it('keeps planned checks when action planning fails', () => {
    const model = defineModel({
      name: 'checkable-failure-model',
      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs' }),
        }
      },
      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },
      actions: {
        calculate() {
          throw new Error('cannot write without a result target')
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    expect(result.status).toBe('failed')
    if (result.status === 'failed') {
      expect(result.modelName).toBe('checkable-failure-model')
      expect(result.actionName).toBe('calculate')
      expect(result.checks).toEqual([
        {
          status: 'planned',
          kind: 'exists',
          target: expect.objectContaining({
            kind: 'table',
            name: 'Inputs',
          }),
          message: 'Inputs exists',
        },
      ])
      expect(result.errors).toEqual([
        {
          code: 'action_failed',
          message: 'cannot write without a result target',
        },
      ])
    }
  })

  it('rejects accessor-backed returned check arrays without invoking getters', () => {
    let getterInvoked = false
    const model = defineModel({
      name: 'returned-check-array-proof-model',
      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs' }),
        }
      },
      checks() {
        // @ts-expect-error exercising JS callers that bypass the returned check array type
        return accessorArray(() => {
          getterInvoked = true
          throw new Error('getter must not run')
        })
      },
      actions: {
        inspect() {},
      },
    })

    expect(planWorkbookAction(model, 'inspect')).toEqual({
      status: 'failed',
      modelName: 'returned-check-array-proof-model',
      actionName: 'inspect',
      checks: [],
      errors: [
        {
          code: 'checks_failed',
          message: 'Workbook check at checks[0] must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects accessor-backed returned check fields without invoking getters', () => {
    let getterInvoked = false
    const model = defineModel({
      name: 'returned-check-field-proof-model',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      checks({ refs }) {
        const checkResult = {
          status: 'planned',
          target: refs.output,
          message: 'Output exists',
        }
        Object.defineProperty(checkResult, 'kind', {
          enumerable: true,
          get() {
            getterInvoked = true
            throw new Error('getter must not run')
          },
        })
        // @ts-expect-error exercising JS callers that bypass the returned check type
        return [checkResult]
      },
      actions: {
        inspect() {},
      },
    })

    expect(planWorkbookAction(model, 'inspect')).toEqual({
      status: 'failed',
      modelName: 'returned-check-field-proof-model',
      actionName: 'inspect',
      checks: [],
      errors: [
        {
          code: 'checks_failed',
          message: 'Workbook check at checks[0].kind must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })

  it('rejects array-backed returned checks as uninspectable data', () => {
    const model = defineModel({
      name: 'returned-check-array-backed-proof-model',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      checks({ refs }) {
        // @ts-expect-error exercising JS callers that bypass the returned check type
        return [
          arrayBackedRecord({
            status: 'planned',
            kind: 'exists',
            target: refs.output,
            message: 'Output exists',
          }),
        ]
      },
      actions: {
        inspect() {},
      },
    })

    expect(planWorkbookAction(model, 'inspect')).toEqual({
      status: 'failed',
      modelName: 'returned-check-array-backed-proof-model',
      actionName: 'inspect',
      checks: [],
      errors: [
        {
          code: 'checks_failed',
          message: 'Workbook check at checks[0] must be an object',
        },
      ],
    })
  })

  it('rejects custom-prototype returned checks as uninspectable data', () => {
    const model = defineModel({
      name: 'returned-check-custom-prototype-proof-model',
      find(workbook) {
        return {
          output: workbook.findRange({ sheetName: 'Sheet1', address: 'B2' }),
        }
      },
      checks({ refs }) {
        // @ts-expect-error exercising JS callers that bypass the returned check type
        return [
          customPrototypeRecord({
            status: 'planned',
            kind: 'exists',
            target: refs.output,
            message: 'Output exists',
          }),
        ]
      },
      actions: {
        inspect() {},
      },
    })

    expect(planWorkbookAction(model, 'inspect')).toEqual({
      status: 'failed',
      modelName: 'returned-check-custom-prototype-proof-model',
      actionName: 'inspect',
      checks: [],
      errors: [
        {
          code: 'checks_failed',
          message: 'Workbook check at checks[0] must be an object',
        },
      ],
    })
  })

  it('describes failed plan results without raw workbook refs', () => {
    const model = defineModel({
      name: 'described-failure-model',

      find(workbook) {
        return {
          table: workbook.findTable({ name: 'Inputs' }),
        }
      },

      checks({ refs, workbook }) {
        return [workbook.check.exists(refs.table)]
      },

      actions: {
        calculate() {
          throw new Error('missing output target')
        },
      },
    })

    const result = planWorkbookAction(model, 'calculate')
    const described = describePlanResult(result)

    expect(described).toEqual({
      status: 'failed',
      modelName: 'described-failure-model',
      actionName: 'calculate',
      errors: [
        {
          code: 'action_failed',
          message: 'missing output target',
        },
      ],
      checks: [
        {
          status: 'planned',
          kind: 'exists',
          target: {
            kind: 'table',
            id: 'table_Inputs',
            label: 'Inputs',
            name: 'Inputs',
          },
          message: 'Inputs exists',
        },
      ],
    })
    expect(JSON.parse(JSON.stringify(described))).toEqual(described)
    expect(Object.isFrozen(described)).toBe(true)
    expect(Object.isFrozen(described.errors)).toBe(true)
    expect(Object.isFrozen(described.errors[0])).toBe(true)
    expect(Object.isFrozen(described.checks)).toBe(true)
    expect(Object.isFrozen(described.checks[0])).toBe(true)
    expect(Object.isFrozen(described.checks[0]?.target)).toBe(true)
  })
})
