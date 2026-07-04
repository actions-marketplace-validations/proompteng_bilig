import { describe, expect, it } from 'vitest'
import {
  checkWorkbookCommandResult,
  checkWorkbookCommandResultForBundle,
  normalizeWorkbookCommandBundle,
  workbookCommandResultFor,
  workbookCommandResultForReceipts,
  workbookOpCommandReceipt,
} from '../index.js'

const mutationRequest = {
  featureId: 'cells',
  commandId: 'cells.setValue',
  category: 'mutation',
  mode: 'applyAndVerify',
} as const

const setCellValueOp = {
  kind: 'setCellValue',
  sheetName: 'Sheet1',
  address: 'A1',
  value: 1,
} as const

describe('@bilig/workbook command bundle api receipt validation', () => {
  it('rejects low-level op receipts that do not prove the planned op', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-op',
      targetRevision: 7,
      idempotencyKey: 'bundle-op',
      commands: [
        {
          id: 'op-write-a1',
          kind: 'op',
          destructive: true,
          op: setCellValueOp,
          touchedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ],
    })
    const [command] = bundle.commands
    if (command === undefined) {
      throw new Error('expected command')
    }
    const wrongOp = {
      ...setCellValueOp,
      value: 2,
    } as const

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        workbookOpCommandReceipt(command, 0, {
          status: 'applied',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        }),
      ]),
    ).toThrow('Workbook command result is invalid: Workbook command result receipt 0 appliedOps must equal commands[0].op')

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        workbookOpCommandReceipt(command, 0, {
          status: 'applied',
          previewOps: [wrongOp],
          appliedOps: [wrongOp],
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        }),
      ]),
    ).toThrow('Workbook command result is invalid: Workbook command result receipt 0 previewOps must equal commands[0].op')

    expect(
      checkWorkbookCommandResultForBundle(bundle, {
        ...workbookCommandResultFor(bundle),
        status: 'applied',
        revision: 8,
        receipts: [
          {
            status: 'applied',
            featureId: 'workbook-op',
            commandId: 'op-write-a1',
            category: 'operation',
            appliedOps: [wrongOp],
            changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
          },
        ],
        matched: null,
        changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'receipt_command_mismatch',
          path: 'receipts[0].appliedOps',
          message: 'Workbook command result receipt 0 appliedOps must equal commands[0].op',
        },
      ],
    })
  })

  it('rejects command result receipts that do not match request commands', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-3',
      targetRevision: 7,
      idempotencyKey: 'bundle-3',
      commands: [
        {
          kind: 'request',
          destructive: true,
          request: mutationRequest,
        },
      ],
    })

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        {
          status: 'applied',
          featureId: 'other',
          commandId: 'cells.setValue',
          category: 'mutation',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ]),
    ).toThrow('Workbook command result is invalid: receipts[0] does not match commands[0].request')
  })

  it('rejects command result receipts that do not match op commands', () => {
    const bundle = normalizeWorkbookCommandBundle({
      id: 'bundle-op',
      targetRevision: 7,
      idempotencyKey: 'bundle-op',
      commands: [
        {
          id: 'op-write-a1',
          kind: 'op',
          destructive: true,
          op: setCellValueOp,
        },
      ],
    })

    expect(() =>
      workbookCommandResultForReceipts(bundle, [
        {
          status: 'applied',
          featureId: 'other',
          commandId: 'op-write-a1',
          category: 'operation',
          changedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }],
        },
      ]),
    ).toThrow('Workbook command result is invalid: receipts[0] does not match commands[0].op')
  })

  it('rejects uninspectable command result data without invoking getters', () => {
    let getterInvoked = false
    const result = {
      status: 'applied',
      targetRevision: 7,
      idempotencyKey: 'bundle-4',
      commandCount: 1,
      touchedRanges: [],
      touchedCellCount: 0,
      receipts: [
        {
          status: 'applied',
          featureId: 'cells',
          commandId: 'cells.setValue',
          category: 'mutation',
        },
      ],
      matched: null,
      changedRanges: [],
    }
    Object.defineProperty(result.receipts[0], 'proof', {
      enumerable: true,
      get() {
        getterInvoked = true
        throw new Error('getter must not run')
      },
    })

    expect(checkWorkbookCommandResult(result)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_receipt',
          path: 'receipts[0].proof',
          message: 'Workbook command receipt proof must be a data property',
        },
      ],
    })
    expect(getterInvoked).toBe(false)
  })
})
