// @vitest-environment jsdom
import { act, createElement, useEffect } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { WorkerHandle, WorkerRuntimeSelection } from '../runtime-session.js'
import { flushScheduledSelectionPersistence } from '../selection-persistence.js'
import { useWorkerWorkbookInteractionState } from '../use-worker-workbook-interaction-state.js'

function InteractionHarness(props: {
  documentId: string
  selection: WorkerRuntimeSelection
  selectedCell: CellSnapshot
  workerHandle: WorkerHandle
  invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  sendSelectionChanged: (selection: WorkerRuntimeSelection) => void
  capture: (value: ReturnType<typeof useWorkerWorkbookInteractionState>) => void
}) {
  const state = useWorkerWorkbookInteractionState({
    documentId: props.documentId,
    currentUserId: 'test-user',
    selection: props.selection,
    selectedCell: props.selectedCell,
    workerHandle: props.workerHandle,
    workerHandleRef: { current: props.workerHandle },
    writesAllowed: true,
    invokeMutation: props.invokeMutation,
    perfSession: {
      scope: 'test',
      markShellMounted() {},
      noteBootstrapResult() {},
      markFirstAuthoritativePatchVisible() {},
      markFirstReconcileStarted() {},
      markFirstReconcileSettled() {},
      markFirstSelectionVisible() {},
      markFirstLocalEditApplied() {},
      markFirstPasteApplied() {},
    },
    reportRuntimeError: vi.fn(),
    sendSelectionChanged: props.sendSelectionChanged,
  })

  useEffect(() => {
    props.capture(state)
  }, [props, state])

  return createElement('div')
}

function mountHarness(): {
  root: Root
  render: (props: Parameters<typeof InteractionHarness>[0]) => Promise<void>
} {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const root = createRoot(host)
  const render = async (props: Parameters<typeof InteractionHarness>[0]) => {
    await act(async () => {
      root.render(createElement(InteractionHarness, props))
    })
  }
  return { root, render }
}

describe('useWorkerWorkbookInteractionState hydrated edit base', () => {
  let originalRequestIdleCallback: Window['requestIdleCallback'] | undefined

  beforeEach(() => {
    originalRequestIdleCallback = window.requestIdleCallback
    Object.defineProperty(window, 'requestIdleCallback', {
      configurable: true,
      value: undefined,
      writable: true,
    })
  })

  afterEach(() => {
    Object.defineProperty(window, 'requestIdleCallback', {
      configurable: true,
      value: originalRequestIdleCallback,
      writable: true,
    })
    flushScheduledSelectionPersistence()
    document.body.innerHTML = ''
    vi.clearAllMocks()
  })

  it('adopts the first hydrated selected-cell value as the edit base after name-box navigation', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const selectedCell = stringCell('Sheet1', 'A1', 'one')
    const workerHandle = {
      viewportStore: createViewportStoreMapStub([selectedCell]),
    }
    const invokeMutation = vi.fn(async () => undefined)
    const sendSelectionChanged = vi.fn()
    const harness = mountHarness()
    let captured: ReturnType<typeof useWorkerWorkbookInteractionState> | null = null

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell,
      workerHandle,
      invokeMutation,
      sendSelectionChanged,
      capture: (value) => {
        captured = value
      },
    })
    if (!captured) {
      throw new Error('Expected interaction state capture')
    }

    await act(async () => {
      captured?.selectAddress('Sheet1', 'F6')
      captured?.beginEditing('', 'select-all', 'formula')
      captured?.handleEditorChange('7777777')
      workerHandle.viewportStore.setCellSnapshot(stringCell('Sheet1', 'F6', 'note-5-5'))
    })

    await harness.render({
      documentId: 'doc-1',
      selection: { sheetName: 'Sheet1', address: 'A1' },
      selectedCell: stringCell('Sheet1', 'A1', 'one'),
      workerHandle,
      invokeMutation,
      sendSelectionChanged,
      capture: (value) => {
        captured = value
      },
    })

    let commitResult: boolean | undefined
    await act(async () => {
      commitResult = captured?.commitEditor(undefined, '7777777')
      await Promise.resolve()
    })

    expect(commitResult).toBe(true)
    expect(captured?.editorConflictBanner).toBeNull()
    expect(invokeMutation).not.toHaveBeenCalled()

    await act(async () => {
      await captured?.flushPendingEditCommit()
    })

    expect(invokeMutation).toHaveBeenCalledWith('setCellValue', 'Sheet1', 'F6', 7777777)

    await act(async () => {
      harness.root.unmount()
    })
  })
})

function createViewportStoreMapStub(cells: readonly CellSnapshot[]) {
  const cellMap = new Map(cells.map((cell) => [`${cell.sheetName}!${cell.address}`, cell] as const))
  return {
    getCell(targetSheetName: string, targetAddress: string) {
      return cellMap.get(`${targetSheetName}!${targetAddress}`) ?? stringCell(targetSheetName, targetAddress, '')
    },
    hasCellSnapshot(targetSheetName: string, targetAddress: string) {
      return cellMap.has(`${targetSheetName}!${targetAddress}`)
    },
    setCellSnapshot: vi.fn((snapshot: CellSnapshot) => {
      cellMap.set(`${snapshot.sheetName}!${snapshot.address}`, snapshot)
    }),
  }
}

function stringCell(sheetName: string, address: string, value: string): CellSnapshot {
  return {
    sheetName,
    address,
    value: {
      tag: ValueTag.String,
      value,
    },
    input: value,
    flags: 0,
    version: 1,
  }
}
