// @vitest-environment jsdom
import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, test, vi } from 'vitest'
import type { captureGridClipboardSelection } from '../gridClipboardKeyboardController.js'
import {
  handleGridKey,
  isGridKeyboardEditableTarget,
  shouldHandleGridSurfaceKey,
  shouldHandleGridWindowKey,
  shouldSuppressWorkbookChromeClearKey,
  shouldSuppressWorkbookChromeSelectionKeyUp,
} from '../gridClipboardKeyboardController.js'
import { createGridSelection, createRangeSelection } from '../gridSelection.js'
import { useWorkbookGridKeyboardHandler } from '../useWorkbookGridKeyboardHandler.js'

function createCellSnapshot(address: string, input: string): CellSnapshot {
  return {
    sheetName: 'Sheet1',
    address,
    input,
    value: { tag: ValueTag.String, value: input, stringId: 0 },
    flags: 0,
    version: 0,
  }
}

function KeyboardHandlerHarness(props: {
  applyClipboardValues?: ReturnType<typeof vi.fn>
  beginSelectedEdit: ReturnType<typeof vi.fn>
  captureInternalClipboardSelection?: ReturnType<typeof vi.fn>
  getGridSelection?: () => ReturnType<typeof createGridSelection>
  internalClipboardRef?: { current: ReturnType<typeof captureGridClipboardSelection> }
  lastKeyboardClipboardRef?: { current: ReturnType<typeof captureGridClipboardSelection> }
  onCancelEdit?: ReturnType<typeof vi.fn>
  onClearCell?: ReturnType<typeof vi.fn>
  onCommitEdit?: ReturnType<typeof vi.fn>
  onFillRange?: ReturnType<typeof vi.fn>
  scrollActiveCellIntoView?: ReturnType<typeof vi.fn>
  setGridSelection: ReturnType<typeof vi.fn>
  onSelectionChange: ReturnType<typeof vi.fn>
}) {
  const hostRef = { current: null as HTMLDivElement | null }
  const internalClipboardRef = props.internalClipboardRef ?? { current: null }

  useWorkbookGridKeyboardHandler({
    applyClipboardValues: props.applyClipboardValues ?? vi.fn(),
    beginSelectedEdit: props.beginSelectedEdit,
    captureInternalClipboardSelection: props.captureInternalClipboardSelection ?? vi.fn(),
    editorValue: '',
    engine: {
      getCell: (_sheetName, address) => createCellSnapshot(address, 'value'),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    },
    gridSelection: createGridSelection(1, 1),
    getGridSelection: props.getGridSelection,
    hostRef,
    internalClipboardRef,
    isEditingCell: false,
    onCancelEdit: props.onCancelEdit ?? vi.fn(),
    onClearCell: props.onClearCell ?? vi.fn(),
    onCommitEdit: props.onCommitEdit ?? vi.fn(),
    onEditorChange: vi.fn(),
    onFillRange: props.onFillRange ?? vi.fn(),
    onSelectionChange: props.onSelectionChange,
    scrollActiveCellIntoView: props.scrollActiveCellIntoView ?? vi.fn(),
    lastKeyboardClipboardRef: props.lastKeyboardClipboardRef ?? { current: null },
    pendingClipboardCopySequenceRef: { current: 0 },
    pendingKeyboardPasteIntentRef: { current: null },
    pendingKeyboardPasteSequenceRef: { current: 0 },
    pendingTypeSeedRef: { current: null },
    selectedCell: { col: 1, row: 1 },
    setGridSelection: props.setGridSelection,
    sheetName: 'Sheet1',
    suppressNextNativePasteRef: { current: false },
    toggleSelectedBooleanCell: vi.fn(),
  })

  return createElement('div', {
    ref: (node: HTMLDivElement | null) => {
      hostRef.current = node
    },
  })
}

describe('gridClipboardKeyboardController clear delete and history shortcuts', () => {
  test('routes primary-modified Backspace to active-cell scrolling without clearing content', () => {
    const onClearCell = vi.fn()
    const scrollActiveCellIntoView = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'Backspace',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createGridSelection(3, 7),
      internalClipboardRef: { current: null },
      lastKeyboardClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell,
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange: vi.fn(),
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView,
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteIntentRef: { current: null },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 3, row: 7 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(scrollActiveCellIntoView).toHaveBeenCalledTimes(1)
    expect(onClearCell).not.toHaveBeenCalled()
  })

  test('suppresses no-op fill shortcuts so the browser does not steal grid focus', () => {
    const onFillRange = vi.fn()
    const preventDefault = vi.fn()

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: '',
      event: {
        key: 'd',
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createRangeSelection(createGridSelection(1, 1), [1, 1], [3, 1]),
      internalClipboardRef: { current: null },
      lastKeyboardClipboardRef: { current: null },
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onFillRange,
      onSelectionChange: vi.fn(),
      scrollActiveCellIntoView: vi.fn(),
      pendingClipboardCopySequenceRef: { current: 0 },
      pendingKeyboardPasteIntentRef: { current: null },
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      sheetName: 'Sheet1',
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    })

    expect(preventDefault).toHaveBeenCalled()
    expect(onFillRange).not.toHaveBeenCalled()
  })

  test('only claims global grid shortcuts when focus is on the document body', () => {
    const host = document.createElement('div')
    document.body.append(host)

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(true)

    const input = document.createElement('input')
    document.body.append(input)
    input.focus()
    expect(shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, input, host)).toBe(
      false,
    )

    const textarea = document.createElement('textarea')
    textarea.dataset['testid'] = 'cell-editor-input'
    document.body.append(textarea)
    textarea.focus()
    expect(isGridKeyboardEditableTarget(textarea)).toBe(true)
    expect(shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'r', metaKey: false, shiftKey: false }, textarea, host)).toBe(
      false,
    )
  })

  test('does not claim global grid shortcuts while a modal dialog is open', () => {
    const host = document.createElement('div')
    document.body.append(host)

    const modal = document.createElement('div')
    modal.setAttribute('aria-modal', 'true')
    modal.setAttribute('role', 'dialog')
    document.body.append(modal)

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'r', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Tab', metaKey: false, shiftKey: false }, document.body, host),
    ).toBe(false)
  })

  test('ignores globally prevented keydown events before routing them into the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const beginSelectedEdit = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: '?',
      shiftKey: true,
    })
    event.preventDefault()

    await act(async () => {
      window.dispatchEvent(event)
    })

    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes macOS clipboard paste-values-only keyup fallback when the browser suppresses V keydown', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const originalClipboard = navigator.clipboard
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: vi.fn().mockResolvedValue('3\t=B2*2\n4\t=B3*2'),
      },
    })

    const applyClipboardValues = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(
          createElement(KeyboardHandlerHarness, {
            applyClipboardValues,
            beginSelectedEdit: vi.fn(),
            getGridSelection: () => createGridSelection(3, 1),
            internalClipboardRef: {
              current: {
                operation: 'copy',
                sourceStartAddress: 'B2',
                sourceEndAddress: 'C3',
                signature: '3\u001f=B2*2\u001e4\u001f=B3*2',
                plainText: '3\t=B2*2\n4\t=B3*2',
                valuesOnlyPlainText: '3\t6\n4\t8',
                rowCount: 2,
                colCount: 2,
              },
            },
            onSelectionChange: vi.fn(),
            setGridSelection: vi.fn(),
          }),
        )
      })

      const event = new KeyboardEvent('keyup', {
        bubbles: true,
        cancelable: true,
        code: 'KeyV',
        key: 'V',
        metaKey: true,
        shiftKey: true,
      })

      await act(async () => {
        window.dispatchEvent(event)
        await Promise.resolve()
      })

      expect(event.defaultPrevented).toBe(true)
      expect(applyClipboardValues).toHaveBeenCalledWith(
        [3, 1],
        [
          ['3', '6'],
          ['4', '8'],
        ],
        { pasteValuesOnly: true },
      )
    } finally {
      await act(async () => {
        root.unmount()
      })
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: originalClipboard,
      })
    }
  })

  test('does not route Delete from editable event targets into the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const formulaInput = document.createElement('input')
    document.body.append(formulaInput)
    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Delete',
    })

    await act(async () => {
      formulaInput.dispatchEvent(event)
    })

    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes high-confidence grid shortcuts from workbook chrome without stealing button activation keys or delete ownership', () => {
    document.body.innerHTML = ''
    const scope = document.createElement('section')
    scope.dataset['workbookKeyboardScope'] = 'true'
    const gridHost = document.createElement('div')
    const toolbarButton = document.createElement('button')
    scope.append(gridHost, toolbarButton)
    document.body.append(scope)
    toolbarButton.focus()

    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldSuppressWorkbookChromeClearKey(
        { altKey: false, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldSuppressWorkbookChromeClearKey(
        { altKey: false, ctrlKey: false, key: 'Backspace', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: 'c', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: true, ctrlKey: true, key: '-', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: true }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: true, key: ' ', metaKey: false, shiftKey: true }, toolbarButton, gridHost),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey(
        { altKey: false, ctrlKey: false, key: 'ArrowDown', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldHandleGridWindowKey(
        { altKey: true, ctrlKey: false, key: 'ArrowLeft', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: true, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldSuppressWorkbookChromeClearKey(
        { altKey: true, ctrlKey: false, key: 'Delete', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'Enter', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldHandleGridWindowKey({ altKey: false, ctrlKey: false, key: 'x', metaKey: false, shiftKey: false }, toolbarButton, gridHost),
    ).toBe(false)
    expect(
      shouldSuppressWorkbookChromeSelectionKeyUp(
        { altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: true },
        toolbarButton,
        gridHost,
      ),
    ).toBe(true)
    expect(
      shouldSuppressWorkbookChromeSelectionKeyUp(
        { altKey: false, ctrlKey: false, key: ' ', metaKey: false, shiftKey: false },
        toolbarButton,
        gridHost,
      ),
    ).toBe(false)
  })

  test('suppresses browser clear-key defaults from workbook chrome without clearing the grid', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const scope = document.createElement('section')
    scope.dataset['workbookKeyboardScope'] = 'true'
    const toolbarButton = document.createElement('button')
    const host = document.createElement('div')
    scope.append(toolbarButton, host)
    document.body.append(scope)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          getGridSelection: () => createGridSelection(4, 8),
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    toolbarButton.focus()
    const deleteEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Delete' })
    const backspaceEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Backspace' })
    await act(async () => {
      toolbarButton.dispatchEvent(deleteEvent)
      toolbarButton.dispatchEvent(backspaceEvent)
    })

    expect(deleteEvent.defaultPrevented).toBe(true)
    expect(backspaceEvent.defaultPrevented).toBe(true)
    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes row, column, and full-sheet selection shortcuts from workbook chrome', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const scope = document.createElement('section')
    scope.dataset['workbookKeyboardScope'] = 'true'
    const toolbarButton = document.createElement('button')
    const host = document.createElement('div')
    scope.append(toolbarButton, host)
    document.body.append(scope)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          getGridSelection: () => createGridSelection(3, 7),
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    toolbarButton.focus()
    const rowEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: ' ', shiftKey: true })
    const columnEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: true, key: ' ' })
    const sheetEvent = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ctrlKey: true, key: ' ', shiftKey: true })
    await act(async () => {
      toolbarButton.dispatchEvent(rowEvent)
      toolbarButton.dispatchEvent(columnEvent)
      toolbarButton.dispatchEvent(sheetEvent)
    })

    expect(rowEvent.defaultPrevented).toBe(true)
    expect(columnEvent.defaultPrevented).toBe(true)
    expect(sheetEvent.defaultPrevented).toBe(true)
    expect(setGridSelection).toHaveBeenCalledTimes(3)
    expect(setGridSelection.mock.calls[0]?.[0]?.rows.first()).toBe(7)
    expect(setGridSelection.mock.calls[0]?.[0]?.columns.first()).toBeUndefined()
    expect(setGridSelection.mock.calls[1]?.[0]?.columns.first()).toBe(3)
    expect(setGridSelection.mock.calls[1]?.[0]?.rows.first()).toBeUndefined()
    expect(setGridSelection.mock.calls[2]?.[0]?.current?.cell).toEqual([0, 0])
    expect(setGridSelection.mock.calls[2]?.[0]?.rows.first()).toBe(0)
    expect(setGridSelection.mock.calls[2]?.[0]?.columns.first()).toBe(0)
    expect(onSelectionChange).toHaveBeenCalledTimes(3)
    expect(beginSelectedEdit).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('routes only primary Backspace among modified delete keys', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const scrollActiveCellIntoView = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          onClearCell,
          onSelectionChange,
          scrollActiveCellIntoView,
          setGridSelection,
        }),
      )
    })

    const events = [
      { key: 'Delete', ctrlKey: true },
      { key: 'Delete', metaKey: true },
      { key: 'Delete', altKey: true },
      { key: 'Delete', shiftKey: true },
      { key: 'Backspace', altKey: true },
      { key: 'Backspace', shiftKey: true },
    ].map(
      (eventInit) =>
        new KeyboardEvent('keydown', {
          bubbles: true,
          cancelable: true,
          ...eventInit,
        }),
    )

    await act(async () => {
      for (const event of events) {
        window.dispatchEvent(event)
      }
    })

    for (const event of events) {
      expect(event.defaultPrevented).toBe(false)
    }

    expect(onClearCell).not.toHaveBeenCalled()
    expect(beginSelectedEdit).not.toHaveBeenCalled()
    expect(setGridSelection).not.toHaveBeenCalled()
    expect(onSelectionChange).not.toHaveBeenCalled()
    expect(scrollActiveCellIntoView).not.toHaveBeenCalled()

    const primaryBackspace = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      ctrlKey: true,
      key: 'Backspace',
    })
    await act(async () => {
      window.dispatchEvent(primaryBackspace)
    })
    expect(primaryBackspace.defaultPrevented).toBe(true)
    expect(scrollActiveCellIntoView).toHaveBeenCalledTimes(1)
    expect(onClearCell).not.toHaveBeenCalled()

    await act(async () => {
      root.unmount()
    })
  })

  test('reads the latest runtime selection when routing global clear keys', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    document.body.innerHTML = ''
    const beginSelectedEdit = vi.fn()
    const onClearCell = vi.fn()
    const setGridSelection = vi.fn()
    const onSelectionChange = vi.fn()
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        createElement(KeyboardHandlerHarness, {
          beginSelectedEdit,
          getGridSelection: () => createGridSelection(2, 3),
          onClearCell,
          onSelectionChange,
          setGridSelection,
        }),
      )
    })

    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Delete',
    })

    await act(async () => {
      window.dispatchEvent(event)
    })

    expect(onClearCell).toHaveBeenCalledWith({
      address: 'C4',
      kind: 'cell',
      range: { startAddress: 'C4', endAddress: 'C4' },
      sheetName: 'Sheet1',
    })

    await act(async () => {
      root.unmount()
    })
  })

  test('filters grid-surface key handling to grid-relevant keys', () => {
    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: 'Enter',
        metaKey: false,
      }),
    ).toBe(true)

    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: 'Shift',
        metaKey: false,
      }),
    ).toBe(false)

    expect(
      shouldHandleGridSurfaceKey({
        altKey: true,
        ctrlKey: false,
        key: '-',
        metaKey: true,
      }),
    ).toBe(true)

    expect(
      shouldHandleGridSurfaceKey({
        altKey: true,
        ctrlKey: false,
        key: 'ArrowLeft',
        metaKey: false,
      }),
    ).toBe(false)
  })
})
