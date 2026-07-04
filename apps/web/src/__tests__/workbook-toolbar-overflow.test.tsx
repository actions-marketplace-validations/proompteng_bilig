// @vitest-environment jsdom
import { ValueTag, type CellRangeRef, type CellStyleRecord } from '@bilig/protocol'
import { act, type MutableRefObject } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { useWorkbookToolbar } from '../use-workbook-toolbar.js'
import { WorkbookToolbar } from '../workbook-toolbar.js'

function ToolbarHookHarness(props: {
  readonly invokeMutation: (method: string, ...args: unknown[]) => Promise<void>
  readonly onRedo?: (() => void) | undefined
  readonly onUndo?: (() => void) | undefined
  readonly onCreateTableFromSelection?: (() => void) | undefined
  readonly selectionRangeRef: MutableRefObject<CellRangeRef>
  readonly selectedStyle?: CellStyleRecord | undefined
  readonly canHideCurrentRow?: boolean | undefined
  readonly canUnmergeSelection?: boolean | undefined
  readonly writesAllowed?: boolean | undefined
}) {
  const { ribbon } = useWorkbookToolbar({
    canHideCurrentColumn: false,
    canHideCurrentRow: props.canHideCurrentRow ?? false,
    canRedo: false,
    canUndo: false,
    canUnhideCurrentColumn: false,
    canUnhideCurrentRow: false,
    canUnmergeSelection: props.canUnmergeSelection ?? false,
    connectionStateName: 'connected',
    currentFillColor: '#ffffff',
    currentNumberFormatKind: 'general',
    currentTextColor: '#111827',
    horizontalAlignment: null,
    invokeMutation: props.invokeMutation,
    localPersistenceMode: 'ephemeral',
    onApplyBorderPreset: () => {},
    onClearStyle: () => {},
    onFillColorReset: () => {},
    onFillColorSelect: () => {},
    onFontSizeChange: () => {},
    onHideCurrentColumn: () => {},
    onHideCurrentRow: () => {},
    onCreateTableFromSelection: props.onCreateTableFromSelection,
    onHorizontalAlignmentChange: () => {},
    onNumberFormatChange: () => {},
    onRedo: props.onRedo ?? (() => {}),
    onTextColorReset: () => {},
    onTextColorSelect: () => {},
    onToggleBold: () => {},
    onToggleItalic: () => {},
    onToggleUnderline: () => {},
    onToggleWrap: () => {},
    onUndo: props.onUndo ?? (() => {}),
    onUnhideCurrentColumn: () => {},
    onUnhideCurrentRow: () => {},
    remoteSyncAvailable: true,
    runtimeReady: true,
    selectedCell: {
      address: 'A1',
      sheetName: 'Sheet1',
      flags: 0,
      value: { tag: ValueTag.Empty },
      version: 0,
    },
    selectedStyle: props.selectedStyle,
    selection: { sheetName: 'Sheet1' },
    selectionRangeRef: props.selectionRangeRef,
    trailingContent: null,
    writesAllowed: props.writesAllowed ?? true,
    zeroConfigured: true,
    zeroHealthReady: true,
  })

  return <>{ribbon}</>
}

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

function setScrollGeometry(
  element: Element,
  geometry: {
    readonly clientWidth: number
    readonly scrollLeft?: number | undefined
    readonly scrollWidth: number
  },
) {
  Object.defineProperties(element, {
    clientWidth: {
      configurable: true,
      value: geometry.clientWidth,
    },
    scrollLeft: {
      configurable: true,
      value: geometry.scrollLeft ?? 0,
      writable: true,
    },
    scrollWidth: {
      configurable: true,
      value: geometry.scrollWidth,
    },
  })
}
function createWorkbookKeyboardScopeHost() {
  const scope = document.createElement('div')
  scope.dataset['workbookKeyboardScope'] = 'true'
  const host = document.createElement('div')
  scope.appendChild(host)
  document.body.appendChild(scope)
  return host
}

async function flushToolbarMutationQueue(cycles = 3): Promise<void> {
  await Promise.resolve()
  if (cycles > 1) {
    await flushToolbarMutationQueue(cycles - 1)
  }
}

describe('WorkbookToolbar borders overflow and responsive cues', () => {
  it('keeps global formatting shortcut capture mounted across active style rerenders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)
    const addSpy = vi.spyOn(window, 'addEventListener')
    const removeSpy = vi.spyOn(window, 'removeEventListener')
    const keydownAdds = () => addSpy.mock.calls.filter(([type]) => type === 'keydown').length
    const keydownRemoves = () => removeSpy.mock.calls.filter(([type]) => type === 'keydown').length

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(0)

    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{ id: 'style-bold', font: { bold: true } }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })
    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{ id: 'style-bold-italic', font: { bold: true, italic: true } }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(0)

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
    })

    expect(invokeMutation).toHaveBeenLastCalledWith(
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { underline: true },
      },
    )

    await act(async () => {
      root.unmount()
    })

    expect(keydownAdds()).toBe(1)
    expect(keydownRemoves()).toBe(1)
  })

  it('optimistically marks formatting shortcut buttons while mutations are pending', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const pendingMutationResolvers: Array<() => void> = []
    const pendingMutations: Array<Promise<void>> = []
    const invokeMutation = vi.fn(() => {
      const pendingMutation = new Promise<void>((resolve) => {
        pendingMutationResolvers.push(resolve)
      })
      pendingMutations.push(pendingMutation)
      return pendingMutation
    })
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'b', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      1,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { bold: true },
      },
    )

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'i', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)

    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'u', metaKey: true }))
      await flushToolbarMutationQueue()
    })
    expect(host.querySelector("[aria-label='Bold']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Italic']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(host.querySelector("[aria-label='Underline']")?.className).toContain('bg-[var(--wb-accent-soft)]')
    expect(invokeMutation).toHaveBeenCalledTimes(1)

    await act(async () => {
      pendingMutationResolvers[0]?.()
      await flushToolbarMutationQueue()
    })
    expect(invokeMutation).toHaveBeenCalledTimes(2)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      2,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { italic: true },
      },
    )

    await act(async () => {
      pendingMutationResolvers[1]?.()
      await flushToolbarMutationQueue()
    })
    expect(invokeMutation).toHaveBeenCalledTimes(3)
    expect(invokeMutation).toHaveBeenNthCalledWith(
      3,
      'setRangeStyle',
      {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      {
        font: { underline: true },
      },
    )

    await act(async () => {
      pendingMutationResolvers[2]?.()
      await flushToolbarMutationQueue()
      await Promise.all(pendingMutations)
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('marks the border menu active when the selected cell has borders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <ToolbarHookHarness
          invokeMutation={invokeMutation}
          selectedStyle={{
            id: 'style-bordered',
            borders: {
              top: {
                color: '#111827',
                style: 'solid',
                weight: 'thin',
              },
            },
          }}
          selectionRangeRef={selectionRangeRef}
        />,
      )
    })

    const borderTrigger = host.querySelector("[aria-label='Borders']")
    expect(borderTrigger?.getAttribute('aria-pressed')).toBe('true')
    expect(borderTrigger?.className).toContain('bg-[var(--wb-accent-soft)]')

    await act(async () => {
      root.unmount()
    })
  })

  it('applies bottom border presets to the live selection range', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const invokeMutation = vi.fn(async () => {})
    const selectionRangeRef: MutableRefObject<CellRangeRef> = {
      current: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
    }
    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(<ToolbarHookHarness invokeMutation={invokeMutation} selectionRangeRef={selectionRangeRef} />)
    })

    selectionRangeRef.current = {
      sheetName: 'Sheet1',
      startAddress: 'B2',
      endAddress: 'D5',
    }

    await act(async () => {
      host.querySelector("[aria-label='Borders']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    await act(async () => {
      document.querySelector("[aria-label='Bottom border']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(invokeMutation.mock.calls).toEqual([
      [
        'clearRangeStyle',
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'D5',
        },
        ['borderTop', 'borderRight', 'borderBottom', 'borderLeft'],
      ],
      [
        'setRangeStyle',
        {
          sheetName: 'Sheet1',
          startAddress: 'B5',
          endAddress: 'D5',
        },
        {
          borders: {
            bottom: {
              style: 'solid',
              weight: 'thin',
              color: '#111827',
            },
          },
        },
      ],
    ])

    await act(async () => {
      root.unmount()
    })
  })

  it('hides native toolbar overflow scrollbars while preserving horizontal scrolling', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      )
    })

    const toolbar = host.querySelector("[aria-label='Formatting toolbar']")
    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    expect(toolbar?.className).toContain('overflow-hidden')
    expect(formattingScroll?.className).toContain('overflow-x-auto')
    expect(formattingScroll?.className).toContain('overflow-y-hidden')
    expect(formattingScroll?.className).toContain('wb-scrollbar-none')

    await act(async () => {
      root.unmount()
    })
  })

  it('signals hidden formatting actions on narrow toolbar scroll regions', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      )
    })

    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    if (!formattingScroll) {
      throw new Error('Expected formatting toolbar scroll region to render')
    }

    setScrollGeometry(formattingScroll, { clientWidth: 220, scrollWidth: 870 })
    const scrollBy = vi.fn()
    Object.defineProperty(formattingScroll, 'scrollBy', {
      configurable: true,
      value: scrollBy,
    })

    await act(async () => {
      window.dispatchEvent(new Event('resize'))
    })

    const cue = host.querySelector("[data-testid='toolbar-overflow-cue']")
    expect(host.querySelector("[data-testid='toolbar-overflow-back-cue']")).toBeNull()
    expect(cue).not.toBeNull()
    expect(cue?.getAttribute('aria-label')).toBe('Show more toolbar actions')
    expect(formattingScroll.className).not.toContain('pr-7')
    expect(cue?.className).not.toContain('absolute')
    expect(cue?.className).toContain('flex-none')
    expect(cue?.className).toContain('text-[var(--wb-accent)]')

    await act(async () => {
      cue?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(scrollBy).toHaveBeenCalledWith({
      behavior: 'smooth',
      left: 220,
    })

    setScrollGeometry(formattingScroll, {
      clientWidth: 220,
      scrollLeft: 650,
      scrollWidth: 870,
    })

    await act(async () => {
      formattingScroll?.dispatchEvent(new Event('scroll'))
    })

    expect(host.querySelector("[data-testid='toolbar-overflow-cue']")).toBeNull()
    const backCue = host.querySelector("[data-testid='toolbar-overflow-back-cue']")
    expect(backCue).not.toBeNull()
    expect(backCue?.getAttribute('aria-label')).toBe('Show previous toolbar actions')
    expect(backCue?.className).toContain('flex-none')
    expect(backCue?.className).toContain('border-r')

    await act(async () => {
      backCue?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    expect(scrollBy).toHaveBeenLastCalledWith({
      behavior: 'smooth',
      left: -275,
    })

    await act(async () => {
      root.unmount()
    })
  })

  it('refreshes toolbar overflow cues when mounted toolbar contents change size', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

    const host = createWorkbookKeyboardScopeHost()
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookToolbar
          canHideCurrentColumn={false}
          canHideCurrentRow={false}
          canMergeSelection={false}
          canUnmergeSelection={false}
          canRedo={false}
          canUndo={false}
          canUnhideCurrentColumn={false}
          canUnhideCurrentRow={false}
          currentFillColor="#ffffff"
          currentNumberFormatKind="general"
          currentTextColor="#111827"
          horizontalAlignment={null}
          isBoldActive={false}
          isItalicActive={false}
          isUnderlineActive={false}
          isWrapActive={false}
          onApplyBorderPreset={() => {}}
          onClearStyle={() => {}}
          onFillColorReset={() => {}}
          onFillColorSelect={() => {}}
          onFontSizeChange={() => {}}
          onHideCurrentColumn={() => {}}
          onHideCurrentRow={() => {}}
          onMergeSelectedCells={() => {}}
          onHorizontalAlignmentChange={() => {}}
          onNumberFormatChange={() => {}}
          onRedo={() => {}}
          onTextColorReset={() => {}}
          onTextColorSelect={() => {}}
          onToggleBold={() => {}}
          onToggleItalic={() => {}}
          onToggleUnderline={() => {}}
          onToggleWrap={() => {}}
          onUndo={() => {}}
          onUnmergeSelectedCells={() => {}}
          onUnhideCurrentColumn={() => {}}
          onUnhideCurrentRow={() => {}}
          recentFillColors={[]}
          recentTextColors={[]}
          selectedFontSize="11"
          trailingContent={<div>Trailing</div>}
          writesAllowed
        />,
      )
    })

    const formattingScroll = host.querySelector("[data-testid='toolbar-formatting-scroll']")
    if (!formattingScroll) {
      throw new Error('Expected formatting toolbar scroll region to render')
    }

    setScrollGeometry(formattingScroll, { clientWidth: 420, scrollWidth: 420 })
    await act(async () => {
      window.dispatchEvent(new Event('resize'))
    })
    expect(host.querySelector("[data-testid='toolbar-overflow-cue']")).toBeNull()

    setScrollGeometry(formattingScroll, { clientWidth: 220, scrollWidth: 870 })
    await act(async () => {
      formattingScroll.appendChild(document.createElement('span'))
      await Promise.resolve()
    })

    expect(host.querySelector("[data-testid='toolbar-overflow-cue']")).not.toBeNull()

    await act(async () => {
      root.unmount()
    })
  })
})
