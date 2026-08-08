// @vitest-environment jsdom
import { isWorkbookAgentCommandBundle, toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
import { ValueTag } from '@bilig/protocol'
import { act, useCallback, useEffect, useRef, useState } from 'react'
import { toast } from 'sonner'
import { afterEach, beforeEach, vi } from 'vitest'
import { WorkbookToastRegion } from '../WorkbookToastRegion.js'
import { useWorkbookAgentPane } from '../use-workbook-agent-pane.js'
import { resetWorkbookAgentClientTransportStateForTests } from '../workbook-agent-client.js'
import { clearWorkbookAgentPreviewCache } from '../workbook-agent-preview-cache.js'
export { isWorkbookAgentCommandBundle, toWorkbookAgentReviewQueueItem, type WorkbookAgentCommandBundle } from '@bilig/agent-api'
export { ValueTag } from '@bilig/protocol'
export { act, useCallback, useEffect, useRef, useState } from 'react'
export { createRoot } from 'react-dom/client'
export { toast } from 'sonner'
export { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
export { WorkbookToastRegion } from '../WorkbookToastRegion.js'
export { useWorkbookAgentPane } from '../use-workbook-agent-pane.js'
export { resetWorkbookAgentClientTransportStateForTests } from '../workbook-agent-client.js'
export { clearWorkbookAgentPreviewCache } from '../workbook-agent-preview-cache.js'

export function agentStorageKey(userId = 'alex@example.com'): string {
  return `bilig:workbook-agent:doc-1:${encodeURIComponent(userId)}`
}

export async function flushToasts(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

export class MockEventSource {
  static latest: MockEventSource | null = null
  readonly url: string
  private readonly listeners = new Map<string, Set<(event: Event) => void>>()

  constructor(url: string) {
    this.url = url
    MockEventSource.latest = this
  }

  close() {}

  addEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type) ?? new Set()
    entries.add(listener)
    this.listeners.set(type, entries)
  }

  removeEventListener(type: string, listener: (event: Event) => void): void {
    const entries = this.listeners.get(type)
    if (!entries) {
      return
    }
    entries.delete(listener)
    if (entries.size === 0) {
      this.listeners.delete(type)
    }
  }

  emit(data: unknown): void {
    this.listeners.get('message')?.forEach((listener) => {
      listener(
        new MessageEvent('message', {
          data: JSON.stringify(data),
        }),
      )
    })
  }

  emitRaw(data: string): void {
    this.listeners.get('message')?.forEach((listener) => {
      listener(
        new MessageEvent('message', {
          data,
        }),
      )
    })
  }

  emitError(): void {
    this.listeners.get('error')?.forEach((listener) => {
      listener(new Event('error'))
    })
  }
}

export function requestUrl(input: RequestInfo | URL): string {
  if (typeof input === 'string') {
    return input
  }
  if (input instanceof URL) {
    return input.href
  }
  return input.url
}

export function requestBody(init: RequestInit | undefined): unknown {
  if (!init || typeof init.body !== 'string') {
    return null
  }
  return JSON.parse(init.body) as unknown
}

export function requestMethod(init: RequestInit | undefined): string {
  return init?.method ?? 'GET'
}

export function createDefaultWorkflowContext() {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'A1',
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 5,
    },
  }
}

export function createReviewQueueItem(bundle: WorkbookAgentCommandBundle) {
  return toWorkbookAgentReviewQueueItem({
    bundle,
    reviewMode: bundle.sharedReview ? 'ownerReview' : 'manual',
    ...(bundle.sharedReview ? { sharedReview: bundle.sharedReview } : {}),
  })
}

export function createSnapshot(overrides: Record<string, unknown> = {}) {
  const reviewBundleOverride = overrides['reviewBundle']
  const reviewQueueItemsOverride = overrides['reviewQueueItems']
  const { reviewBundle: _reviewBundle, ...restOverrides } = overrides
  const overrideEntries = Array.isArray(overrides['entries'])
    ? overrides['entries'].map((entry) =>
        typeof entry === 'object' && entry !== null && !('citations' in entry)
          ? {
              ...entry,
              citations: [],
            }
          : entry,
      )
    : undefined
  return {
    documentId: 'doc-1',
    threadId: 'thr-1',
    scope: 'private',
    executionPolicy: 'autoApplyAll',
    status: 'idle',
    activeTurnId: null,
    lastError: null,
    context: createDefaultWorkflowContext(),
    entries: [
      {
        id: 'assistant-1',
        kind: 'assistant',
        turnId: 'turn-1',
        text: '',
        phase: null,
        toolName: null,
        toolStatus: null,
        argumentsText: null,
        outputText: null,
        success: null,
        citations: [],
      },
    ],
    reviewQueueItems: Array.isArray(reviewQueueItemsOverride)
      ? reviewQueueItemsOverride
      : isWorkbookAgentCommandBundle(reviewBundleOverride)
        ? [createReviewQueueItem(reviewBundleOverride)]
        : [],
    executionRecords: [],
    workflowRuns: [],
    ...restOverrides,
    ...(overrideEntries ? { entries: overrideEntries } : {}),
  }
}

export function createPreviewSummary(overrides: Record<string, unknown> = {}) {
  return {
    ranges: [],
    structuralChanges: [],
    cellDiffs: [],
    effectSummary: {
      displayedCellDiffCount: 0,
      truncatedCellDiffs: false,
      inputChangeCount: 0,
      formulaChangeCount: 0,
      styleChangeCount: 0,
      numberFormatChangeCount: 0,
      structuralChangeCount: 0,
    },
    ...overrides,
  }
}

export function createThreadSummary(overrides: Record<string, unknown> = {}) {
  return {
    threadId: 'thr-1',
    scope: 'private',
    ownerUserId: 'alex@example.com',
    updatedAtUnixMs: 100,
    entryCount: 1,
    reviewQueueItemCount: typeof overrides['reviewQueueItemCount'] === 'number' ? overrides['reviewQueueItemCount'] : 0,
    latestEntryText: null,
    ...overrides,
  }
}

export interface MockZeroAgentHarness {
  readonly zero: {
    materialize(query: unknown): {
      readonly data: unknown
      addListener(listener: (value: unknown) => void): () => void
      destroy(): void
    }
  }
}

export function createMockZeroAgentHarness(input: {
  readonly initialThreadSummaries: unknown
  readonly initialWorkflowRuns: unknown
}): MockZeroAgentHarness {
  let threadSummaryValue = input.initialThreadSummaries
  let workflowRunValue = input.initialWorkflowRuns
  const threadSummaryListeners = new Set<(value: unknown) => void>()
  const workflowRunListeners = new Set<(value: unknown) => void>()
  let materializeCallCount = 0

  return {
    zero: {
      materialize(_query: unknown) {
        const isThreadSummaryQuery = materializeCallCount === 0
        materializeCallCount += 1
        return {
          get data() {
            return isThreadSummaryQuery ? threadSummaryValue : workflowRunValue
          },
          addListener(listener: (value: unknown) => void) {
            const listeners = isThreadSummaryQuery ? threadSummaryListeners : workflowRunListeners
            listeners.add(listener)
            return () => {
              listeners.delete(listener)
            }
          },
          destroy() {},
        }
      },
    },
  }
}

export function AgentHarness(props: {
  readonly currentUserId?: string
  readonly previewCommandBundle?: Parameters<typeof useWorkbookAgentPane>[0]['previewCommandBundle']
  readonly syncAuthoritativeRevision?: Parameters<typeof useWorkbookAgentPane>[0]['syncAuthoritativeRevision']
  readonly zero?: Parameters<typeof useWorkbookAgentPane>[0]['zero']
  readonly zeroEnabled?: boolean
  readonly apiEnabled?: boolean
  readonly showNewThreadControl?: boolean
}) {
  const { agentError, agentPanel, clearAgentError, startNewThread } = useWorkbookAgentPane({
    currentUserId: props.currentUserId ?? 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext: () => ({
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }),
    previewCommandBundle: props.previewCommandBundle ?? vi.fn(async () => createPreviewSummary()),
    ...(props.syncAuthoritativeRevision ? { syncAuthoritativeRevision: props.syncAuthoritativeRevision } : {}),
    ...(props.apiEnabled !== undefined ? { apiEnabled: props.apiEnabled } : {}),
    ...(props.zero ? { zero: props.zero } : {}),
    ...(props.zeroEnabled !== undefined ? { zeroEnabled: props.zeroEnabled } : {}),
  })

  return (
    <div>
      <WorkbookToastRegion
        toasts={
          agentError
            ? [
                {
                  id: 'agent-error',
                  tone: 'error',
                  message: agentError,
                  onDismiss: clearAgentError,
                },
              ]
            : []
        }
      />
      {props.showNewThreadControl ? (
        <button data-testid="test-start-new-thread" type="button" onClick={startNewThread}>
          New thread
        </button>
      ) : null}
      {agentPanel}
    </div>
  )
}

export function UnstableLiveThreadSummaryHarness(props: { readonly zero: Parameters<typeof useWorkbookAgentPane>[0]['zero'] }) {
  const getContext = useCallback(() => createDefaultWorkflowContext(), [])
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    applyContext: () => undefined,
    previewCommandBundle,
    syncAuthoritativeRevision: () => undefined,
    zero: props.zero,
    zeroEnabled: true,
  })

  return <div>{agentPanel}</div>
}

export function LaggyContextHarness() {
  const [selection, setSelection] = useState({
    sheetName: 'Sheet1',
    address: 'A1',
  })
  const selectionRef = useRef(selection)

  useEffect(() => {
    selectionRef.current = selection
  }, [selection])

  const getContext = useCallback(() => {
    const currentSelection = selectionRef.current
    return {
      selection: {
        sheetName: currentSelection.sheetName,
        address: currentSelection.address,
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }
  }, [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    previewCommandBundle: vi.fn(async () => createPreviewSummary()),
  })

  return (
    <div>
      <button
        data-testid="switch-context"
        type="button"
        onClick={() => {
          setSelection({
            sheetName: 'sheet3',
            address: 'A1',
          })
        }}
      >
        Switch
      </button>
      {agentPanel}
    </div>
  )
}

export function RapidSelectionContextHarness() {
  const [row, setRow] = useState(1)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      selection: {
        sheetName: 'Sheet1',
        address: `A${row}`,
      },
      viewport: {
        rowStart: row - 1,
        rowEnd: row + 9,
        colStart: 0,
        colEnd: 5,
      },
    }),
    [row],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: `Sheet1!A${row}`,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-selection-context"
        type="button"
        onClick={() => {
          setRow((current) => current + 1)
        }}
      >
        {row}
      </button>
      {agentPanel}
    </div>
  )
}

export function ToggleableContextSyncHarness() {
  const [enabled, setEnabled] = useState(true)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(() => createDefaultWorkflowContext(), [])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled,
    getContext,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="disable-agent-context"
        type="button"
        onClick={() => {
          setEnabled(false)
        }}
      >
        Disable
      </button>
      {agentPanel}
    </div>
  )
}

export function VolatileRenderedContextHarness() {
  const [renderCount, setRenderCount] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId: 11,
        selection: null,
        visibleRange: null,
      },
    }),
    [],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="force-render"
        type="button"
        onClick={() => {
          setRenderCount((current) => current + 1)
        }}
      >
        {renderCount}
      </button>
      {agentPanel}
    </div>
  )
}

export function VolatileRenderedBatchContextHarness() {
  const [batchId, setBatchId] = useState(11)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId,
        selection: null,
        visibleRange: null,
      },
    }),
    [batchId],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-render-batch"
        type="button"
        onClick={() => {
          setBatchId((current) => current + 1)
        }}
      >
        {batchId}
      </button>
      {agentPanel}
    </div>
  )
}

export function RapidRenderedRevisionContextHarness() {
  const [capturedRevision, setCapturedRevision] = useState(7)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision,
        batchId: capturedRevision,
        selection: null,
        visibleRange: null,
      },
    }),
    [capturedRevision],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-render-revision"
        type="button"
        onClick={() => {
          setCapturedRevision((current) => current + 1)
        }}
      >
        {capturedRevision}
      </button>
      {agentPanel}
    </div>
  )
}

export function RapidRenderedRangeContextHarness() {
  const [renderedVersion, setRenderedVersion] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 20 + renderedVersion,
        batchId: 20 + renderedVersion,
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'A1',
                input: `rendered-${renderedVersion}`,
                value: {
                  tag: ValueTag.String,
                  value: `rendered-${renderedVersion}`,
                  stringId: renderedVersion,
                },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }),
    [renderedVersion],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-rendered-range"
        type="button"
        onClick={() => {
          setRenderedVersion((current) => current + 1)
        }}
      >
        {renderedVersion}
      </button>
      {agentPanel}
    </div>
  )
}

export function VolatileRenderedStringIdContextHarness() {
  const [stringId, setStringId] = useState(1)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const getContext = useCallback(
    () => ({
      ...createDefaultWorkflowContext(),
      rendered: {
        capturedAtUnixMs: Date.now(),
        capturedRevision: 7,
        batchId: 11,
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'A1',
                input: 'same visible value',
                value: {
                  tag: ValueTag.String,
                  value: 'same visible value',
                  stringId,
                },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }),
    [stringId],
  )

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: 'Sheet1!A1',
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="advance-string-id"
        type="button"
        onClick={() => {
          setStringId((current) => current + 1)
        }}
      >
        {stringId}
      </button>
      {agentPanel}
    </div>
  )
}

export function VersionedContextRenderHarness(props: { readonly onBuildContext: (address: string) => void }) {
  const { onBuildContext } = props
  const [renderCount, setRenderCount] = useState(0)
  const [contextVersion, setContextVersion] = useState(0)
  const previewCommandBundle = useCallback(async () => createPreviewSummary(), [])
  const address = `A${String(contextVersion + 1)}`
  const getContext = useCallback(() => {
    onBuildContext(address)
    return {
      selection: {
        sheetName: 'Sheet1',
        address,
      },
      viewport: {
        rowStart: contextVersion,
        rowEnd: contextVersion + 10,
        colStart: 0,
        colEnd: 5,
      },
    }
  }, [address, contextVersion, onBuildContext])

  const { agentPanel } = useWorkbookAgentPane({
    currentUserId: 'alex@example.com',
    documentId: 'doc-1',
    enabled: true,
    getContext,
    activeContextLabel: `Sheet1!${address}`,
    contextVersion,
    previewCommandBundle,
  })

  return (
    <div>
      <button
        data-testid="force-versioned-context-render"
        type="button"
        onClick={() => {
          setRenderCount((current) => current + 1)
        }}
      >
        {renderCount}
      </button>
      <button
        data-testid="advance-versioned-context"
        type="button"
        onClick={() => {
          setContextVersion((current) => current + 1)
        }}
      >
        {contextVersion}
      </button>
      {agentPanel}
    </div>
  )
}

beforeEach(() => {
  vi.stubGlobal('EventSource', MockEventSource)
  window.sessionStorage.clear()
  clearWorkbookAgentPreviewCache()
})

afterEach(() => {
  toast.dismiss()
  vi.restoreAllMocks()
  resetWorkbookAgentClientTransportStateForTests()
  window.sessionStorage.clear()
  clearWorkbookAgentPreviewCache()
  document.body.innerHTML = ''
})
