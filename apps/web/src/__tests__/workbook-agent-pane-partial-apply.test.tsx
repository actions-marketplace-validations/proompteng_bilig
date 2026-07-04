// @vitest-environment jsdom
import {
  act,
  AgentHarness,
  agentStorageKey,
  createPreviewSummary,
  createRoot,
  createSnapshot,
  describe,
  expect,
  it,
  requestMethod,
  requestUrl,
  vi,
} from './workbook-agent-pane-test-helpers.js'

describe('workbook agent pane selected command apply', () => {
  it('re-previews and applies only the selected command subset', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fullPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C3',
          role: 'target' as const,
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 2,
        truncatedCellDiffs: false,
        inputChangeCount: 2,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const subsetPreview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          endAddress: 'C3',
          role: 'target' as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: 'Sheet1',
          address: 'C3',
          beforeInput: null,
          beforeFormula: null,
          afterInput: 2,
          afterFormula: null,
          changeKinds: ['input'],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 1,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const previewBundle = vi.fn(async (bundle) => (bundle.commands.length === 1 ? subsetPreview : fullPreview))
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-1',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Update two cells',
                summary: 'Write cells in Sheet1!B2 and 1 more change',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 3,
                createdAtUnixMs: 10,
                context: {
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
                },
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[1]],
                  },
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'C3',
                    values: [[2]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                    role: 'target',
                  },
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'C3',
                    endAddress: 'C3',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 2,
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-1/apply')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-2',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Update two cells',
                summary: 'Write cells in Sheet1!B2',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: {
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
                },
                commands: [
                  {
                    kind: 'writeRange',
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    values: [[1]],
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'B2',
                    endAddress: 'B2',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
              },
              executionRecords: [
                {
                  id: 'run-1',
                  bundleId: 'bundle-1',
                  documentId: 'doc-1',
                  threadId: 'thr-1',
                  turnId: 'turn-1',
                  actorUserId: 'user@example.com',
                  goalText: 'Update two cells',
                  planText: 'Apply only the second cell',
                  summary: 'Write cells in Sheet1!C3',
                  scope: 'sheet',
                  riskClass: 'medium',
                  acceptedScope: 'partial',
                  appliedBy: 'user',
                  baseRevision: 3,
                  appliedRevision: 4,
                  createdAtUnixMs: 10,
                  appliedAtUnixMs: 20,
                  context: {
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
                  },
                  commands: [
                    {
                      kind: 'writeRange',
                      sheetName: 'Sheet1',
                      startAddress: 'C3',
                      values: [[2]],
                    },
                  ],
                  preview: subsetPreview,
                },
              ],
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness previewCommandBundle={previewBundle} />)
    })

    expect(previewBundle).toHaveBeenCalledTimes(1)
    expect(host.textContent).toContain('2/2')

    const firstToggle = host.querySelector("[data-testid='workbook-agent-review-command-toggle-0']")
    expect(firstToggle instanceof HTMLInputElement).toBe(true)

    await act(async () => {
      firstToggle?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).toHaveBeenCalledTimes(2)
    expect(previewBundle.mock.calls[1]?.[0]).toEqual(
      expect.objectContaining({
        commands: [
          {
            kind: 'writeRange',
            sheetName: 'Sheet1',
            startAddress: 'C3',
            values: [[2]],
          },
        ],
      }),
    )
    expect(host.textContent).toContain('1/2')

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    expect(applyButton).toBeTruthy()

    await act(async () => {
      applyButton?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })

    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-1/apply'))
    expect(applyCall?.[1]?.body).toBe(
      JSON.stringify({
        appliedBy: 'user',
        commandIndexes: [1],
      }),
    )
    expect(host.textContent).toContain('Write cells in Sheet1!B2')
    expect(host.textContent).toContain('Sheet1!C3')
    expect(host.textContent).not.toContain('Recent changes')
    expect(host.textContent).not.toContain('Run again')

    await act(async () => {
      root.unmount()
    })
  })
})
