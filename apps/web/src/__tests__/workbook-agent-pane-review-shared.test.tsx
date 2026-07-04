// @vitest-environment jsdom
import {
  act,
  AgentHarness,
  agentStorageKey,
  createPreviewSummary,
  createRoot,
  createSnapshot,
  createThreadSummary,
  describe,
  expect,
  it,
  MockEventSource,
  RapidRenderedRangeContextHarness,
  requestBody,
  requestMethod,
  requestUrl,
  vi,
  VolatileRenderedStringIdContextHarness,
} from './workbook-agent-pane-test-helpers.js'

describe('workbook agent pane review and shared-thread controls', () => {
  it('requests the exact authoritative revision from applied assistant execution records', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const syncAuthoritativeRevision = vi.fn(async () => undefined)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
              executionRecords: [
                {
                  id: 'run-1',
                  appliedRevision: 12,
                },
              ],
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    try {
      await act(async () => {
        root.render(<AgentHarness syncAuthoritativeRevision={syncAuthoritativeRevision} />)
      })
      await act(async () => {
        await Promise.resolve()
        await Promise.resolve()
      })

      expect(syncAuthoritativeRevision).toHaveBeenCalledTimes(1)
      expect(syncAuthoritativeRevision).toHaveBeenCalledWith(12)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not treat rendered range churn as immediate context and flood sync requests', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidRenderedRangeContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRenderedRange = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-rendered-range']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await advanceRenderedRange()
      await advanceRenderedRange()
      await advanceRenderedRange()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 1_000))
      })

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 20,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('throttles rendered range context churn while the assistant turn is active', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              activeTurnId: 'turn-1',
              status: 'inProgress',
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidRenderedRangeContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRenderedRange = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-rendered-range']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 800))
        })
      }

      await advanceRenderedRange()
      await advanceRenderedRange()
      await advanceRenderedRange()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 3_000))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 23,
            batchId: 23,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  }, 10_000)

  it('does not resync workbook context only because rendered string intern ids change', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<VolatileRenderedStringIdContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceStringId = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-string-id']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 220))
        })
      }

      await advanceStringId()
      await advanceStringId()
      await advanceStringId()

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('recreates the assistant session and reconnects the stream after a stale session error', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    let resumeCount = 0
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        resumeCount += 1
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      MockEventSource.latest?.emitError()
      await Promise.resolve()
      await Promise.resolve()
    })

    const sessionCalls = fetchSpy.mock.calls.filter(
      ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET',
    )
    expect(sessionCalls).toHaveLength(2)
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')
    expect(window.sessionStorage.getItem(agentStorageKey())).toBe(
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('bootstraps from a stored durable thread id without requiring a stored session id', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
            }),
          ),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const bootstrapSessionCall = fetchSpy.mock.calls.find(([input, init]) => {
      return requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET'
    })
    expect(bootstrapSessionCall).toBeDefined()
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')
    expect(window.sessionStorage.getItem(agentStorageKey())).toContain('"threadId":"thr-1"')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not render private review controls for restored private review items', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
          role: 'target' as const,
        },
      ],
      cellDiffs: [
        {
          sheetName: 'Sheet1',
          address: 'A1',
          beforeInput: 1,
          beforeFormula: null,
          afterInput: 1,
          afterFormula: null,
          changeKinds: ['style'],
        },
      ],
      effectSummary: {
        displayedCellDiffCount: 1,
        truncatedCellDiffs: false,
        inputChangeCount: 0,
        formulaChangeCount: 0,
        styleChangeCount: 1,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              reviewBundle: {
                id: 'bundle-1',
                documentId: 'doc-1',
                threadId: 'thr-1',
                turnId: 'turn-1',
                goalText: 'Bold the selected cell',
                summary: 'Format Sheet1!A1',
                scope: 'selection',
                riskClass: 'low',
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
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A1',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A1',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
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
              reviewQueueItems: [],
              executionRecords: [
                {
                  id: 'run-1',
                  bundleId: 'bundle-1',
                  documentId: 'doc-1',
                  threadId: 'thr-1',
                  turnId: 'turn-1',
                  actorUserId: 'user@example.com',
                  goalText: 'Bold the selected cell',
                  planText: 'Apply bold formatting',
                  summary: 'Format Sheet1!A1',
                  scope: 'selection',
                  riskClass: 'low',
                  acceptedScope: 'full',
                  appliedBy: 'auto',
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
                      kind: 'formatRange',
                      range: {
                        sheetName: 'Sheet1',
                        startAddress: 'A1',
                        endAddress: 'A1',
                      },
                      patch: {
                        font: {
                          bold: true,
                        },
                      },
                    },
                  ],
                  preview,
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

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).not.toHaveBeenCalled()
    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-1/apply'))
    expect(applyCall).toBeUndefined()
    expect(host.textContent).not.toContain('Apply')
    expect(host.textContent).not.toContain('Executions')
    expect(host.textContent).not.toContain('Replay')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not auto-apply low-risk review items on shared threads', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      ranges: [
        {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
          role: 'target' as const,
        },
      ],
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-1',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-1',
                goalText: 'Bold the selected cell',
                summary: 'Format Sheet1!A1',
                scope: 'selection',
                riskClass: 'low',
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
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A1',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A1',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 1,
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness previewCommandBundle={previewBundle} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    expect(previewBundle).toHaveBeenCalled()
    expect(previewBundle.mock.calls[0]?.[0]).toMatchObject({
      id: 'bundle-shared-1',
    })
    const applyCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-1/apply'))
    expect(applyCall).toBeUndefined()

    await act(async () => {
      root.unmount()
    })
  })

  it('blocks collaborator approval of shared medium-risk bundles in the panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey('casey@example.com'),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Format selected range'],
    })
    const previewBundle = vi.fn(async () => preview)
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-2',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
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
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A20',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="casey@example.com" previewCommandBundle={previewBundle} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected apply button to render')
    }
    expect(applyButton.disabled).toBe(true)
    expect(host.textContent).toContain('Owner review routes medium/high-risk changes to Alex on this shared thread.')
    expect(host.textContent).toContain('Owner review is in progress with Alex.')

    await act(async () => {
      root.unmount()
    })
  })

  it('lets the shared thread owner approve a medium-risk bundle before apply', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Normalize selected range'],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-owner',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
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
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [
                  {
                    sheetName: 'Sheet1',
                    startAddress: 'A1',
                    endAddress: 'A20',
                    role: 'target',
                  },
                ],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-shared-owner/review')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-owner',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'approved',
                  decidedByUserId: 'alex@example.com',
                  decidedAtUnixMs: 25,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="alex@example.com" previewCommandBundle={vi.fn(async () => preview)} />)
    })

    await act(async () => {
      await Promise.resolve()
    })

    const applyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    const approveButton = host.querySelector("[data-testid='workbook-agent-review-item-approve']")
    if (!(applyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected apply button')
    }
    if (!(approveButton instanceof HTMLButtonElement)) {
      throw new Error('Expected approve button')
    }
    expect(applyButton.disabled).toBe(true)

    await act(async () => {
      approveButton.click()
    })

    const reviewCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-owner/review'))
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: 'approved',
    })
    expect(host.textContent).toContain('Approved by Alex.')
    const refreshedApplyButton = host.querySelector("[data-testid='workbook-agent-apply-review-item']")
    if (!(refreshedApplyButton instanceof HTMLButtonElement)) {
      throw new Error('Expected refreshed apply button')
    }
    expect(refreshedApplyButton.disabled).toBe(false)

    await act(async () => {
      root.unmount()
    })
  })

  it('lets collaborators recommend approval on shared medium-risk bundles', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey('casey@example.com'),
      JSON.stringify({
        threadId: 'thr-shared',
      }),
    )
    const preview = createPreviewSummary({
      structuralChanges: ['Normalize selected range'],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && (init?.method ?? 'GET') === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              ownerUserId: 'alex@example.com',
              entryCount: 3,
              reviewQueueItemCount: 1,
              latestEntryText: 'Review item queued',
            }),
          ]),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/chat/threads/thr-shared') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-collab',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      if (url.endsWith('/review-items/bundle-shared-collab/review')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-shared',
              scope: 'shared',
              reviewBundle: {
                id: 'bundle-shared-collab',
                documentId: 'doc-1',
                threadId: 'thr-shared',
                turnId: 'turn-2',
                goalText: 'Normalize the imported sheet',
                summary: 'Normalize Sheet1!A1:A20',
                scope: 'sheet',
                riskClass: 'medium',
                baseRevision: 4,
                createdAtUnixMs: 20,
                context: null,
                commands: [
                  {
                    kind: 'formatRange',
                    range: {
                      sheetName: 'Sheet1',
                      startAddress: 'A1',
                      endAddress: 'A20',
                    },
                    patch: {
                      font: {
                        bold: true,
                      },
                    },
                  },
                ],
                affectedRanges: [],
                estimatedAffectedCells: 20,
                sharedReview: {
                  ownerUserId: 'alex@example.com',
                  status: 'pending',
                  decidedByUserId: null,
                  decidedAtUnixMs: null,
                  recommendations: [
                    {
                      userId: 'casey@example.com',
                      decision: 'approved',
                      decidedAtUnixMs: 30,
                    },
                  ],
                },
              },
            }),
          ),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness currentUserId="casey@example.com" previewCommandBundle={async () => preview} />)
    })

    const approveButton = host.querySelector("[data-testid='workbook-agent-review-item-approve']")
    const clearButton = host.querySelector("[data-testid='workbook-agent-dismiss-review-item']")
    expect(approveButton instanceof HTMLButtonElement).toBe(true)
    expect(clearButton instanceof HTMLButtonElement).toBe(true)
    expect(clearButton instanceof HTMLButtonElement ? clearButton.disabled : false).toBe(true)
    expect(host.textContent).toContain('Owner review is in progress with Alex.')

    await act(async () => {
      if (!(approveButton instanceof HTMLButtonElement)) {
        throw new Error('Expected recommend approve button')
      }
      approveButton.click()
    })

    const reviewCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/review-items/bundle-shared-collab/review'))
    expect(requestBody(reviewCall?.[1])).toEqual({
      decision: 'approved',
    })
    expect(host.textContent).toContain('1 approval recommendation')
    expect(host.textContent).toContain('You recommended approval.')

    await act(async () => {
      root.unmount()
    })
  })
})
