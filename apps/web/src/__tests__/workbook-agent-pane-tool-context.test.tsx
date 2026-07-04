// @vitest-environment jsdom
import { captureExpectedConsoleDebug } from './expected-console.js'
import {
  act,
  AgentHarness,
  agentStorageKey,
  createMockZeroAgentHarness,
  createRoot,
  createSnapshot,
  createThreadSummary,
  describe,
  expect,
  flushToasts,
  it,
  LaggyContextHarness,
  MockEventSource,
  RapidRenderedRevisionContextHarness,
  RapidSelectionContextHarness,
  requestBody,
  requestMethod,
  requestUrl,
  ToggleableContextSyncHarness,
  UnstableLiveThreadSummaryHarness,
  VersionedContextRenderHarness,
  vi,
  VolatileRenderedBatchContextHarness,
  VolatileRenderedContextHarness,
} from './workbook-agent-pane-test-helpers.js'

function expectLoggedErrorMessage(payload: unknown, expectedMessage: string): void {
  expect(payload).toMatchObject({ error: expect.any(Error) })
  if (typeof payload !== 'object' || payload === null || !('error' in payload) || !(payload.error instanceof Error)) {
    throw new Error('Expected console debug payload to include an Error')
  }
  expect(payload.error.message).toBe(expectedMessage)
}

describe('workbook agent pane tool rows and context sync', () => {
  it('summarizes attached selection ranges in tool rows', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'tool-context',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'get_context',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      selection: {
                        sheetName: 'Sheet1',
                        address: 'E20',
                        range: {
                          startAddress: 'C11',
                          endAddress: 'F20',
                        },
                      },
                      visibleRange: {
                        sheetName: 'Sheet1',
                        startAddress: 'A1',
                        endAddress: 'J20',
                      },
                    }),
                    success: true,
                  },
                ],
              }),
            ),
            {
              status: 200,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('Get Context')
    expect(host.textContent).toContain('Sheet1!C11:F20')
    expect(host.textContent).not.toContain('Sheet1!E20')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not poll assistant thread APIs when the runtime reports the assistant service disabled', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async () => new Response(JSON.stringify([]), { status: 200 }))
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness apiEnabled={false} />)
    })

    expect(fetchSpy).not.toHaveBeenCalled()
    expect(host.textContent).not.toContain('No messages yet')

    await act(async () => {
      root.unmount()
    })
  })

  it('hides raw app-server protocol errors behind user-facing copy', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify({
              error: 'WORKBOOK_AGENT_RUNTIME_UNAVAILABLE',
              message: 'thread/start.dynamicTools requires experimentalApi capability',
              retryable: true,
            }),
            {
              status: 503,
              headers: { 'content-type': 'application/json' },
            },
          ),
      ),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const input = host.querySelector("[data-testid='workbook-agent-input']")
    expect(input instanceof HTMLTextAreaElement).toBe(true)

    await act(async () => {
      if (!(input instanceof HTMLTextAreaElement)) {
        throw new Error('Agent input not found')
      }
      const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
      const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
      if (typeof valueSetter !== 'function') {
        throw new Error('Textarea value setter not found')
      }
      Reflect.apply(valueSetter, input, ['Summarize this sheet'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })

    const submit = host.querySelector("[data-testid='workbook-agent-send']")
    await act(async () => {
      if (!(submit instanceof HTMLButtonElement)) {
        throw new Error('Send button not found')
      }
      submit.click()
    })
    await flushToasts()

    expect(host.textContent).toContain('Retry in a moment.')
    expect(host.textContent).not.toContain('thread/start.dynamicTools requires experimentalApi capability')

    await act(async () => {
      root.unmount()
    })
  })

  it('bootstraps the assistant session and streams assistant deltas into the rail', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Thinking')
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'assistant-1',
        turnId: 'turn-1',
        entryKind: 'assistant',
        delta: 'Updated Sheet1',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Updated Sheet1')

    await act(async () => {
      root.unmount()
    })
  })

  it('surfaces malformed assistant stream payloads with stable copy', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(JSON.stringify([]), { status: 200, headers: { 'content-type': 'application/json' } })
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    await act(async () => {
      MockEventSource.latest?.emitRaw('{')
    })
    await flushToasts()

    expect(host.textContent).toContain('Assistant stream returned malformed event data.')
    expect(host.textContent).not.toContain('SyntaxError')

    await act(async () => {
      root.unmount()
    })
  })

  it('streams command execution output deltas into command tool rows', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'cmd-1',
                    kind: 'system',
                    turnId: 'turn-1',
                    text: 'Codex emitted commandExecution.',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
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
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Codex emitted commandExecution.')
    expect(host.querySelector("[data-testid='workbook-agent-empty-state']")).toBeNull()

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryToolOutputDelta',
        itemId: 'cmd-1',
        turnId: 'turn-1',
        delta: 'hi\n',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Command')
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Codex emitted commandExecution.')

    const toggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-cmd-1']")
    expect(toggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(toggle instanceof HTMLButtonElement)) {
        throw new Error('Command execution toggle not found')
      }
      toggle.click()
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('hi')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders reasoning text immediately from streamed deltas without waiting for a snapshot refresh', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'optimistic-user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Check version issues',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
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
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).not.toContain('Thought')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'reasoning-1',
        turnId: 'turn-1',
        entryKind: 'reasoning',
        delta: 'Examining version issues',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Thought')
    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain('Examining version issues')

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'entryTextDelta',
        itemId: 'reasoning-1',
        turnId: 'turn-1',
        entryKind: 'reasoning',
        delta: ' before deciding whether staged changes must be cleared.',
      })
    })

    expect(host.querySelector("[data-testid='workbook-agent-panel']")?.textContent).toContain(
      'Examining version issues before deciding whether staged changes must be cleared.',
    )

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps the thinking row visible while tool activity is still streaming', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
          return new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'optimistic-user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Build the operating plan',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [],
                  },
                  {
                    id: 'tool-1',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: '',
                    phase: null,
                    toolName: 'bilig_read_workbook',
                    toolStatus: 'completed',
                    argumentsText: null,
                    outputText: null,
                    success: true,
                    citations: [],
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
        return new Response(
          JSON.stringify({
            ok: true,
          }),
          { status: 200, headers: { 'content-type': 'application/json' } },
        )
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).toContain('Read Workbook')
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull()
    expect(host.textContent).toContain('Thinking')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not refetch thread summaries when stream snapshots arrive', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify([createThreadSummary({ threadId: 'thr-1' })]), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
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

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)

    await act(async () => {
      MockEventSource.latest?.emit({
        type: 'snapshot',
        snapshot: createSnapshot({
          threadId: 'thr-1',
          status: 'inProgress',
          activeTurnId: 'turn-2',
        }),
      })
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('does not restart live thread bootstrap when callback props churn across internal rerenders', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [],
      initialWorkflowRuns: [],
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
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

    await act(async () => {
      root.render(<UnstableLiveThreadSummaryHarness zero={zero.zero} />)
      await Promise.resolve()
      await Promise.resolve()
    })

    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET'),
    ).toHaveLength(1)
    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(0)
    expect(MockEventSource.latest?.url).toContain('/v2/documents/doc-1/chat/threads/thr-1/events')

    await act(async () => {
      root.unmount()
    })
  })

  it('syncs the latest workbook context after a sheet change even when the context getter reads laggy refs', async () => {
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

    try {
      await act(async () => {
        root.render(<LaggyContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      await act(async () => {
        host.querySelector("[data-testid='switch-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      const contextCalls = fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )
      expect(contextCalls.length).toBeGreaterThan(0)
      expect(requestBody(contextCalls.at(-1)?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'sheet3',
            address: 'A1',
          },
        },
      })
    } finally {
      // no-op
    }

    await act(async () => {
      root.unmount()
    })
  })

  it('keeps workbook context sync single-flight when selection changes faster than the backend responds', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const contextResponses: Array<() => void> = []
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Promise<Response>((resolve) => {
          contextResponses.push(() => {
            resolve(
              new Response(JSON.stringify({ ok: true }), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              }),
            )
          })
        })
      }
      return Promise.reject(new Error(`Unexpected fetch to ${url}`))
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    const advanceSelection = async () => {
      await act(async () => {
        host.querySelector("[data-testid='advance-selection-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })
      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })
    }

    try {
      await act(async () => {
        root.render(<RapidSelectionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      await advanceSelection()
      await advanceSelection()
      await advanceSelection()

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        contextResponses[0]?.()
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A4',
          },
        },
      })

      await act(async () => {
        contextResponses[1]?.()
        await Promise.resolve()
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('retries workbook context sync after a failed server response without marking stale context as synced', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let contextAttempts = 0
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        contextAttempts += 1
        if (contextAttempts === 1) {
          return new Response(JSON.stringify({ message: 'temporary context failure' }), {
            status: 503,
            headers: { 'content-type': 'application/json' },
          })
        }
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)
    const contextSyncDebug = captureExpectedConsoleDebug('Failed to sync agent context update')

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<RapidSelectionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 1_300))
      })

      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A1',
          },
        },
      })
      contextSyncDebug.expectLogCount(1)
      expectLoggedErrorMessage(contextSyncDebug.payloads()[0], 'temporary context failure')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not retry a failed in-flight context sync after the assistant is disabled', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let failContextSync: ((error: Error) => void) | null = null
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1' })), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST') {
        return new Promise<Response>((_resolve, reject) => {
          failContextSync = reject
        })
      }
      return Promise.reject(new Error(`Unexpected fetch to ${url}`))
    })
    vi.stubGlobal('fetch', fetchSpy)
    const contextSyncDebug = captureExpectedConsoleDebug('Failed to sync agent context update')

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const contextCalls = () =>
      fetchSpy.mock.calls.filter(
        ([input, init]) => requestUrl(input).endsWith('/chat/threads/thr-1/context') && requestMethod(init) === 'POST',
      )

    try {
      await act(async () => {
        root.render(<ToggleableContextSyncHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 220))
      })

      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        host.querySelector("[data-testid='disable-agent-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        await Promise.resolve()
      })

      await act(async () => {
        failContextSync?.(new Error('context transport down'))
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
      contextSyncDebug.expectLogCount(1)
      expectLoggedErrorMessage(contextSyncDebug.payloads()[0], 'context transport down')
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because the rendered capture timestamp changes', async () => {
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
        root.render(<VolatileRenderedContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const forceInertRender = async () => {
        await act(async () => {
          host.querySelector("[data-testid='force-render']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await forceInertRender()
      await forceInertRender()
      await forceInertRender()

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
            batchId: 11,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not rebuild workbook context on unrelated assistant pane renders when a context version is provided', async () => {
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
    const buildContext = vi.fn()
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
        root.render(<VersionedContextRenderHarness onBuildContext={buildContext} />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(1)
      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        const button = host.querySelector("[data-testid='force-versioned-context-render']")
        for (let index = 0; index < 20; index += 1) {
          button?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        }
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(1)
      expect(contextCalls()).toHaveLength(1)

      await act(async () => {
        host.querySelector("[data-testid='advance-versioned-context']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(buildContext).toHaveBeenCalledTimes(2)
      expect(buildContext).toHaveBeenLastCalledWith('A2')
      expect(contextCalls()).toHaveLength(2)
      expect(requestBody(contextCalls()[1]?.[1])).toMatchObject({
        context: {
          selection: {
            address: 'A2',
            sheetName: 'Sheet1',
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because the rendered batch id changes', async () => {
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
        root.render(<VolatileRenderedBatchContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceBatch = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-render-batch']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 200))
        })
      }

      await advanceBatch()
      await advanceBatch()
      await advanceBatch()

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
            batchId: 11,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not resync workbook context only because rendered proof revisions change', async () => {
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
        root.render(<RapidRenderedRevisionContextHarness />)
      })

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 200))
      })

      expect(contextCalls()).toHaveLength(1)

      const advanceRevision = async () => {
        await act(async () => {
          host.querySelector("[data-testid='advance-render-revision']")?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
        })
        await act(async () => {
          await Promise.resolve()
          await new Promise((resolve) => setTimeout(resolve, 100))
        })
      }
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()
      await advanceRevision()

      await act(async () => {
        await Promise.resolve()
        await new Promise((resolve) => setTimeout(resolve, 900))
      })

      expect(contextCalls()).toHaveLength(1)
      expect(requestBody(contextCalls()[0]?.[1])).toMatchObject({
        context: {
          rendered: {
            capturedRevision: 7,
          },
        },
      })
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('does not poll authoritative revisions only because an assistant turn is in progress', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.useFakeTimers()
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
              status: 'inProgress',
              activeTurnId: 'turn-1',
              executionRecords: [],
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

      expect(syncAuthoritativeRevision).not.toHaveBeenCalled()

      await act(async () => {
        await vi.advanceTimersByTimeAsync(6_000)
      })

      expect(syncAuthoritativeRevision).not.toHaveBeenCalled()
    } finally {
      await act(async () => {
        root.unmount()
      })
      vi.useRealTimers()
    }
  })
})
