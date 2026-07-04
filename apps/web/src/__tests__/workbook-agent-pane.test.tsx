// @vitest-environment jsdom
import {
  act,
  AgentHarness,
  agentStorageKey,
  createDefaultWorkflowContext,
  createMockZeroAgentHarness,
  createRoot,
  createSnapshot,
  createThreadSummary,
  describe,
  expect,
  flushToasts,
  it,
  MockEventSource,
  requestBody,
  requestMethod,
  requestUrl,
  vi,
} from './workbook-agent-pane-test-helpers.js'

describe('workbook agent pane rendering composer and stream basics', () => {
  it('renders the assistant panel without the skill-card strip', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(JSON.stringify(createSnapshot()), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
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
    expect(input instanceof HTMLTextAreaElement ? input.value : '').toBe('')
    expect(host.textContent).not.toContain('Local Skills')
    expect(host.textContent).not.toContain('Inspect Selection')
    expect(host.textContent).not.toContain('Ask the assistant to inspect, edit, or restructure this workbook.')
    expect(host.querySelector("[data-testid='workbook-agent-empty-state']")).toBeNull()
    expect(host.textContent).not.toContain('No messages yet')
    expect(host.textContent).not.toContain('Active context: Sheet1!A1')
    expect(input instanceof HTMLTextAreaElement ? input.getAttribute('placeholder') : null).toBe('Ask the workbook assistant')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders durable thread summaries and workflow runs from Zero projections', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const zero = createMockZeroAgentHarness({
      initialThreadSummaries: [
        createThreadSummary({
          threadId: 'thr-1',
          scope: 'shared',
          ownerUserId: 'casey@example.com',
          latestEntryText: 'Completed workflow: Summarize Workbook',
        }),
      ],
      initialWorkflowRuns: [
        {
          runId: 'wf-zero-1',
          threadId: 'thr-1',
          startedByUserId: 'casey@example.com',
          workflowTemplate: 'summarizeWorkbook',
          title: 'Summarize Workbook',
          summary: 'Summarized workbook structure across 2 sheets.',
          status: 'completed',
          createdAtUnixMs: 1,
          updatedAtUnixMs: 2,
          completedAtUnixMs: 2,
          errorMessage: null,
          steps: [
            {
              stepId: 'inspect-workbook',
              label: 'Inspect workbook structure',
              status: 'completed',
              summary: 'Read durable workbook structure across 2 sheets.',
              updatedAtUnixMs: 1,
            },
          ],
          artifact: {
            kind: 'markdown',
            title: 'Workbook Summary',
            text: '## Workbook Summary',
          },
        },
      ],
    })
    sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1')) {
        return new Response(JSON.stringify(createSnapshot({ threadId: 'thr-1', workflowRuns: [] })), {
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
      root.render(<AgentHarness zero={zero.zero} zeroEnabled />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()
    expect(host.textContent).toContain('Workflows')
    expect(host.textContent).toContain('Summarize Workbook')
    expect(host.textContent).toContain('Workbook Summary')
    expect(
      fetchSpy.mock.calls.filter(([input, init]) => requestUrl(input).endsWith('/chat/threads') && requestMethod(init) === 'GET'),
    ).toHaveLength(0)

    await act(async () => {
      root.unmount()
    })
  })

  it('hides applied preview system timeline entries from the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads')) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        if (url.endsWith('/chat/threads/thr-1')) {
          return new Response(
            JSON.stringify(
              createSnapshot({
                entries: [
                  {
                    id: 'system-apply:run-1',
                    kind: 'system',
                    turnId: 'turn-1',
                    text: 'Applied workbook change set at revision r7: Write cells in Sheet1!B2',
                    phase: null,
                    toolName: null,
                    toolStatus: null,
                    argumentsText: null,
                    outputText: null,
                    success: null,
                    citations: [
                      {
                        kind: 'range',
                        sheetName: 'Sheet1',
                        startAddress: 'B2',
                        endAddress: 'B2',
                        role: 'target',
                      },
                      {
                        kind: 'revision',
                        revision: 7,
                      },
                    ],
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
        throw new Error(`Unexpected fetch to ${url}`)
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.textContent).not.toContain('Applied workbook change set at revision r7')
    expect(host.textContent).not.toContain('Sheet1!B2')
    expect(host.textContent).not.toContain('r7')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders durable workflow runs in the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(agentStorageKey(), JSON.stringify({ threadId: 'thr-1' }))
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () =>
          new Response(
            JSON.stringify(
              createSnapshot({
                workflowRuns: [
                  {
                    runId: 'wf-1',
                    threadId: 'thr-1',
                    startedByUserId: 'alex@example.com',
                    workflowTemplate: 'summarizeWorkbook',
                    title: 'Summarize Workbook',
                    summary: 'Summarized workbook structure across 2 sheets.',
                    status: 'completed',
                    createdAtUnixMs: 1,
                    updatedAtUnixMs: 2,
                    completedAtUnixMs: 2,
                    errorMessage: null,
                    steps: [
                      {
                        stepId: 'inspect-workbook',
                        label: 'Inspect workbook structure',
                        status: 'completed',
                        summary: 'Read durable workbook structure across 2 sheets.',
                        updatedAtUnixMs: 1,
                      },
                      {
                        stepId: 'draft-summary',
                        label: 'Draft summary artifact',
                        status: 'completed',
                        summary: 'Prepared the durable workbook summary artifact for the thread.',
                        updatedAtUnixMs: 2,
                      },
                    ],
                    artifact: {
                      kind: 'markdown',
                      title: 'Workbook Summary',
                      text: '## Workbook Summary\n\nSheets: 2\n### Sheets\n- Sheet1',
                    },
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

    expect(host.textContent).toContain('Workflows')
    expect(host.textContent).toContain('Summarize Workbook')
    expect(host.textContent).toContain('Inspect workbook structure')
    expect(host.textContent).toContain('Workbook Summary')
    expect(host.textContent).toContain('Sheets: 2')
    expect(host.textContent).toContain('Done')

    await act(async () => {
      root.unmount()
    })
  })

  it('loads durable thread summaries into the assistant panel', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads')) {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-shared',
              scope: 'shared',
              entryCount: 4,
              reviewQueueItemCount: 1,
              latestEntryText: 'Applied workbook change set at revision r7',
            }),
            createThreadSummary({
              threadId: 'thr-private',
              scope: 'private',
              entryCount: 2,
              latestEntryText: 'Review item queued',
            }),
          ]),
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

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-shared']")).not.toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-private']")).not.toBeNull()
    expect(host.textContent).toContain('Shared')
    expect(host.textContent).toContain('Review')
    expect(host.textContent).toContain('4 items')
    expect(host.textContent).toContain('Applied workbook change set at revision r7')

    await act(async () => {
      root.unmount()
    })
  })

  it('switches to a durable thread from the summary strip', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-2',
              scope: 'shared',
              entryCount: 3,
            }),
          ]),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-2')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-2',
              scope: 'shared',
              entries: [],
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

    const threadButton = host.querySelector("[data-testid='workbook-agent-thread-thr-2']")
    expect(threadButton instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(threadButton instanceof HTMLButtonElement)) {
        throw new Error('Thread button not found')
      }
      threadButton.click()
    })

    expect(MockEventSource.latest?.url).toBe('/v2/documents/doc-1/chat/threads/thr-2/events')
    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-2']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()
    expect(fetchSpy).toHaveBeenCalledWith('/v2/documents/doc-1/chat/threads/thr-2')

    await act(async () => {
      root.unmount()
    })
  })

  it('hides the summary strip when it would only repeat the active thread', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return new Response(
          JSON.stringify([
            createThreadSummary({
              threadId: 'thr-1',
              scope: 'private',
              entryCount: 64,
              latestEntryText: 'Done - operating plan now exists as a sheet.',
            }),
          ]),
          {
            status: 200,
            headers: { 'content-type': 'application/json' },
          },
        )
      }
      if (url.endsWith('/chat/threads/thr-1')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              threadId: 'thr-1',
              scope: 'private',
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

    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-thread-thr-1']")).toBeNull()
    expect(host.textContent).not.toContain('64 items')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not render thread scope controls', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        throw new Error(`Unexpected fetch to ${url}`)
      }),
    )

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    expect(host.querySelector("[data-testid='workbook-agent-scope-private']")).toBeNull()
    expect(host.querySelector("[data-testid='workbook-agent-scope-shared']")).toBeNull()

    await act(async () => {
      root.unmount()
    })
  })

  it('restores a new-thread draft after remount', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.stubGlobal(
      'fetch',
      vi.fn(async (input: RequestInfo | URL) => {
        const url = requestUrl(input)
        if (url.endsWith('/chat/threads')) {
          return new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          })
        }
        throw new Error(`Unexpected fetch to ${url}`)
      }),
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
      Reflect.apply(valueSetter, input, ['Persisted draft'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })

    await act(async () => {
      root.unmount()
    })

    const remountRoot = createRoot(host)
    await act(async () => {
      remountRoot.render(<AgentHarness />)
    })

    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe('Persisted draft')

    await act(async () => {
      remountRoot.unmount()
    })
  })

  it('submits the draft on Enter from the chat composer', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/turns')) {
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
                  text: 'Summarize this sheet',
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
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

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
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(turnCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/turns')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')
    const nextInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(nextInput instanceof HTMLTextAreaElement ? nextInput.value : null).toBe('')

    await act(async () => {
      root.unmount()
    })
  })

  it('coalesces rapid first prompt submits into one assistant session and one turn', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    let resolveSession: (() => void) | null = null
    const fetchSpy = vi.fn((input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'GET') {
        return Promise.resolve(
          new Response(JSON.stringify([]), {
            status: 200,
            headers: { 'content-type': 'application/json' },
          }),
        )
      }
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Promise<Response>((resolve) => {
          resolveSession = () => {
            resolve(
              new Response(JSON.stringify(createSnapshot({ entries: [] })), {
                status: 200,
                headers: { 'content-type': 'application/json' },
              }),
            )
          }
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return Promise.resolve(
          new Response(
            JSON.stringify(
              createSnapshot({
                status: 'inProgress',
                activeTurnId: 'turn-1',
                entries: [
                  {
                    id: 'user:turn-1',
                    kind: 'user',
                    turnId: 'turn-1',
                    text: 'Summarize this sheet',
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
          ),
        )
      }
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const callsTo = (suffix: string, method: string) =>
      fetchSpy.mock.calls.filter(([requestInput, init]) => requestUrl(requestInput).endsWith(suffix) && requestMethod(init) === method)

    try {
      await act(async () => {
        root.render(<AgentHarness />)
      })

      const input = host.querySelector("[data-testid='workbook-agent-input']")
      const submit = host.querySelector("[data-testid='workbook-agent-send']")
      expect(input instanceof HTMLTextAreaElement).toBe(true)
      expect(submit instanceof HTMLButtonElement).toBe(true)

      await act(async () => {
        if (!(input instanceof HTMLTextAreaElement) || !(submit instanceof HTMLButtonElement)) {
          throw new Error('Assistant composer not found')
        }
        const valueDescriptor = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')
        const valueSetter = valueDescriptor ? Reflect.get(valueDescriptor, 'set') : null
        if (typeof valueSetter !== 'function') {
          throw new Error('Textarea value setter not found')
        }
        Reflect.apply(valueSetter, input, ['Summarize this sheet'])
        input.dispatchEvent(new Event('input', { bubbles: true }))
        submit.click()
        submit.click()
      })

      expect(callsTo('/chat/threads', 'POST')).toHaveLength(1)
      expect(callsTo('/chat/threads/thr-1/turns', 'POST')).toHaveLength(0)

      await act(async () => {
        resolveSession?.()
        await Promise.resolve()
        await Promise.resolve()
      })

      expect(callsTo('/chat/threads', 'POST')).toHaveLength(1)
      expect(callsTo('/chat/threads/thr-1/turns', 'POST')).toHaveLength(1)
    } finally {
      await act(async () => {
        root.unmount()
      })
    }
  })

  it('submits follow-up prompts through the durable thread route when a thread is already active', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads') && requestMethod(init) === 'POST') {
        return new Response(JSON.stringify(createSnapshot({ entries: [] })), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-2',
              entries: [
                {
                  id: 'optimistic-user:turn-2',
                  kind: 'user',
                  turnId: 'turn-2',
                  text: 'Continue working',
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
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

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
      Reflect.apply(valueSetter, input, ['Continue working'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(turnCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/turns')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')

    await act(async () => {
      root.unmount()
    })
  })

  it('restores the draft and shows the server message when a turn request fails', async () => {
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
        return new Response(JSON.stringify([createThreadSummary()]), {
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
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return new Response(
          JSON.stringify({
            message: 'Prompt rejected by server',
          }),
          {
            status: 422,
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
      Reflect.apply(valueSetter, input, ['Continue working'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
    })
    await flushToasts()

    const turnCall = fetchSpy.mock.calls.find(([requestInput]) => requestUrl(requestInput).endsWith('/chat/threads/thr-1/turns'))
    expect(requestBody(turnCall?.[1])).toEqual({
      prompt: 'Continue working',
      context: createDefaultWorkflowContext(),
    })
    expect(host.textContent).toContain('Prompt rejected by server')
    expect(host.textContent).not.toContain('Workbook agent request failed with status 422')
    const restoredInput = host.querySelector("[data-testid='workbook-agent-input']")
    expect(restoredInput instanceof HTMLTextAreaElement ? restoredInput.value : null).toBe('Continue working')

    await act(async () => {
      root.unmount()
    })
  })

  it('does not inject a synthetic progress row before the turn request resolves', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    window.sessionStorage.setItem(
      agentStorageKey(),
      JSON.stringify({
        threadId: 'thr-1',
      }),
    )
    let resolveTurnResponse: ((response: Response) => void) | null = null
    const turnResponse = new Promise<Response>((resolve) => {
      resolveTurnResponse = resolve
    })
    const fetchSpy = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = requestUrl(input)
      if (url.endsWith('/chat/threads/thr-1') && requestMethod(init) === 'GET') {
        return new Response(JSON.stringify(createSnapshot()), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      }
      if (url.endsWith('/chat/threads/thr-1/turns')) {
        return await turnResponse
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
      Reflect.apply(valueSetter, input, ['yo'])
      input.dispatchEvent(new Event('input', { bubbles: true }))
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          bubbles: true,
          key: 'Enter',
        }),
      )
      await Promise.resolve()
    })

    expect(host.textContent).toContain('yo')
    expect(host.textContent).not.toContain('Reviewing workbook context and drafting a response.')
    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).toBeNull()

    await act(async () => {
      resolveTurnResponse?.(
        new Response(
          JSON.stringify(
            createSnapshot({
              status: 'inProgress',
              activeTurnId: 'turn-3',
              entries: [
                {
                  id: 'optimistic-user:turn-3',
                  kind: 'user',
                  turnId: 'turn-3',
                  text: 'yo',
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
        ),
      )
      await Promise.resolve()
    })

    expect(host.querySelector("[data-testid='workbook-agent-progress-row']")).not.toBeNull()
    expect(host.textContent).toContain('Thinking')

    await act(async () => {
      root.unmount()
    })
  })

  it('uses the composer button to interrupt an active turn', async () => {
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
              status: 'inProgress',
              activeTurnId: 'turn-1',
              entries: [
                {
                  id: 'assistant-1',
                  kind: 'assistant',
                  turnId: 'turn-1',
                  text: 'Working',
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
      if (url.endsWith('/interrupt')) {
        return new Response(
          JSON.stringify(
            createSnapshot({
              status: 'idle',
              activeTurnId: null,
              entries: [
                {
                  id: 'assistant-1',
                  kind: 'assistant',
                  turnId: 'turn-1',
                  text: 'Working',
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
      throw new Error(`Unexpected fetch to ${url}`)
    })
    vi.stubGlobal('fetch', fetchSpy)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(<AgentHarness />)
    })

    const button = host.querySelector("[data-testid='workbook-agent-send']")
    expect(button instanceof HTMLButtonElement).toBe(true)
    expect(button instanceof HTMLButtonElement ? button.getAttribute('aria-label') : null).toBe('Stop')

    await act(async () => {
      if (!(button instanceof HTMLButtonElement)) {
        throw new Error('Agent button not found')
      }
      button.click()
    })

    const interruptCall = fetchSpy.mock.calls.find(([input]) => requestUrl(input).endsWith('/chat/threads/thr-1/interrupt'))
    expect(interruptCall?.[0]).toBe('/v2/documents/doc-1/chat/threads/thr-1/interrupt')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders structured workbook comprehension tool results in the rail', async () => {
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
                    id: 'tool-search',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'search_workbook',
                    toolStatus: 'completed',
                    argumentsText: '{"query":"gross margin"}',
                    outputText: JSON.stringify({
                      query: 'gross margin',
                      summary: { matchCount: 1, truncated: false },
                      matches: [
                        {
                          kind: 'cell',
                          sheetName: 'Sheet1',
                          address: 'A2',
                          snippet: 'Gross Margin',
                          reasons: ['value'],
                          score: 65,
                        },
                      ],
                    }),
                    success: true,
                  },
                  {
                    id: 'tool-issues',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'find_formula_issues',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      summary: {
                        issueCount: 1,
                        scannedFormulaCells: 3,
                        errorCount: 1,
                        cycleCount: 0,
                        unsupportedCount: 0,
                      },
                      issues: [
                        {
                          sheetName: 'Sheet1',
                          address: 'C1',
                          formula: '=1/0',
                          valueText: '#DIV/0!',
                          issueKinds: ['error'],
                        },
                      ],
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

    expect(host.querySelector("[data-testid='workbook-agent-panel-scroll-viewport']")).not.toBeNull()
    expect(host.textContent).toContain('Search Workbook')
    expect(host.textContent).toContain('Find Formula Issues')
    expect(host.textContent).not.toContain('Gross Margin')
    expect(host.textContent).not.toContain('gross margin')
    expect(host.textContent).not.toContain('C1')

    const searchToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-search']")
    const issuesToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-issues']")
    expect(searchToggle instanceof HTMLButtonElement).toBe(true)
    expect(issuesToggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(searchToggle instanceof HTMLButtonElement) || !(issuesToggle instanceof HTMLButtonElement)) {
        throw new Error('Tool toggles not found')
      }
      searchToggle.click()
      issuesToggle.click()
    })

    expect(host.textContent).toContain('Gross Margin')
    expect(host.textContent).toContain('gross margin')
    expect(host.textContent).toContain('C1')

    await act(async () => {
      root.unmount()
    })
  })

  it('renders workbook inspection tool payloads as structured result cards instead of raw JSON blobs', async () => {
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
                    id: 'tool-tables',
                    kind: 'tool',
                    turnId: 'turn-1',
                    text: null,
                    phase: null,
                    toolName: 'list_tables',
                    toolStatus: 'completed',
                    argumentsText: '{}',
                    outputText: JSON.stringify({
                      documentId: 'bilig-demo',
                      tableCount: 1,
                      tables: [
                        {
                          name: 'OperatingPlan',
                          sheetName: 'sheet3',
                          startAddress: 'A6',
                          endAddress: 'K10',
                          headerRowCount: 1,
                          rowCount: 4,
                          columnCount: 11,
                          columnNames: ['Item', 'Vendor', 'Category'],
                        },
                      ],
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

    expect(host.textContent).toContain('List Tables')
    expect(host.textContent).not.toContain('"documentId": "bilig-demo"')
    expect(host.textContent).not.toContain('"tableCount": 1')

    const readToggle = host.querySelector("[data-testid='workbook-agent-tool-toggle-tool-tables']")
    expect(readToggle instanceof HTMLButtonElement).toBe(true)

    await act(async () => {
      if (!(readToggle instanceof HTMLButtonElement)) {
        throw new Error('List tables tool toggle not found')
      }
      readToggle.click()
    })

    const readPanelViewport = host.querySelector("[data-testid='workbook-agent-tool-panel-tool-tables-viewport']")
    expect(readPanelViewport instanceof HTMLDivElement).toBe(true)
    expect(readPanelViewport?.className).toContain('h-44')
    expect(host.textContent).toContain('1 table')
    expect(host.textContent).toContain('OperatingPlan')
    expect(host.textContent).toContain('sheet3!A6:K10')
    expect(host.textContent).toContain('4 rows')
    expect(host.textContent).toContain('11 columns')
    expect(host.textContent).not.toContain('"tables": [')

    await act(async () => {
      root.unmount()
    })
  })
})
