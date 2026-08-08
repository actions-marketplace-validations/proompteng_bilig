import { EventEmitter } from 'node:events'
import { describe, expect, it, vi } from 'vitest'
import { createApplicationShutdown, registerApplicationShutdownSignals } from '../app-shutdown.js'

describe('application shutdown', () => {
  it('starts HTTP and workflow draining together, then closes persistence exactly once', async () => {
    const order: string[] = []
    const shutdown = createApplicationShutdown({
      closeHttpServer: async () => {
        order.push('http')
      },
      closeWorkbookAgent: async () => {
        order.push('workflows')
      },
      closePersistence: async () => {
        order.push('persistence')
      },
    })

    await Promise.all([shutdown('SIGTERM'), shutdown('SIGINT')])

    expect(order.slice(0, 2).toSorted()).toEqual(['http', 'workflows'])
    expect(order).toHaveLength(3)
    expect(order[2]).toBe('persistence')
  })

  it('still closes persistence and reports all boundary failures', async () => {
    const closePersistence = vi.fn(async () => {})
    const shutdown = createApplicationShutdown({
      closeHttpServer: async () => {
        throw new Error('http failed')
      },
      closeWorkbookAgent: async () => {
        throw new Error('workflow failed')
      },
      closePersistence,
    })

    await expect(shutdown('SIGTERM')).rejects.toMatchObject({ errors: expect.arrayContaining([expect.any(Error), expect.any(Error)]) })
    expect(closePersistence).toHaveBeenCalledOnce()
  })

  it('registers one-shot signal handlers and removes the unused handler', async () => {
    const signalSource = new EventEmitter()
    const shutdown = vi.fn(async () => {})
    const onError = vi.fn()
    registerApplicationShutdownSignals({ shutdown, onError, signalSource })

    signalSource.emit('SIGTERM')
    await Promise.resolve()

    expect(shutdown).toHaveBeenCalledWith('SIGTERM')
    expect(signalSource.listenerCount('SIGINT')).toBe(0)
    expect(signalSource.listenerCount('SIGTERM')).toBe(0)
    expect(onError).not.toHaveBeenCalled()
  })
})
