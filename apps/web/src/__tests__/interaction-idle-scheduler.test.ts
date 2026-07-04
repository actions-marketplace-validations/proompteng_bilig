// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { deferInteractionPersistence } from '../interaction-idle-scheduler.js'

describe('interaction idle scheduler', () => {
  const originalRequestAnimationFrame = window.requestAnimationFrame
  const originalRequestIdleCallback = window.requestIdleCallback

  afterEach(() => {
    vi.useRealTimers()
    Object.defineProperty(window, 'requestAnimationFrame', {
      configurable: true,
      value: originalRequestAnimationFrame,
      writable: true,
    })
    Object.defineProperty(window, 'requestIdleCallback', {
      configurable: true,
      value: originalRequestIdleCallback,
      writable: true,
    })
  })

  it('falls back to a timer when requestIdleCallback is removed before the idle window opens', async () => {
    vi.useFakeTimers()
    let frameCallback: FrameRequestCallback | null = null
    Object.defineProperty(window, 'requestAnimationFrame', {
      configurable: true,
      value: (callback: FrameRequestCallback) => {
        frameCallback = callback
        return 1
      },
      writable: true,
    })
    Object.defineProperty(window, 'requestIdleCallback', {
      configurable: true,
      value: vi.fn(),
      writable: true,
    })

    const scheduled = deferInteractionPersistence()
    if (!frameCallback) {
      throw new Error('requestAnimationFrame was not scheduled')
    }
    frameCallback(16)
    await Promise.resolve()
    Object.defineProperty(window, 'requestIdleCallback', {
      configurable: true,
      value: undefined,
      writable: true,
    })

    await vi.advanceTimersByTimeAsync(350)
    await vi.runOnlyPendingTimersAsync()
    await scheduled

    expect(window.requestIdleCallback).toBeUndefined()
  })
})
