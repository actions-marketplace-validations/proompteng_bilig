const INTERACTION_PERSISTENCE_DELAY_MS = 350
const INTERACTION_IDLE_TIMEOUT_MS = 1_000

type BrowserWindowWithIdle = Window & {
  requestIdleCallback?: ((callback: () => void, options?: { timeout?: number }) => number) | undefined
}

export function deferNextInteractionFrame(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(() => resolve())
      return
    }
    setTimeout(resolve, 0)
  })
}

export async function deferInteractionPersistence(): Promise<void> {
  if (typeof window === 'undefined') {
    return
  }
  const browserWindow = window as BrowserWindowWithIdle
  if (typeof browserWindow.requestIdleCallback !== 'function') {
    await deferNextInteractionFrame()
    return
  }
  await deferNextInteractionFrame()
  await new Promise<void>((resolve) => {
    window.setTimeout(() => {
      const requestIdle = browserWindow.requestIdleCallback
      if (typeof requestIdle !== 'function') {
        window.setTimeout(resolve, 0)
        return
      }
      requestIdle(() => resolve(), { timeout: INTERACTION_IDLE_TIMEOUT_MS })
    }, INTERACTION_PERSISTENCE_DELAY_MS)
  })
}
