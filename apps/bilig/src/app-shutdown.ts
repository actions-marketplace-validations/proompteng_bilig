export type ApplicationShutdownReason = 'SIGINT' | 'SIGTERM' | 'startup-error'

interface ApplicationShutdownOptions {
  readonly closeHttpServer: () => Promise<void>
  readonly closeWorkbookAgent: () => Promise<void>
  readonly closePersistence: () => Promise<void>
}

interface ShutdownSignalSource {
  once(signal: 'SIGINT' | 'SIGTERM', listener: () => void): unknown
  removeListener(signal: 'SIGINT' | 'SIGTERM', listener: () => void): unknown
}

export function createApplicationShutdown(options: ApplicationShutdownOptions) {
  let shutdownPromise: Promise<void> | null = null
  return (_reason: ApplicationShutdownReason): Promise<void> => {
    shutdownPromise ??= closeApplication(options)
    return shutdownPromise
  }
}

export function registerApplicationShutdownSignals(input: {
  readonly shutdown: (reason: ApplicationShutdownReason) => Promise<void>
  readonly onError: (error: unknown) => void
  readonly signalSource?: ShutdownSignalSource
}): () => void {
  const signalSource = input.signalSource ?? process
  let removed = false
  const remove = () => {
    if (removed) {
      return
    }
    removed = true
    signalSource.removeListener('SIGINT', onSigint)
    signalSource.removeListener('SIGTERM', onSigterm)
  }
  const handle = (reason: 'SIGINT' | 'SIGTERM') => {
    remove()
    const shutdown = async () => {
      try {
        await input.shutdown(reason)
      } catch (error) {
        input.onError(error)
      }
    }
    void shutdown()
  }
  const onSigint = () => handle('SIGINT')
  const onSigterm = () => handle('SIGTERM')
  signalSource.once('SIGINT', onSigint)
  signalSource.once('SIGTERM', onSigterm)
  return remove
}

async function closeApplication(options: ApplicationShutdownOptions): Promise<void> {
  const failures: unknown[] = []
  const boundaryResults = await Promise.allSettled([options.closeHttpServer(), options.closeWorkbookAgent()])
  for (const result of boundaryResults) {
    if (result.status === 'rejected' && !failures.includes(result.reason)) {
      failures.push(result.reason)
    }
  }
  try {
    await options.closePersistence()
  } catch (error) {
    failures.push(error)
  }
  if (failures.length > 0) {
    throw new AggregateError(failures, 'Failed to shut down the Bilig application cleanly')
  }
}
