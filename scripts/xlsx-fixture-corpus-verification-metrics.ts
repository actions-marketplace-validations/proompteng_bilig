import type { XlsxFixtureCorpusWorkerOptions } from './xlsx-fixture-corpus-footprint.ts'
import type {
  XlsxFixtureCorpusCase,
  XlsxFixtureVerificationPhase,
  XlsxFixtureVerificationPhaseTiming,
} from './xlsx-fixture-corpus-types.ts'

interface VerificationRuntimeMetrics {
  readonly startedAt: number
  readonly phaseTimings: XlsxFixtureVerificationPhaseTiming[]
}

export function startVerificationRuntimeMetrics(): VerificationRuntimeMetrics {
  return {
    startedAt: performance.now(),
    phaseTimings: [],
  }
}

export async function timeVerificationPhase<T>(
  metrics: VerificationRuntimeMetrics,
  workerOptions: XlsxFixtureCorpusWorkerOptions,
  phase: XlsxFixtureVerificationPhase,
  fn: () => T | Promise<T>,
): Promise<T> {
  workerOptions.onPhase?.(phase)
  const startedAt = performance.now()
  try {
    return await fn()
  } finally {
    metrics.phaseTimings.push({ phase, elapsedMs: roundElapsedMs(performance.now() - startedAt) })
  }
}

export function withVerificationRuntimeMetrics(
  corpusCase: XlsxFixtureCorpusCase,
  metrics: VerificationRuntimeMetrics,
  peakRssBytes?: number,
): XlsxFixtureCorpusCase {
  return withPeakRssBytes(
    {
      ...corpusCase,
      elapsedMs: roundElapsedMs(performance.now() - metrics.startedAt),
      phaseTimings: metrics.phaseTimings,
    },
    peakRssBytes,
  )
}

export function withPeakRssBytes(corpusCase: XlsxFixtureCorpusCase, peakRssBytes: number): XlsxFixtureCorpusCase
export function withPeakRssBytes(corpusCase: XlsxFixtureCorpusCase, peakRssBytes?: number): XlsxFixtureCorpusCase
export function withPeakRssBytes(corpusCase: XlsxFixtureCorpusCase, peakRssBytes?: number): XlsxFixtureCorpusCase {
  if (peakRssBytes === undefined || peakRssBytes <= 0) {
    return corpusCase
  }
  return {
    ...corpusCase,
    peakRssBytes: Math.max(corpusCase.peakRssBytes ?? 0, Math.trunc(peakRssBytes)),
  }
}

function roundElapsedMs(value: number): number {
  return Math.max(0, Math.round(value))
}
