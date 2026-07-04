import { expect, vi } from 'vitest'

export function captureExpectedConsoleDebug(expectedMessage: string) {
  const captured = captureExpectedConsoleDebugMessages()

  return {
    expectLogCount(expectedCount: number): void {
      captured.expectMessageCount(expectedMessage, expectedCount)
    },
    payloads(): readonly unknown[] {
      return captured.payloads(expectedMessage)
    },
  }
}

export function captureExpectedConsoleDebugMessages() {
  const spy = vi.spyOn(console, 'debug').mockImplementation(() => undefined)

  const matchingCalls = (expectedMessage: string) =>
    spy.mock.calls.filter((call) => call[0] === '[bilig-web]' && call[1] === expectedMessage)

  return {
    expectMessageCount(expectedMessage: string, expectedCount: number): void {
      expect(matchingCalls(expectedMessage)).toHaveLength(expectedCount)
    },
    payloads(expectedMessage: string): readonly unknown[] {
      return matchingCalls(expectedMessage).map((call) => call[2])
    },
  }
}
