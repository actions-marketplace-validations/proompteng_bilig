import { describe, expect, it } from 'vitest'

import { DEFAULT_TEST_MAX_LINES, parseTestMaxLines } from '../test-file-size-config.js'

describe('test file size config', () => {
  it('defaults to the repository test size limit', () => {
    expect(DEFAULT_TEST_MAX_LINES).toBe(1500)
    expect(parseTestMaxLines(undefined)).toBe(1500)
  })

  it('accepts explicit positive safe integer limits', () => {
    expect(parseTestMaxLines('1')).toBe(1)
    expect(parseTestMaxLines('2000')).toBe(2000)
  })

  it.each(['', '0', '-1', '1000.5', '1000abc'])('rejects malformed BILIG_TEST_MAX_LINES=%s', (value) => {
    expect(() => parseTestMaxLines(value)).toThrow(`BILIG_TEST_MAX_LINES must be a positive integer, got ${value}`)
  })

  it('rejects unsafe integer limits', () => {
    expect(() => parseTestMaxLines('9007199254740992')).toThrow('BILIG_TEST_MAX_LINES must be a safe integer, got 9007199254740992')
  })
})
