import type { Unzipped } from 'fflate'
import { describe, expect, it } from 'vitest'

import {
  assertXlsxZipWithinMaterializationLimits,
  planXlsxImportRoute,
  resolveXlsxImportLimits,
  type XlsxImportSizeLimitExceededError,
} from '../xlsx-import-limits.js'

describe('xlsx import route planning', () => {
  it('keeps small default imports on the full-fidelity fallback threshold', () => {
    const route = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 8_000,
      options: {},
      inspection: null,
    })

    expect(route.createLargeSimpleImportOptions().minByteLength).toBeUndefined()
  })

  it('allows formula recalc callers to opt into small native simple imports', () => {
    const route = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 8_000,
      options: { preferNativeSimpleImport: true },
      inspection: null,
    })

    expect(route.createLargeSimpleImportOptions().minByteLength).toBe(0)
  })

  it('keeps limits false bounded unless large legacy fallback is explicit', () => {
    const defaultRoute = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 1_100_000,
      options: { limits: false },
      inspection: null,
    })
    const legacyRoute = planXlsxImportRoute({
      workbookZip: zipWith(['xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet1.xml']),
      sourceByteLength: 1_100_000,
      options: { allowLegacyLargeSheetJsFallback: true, limits: false },
      inspection: null,
    })

    expect(defaultRoute.shouldInspectBeforeSheetJsFallback).toBe(true)
    expect(defaultRoute.inspectionOptions?.minByteLength).toBe(0)
    expect(legacyRoute.shouldInspectBeforeSheetJsFallback).toBe(false)
    expect(legacyRoute.inspectionOptions).toBeUndefined()
  })

  it('rejects zip containers whose cumulative expanded entries exceed the budget', () => {
    const zip = zipWithSizes({
      'xl/workbook.xml': 60,
      'xl/worksheets/sheet1.xml': 50,
    })
    const limits = resolveXlsxImportLimits({ limits: { maxUncompressedBytes: 100 } })

    expect(() => assertXlsxZipWithinMaterializationLimits(zip, limits)).toThrowError(
      expect.objectContaining<XlsxImportSizeLimitExceededError>({
        reason: 'uncompressed-byte-count',
        uncompressedByteLength: 110,
      }),
    )
  })

  it.each([NaN, -1, 1.5])('rejects invalid programmatic import limits (%s)', (value) => {
    expect(() => resolveXlsxImportLimits({ limits: { maxMaterializedCells: value } })).toThrowError(
      'maxMaterializedCells must be a non-negative safe integer or Infinity',
    )
  })
})

function zipWith(paths: readonly string[]): Unzipped {
  return Object.fromEntries(paths.map((path) => [path, new Uint8Array()]))
}

function zipWithSizes(entries: Readonly<Record<string, number>>): Unzipped {
  return Object.fromEntries(Object.entries(entries).map(([path, byteLength]) => [path, new Uint8Array(byteLength)]))
}
