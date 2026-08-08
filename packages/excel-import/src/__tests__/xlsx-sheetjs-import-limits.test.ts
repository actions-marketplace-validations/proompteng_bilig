import { describe, expect, it } from 'vitest'

import { LEGACY_XLS_CONTENT_TYPE, type XlsxImportSizeLimitExceededError } from '../index.js'
import { importParsedSheetJsWorkbook } from '../xlsx-sheetjs-import.js'
import type { SheetJsWorkBook } from '../xlsx-sheetjs-types.js'

describe('SheetJS import materialization limits', () => {
  it('rejects workbooks while materializing cells beyond the configured budget', () => {
    expect(() =>
      importParsedSheetJsWorkbook({
        workbook: workbookWithCells(),
        fileName: 'bounded.xls',
        contentType: LEGACY_XLS_CONTENT_TYPE,
        workbookZip: null,
        fallbackArtifactSource: new Uint8Array(),
        sourceFileSizeBytes: 0,
        options: { limits: { maxMaterializedCells: 2 } },
      }),
    ).toThrowError(
      expect.objectContaining<XlsxImportSizeLimitExceededError>({
        reason: 'cell-count',
        observedCount: 3,
      }),
    )
  })

  it('rejects legacy XLS workbooks while materializing formulas beyond the configured budget', () => {
    expect(() =>
      importParsedSheetJsWorkbook({
        workbook: workbookWithCells(),
        fileName: 'bounded.xls',
        contentType: LEGACY_XLS_CONTENT_TYPE,
        workbookZip: null,
        fallbackArtifactSource: new Uint8Array(),
        sourceFileSizeBytes: 0,
        options: { limits: { maxMaterializedFormulaCells: 0 } },
      }),
    ).toThrowError(
      expect.objectContaining<XlsxImportSizeLimitExceededError>({
        reason: 'formula-cell-count',
        observedCount: 1,
      }),
    )
  })
})

function workbookWithCells(): SheetJsWorkBook {
  return {
    SheetNames: ['Sheet1'],
    Sheets: {
      Sheet1: {
        '!ref': 'A1:C1',
        A1: { t: 'n', v: 1 },
        B1: { t: 'n', v: 2 },
        C1: { t: 'n', v: 3, f: 'A1+B1' },
      },
    },
  }
}
