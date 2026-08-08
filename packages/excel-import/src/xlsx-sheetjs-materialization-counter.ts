import type { WorkbookSnapshot } from '@bilig/protocol'

import { assertXlsxMaterializedCountsWithinLimits, resolveXlsxSheetJsFallbackLimits, type XlsxImportOptions } from './xlsx-import-limits.js'

type ImportedSnapshotCell = WorkbookSnapshot['sheets'][number]['cells'][number]

export interface SheetJsMaterializationCounter {
  readonly cellCount: number
  readonly formulaCellCount: number
  recordCell(cell: ImportedSnapshotCell): void
}

export function createSheetJsMaterializationCounter(options: XlsxImportOptions): SheetJsMaterializationCounter {
  const limits = resolveXlsxSheetJsFallbackLimits(options)
  let cellCount = 0
  let formulaCellCount = 0
  return {
    get cellCount() {
      return cellCount
    },
    get formulaCellCount() {
      return formulaCellCount
    },
    recordCell(cell) {
      const nextCounts = {
        cellCount: cellCount + 1,
        formulaCellCount: formulaCellCount + (cell.formula === undefined ? 0 : 1),
      }
      assertXlsxMaterializedCountsWithinLimits(nextCounts, limits)
      cellCount = nextCounts.cellCount
      formulaCellCount = nextCounts.formulaCellCount
    },
  }
}
