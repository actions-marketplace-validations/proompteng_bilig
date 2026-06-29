import { ValueTag } from '@bilig/protocol'
import type { RuntimeColumnView } from './runtime-column-store-service.js'
import { decodeValueTag } from './criterion-range-values.js'

export function countNonEmptyRowsInView(view: RuntimeColumnView): number {
  let count = 0
  view.owner.pages.forEach((page) => {
    const rowStart = Math.max(view.rowStart, page.rowStart)
    const rowEnd = Math.min(view.rowEnd, page.rowStart + page.tags.length - 1)
    if (rowStart > rowEnd || page.nonEmptyCount === 0) {
      return
    }
    if (rowStart === page.rowStart && rowEnd === page.rowStart + page.tags.length - 1) {
      count += page.nonEmptyCount
      return
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (decodeValueTag(page.tags[row - page.rowStart]) !== ValueTag.Empty) {
        count += 1
      }
    }
  })
  return count
}

export function forEachNonEmptyRowOffsetInView(view: RuntimeColumnView, fn: (rowOffset: number) => void): void {
  view.owner.pages.forEach((page) => {
    const rowStart = Math.max(view.rowStart, page.rowStart)
    const rowEnd = Math.min(view.rowEnd, page.rowStart + page.tags.length - 1)
    if (rowStart > rowEnd || page.nonEmptyCount === 0) {
      return
    }
    if (rowStart === page.rowStart && rowEnd === page.rowStart + page.tags.length - 1 && page.nonEmptyCount === page.tags.length) {
      for (let row = rowStart; row <= rowEnd; row += 1) {
        fn(row - view.rowStart)
      }
      return
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      if (decodeValueTag(page.tags[row - page.rowStart]) !== ValueTag.Empty) {
        fn(row - view.rowStart)
      }
    }
  })
}
