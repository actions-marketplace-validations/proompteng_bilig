import type { Item } from '../gridTypes.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import { DirtyMaskV3, type DirtyTileLocalSpanV3 } from '../renderer-v3/tile-damage-index.js'

const TEXT_ONLY_DIRTY_MASK_V3 = DirtyMaskV3.Value | DirtyMaskV3.Text

export function appendLocalTextDirtySpan(
  spansByTile: Map<number, DirtyTileLocalSpanV3[]>,
  tileKey: number,
  span: DirtyTileLocalSpanV3 | null,
): void {
  if (!span) {
    return
  }
  const spans = spansByTile.get(tileKey)
  if (spans) {
    spans.push(span)
    return
  }
  spansByTile.set(tileKey, [span])
}

export function appendLocalDirtySpans(
  spansByTile: Map<number, DirtyTileLocalSpanV3[]>,
  tileKey: number,
  spans: readonly DirtyTileLocalSpanV3[],
): void {
  if (spans.length === 0) {
    return
  }
  const existing = spansByTile.get(tileKey)
  if (existing) {
    existing.push(...spans)
    return
  }
  spansByTile.set(tileKey, [...spans])
}

export function mergeLocalDirtySpans(
  authoritativeSpans: readonly DirtyTileLocalSpanV3[],
  localTextSpans: readonly DirtyTileLocalSpanV3[] | undefined,
): readonly DirtyTileLocalSpanV3[] {
  if (!localTextSpans || localTextSpans.length === 0) {
    return authoritativeSpans
  }
  if (authoritativeSpans.length === 0) {
    return localTextSpans
  }
  return [...authoritativeSpans, ...localTextSpans]
}

export function dirtySpansFromTile(tile: GridRenderTile): readonly DirtyTileLocalSpanV3[] {
  const dirtyMasks = tile.dirtyMasks
  const dirtyLocalRows = tile.dirtyLocalRows
  const dirtyLocalCols = tile.dirtyLocalCols
  if (!dirtyMasks || !dirtyLocalRows || !dirtyLocalCols || dirtyMasks.length === 0) {
    return []
  }
  if (dirtyLocalRows.length !== dirtyMasks.length * 2 || dirtyLocalCols.length !== dirtyMasks.length * 2) {
    return []
  }
  return Array.from(dirtyMasks, (mask, index) => {
    const offset = index * 2
    return {
      colEnd: dirtyLocalCols[offset + 1] ?? dirtyLocalCols[offset] ?? 0,
      colStart: dirtyLocalCols[offset] ?? 0,
      mask,
      rowEnd: dirtyLocalRows[offset + 1] ?? dirtyLocalRows[offset] ?? 0,
      rowStart: dirtyLocalRows[offset] ?? 0,
    }
  })
}

export function textDirtySpanForCell(tile: GridRenderTile, cell: Item | null | undefined): DirtyTileLocalSpanV3 | null {
  if (!cell) {
    return null
  }
  const [col, row] = cell
  if (row < tile.bounds.rowStart || row > tile.bounds.rowEnd || col < tile.bounds.colStart || col > tile.bounds.colEnd) {
    return null
  }
  return {
    colEnd: col - tile.bounds.colStart,
    colStart: col - tile.bounds.colStart,
    mask: TEXT_ONLY_DIRTY_MASK_V3,
    rowEnd: row - tile.bounds.rowStart,
    rowStart: row - tile.bounds.rowStart,
  }
}

export function fullTextDirtySpanForTile(tile: GridRenderTile): DirtyTileLocalSpanV3 {
  return {
    colEnd: tile.bounds.colEnd - tile.bounds.colStart,
    colStart: 0,
    mask: TEXT_ONLY_DIRTY_MASK_V3,
    rowEnd: tile.bounds.rowEnd - tile.bounds.rowStart,
    rowStart: 0,
  }
}
