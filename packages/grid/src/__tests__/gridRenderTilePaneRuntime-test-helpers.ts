import { ValueTag, type CellSnapshot } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import { getGridMetrics } from '../gridMetrics.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import type {
  GridRenderTile,
  GridRenderTileDeltaSubscription,
  GridRenderTileSceneChange,
  GridRenderTileSource,
} from '../renderer-v3/render-tile-source.js'
import { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
import { DirtyMaskV3, type WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import type { GridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
import { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'
export { formatAddress } from '@bilig/formula'
export { ValueTag, type CellSnapshot, type CellStyleRecord } from '@bilig/protocol'
export { describe, expect, it } from 'vitest'
export type { GridEngineLike } from '../grid-engine.js'
export { getGridMetrics } from '../gridMetrics.js'
export { materializeGridRenderTileV3 } from '../renderer-v3/grid-tile-materializer.js'
export { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
export type {
  GridRenderTile,
  GridRenderTileDeltaSubscription,
  GridRenderTileSceneChange,
  GridRenderTileSource,
} from '../renderer-v3/render-tile-source.js'
export { GRID_TEXT_METRIC_FLOAT_COUNT_V3 } from '../renderer-v3/text-run-buffer.js'
export { DirtyMaskV3, type WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
export { packTileKey53 } from '../renderer-v3/tile-key.js'
export { getGridRenderTilePaneRuntime, GridRenderTilePaneRuntime } from '../runtime/gridRenderTilePaneRuntime.js'
export { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
export { WORKBOOK_DEFAULT_FONT_SIZE, WORKBOOK_FONT_SANS, workbookFontPointSizeToCssPx } from '../workbookTheme.js'

export const DEFAULT_TEST_TEXT_COLOR = '#1f2933'
export const DEFAULT_TEST_FONT_SIZE = workbookFontPointSizeToCssPx(WORKBOOK_DEFAULT_FONT_SIZE)
export const DEFAULT_TEST_FONT = `400 ${DEFAULT_TEST_FONT_SIZE}px ${WORKBOOK_FONT_SANS}`

export const TEST_ENGINE: GridEngineLike = {
  getCell: (_sheetName, address) => createEmptyCellSnapshot(address),
  getCellStyle: () => undefined,
  subscribeCells: () => () => {},
  workbook: {
    getSheet: () => undefined,
  },
}

export function createEmptyCellSnapshot(address: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: '',
    sheetName: 'Sheet1',
    value: { tag: ValueTag.Empty },
    version: 0,
  }
}

export function createStringCellSnapshot(address: string, value: string): CellSnapshot {
  return {
    address,
    flags: 0,
    input: value,
    sheetName: 'Sheet1',
    value: { tag: ValueTag.String, value, stringId: 0 },
    version: 1,
  }
}

export function createStyledStringCellSnapshot(address: string, value: string, styleId: string): CellSnapshot {
  return {
    ...createStringCellSnapshot(address, value),
    styleId,
  }
}

export const LOCAL_EMPTY_ENGINE: GridEngineLike = {
  getCell: (_sheetName, address) => createEmptyCellSnapshot(address),
  getCellStyle: () => undefined,
  subscribeCells: () => () => {},
  workbook: {
    getSheet: () => undefined,
  },
}

export function expectedGridBorderRectCount(bounds: GridRenderTile['bounds']): number {
  return bounds.rowEnd - bounds.rowStart + 1 + bounds.colEnd - bounds.colStart + 1
}

export function createGridBorderRectInstances(rectCount: number): Float32Array {
  const rectInstances = new Float32Array(rectCount * GRID_RECT_INSTANCE_FLOAT_COUNT_V3)
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    rectInstances[offset + 2] = index % 2 === 0 ? 100 : 1
    rectInstances[offset + 3] = index % 2 === 0 ? 1 : 20
    rectInstances[offset + 11] = 1
    rectInstances[offset + 13] = 1
  }
  return rectInstances
}

export function hasOpaqueGreenFillRect(tile: GridRenderTile | undefined): boolean {
  if (!tile) {
    return false
  }
  for (let index = 0; index < tile.rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const red = tile.rectInstances[offset + 4] ?? 1
    const green = tile.rectInstances[offset + 5] ?? 0
    const blue = tile.rectInstances[offset + 6] ?? 1
    const alpha = tile.rectInstances[offset + 7] ?? 0
    const instanceKind = tile.rectInstances[offset + 13] ?? -1
    if (instanceKind === 0 && red < 0.05 && green > 0.95 && blue < 0.05 && alpha > 0.95) {
      return true
    }
  }
  return false
}

export function createHost(): GridRuntimeHost {
  return new GridRuntimeHost({
    columnCount: 1000,
    defaultColumnWidth: 100,
    defaultRowHeight: 20,
    gridMetrics: getGridMetrics(),
    rowCount: 1000,
    viewportHeight: 400,
    viewportWidth: 800,
  })
}

export function createRenderTile(tileId: number, sheetId = 7, sheetOrdinal = sheetId): GridRenderTile {
  const bounds = { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 }
  const rectCount = expectedGridBorderRectCount(bounds)
  return {
    bounds,
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId,
      sheetOrdinal,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount,
    rectInstances: createGridBorderRectInstances(rectCount),
    textCount: 0,
    textMetrics: new Float32Array(GRID_TEXT_METRIC_FLOAT_COUNT_V3),
    textRuns: [],
    tileId,
    version: {
      axisX: 1,
      axisY: 1,
      freeze: 0,
      styles: 1,
      text: 1,
      values: 1,
    },
  }
}

export function createRenderTileSource(tiles: readonly GridRenderTile[]): GridRenderTileSource {
  const byId = new Map(tiles.map((tile) => [tile.tileId, tile]))
  return {
    peekRenderTile: (tileId) => byId.get(tileId) ?? null,
    subscribeRenderTileDeltas: () => () => {},
  }
}

export function createCapturingRenderTileSource(): {
  readonly source: GridRenderTileSource
  readonly captured: () => GridRenderTileDeltaSubscription | null
  readonly subscribeCount: () => number
  readonly unsubscribed: () => boolean
  readonly unsubscribeCount: () => number
} {
  let captured: GridRenderTileDeltaSubscription | null = null
  let subscribeCount = 0
  let unsubscribeCount = 0
  return {
    captured: () => captured,
    source: {
      peekRenderTile: () => null,
      subscribeRenderTileDeltas: (subscription) => {
        captured = subscription
        subscribeCount += 1
        return () => {
          unsubscribeCount += 1
        }
      },
    },
    subscribeCount: () => subscribeCount,
    unsubscribed: () => unsubscribeCount > 0,
    unsubscribeCount: () => unsubscribeCount,
  }
}

export function createMutableRenderTileSource(tiles: readonly GridRenderTile[] = []): {
  readonly source: GridRenderTileSource
  readonly deleteTile: (tileId: number) => void
  readonly emit: (change: {
    readonly batchId?: number | undefined
    readonly cameraSeq?: number | undefined
    readonly changedTileIds?: readonly number[] | undefined
    readonly invalidatedTileIds?: readonly number[] | undefined
  }) => void
  readonly setTile: (tile: GridRenderTile) => void
  readonly unsubscribed: () => boolean
} {
  const byId = new Map(tiles.map((tile) => [tile.tileId, tile]))
  let listener: ((change: GridRenderTileSceneChange) => void) | null = null
  let unsubscribed = false
  return {
    deleteTile: (tileId) => {
      byId.delete(tileId)
    },
    emit: (change) =>
      listener?.({
        batchId: change.batchId ?? 2,
        cameraSeq: change.cameraSeq ?? 3,
        changedTileIds: change.changedTileIds ?? [],
        invalidatedTileIds: change.invalidatedTileIds ?? [],
        structural: false,
      }),
    setTile: (tile) => byId.set(tile.tileId, tile),
    source: {
      peekRenderTile: (tileId) => byId.get(tileId) ?? null,
      subscribeRenderTileDeltas: (_subscription, nextListener) => {
        listener = nextListener
        return () => {
          listener = null
          unsubscribed = true
        }
      },
    },
    unsubscribed: () => unsubscribed,
  }
}

export function createWorkbookDeltaSource(): {
  readonly source: GridRenderTileSource
  readonly emit: (batch: WorkbookDeltaBatchLikeV3) => void
  readonly unsubscribed: () => boolean
} {
  let listener: ((batch: WorkbookDeltaBatchLikeV3) => void) | null = null
  let unsubscribed = false
  return {
    emit: (batch) => listener?.(batch),
    source: {
      peekRenderTile: () => null,
      subscribeRenderTileDeltas: () => () => {},
      subscribeWorkbookDeltas: (nextListener) => {
        listener = nextListener
        return () => {
          listener = null
          unsubscribed = true
        }
      },
    },
    unsubscribed: () => unsubscribed,
  }
}

export function createMutableWorkbookDeltaRenderTileSource(tiles: readonly GridRenderTile[] = []): {
  readonly source: GridRenderTileSource
  readonly deleteTile: (tileId: number) => void
  readonly emitWorkbookDelta: (batch: WorkbookDeltaBatchLikeV3) => void
  readonly emitRenderTileDelta: (change: {
    readonly batchId?: number | undefined
    readonly cameraSeq?: number | undefined
    readonly changedTileIds?: readonly number[] | undefined
    readonly invalidatedTileIds?: readonly number[] | undefined
  }) => void
  readonly setTile: (tile: GridRenderTile) => void
} {
  const renderTiles = createMutableRenderTileSource(tiles)
  let workbookDeltaListener: ((batch: WorkbookDeltaBatchLikeV3) => void) | null = null
  return {
    deleteTile: renderTiles.deleteTile,
    emitRenderTileDelta: (change) => renderTiles.emit(change),
    emitWorkbookDelta: (batch) => workbookDeltaListener?.(batch),
    setTile: renderTiles.setTile,
    source: {
      peekRenderTile: (tileId) => renderTiles.source.peekRenderTile(tileId),
      subscribeRenderTileDeltas: (subscription, listener) => renderTiles.source.subscribeRenderTileDeltas(subscription, listener),
      subscribeWorkbookDeltas: (listener) => {
        workbookDeltaListener = listener
        return () => {
          workbookDeltaListener = null
        }
      },
    },
  }
}

export function createWorkbookDeltaBatch(overrides: Partial<WorkbookDeltaBatchLikeV3> = {}): WorkbookDeltaBatchLikeV3 {
  return {
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: new Uint32Array([0, 0, 0, 0, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
    },
    seq: 1,
    sheetId: 7,
    sheetOrdinal: 7,
    ...overrides,
  }
}

export function createInput(
  overrides: Partial<Parameters<GridRenderTilePaneRuntime['resolve']>[0]> = {},
): Parameters<GridRenderTilePaneRuntime['resolve']>[0] {
  return {
    columnWidths: {},
    dprBucket: 1,
    engine: TEST_ENGINE,
    freezeCols: 0,
    freezeRows: 0,
    frozenColumnWidth: 0,
    frozenRowHeight: 0,
    gridMetrics: getGridMetrics(),
    gridRuntimeHost: createHost(),
    hostClientHeight: 400,
    hostClientWidth: 800,
    hostReady: true,
    renderTileSource: undefined,
    renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    residentViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    rowHeights: {},
    sceneRevision: 1,
    sheetId: 7,
    sheetName: 'Sheet1',
    sortedColumnWidthOverrides: [],
    sortedRowHeightOverrides: [],
    visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    ...overrides,
  }
}
