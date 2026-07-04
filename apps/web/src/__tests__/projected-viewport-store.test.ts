import { SpreadsheetEngine } from '@bilig/core'
import { getGridMetrics } from '@bilig/grid'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import type { ViewportPatch, WorkerEngineClient } from '@bilig/worker-transport'
import { describe, expect, it, vi } from 'vitest'
import { buildLocalFixedRenderTiles } from '../../../../packages/grid/src/renderer-v3/local-render-tile-materializer.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../../../../packages/grid/src/renderer-v3/rect-instance-buffer.js'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { ProjectedViewportStore } from '../projected-viewport-store.js'
import { LOCAL_CELL_VISUAL_DIRTY_MASK } from '../projected-workbook-local-delta.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from '../workbook-optimistic-cell-flags.js'
import { applyOptimisticClearRange } from '../workbook-optimistic-range.js'
import type { WorkerEngine } from '../worker-runtime-support.js'
import { buildViewportPatchFromEngine, DEFAULT_STYLE_ID } from '../worker-runtime-viewport.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}
const LOCAL_FILL_DIRTY_MASK = DirtyMaskV3.Style | DirtyMaskV3.Rect

function createPatch(styleId?: string): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 0,
    freezeCols: 0,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 3,
      rowEnd: 7,
      colStart: 2,
      colEnd: 4,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [
      {
        row: 4,
        col: 3,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'D5',
          value: { tag: ValueTag.Empty },
          flags: 0,
          version: 1,
          ...(styleId ? { styleId } : {}),
        },
        displayText: '',
        copyText: '',
        editorText: '',
        formatId: 0,
        styleId: styleId ?? 'style-0',
      },
    ],
    columns: [],
    rows: [],
  }
}

function countSheetCells(cache: ProjectedViewportStore, sheetName: string): number {
  let count = 0
  cache.workbook.getSheet(sheetName)?.grid.forEachCellEntry(() => {
    count += 1
  })
  return count
}

function hasOpaqueGreenFillRect(rectInstances: Float32Array, rectCount: number): boolean {
  for (let index = 0; index < rectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const red = rectInstances[offset + 4] ?? 1
    const green = rectInstances[offset + 5] ?? 0
    const blue = rectInstances[offset + 6] ?? 1
    const alpha = rectInstances[offset + 7] ?? 0
    const instanceKind = rectInstances[offset + 13] ?? -1
    if (instanceKind === 0 && red < 0.05 && green > 0.95 && blue < 0.05 && alpha > 0.95) {
      return true
    }
  }
  return false
}

function createNoopWorkerEngineClient(): WorkerEngineClient {
  return {
    dispose: vi.fn(),
    invoke: vi.fn(async () => undefined),
    ready: vi.fn(async () => undefined),
    subscribe: vi.fn(() => () => undefined),
    subscribeBatches: vi.fn(() => () => undefined),
    subscribeRenderTileDeltas: vi.fn(() => () => undefined),
    subscribeViewportPatches: vi.fn(() => () => undefined),
    subscribeWorkbookDeltas: vi.fn(() => () => undefined),
  }
}

describe('ProjectedViewportStore style projection and render revisions', () => {
  it('renders range formatting for empty cells when the patch carries style ids outside the snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      styles: [{ id: 'style-green', fill: { backgroundColor: '#00ff00' } }],
      cells: createPatch().cells.map((cell) => Object.assign({}, cell, { styleId: 'style-green' })),
    })

    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-green')
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'D5').styleId)).toEqual({
      id: 'style-green',
      fill: { backgroundColor: '#00ff00' },
    })
  })

  it('hydrates authoritative snapshot style ranges for empty cells into the viewport cache', async () => {
    const seed = new SpreadsheetEngine({ workbookName: 'viewport-style-range-seed' })
    await seed.ready()
    seed.createSheet('Sheet1')
    seed.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'E6', endAddress: 'E6' }, { fill: { backgroundColor: '#00ff00' } })

    const restored = new SpreadsheetEngine({ workbookName: 'viewport-style-range-restored' }) as SpreadsheetEngine & WorkerEngine
    await restored.ready()
    restored.importSnapshot(seed.exportSnapshot())
    const cache = new ProjectedViewportStore()
    const patch = buildViewportPatchFromEngine({
      authoritativeRevision: 3,
      emptyCellSnapshot: (sheetName, address) => ({
        sheetName,
        address,
        flags: 0,
        value: { tag: ValueTag.Empty },
        version: 0,
      }),
      engine: restored,
      event: null,
      getFormatId: () => 0,
      getStyleRecord: (styleId) => restored.getCellStyle(styleId) ?? { id: DEFAULT_STYLE_ID },
      metrics: { ...TEST_METRICS, batchId: 3 },
      sheetImpact: null,
      state: {
        knownStyleIds: new Set(),
        lastCellSignatures: new Map(),
        lastColumnSignatures: new Map(),
        lastMergeSignatures: new Map(),
        lastRowSignatures: new Map(),
        lastStyleSignatures: new Map(),
        listener: () => undefined,
        nextVersion: 1,
        subscription: {
          sheetName: 'Sheet1',
          rowStart: 0,
          rowEnd: 31,
          colStart: 0,
          colEnd: 127,
        },
      },
    })

    cache.applyViewportPatch(patch)

    expect(restored.getCell('Sheet1', 'E6').styleId).toBeDefined()
    expect(cache.getCell('Sheet1', 'E6').styleId).toBe(restored.getCell('Sheet1', 'E6').styleId)
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'E6').styleId)).toEqual({
      id: restored.getCell('Sheet1', 'E6').styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)
  })

  it('optimistically styles visible empty cells and publishes precise fill tile damage', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const deltaListener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(deltaListener)
    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'D5' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    expect(rollback).toEqual(expect.any(Function))
    const styledCell = cache.getCell('Sheet1', 'D5')
    expect(styledCell.styleId).toMatch(/^style-local-/)
    expect(cache.getCellStyle(styledCell.styleId)).toEqual({
      id: styledCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, LOCAL_FILL_DIRTY_MASK]),
        },
        source: 'localOptimistic',
      }),
    )

    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)

    deltaListener.mockClear()
    rollback?.()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, LOCAL_FILL_DIRTY_MASK]),
        },
        source: 'localOptimistic',
      }),
    )
    unsubscribeViewport()
    unsubscribeDeltas()
  })

  it('publishes style-only damage for bold applied to a visible empty cell', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const deltaListener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(deltaListener)
    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 12,
        colStart: 0,
        colEnd: 12,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    const rollback = cache.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'D5' }, { font: { bold: true } })

    expect(rollback).toEqual(expect.any(Function))
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, DirtyMaskV3.Style]),
        },
        source: 'localOptimistic',
      }),
    )

    deltaListener.mockClear()
    rollback?.()
    expect(deltaListener).toHaveBeenCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 4, 3, 3, DirtyMaskV3.Style]),
        },
        source: 'localOptimistic',
      }),
    )
    unsubscribeViewport()
    unsubscribeDeltas()
  })

  it('resolves large range styles for cells that become visible after the mutation', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'F900' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    expect(rollback).toEqual(expect.any(Function))
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    const offscreenCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(offscreenCell.styleId)).toEqual({
      id: offscreenCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    const deltaListener = vi.fn()
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(deltaListener)

    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 695,
        rowEnd: 704,
        colStart: 3,
        colEnd: 5,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(deltaListener).toHaveBeenCalledTimes(1)
    expect(deltaListener).toHaveBeenLastCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: Uint32Array.from(
            Array.from({ length: 30 }, (_, index) => {
              const row = 695 + Math.floor(index / 3)
              const col = 3 + (index % 3)
              return [row, row, col, col, LOCAL_CELL_VISUAL_DIRTY_MASK]
            }).flat(),
          ),
        },
        source: 'localOptimistic',
      }),
    )
    expect(countSheetCells(cache, 'Sheet1')).toBeLessThanOrEqual(30)
    const materializedCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(materializedCell.styleId)).toEqual({
      id: materializedCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    const tiles = buildLocalFixedRenderTiles({
      cameraSeq: 1,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 3,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 695,
        rowEnd: 704,
        colStart: 3,
        colEnd: 5,
      },
    })
    expect(tiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(true)

    deltaListener.mockClear()
    rollback?.()
    expect(deltaListener).toHaveBeenCalledTimes(1)
    expect(deltaListener).toHaveBeenLastCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: Uint32Array.from(
            Array.from({ length: 30 }, (_, index) => {
              const row = 695 + Math.floor(index / 3)
              const col = 3 + (index % 3)
              return [row, row, col, col, LOCAL_CELL_VISUAL_DIRTY_MASK]
            }).flat(),
          ),
        },
        source: 'localOptimistic',
      }),
    )
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.peekCell('Sheet1', 'E700')).toBeUndefined()
    const rollbackTiles = buildLocalFixedRenderTiles({
      cameraSeq: 2,
      columnWidths: cache.getColumnWidths('Sheet1'),
      dprBucket: 1,
      engine: cache,
      generation: 4,
      gridMetrics: getGridMetrics(),
      rowHeights: cache.getRowHeights('Sheet1'),
      sheetId: 7,
      sheetOrdinal: 7,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      viewport: {
        sheetName: 'Sheet1',
        rowStart: 695,
        rowEnd: 704,
        colStart: 3,
        colEnd: 5,
      },
    })
    expect(rollbackTiles.some((tile) => hasOpaqueGreenFillRect(tile.rectInstances, tile.rectCount))).toBe(false)
    unsubscribeViewport()
    unsubscribeDeltas()
  })

  it('publishes coarse local range deltas for large background overlays set and cleared before viewport materialization', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const deltaListener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(deltaListener)
    const range = { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'F900' }

    const rollbackSet = cache.setRangeStyle(range, { fill: { backgroundColor: '#00ff00' } })

    expect(rollbackSet).toEqual(expect.any(Function))
    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(1)
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'E700').styleId)).toEqual({
      id: cache.getCell('Sheet1', 'E700').styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(deltaListener).toHaveBeenCalledTimes(1)
    expect(deltaListener).toHaveBeenLastCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 899, 3, 5, LOCAL_CELL_VISUAL_DIRTY_MASK]),
        },
        seq: 1,
        source: 'localOptimistic',
      }),
    )

    const rollbackClear = cache.clearRangeStyle(range, ['backgroundColor'])

    expect(rollbackClear).toEqual(expect.any(Function))
    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(2)
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.getCellStyle(cache.getCell('Sheet1', 'E700').styleId)?.fill).toBeUndefined()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(deltaListener).toHaveBeenCalledTimes(2)
    expect(deltaListener).toHaveBeenLastCalledWith(
      expect.objectContaining({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([4, 899, 3, 5, LOCAL_CELL_VISUAL_DIRTY_MASK]),
        },
        seq: 2,
        source: 'localOptimistic',
      }),
    )

    unsubscribeDeltas()
  })

  it('lets newer small style edits override an older large range overlay', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())

    const rollbackLarge = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'F900' },
      { fill: { backgroundColor: '#00ff00' } },
    )
    const rollbackSmall = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'E700', endAddress: 'E700' },
      { fill: { backgroundColor: '#a4c2f4' } },
    )

    const overriddenCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(overriddenCell.styleId)).toEqual({
      id: overriddenCell.styleId,
      fill: { backgroundColor: '#a4c2f4' },
    })
    const neighborCell = cache.getCell('Sheet1', 'D700')
    expect(cache.getCellStyle(neighborCell.styleId)).toEqual({
      id: neighborCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    rollbackSmall?.()
    const restoredCell = cache.getCell('Sheet1', 'E700')
    expect(cache.getCellStyle(restoredCell.styleId)).toEqual({
      id: restoredCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })
    rollbackLarge?.()
    expect(cache.peekCell('Sheet1', 'E700')).toBeUndefined()
  })

  it('keeps large clears visible for future viewports without materializing the whole range', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())

    const rollback = applyOptimisticClearRange(cache, {
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D3000',
    })

    expect(rollback).toEqual(expect.any(Function))
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    const optimisticClear = cache.peekCell('Sheet1', 'D2500')
    expect(optimisticClear).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      value: { tag: ValueTag.Empty },
    })
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)

    const unsubscribeViewport = cache.subscribeViewport(
      'Sheet1',
      {
        sheetName: 'Sheet1',
        rowStart: 2495,
        rowEnd: 2504,
        colStart: 0,
        colEnd: 3,
      },
      () => undefined,
      { initialPatch: 'none' },
    )

    expect(countSheetCells(cache, 'Sheet1')).toBeLessThanOrEqual(40)
    expect(cache.getCell('Sheet1', 'D2500')).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      value: { tag: ValueTag.Empty },
    })

    rollback?.()
    expect(countSheetCells(cache, 'Sheet1')).toBe(0)
    expect(cache.peekCell('Sheet1', 'D2500')).toBeUndefined()
    unsubscribeViewport()
  })

  it('optimistically styles protected local edit snapshots instead of waiting for worker readback', () => {
    const cache = new ProjectedViewportStore()
    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'moved-fill-proof',
      value: { tag: ValueTag.String, value: 'moved-fill-proof', stringId: 0 },
      version: 4,
    })

    const rollback = cache.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'D5', endAddress: 'D5' },
      { fill: { backgroundColor: '#00ff00' } },
    )

    const styledCell = cache.getCell('Sheet1', 'D5')
    expect(styledCell).toMatchObject({
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'moved-fill-proof',
      value: { tag: ValueTag.String, value: 'moved-fill-proof' },
    })
    expect(cache.getCellStyle(styledCell.styleId)).toEqual({
      id: styledCell.styleId,
      fill: { backgroundColor: '#00ff00' },
    })

    rollback?.()
    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('accepts equal-version empty snapshots that clear stale styling', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    expect(cache.getCell('Sheet1', 'D5').styleId).toBe('style-red')

    cache.applyViewportPatch(createPatch())

    expect(cache.getCell('Sheet1', 'D5').styleId).toBeUndefined()
  })

  it('exposes authoritative, projected, and tile-scene render revisions for visible proof checks', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 17,
      metrics: {
        ...TEST_METRICS,
        batchId: 23,
      },
    })

    expect(cache.getRenderRevisionSnapshot()).toEqual({
      authoritativeRevision: 17,
      localRevision: 0,
      projectedRevision: 23,
      tileSceneCameraSeq: null,
      tileSceneRevision: null,
    })
  })

  it('rejects stale authoritative viewport patches before they regress rendered cells, axes, or proof revisions', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 17,
      metrics: {
        ...TEST_METRICS,
        batchId: 23,
      },
      columns: [{ index: 0, size: 93, hidden: true }],
      rows: [{ index: 0, size: 44, hidden: true }],
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'fresh', stringId: 1 },
            input: 'fresh',
            flags: 0,
            version: 17,
          },
          displayText: 'fresh',
          copyText: 'fresh',
          editorText: 'fresh',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const staleDamage = cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 12,
      metrics: {
        ...TEST_METRICS,
        batchId: 18,
      },
      freezeRows: 4,
      freezeCols: 3,
      columns: [{ index: 0, size: 68, hidden: false }],
      rows: [{ index: 0, size: 30, hidden: false }],
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.String, value: 'stale', stringId: 2 },
            input: 'stale',
            flags: 0,
            version: 99,
          },
          displayText: 'stale',
          copyText: 'stale',
          editorText: 'stale',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(staleDamage).toEqual([])
    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'fresh', stringId: 1 },
      input: 'fresh',
      version: 17,
    })
    expect(cache.getColumnWidths('Sheet1')[0]).toBe(0)
    expect(cache.getColumnSizes('Sheet1')[0]).toBe(93)
    expect(cache.getHiddenColumns('Sheet1')[0]).toBe(true)
    expect(cache.getRowHeights('Sheet1')[0]).toBe(0)
    expect(cache.getRowSizes('Sheet1')[0]).toBe(44)
    expect(cache.getHiddenRows('Sheet1')[0]).toBe(true)
    expect(cache.getFreezeRows('Sheet1')).toBe(0)
    expect(cache.getFreezeCols('Sheet1')).toBe(0)
    expect(cache.getRenderRevisionSnapshot()).toEqual({
      authoritativeRevision: 17,
      localRevision: 0,
      projectedRevision: 23,
      tileSceneCameraSeq: null,
      tileSceneRevision: null,
    })
  })

  it('increments the local render revision for optimistic cell snapshots', () => {
    const cache = new ProjectedViewportStore()

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: 0,
      input: 'local',
      value: { tag: ValueTag.String, value: 'local', stringId: 0 },
      version: 1,
    })

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(1)
  })

  it('publishes local dirty ranges when optimistic trust flags are cleared', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 3 }])
    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'B2',
        flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
        value: { tag: ValueTag.Number, value: 17 },
        version: 12,
      },
      { emitLocalDelta: false },
    )
    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'D4',
        flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
        value: { tag: ValueTag.String, value: 'trusted', stringId: 9 },
        version: 15,
      },
      { emitLocalDelta: false },
    )
    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'E4',
        flags: 0,
        value: { tag: ValueTag.String, value: 'stable', stringId: 10 },
        version: 16,
      },
      { emitLocalDelta: false },
    )
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)

    cache.clearOptimisticCellFlagsForSheet('Sheet1')

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(1)
    expect(cache.getCell('Sheet1', 'B2').flags).toBe(0)
    expect(cache.getCell('Sheet1', 'D4').flags).toBe(0)
    expect(cache.getCell('Sheet1', 'E4').flags).toBe(0)
    expect(listener).toHaveBeenCalledTimes(1)
    const batch = listener.mock.calls[0]?.[0]
    expect(batch).toMatchObject({
      calcSeq: 15,
      seq: 1,
      sheetId: 7,
      sheetOrdinal: 3,
      source: 'localOptimistic',
      styleSeq: 15,
      valueSeq: 15,
    })
    expect(batch?.dirty.axisX).toEqual(new Uint32Array())
    expect(batch?.dirty.axisY).toEqual(new Uint32Array())
    expect(batch?.dirty.cellRanges).toEqual(
      new Uint32Array([1, 1, 1, 1, LOCAL_CELL_VISUAL_DIRTY_MASK, 3, 3, 3, 3, LOCAL_CELL_VISUAL_DIRTY_MASK]),
    )

    listener.mockClear()
    cache.clearOptimisticCellFlagsForSheet('Sheet1')

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(1)
    expect(listener).not.toHaveBeenCalled()

    unsubscribeDeltas()
  })

  it('hydrates selected-cell cache without publishing render tile deltas', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const listener = vi.fn()
    cache.setSheetIdentities([{ id: 7, name: 'Sheet1', order: 0 }])
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas(listener)

    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)

    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'D5',
        flags: 0,
        input: 'hydrated',
        value: { tag: ValueTag.String, value: 'hydrated', stringId: 0 },
        version: 12,
      },
      { emitLocalDelta: false },
    )

    expect(cache.getCell('Sheet1', 'D5').input).toBe('hydrated')
    expect(cache.getRenderRevisionSnapshot().localRevision).toBe(0)
    expect(listener).not.toHaveBeenCalled()

    unsubscribeDeltas()
  })

  it('can force authoritative selected-cell hydration over stale optimistic input', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      input: 'ghost-content',
      value: { tag: ValueTag.String, value: 'ghost-content', stringId: 0 },
      version: 8,
    })

    cache.setCellSnapshot(
      {
        sheetName: 'Sheet1',
        address: 'D5',
        flags: 0,
        value: { tag: ValueTag.Empty },
        version: 0,
      },
      { force: true, forceOptimistic: true },
    )

    expect(cache.getCell('Sheet1', 'D5')).toEqual({
      sheetName: 'Sheet1',
      address: 'D5',
      flags: 0,
      value: { tag: ValueTag.Empty },
      version: 0,
    })
  })

  it('accepts partial reset-empty patches that clear newer local input after structural deletes', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'stale-tail', stringId: 1 },
      input: 'stale-tail',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      authoritativeRevision: 0,
      full: false,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5').value).toEqual({ tag: ValueTag.Empty })
    expect(cache.getCell('Sheet1', 'D5').input).toBeUndefined()
  })

  it('keeps newer local input when a full reset-empty patch arrives during hydration', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.String, value: 'local-edit', stringId: 1 },
      input: 'local-edit',
      flags: 0,
      version: 7,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Empty },
            flags: 0,
            version: 0,
          },
          displayText: '',
          copyText: '',
          editorText: '',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.String, value: 'local-edit' },
      input: 'local-edit',
      version: 7,
    })
  })

  it('keeps an equal-version local formula snapshot when a later patch drops the formula', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 3,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a newer eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps a local formula snapshot when a direct cell refresh drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: 0,
      version: 3,
    })

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'D5',
      value: { tag: ValueTag.Boolean, value: false },
      flags: 0,
      version: 4,
    })

    expect(cache.getCell('Sheet1', 'D5')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      version: 3,
    })
  })

  it('keeps an optimistic formula snapshot when an eval-only patch drops source metadata', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A2',
      value: { tag: ValueTag.Boolean, value: true },
      input: '=A1="HELLO"',
      formula: 'A1="HELLO"',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 1,
          col: 0,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'A2',
            value: { tag: ValueTag.Boolean, value: false },
            flags: 0,
            version: 1,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    expect(cache.getCell('Sheet1', 'A2')).toMatchObject({
      value: { tag: ValueTag.Boolean, value: true },
      formula: 'A1="HELLO"',
      flags: OPTIMISTIC_CELL_SNAPSHOT_FLAG,
      version: 1,
    })
  })

  it('accepts a newer literal snapshot when the source input is present', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: true },
            input: '=A1="HELLO"',
            formula: 'A1="HELLO"',
            flags: 0,
            version: 3,
          },
          displayText: 'TRUE',
          copyText: '=A1="HELLO"',
          editorText: '=A1="HELLO"',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    cache.applyViewportPatch({
      ...createPatch(),
      cells: [
        {
          row: 4,
          col: 3,
          snapshot: {
            sheetName: 'Sheet1',
            address: 'D5',
            value: { tag: ValueTag.Boolean, value: false },
            input: false,
            flags: 0,
            version: 4,
          },
          displayText: 'FALSE',
          copyText: 'FALSE',
          editorText: 'FALSE',
          formatId: 0,
          styleId: 'style-0',
        },
      ],
    })

    const snapshot = cache.getCell('Sheet1', 'D5')
    expect(snapshot).toMatchObject({
      value: { tag: ValueTag.Boolean, value: false },
      input: false,
      version: 4,
    })
    expect('formula' in snapshot).toBe(false)
  })

  it('reports damage when a style record changes without a newer cell snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#c9daf8' } }],
    })

    const damage = cache.applyViewportPatch({
      ...createPatch('style-fill'),
      styles: [{ id: 'style-fill', fill: { backgroundColor: '#a4c2f4' } }],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.getCellStyle('style-fill')).toEqual({
      id: 'style-fill',
      fill: { backgroundColor: '#a4c2f4' },
    })
  })

  it('tracks freeze pane metadata from viewport patches', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 2,
      freezeCols: 1,
    })

    expect(cache.getFreezeRows('Sheet1')).toBe(2)
    expect(cache.getFreezeCols('Sheet1')).toBe(1)
  })

  it('does not notify freeze subscribers when a patch confirms the default unfrozen state', () => {
    const cache = new ProjectedViewportStore()
    const freezeListener = vi.fn()

    const unsubscribeFreeze = cache.subscribeSheetChannel('Sheet1', 'freeze', freezeListener)

    cache.applyViewportPatch({
      ...createPatch(),
      freezeRows: 0,
      freezeCols: 0,
    })

    expect(freezeListener).not.toHaveBeenCalled()

    unsubscribeFreeze()
  })

  it('clears stale viewport cells on full patches without dropping cells outside the viewport', () => {
    const cache = new ProjectedViewportStore()

    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'A1',
      value: { tag: ValueTag.String, value: 'pinned', stringId: 1 },
      flags: 0,
      version: 1,
    })
    cache.applyViewportPatch({ ...createPatch(), full: true })

    const damage = cache.applyViewportPatch({
      ...createPatch(),
      full: true,
      cells: [],
    })

    expect(damage).toEqual([{ cell: [3, 4] }])
    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.String,
      value: 'pinned',
      stringId: 1,
    })
  })

  it('drops stale sheet cache entries when sheets disappear', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch())
    expect(cache.peekCell('Sheet1', 'D5')).toBeDefined()

    cache.setKnownSheets(['Sheet2'])

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
  })

  it('resets same-sheet projected state before installing a replacement authoritative snapshot', () => {
    const cache = new ProjectedViewportStore()

    cache.applyViewportPatch(createPatch('style-red'))
    cache.setColumnWidth('Sheet1', 0, 68)
    cache.setRowHeight('Sheet1', 0, 240)

    cache.resetProjectionState()

    expect(cache.peekCell('Sheet1', 'D5')).toBeUndefined()
    expect(cache.workbook.getSheet('Sheet1')).toBeDefined()
    expect(cache.getColumnWidths('Sheet1')[0]).toBeUndefined()
    expect(cache.getRowHeights('Sheet1')[0]).toBeUndefined()
  })

  it('publishes sparse local column axis deltas without materializing default axis sizes', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'columnWidths', () => {
      events.push(`axis:${cache.getColumnWidths('Sheet1')[2]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${cache.getColumnWidths('Sheet1')[2]}:${[...batch.dirty.axisX].join(':')}:${batch.dirty.axisY.length}`)
    })

    cache.setColumnWidth('Sheet1', 2, 144)

    expect(events).toEqual(['axis:144', 'delta:144:2:2:44:0'])
    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })

  it('can apply optimistic axis sizes without publishing duplicate local tile deltas', () => {
    const cache = new ProjectedViewportStore(createNoopWorkerEngineClient())
    const events: string[] = []
    const unsubscribeRenderTiles = cache.subscribeRenderTileDeltas(
      {
        sheetId: 7,
        sheetName: 'Sheet1',
        sheetOrdinal: 3,
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 63,
      },
      () => undefined,
    )
    const unsubscribeAxis = cache.subscribeSheetChannel('Sheet1', 'columnWidths', () => {
      events.push(`axis:${cache.getColumnWidths('Sheet1')[2]}`)
    })
    const unsubscribeDeltas = cache.subscribeWorkbookDeltas((batch) => {
      events.push(`delta:${batch.dirty.axisX.length}`)
    })

    cache.setColumnWidth('Sheet1', 2, 144, { emitLocalDelta: false })

    expect(events).toEqual(['axis:144'])
    expect(cache.getColumnWidths('Sheet1')).toEqual({ 2: 144 })

    unsubscribeAxis()
    unsubscribeDeltas()
    unsubscribeRenderTiles()
  })
})
