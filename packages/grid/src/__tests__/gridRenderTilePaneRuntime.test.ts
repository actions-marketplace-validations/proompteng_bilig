import type { CellStyleRecord, GridEngineLike, GridRenderTile } from './gridRenderTilePaneRuntime-test-helpers.js'
import {
  DEFAULT_TEST_FONT,
  DEFAULT_TEST_FONT_SIZE,
  DEFAULT_TEST_TEXT_COLOR,
  DirtyMaskV3,
  GRID_TEXT_METRIC_FLOAT_COUNT_V3,
  GridRenderTilePaneRuntime,
  LOCAL_EMPTY_ENGINE,
  createCapturingRenderTileSource,
  createEmptyCellSnapshot,
  createGridBorderRectInstances,
  createHost,
  createInput,
  createMutableWorkbookDeltaRenderTileSource,
  createRenderTile,
  createRenderTileSource,
  createStringCellSnapshot,
  createStyledStringCellSnapshot,
  createWorkbookDeltaBatch,
  describe,
  expect,
  expectedGridBorderRectCount,
  formatAddress,
  getGridMetrics,
  getGridRenderTilePaneRuntime,
  hasOpaqueGreenFillRect,
  it,
  materializeGridRenderTileV3,
  packTileKey53,
} from './gridRenderTilePaneRuntime-test-helpers.js'

describe('GridRenderTilePaneRuntime remote and local pane resolution', () => {
  it('replaces stale runtime refs from live reloads', () => {
    const runtime = new GridRenderTilePaneRuntime()

    expect(getGridRenderTilePaneRuntime(runtime)).toBe(runtime)
    expect(getGridRenderTilePaneRuntime({})).toBeInstanceOf(GridRenderTilePaneRuntime)
    expect(getGridRenderTilePaneRuntime(null)).toBeInstanceOf(GridRenderTilePaneRuntime)
  })

  it('publishes bridge revisions from the runtime-owned external store', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const snapshots: unknown[] = []
    const unsubscribe = runtime.subscribeBridgeState(() => {
      snapshots.push(runtime.snapshotBridgeState())
    })

    runtime.noteRenderTileDelta()
    runtime.noteWorkbookDeltaDamage()
    runtime.noteLocalFallbackInvalidation()
    unsubscribe()
    runtime.noteRenderTileDelta()

    expect(snapshots).toEqual([
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 1,
      },
      {
        forceLocalTiles: false,
        localFallbackRevision: 0,
        renderTileRevision: 2,
      },
      {
        forceLocalTiles: true,
        localFallbackRevision: 1,
        renderTileRevision: 2,
      },
    ])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 3,
    })
  })

  it('resolves remote render tiles into V3 pane placements', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const state = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )

    expect(state.residentBodyPane?.tile.tileId).toBe(tileId)
    expect(state.needsLocalCellInvalidation).toBe(false)
    expect(state.renderTilePanes).toHaveLength(1)
    expect(state.preloadDataPanes).toHaveLength(0)
    expect(state.tileReadiness).toMatchObject({
      exactHits: [tileId],
      misses: [],
      staleHits: [],
      visibleDirtyTileKeys: [],
    })
    expect(host.tiles.residency.getExact(tileId)?.packet).toMatchObject({
      tileId,
      version: {
        values: 1,
      },
    })
  })

  it('resolves available warm remote tiles into the V3 preload lane without drawing them', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const visibleTileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const warmTileId = packTileKey53({
      colTile: 1,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })
    const baseWarmTile = createRenderTile(warmTileId)
    const warmTile: GridRenderTile = {
      ...baseWarmTile,
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        ...baseWarmTile.coord,
        colTile: 1,
        rowTile: 0,
      },
    }

    const state = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(visibleTileId), warmTile]),
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual([visibleTileId])
    expect(state.preloadDataPanes.map((pane) => pane.tile.tileId)).toContain(warmTileId)
    expect(state.preloadDataPanes.map((pane) => pane.tile.tileId)).not.toContain(visibleTileId)
    expect(host.tiles.residency.getExact(warmTileId)?.packet).toMatchObject({
      tileId: warmTileId,
    })
  })

  it('refreshes host-owned tile revisions when remote tile contents update', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([
          {
            ...createRenderTile(tileId),
            version: {
              axisX: 5,
              axisY: 6,
              freeze: 7,
              styles: 8,
              text: 9,
              values: 10,
            },
          },
        ]),
      }),
    )

    expect(host.tiles.residency.getExact(tileId)).toMatchObject({
      axisSeqX: 5,
      axisSeqY: 6,
      freezeSeq: 7,
      styleSeq: 8,
      textSeq: 9,
      valueSeq: 10,
    })
  })

  it('retains the previous same-sheet panes while a remote tile is temporarily unavailable', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const ready = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    const missing = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(missing.residentDataPanes.map((pane) => pane.tile.tileId)).toEqual([tileId])
    expect(missing.residentBodyPane?.tile).toBe(ready.residentBodyPane?.tile)
    expect(missing.needsLocalCellInvalidation).toBe(false)
  })

  it('does not resurrect retained background fills after local style clear when the remote source misses', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const fullVisualDirtyMask = DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect | DirtyMaskV3.Border
    const styles: Record<string, CellStyleRecord> = {
      'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
    }
    const greenFillEngine: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A1' ? createStyledStringCellSnapshot('A1', 'A1', 'style-green') : createEmptyCellSnapshot(address),
      getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const remoteGreenTile = materializeGridRenderTileV3({
      axisSeqX: 1,
      axisSeqY: 1,
      cameraSeq: 1,
      columnWidths: {},
      dprBucket: 1,
      engine: greenFillEngine,
      freezeSeq: 0,
      gridMetrics: getGridMetrics(),
      materializedAtSeq: 1,
      packetSeq: 1,
      rectSeq: 1,
      rowHeights: {},
      sheetId: 7,
      sheetName: 'Sheet1',
      sheetOrdinal: 7,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 1,
      textSeq: 1,
      valueSeq: 1,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    expect(remoteGreenTile.tileId).toBe(tileId)
    expect(hasOpaqueGreenFillRect(remoteGreenTile)).toBe(true)
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([remoteGreenTile])

    const ready = runtime.resolve(
      createInput({
        engine: greenFillEngine,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(ready.residentBodyPane?.tile).toBe(remoteGreenTile)
    expect(hasOpaqueGreenFillRect(ready.residentBodyPane?.tile)).toBe(true)

    runtime.connectWorkbookDeltaDamage({
      dprBucket: 1,
      gridRuntimeHost: host,
      renderTileSource: renderTileSource.source,
      sheetId: 7,
    })
    renderTileSource.emitWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([0, 0, 0, 0, fullVisualDirtyMask]),
        },
        source: 'localOptimistic',
      }),
    )
    renderTileSource.deleteTile(tileId)

    const cleared = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )

    expect(cleared.residentBodyPane?.tile).not.toBe(remoteGreenTile)
    expect(hasOpaqueGreenFillRect(cleared.residentBodyPane?.tile)).toBe(false)
    expect(cleared.residentBodyPane?.tile.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(cleared.residentBodyPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 0]))
    expect(cleared.residentBodyPane?.tile.dirtyMasks).toEqual(new Uint32Array([fullVisualDirtyMask]))
    expect(host.tiles.dirtyTiles.getUnconsumedMask(tileId)).toBe(0)
  })

  it('builds a local visible tile when remote tiles are unavailable before same-sheet retention exists', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(state.residentBodyPane?.tile.tileId).toBe(tileId)
    expect(state.residentBodyPane?.tile.rectCount).toBeGreaterThan(0)
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.residentDataPanes).toHaveLength(1)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('fills visible remote tile holes with local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0])]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness).toMatchObject({
      exactHits: tileIds,
      misses: [],
    })
  })

  it('localizes a selected-cell tile when the remote tile is missing selected text', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const state = runtime.resolve(
      createInput({
        engine: {
          ...LOCAL_EMPTY_ENGINE,
          getCell: (_sheetName, address) =>
            address === 'D53' ? createStringCellSnapshot('D53', 'Month 1') : createEmptyCellSnapshot(address),
        },
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0])]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createStringCellSnapshot('D53', 'Month 1'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).toContain('Month 1')
  })

  it('localizes a selected-cell tile when the remote tile still has text for a cleared cell', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[0]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 300,
          clipY: 400,
          color: '#111827',
          col: 3,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 52,
          strike: false,
          text: 'Month 1',
          underline: false,
          width: 100,
          x: 300,
          y: 400,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createEmptyCellSnapshot('D53'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).not.toContain('Month 1')
  })

  it('localizes an active editor cell tile so remote text cannot show under the editor', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[0]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 300,
          clipY: 400,
          color: '#111827',
          col: 3,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 52,
          strike: false,
          text: 'Month 1',
          underline: false,
          width: 100,
          x: 300,
          y: 400,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        editingCell: [3, 52],
        engine: {
          ...LOCAL_EMPTY_ENGINE,
          getCell: (_sheetName, address) =>
            address === 'D53' ? createStringCellSnapshot('D53', 'Month 1') : createEmptyCellSnapshot(address),
        },
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
        selectedCell: [3, 52],
        selectedCellSnapshot: createStringCellSnapshot('D53', 'Month 1'),
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      }),
    )

    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).not.toContain('Month 1')
  })

  it('localizes visible neighboring tiles while editing so overflow text cannot ghost beside the editor', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const viewport = { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 }
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport,
    })
    const sourceTile = createRenderTile(tileIds[0])
    const adjacentRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      coord: {
        colTile: 1,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 0,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 260,
          clipX: -20,
          clipY: 40,
          color: '#111827',
          col: 127,
          font: '400 12px Arial',
          fontSize: 12,
          height: 20,
          row: 2,
          strike: false,
          text: 'overflow-editor-ghost',
          underline: false,
          width: 260,
          x: -20,
          y: 40,
        },
      ],
    }
    const editAddress = formatAddress(2, 127)
    const state = runtime.resolve(
      createInput({
        editingCell: [127, 2],
        engine: {
          ...LOCAL_EMPTY_ENGINE,
          getCell: (_sheetName, address) =>
            address === editAddress ? createStringCellSnapshot(editAddress, 'overflow-editor-ghost') : createEmptyCellSnapshot(address),
        },
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([sourceTile, adjacentRemoteTile]),
        renderTileViewport: viewport,
        residentViewport: viewport,
        selectedCell: [127, 2],
        selectedCellSnapshot: createStringCellSnapshot(editAddress, 'overflow-editor-ghost'),
        visibleViewport: viewport,
      }),
    )

    const adjacentPane = state.renderTilePanes.find((pane) => pane.tile.tileId === tileIds[1])
    expect(adjacentPane?.tile).not.toBe(adjacentRemoteTile)
    expect(state.renderTilePanes.flatMap((pane) => pane.tile.textRuns.map((run) => run.text))).not.toContain('overflow-editor-ghost')
  })

  it('rebuilds visible remote tiles with missing grid payloads as local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const emptyRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 0,
      textMetrics: new Float32Array(),
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), emptyRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[0]?.tile.rectCount).toBe(expectedGridBorderRectCount(state.renderTilePanes[0].tile.bounds))
    expect(state.renderTilePanes[1]?.tile).not.toBe(emptyRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible remote tiles with partial gridline payloads as local grid tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const partialRectCount = 12
    const partialRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: partialRectCount,
      rectInstances: createGridBorderRectInstances(partialRectCount),
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), partialRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile).not.toBe(partialRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(partialRectCount)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible text-only remote tiles so blank cells keep gridlines', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const textOnlyRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      rectCount: 0,
      rectInstances: new Float32Array(),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 0,
          clipY: 0,
          color: DEFAULT_TEST_TEXT_COLOR,
          col: 0,
          font: DEFAULT_TEST_FONT,
          fontSize: DEFAULT_TEST_FONT_SIZE,
          height: 20,
          row: 32,
          strike: false,
          text: 'remote text without grid rects',
          underline: false,
          width: 120,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), textOnlyRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(tileIds)
    expect(state.renderTilePanes[1]?.tile).not.toBe(textOnlyRemoteTile)
    expect(state.renderTilePanes[1]?.tile.rectCount).toBeGreaterThan(0)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible row-tile boundary text from authoritative cache when remote text is stale', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
    })
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileIds[1]),
      bounds: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 32 },
      coord: {
        colTile: 0,
        dprBucket: 1,
        paneKind: 'body',
        rowTile: 1,
        sheetId: 7,
        sheetOrdinal: 7,
      },
      textCount: 0,
      textMetrics: new Float32Array(GRID_TEXT_METRIC_FLOAT_COUNT_V3),
      textRuns: [],
    }
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) =>
        address === 'C33' ? createStringCellSnapshot(address, 'Annual software subscription') : createEmptyCellSnapshot(address),
    }
    const state = runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileIds[0]), staleRemoteTile]),
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 63, rowStart: 0 },
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 42, rowStart: 11 },
      }),
    )

    const boundaryTile = state.renderTilePanes.find((pane) => pane.tile.bounds.rowStart === 32)?.tile
    expect(boundaryTile).toBeDefined()
    expect(boundaryTile).not.toBe(staleRemoteTile)
    expect(boundaryTile?.textRuns.some((run) => run.row === 32 && run.col === 2 && run.text === 'Annual software subscription')).toBe(true)
    expect(state.tileReadiness.misses).toEqual([])
  })

  it('rebuilds visible remote tiles when deleted local text makes remote text stale', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 100,
          clipY: 80,
          color: '#000000',
          col: 1,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 4,
          strike: false,
          text: 'deleted value',
          underline: false,
          width: 120,
          wrap: false,
          x: 100,
          y: 80,
        },
      ],
    }

    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(state.residentBodyPane?.tile).not.toBe(staleRemoteTile)
    expect(state.residentBodyPane?.tile.textRuns.some((run) => run.text === 'deleted value')).toBe(false)
  })

  it('localizes visible dirty remote tiles even when the stale remote payload looks complete', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const remoteTile = createRenderTile(tileId)
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([0, 0, 0, 0, DirtyMaskV3.Style | DirtyMaskV3.Rect]),
        },
      }),
      { dprBucket: 1 },
    )

    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(state.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(state.residentBodyPane?.tile.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(state.residentBodyPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 0]))
    expect(state.residentBodyPane?.tile.dirtyMasks).toEqual(new Uint32Array([DirtyMaskV3.Style | DirtyMaskV3.Rect]))
    expect(state.tileReadiness.visibleDirtyTileKeys).toContain(tileId)
    expect(host.tiles.dirtyTiles.getUnconsumedMask(tileId)).toBe(0)
  })

  it('keeps clean neighboring remote tiles resident while localizing dirty visible tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [dirtyTileId, cleanTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (dirtyTileId === undefined || cleanTileId === undefined) {
      throw new Error('Expected two render tile keys for the test viewport')
    }
    const dirtyRemoteTile = createRenderTile(dirtyTileId)
    const cleanRemoteTextAddress = formatAddress(0, 128)
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) =>
        address === cleanRemoteTextAddress
          ? createStringCellSnapshot(cleanRemoteTextAddress, 'clean remote text')
          : createEmptyCellSnapshot(address),
    }
    const cleanRemoteTile = materializeGridRenderTileV3({
      axisSeqX: 1,
      axisSeqY: 1,
      cameraSeq: 1,
      columnWidths: {},
      dprBucket: 1,
      engine,
      freezeSeq: 1,
      gridMetrics: getGridMetrics(),
      materializedAtSeq: 1,
      packetSeq: 1,
      rectSeq: 1,
      rowHeights: {},
      sheetId: 7,
      sheetName: 'Sheet1',
      sheetOrdinal: 7,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 1,
      textSeq: 1,
      valueSeq: 1,
      viewport: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
    })
    expect(cleanRemoteTile.tileId).toBe(cleanTileId)
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([0, 0, 0, 0, DirtyMaskV3.Style | DirtyMaskV3.Rect]),
        },
      }),
      { dprBucket: 1 },
    )

    const state = runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([dirtyRemoteTile, cleanRemoteTile]),
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    expect(state.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual([dirtyTileId, cleanTileId])
    expect(state.renderTilePanes[0]?.tile).not.toBe(dirtyRemoteTile)
    expect(state.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(state.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 0]))
    expect(state.renderTilePanes[0]?.tile.dirtyMasks).toEqual(new Uint32Array([DirtyMaskV3.Style | DirtyMaskV3.Rect]))
    expect(state.renderTilePanes[1]?.tile).toBe(cleanRemoteTile)
    expect(state.renderTilePanes[1]?.tile.textRuns[0]?.text).toBe('clean remote text')
    expect(host.tiles.dirtyTiles.getUnconsumedMask(dirtyTileId)).toBe(0)
  })

  it('keeps clean remote tiles resident during pending local projections', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [dirtyTileId, cleanTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (dirtyTileId === undefined || cleanTileId === undefined) {
      throw new Error('Expected two render tile keys for the test viewport')
    }
    const dirtyRemoteTile = createRenderTile(dirtyTileId)
    const cleanRemoteTextAddress = formatAddress(0, 128)
    let localRevision = 0
    let projectedRevision = 1
    let cleanTileReadCount = 0
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) => {
        if (address === cleanRemoteTextAddress) {
          cleanTileReadCount += 1
          return createStringCellSnapshot(cleanRemoteTextAddress, 'clean remote text')
        }
        return address === 'B2' ? createStyledStringCellSnapshot('B2', '', 'style-green') : createEmptyCellSnapshot(address)
      },
      getCellStyle: (styleId) => (styleId === 'style-green' ? { id: 'style-green', fill: { backgroundColor: '#00ff00' } } : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 0,
        localRevision,
        projectedRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
    }
    const cleanRemoteTile = materializeGridRenderTileV3({
      axisSeqX: 1,
      axisSeqY: 1,
      cameraSeq: 1,
      columnWidths: {},
      dprBucket: 1,
      engine,
      freezeSeq: 1,
      gridMetrics: getGridMetrics(),
      materializedAtSeq: 1,
      packetSeq: 1,
      rectSeq: 1,
      rowHeights: {},
      sheetId: 7,
      sheetName: 'Sheet1',
      sheetOrdinal: 7,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      styleSeq: 1,
      textSeq: 1,
      valueSeq: 1,
      viewport: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
    })
    expect(cleanRemoteTile.tileId).toBe(cleanTileId)
    const renderTileSource = createRenderTileSource([dirtyRemoteTile, cleanRemoteTile])
    const viewport = { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 }

    runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: viewport,
        residentViewport: viewport,
        visibleViewport: viewport,
      }),
    )
    cleanTileReadCount = 0
    localRevision = 1
    projectedRevision = 2
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([1, 1, 1, 1, DirtyMaskV3.Style | DirtyMaskV3.Rect]),
        },
        seq: 2,
        source: 'localOptimistic',
      }),
      { dprBucket: 1 },
    )

    const state = runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: viewport,
        residentViewport: viewport,
        visibleViewport: viewport,
      }),
    )

    expect(state.renderTilePanes[0]?.tile).not.toBe(dirtyRemoteTile)
    expect(state.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([1, 1]))
    expect(state.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([1, 1]))
    expect(state.renderTilePanes[1]?.tile).toBe(cleanRemoteTile)
    expect(cleanTileReadCount).toBe(0)
  })

  it('preserves dirty spans from stale projected remote tiles when localizing them', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key for the test viewport')
    }
    const dirtyMask = DirtyMaskV3.Value | DirtyMaskV3.Text
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      dirtyLocalCols: new Uint32Array([5, 5]),
      dirtyLocalRows: new Uint32Array([6, 6]),
      dirtyMasks: new Uint32Array([dirtyMask]),
      lastBatchId: 1,
    }
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 0,
        localRevision: 0,
        projectedRevision: 2,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
    }

    const state = runtime.resolve(
      createInput({
        engine,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([staleRemoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(state.residentBodyPane?.tile).not.toBe(staleRemoteTile)
    expect(state.residentBodyPane?.tile.rectInstances).toBe(staleRemoteTile.rectInstances)
    expect(state.residentBodyPane?.tile.dirtyLocalRows).toEqual(new Uint32Array([6, 6]))
    expect(state.residentBodyPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([5, 5]))
    expect(state.residentBodyPane?.tile.dirtyMasks).toEqual(new Uint32Array([dirtyMask]))
    expect(state.residentBodyPane?.tile.dirty.rectSpans).toEqual([])
  })

  it('requires coherent sheet id and ordinal for remote tiles when both are known', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]

    const sameOrdinalWrongSheetId = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId, 99, 7)]),
        sheetId: 7,
        sheetOrdinal: 7,
      }),
    )
    const sameSheetIdWrongOrdinal = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId, 7, 2)]),
        sheetId: 7,
        sheetOrdinal: 7,
      }),
    )

    expect(sameOrdinalWrongSheetId.residentDataPanes).toHaveLength(1)
    expect(sameOrdinalWrongSheetId.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 7, sheetOrdinal: 7 })
    expect(sameOrdinalWrongSheetId.residentDataPanes[0]?.tile.rectCount).toBeGreaterThan(0)
    expect(sameSheetIdWrongOrdinal.residentDataPanes).toHaveLength(1)
    expect(sameSheetIdWrongOrdinal.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 7, sheetOrdinal: 7 })
    expect(sameSheetIdWrongOrdinal.residentDataPanes[0]?.tile.rectCount).toBeGreaterThan(0)
  })

  it('does not retain local fallback panes as authoritative remote panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const first = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        renderTileSource: createRenderTileSource([]),
        sceneRevision: 1,
      }),
    )
    const second = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        renderTileSource: createRenderTileSource([]),
        sceneRevision: 2,
      }),
    )

    expect(first.residentDataPanes).toHaveLength(1)
    expect(second.residentDataPanes).toHaveLength(1)
    expect(second.residentDataPanes).not.toBe(first.residentDataPanes)
  })

  it('requests local cell invalidation only when local tiles are the active source', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        renderTileSource: undefined,
      }),
    )

    expect(state.residentBodyPane?.tile.coord.sheetId).toBe(7)
    expect(state.needsLocalCellInvalidation).toBe(true)
    expect(state.residentDataPanes).toHaveLength(1)
  })

  it('owns local cell invalidation and clears retained remote panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    let invalidationListener: (() => void) | null = null
    let mergeInvalidationListener: (() => void) | null = null
    let subscribedSheetName = ''
    let subscribedAddresses: readonly string[] = []
    let unsubscribed = false
    let mergeUnsubscribed = false
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      subscribeCells: (sheetName, addresses, listener) => {
        subscribedSheetName = sheetName
        subscribedAddresses = addresses
        invalidationListener = listener
        return () => {
          unsubscribed = true
        }
      },
      subscribeSheetChannel: (_sheetName, channel, listener) => {
        expect(channel).toBe('merges')
        mergeInvalidationListener = listener
        return () => {
          mergeUnsubscribed = true
        }
      },
    }
    const ready = runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([createRenderTile(tileId)]),
      }),
    )
    const invalidations: string[] = []
    const unsubscribe = runtime.connectLocalCellInvalidation(
      {
        engine,
        localInvalidationAddresses: ['A1', 'B2'],
        needsLocalCellInvalidation: true,
        sheetName: 'Sheet1',
      },
      () => invalidations.push('invalidated'),
    )

    expect(subscribedSheetName).toBe('Sheet1')
    expect(subscribedAddresses).toEqual(['A1', 'B2'])
    expect(
      runtime.resolve(
        createInput({
          gridRuntimeHost: host,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentBodyPane?.tile,
    ).toBe(ready.residentBodyPane?.tile)

    invalidationListener?.()
    mergeInvalidationListener?.()

    expect(invalidations).toEqual(['invalidated', 'invalidated'])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 2,
      renderTileRevision: 0,
    })
    expect(
      runtime.resolve(
        createInput({
          engine: LOCAL_EMPTY_ENGINE,
          gridRuntimeHost: host,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentDataPanes,
    ).not.toBe(ready.residentDataPanes)
    unsubscribe?.()
    expect(unsubscribed).toBe(true)
    expect(mergeUnsubscribed).toBe(true)
  })

  it('owns render tile delta subscription stamping in the runtime', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const renderTileSource = createCapturingRenderTileSource()

    const unsubscribe = runtime.connectRenderTileDeltas(
      {
        dprBucket: 2,
        gridRuntimeHost,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 63, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      },
      () => {},
    )

    expect(renderTileSource.captured()).toMatchObject({
      cameraSeq: gridRuntimeHost.snapshot().camera.seq,
      colEnd: 255,
      colStart: 0,
      dprBucket: 2,
      initialDelta: 'full',
      rowEnd: 63,
      rowStart: 0,
      sheetId: 7,
      sheetName: 'Sheet1',
    })
    expect(renderTileSource.captured()?.tileInterest).toMatchObject({
      axisSeqX: gridRuntimeHost.snapshot().axisSeqX,
      axisSeqY: gridRuntimeHost.snapshot().axisSeqY,
      freezeSeq: gridRuntimeHost.snapshot().freezeSeq,
      reason: 'scroll',
      sheetOrdinal: 7,
    })
    expect(renderTileSource.captured()?.tileInterest?.visibleTileKeys).toEqual([
      packTileKey53({
        colTile: 0,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    ])
    expect(renderTileSource.captured()?.warmTileKeys).toContain(
      packTileKey53({
        colTile: 1,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    expect(renderTileSource.captured()?.tileInterest?.warmTileKeys).toContain(
      packTileKey53({
        colTile: 1,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    expect(renderTileSource.captured()?.warmTileKeys).not.toContain(
      packTileKey53({
        colTile: 0,
        dprBucket: 2,
        rowTile: 0,
        sheetOrdinal: 7,
      }),
    )
    unsubscribe?.()
    expect(renderTileSource.unsubscribed()).toBe(true)
  })

  it('builds frozen-pane tile interest from disjoint body and frozen strips instead of the origin-to-body rectangle', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const renderTileSource = createCapturingRenderTileSource()
    const sheetOrdinal = 7
    const dprBucket = 1

    runtime.connectRenderTileDeltas(
      {
        dprBucket,
        freezeCols: 1,
        freezeRows: 1,
        gridRuntimeHost,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 767, colStart: 0, rowEnd: 50079, rowStart: 0 },
        residentViewport: { colEnd: 767, colStart: 512, rowEnd: 50079, rowStart: 49984 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => {},
    )

    const visibleTileKeys = renderTileSource.captured()?.tileInterest?.visibleTileKeys ?? []

    expect(visibleTileKeys).toHaveLength(12)
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 1562, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 0, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 1562, sheetOrdinal }))
    expect(visibleTileKeys).toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 0, sheetOrdinal }))
    expect(visibleTileKeys).not.toContain(packTileKey53({ colTile: 4, dprBucket, rowTile: 100, sheetOrdinal }))
    expect(visibleTileKeys).not.toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 100, sheetOrdinal }))
  })

  it('materializes frozen deep-scroll panes from disjoint resident strips instead of the origin rectangle', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const gridRuntimeHost = createHost()
    const sheetOrdinal = 7
    const dprBucket = 1
    let getCellCount = 0
    const countingEngine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) => {
        getCellCount += 1
        return createEmptyCellSnapshot(address)
      },
    }

    const state = runtime.resolve(
      createInput({
        dprBucket,
        engine: countingEngine,
        freezeCols: 1,
        freezeRows: 1,
        frozenColumnWidth: 100,
        frozenRowHeight: 20,
        gridRuntimeHost,
        renderTileSource: createRenderTileSource([]),
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 959, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 128, rowEnd: 959, rowStart: 928 },
        visibleViewport: { colEnd: 255, colStart: 128, rowEnd: 959, rowStart: 928 },
      }),
    )

    const resolvedTileIds = new Set(state.renderTilePanes.map((pane) => pane.tile.tileId))

    expect(resolvedTileIds).toEqual(
      new Set([
        packTileKey53({ colTile: 1, dprBucket, rowTile: 29, sheetOrdinal }),
        packTileKey53({ colTile: 1, dprBucket, rowTile: 0, sheetOrdinal }),
        packTileKey53({ colTile: 0, dprBucket, rowTile: 29, sheetOrdinal }),
        packTileKey53({ colTile: 0, dprBucket, rowTile: 0, sheetOrdinal }),
      ]),
    )
    expect(state.renderTilePanes).toHaveLength(4)
    expect(getCellCount).toBeLessThanOrEqual(50_000)
    expect(resolvedTileIds).not.toContain(packTileKey53({ colTile: 0, dprBucket, rowTile: 10, sheetOrdinal }))
    expect(resolvedTileIds).not.toContain(packTileKey53({ colTile: 1, dprBucket, rowTile: 10, sheetOrdinal }))
  })
})
