import type { CellStyleRecord, GridEngineLike, GridRenderTile } from './gridRenderTilePaneRuntime-test-helpers.js'
import {
  DEFAULT_TEST_FONT,
  DEFAULT_TEST_FONT_SIZE,
  DEFAULT_TEST_TEXT_COLOR,
  DirtyMaskV3,
  GridRenderTilePaneRuntime,
  LOCAL_EMPTY_ENGINE,
  createEmptyCellSnapshot,
  createGridBorderRectInstances,
  createHost,
  createInput,
  createRenderTile,
  createRenderTileSource,
  createStringCellSnapshot,
  createStyledStringCellSnapshot,
  createWorkbookDeltaBatch,
  createWorkbookDeltaSource,
  describe,
  expect,
  expectedGridBorderRectCount,
  hasOpaqueGreenFillRect,
  it,
} from './gridRenderTilePaneRuntime-test-helpers.js'

describe('GridRenderTilePaneRuntime workbook delta rechecks', () => {
  it('rechecks visible remote tile text when a local workbook delta dirties the tile', () => {
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
    let localRevision = 0
    let visibleText = ''
    let getCellCallCount = 0
    const engineWithLocalVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) => {
        getCellCallCount += 1
        return address === 'B25' && visibleText ? createStringCellSnapshot('B25', visibleText) : createEmptyCellSnapshot(address)
      },
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 1,
        localRevision,
        projectedRevision: 1,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )
    const callsAfterInitialResolve = getCellCallCount
    expect(initial.residentBodyPane?.tile).toBe(remoteTile)

    runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )
    expect(getCellCallCount).toBe(callsAfterInitialResolve)

    localRevision = 1
    visibleText = 'abcdef'
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([24, 24, 1, 1, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        },
        seq: 2,
        source: 'localOptimistic',
      }),
      { dprBucket: 1 },
    )

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithLocalVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 24 && run.col === 1 && run.text === visibleText)).toBe(true)
  })

  it('rebuilds visible remote tiles when style revisions are stale even if text still matches', () => {
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
    const remoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      lastBatchId: 7,
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 0,
          clipY: 0,
          color: DEFAULT_TEST_TEXT_COLOR,
          col: 0,
          font: DEFAULT_TEST_FONT,
          fontSize: DEFAULT_TEST_FONT_SIZE,
          height: 20,
          row: 0,
          strike: false,
          text: 'A1',
          underline: false,
          width: 100,
          x: 0,
          y: 0,
        },
      ],
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 1,
        text: 7,
        values: 7,
      },
    }
    const styles: Record<string, CellStyleRecord> = {
      'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
    }
    const engineWithFreshFill: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A1' ? createStyledStringCellSnapshot('A1', 'A1', 'style-green') : createEmptyCellSnapshot(address),
      getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 7,
        localRevision: 7,
        projectedRevision: 7,
        tileSceneCameraSeq: 7,
        tileSceneRevision: 7,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithFreshFill,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(hasOpaqueGreenFillRect(refreshed.residentBodyPane?.tile)).toBe(true)
  })

  it('rebuilds current-revision remote tiles that are missing visible fill rects', () => {
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
    const remoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      lastBatchId: 7,
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 100,
          clipX: 0,
          clipY: 0,
          color: DEFAULT_TEST_TEXT_COLOR,
          col: 0,
          font: DEFAULT_TEST_FONT,
          fontSize: DEFAULT_TEST_FONT_SIZE,
          height: 20,
          row: 0,
          strike: false,
          text: 'A1',
          underline: false,
          width: 100,
          x: 0,
          y: 0,
        },
      ],
      version: {
        axisX: 7,
        axisY: 7,
        freeze: 7,
        styles: 7,
        text: 7,
        values: 7,
      },
    }
    const styles: Record<string, CellStyleRecord> = {
      'style-green': { id: 'style-green', fill: { backgroundColor: '#00ff00' } },
    }
    const engineWithFreshFill: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A1' ? createStyledStringCellSnapshot('A1', 'A1', 'style-green') : createEmptyCellSnapshot(address),
      getCellStyle: (styleId) => (styleId ? styles[styleId] : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 7,
        localRevision: 7,
        projectedRevision: 7,
        tileSceneCameraSeq: 7,
        tileSceneRevision: 7,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithFreshFill,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(hasOpaqueGreenFillRect(refreshed.residentBodyPane?.tile)).toBe(true)
  })

  it('does not localize offscreen resident text when the logical visible window is shorter', () => {
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
    const engineWithResidentText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'B25' ? createStringCellSnapshot('B25', 'click-away text') : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 1,
        localRevision: 1,
        projectedRevision: 1,
        tileSceneCameraSeq: 1,
        tileSceneRevision: 1,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTile])

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithResidentText,
        gridRuntimeHost: host,
        renderTileSource,
        residentViewport: { colEnd: 10, colStart: 0, rowEnd: 30, rowStart: 0 },
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).toBe(remoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 24 && run.col === 1 && run.text === 'click-away text')).toBe(
      false,
    )
  })

  it('reuses remote static rect buffers for text-only local dirty tiles', () => {
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
    const baseTile = createRenderTile(tileId)
    const rectCount = expectedGridBorderRectCount(baseTile.bounds)
    const rectInstances = createGridBorderRectInstances(rectCount)
    rectInstances[0] = 42
    const remoteTile: GridRenderTile = {
      ...baseTile,
      rectCount,
      rectInstances,
    }

    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([6, 6, 5, 5, DirtyMaskV3.Value | DirtyMaskV3.Text]),
        },
      }),
      { dprBucket: 1 },
    )
    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.rectCount).toBe(rectCount)
    expect(fallback.renderTilePanes[0]?.tile.rectInstances).toBe(rectInstances)
    expect(fallback.renderTilePanes[0]?.tile.dirty?.rectSpans).toEqual([])
    expect(fallback.renderTilePanes[0]?.tile.dirty?.textSpans).toEqual([])
  })

  it('reuses resident static rect buffers for text-only local dirty tiles when source misses', () => {
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
    const baseTile = createRenderTile(tileId)
    const rectCount = expectedGridBorderRectCount(baseTile.bounds)
    const rectInstances = createGridBorderRectInstances(rectCount)
    rectInstances[0] = 77
    const residentTile: GridRenderTile = {
      ...baseTile,
      rectCount,
      rectInstances,
    }

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([residentTile]),
      }),
    )
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([6, 6, 5, 5, DirtyMaskV3.Value | DirtyMaskV3.Text]),
        },
      }),
      { dprBucket: 1 },
    )
    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.rectCount).toBe(rectCount)
    expect(fallback.renderTilePanes[0]?.tile.rectInstances).toBe(rectInstances)
    expect(fallback.renderTilePanes[0]?.tile.dirty?.rectSpans).toEqual([])
  })

  it('applies local optimistic workbook deltas even after higher authoritative seqs on the same sheet', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )

    renderTileSource.emit({
      ...createWorkbookDeltaBatch({ seq: 10 }),
      source: 'workerAuthoritative',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })

    renderTileSource.emit({
      ...createWorkbookDeltaBatch({ seq: 1 }),
      source: 'localOptimistic',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
  })

  it('uses local fallback for local optimistic axis damage while waiting for worker render tiles', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )

    renderTileSource.emit({
      dirty: {
        axisX: new Uint32Array([1, 1, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        axisY: new Uint32Array(),
        cellRanges: new Uint32Array(),
      },
      seq: 1,
      sheetId: 7,
      sheetOrdinal: 7,
      source: 'localOptimistic',
    })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })
  })

  it('does not retain remote panes across sheet switches or before the host is ready', () => {
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

    const switched = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
        sheetId: 8,
      }),
    )

    expect(switched.residentDataPanes).not.toBe(ready.residentDataPanes)
    expect(switched.residentDataPanes).toHaveLength(1)
    expect(switched.residentDataPanes[0]?.tile.coord).toMatchObject({ sheetId: 8, sheetOrdinal: 8 })
    expect(
      runtime.resolve(
        createInput({
          gridRuntimeHost: host,
          hostReady: false,
          renderTileSource: createRenderTileSource([]),
        }),
      ).residentDataPanes,
    ).toEqual([])
  })
})
