import type { GridEngineLike, GridRenderTile, WorkbookDeltaBatchLikeV3 } from './gridRenderTilePaneRuntime-test-helpers.js'
import {
  DEFAULT_TEST_FONT,
  DEFAULT_TEST_FONT_SIZE,
  DEFAULT_TEST_TEXT_COLOR,
  DirtyMaskV3,
  GridRenderTilePaneRuntime,
  LOCAL_EMPTY_ENGINE,
  createCapturingRenderTileSource,
  createEmptyCellSnapshot,
  createHost,
  createInput,
  createMutableRenderTileSource,
  createMutableWorkbookDeltaRenderTileSource,
  createRenderTile,
  createRenderTileSource,
  createStringCellSnapshot,
  createWorkbookDeltaBatch,
  createWorkbookDeltaSource,
  describe,
  expect,
  formatAddress,
  hasOpaqueGreenFillRect,
  it,
  packTileKey53,
} from './gridRenderTilePaneRuntime-test-helpers.js'

describe('GridRenderTilePaneRuntime coordinator and tile delta propagation', () => {
  it('applies render tile delta changes to the host-owned coordinator before React recomputes panes', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const renderTileSource = createMutableRenderTileSource([
      {
        ...createRenderTile(tileId),
        version: {
          axisX: 1,
          axisY: 1,
          freeze: 0,
          styles: 1,
          text: 1,
          values: 1,
        },
      },
    ])

    runtime.resolve(
      createInput({
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    renderTileSource.setTile({
      ...createRenderTile(tileId),
      version: {
        axisX: 2,
        axisY: 3,
        freeze: 4,
        styles: 5,
        text: 6,
        values: 7,
      },
    })
    const listenerChanges: unknown[] = []
    const unsubscribe = runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      (change) => listenerChanges.push(change),
    )

    renderTileSource.emit({ changedTileIds: [tileId] })

    expect(host.tiles.residency.getExact(tileId)).toMatchObject({
      axisSeqX: 2,
      axisSeqY: 3,
      freezeSeq: 4,
      styleSeq: 5,
      textSeq: 6,
      valueSeq: 7,
    })
    expect(listenerChanges).toHaveLength(1)
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })

    renderTileSource.emit({ invalidatedTileIds: [tileId] })
    expect(host.tiles.residency.getExact(tileId)).toBeNull()
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 2,
    })

    unsubscribe?.()
    expect(renderTileSource.unsubscribed()).toBe(true)
  })

  it('skips render tile delta subscription until a remote source and sheet id exist', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const renderTileSource = createCapturingRenderTileSource()
    const input = {
      dprBucket: 1,
      gridRuntimeHost: createHost(),
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetName: 'Sheet1',
    }

    expect(
      runtime.connectRenderTileDeltas(
        {
          ...input,
          renderTileSource: undefined,
          sheetId: 7,
        },
        () => {},
      ),
    ).toBeUndefined()
    expect(
      runtime.connectRenderTileDeltas(
        {
          ...input,
          renderTileSource: renderTileSource.source,
          sheetId: undefined,
        },
        () => {},
      ),
    ).toBeUndefined()
    expect(renderTileSource.captured()).toBeNull()
  })

  it('applies workbook delta damage to the host-owned tile coordinator', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()
    const listenerBatches: WorkbookDeltaBatchLikeV3[] = []
    const tileId = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 7,
    })

    const unsubscribe = runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      (batch) => listenerBatches.push(batch),
    )

    renderTileSource.emit(createWorkbookDeltaBatch())
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 1 }))
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 2, sheetId: 8, sheetOrdinal: 8 }))

    expect(listenerBatches.map((batch) => batch.seq)).toEqual([1])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })
    expect(
      host.tiles.reconcileInterest({
        axisSeqX: 1,
        axisSeqY: 1,
        cameraSeq: 1,
        freezeSeq: 1,
        pinnedTileKeys: [],
        reason: 'mutation',
        seq: 1,
        sheetId: 7,
        sheetOrdinal: 7,
        visibleTileKeys: [tileId],
        warmTileKeys: [],
      }).visibleDirtyTileKeys,
    ).toEqual([tileId])

    unsubscribe?.()
    expect(renderTileSource.unsubscribed()).toBe(true)
  })

  it('reconciles render tile connection lifecycles without React-owned resubscribe churn', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createCapturingRenderTileSource()
    let subscribedSheetName = ''
    let subscribedAddresses: readonly string[] = []
    let localUnsubscribeCount = 0
    const engine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      subscribeCells: (sheetName, addresses) => {
        subscribedSheetName = sheetName
        subscribedAddresses = addresses
        return () => {
          localUnsubscribeCount += 1
        }
      },
    }
    const localInvalidationAddresses = ['A1', 'B2']

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      localInvalidationAddresses,
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
    })
    const firstSubscription = renderTileSource.captured()

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      localInvalidationAddresses: [...localInvalidationAddresses],
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
    })

    expect(renderTileSource.captured()).toBe(firstSubscription)
    expect(renderTileSource.unsubscribed()).toBe(false)
    expect(renderTileSource.subscribeCount()).toBe(1)
    expect(subscribedSheetName).toBe('Sheet1')
    expect(subscribedAddresses).toBe(localInvalidationAddresses)
    expect(localUnsubscribeCount).toBe(0)

    runtime.syncConnections({
      dprBucket: 1,
      engine,
      gridRuntimeHost: host,
      localInvalidationAddresses,
      needsLocalCellInvalidation: true,
      renderTileSource: renderTileSource.source,
      renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
      sheetId: 7,
      sheetName: 'Sheet1',
    })

    expect(renderTileSource.unsubscribed()).toBe(true)
    expect(renderTileSource.subscribeCount()).toBe(2)
    expect(renderTileSource.unsubscribeCount()).toBe(1)
    expect(renderTileSource.captured()).not.toBe(firstSubscription)
    expect(localUnsubscribeCount).toBe(0)

    runtime.disconnectConnections()

    expect(renderTileSource.unsubscribeCount()).toBe(2)
    expect(localUnsubscribeCount).toBe(1)
  })

  it('reconciles render tile interest when only the visible viewport changes inside a resident viewport', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createCapturingRenderTileSource()
    const renderTileViewport = { colEnd: 127, colStart: 0, rowEnd: 95, rowStart: 0 }
    const firstVisibleViewport = { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 }
    const nextVisibleViewport = { colEnd: 127, colStart: 0, rowEnd: 95, rowStart: 64 }

    runtime.syncConnections({
      dprBucket: 1,
      engine: LOCAL_EMPTY_ENGINE,
      gridRuntimeHost: host,
      localInvalidationAddresses: [],
      needsLocalCellInvalidation: false,
      renderTileSource: renderTileSource.source,
      renderTileViewport,
      residentViewport: renderTileViewport,
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleViewport: firstVisibleViewport,
    })
    const firstSubscription = renderTileSource.captured()

    expect(firstSubscription?.tileInterest?.visibleTileKeys).toEqual([
      packTileKey53({ colTile: 0, dprBucket: 1, rowTile: 0, sheetOrdinal: 7 }),
    ])

    runtime.syncConnections({
      dprBucket: 1,
      engine: LOCAL_EMPTY_ENGINE,
      gridRuntimeHost: host,
      localInvalidationAddresses: [],
      needsLocalCellInvalidation: false,
      renderTileSource: renderTileSource.source,
      renderTileViewport,
      residentViewport: renderTileViewport,
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleViewport: nextVisibleViewport,
    })

    expect(renderTileSource.subscribeCount()).toBe(2)
    expect(renderTileSource.unsubscribeCount()).toBe(1)
    expect(renderTileSource.captured()).not.toBe(firstSubscription)
    expect(renderTileSource.captured()?.tileInterest?.visibleTileKeys).toEqual([
      packTileKey53({ colTile: 0, dprBucket: 1, rowTile: 2, sheetOrdinal: 7 }),
    ])

    runtime.syncConnections({
      dprBucket: 1,
      engine: LOCAL_EMPTY_ENGINE,
      gridRuntimeHost: host,
      localInvalidationAddresses: [],
      needsLocalCellInvalidation: false,
      renderTileSource: renderTileSource.source,
      renderTileViewport,
      residentViewport: renderTileViewport,
      sheetId: 7,
      sheetName: 'Sheet1',
      visibleViewport: { ...nextVisibleViewport },
    })

    expect(renderTileSource.subscribeCount()).toBe(2)
    expect(renderTileSource.unsubscribeCount()).toBe(1)
  })

  it('matches workbook delta damage by sheet ordinal when sheet id differs from order', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileSource = createWorkbookDeltaSource()
    const tileId = packTileKey53({
      colTile: 0,
      dprBucket: 1,
      rowTile: 0,
      sheetOrdinal: 2,
    })
    const listenerBatches: WorkbookDeltaBatchLikeV3[] = []

    const unsubscribe = runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 99,
        sheetOrdinal: 2,
      },
      (batch) => listenerBatches.push(batch),
    )

    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 1, sheetId: 99, sheetOrdinal: 2 }))
    renderTileSource.emit(createWorkbookDeltaBatch({ seq: 2, sheetId: 7, sheetOrdinal: 7 }))

    expect(listenerBatches.map((batch) => batch.seq)).toEqual([1])
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })
    expect(
      host.tiles.reconcileInterest({
        axisSeqX: 1,
        axisSeqY: 1,
        cameraSeq: 1,
        freezeSeq: 1,
        pinnedTileKeys: [],
        reason: 'mutation',
        seq: 1,
        sheetId: 99,
        sheetOrdinal: 2,
        visibleTileKeys: [tileId],
        warmTileKeys: [],
      }).visibleDirtyTileKeys,
    ).toEqual([tileId])

    unsubscribe?.()
  })

  it('does not render stale local tiles for worker-authoritative damage before fresh render tiles arrive', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    const remoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          col: 0,
          row: 0,
          text: 'stale remote text',
          x: 0,
          y: 0,
          width: 80,
          height: 20,
          clipX: 0,
          clipY: 0,
          clipWidth: 80,
          clipHeight: 20,
          font: '12px sans-serif',
          fontSize: 12,
          color: '#000000',
          underline: false,
          strike: false,
        },
      ],
    }
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([remoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(initial.residentBodyPane?.tile.textRuns.some((run) => run.text === 'stale remote text')).toBe(false)

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )
    runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => undefined,
    )
    renderTileSource.emitWorkbookDelta({
      ...createWorkbookDeltaBatch(),
      source: 'workerAuthoritative',
    })

    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 1,
    })
    expect(fallback.residentBodyPane?.tile.textRuns.some((run) => run.text === 'stale remote text')).toBe(false)
    expect(fallback.residentBodyPane?.tile.textRuns).toEqual([])

    const freshRemoteTile: GridRenderTile = {
      ...remoteTile,
      lastBatchId: 2,
      lastCameraSeq: 2,
      textRuns: [],
      textCount: 0,
      version: {
        ...remoteTile.version,
        text: remoteTile.version.text + 1,
        values: remoteTile.version.values + 1,
      },
    }
    renderTileSource.setTile(freshRemoteTile)
    renderTileSource.emitRenderTileDelta({ changedTileIds: [tileId] })

    const refreshed = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 0,
      renderTileRevision: 2,
    })
    expect(refreshed.residentBodyPane?.tile.textRuns).toEqual([])
  })

  it('rebuilds resident local fallback tiles when the projected workbook revision advances', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    let projectedRevision = 0
    let greenFillVisible = false
    const engineWithChangingStyle: GridEngineLike = {
      getCell: (_sheetName, address) => ({
        ...(address === 'E6' && greenFillVisible ? { styleId: 'style-green' } : {}),
        ...createEmptyCellSnapshot(address),
      }),
      getCellStyle: (styleId) => (styleId === 'style-green' ? { id: 'style-green', fill: { backgroundColor: '#00ff00' } } : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: projectedRevision,
        projectedRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithChangingStyle,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 0,
      }),
    )
    expect(hasOpaqueGreenFillRect(initial.residentBodyPane?.tile)).toBe(false)

    projectedRevision = 1
    greenFillVisible = true

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithChangingStyle,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 0,
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(initial.residentBodyPane?.tile)
    expect(hasOpaqueGreenFillRect(refreshed.residentBodyPane?.tile)).toBe(true)
  })

  it('keeps worker-authoritative dirty tiles pending until projected state catches up', () => {
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
    let projectedRevision = 1
    let greenFillVisible = false
    const engineWithDelayedProjection: GridEngineLike = {
      getCell: (_sheetName, address) => ({
        ...(address === 'B2' && greenFillVisible ? { styleId: 'style-green' } : {}),
        ...createEmptyCellSnapshot(address),
      }),
      getCellStyle: (styleId) => (styleId === 'style-green' ? { id: 'style-green', fill: { backgroundColor: '#00ff00' } } : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 2,
        projectedRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([createRenderTile(tileId)])

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )
    renderTileSource.emitWorkbookDelta({
      ...createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([1, 1, 1, 1, DirtyMaskV3.Value | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
        },
        seq: 2,
      }),
      source: 'workerAuthoritative',
    })

    const staleProjection = runtime.resolve(
      createInput({
        engine: engineWithDelayedProjection,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sceneRevision: 0,
      }),
    )
    expect(staleProjection.tileReadiness.visibleDirtyTileKeys).toContain(tileId)
    expect(hasOpaqueGreenFillRect(staleProjection.residentBodyPane?.tile)).toBe(false)

    const stillPending = runtime.resolve(
      createInput({
        engine: engineWithDelayedProjection,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sceneRevision: 0,
      }),
    )
    expect(stillPending.tileReadiness.visibleDirtyTileKeys).toContain(tileId)

    projectedRevision = 2
    greenFillVisible = true

    const caughtUp = runtime.resolve(
      createInput({
        engine: engineWithDelayedProjection,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sceneRevision: 0,
      }),
    )
    expect(caughtUp.tileReadiness.visibleDirtyTileKeys).toContain(tileId)
    expect(hasOpaqueGreenFillRect(caughtUp.residentBodyPane?.tile)).toBe(true)

    const acknowledged = runtime.resolve(
      createInput({
        engine: engineWithDelayedProjection,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sceneRevision: 0,
      }),
    )
    expect(acknowledged.tileReadiness.visibleDirtyTileKeys).toEqual([])
    expect(hasOpaqueGreenFillRect(acknowledged.residentBodyPane?.tile)).toBe(true)
  })

  it('rebuilds dirty resident tiles for pending local projections even when the remote batch id is ahead', () => {
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
      lastBatchId: 99,
      version: {
        axisX: 99,
        axisY: 99,
        freeze: 99,
        styles: 99,
        text: 99,
        values: 99,
      },
    }
    const engineWithPendingLocalFill: GridEngineLike = {
      getCell: (_sheetName, address) => ({
        ...(address === 'E6' ? { styleId: 'style-green' } : {}),
        ...createEmptyCellSnapshot(address),
      }),
      getCellStyle: (styleId) => (styleId === 'style-green' ? { id: 'style-green', fill: { backgroundColor: '#00ff00' } } : undefined),
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision: 0,
        localRevision: 24,
        projectedRevision: 2,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array([5, 5, 4, 4, DirtyMaskV3.Style | DirtyMaskV3.Rect]),
        },
        seq: 2,
        source: 'localOptimistic',
      }),
      { dprBucket: 1 },
    )

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithPendingLocalFill,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTile]),
        visibleViewport: { colEnd: 0, colStart: 0, rowEnd: 0, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(remoteTile)
    expect(hasOpaqueGreenFillRect(refreshed.residentBodyPane?.tile)).toBe(true)
  })

  it('keeps local tiles through stale render tile deltas until the renderer batch catches up', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const tileId = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    })[0]
    if (tileId === undefined) {
      throw new Error('Expected a visible render tile key')
    }
    const staleRemoteTile: GridRenderTile = {
      ...createRenderTile(tileId),
      lastBatchId: 9,
      lastCameraSeq: 9,
      textCount: 1,
      textRuns: [
        {
          col: 0,
          row: 0,
          text: 'deleted text from stale tile',
          x: 0,
          y: 0,
          width: 140,
          height: 20,
          clipX: 0,
          clipY: 0,
          clipWidth: 140,
          clipHeight: 20,
          font: '12px sans-serif',
          fontSize: 12,
          color: '#000000',
          underline: false,
          strike: false,
        },
      ],
      version: {
        axisX: 9,
        axisY: 9,
        freeze: 0,
        styles: 9,
        text: 9,
        values: 9,
      },
    }
    const renderTileSource = createMutableWorkbookDeltaRenderTileSource([staleRemoteTile])

    runtime.connectWorkbookDeltaDamage(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        sheetId: 7,
      },
      () => undefined,
    )
    runtime.connectRenderTileDeltas(
      {
        dprBucket: 1,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        sheetId: 7,
        sheetName: 'Sheet1',
      },
      () => undefined,
    )

    renderTileSource.emitWorkbookDelta({
      ...createWorkbookDeltaBatch({ seq: 10 }),
      source: 'localOptimistic',
    })
    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 1,
    })

    renderTileSource.setTile(staleRemoteTile)
    renderTileSource.emitRenderTileDelta({ batchId: 9, cameraSeq: 9, changedTileIds: [tileId] })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: true,
      localFallbackRevision: 1,
      renderTileRevision: 2,
    })
    const staleDeltaState = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(staleDeltaState.residentBodyPane?.tile.textRuns.some((run) => run.text === 'deleted text from stale tile')).toBe(false)
    expect(staleDeltaState.residentBodyPane?.tile.textRuns).toEqual([])

    const freshRemoteTile: GridRenderTile = {
      ...staleRemoteTile,
      lastBatchId: 10,
      lastCameraSeq: 10,
      textCount: 0,
      textRuns: [],
      version: {
        ...staleRemoteTile.version,
        text: 10,
        values: 10,
      },
    }
    renderTileSource.setTile(freshRemoteTile)
    renderTileSource.emitRenderTileDelta({ batchId: 10, cameraSeq: 10, changedTileIds: [tileId] })

    expect(runtime.snapshotBridgeState()).toEqual({
      forceLocalTiles: false,
      localFallbackRevision: 1,
      renderTileRevision: 3,
    })
    const caughtUpState = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        forceLocalTiles: runtime.snapshotBridgeState().forceLocalTiles,
        gridRuntimeHost: host,
        renderTileSource: renderTileSource.source,
      }),
    )
    expect(caughtUpState.residentBodyPane?.tile.textRuns).toEqual([])
  })

  it('keeps clean remote tiles resident when local fallback only needs dirty tiles', () => {
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
    const cleanRemoteTile: GridRenderTile = {
      ...createRenderTile(cleanTileId),
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
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'clean remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(dirtyTileId), cleanRemoteTile])
    const cleanRemoteTextAddress = formatAddress(0, 128)
    const cleanRemoteEngine: GridEngineLike = {
      ...LOCAL_EMPTY_ENGINE,
      getCell: (_sheetName, address) =>
        address === cleanRemoteTextAddress
          ? createStringCellSnapshot(cleanRemoteTextAddress, 'clean remote text')
          : createEmptyCellSnapshot(address),
    }

    host.tiles.applyWorkbookDelta(createWorkbookDeltaBatch(), { dprBucket: 1 })
    const fallback = runtime.resolve(
      createInput({
        engine: cleanRemoteEngine,
        forceLocalTiles: true,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    expect(fallback.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual([dirtyTileId, cleanTileId])
    expect(fallback.renderTilePanes[0]?.tile.textRuns).toEqual([])
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([0, 0]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 0]))
    expect(fallback.renderTilePanes[1]?.tile.textRuns[0]?.text).toBe('clean remote text')
  })

  it('rebuilds dirty resident remote tiles before they first become visible', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [visibleTileId, warmTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (visibleTileId === undefined || warmTileId === undefined) {
      throw new Error('Expected visible and warm render tile keys for the test viewport')
    }
    const staleWarmTile: GridRenderTile = {
      ...createRenderTile(warmTileId),
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
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'stale warm remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(visibleTileId), staleWarmTile])
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array([128, 128, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array(),
        },
      }),
      { dprBucket: 1 },
    )

    const offscreen = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    const offscreenWarmPane = offscreen.renderTilePanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(offscreenWarmPane?.tile).not.toBe(staleWarmTile)
    expect(offscreenWarmPane?.tile.textRuns.some((run) => run.text === 'stale warm remote text')).toBe(false)
    expect(offscreenWarmPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 127]))
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).not.toBe(0)

    const visible = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 255, colStart: 128, rowEnd: 31, rowStart: 0 },
      }),
    )

    const visibleWarmPane = visible.renderTilePanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(visibleWarmPane?.tile.textRuns).toEqual([])
    expect(visibleWarmPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 127]))
    expect(visible.tileReadiness.visibleDirtyTileKeys).toContain(warmTileId)
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).toBe(0)
  })

  it('localizes dirty warm preload tiles before stale remote text can be staged', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const [visibleTileId, warmTileId] = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: { colEnd: 255, colStart: 0, rowEnd: 31, rowStart: 0 },
    })
    if (visibleTileId === undefined || warmTileId === undefined) {
      throw new Error('Expected visible and warm render tile keys for the test viewport')
    }
    const staleWarmTile: GridRenderTile = {
      ...createRenderTile(warmTileId),
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
          clipWidth: 80,
          clipX: 0,
          clipY: 0,
          color: '#000000',
          col: 128,
          font: '12px sans-serif',
          fontSize: 12,
          height: 20,
          row: 0,
          strike: false,
          text: 'stale warm remote text',
          underline: false,
          width: 80,
          wrap: false,
          x: 0,
          y: 0,
        },
      ],
    }
    const renderTileSource = createRenderTileSource([createRenderTile(visibleTileId), staleWarmTile])
    host.tiles.applyWorkbookDelta(
      createWorkbookDeltaBatch({
        dirty: {
          axisX: new Uint32Array([128, 128, DirtyMaskV3.AxisX | DirtyMaskV3.Text | DirtyMaskV3.Rect]),
          axisY: new Uint32Array(),
          cellRanges: new Uint32Array(),
        },
      }),
      { dprBucket: 1 },
    )

    const state = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource,
        renderTileViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        residentViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    const warmPreloadPane = state.preloadDataPanes.find((pane) => pane.tile.tileId === warmTileId)
    expect(warmPreloadPane).toBeDefined()
    expect(warmPreloadPane?.tile).not.toBe(staleWarmTile)
    expect(warmPreloadPane?.tile.textRuns.some((run) => run.text === 'stale warm remote text')).toBe(false)
    expect(warmPreloadPane?.tile.dirtyLocalCols).toEqual(new Uint32Array([0, 127]))
    expect(host.tiles.dirtyTiles.getUnconsumedMask(warmTileId)).not.toBe(0)
  })

  it('preserves dirty spans when local fallback has no remote render tile source', () => {
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
        renderTileSource: undefined,
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalRows).toEqual(new Uint32Array([6, 6]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyLocalCols).toEqual(new Uint32Array([5, 5]))
    expect(fallback.renderTilePanes[0]?.tile.dirtyMasks).toEqual(new Uint32Array([DirtyMaskV3.Value | DirtyMaskV3.Text]))
  })

  it('materializes every missing resident tile when the remote source has not caught up', () => {
    const runtime = new GridRenderTilePaneRuntime()
    const host = createHost()
    const renderTileViewport = { colEnd: 127, colStart: 0, rowEnd: 95, rowStart: 0 }
    const expectedTileIds = host.viewportTileKeys({
      dprBucket: 1,
      sheetOrdinal: 7,
      viewport: renderTileViewport,
    })

    const fallback = runtime.resolve(
      createInput({
        engine: LOCAL_EMPTY_ENGINE,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([]),
        renderTileViewport,
        residentViewport: renderTileViewport,
        visibleViewport: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
      }),
    )

    expect(fallback.renderTilePanes.map((pane) => pane.tile.tileId)).toEqual(expectedTileIds)
  })

  it('materializes a visible remote tile locally when its text payload is incomplete', () => {
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
    const remoteTileWithoutText = createRenderTile(tileId)
    const engineWithVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A15' ? createStringCellSnapshot('A15', 'Amortization Schedule Examples') : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }

    const fallback = runtime.resolve(
      createInput({
        engine: engineWithVisibleText,
        gridRuntimeHost: host,
        renderTileSource: createRenderTileSource([remoteTileWithoutText]),
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 34, rowStart: 3 },
      }),
    )

    expect(fallback.renderTilePanes[0]?.tile.tileId).toBe(tileId)
    expect(fallback.renderTilePanes[0]?.tile).not.toBe(remoteTileWithoutText)
    expect(
      fallback.renderTilePanes[0]?.tile.textRuns.some((run) => run.row === 14 && run.col === 0 && run.text.includes('Amortization')),
    ).toBe(true)
  })

  it('caches visible text freshness for clean remote tiles across selection-only resolves', () => {
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
    const remoteTileWithMatchingText: GridRenderTile = {
      ...createRenderTile(tileId),
      textCount: 1,
      textRuns: [
        {
          align: 'left',
          clipHeight: 20,
          clipWidth: 120,
          clipX: 0,
          clipY: 280,
          color: DEFAULT_TEST_TEXT_COLOR,
          col: 0,
          font: DEFAULT_TEST_FONT,
          fontSize: DEFAULT_TEST_FONT_SIZE,
          height: 20,
          row: 14,
          strike: false,
          text: 'Amortization Schedule Examples',
          underline: false,
          width: 120,
          wrap: false,
          x: 0,
          y: 280,
        },
      ],
    }
    let getCellCallCount = 0
    const engineWithMatchingVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) => {
        getCellCallCount += 1
        return address === 'A15' ? createStringCellSnapshot('A15', 'Amortization Schedule Examples') : createEmptyCellSnapshot(address)
      },
      getCellStyle: () => undefined,
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([remoteTileWithMatchingText])
    const visibleViewport = { colEnd: 10, colStart: 0, rowEnd: 34, rowStart: 3 }

    const initial = runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        selectedCell: [1, 1],
        selectedCellSnapshot: createEmptyCellSnapshot('B2'),
        visibleViewport,
      }),
    )
    const callsAfterInitialResolve = getCellCallCount

    expect(initial.residentBodyPane?.tile).toBe(remoteTileWithMatchingText)
    expect(callsAfterInitialResolve).toBeGreaterThan(0)

    const selectionOnly = runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        selectedCell: [2, 2],
        selectedCellSnapshot: createEmptyCellSnapshot('C3'),
        visibleViewport,
      }),
    )

    expect(selectionOnly.residentBodyPane?.tile).toBe(remoteTileWithMatchingText)
    expect(getCellCallCount).toBe(callsAfterInitialResolve)

    runtime.resolve(
      createInput({
        engine: engineWithMatchingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        sceneRevision: 2,
        selectedCell: [3, 3],
        selectedCellSnapshot: createEmptyCellSnapshot('D4'),
        visibleViewport,
      }),
    )

    expect(getCellCallCount).toBeGreaterThan(callsAfterInitialResolve)
  })

  it('rechecks visible remote tile text when the authoritative workbook revision changes', () => {
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
    const staleRemoteTile = createRenderTile(tileId)
    let authoritativeRevision = 1
    let visibleText = ''
    const engineWithChangingVisibleText: GridEngineLike = {
      getCell: (_sheetName, address) =>
        address === 'A1' && visibleText ? createStringCellSnapshot('A1', visibleText) : createEmptyCellSnapshot(address),
      getCellStyle: () => undefined,
      getRenderRevisionSnapshot: () => ({
        authoritativeRevision,
        projectedRevision: authoritativeRevision,
        tileSceneCameraSeq: null,
        tileSceneRevision: null,
      }),
      subscribeCells: () => () => {},
      workbook: {
        getSheet: () => undefined,
      },
    }
    const renderTileSource = createRenderTileSource([staleRemoteTile])

    const initial = runtime.resolve(
      createInput({
        engine: engineWithChangingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )
    expect(initial.residentBodyPane?.tile).toBe(staleRemoteTile)

    authoritativeRevision = 2
    visibleText = 'Prepaid Expense Template'

    const refreshed = runtime.resolve(
      createInput({
        engine: engineWithChangingVisibleText,
        gridRuntimeHost: host,
        renderTileSource,
        visibleViewport: { colEnd: 10, colStart: 0, rowEnd: 10, rowStart: 0 },
      }),
    )

    expect(refreshed.residentBodyPane?.tile).not.toBe(staleRemoteTile)
    expect(refreshed.residentBodyPane?.tile.textRuns.some((run) => run.row === 0 && run.col === 0 && run.text === visibleText)).toBe(true)
  })
})
