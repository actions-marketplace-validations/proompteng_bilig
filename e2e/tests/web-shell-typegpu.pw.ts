import { expect, test } from '@playwright/test'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH } from '../../apps/web/src/root-route.js'
import {
  clickProductCell,
  createTestDocumentId,
  dragProductBodySelection,
  getProductFillHandleDragPoints,
  gotoWorkbookShell,
  pickToolbarPresetColor,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  settleWorkbookScrollPerf,
  stopWorkbookScrollPerf,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
  warmStartWorkbookScrollPerf,
} from './web-shell-helpers.js'
import {
  allReadbackPointsMatch,
  collectRendererPresentationSamplesDuring,
  dragProductSelectionBorder,
  exerciseClickAwayEditCommit,
  expectNear,
  expectPresentedOverlayRects,
  expectSelectionVisualOpacity,
  expectWorkbookWhitePixel,
  inspectGpuReadback,
  installTypeGpuReadbackHarness,
  isCornflowerBlueFill,
  isResizeGuidePixel,
  isSelectionGreenTint,
  isThemeGreenFill,
  isWorkbookWhitePixel,
  productCellInteriorGridLocalPoint,
  productColumnHeaderGridLocalPoint,
  readNativeRectProofState,
  saveReadbackArtifact,
  selectedRangeFillProbe,
  type TypeGpuReadbackSummary,
  waitForCompositedGridLocalPixel,
  waitForNativeTextRunContent,
  waitForReadback,
  waitForReadbackSequence,
  waitForTypeGpuRenderer,
  waitForVisibleNativeTextRuns,
} from './web-shell-typegpu-helpers.js'

test('@browser-webgpu isolated workbook pane renderer draws grid content through typegpu', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 640, height: 480 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, ISOLATED_WORKBOOK_PANE_RENDERER_PATH)
  await page.waitForSelector('[data-testid="isolated-pane-renderer-route"]', { timeout: 15_000 })
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () => Boolean((window as Window & { __biligGpuReadback?: { readonly ready: boolean } }).__biligGpuReadback?.ready),
    undefined,
    { timeout: 15_000 },
  )

  const summary = await page.evaluate(() => {
    return (window as Window & { __biligGpuReadback?: TypeGpuReadbackSummary }).__biligGpuReadback ?? null
  })

  expect(summary).not.toBeNull()
  expect(summary?.hasGpu).toBe(true)
  expect(summary?.width).toBe(640)
  expect(summary?.height).toBe(480)
  expect(summary?.sequence).toBeGreaterThan(0)
  expect(summary?.points.headerFill).toMatchObject({ r: 243, g: 242, b: 238, a: 255 })
  expect(summary?.points.bodyFill).toMatchObject({ r: 255, g: 255, b: 255, a: 255 })
  expect(summary?.points.selectionBorder.a ?? 0).toBeGreaterThan(150)
  expect(summary?.points.selectionBorder.g ?? 0).toBeGreaterThan(summary?.points.selectionBorder.r ?? 0)
  expect(summary?.points.bodyWhite).toMatchObject({ r: 255, g: 255, b: 255, a: 255 })
  await waitForVisibleNativeTextRuns(
    page,
    [
      { name: 'columnHeaderB', text: 'B', exact: true },
      { name: 'bodyRegion', text: 'Region', exact: true },
      { name: 'bodyNorth', text: 'North', exact: true },
    ],
    (runs) => runs.visibleRunCount > 0 && runs.matches.columnHeaderB && runs.matches.bodyRegion && runs.matches.bodyNorth,
  )
  await expect(page.getByTestId('grid-native-text-layer')).toHaveCount(1)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-draw-text', 'false')
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-native-layer-source', 'browser-native-text-live')
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-native-rect-frame-source', 'presented')
  await expect(page.getByTestId('grid-native-rect-layer')).toHaveCount(1)
  await expect
    .poll(readNativeRectProofState(page), {
      message: 'browser-native rect proof must mirror the presented TypeGPU frame, not a stale live/current frame',
      timeout: 5_000,
    })
    .toMatchObject({
      countPresent: true,
      presentedFrameMatches: true,
      sceneEpochMatches: true,
      signatureMatches: true,
      visibleRevisionMatches: true,
    })

  await saveReadbackArtifact(page, testInfo, 'isolated-pane-renderer-readback.png', 'isolated-pane-renderer-readback')
})

test('main workbook shell mounts typegpu-v3 as the only grid renderer', async ({ page }) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)

  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-renderer-mode', 'typegpu-v3')
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-pane-renderer', 'workbook-pane-renderer-v3')
  await expect(page.locator('[data-pane-renderer="workbook-pane-renderer"]')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-text-overlay')).toHaveCount(0)
})

test('@browser-webgpu @browser-serial main workbook shell grid renders and updates through typegpu', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-grid-updates'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )
  await waitForReadbackSequence(page, 0)

  const initialProbe = {
    points: [
      { name: 'unselectedHeaderFill', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + 20, y: 12 },
      {
        name: 'bodyBlank',
        x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 4 + 14,
        y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 + Math.floor(PRODUCT_ROW_HEIGHT / 2),
      },
    ],
    regions: [
      { name: 'columnHeaderText', x0: 176, y0: 4, x1: 228, y1: 18 },
      { name: 'rowHeaderText', x0: 10, y0: 48, x1: 36, y1: 66 },
    ],
  } as const

  const initialReadback = await waitForReadback(page, initialProbe, (result) => {
    return isWorkbookWhitePixel(result.points.bodyBlank) && result.points.unselectedHeaderFill.a === 255
  })
  await waitForVisibleNativeTextRuns(
    page,
    [
      { name: 'columnHeaderText', text: 'B' },
      { name: 'rowHeaderText', text: '2' },
    ],
    (runs) => runs.rowHeaderRunCount > 0 && runs.visibleRunCount > 0,
  )

  expect(initialReadback.hasGpu).toBe(true)
  expect(initialReadback.width).toBeGreaterThan(400)
  expect(initialReadback.height).toBeGreaterThan(250)
  expect(initialReadback.points.bodyBlank).toMatchObject({ r: 255, g: 255, b: 255, a: 255 })
  await expect(page.getByTestId('grid-native-text-layer')).toHaveCount(1)

  await clickProductCell(page, 2, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C4')

  await page.getByTestId('formula-input').fill('123')
  await page.getByTestId('formula-input').press('Enter')
  await expect(page.getByTestId('formula-input')).toHaveValue('123')
  await expect(page.getByTestId('formula-resolved-value')).toHaveText('123')
  await waitForReadbackSequence(page, initialReadback.sequence)

  const valueProbe = {
    points: [],
    regions: [
      {
        name: 'c4ValueText',
        x0: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + 8,
        y0: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 3 + 4,
        x1: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3 - 8,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 - 4,
      },
    ],
  } as const

  const valueReadback = await waitForReadback(page, valueProbe, (result) => result.sequence > initialReadback.sequence)
  await waitForNativeTextRunContent(page, '123')

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2')
  await clickProductCell(page, 2, 2, { shift: true })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:C3')
  await waitForReadbackSequence(page, valueReadback.sequence)

  const rangeProbe = {
    points: [],
    regions: [],
  } as const

  await expectPresentedOverlayRects(page, 'range selection fill')
  await waitForReadback(page, rangeProbe, (result) => result.sequence > valueReadback.sequence)

  const activeCellFillPoint = await productCellInteriorGridLocalPoint(page, 1, 1)
  const selectedRangeFillPoint = await productCellInteriorGridLocalPoint(page, 2, 2)
  const topHeaderSelectionFillPoint = await productColumnHeaderGridLocalPoint(page, 1)
  const activeCellFill = await waitForCompositedGridLocalPixel(page, activeCellFillPoint, isSelectionGreenTint)
  const selectedRangeFill = await waitForCompositedGridLocalPixel(page, selectedRangeFillPoint, isSelectionGreenTint)
  const topHeaderSelectionFill = await waitForCompositedGridLocalPixel(page, topHeaderSelectionFillPoint, isSelectionGreenTint)

  expect(isSelectionGreenTint(activeCellFill)).toBe(true)
  expect(isSelectionGreenTint(selectedRangeFill)).toBe(true)
  expect(isSelectionGreenTint(topHeaderSelectionFill)).toBe(true)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-presented-overlay-rect-signature', /^[a-z0-9-]+$/)
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-role="header-fill"]'),
    'selected range header fill',
    'visible',
  )
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-role="selection-fill"][data-grid-selection-visual-key^="selection-fill:range"]'),
    'selected range body fill',
    'visible',
  )
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-role="selection-gridline"][data-grid-selection-visual-key^="selection-gridline:range"]'),
    'selected range internal gridlines',
    'visible',
  )
  await expectSelectionVisualOpacity(
    page.locator(['[data-grid-selection-visual-role="fill-handle"]', '[data-grid-selection-visual-role="selection-border"]'].join(',')),
    'selected range chrome',
    'visible',
  )
  await expect(page.getByTestId('grid-pane-renderer-floor')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveCSS('opacity', '1')

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-readback.png', 'main-workbook-grid-readback')
})

test('@browser-webgpu @browser-serial main workbook shell keeps row headers visible after click-away edit', async ({ page }, testInfo) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-click-away-edit-row-headers'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )
  await waitForReadbackSequence(page, 0)

  const rowHeaderProbe = {
    points: [],
    regions: [
      {
        name: 'rowHeaderText',
        x0: 0,
        y0: PRODUCT_HEADER_HEIGHT,
        x1: PRODUCT_ROW_MARKER_WIDTH,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 30,
      },
    ],
  } as const
  const initialReadback = await waitForReadback(page, rowHeaderProbe, (result) => result.opaquePixelCounts.rowHeaderText > 1_000)
  const initialRowHeaderRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => runs.rowHeaderRunCount > 20)
  expect(initialRowHeaderRuns.rowHeaderRunCount).toBeGreaterThan(20)

  await clickProductCell(page, 1, 24)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B25')
  await page.keyboard.press('a')

  const cellEditor = page.getByTestId('cell-editor-input')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue('a')
  await expect
    .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
    .toBe(1)
  const pressRemainingText = async (remainingCharacters: readonly string[], previousText: string): Promise<void> => {
    const [character, ...rest] = remainingCharacters
    if (!character) {
      return
    }
    const nextText = `${previousText}${character}`
    await cellEditor.press(character)
    await expect(cellEditor).toHaveValue(nextText)
    await expect
      .poll(async () => await cellEditor.evaluate((input) => (input instanceof HTMLTextAreaElement ? input.selectionStart : -1)))
      .toBe(nextText.length)
    await pressRemainingText(rest, nextText)
  }
  await pressRemainingText(['b', 'c', 'd', 'e', 'f'], 'a')

  await clickProductCell(page, 3, 25)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D26')
  await expect(cellEditor).toHaveCount(0)
  await waitForReadbackSequence(page, initialReadback.sequence)

  await clickProductCell(page, 1, 24)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B25')
  await expect(page.getByLabel('Formula')).toHaveValue('abcdef')
  await clickProductCell(page, 3, 25)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D26')

  const committedProbe = {
    points: [],
    regions: [
      {
        name: 'rowHeaderText',
        x0: 0,
        y0: PRODUCT_HEADER_HEIGHT,
        x1: PRODUCT_ROW_MARKER_WIDTH,
        y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 30,
      },
    ],
  } as const
  const committedReadback = await waitForReadback(page, committedProbe, (result) => result.opaquePixelCounts.rowHeaderText > 1_000)
  const committedTextRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => runs.rowHeaderRunCount > 20)

  expect(committedReadback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(1_000)
  expect(committedTextRuns.rowHeaderRunCount).toBeGreaterThan(20)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-click-away-edit-row-headers-readback.png',
    'main-workbook-click-away-edit-row-headers-readback',
  )
})

test('@browser-webgpu @browser-serial main workbook shell keeps the live typegpu layer stable during click-away edit commits', async ({
  page,
}) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-click-away-edit-no-proof-blink'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)

  const renderer = page.getByTestId('grid-pane-renderer')
  await expect.poll(async () => await renderer.getAttribute('data-v3-frame-proof-status')).toBe('presented')
  await expect(renderer).toHaveAttribute('data-v3-canvas-proof-layer', 'disabled')
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)

  const formulaInput = page.getByTestId('formula-input')
  await exerciseClickAwayEditCommit(page, {
    address: 'B1',
    awayCol: 2,
    awayRow: 0,
    awaySelection: 'Sheet1!C1',
    col: 1,
    formulaInput,
    renderer,
    row: 0,
    text: 'abc',
  })
  await exerciseClickAwayEditCommit(page, {
    address: 'B2',
    awayCol: 2,
    awayRow: 1,
    awaySelection: 'Sheet1!C2',
    col: 1,
    formulaInput,
    renderer,
    row: 1,
    text: 'def',
  })
})

test('@browser-webgpu @browser-perf main workbook shell keeps resident typegpu content visible while selection moves', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(
    page,
    `/?document=${encodeURIComponent(createTestDocumentId('typegpu-selection-no-flash'))}&benchmarkCorpus=wide-mixed-250k`,
  )
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const probe = {
    points: [],
    regions: [
      { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
      { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
      { name: 'bodyGrid', x0: 70, y0: 48, x1: 360, y1: 150 },
    ],
  } as const

  const initialReadback = await waitForReadback(page, probe, (result) => {
    return (
      result.opaquePixelCounts.columnHeaderText > 100 &&
      result.opaquePixelCounts.rowHeaderText > 100 &&
      result.opaquePixelCounts.bodyGrid > 400
    )
  })
  await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5
  })
  await warmStartWorkbookScrollPerf(page, 'typegpu-selection-overlay-only')
  await settleWorkbookScrollPerf(page, 16)

  const selectionTargets = [
    { col: 2, row: 3, address: '!C4' },
    { col: 5, row: 6, address: '!F7' },
    { col: 1, row: 8, address: '!B9' },
    { col: 3, row: 4, address: '!D5' },
  ] as const
  const sampleSelectionTarget = async (
    targetIndex: number,
    lastSequence: number,
    frames: ReadonlyArray<DynamicReadbackResult>,
  ): Promise<ReadonlyArray<DynamicReadbackResult>> => {
    const target = selectionTargets[targetIndex]
    if (!target) {
      return frames
    }
    await clickProductCell(page, target.col, target.row)
    await expect(page.getByTestId('status-selection')).toContainText(target.address)
    await page.waitForTimeout(50)
    const frame = await inspectGpuReadback(page, probe)
    return await sampleSelectionTarget(targetIndex + 1, Math.max(lastSequence, frame.sequence), [...frames, frame])
  }
  const selectionFrames = await sampleSelectionTarget(0, initialReadback.sequence, [])
  await settleWorkbookScrollPerf(page, 4)
  const perfReport = await stopWorkbookScrollPerf(page)

  expect(selectionFrames.length).toBe(4)
  for (const frame of selectionFrames) {
    expect(frame.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
    expect(frame.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
    expect(frame.opaquePixelCounts.bodyGrid).toBeGreaterThan(400)
  }
  const finalSelectionTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5
  })
  expect(finalSelectionTextRuns.matches.columnHeaderA).toBe(true)
  expect(finalSelectionTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(perfReport).not.toBeNull()
  expect(perfReport?.counters.headerPaneBuilds).toBeLessThanOrEqual(1)
  expect(perfReport?.counters.typeGpuBufferAllocations).toBe(0)
  expect(perfReport?.counters.typeGpuTileMisses).toBe(0)
  expect(perfReport?.counters.rendererTileMisses).toBe(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-selection-no-flash-readback.png',
    'main-workbook-grid-selection-no-flash-readback',
  )
})

test('@browser-webgpu @browser-perf main workbook shell keeps header labels and body text visible while scrolling through typegpu', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, '/?benchmarkCorpus=wide-mixed-250k')
  await waitForWorkbookReady(page)
  await waitForBenchmarkCorpus(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialReadback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
        { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
        { name: 'bodyText', x0: 70, y0: 48, x1: 280, y1: 120 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeaderA', text: 'A', exact: true }], (runs) => {
    return runs.matches.columnHeaderA && runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollLeft = 1_768
    viewport.scrollTop = 418
    viewport.dispatchEvent(new Event('scroll'))
  })
  await waitForReadbackSequence(page, initialReadback.sequence)

  const scrolledReadback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: 60, y0: 4, x1: 220, y1: 18 },
        { name: 'rowHeaderText', x0: 6, y0: 48, x1: 36, y1: 140 },
        { name: 'bodyText', x0: 70, y0: 48, x1: 280, y1: 120 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  const scrolledTextRuns = await waitForVisibleNativeTextRuns(page, [{ name: 'columnHeader', text: 'E', exact: true }], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  expect(scrolledReadback.sequence).toBeGreaterThan(initialReadback.sequence)
  expect(scrolledReadback.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
  expect(scrolledReadback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
  expect(scrolledTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(scrolledTextRuns.visibleRunCount).toBeGreaterThan(20)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-scrolled-readback.png', 'main-workbook-grid-scrolled-readback')
})

test('@browser-webgpu @browser-deep main workbook shell keeps typegpu grid lines exactly aligned after diagonal scroll', async ({
  page,
}, testInfo) => {
  const scrollLeft = PRODUCT_COLUMN_WIDTH * 4 + 17
  const scrollTop = PRODUCT_ROW_HEIGHT * 5 + 9
  const visibleStartCol = Math.floor(scrollLeft / PRODUCT_COLUMN_WIDTH)
  const visibleStartRow = Math.floor(scrollTop / PRODUCT_ROW_HEIGHT)
  const verticalLineAfterCol = visibleStartCol + 3
  const horizontalLineAfterRow = visibleStartRow + 7
  const verticalLineX = PRODUCT_ROW_MARKER_WIDTH + (verticalLineAfterCol + 1) * PRODUCT_COLUMN_WIDTH - scrollLeft - 1
  const horizontalLineY = PRODUCT_HEADER_HEIGHT + (horizontalLineAfterRow + 1) * PRODUCT_ROW_HEIGHT - scrollTop - 1
  const bodyProbeY = PRODUCT_HEADER_HEIGHT + 180
  const bodyProbeX = PRODUCT_ROW_MARKER_WIDTH + 360

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, target) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = target.left
      viewport.scrollTop = target.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: scrollLeft, top: scrollTop },
  )
  await waitForReadbackSequence(page, initialSequence)

  const readback = await waitForReadback(
    page,
    {
      points: [
        { name: 'headerVerticalLine', x: verticalLineX, y: Math.floor(PRODUCT_HEADER_HEIGHT / 2) },
        { name: 'bodyVerticalLine', x: verticalLineX, y: bodyProbeY },
        { name: 'bodyVerticalBlank', x: verticalLineX - 3, y: bodyProbeY },
        { name: 'rowHeaderHorizontalLine', x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), y: horizontalLineY },
        { name: 'bodyHorizontalLine', x: bodyProbeX, y: horizontalLineY },
        { name: 'bodyHorizontalBlank', x: bodyProbeX, y: horizontalLineY - 3 },
      ],
      regions: [],
    },
    (result) =>
      result.points.headerVerticalLine.a > 150 &&
      result.points.bodyVerticalLine.a > 150 &&
      result.points.rowHeaderHorizontalLine.a > 150 &&
      result.points.bodyHorizontalLine.a > 150,
  )

  expect(readback.points.headerVerticalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyVerticalLine.a).toBeGreaterThan(150)
  expectWorkbookWhitePixel(readback.points.bodyVerticalBlank)
  expect(readback.points.rowHeaderHorizontalLine.a).toBeGreaterThan(150)
  expect(readback.points.bodyHorizontalLine.a).toBeGreaterThan(150)
  expectWorkbookWhitePixel(readback.points.bodyHorizontalBlank)
  expect(verticalLineX).toBeGreaterThan(PRODUCT_ROW_MARKER_WIDTH)
  expect(horizontalLineY).toBeGreaterThan(PRODUCT_HEADER_HEIGHT)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-exact-scroll-readback.png', 'main-workbook-grid-exact-scroll-readback')
})

test('@browser-webgpu @browser-deep main workbook shell draws typegpu resize guides at exact geometry positions', async ({
  page,
}, testInfo) => {
  const columnGuideX = PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH - 1
  const rowGuideY = PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT - 1

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.mouse.move(grid.x + columnGuideX, grid.y + Math.floor(PRODUCT_HEADER_HEIGHT / 2))
  await waitForReadbackSequence(page, initialSequence)
  const columnReadback = await waitForReadback(
    page,
    {
      points: [
        { name: 'columnGuideHeader', x: columnGuideX, y: Math.floor(PRODUCT_HEADER_HEIGHT / 2) },
        { name: 'columnGuideBody', x: columnGuideX, y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 },
        { name: 'columnGuideAdjacent', x: columnGuideX - 3, y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 },
      ],
      regions: [],
    },
    (result) => isResizeGuidePixel(result.points.columnGuideHeader) && isResizeGuidePixel(result.points.columnGuideBody),
  )

  expect(isResizeGuidePixel(columnReadback.points.columnGuideHeader)).toBe(true)
  expect(isResizeGuidePixel(columnReadback.points.columnGuideBody)).toBe(true)
  expectWorkbookWhitePixel(columnReadback.points.columnGuideAdjacent)

  await page.mouse.move(grid.x + Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), grid.y + rowGuideY)
  await page.mouse.down()
  await waitForReadbackSequence(page, columnReadback.sequence)
  const rowReadback = await waitForReadback(
    page,
    {
      points: [
        { name: 'rowGuideHeader', x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2), y: rowGuideY },
        { name: 'rowGuideBody', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3, y: rowGuideY },
        { name: 'rowGuideAdjacent', x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3, y: rowGuideY - 3 },
      ],
      regions: [],
    },
    (result) => isResizeGuidePixel(result.points.rowGuideHeader) && isResizeGuidePixel(result.points.rowGuideBody),
  )

  expect(isResizeGuidePixel(rowReadback.points.rowGuideHeader)).toBe(true)
  expect(isResizeGuidePixel(rowReadback.points.rowGuideBody)).toBe(true)
  expectWorkbookWhitePixel(rowReadback.points.rowGuideAdjacent)
  await page.mouse.up()

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-resize-guide-readback.png', 'main-workbook-grid-resize-guide-readback')
})

test('@browser-webgpu @browser-deep main workbook shell keeps DOM editor overlay aligned to typegpu geometry while scrolling', async ({
  page,
}, testInfo) => {
  const targetCol = 3
  const targetRow = 7
  const scrollLeft = 37
  const scrollTop = 13
  const expectedLocalX = PRODUCT_ROW_MARKER_WIDTH + targetCol * PRODUCT_COLUMN_WIDTH - scrollLeft
  const expectedLocalY = PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT - scrollTop

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  await clickProductCell(page, targetCol, targetRow)
  await page.keyboard.press('F2')
  await expect(page.getByTestId('cell-editor-input')).toBeVisible()

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, target) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = target.left
      viewport.scrollTop = target.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: scrollLeft, top: scrollTop },
  )
  await waitForReadbackSequence(page, initialSequence)

  const dpr = await page.evaluate(() => window.devicePixelRatio || 1)
  const tolerance = Math.max(1, 1 / dpr)
  const expectedViewportRect = {
    x: grid.x + expectedLocalX,
    y: grid.y + expectedLocalY,
    width: PRODUCT_COLUMN_WIDTH,
    height: PRODUCT_ROW_HEIGHT,
  }

  await expect
    .poll(
      async () => {
        const box = await page.getByTestId('cell-editor-overlay').boundingBox()
        if (!box) {
          return Number.POSITIVE_INFINITY
        }
        return Math.max(
          Math.abs(box.x - expectedViewportRect.x),
          Math.abs(box.y - expectedViewportRect.y),
          Math.abs(box.width - expectedViewportRect.width),
          Math.abs(box.height - expectedViewportRect.height),
        )
      },
      { timeout: 15_000 },
    )
    .toBeLessThanOrEqual(tolerance)

  const editorBox = await page.getByTestId('cell-editor-overlay').boundingBox()
  if (!editorBox) {
    throw new Error('cell editor overlay is not visible')
  }
  expectNear(editorBox.x, expectedViewportRect.x, tolerance)
  expectNear(editorBox.y, expectedViewportRect.y, tolerance)
  expectNear(editorBox.width, expectedViewportRect.width, tolerance)
  expectNear(editorBox.height, expectedViewportRect.height, tolerance)

  const readback = await waitForReadback(
    page,
    {
      points: [
        { name: 'activeCellTopBorder', x: expectedLocalX + Math.floor(PRODUCT_COLUMN_WIDTH / 2), y: expectedLocalY },
        { name: 'activeCellLeftBorder', x: expectedLocalX, y: expectedLocalY + Math.floor(PRODUCT_ROW_HEIGHT / 2) },
      ],
      regions: [],
    },
    (result) => result.points.activeCellTopBorder.a > 150 && result.points.activeCellLeftBorder.a > 150,
  )

  expect(readback.points.activeCellTopBorder.a).toBeGreaterThan(150)
  expect(readback.points.activeCellLeftBorder.a).toBeGreaterThan(150)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-editor-overlay-readback.png', 'main-workbook-grid-editor-overlay-readback')
})

test('@browser-webgpu @browser-deep main workbook shell refreshes typegpu resident packets after style-only changes', async ({
  page,
}, testInfo) => {
  const fillPoint = {
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-style-refresh'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const initialReadback = await waitForReadback(
    page,
    {
      points: [{ name: 'cellFill', ...fillPoint }],
      regions: [],
    },
    (result) => isWorkbookWhitePixel(result.points.cellFill),
  )

  await clickProductCell(page, 1, 1)
  await expect(page.getByTestId('status-selection')).toContainText('!B2')
  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 3')
  await waitForReadbackSequence(page, initialReadback.sequence)
  const afterStyleSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })
  await clickProductCell(page, 2, 2)
  await waitForReadbackSequence(page, afterStyleSequence)

  const styledReadback = await waitForReadback(
    page,
    {
      points: [{ name: 'cellFill', ...fillPoint }],
      regions: [],
    },
    (result) => result.points.cellFill.a > 200,
  )

  expect(styledReadback.points.cellFill.a).toBe(255)
  expect(styledReadback.points.cellFill.r).toBeGreaterThan(170)
  expect(styledReadback.points.cellFill.r).toBeLessThan(215)
  expect(styledReadback.points.cellFill.g).toBeGreaterThan(190)
  expect(styledReadback.points.cellFill.g).toBeLessThan(230)
  expect(styledReadback.points.cellFill.b).toBeGreaterThan(220)
  expect(styledReadback.points.cellFill.b).toBeLessThan(255)
  expect(styledReadback.points.cellFill.b).toBeGreaterThan(styledReadback.points.cellFill.g)
  expect(styledReadback.points.cellFill.g).toBeGreaterThan(styledReadback.points.cellFill.r)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-style-refresh-readback.png', 'main-workbook-grid-style-refresh-readback')
})

test('@browser-webgpu @browser-deep selected range fill changes stay visually authoritative while selected', async ({ page }, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'topLeft' },
    { ...selectedRangeFillProbe(2, 4), name: 'middle' },
    { ...selectedRangeFillProbe(3, 8), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-selected-fill-refresh'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  await dragProductBodySelection(page, 1, 1, 3, 8)
  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 2')
  const blueReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isCornflowerBlueFill),
  )

  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => result.sequence > blueReadback.sequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )

  for (const [name, point] of Object.entries(greenReadback.points)) {
    expect(point.g, `${name} green channel`).toBeGreaterThan(point.b)
    expect(point.g, `${name} green channel`).toBeGreaterThan(point.r)
  }

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-selected-fill-refresh-readback.png',
    'main-workbook-grid-selected-fill-refresh-readback',
  )
})

test('@browser-webgpu @browser-deep fill-handle drag commits through the typegpu readback layer', async ({ page }, testInfo) => {
  const sourceColumn = 2
  const sourceRow = 5
  const targetRow = 7
  const copiedText = 'fill-drag-pixels'
  const points = [
    { ...selectedRangeFillProbe(sourceColumn, sourceRow), name: 'sourceC6' },
    { ...selectedRangeFillProbe(sourceColumn, sourceRow + 1), name: 'targetC7' },
    { ...selectedRangeFillProbe(sourceColumn, targetRow), name: 'targetC8' },
  ]
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-fill-handle-commit'))}&persist=0`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-backend-status', 'ready')
  await expect(page.getByTestId('grid-native-text-layer')).toHaveCount(1)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const formulaInput = page.getByTestId('formula-input')
  await clickProductCell(page, sourceColumn, sourceRow)
  await formulaInput.fill(copiedText)
  await formulaInput.press('Enter')
  await clickProductCell(page, sourceColumn, sourceRow)
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const sourceReadback = await waitForReadback(
    page,
    {
      points: [points[0]],
      regions: [],
    },
    (result) => isThemeGreenFill(result.points.sourceC6),
  )

  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(
    page,
    sourceColumn,
    sourceRow,
    sourceColumn,
    targetRow,
  )
  const handleState = await page.locator("[data-grid-fill-handle='true']").evaluate((node) => {
    const bounds = node.getBoundingClientRect()
    const style = window.getComputedStyle(node)
    return {
      display: style.display,
      pointerEvents: style.pointerEvents,
      x0: bounds.left,
      x1: bounds.right,
      y0: bounds.top,
      y1: bounds.bottom,
    }
  })
  expect(handleState.display).not.toBe('none')
  expect(handleState.pointerEvents).not.toBe('none')
  expect(sourceX).toBeGreaterThanOrEqual(handleState.x0)
  expect(sourceX).toBeLessThanOrEqual(handleState.x1)
  expect(sourceY).toBeGreaterThanOrEqual(handleState.y0)
  expect(sourceY).toBeLessThanOrEqual(handleState.y1)
  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 10 })
  await expect(page.locator("[data-grid-fill-preview='true']")).toBeVisible()
  await page.mouse.up()

  await expect(page.getByTestId('status-selection')).toContainText('!C6:C8')
  await expect(page.getByTestId('grid-native-text-layer')).toHaveCount(1)
  await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => result.sequence > sourceReadback.sequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )
  await expect
    .poll(async () =>
      page.evaluate(
        (text) => [...document.querySelectorAll('[data-native-text-run]')].filter((run) => run.textContent === text).length,
        copiedText,
      ),
    )
    .toBeGreaterThanOrEqual(3)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-fill-handle-commit-readback.png',
    'main-workbook-grid-fill-handle-commit-readback',
  )
})

test('@browser-webgpu @browser-deep moved content delete preserves selected fill without fallback flashes', async ({ page }, testInfo) => {
  const movedText = 'move-fill-delete-proof'
  const movedTextRegion = {
    name: 'movedTextGlyphs',
    x0: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 3 + 8,
    y0: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 3 + 5,
    x1: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 4 - 8,
    y1: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 4 - 5,
  } as const
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'topLeft' },
    { ...selectedRangeFillProbe(2, 4), name: 'middle' },
    { ...selectedRangeFillProbe(3, 8), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-move-fill-delete-no-flash'))}`)
  await waitForWorkbookReady(page)
  await waitForTypeGpuRenderer(page)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-backend-status', 'ready')
  await expect(page.getByTestId('grid-native-text-layer')).toHaveCount(1)
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-draw-text', 'false')
  await expect(page.getByTestId('grid-pane-renderer')).toHaveAttribute('data-v3-native-layer-source', /browser-native-text/)
  await page.waitForFunction(
    () =>
      Boolean(
        (window as Window & { __biligGpuReadbackInspector?: { readonly isReady: () => boolean } }).__biligGpuReadbackInspector?.isReady(),
      ),
    undefined,
    { timeout: 15_000 },
  )

  const formulaInput = page.getByTestId('formula-input')
  const grid = page.getByTestId('sheet-grid')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill(movedText)
  await formulaInput.press('Enter')
  await expect(formulaInput).toHaveValue(movedText)
  await dragProductSelectionBorder(page, 1, 1, 3, 3)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4')

  await clickProductCell(page, 1, 1)
  await expect(formulaInput).toHaveValue('')
  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue(movedText)
  await waitForNativeTextRunContent(page, movedText)

  await dragProductBodySelection(page, 1, 1, 3, 8)
  await expect(page.getByTestId('status-selection')).toContainText('!B2:D9')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  const deleteSamples = await collectRendererPresentationSamplesDuring(page, () => grid.press('Delete'), 30)
  expect(deleteSamples.filter((sample) => sample.fallbackCanvases !== 0)).toEqual([])
  expect(deleteSamples.filter((sample) => sample.canvasProofLayer !== 'disabled')).toEqual([])

  const afterDeleteReadback = await waitForReadback(
    page,
    {
      points,
      regions: [movedTextRegion],
    },
    (result) =>
      result.sequence >= greenReadback.sequence &&
      allReadbackPointsMatch(result, isThemeGreenFill) &&
      result.darkPixelCounts.movedTextGlyphs <= 2,
  )
  expect(afterDeleteReadback.sequence).toBeGreaterThanOrEqual(greenReadback.sequence)
  expect(afterDeleteReadback.darkPixelCounts.movedTextGlyphs).toBeLessThanOrEqual(2)

  await clickProductCell(page, 3, 3)
  await expect(formulaInput).toHaveValue('')
  await expect
    .poll(async () =>
      page.evaluate((text) => [...document.querySelectorAll('[data-native-text-run]')].some((run) => run.textContent === text), movedText),
    )
    .toBe(false)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-move-fill-delete-no-flash-readback.png',
    'main-workbook-grid-move-fill-delete-no-flash-readback',
  )
})
