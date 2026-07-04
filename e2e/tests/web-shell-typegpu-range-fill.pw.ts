import { expect, test } from '@playwright/test'
import {
  clickProductCell,
  createTestDocumentId,
  dragProductHeaderSelection,
  gotoWorkbookShell,
  pickToolbarPresetColor,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRIMARY_MODIFIER,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  waitForBenchmarkCorpus,
  waitForWorkbookReady,
} from './web-shell-helpers.js'
import {
  allReadbackPointsMatch,
  expectPresentedOverlayRects,
  expectSelectionVisualOpacity,
  installTypeGpuReadbackHarness,
  isCornflowerBlueFill,
  isThemeGreenFill,
  saveReadbackArtifact,
  selectedRangeFillProbe,
  waitForReadback,
  waitForReadbackSequence,
  waitForTypeGpuRenderer,
  waitForVisibleNativeTextRuns,
} from './web-shell-typegpu-helpers.js'

test('@browser-webgpu @browser-deep axis selections paint atomically over the typegpu grid without hiding active cells', async ({
  page,
}, testInfo) => {
  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-axis-selection-proof'))}&persist=0`)
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

  const grid = page.getByTestId('sheet-grid')
  await grid.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * 2 + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!C:C')
  await expectPresentedOverlayRects(page, 'column axis selection fill')
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-key="header-fill:column:2"]'),
    'selected column header fill',
    'visible',
  )
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-role="selection-fill"]'),
    'selected column body fill',
    'visible',
  )
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-key="header-fill:column:3"]')).toHaveCount(0)

  await clickProductCell(page, 1, 1)
  await grid.click({
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * 5 + Math.floor(PRODUCT_ROW_HEIGHT / 2),
    },
  })
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!6:6')
  await expectPresentedOverlayRects(page, 'row axis selection fill')
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-key="header-fill:row:5"]'),
    'selected row header fill',
    'visible',
  )
  await expectSelectionVisualOpacity(
    page.locator('[data-grid-selection-visual-role="selection-fill"]'),
    'selected row body fill',
    'visible',
  )
  await expect(page.locator('[data-grid-selection-visual-role="active-border"]')).toHaveCount(0)
  await expect(page.locator('[data-grid-selection-visual-key="header-fill:row:6"]')).toHaveCount(0)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-axis-selection-readback.png', 'main-workbook-grid-axis-selection-readback')
})

test('@browser-webgpu @browser-deep name-box range fill presents the current frame before success', async ({ page }, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(3, 3), name: 'topLeft' },
    { ...selectedRangeFillProbe(4, 5), name: 'middle' },
    { ...selectedRangeFillProbe(5, 7), name: 'bottomRight' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-name-box-fill-refresh'))}&persist=0`)
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

  const nameBox = page.getByTestId('name-box')
  await nameBox.fill('D4:F8')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F8')

  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const greenReadback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )
  const rendererProof = await page.getByTestId('grid-pane-renderer').evaluate((canvas) => ({
    frameProofStatus: canvas.getAttribute('data-v3-frame-proof-status'),
    projectedRenderRevision: canvas.getAttribute('data-v3-projected-render-revision'),
    tileSceneRevision: canvas.getAttribute('data-v3-tile-scene-revision'),
    visibleProjectedRenderRevision: canvas.getAttribute('data-v3-visible-projected-render-revision'),
    visibleRenderRevision: canvas.getAttribute('data-v3-visible-render-revision'),
  }))
  const gridProjectedRevision = await page.getByTestId('sheet-grid').getAttribute('data-render-projected-revision')

  expect(rendererProof.frameProofStatus).toBe('presented')
  expect(rendererProof.projectedRenderRevision).toBe(gridProjectedRevision)
  expect(rendererProof.visibleProjectedRenderRevision).toBe(gridProjectedRevision)
  expect(rendererProof.visibleRenderRevision).toBe(rendererProof.tileSceneRevision)
  expect(greenReadback.sequence).toBeGreaterThan(0)
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F8')
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-name-box-fill-refresh-readback.png',
    'main-workbook-grid-name-box-fill-refresh-readback',
  )
})

test('@browser-webgpu @browser-deep large range fill remains applied after scrolling into uncached rows', async ({ page }, testInfo) => {
  const initialPoints = [
    { ...selectedRangeFillProbe(3, 3), name: 'topLeft' },
    { ...selectedRangeFillProbe(4, 8), name: 'middle' },
    { ...selectedRangeFillProbe(5, 12), name: 'bottomRight' },
  ]
  const scrolledPoints = [
    { ...selectedRangeFillProbe(3, 5), name: 'scrolledD' },
    { ...selectedRangeFillProbe(4, 8), name: 'scrolledE' },
    { ...selectedRangeFillProbe(5, 11), name: 'scrolledF' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-large-range-fill-scroll'))}&persist=0`)
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

  const nameBox = page.getByTestId('name-box')
  await nameBox.fill('D4:F240')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!D4:F240')
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const initialReadback = await waitForReadback(
    page,
    {
      points: initialPoints,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  await pickToolbarPresetColor(page, 'Fill color', 'light cornflower blue 2')
  const repaintReadback = await waitForReadback(
    page,
    {
      points: initialPoints,
      regions: [],
    },
    (result) => result.sequence > initialReadback.sequence && allReadbackPointsMatch(result, isCornflowerBlueFill),
  )
  expect(repaintReadback.sequence).toBeGreaterThan(initialReadback.sequence)

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport, rowHeight) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollTop = rowHeight * 160
    viewport.dispatchEvent(new Event('scroll'))
  }, PRODUCT_ROW_HEIGHT)
  await waitForReadbackSequence(page, initialReadback.sequence)

  const scrolledReadback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) => result.sequence > repaintReadback.sequence && allReadbackPointsMatch(result, isCornflowerBlueFill),
  )
  expect(scrolledReadback.sequence).toBeGreaterThan(repaintReadback.sequence)

  await clickProductCell(page, 4, 8)
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const overrideReadback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) =>
      result.sequence > scrolledReadback.sequence &&
      isCornflowerBlueFill(result.points.scrolledD) &&
      isThemeGreenFill(result.points.scrolledE) &&
      isCornflowerBlueFill(result.points.scrolledF),
  )
  expect(overrideReadback.sequence).toBeGreaterThan(scrolledReadback.sequence)
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-large-range-fill-scroll-readback.png',
    'main-workbook-grid-large-range-fill-scroll-readback',
  )
})

test('@browser-webgpu @browser-deep keyboard fill down paints uncached rows before sync catch-up', async ({ page }, testInfo) => {
  const scrolledPoints = [
    { ...selectedRangeFillProbe(1, 5), name: 'bVisibleTop' },
    { ...selectedRangeFillProbe(1, 8), name: 'bVisibleMiddle' },
    { ...selectedRangeFillProbe(1, 11), name: 'bVisibleBottom' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-keyboard-filldown-scroll'))}&persist=0`)
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

  const formulaInput = page.getByTestId('formula-input')
  const nameBox = page.getByTestId('name-box')
  const grid = page.getByTestId('sheet-grid')

  await clickProductCell(page, 1, 1)
  await formulaInput.fill('filldown-overlay')
  await formulaInput.press('Enter')
  await clickProductCell(page, 1, 1)
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')

  await nameBox.fill('B2:B20000')
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText('Sheet1!B2:B20000')
  await grid.press(`${PRIMARY_MODIFIER}+D`)

  const filledSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })
  await page.getByTestId('grid-scroll-viewport').evaluate((viewport, rowHeight) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollTop = rowHeight * 160
    viewport.dispatchEvent(new Event('scroll'))
  }, PRODUCT_ROW_HEIGHT)
  await waitForReadbackSequence(page, filledSequence)

  const readback = await waitForReadback(
    page,
    {
      points: scrolledPoints,
      regions: [],
    },
    (result) => result.sequence > filledSequence && allReadbackPointsMatch(result, isThemeGreenFill),
  )
  expect(readback.sequence).toBeGreaterThan(filledSequence)
  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-keyboard-filldown-scroll-readback.png',
    'main-workbook-grid-keyboard-filldown-scroll-readback',
  )
})

test('@browser-webgpu @browser-deep full-column fill changes repaint visible cells without waiting for sync catch-up', async ({
  page,
}, testInfo) => {
  const points = [
    { ...selectedRangeFillProbe(1, 1), name: 'b2' },
    { ...selectedRangeFillProbe(2, 8), name: 'c9' },
    { ...selectedRangeFillProbe(4, 20), name: 'e21' },
  ]

  await page.setViewportSize({ width: 960, height: 720 })
  await installTypeGpuReadbackHarness(page)
  await gotoWorkbookShell(page, `/?document=${encodeURIComponent(createTestDocumentId('typegpu-column-fill-fast'))}&persist=0`)
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

  await dragProductHeaderSelection(page, 'column', 1, 4)
  await expect(page.getByTestId('status-selection')).toContainText('!B:E')
  const startedAt = Date.now()
  await pickToolbarPresetColor(page, 'Fill color', 'theme green')
  const readback = await waitForReadback(
    page,
    {
      points,
      regions: [],
    },
    (result) => allReadbackPointsMatch(result, isThemeGreenFill),
  )

  expect(Date.now() - startedAt).toBeLessThan(3_000)
  expect(readback.sequence).toBeGreaterThan(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-full-column-fill-fast-readback.png',
    'main-workbook-grid-full-column-fill-fast-readback',
  )
})

test('@browser-webgpu @browser-perf main workbook shell keeps typegpu content visible after hover-driven scroll', async ({
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

  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  await page.mouse.move(grid.x + PRODUCT_ROW_MARKER_WIDTH + 180, grid.y + PRODUCT_HEADER_HEIGHT + 96)

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not a div')
    }
    viewport.scrollLeft = 2_600
    viewport.scrollTop = 520
    viewport.dispatchEvent(new Event('scroll'))
  })
  await waitForReadbackSequence(page, initialSequence)

  const readback = await waitForReadback(
    page,
    {
      points: [],
      regions: [
        { name: 'columnHeaderText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: 0, x1: 360, y1: PRODUCT_HEADER_HEIGHT },
        { name: 'rowHeaderText', x0: 0, y0: PRODUCT_HEADER_HEIGHT, x1: PRODUCT_ROW_MARKER_WIDTH, y1: 220 },
        { name: 'bodyText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: PRODUCT_HEADER_HEIGHT, x1: 420, y1: 240 },
      ],
    },
    (result) => result.opaquePixelCounts.columnHeaderText > 100 && result.opaquePixelCounts.rowHeaderText > 100,
  )
  const hoverScrollTextRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })

  expect(readback.opaquePixelCounts.columnHeaderText).toBeGreaterThan(100)
  expect(readback.opaquePixelCounts.rowHeaderText).toBeGreaterThan(100)
  expect(hoverScrollTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(hoverScrollTextRuns.visibleRunCount).toBeGreaterThan(20)

  await saveReadbackArtifact(page, testInfo, 'main-workbook-grid-hover-scroll-readback.png', 'main-workbook-grid-hover-scroll-readback')
})

test('@browser-webgpu @browser-perf main workbook shell keeps grid content and native text visible across tile boundary scroll and resize', async ({
  page,
}, testInfo) => {
  test.slow()
  await page.setViewportSize({ width: 900, height: 680 })
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

  const initialSequence = await page.evaluate(() => {
    return (
      (
        window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } }
      ).__biligGpuReadbackInspector?.getSequence() ?? 0
    )
  })

  await page.getByTestId('grid-scroll-viewport').evaluate(
    (viewport, scrollTarget) => {
      if (!(viewport instanceof HTMLDivElement)) {
        throw new Error('grid scroll viewport is not a div')
      }
      viewport.scrollLeft = scrollTarget.left
      viewport.scrollTop = scrollTarget.top
      viewport.dispatchEvent(new Event('scroll'))
    },
    { left: PRODUCT_COLUMN_WIDTH * 130, top: PRODUCT_ROW_HEIGHT * 34 },
  )
  await waitForReadbackSequence(page, initialSequence)

  await page.setViewportSize({ width: 960, height: 720 })
  await waitForReadbackSequence(page, initialSequence + 1)

  const readback = await waitForReadback(
    page,
    {
      points: [{ name: 'blankBody', x: PRODUCT_ROW_MARKER_WIDTH + 20, y: PRODUCT_HEADER_HEIGHT + 40 }],
      regions: [
        { name: 'canvasDark', x0: 0, y0: 0, x1: 960, y1: 720, threshold: 220 },
        { name: 'columnHeaderText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: 0, x1: 960, y1: PRODUCT_HEADER_HEIGHT },
        { name: 'rowHeaderText', x0: 0, y0: PRODUCT_HEADER_HEIGHT, x1: PRODUCT_ROW_MARKER_WIDTH, y1: 720 },
        { name: 'bodyText', x0: PRODUCT_ROW_MARKER_WIDTH, y0: PRODUCT_HEADER_HEIGHT, x1: 960, y1: 720 },
      ],
    },
    (result) => result.opaquePixelCounts.canvasDark > 200,
  )

  expect(readback.opaquePixelCounts.canvasDark).toBeGreaterThan(200)
  const tileBoundaryTextRuns = await waitForVisibleNativeTextRuns(page, [], (runs) => {
    return runs.rowHeaderRunCount > 5 && runs.visibleRunCount > 20
  })
  expect(tileBoundaryTextRuns.rowHeaderRunCount).toBeGreaterThan(5)
  expect(tileBoundaryTextRuns.visibleRunCount).toBeGreaterThan(20)
  expect(readback.points.blankBody.a).toBeGreaterThanOrEqual(0)

  await saveReadbackArtifact(
    page,
    testInfo,
    'main-workbook-grid-tile-boundary-resize-readback.png',
    'main-workbook-grid-tile-boundary-resize-readback',
  )
})
