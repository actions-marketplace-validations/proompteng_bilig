import { expect, type Page } from '@playwright/test'
import {
  PRIMARY_MODIFIER,
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  countDarkReadbackPixelsInCell,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowHeight,
  getProductRowTop,
  settleWorkbookScrollPerf,
} from './web-shell-helpers.js'

export async function dragSelectedRangeBorderTowardBottom(page: Page): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, 1)
  const targetLeft = await getProductColumnLeft(page, 1)
  const targetWidth = await getProductColumnWidth(page, 1)
  const sourceX = grid.x + startLeft + 3
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT + 2
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + grid.height - 24

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 16 })
}

export async function getProductCellRangeBox(
  page: Page,
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
): Promise<{
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
}> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const leftColumn = Math.min(startColumn, endColumn)
  const rightColumn = Math.max(startColumn, endColumn)
  const topRow = Math.min(startRow, endRow)
  const bottomRow = Math.max(startRow, endRow)
  const left = await getProductColumnLeft(page, leftColumn)
  const right = await getProductColumnLeft(page, rightColumn)
  const rightWidth = await getProductColumnWidth(page, rightColumn)
  const top = await getProductRowTop(page, topRow)
  const bottom = await getProductRowTop(page, bottomRow)
  const bottomHeight = await getProductRowHeight(page, bottomRow)
  return {
    x: grid.x + left,
    y: grid.y + PRODUCT_HEADER_HEIGHT + top,
    width: right + rightWidth - left,
    height: bottom + bottomHeight - top,
  }
}

export async function expectVisualRectNear(
  locator: ReturnType<Page['locator']>,
  expected: {
    readonly x: number
    readonly y: number
    readonly width: number
    readonly height: number
  },
  label: string,
): Promise<void> {
  await expect(locator, `${label} should be unique`).toHaveCount(1)
  const actual = await locator.boundingBox()
  if (!actual) {
    throw new Error(`${label} is not visible`)
  }
  expect(actual.x, `${label} x`).toBeCloseTo(expected.x, 0)
  expect(actual.y, `${label} y`).toBeCloseTo(expected.y, 0)
  expect(actual.width, `${label} width`).toBeCloseTo(expected.width, 0)
  expect(actual.height, `${label} height`).toBeCloseTo(expected.height, 0)
}

export async function expectSelectionVisualRoles(page: Page, roles: readonly string[], expected: 'hidden' | 'visible'): Promise<void> {
  const selector = roles.map((role) => `[data-grid-selection-visual-role="${role}"]`).join(',')
  await expect
    .poll(
      async () => {
        const opacities = await page.locator(selector).evaluateAll((nodes) => nodes.map((node) => window.getComputedStyle(node).opacity))
        if (opacities.length === 0) {
          return false
        }
        const opacityValues = opacities.map((opacity) => Number(opacity))
        const hiddenInDom = opacityValues.every((opacity) => Number.isFinite(opacity) && opacity <= 0.01)
        if (expected === 'hidden') {
          return hiddenInDom
        }
        const visibleInDom = opacityValues.every((opacity) => Number.isFinite(opacity) && opacity > 0.01)
        if (visibleInDom) {
          return true
        }
        if (!(await isTypeGpuCanvasActive(page))) {
          return false
        }
        const overlayRectCount = Number(
          (await page.getByTestId('grid-pane-renderer').getAttribute('data-v3-presented-overlay-rect-count')) ?? '0',
        )
        return Number.isFinite(overlayRectCount) && overlayRectCount > 0
      },
      { message: `selection visual roles ${roles.join(', ')} should be ${expected}` },
    )
    .toBe(true)
}

export async function expectSelectedRangeBodyTint(page: Page, columnIndex: number, rowIndex: number): Promise<void> {
  await expect
    .poll(
      async () => {
        const pixel = await sampleCellInteriorPixel(page, columnIndex, rowIndex)
        return pixel.red < 212 && pixel.green < 230 && pixel.blue < 220 && pixel.green > pixel.red
      },
      {
        message: `cell ${columnIndex}:${rowIndex} should show selected range body tint`,
      },
    )
    .toBe(true)
  const pixel = await sampleCellInteriorPixel(page, columnIndex, rowIndex)
  expect(pixel.red, 'selected range interior should not read as a white hollow rectangle').toBeLessThan(212)
  expect(pixel.green, 'selected range interior should keep a visible spreadsheet selection tint').toBeLessThan(242)
  expect(pixel.blue, 'selected range interior should not read as a white hollow rectangle').toBeLessThan(220)
  expect(pixel.green, 'selected range interior should carry the green Excel-style selection cast').toBeGreaterThan(pixel.red)
}

export async function sampleCellInteriorPixel(
  page: Page,
  columnIndex: number,
  rowIndex: number,
): Promise<{
  readonly blue: number
  readonly green: number
  readonly red: number
}> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const [columnLeft, columnWidth, rowTop, rowHeight, scroll] = await Promise.all([
    getProductColumnLeft(page, columnIndex),
    getProductColumnWidth(page, columnIndex),
    getProductRowTop(page, rowIndex),
    getProductRowHeight(page, rowIndex),
    page.getByTestId('grid-scroll-viewport').evaluate((node) => ({
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    })),
  ])
  const sampleSize = 6
  const sampleX = Math.round(grid.x + columnLeft - scroll.scrollLeft + columnWidth / 2 - sampleSize / 2)
  const sampleY = Math.round(grid.y + PRODUCT_HEADER_HEIGHT + rowTop - scroll.scrollTop + rowHeight / 2 - sampleSize / 2)
  const buffer = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      height: sampleSize,
      width: sampleSize,
      x: sampleX,
      y: sampleY,
    },
  })
  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode selected-range screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for selected-range screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      let red = 0
      let green = 0
      let blue = 0
      let count = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        if (alpha <= 200) {
          continue
        }
        red += pixels[index] ?? 255
        green += pixels[index + 1] ?? 255
        blue += pixels[index + 2] ?? 255
        count += 1
      }
      if (count === 0) {
        throw new Error('No opaque selected-range screenshot pixels sampled')
      }
      return {
        blue: Math.round(blue / count),
        green: Math.round(green / count),
        red: Math.round(red / count),
      }
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}

export async function sampleCompositedViewportPixel(
  page: Page,
  x: number,
  y: number,
): Promise<{
  readonly blue: number
  readonly green: number
  readonly red: number
}> {
  const buffer = await page.screenshot({
    animations: 'disabled',
    caret: 'hide',
    clip: {
      height: 1,
      width: 1,
      x: Math.round(x),
      y: Math.round(y),
    },
  })
  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode active-cell seam screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for active-cell seam screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      let red = 0
      let green = 0
      let blue = 0
      let count = 0
      for (let index = 0; index < pixels.length; index += 4) {
        const alpha = pixels[index + 3] ?? 0
        if (alpha <= 200) {
          continue
        }
        red += pixels[index] ?? 255
        green += pixels[index + 1] ?? 255
        blue += pixels[index + 2] ?? 255
        count += 1
      }
      if (count === 0) {
        throw new Error('No opaque active-cell seam screenshot pixels sampled')
      }
      return {
        blue: Math.round(blue / count),
        green: Math.round(green / count),
        red: Math.round(red / count),
      }
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}

export async function readSelectedRowHeaderNativeTextProof(
  page: Page,
  clip: {
    readonly height: number
    readonly width: number
    readonly x: number
    readonly y: number
  },
): Promise<{
  readonly labels: readonly string[]
  readonly textLayerZIndex: number
  readonly typeGpuLayerZIndex: number
}> {
  return await page.evaluate((viewportClip) => {
    const textLayer = document.querySelector<HTMLElement>('[data-testid="grid-native-text-layer"]')
    const typeGpuLayer = document.querySelector<HTMLElement>('[data-testid="grid-pane-renderer"]')
    if (!textLayer || !typeGpuLayer) {
      throw new Error('Missing workbook text or TypeGPU layer for selected row-header proof')
    }
    const textLayerZIndex = Number.parseInt(window.getComputedStyle(textLayer).zIndex || '0', 10)
    const typeGpuLayerZIndex = Number.parseInt(window.getComputedStyle(typeGpuLayer).zIndex || '0', 10)
    const clipRight = viewportClip.x + viewportClip.width
    const clipBottom = viewportClip.y + viewportClip.height
    const labels = Array.from(textLayer.querySelectorAll<HTMLElement>('[data-native-text-run]'))
      .filter((run) => {
        if (run.getAttribute('data-native-text-run-row') || run.getAttribute('data-native-text-run-col')) {
          return false
        }
        const text = run.textContent?.trim() ?? ''
        if (!/^\d+$/u.test(text)) {
          return false
        }
        const rect = run.getBoundingClientRect()
        if (
          rect.width <= 0 ||
          rect.height <= 0 ||
          rect.right <= viewportClip.x ||
          rect.left >= clipRight ||
          rect.bottom <= viewportClip.y ||
          rect.top >= clipBottom
        ) {
          return false
        }
        const inner = run.firstElementChild
        if (!(inner instanceof HTMLElement)) {
          return false
        }
        const style = window.getComputedStyle(inner)
        return style.display !== 'none' && style.visibility !== 'hidden' && Number.parseFloat(style.opacity || '1') > 0
      })
      .map((run) => run.textContent?.trim() ?? '')
      .toSorted((left, right) => Number(left) - Number(right))

    return { labels, textLayerZIndex, typeGpuLayerZIndex }
  }, clip)
}
export function isSelectionAccentPixel(pixel: { readonly blue: number; readonly green: number; readonly red: number }): boolean {
  return Math.abs(pixel.red - 33) <= 6 && Math.abs(pixel.green - 115) <= 8 && Math.abs(pixel.blue - 70) <= 6
}

export function isGridBorderPixel(pixel: { readonly blue: number; readonly green: number; readonly red: number }): boolean {
  return Math.abs(pixel.red - 221) <= 8 && Math.abs(pixel.green - 216) <= 8 && Math.abs(pixel.blue - 204) <= 8
}

export function maxPixelChannelDistance(
  left: { readonly blue: number; readonly green: number; readonly red: number },
  right: { readonly blue: number; readonly green: number; readonly red: number },
): number {
  return Math.max(Math.abs(left.red - right.red), Math.abs(left.green - right.green), Math.abs(left.blue - right.blue))
}

export async function isTypeGpuCanvasActive(page: Page): Promise<boolean> {
  return (await page.getByTestId('grid-pane-renderer').count()) > 0
}

export async function isTypeGpuTextReadbackActive(page: Page): Promise<boolean> {
  return (await isTypeGpuCanvasActive(page)) && (await page.getByTestId('grid-native-text-layer').count()) === 0
}

export async function expectCellTextPixels(
  page: Page,
  columnIndex: number,
  rowIndex: number,
  expected: 'hidden' | 'visible',
): Promise<void> {
  const poll = expect.poll(async () => await countDarkReadbackPixelsInCell(page, columnIndex, rowIndex), {
    message: `cell ${columnIndex}:${rowIndex} text pixels should be ${expected}`,
  })
  if (expected === 'visible') {
    await poll.toBeGreaterThan(4)
    return
  }
  await poll.toBeLessThanOrEqual(2)
}

export async function expectBorderStyle(
  locator: ReturnType<Page['locator']>,
  expected: {
    readonly boxShadow: string
    readonly color?: string | undefined
    readonly width: string
  },
): Promise<void> {
  const actual = await locator.evaluate((node) => {
    const style = window.getComputedStyle(node)
    const edges = Array.from(node.querySelectorAll<HTMLElement>('[data-grid-selection-visual-edge]')).map((edge) => {
      const edgeStyle = window.getComputedStyle(edge)
      return {
        backgroundColor: edgeStyle.backgroundColor,
        bottom: edgeStyle.bottom,
        edge: edge.dataset.gridSelectionVisualEdge ?? '',
        height: edgeStyle.height,
        left: edgeStyle.left,
        right: edgeStyle.right,
        top: edgeStyle.top,
        width: edgeStyle.width,
      }
    })
    return {
      boxShadow: style.boxShadow,
      edges,
      outlineStyle: style.outlineStyle,
    }
  })
  expect(actual.boxShadow).toBe(expected.boxShadow)
  expect(actual.outlineStyle).toBe('none')
  expect(actual.edges).toHaveLength(4)
  expect(actual.edges.map((edge) => edge.edge)).toEqual(['top', 'right', 'bottom', 'left'])
  for (const edge of actual.edges) {
    expect(edge.backgroundColor).toBe(expected.color)
  }
  expect(actual.edges.find((edge) => edge.edge === 'top')).toMatchObject({
    height: expected.width,
    left: '-1px',
    right: '0px',
    top: '-1px',
  })
  expect(actual.edges.find((edge) => edge.edge === 'right')).toMatchObject({
    bottom: '1px',
    right: '0px',
    top: '0px',
    width: expected.width,
  })
  expect(actual.edges.find((edge) => edge.edge === 'bottom')).toMatchObject({
    bottom: '0px',
    height: expected.width,
    left: '-1px',
    right: '0px',
  })
  expect(actual.edges.find((edge) => edge.edge === 'left')).toMatchObject({
    bottom: '1px',
    left: '-1px',
    top: '0px',
    width: expected.width,
  })
}

export async function dragSelectedRangeBorderPreview(
  page: Page,
  startColumn: number,
  startRow: number,
  targetColumn: number,
  targetRow: number,
): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const sourceX = grid.x + startLeft + 3
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + 2
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
}

export async function getGridScrollTop(page: Page): Promise<number> {
  return await page.getByTestId('grid-scroll-viewport').evaluate((viewport) => {
    if (!(viewport instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return viewport.scrollTop
  })
}

export async function readProductViewportScroll(page: Page): Promise<{
  readonly scrollLeft: number
  readonly scrollTop: number
}> {
  await settleWorkbookScrollPerf(page, 8)
  return await page.getByTestId('grid-scroll-viewport').evaluate((node) => {
    if (!(node instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return {
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    }
  })
}

export async function clickVisibleScrolledBodyCell(page: Page, visibleColumnOffset: number, visibleRowOffset: number): Promise<string> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const scroll = await page.getByTestId('grid-scroll-viewport').evaluate((node) => {
    if (!(node instanceof HTMLDivElement)) {
      throw new Error('grid scroll viewport is not an HTMLDivElement')
    }
    return {
      scrollLeft: node.scrollLeft,
      scrollTop: node.scrollTop,
    }
  })
  const bodyX = visibleColumnOffset * PRODUCT_COLUMN_WIDTH + Math.floor(PRODUCT_COLUMN_WIDTH / 2)
  const bodyY = visibleRowOffset * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)
  const expectedColumnIndex = Math.floor((scroll.scrollLeft + bodyX) / PRODUCT_COLUMN_WIDTH)
  const expectedRowIndex = Math.floor((scroll.scrollTop + bodyY) / PRODUCT_ROW_HEIGHT)

  await gridLocator.click({
    position: {
      x: PRODUCT_ROW_MARKER_WIDTH + bodyX,
      y: PRODUCT_HEADER_HEIGHT + bodyY,
    },
  })

  return formatGridAddress(expectedRowIndex, expectedColumnIndex)
}

export function formatGridAddress(rowIndex: number, columnIndex: number): string {
  return `${formatColumnLabel(columnIndex)}${String(rowIndex + 1)}`
}

export function formatColumnLabel(columnIndex: number): string {
  let remaining = columnIndex + 1
  let label = ''
  while (remaining > 0) {
    const next = (remaining - 1) % 26
    label = String.fromCharCode(65 + next) + label
    remaining = Math.floor((remaining - 1) / 26)
  }
  return label
}

export async function rightClickProductRowHeader(page: Page, rowIndex: number): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const rowTop = await getProductRowTop(page, rowIndex)
  const rowHeight = await getProductRowHeight(page, rowIndex)
  await gridLocator.click({
    button: 'right',
    position: {
      x: Math.floor(PRODUCT_ROW_MARKER_WIDTH / 2),
      y: PRODUCT_HEADER_HEIGHT + rowTop + Math.floor(rowHeight / 2),
    },
  })
}

export async function rightClickProductColumnHeader(page: Page, columnIndex: number): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  await gridLocator.click({
    button: 'right',
    position: {
      x: columnLeft + Math.floor(columnWidth / 2),
      y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
    },
  })
}

export async function writeCellValue(page: Page, address: string, value: string): Promise<void> {
  await selectAddress(page, address)
  const formulaInput = page.getByTestId('formula-input')
  await formulaInput.fill(value)
  await formulaInput.press('Enter')
}

export async function dragProductSelectedInterior(
  page: Page,
  startColumn: number,
  startRow: number,
  targetColumn: number,
  targetRow: number,
): Promise<void> {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const startTop = await getProductRowTop(page, startRow)
  const startWidth = await getProductColumnWidth(page, startColumn)
  const startHeight = await getProductRowHeight(page, startRow)
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetTop = await getProductRowTop(page, targetRow)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetHeight = await getProductRowHeight(page, targetRow)

  await page.mouse.move(
    grid.x + startLeft + Math.min(32, Math.floor(startWidth * 0.35)),
    grid.y + PRODUCT_HEADER_HEIGHT + startTop + Math.floor(startHeight / 2),
  )
  await page.mouse.down()
  await page.mouse.move(
    grid.x + targetLeft + Math.floor(targetWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + targetTop + Math.floor(targetHeight / 2),
    { steps: 12 },
  )
  await page.mouse.up()
}

export async function selectAddress(page: Page, address: string): Promise<void> {
  const nameBox = page.getByTestId('name-box')
  await nameBox.fill(address)
  await expect(nameBox).toHaveValue(address)
  await nameBox.press('Enter')
  await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${address}`)
}

export async function readFormulaValue(page: Page): Promise<string> {
  const formulaInput = page.getByTestId('formula-input')
  return await formulaInput.inputValue()
}

export async function pressStructuralDeleteShortcut(page: Page): Promise<void> {
  await page.keyboard.down(PRIMARY_MODIFIER)
  await page.keyboard.down('Alt')
  await page.keyboard.press('Minus')
  await page.keyboard.up('Alt')
  await page.keyboard.up(PRIMARY_MODIFIER)
}
