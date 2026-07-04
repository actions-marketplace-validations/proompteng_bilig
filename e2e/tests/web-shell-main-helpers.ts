import { expect, type Locator, type test } from '@playwright/test'
import {
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  countDarkReadbackPixelsInCell,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductFillHandleDragPoints,
} from './web-shell-helpers.js'

export type BrowserSelectionAction =
  | { kind: 'click'; row: number; col: number }
  | { kind: 'shiftClick'; row: number; col: number }
  | { kind: 'key'; key: 'ArrowLeft' | 'ArrowRight' | 'ArrowUp' | 'ArrowDown'; shift: boolean }

export async function dragProductFillHandle(
  page: Parameters<typeof test>[0]['page'],
  sourceCol: number,
  sourceRow: number,
  targetCol: number,
  targetRow: number,
) {
  const { sourceX, sourceY, targetX, targetY } = await getProductFillHandleDragPoints(page, sourceCol, sourceRow, targetCol, targetRow)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, {
    steps: 10,
  })
  await page.mouse.up()
}

export async function clickSelectionFuzzCell(
  page: Parameters<typeof test>[0]['page'],
  columnIndex: number,
  rowIndex: number,
  shift = false,
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  const scrollLeft = await gridLocator.evaluate(
    (node, target) => {
      const scrollViewport = node.querySelector('[aria-hidden="true"]')
      if (!(scrollViewport instanceof HTMLElement)) {
        return 0
      }
      const targetCenter = target.columnLeft + target.columnWidth / 2
      const visibleStart = scrollViewport.scrollLeft
      const visibleEnd = visibleStart + scrollViewport.clientWidth
      if (targetCenter < visibleStart || targetCenter > visibleEnd) {
        scrollViewport.scrollLeft = Math.max(0, targetCenter - scrollViewport.clientWidth / 2)
      }
      return scrollViewport.scrollLeft
    },
    {
      columnLeft,
      columnWidth,
    },
  )
  const point = {
    x: grid.x + columnLeft - scrollLeft + columnWidth / 2,
    y: grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT + PRODUCT_ROW_HEIGHT / 2,
  }
  if (shift) {
    await page.keyboard.down('Shift')
  }
  try {
    await page.mouse.click(point.x, point.y)
  } finally {
    if (shift) {
      await page.keyboard.up('Shift')
    }
  }
}

export async function runSelectionFuzzActions(
  page: Parameters<typeof test>[0]['page'],
  grid: Locator,
  actions: readonly BrowserSelectionAction[],
  index = 0,
): Promise<void> {
  const action = actions[index]
  if (!action) {
    return
  }

  if (action.kind === 'click') {
    await clickSelectionFuzzCell(page, action.col, action.row)
  } else if (action.kind === 'shiftClick') {
    await clickSelectionFuzzCell(page, action.col, action.row, true)
  } else {
    await grid.press(action.shift ? `Shift+${action.key}` : action.key)
  }

  const selection = await page.getByTestId('status-selection').textContent()
  expect(selection).toMatch(/^Sheet1!(?:[A-Z]+[0-9]+(?::[A-Z]+[0-9]+)?|[A-Z]+:[A-Z]+|[0-9]+:[0-9]+|All)$/)

  const focusInsideShell = await page.evaluate(() => {
    const active = document.activeElement
    return Boolean(
      active?.closest('[data-testid="sheet-grid"]') ||
      active?.closest('[data-testid="formula-bar"]') ||
      active?.closest('[role="toolbar"]'),
    )
  })
  expect(focusInsideShell).toBe(true)

  await runSelectionFuzzActions(page, grid, actions, index + 1)
}

export async function clickProductSelectedCellTopBorder(page: Parameters<typeof test>[0]['page'], columnIndex: number, rowIndex: number) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const columnLeft = await getProductColumnLeft(page, columnIndex)
  const columnWidth = await getProductColumnWidth(page, columnIndex)
  await page.mouse.click(
    grid.x + columnLeft + Math.floor(columnWidth / 2),
    grid.y + PRODUCT_HEADER_HEIGHT + rowIndex * PRODUCT_ROW_HEIGHT - 1,
  )
}

export async function nativeTextRunsInclude(page: Parameters<typeof test>[0]['page'], text: string): Promise<boolean> {
  return await page.evaluate(
    (needle) => Array.from(document.querySelectorAll('[data-native-text-run]')).some((run) => run.textContent?.includes(needle) ?? false),
    text,
  )
}

export async function expectCellRenderedText(
  page: Parameters<typeof test>[0]['page'],
  columnIndex: number,
  rowIndex: number,
  text: string,
  expected: 'hidden' | 'visible',
): Promise<void> {
  const hasTypeGpuCanvas = (await page.getByTestId('grid-pane-renderer').count()) > 0
  if (!hasTypeGpuCanvas) {
    await expect.poll(() => nativeTextRunsInclude(page, text)).toBe(expected === 'visible')
    return
  }
  const poll = expect.poll(async () => await countDarkReadbackPixelsInCell(page, columnIndex, rowIndex), {
    message: `cell ${columnIndex}:${rowIndex} rendered text should be ${expected}: ${text}`,
  })
  if (expected === 'visible') {
    await poll.toBeGreaterThan(4)
    return
  }
  await poll.toBeLessThanOrEqual(2)
}

export async function textControlValue(locator: Locator): Promise<string> {
  return await locator.evaluate((control) =>
    control instanceof HTMLInputElement || control instanceof HTMLTextAreaElement ? control.value : '',
  )
}

export async function dragProductSelectionBorder(
  page: Parameters<typeof test>[0]['page'],
  startColumn: number,
  startRow: number,
  endColumn: number,
  endRow: number,
  targetColumn: number,
  targetRow: number,
) {
  const gridLocator = page.getByTestId('sheet-grid')
  await expect(gridLocator).toBeVisible()
  const grid = await gridLocator.boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }

  const startLeft = await getProductColumnLeft(page, startColumn)
  const rangeTop = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT
  const sourceX = grid.x + startLeft + 3
  const sourceY = rangeTop + 2
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
  await page.mouse.up()
}
