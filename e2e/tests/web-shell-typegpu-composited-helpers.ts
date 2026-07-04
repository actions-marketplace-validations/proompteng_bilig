import { expect, type Page, type TestInfo } from '@playwright/test'
import { writeFile } from 'node:fs/promises'
import {
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  getProductColumnLeft,
  getProductColumnWidth,
  getProductRowTop,
} from './web-shell-helpers.js'
import type { ReadbackPoint } from './web-shell-typegpu-helpers.js'

export async function sampleCompositedGridLocalPixel(
  page: Page,
  point: { readonly x: number; readonly y: number },
): Promise<ReadbackPoint> {
  const grid = await page.getByTestId('sheet-grid').boundingBox()
  if (!grid) {
    throw new Error('sheet grid is not visible')
  }
  const restoreReadbackCanvas = await page.evaluate(() => {
    const canvas = document.getElementById('gpu-readback-canvas')
    if (!(canvas instanceof HTMLElement)) {
      return null
    }
    const previousVisibility = canvas.style.visibility
    canvas.style.visibility = 'hidden'
    return previousVisibility
  })
  let buffer: Buffer
  try {
    buffer = await page.screenshot({
      animations: 'disabled',
      caret: 'hide',
      clip: {
        height: 1,
        width: 1,
        x: Math.round(grid.x + point.x),
        y: Math.round(grid.y + point.y),
      },
    })
  } finally {
    if (restoreReadbackCanvas !== null) {
      await page.evaluate((previousVisibility) => {
        const canvas = document.getElementById('gpu-readback-canvas')
        if (canvas instanceof HTMLElement) {
          canvas.style.visibility = previousVisibility
        }
      }, restoreReadbackCanvas)
    }
  }
  return await page.evaluate(
    async ({ dataUrl }) => {
      const image = await new Promise<HTMLImageElement>((resolve, reject) => {
        const element = new Image()
        element.addEventListener('load', () => resolve(element), { once: true })
        element.addEventListener('error', () => reject(new Error('Failed to decode composited grid screenshot')), { once: true })
        element.src = dataUrl
      })
      const canvas = document.createElement('canvas')
      canvas.width = image.naturalWidth
      canvas.height = image.naturalHeight
      const context = canvas.getContext('2d')
      if (!context) {
        throw new Error('Missing 2d context for composited grid screenshot analysis')
      }
      context.drawImage(image, 0, 0)
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data
      return {
        a: pixels[3] ?? 0,
        b: pixels[2] ?? 0,
        g: pixels[1] ?? 0,
        r: pixels[0] ?? 0,
      }
    },
    { dataUrl: `data:image/png;base64,${buffer.toString('base64')}` },
  )
}

export async function productCellInteriorGridLocalPoint(
  page: Page,
  columnIndex: number,
  rowIndex: number,
): Promise<{ readonly x: number; readonly y: number }> {
  const [columnLeft, columnWidth, rowTop] = await Promise.all([
    getProductColumnLeft(page, columnIndex),
    getProductColumnWidth(page, columnIndex),
    getProductRowTop(page, rowIndex),
  ])
  return {
    x: columnLeft + Math.floor(columnWidth / 2),
    y: PRODUCT_HEADER_HEIGHT + rowTop + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }
}

export async function productColumnHeaderGridLocalPoint(
  page: Page,
  columnIndex: number,
): Promise<{ readonly x: number; readonly y: number }> {
  const columnLeft = await getProductColumnLeft(page, columnIndex)
  return {
    x: columnLeft + 8,
    y: Math.floor(PRODUCT_HEADER_HEIGHT / 2),
  }
}

export async function waitForCompositedGridLocalPixel(
  page: Page,
  point: { readonly x: number; readonly y: number },
  predicate: (point: ReadbackPoint) => boolean,
): Promise<ReadbackPoint> {
  let lastPoint: ReadbackPoint | null = null
  try {
    await expect
      .poll(
        async () => {
          lastPoint = await sampleCompositedGridLocalPixel(page, point)
          return predicate(lastPoint)
        },
        { timeout: 5_000 },
      )
      .toBe(true)
  } catch (error) {
    throw new Error(`${error instanceof Error ? error.message : String(error)}\nLast composited pixel: ${JSON.stringify(lastPoint)}`, {
      cause: error,
    })
  }
  if (!lastPoint) {
    throw new Error('expected composited grid pixel')
  }
  return lastPoint
}

export async function saveReadbackArtifact(page: Page, testInfo: TestInfo, fileName: string, attachmentName: string): Promise<void> {
  const outputPath = testInfo.outputPath(fileName)
  const dataUrl = await page.evaluate(() => {
    const canvas = document.getElementById('gpu-readback-canvas')
    if (!(canvas instanceof HTMLCanvasElement)) {
      return null
    }
    return canvas.toDataURL('image/png')
  })
  if (!dataUrl) {
    throw new Error('gpu readback canvas unavailable')
  }
  await writeFile(outputPath, dataUrl.replace(/^data:image\/png;base64,/, ''), 'base64')
  await testInfo.attach(attachmentName, {
    path: outputPath,
    contentType: 'image/png',
  })
}
