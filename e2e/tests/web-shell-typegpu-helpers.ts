import { expect, type Locator, type Page } from '@playwright/test'
import {
  PRODUCT_COLUMN_WIDTH,
  PRODUCT_HEADER_HEIGHT,
  PRODUCT_ROW_HEIGHT,
  PRODUCT_ROW_MARKER_WIDTH,
  clickProductCell,
  getProductColumnLeft,
  getProductColumnWidth,
} from './web-shell-helpers.js'

export interface ReadbackPoint {
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

export interface TypeGpuReadbackSummary {
  readonly ready: boolean
  readonly hasGpu: boolean
  readonly width: number
  readonly height: number
  readonly sequence: number
  readonly points: {
    readonly headerFill: ReadbackPoint
    readonly bodyFill: ReadbackPoint
    readonly selectionBorder: ReadbackPoint
    readonly selectionFill: ReadbackPoint
    readonly valueFill: ReadbackPoint
    readonly bodyWhite: ReadbackPoint
  }
  readonly darkPixelCounts: {
    readonly header: number
    readonly body: number
    readonly number: number
  }
}

interface ReadbackInspectorPoint {
  readonly name: string
  readonly x: number
  readonly y: number
}

interface ReadbackInspectorRegion {
  readonly name: string
  readonly x0: number
  readonly y0: number
  readonly x1: number
  readonly y1: number
  readonly threshold?: number
}

interface DynamicReadbackResult {
  readonly ready: boolean
  readonly hasGpu: boolean
  readonly width: number
  readonly height: number
  readonly sequence: number
  readonly points: Record<string, ReadbackPoint>
  readonly darkPixelCounts: Record<string, number>
  readonly opaquePixelCounts: Record<string, number>
}

export function selectedRangeFillProbe(columnIndex: number, rowIndex: number): ReadbackInspectorPoint {
  return {
    name: `${columnIndex}:${rowIndex}`,
    x: PRODUCT_ROW_MARKER_WIDTH + PRODUCT_COLUMN_WIDTH * columnIndex + Math.floor(PRODUCT_COLUMN_WIDTH / 2),
    y: PRODUCT_HEADER_HEIGHT + PRODUCT_ROW_HEIGHT * rowIndex + Math.floor(PRODUCT_ROW_HEIGHT / 2),
  }
}

export async function expectSelectionVisualOpacity(locator: Locator, label: string, expected: 'hidden' | 'visible'): Promise<void> {
  await expect(locator, `${label} should exist`).not.toHaveCount(0)
  const opacities = await locator.evaluateAll((nodes) => nodes.map((node) => window.getComputedStyle(node).opacity))
  expect(
    opacities.every((opacity) => (expected === 'visible' ? opacity !== '0' : opacity === '0')),
    `${label} should be ${expected} in the DOM selection overlay`,
  ).toBe(true)
}

export async function expectPresentedOverlayRects(page: Page, label: string): Promise<void> {
  await expect
    .poll(
      async () => {
        const value = await page.getByTestId('grid-pane-renderer').getAttribute('data-v3-presented-overlay-rect-count')
        return Number(value ?? '0')
      },
      { message: `${label} should be presented by the TypeGPU overlay`, timeout: 5_000 },
    )
    .toBeGreaterThan(0)
}

export function allReadbackPointsMatch(result: DynamicReadbackResult, predicate: (point: ReadbackPoint) => boolean): boolean {
  return Object.values(result.points).every(predicate)
}

export function isCornflowerBlueFill(point: ReadbackPoint): boolean {
  return point.a === 255 && point.b > point.g && point.g > point.r && point.b - point.r > 55
}

export function isSelectionGreenTint(point: ReadbackPoint): boolean {
  return point.a === 255 && point.g > point.r && point.g > point.b && point.g - point.r >= 8
}

export function isThemeGreenFill(point: ReadbackPoint): boolean {
  return point.a === 255 && point.g > point.r + 45 && point.g > point.b + 25
}

export function isWorkbookWhitePixel(point: ReadbackPoint): boolean {
  return point.a === 255 && point.r >= 245 && point.g >= 245 && point.b >= 245
}

export function expectWorkbookWhitePixel(point: ReadbackPoint): void {
  expect(point.a).toBe(255)
  expect(point.r).toBeGreaterThanOrEqual(245)
  expect(point.g).toBeGreaterThanOrEqual(245)
  expect(point.b).toBeGreaterThanOrEqual(245)
}

interface NativeTextExpectation {
  readonly name: string
  readonly text: string
  readonly exact?: boolean | undefined
}

interface NativeTextLayerInspection {
  readonly gridAttrs: Record<string, string>
  readonly matches: Record<string, boolean>
  readonly rendererAttrs: Record<string, string>
  readonly rowHeaderRunCount: number
  readonly sampleTexts: readonly string[]
  readonly visibleRunCount: number
}

interface RendererPresentationSample {
  readonly backendStatus: string | null
  readonly canvasProofLayer: string | null
  readonly editorInputs: number
  readonly editorOverlays: number
  readonly fallbackCanvases: number
  readonly frameProofStatus: string | null
  readonly headerPaneCount: number
  readonly hasPresentedFrame: string | null
  readonly hasPresentedVisibleFrame: string | null
  readonly headerTextRunCount: number
  readonly nativeLayerSource: string | null
  readonly nativeTextLayerMounted: boolean
  readonly nativeTextRunCount: number
  readonly presentedVisibleTextRunCount: number
  readonly presentedTextRunCount: number
  readonly rowHeaderRunCount: number
  readonly selection: string | null
  readonly typeGpuDrawText: string | null
}

export function isResizeGuidePixel(point: ReadbackPoint): boolean {
  const isLegacyBlueGuide = point.b >= 210 && point.g >= 95 && point.g <= 140 && point.r <= 50
  const isWorkbookGreenGuide = point.g >= 80 && point.g > point.r + 35 && point.g > point.b + 20 && point.r <= 70 && point.b <= 105
  return point.a === 255 && (isLegacyBlueGuide || isWorkbookGreenGuide)
}

export function expectNear(actual: number, expected: number, tolerance: number): void {
  expect(Math.abs(actual - expected)).toBeLessThanOrEqual(tolerance)
}

export async function waitForTypeGpuRenderer(page: Page): Promise<void> {
  await page.waitForSelector('[data-testid="grid-pane-renderer"]', { state: 'attached', timeout: 15_000 })
}

export async function collectRendererPresentationSamplesDuring(
  page: Page,
  action: () => Promise<void>,
  sampleCount = 24,
): Promise<readonly RendererPresentationSample[]> {
  const samplesPromise = page.evaluate(
    async ({ count, rowMarkerWidth }) => {
      const samples: RendererPresentationSample[] = []
      await new Promise<void>((resolve) => {
        let index = 0
        const sampleNextFrame = () => {
          const canvas = document.querySelector('[data-testid="grid-pane-renderer"]')
          const grid = document.querySelector('[data-testid="sheet-grid"]')
          const gridRect = grid instanceof HTMLElement ? grid.getBoundingClientRect() : null
          const nativeTextLayer = document.querySelector('[data-testid="grid-native-text-layer"]')
          const nativeTextRuns = [...document.querySelectorAll('[data-native-text-run]')]
          const readNumberAttribute = (name: string) => Number(canvas?.getAttribute(name) ?? '0')
          const rowHeaderRunCount = nativeTextRuns.filter((run) => {
            if (!gridRect || !/^\d+$/.test(run.textContent ?? '')) {
              return false
            }
            const rect = run.getBoundingClientRect()
            return rect.left >= gridRect.left && rect.right <= gridRect.left + rowMarkerWidth + 1
          }).length
          samples.push({
            backendStatus: canvas?.getAttribute('data-v3-backend-status') ?? null,
            canvasProofLayer: canvas?.getAttribute('data-v3-canvas-proof-layer') ?? null,
            editorInputs: document.querySelectorAll('[data-testid="cell-editor-input"]').length,
            editorOverlays: document.querySelectorAll('[data-testid="cell-editor-overlay"]').length,
            fallbackCanvases: document.querySelectorAll('[data-testid="grid-pane-renderer-fallback"]').length,
            frameProofStatus: canvas?.getAttribute('data-v3-frame-proof-status') ?? null,
            headerPaneCount: Number(canvas?.getAttribute('data-v3-header-pane-count') ?? '0'),
            headerTextRunCount: readNumberAttribute('data-v3-header-text-run-count'),
            hasPresentedFrame: canvas?.getAttribute('data-v3-has-presented-frame') ?? null,
            hasPresentedVisibleFrame: canvas?.getAttribute('data-v3-has-presented-visible-frame') ?? null,
            nativeLayerSource: canvas?.getAttribute('data-v3-native-layer-source') ?? null,
            nativeTextLayerMounted: nativeTextLayer instanceof HTMLElement,
            nativeTextRunCount: nativeTextRuns.length,
            presentedVisibleTextRunCount: readNumberAttribute('data-v3-presented-visible-text-run-count'),
            presentedTextRunCount: readNumberAttribute('data-v3-presented-text-run-count'),
            rowHeaderRunCount,
            selection: document.querySelector('[data-testid="status-selection"]')?.textContent ?? null,
            typeGpuDrawText: canvas?.getAttribute('data-v3-draw-text') ?? null,
          })
          index += 1
          if (index >= count) {
            resolve()
            return
          }
          requestAnimationFrame(sampleNextFrame)
        }
        requestAnimationFrame(sampleNextFrame)
      })
      return samples
    },
    { count: sampleCount, rowMarkerWidth: PRODUCT_ROW_MARKER_WIDTH },
  )
  await action()
  return samplesPromise
}

export async function exerciseClickAwayEditCommit(
  page: Page,
  input: {
    readonly address: string
    readonly awayCol: number
    readonly awayRow: number
    readonly awaySelection: string
    readonly col: number
    readonly formulaInput: Locator
    readonly renderer: Locator
    readonly row: number
    readonly text: string
  },
): Promise<void> {
  await clickProductCell(page, input.col, input.row)
  await expect(page.getByTestId('status-selection')).toHaveText(`Sheet1!${input.address}`)

  const firstCharacter = input.text.charAt(0)
  await page.keyboard.press(firstCharacter)

  const cellEditor = page.getByTestId('cell-editor-input')
  await expect(cellEditor).toBeVisible()
  await expect(cellEditor).toHaveValue(firstCharacter)
  await expect(page.getByTestId('cell-editor-overlay')).toHaveCount(1)

  await page.keyboard.type(input.text.slice(1))
  await expect(cellEditor).toHaveValue(input.text)
  await expect
    .poll(async () => await cellEditor.evaluate((editor) => (editor instanceof HTMLTextAreaElement ? editor.selectionStart : -1)))
    .toBe(input.text.length)

  const samples = await collectRendererPresentationSamplesDuring(page, () => clickProductCell(page, input.awayCol, input.awayRow))
  expect(samples.every((sample) => sample.editorInputs <= 1)).toBe(true)
  expect(samples.every((sample) => sample.editorOverlays <= 1)).toBe(true)
  expect(samples.every((sample) => sample.headerPaneCount > 0)).toBe(true)
  expect(samples.every((sample) => sample.nativeLayerSource === 'browser-native-text-live')).toBe(true)
  expect(samples.every((sample) => sample.nativeTextLayerMounted)).toBe(true)
  expect(samples.every((sample) => sample.nativeTextRunCount > 0 || sample.presentedVisibleTextRunCount > 0)).toBe(true)
  expect(samples.every((sample) => sample.typeGpuDrawText === 'false')).toBe(true)
  expect(samples.every((sample) => sample.headerTextRunCount > 0)).toBe(true)
  expect(samples.filter((sample) => sample.fallbackCanvases !== 0)).toEqual([])
  expect(samples.filter((sample) => sample.canvasProofLayer !== 'disabled')).toEqual([])

  await expect(page.getByTestId('status-selection')).toHaveText(input.awaySelection)
  await expect(cellEditor).toHaveCount(0)
  await expect(page.getByTestId('cell-editor-overlay')).toHaveCount(0)
  await expect(page.getByTestId('grid-pane-renderer-fallback')).toHaveCount(0)

  await clickProductCell(page, input.col, input.row)
  await expect(input.formulaInput).toHaveValue(input.text)
  await expect.poll(async () => await input.renderer.getAttribute('data-v3-frame-proof-status')).toBe('presented')
  await expect(input.renderer).toHaveAttribute('data-v3-canvas-proof-layer', 'disabled')
}

export async function dragProductSelectionBorder(
  page: Page,
  startColumn: number,
  startRow: number,
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
  const sourceX = grid.x + startLeft + 3
  const sourceY = grid.y + PRODUCT_HEADER_HEIGHT + startRow * PRODUCT_ROW_HEIGHT + 2
  const targetLeft = await getProductColumnLeft(page, targetColumn)
  const targetWidth = await getProductColumnWidth(page, targetColumn)
  const targetX = grid.x + targetLeft + Math.floor(targetWidth / 2)
  const targetY = grid.y + PRODUCT_HEADER_HEIGHT + targetRow * PRODUCT_ROW_HEIGHT + Math.floor(PRODUCT_ROW_HEIGHT / 2)

  await page.mouse.move(sourceX, sourceY)
  await page.mouse.down()
  await page.mouse.move(targetX, targetY, { steps: 12 })
  await page.mouse.up()
}

export async function installTypeGpuReadbackHarness(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const globalWindow = window as Window & {
      __biligGpuReadbackState?: {
        bgra: Uint8Array
        bytesPerRow: number
        hasGpu: boolean
        height: number
        ready: boolean
        sequence: number
        width: number
      }
      __biligGpuReadback?: TypeGpuReadbackSummary
      __biligTypeGpuHarnessInstalled?: boolean
      __biligGpuReadbackInspector?: {
        readonly isReady: () => boolean
        readonly getSequence: () => number
        readonly getSize: () => { readonly width: number; readonly height: number }
        readonly samplePoints: (
          points: readonly { readonly name: string; readonly x: number; readonly y: number }[],
        ) => Record<string, ReadbackPoint>
        readonly countDarkPixels: (
          regions: readonly {
            readonly name: string
            readonly x0: number
            readonly y0: number
            readonly x1: number
            readonly y1: number
            readonly threshold?: number
          }[],
        ) => Record<string, number>
        readonly countOpaquePixels: (
          regions: readonly {
            readonly name: string
            readonly x0: number
            readonly y0: number
            readonly x1: number
            readonly y1: number
            readonly threshold?: number
          }[],
        ) => Record<string, number>
      }
    }

    const readbackState = globalWindow.__biligGpuReadbackState ?? {
      bgra: new Uint8Array(0),
      bytesPerRow: 0,
      hasGpu: Boolean(navigator.gpu),
      height: 0,
      ready: false,
      sequence: 0,
      width: 0,
    }
    globalWindow.__biligGpuReadbackState = readbackState
    readbackState.bgra = new Uint8Array(0)
    readbackState.bytesPerRow = 0
    readbackState.hasGpu = Boolean(navigator.gpu)
    readbackState.height = 0
    readbackState.ready = false
    readbackState.sequence = 0
    readbackState.width = 0

    const pointAt = (x: number, y: number): ReadbackPoint => {
      if (!readbackState.ready || x < 0 || y < 0 || x >= readbackState.width || y >= readbackState.height) {
        return { r: 0, g: 0, b: 0, a: 0 }
      }
      const offset = y * readbackState.bytesPerRow + x * 4
      return {
        r: readbackState.bgra[offset + 2] ?? 0,
        g: readbackState.bgra[offset + 1] ?? 0,
        b: readbackState.bgra[offset + 0] ?? 0,
        a: readbackState.bgra[offset + 3] ?? 0,
      }
    }

    const countDarkPixels = (x0: number, y0: number, x1: number, y1: number, threshold = 120): number => {
      let count = 0
      for (let y = Math.max(0, y0); y < Math.min(readbackState.height, y1); y += 1) {
        for (let x = Math.max(0, x0); x < Math.min(readbackState.width, x1); x += 1) {
          const point = pointAt(x, y)
          if (point.a > 0 && point.r < threshold && point.g < threshold && point.b < threshold) {
            count += 1
          }
        }
      }
      return count
    }

    const countOpaquePixels = (x0: number, y0: number, x1: number, y1: number, threshold = 1): number => {
      let count = 0
      for (let y = Math.max(0, y0); y < Math.min(readbackState.height, y1); y += 1) {
        for (let x = Math.max(0, x0); x < Math.min(readbackState.width, x1); x += 1) {
          if (pointAt(x, y).a >= threshold) {
            count += 1
          }
        }
      }
      return count
    }

    globalWindow.__biligGpuReadbackInspector = {
      countDarkPixels(regions) {
        return Object.fromEntries(
          regions.map((region) => [region.name, countDarkPixels(region.x0, region.y0, region.x1, region.y1, region.threshold)]),
        )
      },
      countOpaquePixels(regions) {
        return Object.fromEntries(
          regions.map((region) => [region.name, countOpaquePixels(region.x0, region.y0, region.x1, region.y1, region.threshold)]),
        )
      },
      getSequence() {
        return readbackState.sequence
      },
      getSize() {
        return { height: readbackState.height, width: readbackState.width }
      },
      isReady() {
        return readbackState.ready
      },
      samplePoints(points) {
        return Object.fromEntries(points.map((point) => [point.name, pointAt(point.x, point.y)]))
      },
    }

    globalWindow.__biligGpuReadback = {
      ready: readbackState.ready,
      hasGpu: readbackState.hasGpu,
      width: readbackState.width,
      height: readbackState.height,
      sequence: readbackState.sequence,
      points: {
        headerFill: { r: 0, g: 0, b: 0, a: 0 },
        bodyFill: { r: 0, g: 0, b: 0, a: 0 },
        selectionBorder: { r: 0, g: 0, b: 0, a: 0 },
        selectionFill: { r: 0, g: 0, b: 0, a: 0 },
        valueFill: { r: 0, g: 0, b: 0, a: 0 },
        bodyWhite: { r: 0, g: 0, b: 0, a: 0 },
      },
      darkPixelCounts: {
        header: 0,
        body: 0,
        number: 0,
      },
    }

    if (globalWindow.__biligTypeGpuHarnessInstalled) {
      return
    }

    if (!navigator.gpu) {
      return
    }

    globalWindow.__biligTypeGpuHarnessInstalled = true

    const functionKind = 'function'
    const isCanvasContextConfigure = (value: unknown): value is (this: GPUCanvasContext, descriptor: GPUCanvasConfiguration) => void =>
      typeof value === functionKind
    const isCanvasContextGetCurrentTexture = (value: unknown): value is (this: GPUCanvasContext) => GPUTexture =>
      typeof value === functionKind
    const readbackCanvasId = 'gpu-readback-canvas'
    const originalConfigure = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'configure')?.value
    if (!isCanvasContextConfigure(originalConfigure)) {
      return
    }
    GPUCanvasContext.prototype.configure = function configureWithCopySrc(descriptor: GPUCanvasConfiguration) {
      return originalConfigure.call(this, {
        ...descriptor,
        usage: (descriptor.usage ?? GPUTextureUsage.RENDER_ATTACHMENT) | GPUTextureUsage.COPY_SRC,
      })
    }

    const originalRequestAdapter = navigator.gpu.requestAdapter.bind(navigator.gpu)
    navigator.gpu.requestAdapter = async (...adapterArgs) => {
      const adapter = await originalRequestAdapter(...adapterArgs)
      if (!adapter) {
        return adapter
      }

      const originalRequestDevice = adapter.requestDevice.bind(adapter)
      adapter.requestDevice = async (...deviceArgs) => {
        const device = await originalRequestDevice(...deviceArgs)
        let lastTexture: GPUTexture | null = null
        let lastWidth = 0
        let lastHeight = 0
        let lastFallbackTexture: GPUTexture | null = null
        let lastFallbackWidth = 0
        let lastFallbackHeight = 0
        let readbackSerial = 0
        let committedReadbackSerial = 0

        const originalGetCurrentTexture = Object.getOwnPropertyDescriptor(GPUCanvasContext.prototype, 'getCurrentTexture')?.value
        if (!isCanvasContextGetCurrentTexture(originalGetCurrentTexture)) {
          return device
        }
        GPUCanvasContext.prototype.getCurrentTexture = function recordCurrentTexture() {
          const texture = originalGetCurrentTexture.call(this)
          if (this.canvas instanceof HTMLCanvasElement) {
            const testId = this.canvas.getAttribute('data-testid')
            const dataPaneRenderer = this.canvas.getAttribute('data-pane-renderer')
            const tracked = testId === 'grid-pane-renderer' || dataPaneRenderer === 'workbook-pane-renderer'
            lastFallbackTexture = texture
            lastFallbackWidth = this.canvas.width
            lastFallbackHeight = this.canvas.height
            if (!tracked) {
              return texture
            }
            lastTexture = texture
            lastWidth = this.canvas.width
            lastHeight = this.canvas.height
          }
          return texture
        }

        const originalSubmit = device.queue.submit.bind(device.queue)
        device.queue.submit = (buffers: Iterable<GPUCommandBuffer>) => {
          const commandBuffers = Array.from(buffers)
          const targetTexture = lastTexture ?? lastFallbackTexture
          const targetWidth = lastTexture ? lastWidth : lastFallbackWidth
          const targetHeight = lastTexture ? lastHeight : lastFallbackHeight
          if (targetTexture && targetWidth > 0 && targetHeight > 0) {
            readbackSerial += 1
            const serial = readbackSerial
            const bytesPerRow = Math.ceil((targetWidth * 4) / 256) * 256
            const buffer = device.createBuffer({
              size: bytesPerRow * targetHeight,
              usage: GPUBufferUsage.COPY_DST | GPUBufferUsage.MAP_READ,
            })
            const encoder = device.createCommandEncoder()
            encoder.copyTextureToBuffer(
              { texture: targetTexture },
              { buffer, bytesPerRow, rowsPerImage: targetHeight },
              { width: targetWidth, height: targetHeight, depthOrArrayLayers: 1 },
            )
            const result = originalSubmit([...commandBuffers, encoder.finish()])
            void buffer
              .mapAsync(GPUMapMode.READ)
              .then(() => {
                const mapped = new Uint8Array(buffer.getMappedRange())
                if (serial <= committedReadbackSerial) {
                  return mapped
                }
                const bgra = new Uint8Array(mapped)
                committedReadbackSerial = serial
                readbackState.bgra = bgra
                readbackState.bytesPerRow = bytesPerRow
                readbackState.hasGpu = true
                readbackState.height = targetHeight
                readbackState.ready = true
                readbackState.sequence += 1
                readbackState.width = targetWidth
                globalWindow.__biligGpuReadback = buildReadbackSummary({
                  width: targetWidth,
                  height: targetHeight,
                  bytesPerRow,
                  bgra,
                  hasGpu: true,
                  sequence: readbackState.sequence,
                })
                renderReadbackCanvas({ width: targetWidth, height: targetHeight, bytesPerRow, bgra })
                return bgra
              })
              .finally(() => {
                try {
                  buffer.unmap()
                } catch (error) {
                  console.warn('Ignoring GPU buffer unmap error', error)
                }
                buffer.destroy()
              })
            return result
          }
          return originalSubmit(commandBuffers)
        }

        return device
      }

      return adapter
    }

    function buildReadbackSummary(input: {
      readonly width: number
      readonly height: number
      readonly bytesPerRow: number
      readonly bgra: Uint8Array
      readonly hasGpu: boolean
      readonly sequence: number
    }): TypeGpuReadbackSummary {
      const samplePoint = (x: number, y: number): ReadbackPoint => {
        const offset = y * input.bytesPerRow + x * 4
        return {
          r: input.bgra[offset + 2] ?? 0,
          g: input.bgra[offset + 1] ?? 0,
          b: input.bgra[offset + 0] ?? 0,
          a: input.bgra[offset + 3] ?? 0,
        }
      }

      const sampleDarkPixels = (x0: number, y0: number, x1: number, y1: number): number => {
        let count = 0
        for (let y = y0; y < y1; y += 1) {
          for (let x = x0; x < x1; x += 1) {
            const point = samplePoint(x, y)
            if (point.a > 0 && point.r < 120 && point.g < 120 && point.b < 120) {
              count += 1
            }
          }
        }
        return count
      }

      return {
        ready: true,
        hasGpu: input.hasGpu,
        width: input.width,
        height: input.height,
        sequence: input.sequence,
        points: {
          headerFill: samplePoint(20, 12),
          bodyFill: samplePoint(60, 40),
          selectionBorder: samplePoint(200, 68),
          selectionFill: samplePoint(260, 100),
          valueFill: samplePoint(520, 140),
          bodyWhite: samplePoint(400, 300),
        },
        darkPixelCounts: {
          header: sampleDarkPixels(80, 4, 120, 18),
          body: sampleDarkPixels(58, 48, 110, 66),
          number: sampleDarkPixels(532, 48, 620, 70),
        },
      }
    }

    function renderReadbackCanvas(input: {
      readonly width: number
      readonly height: number
      readonly bytesPerRow: number
      readonly bgra: Uint8Array
    }): void {
      const existing = globalWindow.document.getElementById(readbackCanvasId)
      existing?.remove()

      const rgba = new Uint8ClampedArray(input.width * input.height * 4)
      for (let y = 0; y < input.height; y += 1) {
        const rowOffset = y * input.bytesPerRow
        for (let x = 0; x < input.width; x += 1) {
          const src = rowOffset + x * 4
          const dst = (y * input.width + x) * 4
          rgba[dst + 0] = input.bgra[src + 2] ?? 0
          rgba[dst + 1] = input.bgra[src + 1] ?? 0
          rgba[dst + 2] = input.bgra[src + 0] ?? 0
          rgba[dst + 3] = input.bgra[src + 3] ?? 0
        }
      }

      const canvas = globalWindow.document.createElement('canvas')
      canvas.id = readbackCanvasId
      canvas.width = input.width
      canvas.height = input.height
      canvas.style.position = 'fixed'
      canvas.style.left = '0'
      canvas.style.top = '0'
      canvas.style.zIndex = '99999'
      canvas.style.pointerEvents = 'none'
      const context = canvas.getContext('2d')
      if (!context) {
        return
      }
      context.putImageData(new ImageData(rgba, input.width, input.height), 0, 0)
      globalWindow.document.body.appendChild(canvas)
    }
  })
}

export async function inspectGpuReadback(
  page: Page,
  input: {
    readonly points: readonly ReadbackInspectorPoint[]
    readonly regions: readonly ReadbackInspectorRegion[]
  },
): Promise<DynamicReadbackResult> {
  const result = await page.evaluate(({ points, regions }) => {
    const inspector = (
      window as Window & {
        __biligGpuReadbackInspector?: {
          readonly isReady: () => boolean
          readonly getSequence: () => number
          readonly getSize: () => { readonly width: number; readonly height: number }
          readonly samplePoints: (points: readonly ReadbackInspectorPoint[]) => Record<string, ReadbackPoint>
          readonly countDarkPixels: (regions: readonly ReadbackInspectorRegion[]) => Record<string, number>
          readonly countOpaquePixels: (regions: readonly ReadbackInspectorRegion[]) => Record<string, number>
        }
        __biligGpuReadback?: { readonly hasGpu: boolean }
      }
    ).__biligGpuReadbackInspector
    const hasGpu = Boolean((window as Window & { __biligGpuReadback?: { readonly hasGpu: boolean } }).__biligGpuReadback?.hasGpu)

    if (!inspector) {
      return {
        ready: false,
        hasGpu,
        width: 0,
        height: 0,
        sequence: 0,
        points: {},
        darkPixelCounts: {},
        opaquePixelCounts: {},
      }
    }

    const size = inspector.getSize()
    return {
      ready: inspector.isReady(),
      hasGpu,
      width: size.width,
      height: size.height,
      sequence: inspector.getSequence(),
      points: inspector.samplePoints(points),
      darkPixelCounts: inspector.countDarkPixels(regions),
      opaquePixelCounts: inspector.countOpaquePixels(regions),
    }
  }, input)

  expect(result.ready).toBe(true)
  return result
}

async function inspectVisibleNativeTextRuns(
  page: Page,
  textExpectations: readonly NativeTextExpectation[],
): Promise<NativeTextLayerInspection> {
  return await page.evaluate(
    (payload) => {
      const sheetGrid = document.querySelector('[data-testid="sheet-grid"]')
      const renderer = document.querySelector('[data-testid="grid-pane-renderer"]')
      const gridRect = sheetGrid instanceof HTMLElement ? sheetGrid.getBoundingClientRect() : null
      const nativeTextRuns = [...document.querySelectorAll<HTMLElement>('[data-native-text-run]')]
      const nativeTextRunTexts = nativeTextRuns.map((run) => (run.textContent ?? '').trim()).filter(Boolean)
      const readNumberAttribute = (name: string) => Number(renderer?.getAttribute(name) ?? '0')
      const headerTextRunCount =
        readNumberAttribute('data-v3-presented-header-text-run-count') || readNumberAttribute('data-v3-header-text-run-count')
      const visibleTextRunCount = readNumberAttribute('data-v3-presented-visible-text-run-count')
      const bodyTextRunCount =
        visibleTextRunCount ||
        readNumberAttribute('data-v3-presented-text-run-count') ||
        readNumberAttribute('data-v3-text-run-count') ||
        nativeTextRuns.length
      const rowHeaderRunCount =
        nativeTextRuns.filter((run) => {
          if (!gridRect || !/^\d+$/.test(run.textContent ?? '')) {
            return false
          }
          const rect = run.getBoundingClientRect()
          return rect.left >= gridRect.left && rect.right <= gridRect.left + 46 + 1
        }).length || headerTextRunCount
      return {
        gridAttrs: Object.fromEntries(
          [...(sheetGrid?.attributes ?? [])]
            .filter((attribute) => attribute.name.includes('render') || attribute.name.includes('revision'))
            .map((attribute) => [attribute.name, attribute.value]),
        ),
        matches: Object.fromEntries(
          payload.textExpectations.map((expectation) => {
            const matchesText =
              nativeTextRunTexts.length > 0
                ? nativeTextRunTexts.some((text) => (expectation.exact ? text === expectation.text : text.includes(expectation.text)))
                : expectation.name.toLowerCase().includes('header')
                  ? headerTextRunCount > 0
                  : bodyTextRunCount > 0
            return [expectation.name, matchesText]
          }),
        ),
        rendererAttrs: Object.fromEntries(
          [...(renderer?.attributes ?? [])]
            .filter((attribute) => attribute.name.includes('v3'))
            .map((attribute) => [attribute.name, attribute.value]),
        ),
        rowHeaderRunCount,
        sampleTexts: nativeTextRunTexts.slice(0, 32),
        visibleRunCount: nativeTextRuns.length || headerTextRunCount + bodyTextRunCount,
      }
    },
    { textExpectations },
  )
}

export function readNativeRectProofState(page: Page): () => Promise<{
  readonly countPresent: boolean
  readonly presentedFrameMatches: boolean
  readonly sceneEpochMatches: boolean
  readonly signatureMatches: boolean
  readonly visibleRevisionMatches: boolean
}> {
  return async () =>
    await page.evaluate(() => {
      const renderer = document.querySelector('[data-testid="grid-pane-renderer"]')
      const nativeRectLayer = document.querySelector('[data-testid="grid-native-rect-layer"]')
      const rendererAttr = (name: string) => (renderer instanceof HTMLElement ? (renderer.getAttribute(name) ?? '') : '')
      const rectAttr = (name: string) => (nativeRectLayer instanceof HTMLElement ? (nativeRectLayer.getAttribute(name) ?? '') : '')
      const rectCount = Number.parseInt(rectAttr('data-v3-native-rect-count'), 10) || 0
      return {
        countPresent: rectCount > 0,
        presentedFrameMatches:
          rectAttr('data-v3-native-rect-presented-frame-id').length > 0 &&
          rectAttr('data-v3-native-rect-presented-frame-id') === rendererAttr('data-v3-presented-frame-proof-signature'),
        sceneEpochMatches:
          rectAttr('data-v3-native-rect-scene-epoch').length > 0 &&
          rectAttr('data-v3-native-rect-scene-epoch') === rendererAttr('data-v3-presented-scene-epoch'),
        signatureMatches:
          rectAttr('data-v3-native-rect-signature').length > 0 &&
          rectAttr('data-v3-native-rect-signature') === rendererAttr('data-v3-presented-rect-signature'),
        visibleRevisionMatches:
          rectAttr('data-v3-native-rect-visible-render-revision').length > 0 &&
          rectAttr('data-v3-native-rect-visible-render-revision') === rendererAttr('data-v3-visible-render-revision'),
      }
    })
}

export async function waitForVisibleNativeTextRuns(
  page: Page,
  expectations: readonly NativeTextExpectation[],
  predicate: (runs: NativeTextLayerInspection) => boolean,
): Promise<NativeTextLayerInspection> {
  let lastResult: NativeTextLayerInspection | null = null
  try {
    await expect
      .poll(
        async () => {
          lastResult = await inspectVisibleNativeTextRuns(page, expectations)
          return predicate(lastResult)
        },
        { timeout: 10_000 },
      )
      .toBe(true)
  } catch (error) {
    throw new Error(`${error instanceof Error ? error.message : String(error)}\nLast native text runs: ${JSON.stringify(lastResult)}`, {
      cause: error,
    })
  }
  if (!lastResult) {
    throw new Error('expected native text run result')
  }
  return lastResult
}

export async function waitForNativeTextRunContent(page: Page, text: string): Promise<void> {
  await expect
    .poll(
      async () =>
        await page.evaluate(
          (needle) => [...document.querySelectorAll('[data-native-text-run]')].some((run) => run.textContent?.includes(needle) ?? false),
          text,
        ),
      { timeout: 10_000 },
    )
    .toBe(true)
}

export {
  productCellInteriorGridLocalPoint,
  productColumnHeaderGridLocalPoint,
  sampleCompositedGridLocalPixel,
  saveReadbackArtifact,
  waitForCompositedGridLocalPixel,
} from './web-shell-typegpu-composited-helpers.js'

export async function waitForReadback(
  page: Page,
  input: {
    readonly points: readonly ReadbackInspectorPoint[]
    readonly regions: readonly ReadbackInspectorRegion[]
  },
  predicate: (result: DynamicReadbackResult) => boolean,
): Promise<DynamicReadbackResult> {
  let lastResult: DynamicReadbackResult | null = null
  try {
    await expect
      .poll(
        async () => {
          lastResult = await inspectGpuReadback(page, input)
          return lastResult.ready && lastResult.hasGpu && predicate(lastResult)
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } catch (error) {
    throw new Error(`${error instanceof Error ? error.message : String(error)}\nLast readback: ${JSON.stringify(lastResult)}`, {
      cause: error,
    })
  }
  if (!lastResult) {
    throw new Error('expected readback result')
  }
  return lastResult
}

export async function waitForReadbackSequence(page: Page, previousSequence: number): Promise<void> {
  await page.waitForFunction(
    (sequence) => {
      const inspector = (window as Window & { __biligGpuReadbackInspector?: { readonly getSequence: () => number } })
        .__biligGpuReadbackInspector
      return (inspector?.getSequence() ?? 0) > sequence
    },
    previousSequence,
    { timeout: 30_000 },
  )
}
