;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const canvasGlobal = globalThis as typeof globalThis & {
  HTMLCanvasElement?: typeof HTMLCanvasElement
}

if (typeof canvasGlobal.HTMLCanvasElement !== 'undefined') {
  const prototype = canvasGlobal.HTMLCanvasElement.prototype
  Object.defineProperty(prototype, 'getContext', {
    configurable: true,
    value(this: HTMLCanvasElement, contextId: string) {
      if (contextId !== '2d') {
        return null
      }
      return createVitestCanvas2DContext(this)
    },
  })
}

interface VitestCanvasTextMetrics {
  readonly actualBoundingBoxAscent: number
  readonly actualBoundingBoxDescent: number
  readonly actualBoundingBoxLeft: number
  readonly actualBoundingBoxRight: number
  readonly width: number
}

interface VitestCanvas2DContext {
  readonly canvas: HTMLCanvasElement
  clearRect: () => void
  drawImage: () => void
  fillRect: () => void
  fillStyle: string
  fillText: () => void
  font: string
  fontKerning: CanvasFontKerning
  getImageData: (_sx: number, _sy: number, sw: number, sh: number) => ImageData
  imageSmoothingEnabled: boolean
  imageSmoothingQuality: ImageSmoothingQuality
  measureText: (text: string) => VitestCanvasTextMetrics
  putImageData: () => void
  resetTransform: () => void
  scale: () => void
  setTransform: () => void
  textBaseline: CanvasTextBaseline
  textRendering: 'auto' | 'optimizeSpeed' | 'optimizeLegibility' | 'geometricPrecision'
}

function createVitestCanvas2DContext(canvas: HTMLCanvasElement): VitestCanvas2DContext {
  const context: VitestCanvas2DContext = {
    canvas,
    clearRect: () => {},
    drawImage: () => {},
    fillRect: () => {},
    fillStyle: '#000000',
    fillText: () => {},
    font: '10px sans-serif',
    fontKerning: 'auto',
    getImageData: (_sx: number, _sy: number, sw: number, sh: number) => new ImageData(Math.max(1, sw), Math.max(1, sh)),
    imageSmoothingEnabled: true,
    imageSmoothingQuality: 'low',
    measureText: (text: string) => {
      const fontSize = parseCanvasFontSize(context.font)
      const width = Math.max(1, text.length * fontSize * 0.6)
      return {
        actualBoundingBoxAscent: fontSize,
        actualBoundingBoxDescent: Math.max(1, fontSize * 0.2),
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: width,
        width,
      }
    },
    putImageData: () => {},
    resetTransform: () => {},
    scale: () => {},
    setTransform: () => {},
    textBaseline: 'alphabetic',
    textRendering: 'auto',
  }

  return context
}

function parseCanvasFontSize(font: string): number {
  const match = font.match(/(\d+(?:\.\d+)?)px/)
  return match ? Number(match[1]) : 10
}
