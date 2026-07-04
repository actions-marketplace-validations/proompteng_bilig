import { expect } from 'vitest'
import { buildTextQuadsFromRunsWithSpans } from '../renderer-v3/line-text-quad-buffer.js'
import type { GridRenderTile } from '../renderer-v3/render-tile-source.js'
import {
  resolveGridRectTileRevisionKeyV3,
  resolveGridTextTileRevisionKeyV3,
  shouldAttemptAxisOnlyTileTextGeometryResourceSync,
  syncAxisOnlyTileTextGeometryResource,
  type TypeGpuTileContentResourceEntryV3,
} from '../renderer-v3/typegpu-tile-buffer-pool.js'

export function createTile(overrides: Partial<GridRenderTile> = {}): GridRenderTile {
  const version = {
    axisX: 1,
    axisY: 1,
    freeze: 0,
    styles: 1,
    text: 1,
    values: 1,
    ...overrides.version,
  }
  return {
    bounds: { colEnd: 127, colStart: 0, rowEnd: 31, rowStart: 0 },
    coord: {
      colTile: 0,
      dprBucket: 1,
      paneKind: 'body',
      rowTile: 0,
      sheetId: 7,
    },
    lastBatchId: 1,
    lastCameraSeq: 1,
    rectCount: 1,
    rectInstances: new Float32Array([0, 0, 104, 22, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 200, 100]),
    textCount: 0,
    textMetrics: new Float32Array(),
    textRuns: [],
    tileId: 101,
    version,
    ...overrides,
  }
}

export function rectRevisionKey(tile: GridRenderTile): ReturnType<typeof resolveGridRectTileRevisionKeyV3> {
  return resolveGridRectTileRevisionKeyV3({ tile })
}

export function contentEntry(overrides: Partial<TypeGpuTileContentResourceEntryV3> = {}): TypeGpuTileContentResourceEntryV3 {
  return {
    decorationCellKeys: null,
    decorationRects: null,
    rectBaseCount: overrides.rectBaseCount ?? overrides.rectCount ?? 1,
    rectCount: 1,
    rectHandle: null,
    rectRevisionKey: null,
    textCount: 1,
    textAtlasDependencyVersion: 1,
    textAtlasGeometryVersion: 1,
    textGlyphIds: null,
    textGlyphPageIds: null,
    textHandle: null,
    textRunGlyphIds: null,
    textRunCount: 1,
    textRunPayloads: null,
    textRunQuadSpans: null,
    textRevisionKey: null,
    ...overrides,
  }
}

export function createTextRun(overrides: Partial<GridRenderTile['textRuns'][number]> = {}): GridRenderTile['textRuns'][number] {
  return {
    align: 'left',
    clipHeight: 22,
    clipWidth: 104,
    clipX: 0,
    clipY: 0,
    color: '#111111',
    font: '400 11px sans-serif',
    fontSize: 11,
    height: 22,
    strike: false,
    text: 'A1',
    underline: false,
    width: 104,
    wrap: false,
    x: 0,
    y: 0,
    ...overrides,
  }
}

interface RecordedVertexWrite {
  readonly endOffset?: number | undefined
  readonly floats: readonly number[]
  readonly startOffset?: number | undefined
}

export function createRecordedTextHandle(
  writes: RecordedVertexWrite[],
): NonNullable<Parameters<typeof syncAxisOnlyTileTextGeometryResource>[0]['handle']> {
  return {
    buffer: {
      write(source: ArrayBuffer, options?: { readonly startOffset?: number | undefined; readonly endOffset?: number | undefined }) {
        writes.push({
          endOffset: options?.endOffset,
          floats: Array.from(new Float32Array(source)),
          startOffset: options?.startOffset,
        })
      },
    },
    capacityBytes: 4096,
    classId: 4,
    layout: 'textRuns',
    usedBytes: 0,
  }
}

export function createRecordedRectHandle(): NonNullable<TypeGpuTileContentResourceEntryV3['rectHandle']> {
  return {
    buffer: {
      write() {},
    },
    capacityBytes: 4096,
    classId: 4,
    layout: 'rectInstances',
    usedBytes: 0,
  }
}

export function expectAxisOnlySyncMatchesFullTextPayload(input: {
  readonly dirtyMasks: Uint32Array
  readonly name: string
  readonly baseRun: GridRenderTile['textRuns'][number]
  readonly shiftedRun: GridRenderTile['textRuns'][number]
}): void {
  const writes: RecordedVertexWrite[] = []
  const atlas = createTestAtlas()
  const handle = createRecordedTextHandle(writes)
  const baseRuns: GridRenderTile['textRuns'] = [input.baseRun]
  const baseTile = createTile({
    rectCount: 0,
    rectInstances: new Float32Array(),
    textCount: baseRuns.length,
    textRuns: baseRuns,
  })
  const basePayload = buildTextQuadsFromRunsWithSpans(baseRuns, atlas)
  const content = contentEntry({
    rectBaseCount: 0,
    rectCount: 0,
    textCount: basePayload.quadCount,
    textGlyphIds: basePayload.glyphIds,
    textGlyphPageIds: basePayload.pageIds,
    textRunCount: baseRuns.length,
    textRunGlyphIds: basePayload.runGlyphIds,
    textRunPayloads: basePayload.runPayloads,
    textRunQuadSpans: basePayload.runSpans,
    textRevisionKey: resolveGridTextTileRevisionKeyV3(baseTile),
  })
  const shiftedRuns: GridRenderTile['textRuns'] = [input.shiftedRun]
  const shiftedTile = createTile({
    ...baseTile,
    dirty: {
      glyphSpans: [],
      rectSpans: [],
      textSpans: [{ offset: 0, length: 1 }],
    },
    dirtyMasks: input.dirtyMasks,
    textRuns: shiftedRuns,
    version: {
      ...baseTile.version,
      axisX: 2,
      axisY: 2,
    },
  })
  expect(
    shouldAttemptAxisOnlyTileTextGeometryResourceSync({
      contentRevisionKey: content.textRevisionKey,
      dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
      textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
      tile: shiftedTile,
    }),
    input.name,
  ).toBe(true)

  const didSync = syncAxisOnlyTileTextGeometryResource({
    content,
    dirtyTextRunSpans: shiftedTile.dirty?.textSpans,
    handle,
    label: 'tile-text:test',
    textRevisionKey: resolveGridTextTileRevisionKeyV3(shiftedTile),
    tile: shiftedTile,
  })
  const fullPayload = buildTextQuadsFromRunsWithSpans(shiftedRuns, atlas)

  expect(didSync, input.name).toBe(true)
  if (areNumberArraysEqual(Array.from(basePayload.floats), Array.from(fullPayload.floats))) {
    expect(writes, input.name).toHaveLength(0)
  } else {
    expect(writes, input.name).toHaveLength(1)
    expect(writes[0]?.startOffset, input.name).toBe(0)
    expect(writes[0]?.endOffset, input.name).toBe(fullPayload.floats.length * Float32Array.BYTES_PER_ELEMENT)
    expect(writes[0]?.floats, input.name).toEqual(Array.from(fullPayload.floats))
  }
  expect(Array.from(content.textRunPayloads?.[0]?.floats ?? []), input.name).toEqual(Array.from(fullPayload.runPayloads[0]?.floats ?? []))
}

function areNumberArraysEqual(left: readonly number[], right: readonly number[]): boolean {
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

export function createTestAtlas(): Parameters<typeof buildTextQuadsFromRunsWithSpans>[1] {
  return {
    getGlyphGeometryVersion: () => 1,
    getVersion: () => 1,
    intern(font: string, glyph: string) {
      const advance = Math.max(0, glyph.length * 8)
      return {
        advance,
        baseline: 10,
        font,
        glyph,
        glyphId: glyph.codePointAt(0) ?? 0,
        height: 12,
        key: `atlas:${glyph}`,
        originOffsetX: 0,
        pageId: 1,
        u0: 0,
        u1: 1,
        v0: 0,
        v1: 1,
        width: advance,
        x: 0,
        y: 0,
      }
    },
  }
}
