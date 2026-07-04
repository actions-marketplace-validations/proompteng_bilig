import { strToU8, zipSync } from 'fflate'

export function buildLargeSimpleWorkbook(input: {
  worksheetXml: string
  definedNamesXml?: string
  extraEntries?: Record<string, string | Uint8Array>
  includeSharedStrings?: boolean
  includeStyles?: boolean
  sharedStringsXml?: string
  sheetRelationshipsXml?: string
}): Uint8Array {
  const includeSharedStrings = input.includeSharedStrings ?? true
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
  ${input.definedNamesXml ?? ''}
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  ${
    includeSharedStrings
      ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
      : ''
  }
</Relationships>`),
    ...(includeSharedStrings
      ? {
          'xl/sharedStrings.xml': strToU8(
            input.sharedStringsXml ??
              `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Alpha &amp;#8211;</t></si>
  <si><t>Line_x000a_Break</t></si>
</sst>`,
          ),
        }
      : {}),
    ...(input.sheetRelationshipsXml ? { 'xl/worksheets/_rels/sheet1.xml.rels': strToU8(input.sheetRelationshipsXml) } : {}),
    ...(input.includeStyles
      ? {
          'xl/styles.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFCC00"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0"/></cellXfs>
</styleSheet>`),
        }
      : {}),
    'xl/worksheets/sheet1.xml': strToU8(input.worksheetXml),
    ...Object.fromEntries(
      Object.entries(input.extraEntries ?? {}).map(([path, value]) => [path, typeof value === 'string' ? strToU8(value) : value]),
    ),
  })
}

export function countLazyZipEntryStreams(zip: Record<string, Uint8Array>, path: string): () => number {
  const metadata = readLazyZipMetadata(zip)
  const entry = metadata?.entriesByPath.get(path)
  if (!metadata || !entry) {
    throw new Error(`Missing lazy ZIP metadata for ${path}`)
  }
  const source = metadata.source
  const localHeader = source.readRange(entry.localHeaderOffset, entry.localHeaderOffset + 30)
  const fileNameLength = readLittleEndianUint16(localHeader, 26)
  const extraFieldLength = readLittleEndianUint16(localHeader, 28)
  const dataStart = entry.localHeaderOffset + 30 + fileNameLength + extraFieldLength
  const dataEnd = dataStart + entry.compressedSize
  let streamCount = 0
  metadata.source = new Proxy(source, {
    get(target, property) {
      if (property === 'readRange') {
        return (start?: number, end?: number) => {
          if (start === dataStart && end === dataEnd) {
            streamCount += 1
          }
          return target.readRange(start ?? 0, end ?? target.byteLength)
        }
      }
      const value = Reflect.get(target, property, target)
      return typeof value === 'function' ? value.bind(target) : value
    },
  })
  return () => streamCount
}

export function countByteSourceZipEntryStreams(
  bytes: Uint8Array,
  path: string,
): {
  readonly source: {
    readonly byteLength: number
    readRange(start: number, end: number): Uint8Array
    readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array
  }
  readonly count: () => number
  readonly readIntoCount: () => number
} {
  const [dataStart] = findLocalZipEntryDataRange(bytes, path)
  let streamCount = 0
  let readIntoCount = 0
  return {
    source: {
      byteLength: bytes.byteLength,
      readRange(start = 0, end = bytes.byteLength): Uint8Array {
        if (start === dataStart) {
          streamCount += 1
        }
        return bytes.subarray(start, end)
      },
      readRangeInto(start: number, end: number, target: Uint8Array): Uint8Array {
        const length = end - start
        if (start === dataStart) {
          streamCount += 1
        }
        readIntoCount += 1
        target.set(bytes.subarray(start, end), 0)
        return target.subarray(0, length)
      },
    },
    count: () => streamCount,
    readIntoCount: () => readIntoCount,
  }
}

export function byteSourceFor(bytes: Uint8Array): { readonly byteLength: number; readRange(start: number, end: number): Uint8Array } {
  return {
    byteLength: bytes.byteLength,
    readRange(start, end) {
      return bytes.subarray(start, end)
    },
  }
}

function findLocalZipEntryDataRange(bytes: Uint8Array, path: string): readonly [number, number] {
  const decoder = new TextDecoder()
  let offset = 0
  while (offset + 30 <= bytes.byteLength && readLittleEndianUint32(bytes, offset) === 0x04034b50) {
    const compressedSize = readLittleEndianUint32(bytes, offset + 18)
    const fileNameLength = readLittleEndianUint16(bytes, offset + 26)
    const extraFieldLength = readLittleEndianUint16(bytes, offset + 28)
    const fileNameStart = offset + 30
    const fileNameEnd = fileNameStart + fileNameLength
    const dataStart = fileNameEnd + extraFieldLength
    const dataEnd = dataStart + compressedSize
    if (decoder.decode(bytes.subarray(fileNameStart, fileNameEnd)) === path) {
      return [dataStart, dataEnd]
    }
    offset = dataEnd
  }
  throw new Error(`Missing local ZIP header for ${path}`)
}

function readLazyZipMetadata(zip: Record<string, Uint8Array>):
  | {
      source: XlsxLazyZipByteSource
      readonly entriesByPath: ReadonlyMap<
        string,
        {
          readonly localHeaderOffset: number
          readonly compressedSize: number
        }
      >
    }
  | undefined {
  for (const symbol of Object.getOwnPropertySymbols(zip)) {
    const value = Reflect.get(zip, symbol) as unknown
    if (isLazyZipMetadata(value)) {
      return value
    }
  }
  return undefined
}

function isLazyZipMetadata(value: unknown): value is {
  source: XlsxLazyZipByteSource
  readonly entriesByPath: ReadonlyMap<
    string,
    {
      readonly localHeaderOffset: number
      readonly compressedSize: number
    }
  >
} {
  return (
    typeof value === 'object' &&
    value !== null &&
    'source' in value &&
    isLazyZipByteSource(value.source) &&
    'entriesByPath' in value &&
    value.entriesByPath instanceof Map
  )
}

interface XlsxLazyZipByteSource {
  readonly byteLength: number
  readRange(start: number, end: number): Uint8Array
}

function isLazyZipByteSource(value: unknown): value is XlsxLazyZipByteSource {
  return (
    typeof value === 'object' &&
    value !== null &&
    'byteLength' in value &&
    typeof value.byteLength === 'number' &&
    'readRange' in value &&
    typeof value.readRange === 'function'
  )
}

function readLittleEndianUint16(source: Uint8Array, offset: number): number {
  return source[offset] | (source[offset + 1] << 8)
}

function readLittleEndianUint32(source: Uint8Array, offset: number): number {
  return (source[offset] | (source[offset + 1] << 8) | (source[offset + 2] << 16) | (source[offset + 3] << 24)) >>> 0
}

export function deterministicBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length)
  let state = 0x12345678
  for (let index = 0; index < length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    bytes[index] = (state >>> 24) & 0xff
  }
  return bytes
}

export function decodeBase64(value: string): Uint8Array {
  return new Uint8Array(Buffer.from(value, 'base64'))
}

export function encodeColumnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    value -= 1
    output = String.fromCharCode(65 + (value % 26)) + output
    value = Math.floor(value / 26)
  }
  return output
}
