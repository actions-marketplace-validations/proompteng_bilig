import { describe, expect, it } from 'vitest'

import {
  AGENT_PROTOCOL_VERSION,
  AGENT_STDIN_MAGIC,
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  decodeAgentFrame,
  decodeStdioMessages,
  encodeAgentFrame,
  encodeStdioMessage,
  normalizeWorkbookImportContentType,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
} from '../index.js'

describe('agent api', () => {
  it('roundtrips request and response envelopes', () => {
    const frame = {
      kind: 'request' as const,
      request: {
        kind: 'openWorkbookSession' as const,
        id: 'req-1',
        documentId: 'book-1',
        replicaId: 'agent-local',
      },
    }

    expect(decodeAgentFrame(encodeAgentFrame(frame))).toEqual(frame)
  })

  it('decodes multiple stdio messages from one buffer', () => {
    const encoded = new Uint8Array([
      ...encodeStdioMessage({
        kind: 'request',
        request: {
          kind: 'getMetrics',
          id: 'req-1',
          sessionId: 'sess-1',
        },
      }),
      ...encodeStdioMessage({
        kind: 'response',
        response: {
          kind: 'ok',
          id: 'req-1',
          value: { ok: true },
        },
      }),
    ])

    const decoded = decodeStdioMessages(encoded)
    expect(decoded.frames).toHaveLength(2)
    expect(decoded.remainder.byteLength).toBe(0)
  })

  it('rejects oversized stdio frames before waiting for their declared payload', () => {
    const prefix = new Uint8Array(4)
    new DataView(prefix.buffer).setUint32(0, 0xffff_ffff, true)

    expect(() => decodeStdioMessages(prefix)).toThrow('Agent frame exceeds the protocol size limit')
  })

  it('roundtrips workbook file load requests and responses', () => {
    const requestFrame = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-1',
        replicaId: 'agent-local',
        openMode: 'create' as const,
        fileName: 'report.xlsx',
        contentType: XLSX_CONTENT_TYPE,
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(requestFrame))).toEqual(requestFrame)

    const csvRequestFrame = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-2',
        replicaId: 'agent-local',
        openMode: 'replace' as const,
        documentId: 'doc-1',
        fileName: 'report.csv',
        contentType: CSV_CONTENT_TYPE,
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(csvRequestFrame))).toEqual(csvRequestFrame)

    const csvRequestFrameWithUploadMetadata = {
      kind: 'request' as const,
      request: {
        kind: 'loadWorkbookFile' as const,
        id: 'upload-3',
        replicaId: 'agent-local',
        openMode: 'create' as const,
        fileName: 'report.csv',
        contentType: 'TEXT/CSV; charset=utf-8',
        bytesBase64: 'QUJD',
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(csvRequestFrameWithUploadMetadata))).toEqual(csvRequestFrameWithUploadMetadata)

    const responseFrame = {
      kind: 'response' as const,
      response: {
        kind: 'workbookLoaded' as const,
        id: 'upload-1',
        documentId: 'xlsx:abc123',
        sessionId: 'xlsx:abc123:agent-local',
        workbookName: 'report.xlsx',
        sheetNames: ['Sheet1'],
        serverUrl: 'http://127.0.0.1:4321',
        browserUrl: 'http://127.0.0.1:4173/?document=xlsx%3Aabc123',
        warnings: [],
      },
    }
    expect(decodeAgentFrame(encodeAgentFrame(responseFrame))).toEqual(responseFrame)
  })

  it('normalizes workbook import content types from upload metadata', () => {
    expect(normalizeWorkbookImportContentType(' text/csv; charset=utf-8 ')).toBe(CSV_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('TEXT/CSV')).toBe(CSV_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType(`${XLSX_CONTENT_TYPE}; charset=binary`)).toBe(XLSX_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType(XLSX_CONTENT_TYPE.toUpperCase())).toBe(XLSX_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel.sheet.macroEnabled.12')).toBe(XLSM_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel.sheet.binary.macroEnabled.12')).toBe(XLSB_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/vnd.ms-excel; charset=binary')).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(normalizeWorkbookImportContentType('application/octet-stream')).toBeNull()
  })

  it.each([
    {
      name: 'missing write range fields',
      frame: { kind: 'request', request: { kind: 'writeRange', id: 'req-1' } },
    },
    {
      name: 'unknown request kind',
      frame: { kind: 'request', request: { kind: 'deleteEverything', id: 'req-1' } },
    },
    {
      name: 'malformed workbook upload base64',
      frame: {
        kind: 'request',
        request: {
          kind: 'loadWorkbookFile',
          id: 'upload-invalid',
          replicaId: 'agent-local',
          openMode: 'create',
          fileName: 'report.csv',
          contentType: CSV_CONTENT_TYPE,
          bytesBase64: 'not%%%base64',
        },
      },
    },
    {
      name: 'invalid response payload',
      frame: { kind: 'response', response: { kind: 'rangeValues', id: 'req-1', values: 'not-a-matrix' } },
    },
    {
      name: 'invalid event payload',
      frame: { kind: 'event', event: { kind: 'syncState', state: 'live' } },
    },
    {
      name: 'unknown event kind',
      frame: { kind: 'event', event: { kind: 'workbookDeleted' } },
    },
  ])('rejects $name', ({ frame }) => {
    expect(() => decodeAgentFrame(encodeUntrustedAgentFrame(frame))).toThrow('Invalid agent frame payload')
  })

  it('rejects unexpected fields instead of silently widening the protocol', () => {
    expect(() =>
      decodeAgentFrame(
        encodeUntrustedAgentFrame({
          kind: 'request',
          request: {
            kind: 'openWorkbookSession',
            id: 'req-1',
            documentId: 'book-1',
            replicaId: 'agent-local',
            grantsAdmin: true,
          },
        }),
      ),
    ).toThrow('Invalid agent frame payload')

    expect(() =>
      decodeAgentFrame(
        encodeUntrustedAgentFrame({
          kind: 'request',
          request: {
            kind: 'readRange',
            id: 'req-2',
            sessionId: 'book-1:agent-local',
            range: {
              sheetName: 'Sheet1',
              startAddress: 'A1',
              endAddress: 'B2',
              trustsCaller: true,
            },
          },
        }),
      ),
    ).toThrow('Invalid agent frame payload')
  })
})

function encodeUntrustedAgentFrame(frame: unknown): Uint8Array {
  const payload = new TextEncoder().encode(JSON.stringify(frame))
  const output = new Uint8Array(10 + payload.byteLength)
  const view = new DataView(output.buffer)
  view.setUint32(0, AGENT_STDIN_MAGIC, true)
  view.setUint16(4, AGENT_PROTOCOL_VERSION, true)
  view.setUint32(6, payload.byteLength, true)
  output.set(payload, 10)
  return output
}
