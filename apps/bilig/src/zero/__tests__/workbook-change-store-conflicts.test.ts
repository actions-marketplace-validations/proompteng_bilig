import type { Row } from '@rocicorp/zero'
import { describe, expect, it } from 'vitest'
import type { QueryResultRow, Queryable } from '../store.js'
import { loadLatestRedoableWorkbookChange, loadLatestUndoableWorkbookChange } from '../workbook-change-store.js'

interface RecordedQuery {
  readonly text: string
  readonly values: readonly unknown[] | undefined
}

class FakeQueryable implements Queryable {
  readonly calls: RecordedQuery[] = []

  constructor(
    private readonly responders: readonly ((text: string, values: readonly unknown[] | undefined) => QueryResultRow[] | null)[] = [],
  ) {}

  async query<T extends QueryResultRow = QueryResultRow>(text: string, values?: unknown[]): Promise<{ rows: T[] }> {
    this.calls.push({ text, values })
    for (const responder of this.responders) {
      const rows = responder(text, values)
      if (rows) {
        return {
          rows: rows.filter((row): row is T => row !== null),
        }
      }
    }
    return { rows: [] }
  }

  async loadWorkbookChangeRow(input: { readonly documentId: string; readonly revision: number }): Promise<ZeroWorkbookChangeRow | null> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1 AND revision = $2
      `,
      [input.documentId, input.revision],
    )
    return result.rows[0] ? toZeroWorkbookChangeRow(result.rows[0]) : null
  }

  async listWorkbookChangesAfterRevisionRows(input: {
    readonly documentId: string
    readonly revision: number
  }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1 AND revision > $2
        ORDER BY revision ASC
      `,
      [input.documentId, input.revision],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }

  async listWorkbookHistoryRows(input: { readonly documentId: string }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1
        ORDER BY revision ASC
      `,
      [input.documentId],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }

  async listRecentWorkbookChangeRows(input: {
    readonly documentId: string
    readonly limit: number
  }): Promise<readonly ZeroWorkbookChangeRow[]> {
    const result = await this.query(
      `
        FROM workbook_change
        WHERE workbook_id = $1
        ORDER BY revision DESC
      `,
      [input.documentId, input.limit],
    )
    return result.rows.map(toZeroWorkbookChangeRow)
  }
}

type ZeroWorkbookChangeRow = Row['workbook_change']

type JsonValue = string | number | boolean | null | readonly JsonValue[] | { readonly [key: string]: JsonValue }

function isJsonValue(value: unknown): value is JsonValue {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return true
  }
  if (Array.isArray(value)) {
    return value.every((entry) => isJsonValue(entry))
  }
  if (typeof value !== 'object') {
    return false
  }
  return Object.values(value).every((entry) => isJsonValue(entry))
}

function stringValue(value: unknown, fallback: string): string {
  return typeof value === 'string' ? value : fallback
}

function optionalStringValue(value: unknown): string | null {
  return typeof value === 'string' ? value : null
}

function numberValue(value: unknown, fallback: number): number {
  return typeof value === 'number' ? value : fallback
}

function optionalNumberValue(value: unknown): number | null {
  return typeof value === 'number' ? value : null
}

function jsonValue(value: unknown): JsonValue | null {
  return isJsonValue(value) ? value : null
}

function toZeroWorkbookChangeRow(row: QueryResultRow): ZeroWorkbookChangeRow {
  return {
    workbookId: stringValue(row['workbookId'], 'doc-1'),
    revision: numberValue(row['revision'], -1),
    actorUserId: stringValue(row['actorUserId'], ''),
    clientMutationId: optionalStringValue(row['clientMutationId']),
    eventKind: stringValue(row['eventKind'], ''),
    summary: stringValue(row['summary'], ''),
    sheetId: optionalNumberValue(row['sheetId']),
    sheetName: optionalStringValue(row['sheetName']),
    anchorAddress: optionalStringValue(row['anchorAddress']),
    rangeJson: jsonValue(row['rangeJson']),
    undoBundleJson: jsonValue(row['undoBundleJson']),
    revertedByRevision: optionalNumberValue(row['revertedByRevision']),
    revertsRevision: optionalNumberValue(row['revertsRevision']),
    createdAt: numberValue(row['createdAt'] ?? row['createdAtUnixMs'], 0),
  }
}

function latestQuery(queryable: FakeQueryable): RecordedQuery {
  const query = queryable.calls.at(-1)
  if (!query) {
    throw new Error('Expected at least one query')
  }
  return query
}

describe('workbook-change-store undo redo conflict visibility', () => {
  it('loads latest undoable and redoable changes for an actor', async () => {
    const queryable = new FakeQueryable([
      (text) => {
        if (!text.includes('FROM workbook_change')) {
          return null
        }
        return [
          {
            revision: 14,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-14',
            eventKind: 'revertChange',
            summary: 'Reverted r13: Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 5 }],
            },
            revertedByRevision: null,
            revertsRevision: 13,
            createdAtUnixMs: 123_460,
          } satisfies QueryResultRow,
          {
            revision: 13,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-13',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
            },
            revertedByRevision: null,
            revertsRevision: null,
            createdAtUnixMs: 123_456,
          } satisfies QueryResultRow,
        ]
      },
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toMatchObject({
      revision: 14,
      eventKind: 'revertChange',
    })
    expect(latestQuery(queryable).values).toEqual(['doc-1'])
  })

  it('does not expose redo after a fresh authored change branches history after an undo', async () => {
    const queryable = new FakeQueryable([
      (text) => {
        if (!text.includes('FROM workbook_change')) {
          return null
        }
        if (text.includes("event_kind = 'revertChange'")) {
          return [
            {
              revision: 22,
              actorUserId: 'alex@example.com',
              clientMutationId: 'mutation-22',
              eventKind: 'revertChange',
              summary: 'Reverted r21: Updated Sheet1!A1',
              sheetId: 1,
              sheetName: 'Sheet1',
              anchorAddress: 'A1',
              rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
              undoBundleJson: {
                kind: 'engineOps',
                ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
              },
              revertedByRevision: null,
              revertsRevision: 21,
              createdAtUnixMs: 123_460,
            } satisfies QueryResultRow,
          ]
        }
        return [
          {
            revision: 23,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-23',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!C1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'C1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C1' }],
            },
            revertedByRevision: null,
            revertsRevision: null,
            createdAtUnixMs: 123_470,
          } satisfies QueryResultRow,
          {
            revision: 22,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-22',
            eventKind: 'revertChange',
            summary: 'Reverted r21: Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
            },
            revertedByRevision: null,
            revertsRevision: 21,
            createdAtUnixMs: 123_460,
          } satisfies QueryResultRow,
          {
            revision: 21,
            actorUserId: 'alex@example.com',
            clientMutationId: 'mutation-21',
            eventKind: 'setCellValue',
            summary: 'Updated Sheet1!A1',
            sheetId: 1,
            sheetName: 'Sheet1',
            anchorAddress: 'A1',
            rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
            undoBundleJson: {
              kind: 'engineOps',
              ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
            },
            revertedByRevision: 22,
            revertsRevision: null,
            createdAtUnixMs: 123_450,
          } satisfies QueryResultRow,
        ]
      },
    ])

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('does not expose undo after another actor changes an overlapping range', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 32,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-32',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
              {
                revision: 31,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-31',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_470,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('does not expose undo when a later malformed range would otherwise fall back to a disjoint anchor', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 52,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-52',
                eventKind: 'insertRows',
                summary: 'Inserted rows 3:4 on Sheet1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A3',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A3', endAddress: 'A4', scope: 'row-band' },
                undoBundleJson: null,
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_490,
              } satisfies QueryResultRow,
              {
                revision: 51,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-51',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!Z99',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'Z99',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'Z99', endAddress: 'Z99' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'Z99' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestUndoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toBeNull()
  })

  it('keeps redo after another actor changes a disjoint range', async () => {
    const queryable = new FakeQueryable([
      (text) =>
        text.includes('FROM workbook_change')
          ? [
              {
                revision: 42,
                actorUserId: 'morgan@example.com',
                clientMutationId: 'mutation-42',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!C1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'C1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'C1' }],
                },
                revertedByRevision: null,
                revertsRevision: null,
                createdAtUnixMs: 123_480,
              } satisfies QueryResultRow,
              {
                revision: 41,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-41',
                eventKind: 'revertChange',
                summary: 'Reverted r40: Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 'seed' }],
                },
                revertedByRevision: null,
                revertsRevision: 40,
                createdAtUnixMs: 123_470,
              } satisfies QueryResultRow,
              {
                revision: 40,
                actorUserId: 'alex@example.com',
                clientMutationId: 'mutation-40',
                eventKind: 'setCellValue',
                summary: 'Updated Sheet1!A1',
                sheetId: 1,
                sheetName: 'Sheet1',
                anchorAddress: 'A1',
                rangeJson: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
                undoBundleJson: {
                  kind: 'engineOps',
                  ops: [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }],
                },
                revertedByRevision: 41,
                revertsRevision: null,
                createdAtUnixMs: 123_460,
              } satisfies QueryResultRow,
            ]
          : null,
    ])

    await expect(
      loadLatestRedoableWorkbookChange(queryable, {
        documentId: 'doc-1',
        actorUserId: 'alex@example.com',
      }),
    ).resolves.toMatchObject({
      revision: 41,
    })
  })
})
