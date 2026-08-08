import {
  CELL_BORDER_STYLE_VALUES,
  CELL_BORDER_WEIGHT_VALUES,
  CELL_HORIZONTAL_ALIGNMENT_VALUES,
  CELL_NUMBER_FORMAT_KIND_VALUES,
  CELL_NUMBER_NEGATIVE_STYLE_VALUES,
  CELL_NUMBER_ZERO_STYLE_VALUES,
  CELL_DATE_STYLE_VALUES,
  CELL_STYLE_FIELD_VALUES,
  CELL_VERTICAL_ALIGNMENT_VALUES,
  isCellValue,
  isWorkbookSnapshot,
  type CellRangeRef,
  type CellValue,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import { z } from 'zod'
import { MAX_AGENT_WORKBOOK_IMPORT_BASE64_LENGTH } from './agent-frame-limits.js'
import type { AgentFrame } from './index.js'
import { normalizeWorkbookImportContentType } from './workbook-import-content-types.js'

const MAX_IDENTIFIER_LENGTH = 512
const MAX_TEXT_LENGTH = 32_767
const MAX_MATRIX_ROWS = 1_048_576
const MAX_MATRIX_COLUMNS = 16_384
const MAX_MATRIX_CELLS = 1_000_000

const identifierSchema = z.string().min(1).max(MAX_IDENTIFIER_LENGTH)
const textSchema = z.string().max(MAX_TEXT_LENGTH)
const finiteNumberSchema = z.number().finite()
const nullableBooleanSchema = z.boolean().nullable().optional()
const nullableNumberSchema = finiteNumberSchema.nullable().optional()
const nullableTextSchema = textSchema.nullable().optional()

const rangeSchema = z.strictObject({
  sheetId: z.number().int().positive().optional(),
  sheetName: identifierSchema,
  startAddress: identifierSchema,
  endAddress: identifierSchema,
}) satisfies z.ZodType<CellRangeRef>

const literalInputSchema = z.union([z.null(), z.boolean(), finiteNumberSchema, textSchema])
const literalMatrixSchema = z
  .array(z.array(literalInputSchema).max(MAX_MATRIX_COLUMNS))
  .max(MAX_MATRIX_ROWS)
  .refine((rows) => rows.reduce((count, row) => count + row.length, 0) <= MAX_MATRIX_CELLS)
const formulaMatrixSchema = z
  .array(z.array(textSchema).max(MAX_MATRIX_COLUMNS))
  .max(MAX_MATRIX_ROWS)
  .refine((rows) => rows.reduce((count, row) => count + row.length, 0) <= MAX_MATRIX_CELLS)

const borderSidePatchSchema = z.strictObject({
  style: z.enum(CELL_BORDER_STYLE_VALUES).nullable().optional(),
  weight: z.enum(CELL_BORDER_WEIGHT_VALUES).nullable().optional(),
  color: nullableTextSchema,
})

const stylePatchSchema = z.strictObject({
  fill: z
    .strictObject({
      backgroundColor: nullableTextSchema,
    })
    .nullable()
    .optional(),
  font: z
    .strictObject({
      family: nullableTextSchema,
      size: nullableNumberSchema,
      bold: nullableBooleanSchema,
      italic: nullableBooleanSchema,
      underline: nullableBooleanSchema,
      color: nullableTextSchema,
    })
    .nullable()
    .optional(),
  alignment: z
    .strictObject({
      horizontal: z.enum(CELL_HORIZONTAL_ALIGNMENT_VALUES).nullable().optional(),
      vertical: z.enum(CELL_VERTICAL_ALIGNMENT_VALUES).nullable().optional(),
      wrap: nullableBooleanSchema,
      indent: nullableNumberSchema,
      shrinkToFit: nullableBooleanSchema,
      readingOrder: nullableNumberSchema,
      textRotation: nullableNumberSchema,
      justifyLastLine: nullableBooleanSchema,
    })
    .nullable()
    .optional(),
  borders: z
    .strictObject({
      top: borderSidePatchSchema.nullable().optional(),
      right: borderSidePatchSchema.nullable().optional(),
      bottom: borderSidePatchSchema.nullable().optional(),
      left: borderSidePatchSchema.nullable().optional(),
    })
    .nullable()
    .optional(),
})

const numberFormatSchema = z.union([
  textSchema,
  z.strictObject({
    kind: z.enum(CELL_NUMBER_FORMAT_KIND_VALUES),
    currency: textSchema.optional(),
    decimals: z.number().int().min(0).max(30).optional(),
    useGrouping: z.boolean().optional(),
    negativeStyle: z.enum(CELL_NUMBER_NEGATIVE_STYLE_VALUES).optional(),
    zeroStyle: z.enum(CELL_NUMBER_ZERO_STYLE_VALUES).optional(),
    dateStyle: z.enum(CELL_DATE_STYLE_VALUES).optional(),
  }),
])

const pivotValueSchema = z.strictObject({
  sourceColumn: identifierSchema,
  summarizeBy: z.enum(['sum', 'count', 'countNums', 'average', 'min', 'max', 'product']),
  outputLabel: textSchema.optional(),
})

const sessionRequestBase = {
  id: identifierSchema,
  sessionId: identifierSchema,
} as const

const agentRequestSchema = z.discriminatedUnion('kind', [
  z.strictObject({
    kind: z.literal('openWorkbookSession'),
    id: identifierSchema,
    documentId: identifierSchema,
    replicaId: identifierSchema,
  }),
  z.strictObject({ kind: z.literal('closeWorkbookSession'), ...sessionRequestBase }),
  z.strictObject({ kind: z.literal('readRange'), ...sessionRequestBase, range: rangeSchema }),
  z.strictObject({ kind: z.literal('writeRange'), ...sessionRequestBase, range: rangeSchema, values: literalMatrixSchema }),
  z.strictObject({ kind: z.literal('setRangeFormulas'), ...sessionRequestBase, range: rangeSchema, formulas: formulaMatrixSchema }),
  z.strictObject({ kind: z.literal('setRangeStyle'), ...sessionRequestBase, range: rangeSchema, patch: stylePatchSchema }),
  z.strictObject({
    kind: z.literal('clearRangeStyle'),
    ...sessionRequestBase,
    range: rangeSchema,
    fields: z.array(z.enum(CELL_STYLE_FIELD_VALUES)).max(CELL_STYLE_FIELD_VALUES.length).optional(),
  }),
  z.strictObject({ kind: z.literal('setRangeNumberFormat'), ...sessionRequestBase, range: rangeSchema, format: numberFormatSchema }),
  z.strictObject({ kind: z.literal('clearRangeNumberFormat'), ...sessionRequestBase, range: rangeSchema }),
  z.strictObject({ kind: z.literal('clearRange'), ...sessionRequestBase, range: rangeSchema }),
  z.strictObject({ kind: z.literal('fillRange'), ...sessionRequestBase, source: rangeSchema, target: rangeSchema }),
  z.strictObject({ kind: z.literal('copyRange'), ...sessionRequestBase, source: rangeSchema, target: rangeSchema }),
  z.strictObject({ kind: z.literal('moveRange'), ...sessionRequestBase, source: rangeSchema, target: rangeSchema }),
  z.strictObject({ kind: z.literal('pasteRange'), ...sessionRequestBase, source: rangeSchema, target: rangeSchema }),
  z.strictObject({
    kind: z.literal('getDependents'),
    ...sessionRequestBase,
    sheetName: identifierSchema,
    address: identifierSchema,
  }),
  z.strictObject({
    kind: z.literal('getPrecedents'),
    ...sessionRequestBase,
    sheetName: identifierSchema,
    address: identifierSchema,
  }),
  z.strictObject({
    kind: z.literal('subscribeRange'),
    ...sessionRequestBase,
    range: rangeSchema,
    subscriptionId: identifierSchema,
  }),
  z.strictObject({ kind: z.literal('unsubscribe'), ...sessionRequestBase, subscriptionId: identifierSchema }),
  z.strictObject({ kind: z.literal('exportSnapshot'), ...sessionRequestBase }),
  z.strictObject({
    kind: z.literal('importSnapshot'),
    ...sessionRequestBase,
    snapshot: z.custom<WorkbookSnapshot>(isWorkbookSnapshot),
  }),
  z.strictObject({ kind: z.literal('getMetrics'), ...sessionRequestBase }),
  z.strictObject({
    kind: z.literal('createPivotTable'),
    ...sessionRequestBase,
    name: identifierSchema,
    sheetName: identifierSchema,
    address: identifierSchema,
    source: rangeSchema,
    groupBy: z.array(identifierSchema).max(MAX_MATRIX_COLUMNS),
    values: z.array(pivotValueSchema).max(MAX_MATRIX_COLUMNS),
  }),
  z
    .strictObject({
      kind: z.literal('loadWorkbookFile'),
      id: identifierSchema,
      replicaId: identifierSchema,
      openMode: z.enum(['create', 'replace']),
      documentId: identifierSchema.optional(),
      fileName: identifierSchema,
      contentType: textSchema.refine((value) => normalizeWorkbookImportContentType(value) !== null),
      bytesBase64: z.base64().min(1).max(MAX_AGENT_WORKBOOK_IMPORT_BASE64_LENGTH),
    })
    .refine((request) => request.openMode !== 'replace' || request.documentId !== undefined),
])

const cellValueMatrixSchema = z
  .array(z.array(z.custom<CellValue>(isCellValue)).max(MAX_MATRIX_COLUMNS))
  .max(MAX_MATRIX_ROWS)
  .refine((rows) => rows.reduce((count, row) => count + row.length, 0) <= MAX_MATRIX_CELLS)
const workbookLoadedResponseSchema = z.strictObject({
  kind: z.literal('workbookLoaded'),
  id: identifierSchema,
  documentId: identifierSchema,
  sessionId: identifierSchema,
  workbookName: identifierSchema,
  sheetNames: z.array(identifierSchema).max(MAX_MATRIX_ROWS),
  serverUrl: z.url(),
  browserUrl: z.url().optional(),
  warnings: z.array(textSchema).max(10_000),
})

const agentResponseSchema = z.discriminatedUnion('kind', [
  z.strictObject({ kind: z.literal('ok'), id: identifierSchema, sessionId: identifierSchema.optional(), value: z.unknown().optional() }),
  z.strictObject({ kind: z.literal('rangeValues'), id: identifierSchema, values: cellValueMatrixSchema }),
  z.strictObject({ kind: z.literal('dependencies'), id: identifierSchema, addresses: z.array(identifierSchema).max(MAX_MATRIX_CELLS) }),
  z.strictObject({ kind: z.literal('snapshot'), id: identifierSchema, snapshot: z.custom<WorkbookSnapshot>(isWorkbookSnapshot) }),
  z.strictObject({ kind: z.literal('metrics'), id: identifierSchema, value: z.unknown() }),
  workbookLoadedResponseSchema,
  z.strictObject({
    kind: z.literal('error'),
    id: identifierSchema,
    code: identifierSchema,
    message: textSchema,
    retryable: z.boolean(),
  }),
])

const agentEventSchema = z.discriminatedUnion('kind', [
  z.strictObject({
    kind: z.literal('rangeChanged'),
    subscriptionId: identifierSchema,
    range: rangeSchema,
    changedAddresses: z.array(identifierSchema).max(MAX_MATRIX_CELLS),
  }),
  z.strictObject({
    kind: z.literal('syncState'),
    sessionId: identifierSchema,
    state: z.enum(['local-only', 'syncing', 'live', 'behind', 'reconnecting']),
  }),
])

const agentFrameSchema = z.discriminatedUnion('kind', [
  z.strictObject({ kind: z.literal('request'), request: agentRequestSchema }),
  z.strictObject({ kind: z.literal('response'), response: agentResponseSchema }),
  z.strictObject({ kind: z.literal('event'), event: agentEventSchema }),
])

export function isAgentFramePayload(value: unknown): value is AgentFrame {
  return agentFrameSchema.safeParse(value).success
}
