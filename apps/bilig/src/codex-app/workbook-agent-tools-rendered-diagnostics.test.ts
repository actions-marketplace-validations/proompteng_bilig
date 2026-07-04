import type { WorkbookAgentCommandBundle } from './workbook-agent-tools-test-helpers.js'
import {
  applyWorkbookAgentCommandBundleWithUndoCapture,
  createBundle,
  createEngine,
  createVisibleSceneProof,
  createZeroSyncHarness,
  describe,
  expect,
  handleWorkbookAgentToolCall,
  invariantPayloadSchema,
  it,
  readToolJson,
  renderedSelectionPayloadSchema,
  tableListPayloadSchema,
  vi,
  z,
} from './workbook-agent-tools-test-helpers.js'

describe('workbook agent tools rendered readback and diagnostics', () => {
  it('keeps table metadata on existing sheets after rename and table creation', async () => {
    const engine = await createEngine()
    engine.createSheet('Sheet3')
    engine.setCellValue('Sheet3', 'A1', 'Vendor')
    engine.setCellValue('Sheet3', 'B1', 'Amount')
    engine.setCellValue('Sheet3', 'A2', 'Insurance')
    engine.setCellValue('Sheet3', 'B2', 1200)
    engine.renameSheet('Sheet3', 'Prepaid Template')
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => {
      const bundle = createBundle(command)
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      return {
        bundle,
        executionRecord: {
          id: 'run-table',
          bundleId: bundle.id,
          documentId: bundle.documentId,
          threadId: bundle.threadId,
          turnId: bundle.turnId,
          actorUserId: 'alex@example.com',
          goalText: bundle.goalText,
          planText: null,
          summary: bundle.summary,
          scope: bundle.scope,
          riskClass: bundle.riskClass,
          acceptedScope: 'full' as const,
          appliedBy: 'auto' as const,
          baseRevision: bundle.baseRevision,
          appliedRevision: 3,
          context: bundle.context,
          commands: bundle.commands,
          preview: null,
          createdAtUnixMs: 3,
          appliedAtUnixMs: 3,
        },
      }
    })

    const createResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-create-prepaid-table',
        tool: 'create_table',
        arguments: {
          name: 'Prepaids',
          range: {
            sheetName: 'Prepaid Template',
            startAddress: 'A1',
            endAddress: 'B2',
          },
          headerRow: true,
        },
      },
    )
    expect(createResponse.success).toBe(true)

    const tablesResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-list-prepaid-tables',
        tool: 'list_tables',
        arguments: {},
      },
    )
    const tables = tableListPayloadSchema.parse(readToolJson(tablesResponse))
    expect(tables.tables).toEqual([expect.objectContaining({ sheetName: 'Prepaid Template' })])

    const invariantResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-table-invariants',
        tool: 'verify_invariants',
        arguments: {},
      },
    )
    const invariants = invariantPayloadSchema.parse(readToolJson(invariantResponse))
    expect(invariants.summary.ok).toBe(true)
    expect(JSON.stringify(invariants.problems)).not.toContain('missing sheet')
  })

  it('reads cached rendered selection state alongside authoritative state', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'E5', 'changed')
    engine.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'E5', endAddress: 'E5' },
      {
        fill: { backgroundColor: '#93c47d' },
      },
    )
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const context = {
      selection: {
        sheetName: 'Sheet1',
        address: 'E5',
        range: {
          startAddress: 'E5',
          endAddress: 'E5',
        },
      },
      viewport: {
        rowStart: 4,
        rowEnd: 4,
        colStart: 4,
        colEnd: 4,
      },
      rendered: {
        capturedAtUnixMs: 10,
        capturedRevision: 11,
        batchId: 1,
        visibleSceneProof: createVisibleSceneProof(11),
        selection: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'E5',
            endAddress: 'E5',
          },
          rowCount: 1,
          columnCount: 1,
          cellCount: 1,
          truncated: false,
          rows: [
            [
              {
                address: 'E5',
                input: 'changed',
                value: { tag: 2, value: 'changed' },
                formula: null,
                displayFormat: 'changed',
                styleId: 'style-rendered',
                numberFormatId: null,
                style: { fill: { backgroundColor: '#93c47d' } },
              },
            ],
          ],
        },
        visibleRange: null,
      },
    }

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: context,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-rendered-selection',
        tool: 'read_rendered_selection',
        arguments: {},
      },
    )

    const payload = renderedSelectionPayloadSchema.parse(readToolJson(response))
    expect(payload.authoritativeReadback.rows[0]?.[0]?.value).toBe('changed')
    expect(payload.renderedReadback.available).toBe(true)
    expect(payload.renderedReadback.capturedRevision).toBe(11)
    expect(payload.renderedReadback.capturedBatchId).toBe(1)
    expect(JSON.stringify(payload.renderedReadback.range.rows[0]?.[0]?.style)).toContain('#93c47d')
  })

  it('reads rendered ranges from a captured visible viewport subset with freshness proof', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'B2', 'viewport proof')
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const context = {
      selection: {
        sheetName: 'Sheet1',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      rendered: {
        capturedAtUnixMs: 10,
        capturedRevision: 11,
        batchId: 1,
        visibleSceneProof: createVisibleSceneProof(11),
        selection: null,
        visibleRange: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'C3',
          },
          rowCount: 3,
          columnCount: 3,
          cellCount: 9,
          truncated: false,
          rows: [
            [
              {
                address: 'A1',
                input: 42,
                value: { tag: 1, value: 42 },
                formula: null,
                displayFormat: '$42.00',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B1',
                input: null,
                value: { tag: 1, value: 42 },
                formula: '=SUM(A1:A1)',
                displayFormat: '42',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C1',
                input: null,
                value: { tag: 4, code: 7 },
                formula: '=1/0',
                displayFormat: '#DIV/0!',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
            [
              {
                address: 'A2',
                input: 'Gross Margin',
                value: { tag: 3, value: 'Gross Margin' },
                formula: null,
                displayFormat: 'Gross Margin',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B2',
                input: 'viewport proof',
                value: { tag: 3, value: 'viewport proof' },
                formula: null,
                displayFormat: 'viewport proof',
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C2',
                input: null,
                value: { tag: 0 },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
            [
              {
                address: 'A3',
                input: null,
                value: { tag: 0 },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'B3',
                input: null,
                value: { tag: 0 },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
              {
                address: 'C3',
                input: null,
                value: { tag: 0 },
                formula: null,
                displayFormat: null,
                styleId: null,
                numberFormatId: null,
                style: null,
              },
            ],
          ],
        },
      },
    }

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: context,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-rendered-subset',
        tool: 'read_rendered_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B2',
        },
      },
    )

    const payload = z
      .object({
        renderedReadback: z.object({
          available: z.literal(true),
          matched: z.literal(true),
          stale: z.literal(false),
          capturedRevision: z.literal(11),
          capturedBatchId: z.literal(1),
          sourceRange: z.object({
            startAddress: z.literal('A1'),
            endAddress: z.literal('C3'),
          }),
          capturedRange: z.object({
            startAddress: z.literal('B2'),
            endAddress: z.literal('B2'),
          }),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.renderedReadback.sourceRange.startAddress).toBe('A1')
  })

  it('returns a chunk plan instead of throwing when read_range exceeds the single-call cell limit', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-chunked-read',
        tool: 'read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A4001',
        },
      },
    )

    const payload = z
      .object({
        chunked: z.literal(true),
        truncated: z.literal(true),
        totalCells: z.literal(4001),
        currentChunk: z.object({
          startAddress: z.literal('A1'),
          endAddress: z.literal('A4000'),
          cellCount: z.literal(4000),
        }),
        nextChunk: z.object({
          startAddress: z.literal('A4001'),
          endAddress: z.literal('A4001'),
          cellCount: z.literal(1),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.nextChunk.cellCount).toBe(1)
  })

  it('returns actionable formula diagnostics with exact dependencies and recalc status', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-formula-diagnostics',
        tool: 'find_formula_issues',
        arguments: {
          sheetName: 'Sheet1',
          limit: 10,
        },
      },
    )

    const payload = z
      .object({
        actionableIssues: z.array(
          z.object({
            formula: z.string(),
            errorText: z.string().nullable(),
            recalculationStatus: z.enum(['upToDate', 'stale']),
            directPrecedents: z.array(z.string()),
            directDependents: z.array(z.string()),
            suggestedNextInspectionRanges: z.array(z.string()),
          }),
        ),
      })
      .parse(readToolJson(response))
    expect(payload.actionableIssues.some((issue) => issue.formula === '=1/0' && issue.errorText !== null)).toBe(true)
    expect(payload.actionableIssues[0]?.suggestedNextInspectionRanges.length).toBeGreaterThan(0)
  })

  it('stages format commands with normalized number format presets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-3',
        tool: 'bilig_format_range',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A2',
          },
          numberFormat: {
            kind: 'currency',
            currency: 'USD',
          },
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A2',
      },
      numberFormat: {
        kind: 'currency',
        currency: 'USD',
        decimals: 2,
        useGrouping: true,
        negativeStyle: 'minus',
        zeroStyle: 'zero',
      },
    })
  })

  it('normalizes flat format_range patch aliases into workbook style patches', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-flat-format-range',
        tool: 'format_range',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'K1',
          },
          patch: {
            fontWeight: '700',
            fontSize: 16,
            fillColor: '#E8F0FE',
            horizontalAlignment: 'left',
            verticalAlignment: 'middle',
            borderBottom: { style: 'solid', color: '#C7D2FE' },
          },
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'K1',
      },
      patch: {
        fill: {
          backgroundColor: '#E8F0FE',
        },
        font: {
          bold: true,
          size: 16,
        },
        alignment: {
          horizontal: 'left',
          vertical: 'middle',
        },
        borders: {
          bottom: {
            style: 'solid',
            color: '#C7D2FE',
          },
        },
      },
    })
  })

  it('expands border shorthand across all sides for format_range', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-border-shorthand-format-range',
        tool: 'format_range',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'D1',
            endAddress: 'E5',
          },
          patch: {
            border: { style: 'solid', color: '#CBD5E1' },
          },
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'D1',
        endAddress: 'E5',
      },
      patch: {
        borders: {
          top: { style: 'solid', color: '#CBD5E1' },
          right: { style: 'solid', color: '#CBD5E1' },
          bottom: { style: 'solid', color: '#CBD5E1' },
          left: { style: 'solid', color: '#CBD5E1' },
        },
      },
    })
  })

  it('rejects empty format_range style patches instead of staging a no-op', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-empty-format-range',
        tool: 'format_range',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'A1',
          },
          patch: {},
        },
      },
    )

    expect(response.success).toBe(false)
    expect(response.contentItems[0]).toEqual(
      expect.objectContaining({
        type: 'inputText',
        text: expect.stringContaining('format_range patch did not include any supported style fields'),
      }),
    )
    expect(stageCommand).not.toHaveBeenCalled()
  })

  it('stages row metadata commands for hide and resize operations', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-4',
        tool: 'bilig_update_row_metadata',
        arguments: {
          sheetName: 'Sheet1',
          startRow: 1,
          count: 2,
          hidden: true,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'updateRowMetadata',
      sheetName: 'Sheet1',
      startRow: 1,
      count: 2,
      hidden: true,
    })
  })

  it('stages structural row insertion commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-insert-rows',
        tool: 'insert_rows',
        arguments: {
          sheetName: 'Sheet1',
          start: 1,
          count: 2,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'insertRows',
      sheetName: 'Sheet1',
      start: 1,
      count: 2,
    })
  })

  it('stages structural column deletion commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-delete-columns',
        tool: 'delete_columns',
        arguments: {
          sheetName: 'Sheet1',
          start: 0,
          count: 1,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'deleteColumns',
      sheetName: 'Sheet1',
      start: 0,
      count: 1,
    })
  })

  it.each([
    {
      name: 'delete_sheet',
      request: { name: 'Imports' },
      expected: { kind: 'deleteSheet', name: 'Imports' },
    },
    {
      name: 'set_freeze_pane',
      request: { sheetName: 'Sheet1', rows: 1, cols: 2 },
      expected: { kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 2 },
    },
    {
      name: 'set_filter',
      request: {
        range: { sheetName: 'Sheet1', startAddress: 'D4', endAddress: 'A1' },
      },
      expected: {
        kind: 'setFilter',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D4' },
      },
    },
    {
      name: 'clear_filter',
      request: {
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D4' },
      },
      expected: {
        kind: 'clearFilter',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D4' },
      },
    },
    {
      name: 'set_sort',
      request: {
        range: { sheetName: 'Sheet1', startAddress: 'B3', endAddress: 'A1' },
        keys: [{ keyAddress: 'B1', direction: 'desc' as const }],
      },
      expected: {
        kind: 'setSort',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
        keys: [{ keyAddress: 'B1', direction: 'desc' as const }],
      },
    },
    {
      name: 'clear_sort',
      request: {
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      },
      expected: {
        kind: 'clearSort',
        range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      },
    },
  ])('stages $name commands', async ({ name, request, expected }) => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: `call-${name}`,
        tool: name,
        arguments: request,
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith(expected)
  })

  it('starts built-in durable workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'summarizeWorkbook' as const,
      title: 'Summarize Workbook',
      summary: 'Summarized workbook structure across 2 sheets.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-workbook',
          label: 'Inspect workbook structure',
          status: 'completed' as const,
          summary: 'Read durable workbook structure across 2 sheets.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'draft-summary',
          label: 'Draft summary artifact',
          status: 'completed' as const,
          summary: 'Prepared the durable workbook summary artifact for the thread.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Workbook Summary',
        text: '## Workbook Summary',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'summarizeWorkbook',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'summarizeWorkbook',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"runId": "wf-1"')
    expect(output && 'text' in output ? output.text : '').toContain('"title": "Workbook Summary"')
  })

  it('starts structural create-sheet workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-create-sheet-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'createSheet' as const,
      title: 'Create Sheet',
      summary: 'Staged a structural change set to create Forecast.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'plan-sheet-create',
          label: 'Plan sheet creation',
          status: 'completed' as const,
          summary: 'Prepared the semantic sheet-creation command for Forecast.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'stage-structural-preview',
          label: 'Stage structural preview',
          status: 'completed' as const,
          summary: 'Staged the structural change set in the thread panel.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Create Sheet Preview',
        text: '## Create Sheet Preview',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-create-sheet-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'createSheet',
          name: 'Forecast',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'createSheet',
      name: 'Forecast',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "createSheet"')
  })

  it('starts query-driven workbook search workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-search-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'searchWorkbookQuery' as const,
      title: 'Search Workbook',
      summary: 'Found 2 workbook matches for "revenue".',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'search-workbook',
          label: 'Search workbook',
          status: 'completed' as const,
          summary: 'Searched workbook sheets, formulas, values, and addresses for "revenue" and found 2 matches.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'draft-search-report',
          label: 'Draft search report',
          status: 'completed' as const,
          summary: 'Prepared the durable workbook search report for the thread.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Workbook Search',
        text: '## Workbook Search',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-search-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'searchWorkbookQuery',
          query: 'revenue',
          limit: 10,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'searchWorkbookQuery',
      query: 'revenue',
      limit: 10,
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"runId": "wf-search-1"')
    expect(output && 'text' in output ? output.text : '').toContain('"title": "Search Workbook"')
  })

  it('starts current-sheet summary workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-sheet-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'summarizeCurrentSheet' as const,
      title: 'Summarize Current Sheet',
      summary: 'Summarized Revenue with 24 populated cells and 1 table.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-current-sheet',
          label: 'Inspect current sheet',
          status: 'completed' as const,
          summary: 'Read durable metadata for Revenue, including used range and tables.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'draft-sheet-summary',
          label: 'Draft current sheet summary',
          status: 'completed' as const,
          summary: 'Prepared the durable current-sheet summary artifact for the thread.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Current Sheet Summary',
        text: '## Current Sheet Summary',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-sheet-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'summarizeCurrentSheet',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'summarizeCurrentSheet',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"runId": "wf-sheet-1"')
    expect(output && 'text' in output ? output.text : '').toContain('"title": "Current Sheet Summary"')
  })

  it('starts sheet-scoped formula issue workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-formula-sheet-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'findFormulaIssues' as const,
      title: 'Find Formula Issues',
      summary: 'Found 2 formula issues on Sheet1 across 3 scanned formula cells.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          status: 'completed' as const,
          summary: 'Scanned 3 formula cells on Sheet1 and found 2 issues.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'draft-issue-report',
          label: 'Draft issue report',
          status: 'completed' as const,
          summary: 'Prepared the durable formula issue report for the thread.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Formula Issues',
        text: '## Formula Issues',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-formula-sheet-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'findFormulaIssues',
          sheetName: 'Sheet1',
          limit: 25,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'findFormulaIssues',
      sheetName: 'Sheet1',
      limit: 25,
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"runId": "wf-formula-sheet-1"')
  })

  it('starts highlight-formula workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-formula-highlight-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'highlightFormulaIssues' as const,
      title: 'Highlight Formula Issues',
      summary: 'Staged highlight formatting for 2 formula issues on Sheet1.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'scan-formula-cells',
          label: 'Scan formula cells',
          status: 'completed' as const,
          summary: 'Scanned 3 formula cells on Sheet1 and found 2 issues.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'stage-issue-highlights',
          label: 'Stage issue highlights',
          status: 'completed' as const,
          summary: 'Prepared 2 semantic formatting commands to highlight the detected formula issues.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Formula Issue Highlights',
        text: '## Highlighted Formula Issues',
      },
    }))

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: null,
        zeroSyncService,
        stageCommand: vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command)),
        startWorkflow,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-workflow-highlight-formula-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'highlightFormulaIssues',
          sheetName: 'Sheet1',
          limit: 25,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'highlightFormulaIssues',
      sheetName: 'Sheet1',
      limit: 25,
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "highlightFormulaIssues"')
  })
})
