import type { JsonValue, WorkbookAgentCommandBundle, WorkbookChangeRecord } from './workbook-agent-tools-test-helpers.js'
import {
  applyWorkbookAgentCommandBundleWithUndoCapture,
  createBundle,
  createDerivedBundle,
  createEngine,
  createZeroSyncHarness,
  describe,
  expect,
  handleWorkbookAgentToolCall,
  it,
  readToolJson,
  vi,
  z,
} from './workbook-agent-tools-test-helpers.js'

describe('workbook agent tools selector-aware mutation staging', () => {
  it('resolves selectors for read and mutation tools before staging commands', async () => {
    const engine = await createEngine()
    engine.setDefinedName('Inputs', {
      kind: 'range-ref',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B1',
    })
    engine.setTable({
      name: 'RevenueTable',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Revenue', 'Margin'],
      headerRow: true,
      totalsRow: false,
    })
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const readResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A2',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 3,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-selector-read-range',
        tool: 'read_range',
        arguments: {
          selector: {
            kind: 'namedRange',
            name: 'Inputs',
          },
        },
      },
    )

    expect(readResponse.success).toBe(true)
    const readText = readResponse.contentItems[0]
    expect(readText?.type).toBe('inputText')
    expect(readText && 'text' in readText ? readText.text : '').toContain('"displayLabel": "Inputs"')
    expect(readText && 'text' in readText ? readText.text : '').toContain('"startAddress": "A1"')

    const formatResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-selector-format-range',
        tool: 'format_range',
        arguments: {
          selector: {
            kind: 'tableColumn',
            table: 'RevenueTable',
            column: 'Margin',
          },
          patch: {
            font: {
              bold: true,
            },
          },
        },
      },
    )

    expect(formatResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'formatRange',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B3',
      },
      patch: {
        font: {
          bold: true,
        },
      },
    })
  })

  it('resolves selectors for structural range tools before staging commands', async () => {
    const engine = await createEngine()
    engine.setTable({
      name: 'RevenueTable',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Revenue', 'Margin'],
      headerRow: true,
      totalsRow: false,
    })
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const setFilterResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-selector-set-filter',
        tool: 'set_filter',
        arguments: {
          selector: {
            kind: 'table',
            table: 'RevenueTable',
          },
        },
      },
    )

    expect(setFilterResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'setFilter',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B3',
      },
    })

    const setSortResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-selector-set-sort',
        tool: 'set_sort',
        arguments: {
          selector: {
            kind: 'tableColumn',
            table: 'RevenueTable',
            column: 'Margin',
          },
          keys: [{ keyAddress: 'B2', direction: 'desc' }],
        },
      },
    )

    expect(setSortResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'setSort',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B3',
      },
      keys: [{ keyAddress: 'B2', direction: 'desc' }],
    })
  })

  it('stages named range, table, and pivot object commands', async () => {
    const engine = await createEngine()
    engine.setTable({
      name: 'RevenueTable',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Revenue', 'Margin'],
      headerRow: true,
      totalsRow: false,
    })
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const namedRangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A2',
            range: {
              startAddress: 'A2',
              endAddress: 'B3',
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-create-named-range',
        tool: 'create_named_range',
        arguments: {
          name: 'Inputs',
          selector: {
            kind: 'currentSelection',
          },
        },
      },
    )

    expect(namedRangeResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'upsertDefinedName',
      name: 'Inputs',
      value: {
        kind: 'range-ref',
        sheetName: 'Sheet1',
        startAddress: 'A2',
        endAddress: 'B3',
      },
    })

    const tableResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-create-table',
        tool: 'create_table',
        arguments: {
          name: 'RegionTable',
          selector: {
            kind: 'table',
            table: 'RevenueTable',
            sheet: 'Sheet1',
          },
          headerRow: true,
        },
      },
    )

    expect(tableResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenNthCalledWith(2, {
      kind: 'upsertTable',
      table: {
        name: 'RegionTable',
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B3',
        columnNames: ['Revenue', 'Margin'],
        headerRow: true,
        totalsRow: false,
      },
    })

    const pivotResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-create-pivot',
        tool: 'create_pivot_table',
        arguments: {
          name: 'RevenuePivot',
          sheetName: 'Sheet1',
          address: 'E2',
          selector: {
            kind: 'table',
            table: 'RevenueTable',
            sheet: 'Sheet1',
          },
          groupBy: ['Revenue'],
          values: [{ sourceColumn: 'Margin', summarizeBy: 'sum' }],
        },
      },
    )

    expect(pivotResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'upsertPivotTable',
      pivot: {
        name: 'RevenuePivot',
        sheetName: 'Sheet1',
        address: 'E2',
        source: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B3',
        },
        groupBy: ['Revenue'],
        values: [{ sourceColumn: 'Margin', summarizeBy: 'sum' }],
        rows: 1,
        cols: 2,
      },
    })
  })

  it('stages selector-aware data validation commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const createResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
            range: {
              startAddress: 'B2',
              endAddress: 'B4',
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-create-data-validation',
        tool: 'create_data_validation',
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
          rule: {
            kind: 'list',
            values: ['Draft', 'Final'],
          },
          allowBlank: false,
          showDropdown: true,
        },
      },
    )

    expect(createResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'setDataValidation',
      validation: {
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'list',
          values: ['Draft', 'Final'],
        },
        allowBlank: false,
        showDropdown: true,
      },
    })

    const removeResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-remove-data-validation',
        tool: 'remove_data_validation',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B4',
          },
        },
      },
    )

    expect(removeResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'clearDataValidation',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
    })
  })

  it('stages comment and note commands against single-cell selector targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const addCommentResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-add-comment',
        tool: 'add_comment',
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
          text: 'Check this total.',
        },
      },
    )

    expect(addCommentResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'upsertCommentThread',
      thread: {
        threadId: expect.any(String),
        sheetName: 'Sheet1',
        address: 'B2',
        comments: [{ id: expect.any(String), body: 'Check this total.' }],
      },
    })

    engine.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'B2',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })

    const replyCommentResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-reply-comment',
        tool: 'reply_comment',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'B2',
            endAddress: 'B2',
          },
          text: 'Looks good.',
        },
      },
    )

    expect(replyCommentResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'upsertCommentThread',
      thread: {
        threadId: 'thread-1',
        sheetName: 'Sheet1',
        address: 'B2',
        comments: [
          { id: 'comment-1', body: 'Check this total.' },
          { id: expect.any(String), body: 'Looks good.' },
        ],
      },
    })

    const addNoteResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-add-note',
        tool: 'add_note',
        arguments: {
          range: {
            sheetName: 'Sheet1',
            startAddress: 'C3',
            endAddress: 'C3',
          },
          text: 'Manual override',
        },
      },
    )

    expect(addNoteResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'upsertNote',
      note: {
        sheetName: 'Sheet1',
        address: 'C3',
        text: 'Manual override',
      },
    })
  })

  it('stages conditional format commands against selector and id targets', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const addResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
            range: {
              startAddress: 'B2',
              endAddress: 'B4',
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-add-conditional-format',
        tool: 'add_conditional_format',
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
          rule: {
            kind: 'cellIs',
            operator: 'greaterThan',
            values: [10],
          },
          style: {
            fill: {
              backgroundColor: '#ff0000',
            },
          },
        },
      },
    )

    expect(addResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'upsertConditionalFormat',
      format: {
        id: expect.any(String),
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        rule: {
          kind: 'cellIs',
          operator: 'greaterThan',
          values: [10],
        },
        style: {
          fill: {
            backgroundColor: '#ff0000',
          },
        },
      },
    })

    engine.setConditionalFormat({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'cellIs',
        operator: 'greaterThan',
        values: [10],
      },
      style: {
        fill: { backgroundColor: '#ff0000' },
      },
    })

    const removeResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-remove-conditional-format',
        tool: 'remove_conditional_format',
        arguments: {
          id: 'cf-1',
        },
      },
    )

    expect(removeResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'deleteConditionalFormat',
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
    })
  })

  it('stages sheet and range protection commands', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => createBundle(command))

    const protectSheetResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'A1',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-protect-sheet',
        tool: 'protect_sheet',
        arguments: {
          hideFormulas: true,
        },
      },
    )

    expect(protectSheetResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'setSheetProtection',
      protection: {
        sheetName: 'Sheet1',
        hideFormulas: true,
      },
    })

    const protectRangeResponse = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B2',
            range: {
              startAddress: 'B2',
              endAddress: 'B4',
            },
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand,
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-protect-range',
        tool: 'protect_range',
        arguments: {
          selector: {
            kind: 'currentSelection',
          },
          hideFormulas: true,
        },
      },
    )

    expect(protectRangeResponse.success).toBe(true)
    expect(stageCommand).toHaveBeenLastCalledWith({
      kind: 'upsertRangeProtection',
      protection: {
        id: expect.any(String),
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
        hideFormulas: true,
      },
    })
  })

  it('writes rectangular ranges through renderCommit with normalized formulas', async () => {
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
        callId: 'call-2',
        tool: 'bilig_write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'C3',
          values: [[1, { formula: '=SUM(A1:A1)' }]],
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'writeRange',
      sheetName: 'Sheet1',
      startAddress: 'C3',
      values: [[1, { formula: '=SUM(A1:A1)' }]],
    })
    expect(response.contentItems).toEqual([
      expect.objectContaining({
        type: 'inputText',
        text: expect.stringContaining('"staged": true'),
      }),
    ])
    expect(response.contentItems[0] && 'text' in response.contentItems[0] ? response.contentItems[0].text : '').toContain(
      '"bundleId": "bundle-1"',
    )
  })

  it('coerces write_range numeric text and typed blanks/formulas before apply verification', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => {
      const bundle = createBundle(command)
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      return {
        bundle,
        executionRecord: {
          id: 'run-1',
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
          appliedRevision: 2,
          appliedBy: 'auto' as const,
          baseRevision: bundle.baseRevision,
          context: bundle.context,
          commands: bundle.commands,
          preview: {
            ranges: [],
            structuralChanges: [],
            cellDiffs: [],
            effectSummary: {
              displayedCellDiffCount: 0,
              truncatedCellDiffs: false,
              inputChangeCount: 0,
              formulaChangeCount: 0,
              styleChangeCount: 0,
              numberFormatChangeCount: 0,
              structuralChangeCount: 0,
            },
          },
          createdAtUnixMs: 2,
          appliedAtUnixMs: 2,
        },
      }
    })

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
        callId: 'call-write-typed-values',
        tool: 'write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'A10',
          values: [
            ['Cost', 'Months', 'Monthly', 'Blank'],
            ['1200', { type: 'number', value: '12' }, { formula: '=A11/B11' }, { type: 'blank' }],
          ],
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'writeRange',
        values: [
          ['Cost', 'Months', 'Monthly', 'Blank'],
          [1200, 12, { formula: '=A11/B11' }, null],
        ],
      }),
    )
    expect(engine.getCell('Sheet1', 'C11').value).toMatchObject({ value: 100 })
    expect(engine.getCell('Sheet1', 'C11').value).not.toHaveProperty('code')
  })

  it('returns deterministic mutation receipts when an applied write lacks rendered proof', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const workbookChanges: WorkbookChangeRecord[] = [
      {
        revision: 2,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle',
        summary: 'Write cells in Sheet1!D10',
        sheetId: null,
        sheetName: 'Sheet1',
        anchorAddress: 'D10',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'D10',
          endAddress: 'D10',
        },
        rangeInvalid: false,
        undoBundle: {
          kind: 'engineOps',
          ops: [],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 2,
      },
    ]
    zeroSyncService.listWorkbookChanges = vi.fn(async () => workbookChanges)
    const stageCommand = vi.fn(async (command: WorkbookAgentCommandBundle['commands'][number]) => {
      const bundle = createDerivedBundle(command)
      applyWorkbookAgentCommandBundleWithUndoCapture(engine, bundle)
      return {
        bundle,
        executionRecord: {
          id: 'run-receipt',
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
          appliedRevision: 2,
          context: bundle.context,
          commands: bundle.commands,
          preview: {
            ranges: bundle.affectedRanges,
            structuralChanges: [],
            cellDiffs: [
              {
                sheetName: 'Sheet1',
                address: 'D10',
                beforeInput: null,
                beforeFormula: null,
                afterInput: 'Verified receipt',
                afterFormula: null,
                changeKinds: ['input' as const],
              },
            ],
            effectSummary: {
              displayedCellDiffCount: 1,
              truncatedCellDiffs: false,
              inputChangeCount: 1,
              formulaChangeCount: 0,
              styleChangeCount: 0,
              numberFormatChangeCount: 0,
              structuralChangeCount: 0,
            },
          },
          createdAtUnixMs: 2,
          appliedAtUnixMs: 2,
        },
      }
    })

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
        callId: 'call-write-receipt',
        tool: 'write_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'D10',
          values: [['Verified receipt']],
        },
      },
    )

    expect(response.success).toBe(true)
    const payload = z
      .object({
        status: z.literal('verification_incomplete'),
        mutationReceipt: z.object({
          status: z.literal('verification_incomplete'),
          revision: z.object({
            before: z.literal(1),
            after: z.literal(2),
          }),
          authoritativeReadback: z.object({
            requested: z.literal(true),
            matched: z.literal(true),
          }),
          renderedReadback: z.object({
            requested: z.literal(true),
            matched: z.null(),
            incompleteReason: z.string(),
          }),
          undo: z.object({
            available: z.literal(true),
            token: z.literal('revision:2'),
          }),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.mutationReceipt.renderedReadback.incompleteReason).toContain('No browser-rendered context')
  })

  it('undoes a specific workbook mutation revision and returns verification context', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'D10', 'temporary value')
    const { zeroSyncService } = createZeroSyncHarness(engine)
    let headRevision = 2
    const applyServerMutator = vi.fn(async (name: string, args: unknown): Promise<void> => {
      expect(name).toBe('workbook.revertChange')
      expect(args).toMatchObject({
        documentId: 'doc-1',
        revision: 2,
      })
      engine.clearCell('Sheet1', 'D10')
      headRevision = 3
    })
    zeroSyncService.applyServerMutator = applyServerMutator
    zeroSyncService.getWorkbookHeadRevision = vi.fn(async () => headRevision)
    const workbookChanges: WorkbookChangeRecord[] = [
      {
        revision: 2,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle',
        summary: 'Write cells in Sheet1!D10',
        sheetId: null,
        sheetName: 'Sheet1',
        anchorAddress: 'D10',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'D10',
          endAddress: 'D10',
        },
        rangeInvalid: false,
        undoBundle: {
          kind: 'engineOps',
          ops: [],
        },
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 2,
      },
    ]
    zeroSyncService.listWorkbookChanges = vi.fn(async () => workbookChanges)

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
        callId: 'call-undo-revision',
        tool: 'undo_workbook_mutation',
        arguments: {
          revision: 2,
        },
      },
    )

    expect(response.success).toBe(true)
    const payload = z
      .object({
        undone: z.literal(true),
        applied: z.literal(false),
        staged: z.literal(false),
        queuedForTurnApply: z.literal(false),
        status: z.literal('verification_incomplete'),
        verificationComplete: z.literal(false),
        revision: z.object({
          before: z.literal(2),
          after: z.literal(3),
          reverted: z.literal(2),
        }),
        targetChange: z.object({
          revision: z.literal(2),
          range: z.object({
            sheetName: z.literal('Sheet1'),
            startAddress: z.literal('D10'),
            endAddress: z.literal('D10'),
          }),
        }),
        verification: z.object({
          appliedRevision: z.literal(3),
          authoritativeReadback: z.array(
            z.object({
              rows: z.array(z.array(z.object({ value: z.null() }))),
            }),
          ),
        }),
      })
      .parse(readToolJson(response))
    expect(payload.verification.authoritativeReadback[0]?.rows[0]?.[0]?.value).toBeNull()
    expect(applyServerMutator).toHaveBeenCalledTimes(1)
  })

  it('stages selector-aware formula writes through set_formula', async () => {
    const engine = await createEngine()
    engine.setTable({
      name: 'RevenueTable',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'B3',
      columnNames: ['Revenue', 'Margin'],
      headerRow: true,
      totalsRow: false,
    })
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
        callId: 'call-set-formula',
        tool: 'set_formula',
        arguments: {
          selector: {
            kind: 'tableColumn',
            table: 'RevenueTable',
            column: 'Margin',
            sheet: 'Sheet1',
          },
          formulas: [['=A2*0.2'], ['=A3*0.25']],
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'setRangeFormulas',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B3',
      },
      formulas: [['=A2*0.2'], ['=A3*0.25']],
    })
  })

  it('normalizes stale sheet context after a Sheet3 rename before read tools answer', async () => {
    const engine = await createEngine()
    engine.createSheet('Sheet3')
    engine.setCellValue('Sheet3', 'A1', 'Prepaid')
    engine.renameSheet('Sheet3', 'Prepaid Template')
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const staleContext = {
      selection: {
        sheetName: 'Sheet3',
        address: 'A1',
      },
      viewport: {
        rowStart: 0,
        rowEnd: 10,
        colStart: 0,
        colEnd: 5,
      },
    }

    const callTool = (tool: string, args: JsonValue = {}) =>
      handleWorkbookAgentToolCall(
        {
          documentId: 'doc-1',
          session: {
            userID: 'alex@example.com',
            roles: ['editor'],
          },
          uiContext: staleContext,
          zeroSyncService,
          stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
        },
        {
          threadId: 'thr-1',
          turnId: 'turn-1',
          callId: `call-${tool}`,
          tool,
          arguments: args,
        },
      )

    const contextResponse = await callTool('get_context')
    const workbookResponse = await callTool('read_workbook')
    const sheetsResponse = await callTool('list_sheets')
    const viewResponse = await callTool('get_sheet_view', { sheetName: 'Prepaid Template' })

    for (const response of [contextResponse, workbookResponse, sheetsResponse, viewResponse]) {
      expect(response.success).toBe(true)
      const text = response.contentItems[0]
      expect(text && 'text' in text ? text.text : '').not.toContain('Sheet3')
    }
    for (const response of [workbookResponse, sheetsResponse, viewResponse]) {
      const text = response.contentItems[0]
      expect(text && 'text' in text ? text.text : '').toContain('Prepaid Template')
    }
  })
})
