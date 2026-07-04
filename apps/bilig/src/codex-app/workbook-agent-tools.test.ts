import {
  createBundle,
  createEngine,
  createZeroSyncHarness,
  describe,
  expect,
  handleWorkbookAgentToolCall,
  it,
  vi,
  workbookSummarySchema,
} from './workbook-agent-tools-test-helpers.js'

describe('workbook agent tools read surfaces', () => {
  it('reads workbook structure with sheet metadata for workbook-wide prompts', async () => {
    const engine = await createEngine()
    engine.updateRowMetadata('Sheet1', 1, 2, 24, true)
    engine.updateColumnMetadata('Sheet1', 0, 1, 110, true)
    engine.setFreezePane('Sheet1', 1, 0)
    engine.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D3' })
    engine.setSort('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D3' }, [{ keyAddress: 'B1', direction: 'desc' }])
    engine.setTable({
      name: 'Sheet1Table',
      sheetName: 'Sheet1',
      startAddress: 'A1',
      endAddress: 'D3',
      columnNames: ['Revenue', 'Formula', 'Error', 'Length'],
      headerRow: true,
      totalsRow: false,
    })
    engine.setSpillRange('Sheet1', 'F1', 2, 2)
    engine.setPivotTable('Ops Search', 'B2', {
      name: 'ImportPivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D3' },
      groupBy: ['Revenue'],
      values: [{ sourceColumn: 'Formula', summarizeBy: 'count' }],
    })
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-read-workbook',
        tool: 'bilig_read_workbook',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const payload = workbookSummarySchema.parse(JSON.parse(textItem && 'text' in textItem ? textItem.text : ''))
    expect(payload.summary).toEqual(
      expect.objectContaining({
        sheetCount: 2,
        tableCount: 1,
        pivotCount: 1,
        spillCount: 1,
        hiddenRowIndexCount: 2,
        hiddenColumnIndexCount: 1,
      }),
    )
    expect(payload.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: 'Sheet1',
          freezePane: { rows: 1, cols: 0 },
          filterCount: 1,
          sortCount: 1,
          tableCount: 1,
          spillCount: 1,
          rowMetadata: expect.objectContaining({
            hiddenIndexCount: 2,
            explicitSizeIndexCount: 2,
          }),
          columnMetadata: expect.objectContaining({
            hiddenIndexCount: 1,
            explicitSizeIndexCount: 1,
          }),
        }),
        expect.objectContaining({
          name: 'Ops Search',
          pivotCount: 1,
        }),
      ]),
    )
  })

  it('reads the current browser selection through the attached workbook context', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
              endAddress: 'D5',
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-selection',
        tool: 'bilig_read_selection',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"startAddress": "B2"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"endAddress": "D5"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"styleId"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"sheetState"')
  })

  it('reads workbook context with selection geometry and sheet state', async () => {
    const engine = await createEngine()
    engine.updateRowMetadata('Sheet1', 1, 2, 28, true)
    engine.updateColumnMetadata('Sheet1', 2, 1, 140, true)
    engine.setFreezePane('Sheet1', 1, 0)
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
              endAddress: 'D5',
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-context',
        tool: 'bilig_get_context',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"kind": "range"')
    expect(text).toContain('"cellCount": 12')
    expect(text).toContain('"freezePane": {')
    expect(text).toContain('"hiddenRows"')
    expect(text).toContain('"hiddenColumns"')
  })

  it('lists sheets and reads sheet-level workbook view metadata', async () => {
    const engine = await createEngine()
    engine.setFreezePane('Sheet1', 1, 0)
    engine.setFilter('Sheet1', { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'D3' })
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const sheetsResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-list-sheets',
        tool: 'list_sheets',
        arguments: {},
      },
    )

    expect(sheetsResponse.success).toBe(true)
    const sheetsText = sheetsResponse.contentItems[0]
    expect(sheetsText?.type).toBe('inputText')
    expect(sheetsText && 'text' in sheetsText ? sheetsText.text : '').toContain('"name": "Sheet1"')
    expect(sheetsText && 'text' in sheetsText ? sheetsText.text : '').toContain('"name": "Ops Search"')

    const sheetViewResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-sheet-view',
        tool: 'get_sheet_view',
        arguments: {
          sheetName: 'Sheet1',
        },
      },
    )

    expect(sheetViewResponse.success).toBe(true)
    const sheetViewText = sheetViewResponse.contentItems[0]
    expect(sheetViewText?.type).toBe('inputText')
    expect(sheetViewText && 'text' in sheetViewText ? sheetViewText.text : '').toContain('"freezePane": {')
    expect(sheetViewText && 'text' in sheetViewText ? sheetViewText.text : '').toContain('"filters": [')
  })

  it('reads used range, current region, and axis metadata', async () => {
    const engine = await createEngine()
    engine.updateRowMetadata('Sheet1', 1, 2, 28, true)
    engine.updateColumnMetadata('Sheet1', 0, 1, 120, true)
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const usedRangeResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-used-range',
        tool: 'get_used_range',
        arguments: {
          sheetName: 'Sheet1',
        },
      },
    )

    expect(usedRangeResponse.success).toBe(true)
    const usedRangeText = usedRangeResponse.contentItems[0]
    expect(usedRangeText?.type).toBe('inputText')
    expect(usedRangeText && 'text' in usedRangeText ? usedRangeText.text : '').toContain('"startAddress": "A1"')

    const currentRegionResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-current-region',
        tool: 'get_current_region',
        arguments: {
          sheetName: 'Sheet1',
          address: 'A1',
        },
      },
    )

    expect(currentRegionResponse.success).toBe(true)
    const currentRegionText = currentRegionResponse.contentItems[0]
    expect(currentRegionText?.type).toBe('inputText')
    expect(currentRegionText && 'text' in currentRegionText ? currentRegionText.text : '').toContain('"derivedA1Ranges": [')

    const rowMetadataResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-row-metadata',
        tool: 'get_row_metadata',
        arguments: {
          sheetName: 'Sheet1',
        },
      },
    )

    expect(rowMetadataResponse.success).toBe(true)
    const rowMetadataText = rowMetadataResponse.contentItems[0]
    expect(rowMetadataText?.type).toBe('inputText')
    expect(rowMetadataText && 'text' in rowMetadataText ? rowMetadataText.text : '').toContain('"hidden": true')

    const columnMetadataResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-column-metadata',
        tool: 'get_column_metadata',
        arguments: {
          sheetName: 'Sheet1',
        },
      },
    )

    expect(columnMetadataResponse.success).toBe(true)
    const columnMetadataText = columnMetadataResponse.contentItems[0]
    expect(columnMetadataText?.type).toBe('inputText')
    expect(columnMetadataText && 'text' in columnMetadataText ? columnMetadataText.text : '').toContain('"size": 120')
  })

  it('reads the visible viewport through the attached browser context', async () => {
    const engine = await createEngine()
    engine.setFreezePane('Sheet1', 1, 0)
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
            rowEnd: 1,
            colStart: 0,
            colEnd: 1,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-visible',
        tool: 'bilig_read_visible_range',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"startAddress": "A1"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"endAddress": "B2"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"freezePane": {')
  })

  it('inspects one cell with formula lineage and runtime metadata', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B1',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 5,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-inspect',
        tool: 'bilig_inspect_cell',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"address": "B1"')
    expect(text).toContain('"formula": "=SUM(A1:A1)"')
    expect(text).toContain('"directPrecedents": [')
    expect(text).toContain('Sheet1!A1')
    expect(text).toContain('"style": {')
    expect(text).toContain('"backgroundColor": "#fef3c7"')
    expect(text).toContain('"numberFormat": {')
  })

  it('scans formula issues through the warm local runtime', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-formula-issues',
        tool: 'bilig_find_formula_issues',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"issueCount": 2')
    expect(text).toContain('"address": "C1"')
    expect(text).toContain('"errorText": "#DIV/0!"')
    expect(text).toContain('"address": "D1"')
    expect(text).toContain('"unsupported"')
  })

  it('searches workbook sheets, cells, formulas, and values', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-search',
        tool: 'bilig_search_workbook',
        arguments: {
          query: 'gross margin',
        },
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"query": "gross margin"')
    expect(text).toContain('"address": "A2"')
    expect(text).toContain('"snippet": "Gross Margin"')
  })

  it('reads recent durable workbook changes', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const listWorkbookChanges = vi.fn(async () => [
      {
        revision: 12,
        actorUserId: 'alex@example.com',
        clientMutationId: null,
        eventKind: 'applyAgentCommandBundle' as const,
        summary: 'Applied workbook change set at revision r12',
        sheetId: 1,
        sheetName: 'Sheet1',
        anchorAddress: 'B2',
        range: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C4',
        },
        rangeInvalid: false,
        undoBundle: null,
        revertedByRevision: null,
        revertsRevision: null,
        createdAtUnixMs: 1_234,
      },
    ])
    zeroSyncService.listWorkbookChanges = listWorkbookChanges

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
        callId: 'call-recent-changes',
        tool: 'bilig_read_recent_changes',
        arguments: {
          limit: 5,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(listWorkbookChanges).toHaveBeenCalledWith('doc-1', 5)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"changeCount": 1')
    expect(text).toContain('"revision": 12')
    expect(text).toContain('"summary": "Applied workbook change set at revision r12"')
    expect(text).toContain('"startAddress": "B2"')
  })

  it('traces multi-hop workbook dependencies from the attached selection', async () => {
    const engine = await createEngine()
    engine.setCellFormula('Sheet1', 'E1', 'B1*2')
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
      {
        documentId: 'doc-1',
        session: {
          userID: 'alex@example.com',
          roles: ['editor'],
        },
        uiContext: {
          selection: {
            sheetName: 'Sheet1',
            address: 'B1',
          },
          viewport: {
            rowStart: 0,
            rowEnd: 10,
            colStart: 0,
            colEnd: 5,
          },
        },
        zeroSyncService,
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-trace',
        tool: 'bilig_trace_dependencies',
        arguments: {
          direction: 'both',
          depth: 2,
        },
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"address": "B1"')
    expect(text).toContain('"precedentCount": 1')
    expect(text).toContain('"dependentCount": 1')
    expect(text).toContain('"address": "A1"')
    expect(text).toContain('"address": "E1"')
  })

  it('reads workbook ranges through the authoritative runtime', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const response = await handleWorkbookAgentToolCall(
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
        stageCommand: vi.fn(async () => createBundle({ kind: 'createSheet', name: 'unused' })),
      },
      {
        threadId: 'thr-1',
        turnId: 'turn-1',
        callId: 'call-1',
        tool: 'bilig_read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'B1',
        },
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"address": "A1"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"value": 42')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"formula": "=SUM(A1:A1)"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"styles": [')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"backgroundColor": "#fef3c7"')
    expect(textItem && 'text' in textItem ? textItem.text : '').toContain('"numberFormats": [')
  })

  it('reads discontiguous selector results as ordered range sets', async () => {
    const engine = await createEngine()
    engine.setCellValue('Sheet1', 'A1', 'Revenue')
    engine.setCellValue('Sheet1', 'B1', 'Margin')
    engine.setCellValue('Sheet1', 'A2', 10)
    engine.setCellValue('Sheet1', 'B2', 2)
    engine.setCellValue('Sheet1', 'A3', 12)
    engine.setCellValue('Sheet1', 'B3', 3)
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
        callId: 'call-read-row-query',
        tool: 'read_range',
        arguments: {
          selector: {
            kind: 'rowQuery',
            sheet: 'Sheet1',
            predicate: {
              column: 'Revenue',
              op: 'gte',
              value: 10,
            },
          },
        },
      },
    )

    expect(response.success).toBe(true)
    const textItem = response.contentItems[0]
    expect(textItem?.type).toBe('inputText')
    const text = textItem && 'text' in textItem ? textItem.text : ''
    expect(text).toContain('"rangeCount": 2')
    expect(text).toContain('"startAddress": "A2"')
    expect(text).toContain('"endAddress": "D3"')
  })

  it('lists named ranges and tables from the authoritative runtime', async () => {
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

    const namedRangesResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-list-named-ranges',
        tool: 'list_named_ranges',
        arguments: {},
      },
    )

    expect(namedRangesResponse.success).toBe(true)
    const namedRangesText = namedRangesResponse.contentItems[0]
    expect(namedRangesText?.type).toBe('inputText')
    expect(namedRangesText && 'text' in namedRangesText ? namedRangesText.text : '').toContain('"name": "Inputs"')
    expect(namedRangesText && 'text' in namedRangesText ? namedRangesText.text : '').toContain('"startAddress": "A1"')

    const tablesResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-list-tables',
        tool: 'list_tables',
        arguments: {},
      },
    )

    expect(tablesResponse.success).toBe(true)
    const tablesText = tablesResponse.contentItems[0]
    expect(tablesText?.type).toBe('inputText')
    expect(tablesText && 'text' in tablesText ? tablesText.text : '').toContain('"name": "RevenueTable"')
    expect(tablesText && 'text' in tablesText ? tablesText.text : '').toContain('"columnNames": [')
  })

  it('lists workbook pivots from the authoritative runtime', async () => {
    const engine = await createEngine()
    engine.setPivotTable('Sheet1', 'E2', {
      name: 'RevenuePivot',
      source: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B3' },
      groupBy: ['Revenue'],
      values: [{ sourceColumn: 'Margin', summarizeBy: 'sum' }],
    })
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
        callId: 'call-list-pivots',
        tool: 'list_pivots',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"name": "RevenuePivot"')
    expect(text && 'text' in text ? text.text : '').toContain('"groupBy": [')
  })

  it('lists workbook data validation rules from the authoritative runtime', async () => {
    const engine = await createEngine()
    engine.setDataValidation({
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
      errorStyle: 'stop',
      errorTitle: 'Status required',
      errorMessage: 'Pick Draft or Final.',
    })
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
        callId: 'call-list-data-validations',
        tool: 'list_data_validation_rules',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"kind": "list"')
    expect(text && 'text' in text ? text.text : '').toContain('"startAddress": "B2"')
  })

  it('includes intersecting data validation metadata in read_range inspection', async () => {
    const engine = await createEngine()
    engine.setDataValidation({
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
    })
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
        callId: 'call-read-range-with-validation',
        tool: 'read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"dataValidations": [')
    expect(text && 'text' in text ? text.text : '').toContain('"Draft"')
  })

  it('lists workbook conditional formats from the authoritative runtime', async () => {
    const engine = await createEngine()
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
        callId: 'call-get-conditional-formats',
        tool: 'get_conditional_formats',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"id": "cf-1"')
    expect(text && 'text' in text ? text.text : '').toContain('"greaterThan"')
  })

  it('includes intersecting conditional format metadata in read_range inspection', async () => {
    const engine = await createEngine()
    engine.setConditionalFormat({
      id: 'cf-1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'B2',
        endAddress: 'B4',
      },
      rule: {
        kind: 'textContains',
        text: 'urgent',
      },
      style: {
        font: { bold: true },
      },
    })
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
        callId: 'call-read-range-with-conditional-format',
        tool: 'read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'B4',
        },
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"conditionalFormats": [')
    expect(text && 'text' in text ? text.text : '').toContain('"urgent"')
  })

  it('lists workbook protection status from the authoritative runtime', async () => {
    const engine = await createEngine()
    engine.setSheetProtection({ sheetName: 'Sheet1', hideFormulas: true })
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'B2',
      },
      hideFormulas: true,
    })
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
        callId: 'call-get-protection-status',
        tool: 'get_protection_status',
        arguments: { sheetName: 'Sheet1' },
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"hideFormulas": true')
    expect(text && 'text' in text ? text.text : '').toContain('"protect-a1"')
  })

  it('masks hidden formulas in read_range and inspect_cell outputs', async () => {
    const engine = await createEngine()
    engine.setCellFormula('Sheet1', 'A1', '2+2')
    engine.setRangeProtection({
      id: 'protect-a1',
      range: {
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'A1',
      },
      hideFormulas: true,
    })
    const { zeroSyncService } = createZeroSyncHarness(engine)

    const rangeResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-read-range-hidden-formula',
        tool: 'read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'A1',
          endAddress: 'A1',
        },
      },
    )

    expect(rangeResponse.success).toBe(true)
    const rangeText = rangeResponse.contentItems[0]
    expect(rangeText?.type).toBe('inputText')
    expect(rangeText && 'text' in rangeText ? rangeText.text : '').toContain('"rangeProtections": [')
    expect(rangeText && 'text' in rangeText ? rangeText.text : '').toContain('"formula": null')

    const inspectResponse = await handleWorkbookAgentToolCall(
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
        callId: 'call-inspect-cell-hidden-formula',
        tool: 'inspect_cell',
        arguments: {
          sheetName: 'Sheet1',
          address: 'A1',
        },
      },
    )

    expect(inspectResponse.success).toBe(true)
    const inspectText = inspectResponse.contentItems[0]
    expect(inspectText?.type).toBe('inputText')
    expect(inspectText && 'text' in inspectText ? inspectText.text : '').toContain('"formula": null')
  })

  it('lists workbook comments and notes from the authoritative runtime', async () => {
    const engine = await createEngine()
    engine.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'B2',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })
    engine.setNote({
      sheetName: 'Sheet1',
      address: 'C3',
      text: 'Manual override',
    })
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
        callId: 'call-get-comments',
        tool: 'get_comments',
        arguments: {},
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"threadId": "thread-1"')
    expect(text && 'text' in text ? text.text : '').toContain('"text": "Manual override"')
  })

  it('includes intersecting comment threads and notes in read_range inspection', async () => {
    const engine = await createEngine()
    engine.setCommentThread({
      threadId: 'thread-1',
      sheetName: 'Sheet1',
      address: 'B2',
      comments: [{ id: 'comment-1', body: 'Check this total.' }],
    })
    engine.setNote({
      sheetName: 'Sheet1',
      address: 'C3',
      text: 'Manual override',
    })
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
        callId: 'call-read-range-with-annotations',
        tool: 'read_range',
        arguments: {
          sheetName: 'Sheet1',
          startAddress: 'B2',
          endAddress: 'C3',
        },
      },
    )

    expect(response.success).toBe(true)
    const text = response.contentItems[0]
    expect(text?.type).toBe('inputText')
    expect(text && 'text' in text ? text.text : '').toContain('"commentThreads": [')
    expect(text && 'text' in text ? text.text : '').toContain('"notes": [')
    expect(text && 'text' in text ? text.text : '').toContain('"Check this total."')
    expect(text && 'text' in text ? text.text : '').toContain('"Manual override"')
  })
})
