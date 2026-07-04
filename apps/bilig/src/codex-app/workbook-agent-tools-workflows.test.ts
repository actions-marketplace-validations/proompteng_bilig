import type { WorkbookAgentCommandBundle } from './workbook-agent-tools-test-helpers.js'
import {
  createBundle,
  createEngine,
  createZeroSyncHarness,
  describe,
  expect,
  handleWorkbookAgentToolCall,
  it,
  vi,
} from './workbook-agent-tools-test-helpers.js'

describe('workbook agent tools workflow starts and metadata commands', () => {
  it('starts repair-formula workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-formula-repair-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'repairFormulaIssues' as const,
      title: 'Repair Formula Issues',
      summary: 'Staged 1 formula repair on Sheet1 from nearby healthy formulas.',
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
          summary: 'Scanned 2 formula cells on Sheet1 and found 1 issue.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'stage-formula-repairs',
          label: 'Stage formula repairs',
          status: 'completed' as const,
          summary: 'Prepared 1 semantic write command for the repair change set.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Formula Repair Preview',
        text: '## Formula Repair Preview',
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
        callId: 'call-workflow-repair-formula-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'repairFormulaIssues',
          sheetName: 'Sheet1',
          limit: 25,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'repairFormulaIssues',
      sheetName: 'Sheet1',
      limit: 25,
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "repairFormulaIssues"')
  })

  it('starts outlier-highlight workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-outlier-highlight-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'highlightCurrentSheetOutliers' as const,
      title: 'Highlight Current Sheet Outliers',
      summary: 'Staged outlier highlights for 2 cells across 1 numeric column on Revenue.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-numeric-columns',
          label: 'Inspect numeric columns',
          status: 'completed' as const,
          summary: 'Loaded numeric cells and header labels from Revenue.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Current Sheet Outlier Highlights',
        text: '## Highlighted Numeric Outliers',
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
        callId: 'call-workflow-outlier-highlight-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'highlightCurrentSheetOutliers',
          sheetName: 'Revenue',
          limit: 10,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'highlightCurrentSheetOutliers',
      sheetName: 'Revenue',
      limit: 10,
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "highlightCurrentSheetOutliers"')
  })

  it('starts header-normalization workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-header-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'normalizeCurrentSheetHeaders' as const,
      title: 'Normalize Current Sheet Headers',
      summary: 'Staged normalized headers for 2 cells on Imports.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-header-row',
          label: 'Inspect header row',
          status: 'completed' as const,
          summary: 'Loaded the used range and current header row from Imports.',
          updatedAtUnixMs: 1,
        },
        {
          stepId: 'stage-header-normalization',
          label: 'Stage header normalization',
          status: 'completed' as const,
          summary: 'Prepared the semantic write preview that normalizes 2 header cells.',
          updatedAtUnixMs: 2,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Header Normalization Preview',
        text: '## Header Normalization Preview',
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
        callId: 'call-workflow-header-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'normalizeCurrentSheetHeaders',
          sheetName: 'Imports',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'normalizeCurrentSheetHeaders',
      sheetName: 'Imports',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "normalizeCurrentSheetHeaders"')
  })

  it('starts number-format-normalization workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-number-format-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'normalizeCurrentSheetNumberFormats' as const,
      title: 'Normalize Current Sheet Number Formats',
      summary: 'Staged normalized number formats for 3 columns on Imports.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-number-columns',
          label: 'Inspect numeric columns',
          status: 'completed' as const,
          summary: 'Loaded numeric cells and header labels from Imports.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Number Format Normalization Preview',
        text: '## Number Format Normalization Preview',
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
        callId: 'call-workflow-number-format-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'normalizeCurrentSheetNumberFormats',
          sheetName: 'Imports',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'normalizeCurrentSheetNumberFormats',
      sheetName: 'Imports',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "normalizeCurrentSheetNumberFormats"')
  })

  it('starts whitespace-normalization workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-whitespace-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'normalizeCurrentSheetWhitespace' as const,
      title: 'Normalize Current Sheet Whitespace',
      summary: 'Staged normalized whitespace for 3 text cells on Imports.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-text-cells',
          label: 'Inspect text cells',
          status: 'completed' as const,
          summary: 'Loaded the used range and string cells from Imports.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Whitespace Normalization Preview',
        text: '## Whitespace Normalization Preview',
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
        callId: 'call-workflow-whitespace-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'normalizeCurrentSheetWhitespace',
          sheetName: 'Imports',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'normalizeCurrentSheetWhitespace',
      sheetName: 'Imports',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "normalizeCurrentSheetWhitespace"')
  })

  it('starts formula fill-down workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-fill-formulas-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'fillCurrentSheetFormulasDown' as const,
      title: 'Fill Current Sheet Formulas Down',
      summary: 'Staged formula fill-down for 1 column on Imports.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-formula-columns',
          label: 'Inspect formula columns',
          status: 'completed' as const,
          summary: 'Loaded formula cells and blank fill gaps from Imports.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Formula Fill-Down Preview',
        text: '## Formula Fill-Down Preview',
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
        callId: 'call-workflow-fill-formulas-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'fillCurrentSheetFormulasDown',
          sheetName: 'Imports',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'fillCurrentSheetFormulasDown',
      sheetName: 'Imports',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "fillCurrentSheetFormulasDown"')
  })

  it('starts header-style workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-style-headers-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'styleCurrentSheetHeaders' as const,
      title: 'Style Current Sheet Headers',
      summary: 'Staged a consistent header style preview for Imports.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-header-row',
          label: 'Inspect header row',
          status: 'completed' as const,
          summary: 'Loaded the used range and header row from Imports.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Header Style Preview',
        text: '## Header Style Preview',
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
        callId: 'call-workflow-style-headers-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'styleCurrentSheetHeaders',
          sheetName: 'Imports',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'styleCurrentSheetHeaders',
      sheetName: 'Imports',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "styleCurrentSheetHeaders"')
  })

  it('starts current-sheet review-tab workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-review-tab-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'createCurrentSheetReviewTab' as const,
      title: 'Create Current Sheet Review Tab',
      summary: 'Staged a review-tab preview for Revenue into Revenue Review.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-source-sheet',
          label: 'Inspect source sheet',
          status: 'completed' as const,
          summary: 'Loaded the used range from Revenue.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Current Sheet Review Tab Preview',
        text: '## Current Sheet Review Tab Preview',
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
        callId: 'call-workflow-review-tab-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'createCurrentSheetReviewTab',
          sheetName: 'Revenue',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'createCurrentSheetReviewTab',
      sheetName: 'Revenue',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "createCurrentSheetReviewTab"')
  })

  it('starts current-sheet rollup workflows from the semantic tool surface', async () => {
    const engine = await createEngine()
    const { zeroSyncService } = createZeroSyncHarness(engine)
    const startWorkflow = vi.fn(async () => ({
      runId: 'wf-rollup-1',
      threadId: 'thr-1',
      startedByUserId: 'alex@example.com',
      workflowTemplate: 'createCurrentSheetRollup' as const,
      title: 'Create Current Sheet Rollup',
      summary: 'Staged a rollup preview for Revenue into Revenue Rollup.',
      status: 'completed' as const,
      createdAtUnixMs: 1,
      updatedAtUnixMs: 2,
      completedAtUnixMs: 2,
      errorMessage: null,
      steps: [
        {
          stepId: 'inspect-source-sheet',
          label: 'Inspect source sheet',
          status: 'completed' as const,
          summary: 'Loaded the used range and numeric columns from Revenue.',
          updatedAtUnixMs: 1,
        },
      ],
      artifact: {
        kind: 'markdown' as const,
        title: 'Current Sheet Rollup Preview',
        text: '## Current Sheet Rollup Preview',
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
        callId: 'call-workflow-rollup-1',
        tool: 'bilig_start_workflow',
        arguments: {
          workflowTemplate: 'createCurrentSheetRollup',
          sheetName: 'Revenue',
        },
      },
    )

    expect(response.success).toBe(true)
    expect(startWorkflow).toHaveBeenCalledWith({
      workflowTemplate: 'createCurrentSheetRollup',
      sheetName: 'Revenue',
    })
    const output = response.contentItems.find((item) => item.type === 'inputText')
    expect(output?.type).toBe('inputText')
    expect(output && 'text' in output ? output.text : '').toContain('"workflowTemplate": "createCurrentSheetRollup"')
  })

  it('stages column metadata commands for resize operations', async () => {
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
        callId: 'call-5',
        tool: 'bilig_update_column_metadata',
        arguments: {
          sheetName: 'Sheet1',
          startCol: 0,
          count: 2,
          width: 120,
        },
      },
    )

    expect(response.success).toBe(true)
    expect(stageCommand).toHaveBeenCalledWith({
      kind: 'updateColumnMetadata',
      sheetName: 'Sheet1',
      startCol: 0,
      count: 2,
      width: 120,
    })
  })
})
