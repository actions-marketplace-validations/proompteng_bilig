import type { EngineEvent, EngineOpBatch, EngineSyncClient } from './engine-test-helpers.js'
import {
  ErrorCode,
  Opcode,
  SpreadsheetEngine,
  ValueTag,
  afterEach,
  describe,
  expect,
  isRuntimeFormulaWithCompiled,
  isRuntimeFormulaWithDirectAggregate,
  isRuntimeFormulaWithRanges,
  it,
  readRuntimeFormula,
  vi,
} from './engine-test-helpers.js'

describe('SpreadsheetEngine runtime API and sync surfaces', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.restoreAllMocks()
  })

  it('stores runtime formula slot and compiled plan metadata separately', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'C3', 'A1*2')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'C3')
    expect(cellIndex).toBeDefined()

    const formulaId = engine.workbook.cellStore.formulaIds[cellIndex!]
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(formulaId).toBeGreaterThan(0)
    expect(isRuntimeFormulaWithCompiled(runtimeFormula)).toBe(true)
    expect(runtimeFormula).toBeDefined()
    expect(runtimeFormula?.formulaSlotId).toBe(formulaId)
    expect(runtimeFormula?.planId).toBe(runtimeFormula?.plan.id)
    expect(runtimeFormula?.compiled).toBe(runtimeFormula?.plan.compiled)
    expect(runtimeFormula?.dependencyEntities.ptr).toBeGreaterThanOrEqual(0)
    expect(runtimeFormula?.compiled.program.length).toBeGreaterThan(0)
    expect(runtimeFormula?.runtimeProgram.length).toBe(0)
    expect(runtimeFormula?.directScalar).toBeDefined()
  })

  it('reuses one compiled plan for identical formula sources while keeping distinct slots', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'shared-plan-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', '1+2')
    engine.setCellFormula('Sheet1', 'B1', '1+2')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()

    const leftFormula = readRuntimeFormula(engine, a1Index!)
    const rightFormula = readRuntimeFormula(engine, b1Index!)

    expect(isRuntimeFormulaWithCompiled(leftFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(rightFormula)).toBe(true)
    expect(leftFormula?.formulaSlotId).not.toBe(rightFormula?.formulaSlotId)
    expect(leftFormula?.planId).toBe(rightFormula?.planId)
    expect(leftFormula?.compiled).toBe(rightFormula?.compiled)
  })

  it('replaces direct lookup range dependencies with lookup-column subscribers and formula-cell deps only', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'lookup-subscriber-spec',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'E1')
    expect(cellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.rangeDependencies).toHaveLength(0)
    expect(runtimeFormula?.dependencyIndices).toEqual(Uint32Array.of(engine.workbook.getCellIndex('Sheet1', 'D1')!))
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getDependencies('Sheet1', 'D1').directDependents).toContain('Sheet1!E1')
    expect(engine.getDependencies('Sheet1', 'A2').directDependents).toEqual([])

    engine.setCellValue('Sheet1', 'A2', 25)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
  })

  it('patches runtime cell and range operands from packed symbolic binding buffers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)+C1')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    const c1Index = engine.workbook.getCellIndex('Sheet1', 'C1')
    expect(cellIndex).toBeDefined()
    expect(c1Index).toBeDefined()

    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(runtimeFormula).toBeDefined()
    const pushCellOpcode = Number(Opcode.PushCell)
    const pushRangeOpcode = Number(Opcode.PushRange)
    const pushCell = runtimeFormula?.runtimeProgram.find((instruction) => instruction >>> 24 === pushCellOpcode)
    const pushRange = runtimeFormula?.runtimeProgram.find((instruction) => instruction >>> 24 === pushRangeOpcode)

    expect(pushCell).toBeDefined()
    expect(pushRange).toBeDefined()
    expect(runtimeFormula?.dependencyIndices).toEqual(Uint32Array.of(c1Index!))
    expect(pushCell! & 0x00ff_ffff).toBe(c1Index)
    expect(pushRange! & 0x00ff_ffff).toBe(runtimeFormula?.rangeDependencies[0])
  })

  it('keeps packed range entity ids stable across structural inserts when the range survives', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-structural-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellValue('Sheet1', 'B1', 3)
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)+C1')

    const beforeCellIndex = engine.workbook.getCellIndex('Sheet1', 'D1')
    expect(beforeCellIndex).toBeDefined()
    const beforeRuntimeFormula = readRuntimeFormula(engine, beforeCellIndex!)
    expect(isRuntimeFormulaWithRanges(beforeRuntimeFormula)).toBe(true)

    const beforeRangeIndex = beforeRuntimeFormula?.rangeDependencies[0]
    expect(beforeRangeIndex).toBeDefined()

    engine.insertRows('Sheet1', 0, 1)

    const afterCellIndex = engine.workbook.getCellIndex('Sheet1', 'D2')
    expect(afterCellIndex).toBeDefined()
    const afterRuntimeFormula = readRuntimeFormula(engine, afterCellIndex!)
    expect(isRuntimeFormulaWithRanges(afterRuntimeFormula)).toBe(true)
    expect(afterRuntimeFormula?.rangeDependencies[0]).toBe(beforeRangeIndex)
    expect(engine.getCellValue('Sheet1', 'D2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('keeps packed range entity ids and plan ids stable across structural row deletes when the range survives', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'symbolic-structural-delete-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row * 10)
      engine.setCellValue('Sheet1', `C${row}`, row * 100)
    }
    engine.setCellFormula('Sheet1', 'D5', 'SUM(A1:B4)+C4')

    const beforeCellIndex = engine.workbook.getCellIndex('Sheet1', 'D5')
    expect(beforeCellIndex).toBeDefined()
    const beforeRuntimeFormula = readRuntimeFormula(engine, beforeCellIndex!)
    expect(isRuntimeFormulaWithRanges(beforeRuntimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(beforeRuntimeFormula)).toBe(true)

    const beforePlanId = beforeRuntimeFormula?.planId
    const beforeRangeIndex = beforeRuntimeFormula?.rangeDependencies[0]
    expect(beforePlanId).toBeDefined()
    expect(beforeRangeIndex).toBeDefined()

    engine.deleteRows('Sheet1', 1, 1)

    const afterCellIndex = engine.workbook.getCellIndex('Sheet1', 'D4')
    expect(afterCellIndex).toBeDefined()
    const afterRuntimeFormula = readRuntimeFormula(engine, afterCellIndex!)
    expect(isRuntimeFormulaWithRanges(afterRuntimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithCompiled(afterRuntimeFormula)).toBe(true)
    expect(afterRuntimeFormula?.planId).toBe(beforePlanId)
    expect(afterRuntimeFormula?.rangeDependencies[0]).toBe(beforeRangeIndex)
    expect(engine.getCell('Sheet1', 'D4').formula).toBe('SUM(A1:B3)+C3')
    expect(engine.getCellValue('Sheet1', 'D4')).toEqual({ tag: ValueTag.Number, value: 488 })
  })

  it('tracks literal-backed ranges through range entities without inflating topo dependency cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'range-topology-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'B1')
    expect(cellIndex).toBeDefined()
    const runtimeFormula = readRuntimeFormula(engine, cellIndex!)

    expect(isRuntimeFormulaWithRanges(runtimeFormula)).toBe(true)
    expect(isRuntimeFormulaWithDirectAggregate(runtimeFormula)).toBe(true)
    expect(runtimeFormula?.dependencyIndices).toEqual(new Uint32Array())
    expect(runtimeFormula?.rangeDependencies).toHaveLength(0)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.setCellValue('Sheet1', 'A2', 4)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
  })

  it('assigns deterministic cycle group ids for cyclic formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'cycle-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'B1+1')
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')

    expect(a1Index).toBeDefined()
    expect(b1Index).toBeDefined()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Cycle,
    })
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBeGreaterThanOrEqual(0)
    expect(engine.workbook.cellStore.cycleGroupIds[a1Index!]).toBe(engine.workbook.cellStore.cycleGroupIds[b1Index!])
  })

  it('assigns topo ranks through range-node dependents deterministically', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'topo-spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'D1', 'SUM(A1:B1)')

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')

    expect(b1Index).toBeDefined()
    expect(d1Index).toBeDefined()
    expect(engine.workbook.cellStore.topoRanks[b1Index!]).toBeLessThan(engine.workbook.cellStore.topoRanks[d1Index!])
  })

  it('notifies per-cell listeners only for the cells that changed', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let a1Notifications = 0
    let b1Notifications = 0
    const unsubscribeA1 = engine.subscribeCell('Sheet1', 'A1', () => {
      a1Notifications += 1
    })
    const unsubscribeB1 = engine.subscribeCell('Sheet1', 'B1', () => {
      b1Notifications += 1
    })

    engine.setCellValue('Sheet1', 'A1', 1)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(0)

    engine.setCellValue('Sheet1', 'B1', 2)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(1)

    engine.setCellValue('Sheet1', 'C1', 3)
    expect(a1Notifications).toBe(1)
    expect(b1Notifications).toBe(1)

    unsubscribeA1()
    unsubscribeB1()
  })

  it('notifies grouped watched cells only when one of them changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    let notifications = 0
    const unsubscribe = engine.subscribeCells('Sheet1', ['A1', 'A2'], () => {
      notifications += 1
    })

    engine.setCellValue('Sheet1', 'B1', 5)
    expect(notifications).toBe(0)

    engine.setCellValue('Sheet1', 'A2', 8)
    expect(notifications).toBe(1)

    unsubscribe()
  })

  it('notifies watched cells when sheet deletion clears them', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)

    let notifications = 0
    const unsubscribe = engine.subscribeCell('Sheet1', 'A1', () => {
      notifications += 1
    })

    engine.deleteSheet('Sheet1')

    expect(notifications).toBe(1)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    unsubscribe()
  })

  it('tracks selection state inside the engine and notifies subscribers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const seen: string[] = []
    const unsubscribe = engine.subscribeSelection(() => {
      const snapshot = engine.getSelectionState()
      seen.push(`${snapshot.sheetName}!${snapshot.address ?? 'null'}`)
    })

    engine.setSelection('Sheet2', 'B3')
    engine.setSelection('Sheet2', 'B3')
    engine.setSelection('Sheet1', 'A1')

    expect(engine.getSelectionState()).toEqual({
      sheetName: 'Sheet1',
      address: 'A1',
      anchorAddress: 'A1',
      range: { startAddress: 'A1', endAddress: 'A1' },
      editMode: 'idle',
    })
    expect(seen).toEqual(['Sheet2!B3', 'Sheet1!A1'])

    unsubscribe()
  })

  it('restores snapshots through transactions without emitting batches or undo history', async () => {
    const source = new SpreadsheetEngine({ workbookName: 'source' })
    await source.ready()
    source.createSheet('Sheet1')
    source.setCellValue('Sheet1', 'A1', 100)
    source.setDefinedName('TaxRate', 0.1)
    source.setCellFormula('Sheet1', 'A2', 'TaxRate*A1')

    const restored = new SpreadsheetEngine({ workbookName: 'restored' })
    await restored.ready()
    const outbound: EngineOpBatch[] = []
    restored.subscribeBatches((batch) => outbound.push(batch))

    restored.importSnapshot(source.exportSnapshot())

    expect(restored.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(restored.getDefinedNames()).toEqual([{ name: 'TaxRate', value: 0.1 }])
    expect(outbound).toEqual([])
    expect(restored.undo()).toBe(false)
  })

  it('supports range mutation helpers and undo/redo over the same local apply path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [3, 4],
    ])
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 4 })

    engine.setRangeFormulas({ sheetName: 'Sheet1', startAddress: 'C1', endAddress: 'C2' }, [['SUM(A1:B1)'], ['SUM(A2:B2)']])
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 7 })

    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'A2', endAddress: 'B2' })
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })

    engine.undo()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 7 })

    engine.redo()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 0 })
  })

  it('clears sparse whole-column ranges without scanning every row address', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sparse-clear-column' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'B1', 'clear-me')
    engine.setCellValue('Sheet1', 'C1', 'keep-me')

    const originalGetCellIndex = engine.workbook.getCellIndex.bind(engine.workbook)
    let lookupCount = 0
    const getCellIndexSpy = vi.spyOn(engine.workbook, 'getCellIndex').mockImplementation((sheetName, address) => {
      lookupCount += 1
      if (lookupCount > 10) {
        throw new Error('clearRange scanned row addresses instead of resident cells')
      }
      return originalGetCellIndex(sheetName, address)
    })

    engine.clearRange({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B1048576' })
    getCellIndexSpy.mockRestore()

    expect(lookupCount).toBeLessThanOrEqual(10)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual(expect.objectContaining({ tag: ValueTag.String, value: 'keep-me' }))
  })

  it('captures undo ops for a local mutation and reapplies raw engine ops deterministically', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const { undoOps } = engine.captureUndoOps(() => {
      engine.setCellValue('Sheet1', 'A1', 'seed')
    })

    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'seed',
    })
    expect(undoOps).not.toBeNull()

    const redoOps = engine.applyOps(undoOps ?? [], { captureUndo: true })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(redoOps).not.toBeNull()

    engine.applyOps(redoOps ?? [], { captureUndo: true })
    expect(engine.getCellValue('Sheet1', 'A1')).toMatchObject({
      tag: ValueTag.String,
      value: 'seed',
    })
  })

  it('emits cell invalidation for local applyOps batches without captured undo', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const events: EngineEvent[] = []
    const unsubscribe = engine.subscribe((event) => {
      events.push(event)
    })

    engine.applyOps([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 7 }], {
      source: 'local',
    })

    unsubscribe()

    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({
      kind: 'batch',
      invalidation: 'cells',
    })
    expect(events[0]?.changedCellIndices.length).toBeGreaterThan(0)
    expect(events[0]?.changedCells.map((change) => `${change.sheetName}!${change.a1}`)).toEqual(['Sheet1!A1', 'Sheet1!B1'])
  })

  it('applies coordinate-native cell mutations with formula recomputation and undo compatibility', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.setCellValueAt(sheetId, 0, 0, 2)
    engine.setCellFormulaAt(sheetId, 0, 1, 'A1*3')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    const undoOps = engine.applyCellMutationsAt(
      [
        { sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 } },
        { sheetId, mutation: { kind: 'clearCell', row: 1, col: 0 } },
      ],
      1,
    )

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(undoOps).not.toBeNull()

    engine.applyOps(undoOps ?? [], { captureUndo: true })

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
  })

  it('applies coordinate-native restore mutations without recording undo history', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const undoOps = engine.applyCellMutationsAtWithOptions(
      [
        { sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 5 } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
      ],
      {
        captureUndo: false,
        potentialNewCells: 2,
        source: 'restore',
      },
    )

    expect(undoOps).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
      }),
    )
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()

    unsubscribe()
  })

  it('applies coordinate-native restore mutations without listeners', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-refs-no-listeners' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const undoOps = engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 9 } }], {
      captureUndo: false,
      potentialNewCells: 1,
      source: 'restore',
    })

    expect(undoOps).toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()
  })

  it('emits standard engine events for restore mutations without tracked listeners', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'restore-cell-mutation-general-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const listener = vi.fn()
    const unsubscribe = engine.subscribe(listener)

    const undoOps = engine.applyCellMutationsAtWithOptions([{ sheetId, mutation: { kind: 'setCellValue', row: 0, col: 0, value: 11 } }], {
      captureUndo: false,
      potentialNewCells: 1,
      source: 'restore',
    })

    expect(undoOps).toBeNull()
    expect(listener).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        invalidation: 'full',
      }),
    )
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.undo()).toBe(true)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()

    unsubscribe()
  })

  it('emits standard engine events for coordinate-native clear cell mutations', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'clear-cell-at-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)
    const listener = vi.fn()
    const unsubscribe = engine.subscribe(listener)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    engine.clearCellAt(sheetId, 0, 0)

    expect(listener).toHaveBeenCalled()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })

    unsubscribe()
  })

  it('reads rectangular range values as a dense matrix without per-cell callers', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'C3' }, [
      [11, 12],
      [13, 14],
    ])
    engine.setCellFormula('Sheet1', 'D2', 'SUM(B2:C2)')

    expect(
      engine.getRangeValues({
        sheetName: 'Sheet1',
        startAddress: 'A1',
        endAddress: 'D3',
      }),
    ).toEqual([
      [{ tag: ValueTag.Empty }, { tag: ValueTag.Empty }, { tag: ValueTag.Empty }, { tag: ValueTag.Empty }],
      [
        { tag: ValueTag.Empty },
        { tag: ValueTag.Number, value: 11 },
        { tag: ValueTag.Number, value: 12 },
        { tag: ValueTag.Number, value: 23 },
      ],
      [{ tag: ValueTag.Empty }, { tag: ValueTag.Number, value: 13 }, { tag: ValueTag.Number, value: 14 }, { tag: ValueTag.Empty }],
    ])
  })

  it('copies and fills rectangular ranges', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }, [
      [1, 2],
      [3, 4],
    ])

    engine.copyRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' },
      { sheetName: 'Sheet1', startAddress: 'D1', endAddress: 'E2' },
    )
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: 4 })

    engine.fillRange(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
      { sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'D5' },
    )
    expect(engine.getCellValue('Sheet1', 'A4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'D5')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('tracks sync client connection state and forwards local batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })
    await engine.ready()

    const forwarded: EngineOpBatch[] = []
    let connected = false
    let disconnected = false

    await engine.connectSyncClient({
      async connect(this: void, handlers: Parameters<EngineSyncClient['connect']>[0]) {
        connected = true
        handlers.setState('behind')
        return {
          send(batch) {
            forwarded.push(batch)
          },
          async disconnect() {
            disconnected = true
          },
        }
      },
    })

    expect(connected).toBe(true)
    expect(engine.getSyncState()).toBe('behind')

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 9)

    expect(forwarded).toHaveLength(2)
    expect(forwarded[1]?.ops).toEqual([{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 9 }])

    await engine.disconnectSyncClient()
    expect(disconnected).toBe(true)
    expect(engine.getSyncState()).toBe('local-only')
  })

  it('disables sync when replica version tracking is off', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'local-only',
      trackReplicaVersions: false,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 5)

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })

    await expect(
      engine.connectSyncClient({
        connect() {
          throw new Error('should not connect')
        },
      }),
    ).rejects.toThrow('Sync is unavailable when trackReplicaVersions is disabled; construct the engine with trackReplicaVersions enabled.')
  })
})
