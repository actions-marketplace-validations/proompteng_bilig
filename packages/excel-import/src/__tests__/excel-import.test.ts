import type { WorkbookSnapshot } from './excel-import-test-helpers.js'
import {
  CSV_CONTENT_TYPE,
  EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES,
  ErrorCode,
  InvalidXlsxZipContainerError,
  LEGACY_XLS_CONTENT_TYPE,
  SpreadsheetEngine,
  ValueTag,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  buildBinaryWorkbook,
  buildCorruptZipBackedWorkbook,
  buildExternalDefinedNamesWorkbook,
  buildExternalGetPivotDataLinkCacheWorkbook,
  buildExternalLinkCacheWorkbook,
  buildExternalLinkRangeCacheWorkbook,
  buildGenericWorkflowWorkbookFixture,
  buildLegacyWorkbook,
  buildMacroEnabledWorkbook,
  buildNamespacedFormulaWorkbook,
  buildRatesWorkbook,
  buildScopedDefinedNamesWorkbook,
  buildSingleCellMergeWorkbook,
  buildSparseExternalLinkRangeCacheWorkbook,
  buildSparseRatesWorkbook,
  buildUnsupportedFunctionCacheWorkbook,
  buildVolatileFormulaWorkbook,
  buildWholeColumnDefinedNamesWorkbook,
  buildWorkbook,
  buildZeroSizeMetadataWorkbook,
  byteSourceFor,
  cellsFromRows,
  describe,
  expect,
  exportXlsx,
  externalWorkbookCompanionAmbiguousMatchWarning,
  externalWorkbookCompanionNoMatchWarning,
  externalWorkbookReferencesWarning,
  importCsv,
  importWorkbookFile,
  importXlsx,
  importXlsxFromZipByteSource,
  inflateXlsxForDenseSheetJsParse,
  it,
  readExternalLinkCacheXml,
  readImportedXlsxCellStyle,
  readRuntimeImage,
  strFromU8,
  unzipSync,
  volatileFormulasWarning,
  writeSimpleXlsxWorkbook,
} from './excel-import-test-helpers.js'

describe('excel import workbook structure and external cache fidelity', () => {
  it('imports sheets, formulas, dimensions, and warnings from xlsx bytes', () => {
    const imported = importXlsx(buildWorkbook(), 'Quarterly Report.xlsx')

    expect(imported.workbookName).toBe('Quarterly Report')
    expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
    expect(imported.snapshot.workbook.name).toBe('Quarterly Report')
    expect(imported.snapshot.sheets).toHaveLength(2)

    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'Sheet1',
      metadata: {
        columns: [
          { index: 0, size: 120 },
          { index: 1, size: 65 },
          { index: 2, size: 80 },
        ],
        rows: [
          { index: 0, size: 30 },
          { index: 1, size: 18 },
        ],
        merges: [{ sheetName: 'Sheet1', startAddress: 'A4', endAddress: 'B4' }],
      },
    })
    expect(imported.snapshot.sheets[0]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 1 })]))
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'C1', formula: 'A1+B1', format: '0.00' })]),
    )
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A1', value: 'hello' })]))
    expect(imported.snapshot.sheets[1]?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'A2', value: true })]))

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'InputBlock', value: { kind: 'range-ref', sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' } },
      { name: 'InputValue', value: { kind: 'cell-ref', sheetName: 'Sheet1', address: 'A1' } },
    ])
    expect(imported.snapshot.sheets[1]?.metadata?.commentThreads).toEqual([
      {
        threadId: 'xlsx-comment:Sheet2:A1',
        sheetName: 'Sheet2',
        address: 'A1',
        comments: [{ id: 'xlsx-comment:Sheet2:A1:1', body: 'comment', authorDisplayName: 'Greg' }],
      },
    ])
    expect(imported.warnings).toEqual([])
    expect(imported.preview.workbookName).toBe('Quarterly Report')
    expect(imported.preview.sheetCount).toBe(2)
    expect(imported.preview.sheets).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          name: 'Sheet1',
          rowCount: 2,
          columnCount: 3,
          nonEmptyCellCount: 4,
          previewRows: [
            ['1', '2', '=A1+B1'],
            ['3', '', ''],
          ],
        }),
      ]),
    )
  })

  it('warns for every formula-compiler volatile function during XLSX import', () => {
    for (const formula of ['INDIRECT("A1")', 'OFFSET(A1,0,0)', '_xlfn.RANDARRAY(2,2)', 'SUBTOTAL(9,A1:A1)', 'AGGREGATE(9,0,A1:A1)']) {
      const imported = importXlsx(buildVolatileFormulaWorkbook(formula), 'volatile-formula.xlsx')

      expect(imported.warnings, formula).toContain(volatileFormulasWarning)
    }
  })

  it('resolves external workbook cell references from saved XLSX external-link caches', async () => {
    const imported = importXlsx(buildExternalLinkCacheWorkbook(), 'external-link-cache.xlsx')
    const formulaCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')

    expect(formulaCell?.formula).toBe('2+3')
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])
    expect(imported.snapshot.workbook.metadata?.externalWorkbookReferences).toEqual([
      {
        bookIndex: 1,
        packagePath: 'xl/externalLinks/externalLink5.xml',
        target: 'file:///tmp/source.xlsx',
        targetMode: 'External',
        workbookName: 'source.xlsx',
        sheetNames: ['External Data'],
      },
    ])
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toEqual([
      {
        kind: 'external-workbook-reference',
        sheetName: 'Report',
        address: 'A1',
        formula: "'[1]External Data'!A1+'[1]External Data'!A2",
        importedFormula: '2+3',
        linkedWorkbooks: [
          {
            bookIndex: 1,
            packagePath: 'xl/externalLinks/externalLink5.xml',
            target: 'file:///tmp/source.xlsx',
            targetMode: 'External',
            workbookName: 'source.xlsx',
            sheetNames: ['External Data'],
          },
        ],
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 2,
        unresolvedExternalReferenceCount: 0,
        reason:
          'Formula depends on an external workbook reference; cached linked values are preserved but linked workbooks are not recalculated during import.',
      },
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Report', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('materializes external workbook range references from saved XLSX external-link caches', async () => {
    const imported = importXlsx(buildExternalLinkRangeCacheWorkbook(), 'external-link-range-cache.xlsx')
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])

    expect(cacheSheet).toMatchObject({
      name: '__bilig_ext_1_Rates',
      order: 1,
      metadata: { visibility: 'veryHidden' },
    })
    expect(cacheCells).toEqual(
      new Map([
        ['A1', 'SKU'],
        ['B1', 'Rate'],
        ['A2', 'A'],
        ['B2', 10],
        ['A3', 'B'],
        ['B3', 20],
        ['A4', 'C'],
        ['B4', 30],
      ]),
    )
    expect(cells.get('C1')).toMatchObject({ formula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1", value: 120 })
    expect(cells.get('C2')).toMatchObject({
      formula: "XLOOKUP(\"B\",'__bilig_ext_1_Rates'!$A$2:$A$4,'__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
      value: 40,
    })
    expect(cells.get('C3')).toMatchObject({
      formula: "SUMPRODUCT('__bilig_ext_1_Rates'!$B$2:$B$4,--('__bilig_ext_1_Rates'!$A$2:$A$4=\"C\"))*B1",
      value: 60,
    })
    expect(imported.warnings).toEqual([externalWorkbookReferencesWarning])
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toEqual([
      expect.objectContaining({
        address: 'C1',
        importedFormula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 3,
        unresolvedExternalReferenceCount: 0,
      }),
      expect.objectContaining({
        address: 'C2',
        importedFormula: "XLOOKUP(\"B\",'__bilig_ext_1_Rates'!$A$2:$A$4,'__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 6,
        unresolvedExternalReferenceCount: 0,
      }),
      expect.objectContaining({
        address: 'C3',
        importedFormula: "SUMPRODUCT('__bilig_ext_1_Rates'!$B$2:$B$4,--('__bilig_ext_1_Rates'!$A$2:$A$4=\"C\"))*B1",
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 6,
        unresolvedExternalReferenceCount: 0,
      }),
    ])

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Model', 'C1')).toEqual({ tag: ValueTag.Number, value: 120 })
    expect(engine.getCellValue('Model', 'C2')).toEqual({ tag: ValueTag.Number, value: 40 })
    expect(engine.getCellValue('Model', 'C3')).toEqual({ tag: ValueTag.Number, value: 60 })

    engine.setCellValue('Model', 'B1', 3)

    expect(engine.getCellValue('Model', 'C1')).toEqual({ tag: ValueTag.Number, value: 180 })
    expect(engine.getCellValue('Model', 'C2')).toEqual({ tag: ValueTag.Number, value: 60 })
    expect(engine.getCellValue('Model', 'C3')).toEqual({ tag: ValueTag.Number, value: 90 })
  })

  it('hydrates saved external-link caches from supplied workbook bytes by package path', () => {
    const imported = importXlsx(buildExternalLinkRangeCacheWorkbook(), 'external-link-range-cache.xlsx', {
      externalWorkbooks: [{ fileName: 'rates.xlsx', bytes: buildRatesWorkbook([20, 30, 40]) }],
    })
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])

    expect(cacheCells.has('A1')).toBe(false)
    expect(cacheCells.has('B1')).toBe(false)
    expect(cacheCells.get('B2')).toBe(20)
    expect(cacheCells.get('B3')).toBe(30)
    expect(cacheCells.get('B4')).toBe(40)
    expect(imported.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 1,
      externalReferenceCount: 1,
      refreshedBookIndices: [1],
      refreshedSheetCount: 1,
      refreshedCellCount: 6,
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 0,
      skippedEmptyRefreshCount: 0,
      references: [
        expect.objectContaining({
          bookIndex: 1,
          status: 'refreshed',
          candidateCount: 1,
          referenceCandidateCount: 1,
          matchKind: 'unique-workbook-identity',
          matchedFileName: 'rates.xlsx',
          refreshedSheetCount: 1,
          refreshedCellCount: 6,
        }),
      ],
    })

    const externalLinkXml = readExternalLinkCacheXml(exportXlsx(imported.snapshot))
    expect(externalLinkXml).toContain('<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>')
    expect(externalLinkXml).toContain('<row r="2">')
    expect(externalLinkXml).toContain('<cell r="B2"><v>20</v></cell>')
    expect(externalLinkXml).toContain('<row r="3">')
    expect(externalLinkXml).toContain('<cell r="B3"><v>30</v></cell>')
    expect(externalLinkXml).toContain('<row r="4">')
    expect(externalLinkXml).toContain('<cell r="B4"><v>40</v></cell>')
    expect(externalLinkXml).not.toContain('<row r="0">')
  })

  it('keeps companion workbook hydration scoped on the dense SheetJS parse path', () => {
    const sourceBytes = inflateXlsxForDenseSheetJsParse(buildExternalLinkRangeCacheWorkbook())
    const imported = importXlsxFromZipByteSource(byteSourceFor(sourceBytes), 'external-link-range-cache-dense.xlsx', {
      externalWorkbooks: [{ fileName: 'rates.xlsx', bytes: buildRatesWorkbook([20, 30, 40]) }],
      limits: { maxMaterializedSourceBytes: sourceBytes.byteLength },
    })
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
    const externalLinkXml = readExternalLinkCacheXml(exportXlsx(imported.snapshot))

    expect(sourceBytes.byteLength).toBeGreaterThan(1_000_000)
    expect(cacheCells).toEqual(
      new Map([
        ['A2', 'A'],
        ['B2', 20],
        ['A3', 'B'],
        ['B3', 30],
        ['A4', 'C'],
        ['B4', 40],
      ]),
    )
    expect(externalLinkXml).toContain('<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>')
    expect(externalLinkXml).toContain('<row r="2">')
    expect(externalLinkXml).toContain('<cell r="B2"><v>20</v></cell>')
    expect(externalLinkXml).toContain('<row r="4">')
    expect(externalLinkXml).toContain('<cell r="B4"><v>40</v></cell>')
  })

  it('hydrates sparse blank and error external ranges from supplied workbook bytes', async () => {
    const imported = importXlsx(buildSparseExternalLinkRangeCacheWorkbook(), 'external-link-sparse-cache.xlsx', {
      externalWorkbooks: [{ fileName: 'rates.xlsx', bytes: buildSparseRatesWorkbook() }],
    })
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cacheCells.get('B2')).toMatchObject({ value: 20 })
    expect(cacheCells.get('B3')).toMatchObject({ value: null })
    expect(cacheCells.get('B4')).toMatchObject({ value: 50 })
    expect(cacheCells.get('B5')).toMatchObject({ formula: '#N/A' })
    expect(cacheCells.get('B6')).toMatchObject({ formula: '#NULL!' })
    expect(imported.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 1,
      externalReferenceCount: 1,
      refreshedBookIndices: [1],
      refreshedSheetCount: 1,
      refreshedCellCount: 5,
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 0,
      skippedEmptyRefreshCount: 0,
    })
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toContainEqual(
      expect.objectContaining({
        address: 'C1',
        importedFormula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 3,
        unresolvedExternalReferenceCount: 0,
      }),
    )

    const formulaCells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])
    expect(formulaCells.get('C1')).toMatchObject({
      formula: "SUM('__bilig_ext_1_Rates'!$B$2:$B$4)*B1",
      value: 60,
    })
    expect(formulaCells.get('C2')).toMatchObject({
      formula: "IFERROR(SUM('__bilig_ext_1_Rates'!$B$2:$B$5),99)",
      value: 60,
    })
    expect(formulaCells.get('C3')).toMatchObject({
      formula: 'IFERROR(SUM(#NULL!),88)',
      value: 60,
    })

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-sparse-cache-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Model', 'C1')).toEqual({ tag: ValueTag.Number, value: 70 })
    expect(engine.getCellValue('Model', 'C2')).toEqual({ tag: ValueTag.Number, value: 99 })
    expect(engine.getCellValue('Model', 'C3')).toEqual({ tag: ValueTag.Number, value: 88 })
    expect(engine.getCellValue('__bilig_ext_1_Rates', 'B5')).toEqual({ tag: ValueTag.Error, code: ErrorCode.NA })
    expect(engine.getCellValue('__bilig_ext_1_Rates', 'B6')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Null })

    const externalLinkXml = readExternalLinkCacheXml(exportXlsx(imported.snapshot))
    expect(externalLinkXml).toContain('<cell r="B2"><v>20</v></cell>')
    expect(externalLinkXml).not.toContain('r="B3"')
    expect(externalLinkXml).toContain('<cell r="B4"><v>50</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B5" t="e"><v>#N/A</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B6" t="e"><v>#NULL!</v></cell>')
  })

  it('does not hydrate external-link caches from an explicitly mismatched target', () => {
    const imported = importXlsx(buildExternalLinkRangeCacheWorkbook(), 'external-link-range-cache.xlsx', {
      externalWorkbooks: [
        {
          fileName: 'rates.xlsx',
          target: 'file:///tmp/other/rates.xlsx',
          bytes: buildRatesWorkbook([20, 30, 40]),
        },
      ],
    })
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
    const externalLinkXml = readExternalLinkCacheXml(exportXlsx(imported.snapshot))

    expect(cacheCells.get('B2')).toBe(10)
    expect(cacheCells.get('B3')).toBe(20)
    expect(cacheCells.get('B4')).toBe(30)
    expect(imported.warnings).toContain(externalWorkbookCompanionNoMatchWarning)
    expect(imported.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 1,
      externalReferenceCount: 1,
      refreshedBookIndices: [],
      skippedNoMatchCount: 1,
      skippedAmbiguousMatchCount: 0,
      references: [
        expect.objectContaining({
          bookIndex: 1,
          status: 'skipped-no-match',
          candidateCount: 0,
        }),
      ],
    })
    expect(externalLinkXml).toContain('<cell r="B2"><v>10</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B3"><v>20</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B4"><v>30</v></cell>')
  })

  it('fails closed when companion workbook basename matching is ambiguous', () => {
    const imported = importXlsx(buildExternalLinkRangeCacheWorkbook(), 'external-link-range-cache.xlsx', {
      externalWorkbooks: [
        { fileName: 'rates.xlsx', bytes: buildRatesWorkbook([20, 30, 40]) },
        { fileName: 'rates.xlsx', bytes: buildRatesWorkbook([200, 300, 400]) },
      ],
    })
    const cacheSheet = imported.snapshot.sheets.find((sheet) => sheet.name === '__bilig_ext_1_Rates')
    const cacheCells = new Map(cacheSheet?.cells.map((cell) => [cell.address, cell.value]) ?? [])
    const externalLinkXml = readExternalLinkCacheXml(exportXlsx(imported.snapshot))

    expect(cacheCells.get('B2')).toBe(10)
    expect(cacheCells.get('B3')).toBe(20)
    expect(cacheCells.get('B4')).toBe(30)
    expect(imported.warnings).toContain(externalWorkbookCompanionAmbiguousMatchWarning)
    expect(imported.warnings).not.toContain(externalWorkbookCompanionNoMatchWarning)
    expect(imported.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 2,
      externalReferenceCount: 1,
      refreshedBookIndices: [],
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 1,
      references: [
        expect.objectContaining({
          bookIndex: 1,
          status: 'skipped-ambiguous-match',
          candidateCount: 2,
          referenceCandidateCount: 1,
          matchKind: 'unique-workbook-identity',
        }),
      ],
    })
    expect(externalLinkXml).toContain('<cell r="B2"><v>10</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B3"><v>20</v></cell>')
    expect(externalLinkXml).toContain('<cell r="B4"><v>30</v></cell>')
  })

  it('materializes external criteria-function ranges as hidden-sheet references', async () => {
    const imported = importXlsx(
      buildExternalLinkRangeCacheWorkbook("SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,\"C\")*B1"),
      'external-link-range-cache-sumifs.xlsx',
    )
    const cells = new Map(imported.snapshot.sheets[0]?.cells.map((cell) => [cell.address, cell]) ?? [])

    expect(cells.get('C3')).toMatchObject({
      formula: "SUMIFS('__bilig_ext_1_Rates'!$B$2:$B$4,'__bilig_ext_1_Rates'!$A$2:$A$4,\"C\")*B1",
      value: 60,
    })
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toContainEqual(
      expect.objectContaining({
        address: 'C3',
        importedFormula: "SUMIFS('__bilig_ext_1_Rates'!$B$2:$B$4,'__bilig_ext_1_Rates'!$A$2:$A$4,\"C\")*B1",
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: false,
        cachedExternalReferenceValuesUsed: true,
        resolvedExternalReferenceCount: 6,
      }),
    )
    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const modelSheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const workbookXml = strFromU8(exportedZip['xl/workbook.xml'] ?? new Uint8Array())

    expect(modelSheetXml).toContain('__bilig_ext_1_Rates')
    expect(modelSheetXml).not.toContain('{')
    expect(workbookXml).toContain('name="__bilig_ext_1_Rates"')
    expect(workbookXml).toContain('state="veryHidden"')

    const engine = new SpreadsheetEngine({ workbookName: 'external-link-range-cache-sumifs-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)

    expect(engine.getCellValue('Model', 'C3')).toEqual({ tag: ValueTag.Number, value: 60 })

    engine.setCellValue('Model', 'B1', 3)

    expect(engine.getCellValue('Model', 'C3')).toEqual({ tag: ValueTag.Number, value: 90 })
  })

  it('preserves external GETPIVOTDATA anchors instead of replacing them with cached labels', () => {
    const imported = importXlsx(buildExternalGetPivotDataLinkCacheWorkbook(), 'external-pivot-link-cache.xlsx')
    const formulaCell = imported.snapshot.sheets[0]?.cells.find((cell) => cell.address === 'A1')

    expect(formulaCell?.formula).toBe('GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")')
    expect(formulaCell?.value).toBe(15)
    expect(imported.warnings).toContain(externalWorkbookReferencesWarning)
    expect(imported.snapshot.workbook.metadata?.externalWorkbookReferences).toEqual([
      {
        bookIndex: 1,
        packagePath: 'xl/externalLinks/externalLink5.xml',
        target: 'file:///tmp/pivot-source.xlsx',
        targetMode: 'External',
        workbookName: 'pivot-source.xlsx',
        sheetNames: ['External Pivot'],
      },
    ])
    expect(imported.snapshot.workbook.metadata?.unsupportedFormulaDependencies).toEqual([
      expect.objectContaining({
        kind: 'external-workbook-reference',
        sheetName: 'Report',
        address: 'A1',
        formula: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
        importedFormula: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
        linkedWorkbooks: [
          {
            bookIndex: 1,
            packagePath: 'xl/externalLinks/externalLink5.xml',
            target: 'file:///tmp/pivot-source.xlsx',
            targetMode: 'External',
            workbookName: 'pivot-source.xlsx',
            sheetNames: ['External Pivot'],
          },
        ],
        cachedValuesUsed: true,
        cachedFormulaValuePreserved: true,
        cachedExternalReferenceValuesUsed: false,
        resolvedExternalReferenceCount: 0,
        unresolvedExternalReferenceCount: 0,
      }),
    ])
  })

  it('retains cached values for imported formula cells that use unavailable add-in functions', () => {
    const imported = importXlsx(buildUnsupportedFunctionCacheWorkbook(), 'udf-cache.xlsx')
    const sheet = imported.snapshot.sheets[0]

    expect(sheet?.cells.find((cell) => cell.address === 'A1')).toMatchObject({
      formula: '_xldudf_WISEPRICE(B1,"Shares Outstanding")',
      value: 14935800000,
    })
    expect(sheet?.cells.find((cell) => cell.address === 'C1')).toMatchObject({
      formula: '_FV(B1,"Ticker symbol",TRUE)',
      value: 'AAPL',
    })
  })

  it('drops degenerate single-cell merge records during import', async () => {
    const imported = importXlsx(buildSingleCellMergeWorkbook(), 'single-cell-merge.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' }])
    const engine = new SpreadsheetEngine({ workbookName: 'single-cell-merge-import' })
    await engine.ready()
    expect(() => engine.importSnapshot(imported.snapshot)).not.toThrow()
  })

  it('ignores zero-size row and column metadata during import', () => {
    const imported = importXlsx(buildZeroSizeMetadataWorkbook(), 'zero-size.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.rows).toBeUndefined()
    expect(imported.snapshot.sheets[0]?.metadata?.columns).toBeUndefined()
  })

  it('preserves hidden row metadata even when the row has no custom height', () => {
    const imported = importXlsx(
      writeSimpleXlsxWorkbook({
        sheets: [
          {
            name: 'Table',
            cells: cellsFromRows([
              ['Header', 'Value'],
              ['Visible', 10],
              ['Hidden', 20],
            ]),
            rows: [{ index: 2, hidden: true }],
          },
        ],
      }),
      'hidden-row.xlsx',
    )

    expect(imported.snapshot.sheets[0]?.metadata?.rows).toEqual([{ id: 'row:2', index: 2, hidden: true }])
  })

  it('canonicalizes imported multiline text to LF line breaks', () => {
    const imported = importXlsx(
      writeSimpleXlsxWorkbook({
        sheets: [{ name: 'Sheet1', cells: cellsFromRows([['Line 1\r\nLine 2\rLine 3']]) }],
      }),
      'multiline.xlsx',
    )

    expect(imported.snapshot.sheets[0]?.cells).toContainEqual({
      address: 'A1',
      row: 0,
      col: 0,
      value: 'Line 1\nLine 2\nLine 3',
    })
  })

  it('preserves external workbook defined names as formulas across export round trips', () => {
    const imported = importXlsx(buildExternalDefinedNamesWorkbook(), 'external-defined-names.xlsx')
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'ExternalBrokenRef', value: { kind: 'formula', formula: '=[2]Sheet1!#REF!' } },
      { name: 'ExternalRange', value: { kind: 'formula', formula: '=[1]Sheet1!$A$1:$A$2' } },
    ])

    const roundTripped = importXlsx(exportXlsx(imported.snapshot), 'external-defined-names.xlsx')
    expect(roundTripped.snapshot.workbook.metadata?.definedNames).toEqual(imported.snapshot.workbook.metadata?.definedNames)
  })

  it('preserves sheet-scoped defined names across import and export round trips', () => {
    const imported = importXlsx(buildScopedDefinedNamesWorkbook(), 'scoped-defined-names.xlsx')

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'LocalBonus', value: { kind: 'cell-ref', sheetName: 'Global', address: 'A1' } },
      { name: 'LocalBonus', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'A1' } },
      { name: 'LocalRevenue', scopeSheetName: 'Local', value: { kind: 'cell-ref', sheetName: 'Local', address: 'B1' } },
    ])

    const roundTripped = importXlsx(exportXlsx(imported.snapshot), 'scoped-defined-names-roundtrip.xlsx')
    expect(roundTripped.warnings).toEqual([])
    expect(roundTripped.snapshot.workbook.metadata?.definedNames).toEqual(imported.snapshot.workbook.metadata?.definedNames)
  })

  it('preserves formula-only cells across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'formula-only-export' },
      sheets: [
        {
          id: 1,
          name: 'Summary',
          order: 0,
          cells: [
            { address: 'A1', value: 12 },
            { address: 'A2', value: 5 },
            { address: 'A3', formula: 'A1-A2' },
          ],
        },
      ],
    }

    const roundTripped = importXlsx(exportXlsx(snapshot), 'formula-only-export.xlsx')

    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([expect.objectContaining({ address: 'A3', formula: 'A1-A2' })]),
    )
  })

  it('preserves sheet names with trailing spaces across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'Trailing Space Workbook' },
      sheets: [
        {
          id: 'sheet-trailing-space',
          name: 'Table 2.1.2  ',
          order: 0,
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'Header' },
            { address: 'B1', row: 0, col: 1, value: 'Value' },
          ],
          metadata: {
            merges: [{ sheetName: 'Table 2.1.2  ', startAddress: 'A1', endAddress: 'B1' }],
          },
        },
      ],
    }

    const roundTripped = importXlsx(exportXlsx(snapshot), 'trailing-space-sheet.xlsx')
    expect(roundTripped.snapshot.sheets[0]?.name).toBe('Table 2.1.2  ')
    expect(roundTripped.snapshot.sheets[0]?.metadata?.merges).toEqual([
      { sheetName: 'Table 2.1.2  ', startAddress: 'A1', endAddress: 'B1' },
    ])
  })

  it('preserves leading-zero number formats across export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: { name: 'leading-zero-number-format' },
      sheets: [
        {
          id: 1,
          name: 'Codes',
          order: 0,
          cells: [{ address: 'A1', value: 7, format: '00' }],
        },
      ],
    }

    const bytes = exportXlsx(snapshot)
    const zip = unzipSync(bytes)
    const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
    const roundTripped = importXlsx(bytes, 'leading-zero-number-format.xlsx')

    expect(stylesXml).toContain('formatCode="00"')
    expect(stylesXml).not.toContain('numFmtId="00"')
    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', row: 0, col: 0, value: 7, format: '00' }])
  })

  it('preserves macro payloads without executing them across macro-enabled workbook import and export', () => {
    const imported = importXlsx(buildMacroEnabledWorkbook(), 'Macro Workbook.xlsm')

    expect(imported.workbookName).toBe('Macro Workbook')
    expect(imported.warnings).toContain('Macros were preserved but not executed during XLSX import.')
    expect(imported.snapshot.workbook.metadata?.macroPayloads).toEqual([
      {
        kind: 'vbaProject',
        storage: 'base64',
        dataBase64: 'AQIDBA==',
        byteLength: 4,
        preservedWithoutExecution: true,
        workbookCodeName: 'ThisWorkbook',
        sheetCodeNames: [{ sheetName: 'Sheet1', codeName: 'Sheet1' }],
      },
    ])
    expect(imported.snapshot.sheets[0]?.cells).toEqual([expect.objectContaining({ address: 'A1', value: 'safe value' })])

    const exported = exportXlsx(imported.snapshot)
    const exportedZip = unzipSync(exported)
    const contentTypesXml = strFromU8(exportedZip['[Content_Types].xml'] ?? new Uint8Array())
    const roundTripped = importXlsx(exported, 'Macro Workbook.xlsm')
    expect(contentTypesXml).toContain('<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>')
    expect(contentTypesXml).toContain(
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>',
    )
    expect(roundTripped.sheetNames).toEqual(['Sheet1'])
    expect(roundTripped.snapshot.workbook.metadata?.macroPayloads).toEqual(imported.snapshot.workbook.metadata?.macroPayloads)
    expect(roundTripped.snapshot.sheets[0]?.cells).toEqual([expect.objectContaining({ address: 'A1', value: 'safe value' })])
  })

  it('maps imported xlsx styles into Bilig style records', () => {
    expect(
      readImportedXlsxCellStyle({
        patternType: 'solid',
        fgColor: { rgb: '1D3989' },
        font: {
          name: 'Aptos',
          sz: 12,
          bold: true,
          italic: true,
          underline: true,
          color: { rgb: 'FFFFFFFF' },
        },
        alignment: {
          horizontal: 'center',
          vertical: 'center',
          wrapText: true,
          indent: 1,
        },
        border: {
          bottom: {
            style: 'thin',
            color: { rgb: 'FF000000' },
          },
        },
        protection: {
          locked: false,
          hidden: true,
        },
      }),
    ).toEqual({
      fill: { backgroundColor: '#1d3989' },
      font: {
        family: 'Aptos',
        size: 12,
        bold: true,
        italic: true,
        underline: true,
        color: '#ffffff',
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrap: true,
        indent: 1,
      },
      borders: {
        bottom: {
          style: 'solid',
          weight: 'thin',
          color: '#000000',
        },
      },
      protection: {
        locked: false,
        hidden: true,
      },
    })
  })

  it('coalesces repeated imported xlsx styles into rectangular ranges', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Styled import',
        metadata: {
          styles: [
            {
              id: 'header-fill',
              fill: { backgroundColor: '#1d3989' },
              font: { bold: true, color: '#ffffff' },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Styled',
          order: 0,
          cells: [
            { address: 'A1', value: 'Header' },
            { address: 'B1', value: 'Header' },
            { address: 'C1', value: 'Header' },
            { address: 'A2', value: 'Header' },
            { address: 'B2', value: 'Header' },
            { address: 'C2', value: 'Header' },
            { address: 'A3', value: 'Header' },
            { address: 'B3', value: 'Header' },
            { address: 'C3', value: 'Header' },
          ],
          metadata: {
            styleRanges: [
              {
                range: { sheetName: 'Styled', startAddress: 'A1', endAddress: 'C3' },
                styleId: 'header-fill',
              },
            ],
          },
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'styled-block.xlsx')
    const styleRanges = imported.snapshot.sheets[0]?.metadata?.styleRanges ?? []
    const styleRange = styleRanges[0]

    expect(styleRanges).toHaveLength(1)
    expect(styleRange?.range).toEqual({ sheetName: 'Styled', startAddress: 'A1', endAddress: 'C3' })
    const style = imported.snapshot.workbook.metadata?.styles?.find((entry) => entry.id === styleRange?.styleId)
    expect(style).toMatchObject({
      fill: { backgroundColor: '#1d3989' },
      font: { bold: true, color: '#ffffff' },
    })
  })

  it('preserves cell-level protection style metadata across XLSX export round trips', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Cell protection styles',
        metadata: {
          styles: [
            {
              id: 'unlocked-input',
              fill: { backgroundColor: '#fff2cc' },
              protection: { locked: false },
            },
            {
              id: 'hidden-formula',
              font: { color: '#000000' },
              protection: { locked: true, hidden: true },
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Protected',
          order: 0,
          metadata: {
            sheetProtection: { sheetName: 'Protected' },
            styleRanges: [
              { range: { sheetName: 'Protected', startAddress: 'B2', endAddress: 'B3' }, styleId: 'unlocked-input' },
              { range: { sheetName: 'Protected', startAddress: 'C2', endAddress: 'C3' }, styleId: 'hidden-formula' },
            ],
          },
          cells: [
            { address: 'B2', value: 10 },
            { address: 'B3', value: 25 },
            { address: 'C2', formula: 'B2*2' },
            { address: 'C3', formula: 'B3*2' },
          ],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const zip = unzipSync(exported)
    const stylesXml = strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())
    const sheetXml = strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
    const imported = importXlsx(exported, 'cell-protection-style-roundtrip.xlsx')
    const stylesById = new Map((imported.snapshot.workbook.metadata?.styles ?? []).map((style) => [style.id, style]))
    const styleForRange = (startAddress: string, endAddress: string) => {
      const styleRange = imported.snapshot.sheets[0]?.metadata?.styleRanges?.find(
        (entry) => entry.range.startAddress === startAddress && entry.range.endAddress === endAddress,
      )
      return styleRange ? stylesById.get(styleRange.styleId) : undefined
    }

    expect(stylesXml).toContain('applyProtection="1"')
    expect(stylesXml).toContain('<protection locked="0"/>')
    expect(stylesXml).toContain('<protection locked="1" hidden="1"/>')
    expect(sheetXml).toContain('<sheetProtection sheet="1"/>')
    expect(styleForRange('B2', 'B3')?.protection).toEqual({ locked: false })
    expect(styleForRange('C2', 'C3')?.protection).toEqual({ locked: true, hidden: true })
  })

  it('imports multiple generic workbook shapes without file-specific dispatch', () => {
    const operations = importXlsx(buildGenericWorkflowWorkbookFixture('multi-sheet-operations'), 'operations-workflow.xlsx')
    expect(operations.sheetNames).toEqual(['Dashboard', 'Ledger', 'Rollforward', 'Lookups'])
    const ledger = operations.snapshot.sheets.find((sheet) => sheet.name === 'Ledger')
    expect(ledger).toMatchObject({
      name: 'Ledger',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 132 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'H1' }],
      },
    })
    expect(ledger?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A4', value: 'OP001' }),
        expect.objectContaining({ address: 'B4', value: 45292 }),
        expect.objectContaining({
          address: 'G4',
          formula: 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)',
        }),
      ]),
    )
    const rollforward = operations.snapshot.sheets.find((sheet) => sheet.name === 'Rollforward')
    expect(rollforward?.cells).toEqual(expect.arrayContaining([expect.objectContaining({ address: 'E5', formula: 'IF(B5=B4,E4+D5,D5)' })]))

    const planning = importXlsx(buildGenericWorkflowWorkbookFixture('single-sheet-planning'), 'monthly-plan.xlsx')
    expect(planning.sheetNames).toEqual(['Monthly Plan'])
    expect(planning.snapshot.sheets[0]).toMatchObject({
      name: 'Monthly Plan',
      metadata: {
        columns: expect.arrayContaining([{ id: 'col:0', index: 0, size: 168 }]),
        rows: expect.arrayContaining([{ id: 'row:0', index: 0, size: 30 }]),
        merges: [{ sheetName: 'Monthly Plan', startAddress: 'A1', endAddress: 'I1' }],
      },
    })
    expect(planning.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A3', value: 'TenantWorks' }),
        expect.objectContaining({
          address: 'F3',
          formula: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
        }),
      ]),
    )
  })

  it('imports csv files into a single-sheet workbook preview', async () => {
    const imported = importCsv('Name,Value\nalpha,12\nbeta,=A2', 'metrics.csv')

    expect(imported.workbookName).toBe('metrics')
    expect(imported.sheetNames).toEqual(['metrics'])
    expect(imported.snapshot.sheets[0]).toMatchObject({
      name: 'metrics',
      cells: [
        { address: 'A1', value: 'Name' },
        { address: 'B1', value: 'Value' },
        { address: 'A2', value: 'alpha' },
        { address: 'B2', value: 12 },
        { address: 'A3', value: 'beta' },
        { address: 'B3', formula: 'A2' },
      ],
    })
    expect(readRuntimeImage(imported.snapshot)?.sheetCells).toEqual([
      {
        sheetName: 'metrics',
        coords: [
          { row: 0, col: 0 },
          { row: 0, col: 1 },
          { row: 1, col: 0 },
          { row: 1, col: 1 },
          { row: 2, col: 0 },
          { row: 2, col: 1 },
        ],
        coordinateOrder: 'dense-row-major',
        dimensions: { width: 2, height: 3 },
        cellCount: 6,
      },
    ])
    const engine = new SpreadsheetEngine({ workbookName: imported.workbookName, replicaId: 'csv-formula-runtime-image-restore' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    expect(engine.getCellValue('metrics', 'B3')).toEqual({ tag: ValueTag.String, value: 'alpha', stringId: expect.any(Number) })
    expect(imported.preview).toMatchObject({
      workbookName: 'metrics',
      sheetCount: 1,
      sheets: [
        {
          name: 'metrics',
          rowCount: 3,
          columnCount: 2,
          nonEmptyCellCount: 6,
          previewRows: [
            ['Name', 'Value'],
            ['alpha', '12'],
            ['beta', '=A2'],
          ],
        },
      ],
    })
  })

  it('attaches runtime coordinates for literal-only dense csv imports', () => {
    const imported = importCsv('Name,Value\nalpha,12\nbeta,24', 'literal-metrics.csv')
    const runtimeImage = readRuntimeImage(imported.snapshot)

    expect(runtimeImage?.sheetCells).toEqual([
      {
        sheetName: 'literal-metrics',
        coords: [
          { row: 0, col: 0 },
          { row: 0, col: 1 },
          { row: 1, col: 0 },
          { row: 1, col: 1 },
          { row: 2, col: 0 },
          { row: 2, col: 1 },
        ],
        coordinateOrder: 'dense-row-major',
        dimensions: { width: 2, height: 3 },
        cellCount: 6,
      },
    ])
  })

  it('parses common accounting number formats from csv imports', () => {
    const imported = importCsv(
      'Account,Amount,Margin,Variance\nRevenue,"$1,234.56",12.5%,"$1,234.56"\nCOGS,"($987.65)",-3.25%,"(987.65)"',
      'accounting.csv',
    )

    expect(imported.warnings).toEqual([])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 1234.56 }),
        expect.objectContaining({ address: 'C2', row: 1, col: 2, value: 0.125 }),
        expect.objectContaining({ address: 'D2', row: 1, col: 3, value: 1234.56 }),
        expect.objectContaining({ address: 'B3', row: 2, col: 1, value: -987.65 }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, value: -0.0325 }),
        expect.objectContaining({ address: 'D3', row: 2, col: 3, value: -987.65 }),
      ]),
    )
  })

  it('imports semicolon-delimited accounting csv files with decimal commas', () => {
    const imported = importCsv('Account;Amount;Tax\n4000;125,50;20,08\n5000;-12,25;0,00', 'locale-accounting.csv')

    expect(imported.warnings).toEqual([])
    expect(imported.preview.sheets[0]?.previewRows).toEqual([
      ['Account', 'Amount', 'Tax'],
      ['4000', '125,50', '20,08'],
      ['5000', '-12,25', '0,00'],
    ])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A1', row: 0, col: 0, value: 'Account' }),
        expect.objectContaining({ address: 'B1', row: 0, col: 1, value: 'Amount' }),
        expect.objectContaining({ address: 'C1', row: 0, col: 2, value: 'Tax' }),
        expect.objectContaining({ address: 'A2', row: 1, col: 0, value: '4000' }),
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 125.5 }),
        expect.objectContaining({ address: 'C2', row: 1, col: 2, value: 20.08 }),
        expect.objectContaining({ address: 'A3', row: 2, col: 0, value: '5000' }),
        expect.objectContaining({ address: 'B3', row: 2, col: 1, value: -12.25 }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, value: 0 }),
      ]),
    )
  })

  it('dispatches workbook import by content type', () => {
    const imported = importWorkbookFile(new TextEncoder().encode('A,B\n1,2'), 'dispatch.csv', CSV_CONTENT_TYPE)

    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['dispatch'])
  })

  it('normalizes workbook import content type parameters and case before dispatching', () => {
    const csvBytes = new TextEncoder().encode('A,B\n1,2')
    const csvVariants = [' text/csv; charset=utf-8 ', 'TEXT/CSV']
    for (const contentType of csvVariants) {
      const imported = importWorkbookFile(csvBytes, 'dispatch.csv', contentType)
      expect(imported.preview.contentType).toBe(CSV_CONTENT_TYPE)
      expect(imported.sheetNames).toEqual(['dispatch'])
    }

    const xlsxBytes = buildWorkbook()
    const xlsxVariants = [`${XLSX_CONTENT_TYPE}; charset=binary`, XLSX_CONTENT_TYPE.toUpperCase()]
    for (const contentType of xlsxVariants) {
      const imported = importWorkbookFile(xlsxBytes, 'dispatch.xlsx', contentType)
      expect(imported.preview.contentType).toBe(XLSX_CONTENT_TYPE)
      expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
    }
  })

  it('dispatches binary Excel workbooks by XLSB content type', () => {
    const imported = importWorkbookFile(buildBinaryWorkbook(), 'dispatch.xlsb', 'application/vnd.ms-excel.sheet.binary.macroEnabled.12')

    expect(imported.preview.contentType).toBe(XLSB_CONTENT_TYPE)
    expect(imported.workbookName).toBe('dispatch')
    expect(imported.sheetNames).toEqual(['Sheet1', 'Sheet2'])
  })

  it('dispatches legacy Excel workbooks by XLS content type', () => {
    expect(EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES).toContain(LEGACY_XLS_CONTENT_TYPE)
    expect(WORKBOOK_IMPORT_CONTENT_TYPES).toContain(LEGACY_XLS_CONTENT_TYPE)

    const imported = importWorkbookFile(buildLegacyWorkbook(), 'legacy-salary.xls', 'application/vnd.ms-excel; charset=binary')

    expect(imported.preview.contentType).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(imported.workbookName).toBe('legacy-salary')
    expect(imported.sheetNames).toEqual(['Salary'])
    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A2', value: 'Operations' }),
        expect.objectContaining({ address: 'B3', value: 1800 }),
        expect.objectContaining({ address: 'C2', value: 3050 }),
      ]),
    )
  })

  it('imports namespaced spreadsheet formulas as executable formulas', async () => {
    const imported = importWorkbookFile(buildNamespacedFormulaWorkbook(), 'legacy-expenses.xls', 'application/vnd.ms-excel')
    const importedCells = imported.snapshot.sheets[0]?.cells

    expect(imported.preview.contentType).toBe(LEGACY_XLS_CONTENT_TYPE)
    expect(imported.preview.sheets[0]?.previewRows[2]).toEqual(['=SUM(A1:A2)', '=SUM(A1:A2)'])
    expect(importedCells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A3', formula: 'SUM(A1:A2)', value: 3 }),
        expect.objectContaining({ address: 'B3', formula: 'SUM(A1:A2)', value: 3 }),
      ]),
    )

    const engine = new SpreadsheetEngine({ workbookName: 'namespaced-formula-import' })
    await engine.ready()
    engine.importSnapshot(imported.snapshot)
    engine.recalculateNow()

    expect(engine.getCellValue('Expenses', 'A3')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Expenses', 'B3')).toEqual({ tag: ValueTag.Number, value: 3 })
  })

  it('dispatches macro-enabled Excel workbooks by standard XLSM content type', () => {
    expect(EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES).toContain(XLSM_CONTENT_TYPE)
    expect(WORKBOOK_IMPORT_CONTENT_TYPES).toContain(XLSM_CONTENT_TYPE)

    const imported = importWorkbookFile(
      buildMacroEnabledWorkbook(),
      'Macro Workbook.xlsm',
      'application/vnd.ms-excel.sheet.macroEnabled.12; charset=binary',
    )

    expect(imported.preview.contentType).toBe(XLSM_CONTENT_TYPE)
    expect(imported.workbookName).toBe('Macro Workbook')
    expect(imported.warnings).toContain('Macros were preserved but not executed during XLSX import.')
    expect(imported.snapshot.workbook.metadata?.macroPayloads).toEqual([
      expect.objectContaining({
        kind: 'vbaProject',
        dataBase64: 'AQIDBA==',
        byteLength: 4,
        preservedWithoutExecution: true,
      }),
    ])
  })

  it('rejects corrupt zip-backed xlsx packages before parsing', () => {
    const bytes = buildCorruptZipBackedWorkbook()

    expect(() => importXlsx(bytes, 'corrupt.xlsx')).toThrow(InvalidXlsxZipContainerError)
    expect(() => importXlsx(bytes, 'corrupt.xlsx')).toThrow('Invalid or corrupt XLSX zip container')
    expect(() => importWorkbookFile(bytes, 'corrupt.xlsx', XLSX_CONTENT_TYPE)).toThrow(InvalidXlsxZipContainerError)
  })

  it('bounds whole-column defined names to the imported sheet extent', () => {
    const imported = importXlsx(buildWholeColumnDefinedNamesWorkbook(), 'nyse.xlsx')

    expect(imported.snapshot.workbook.metadata?.definedNames).toEqual([
      { name: 'Symbol', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'A1', endAddress: 'A3' } },
      { name: 'Total_Revenue', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'C1', endAddress: 'C3' } },
      { name: 'Year_num', value: { kind: 'range-ref', sheetName: 'Projectdata_NYSE', startAddress: 'B1', endAddress: 'B3' } },
    ])
  })
})
