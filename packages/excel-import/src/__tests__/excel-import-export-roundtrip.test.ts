import type { WorkbookSnapshot } from './excel-import-test-helpers.js'
import {
  describe,
  expect,
  exportXlsx,
  importXlsx,
  it,
  projectSupportedSnapshotSemantics,
  strFromU8,
  strToU8,
  unzipSync,
  zipSync,
} from './excel-import-test-helpers.js'

describe('excel import/export roundtrip and formatting fidelity', () => {
  it('exports workbook snapshots to XLSX bytes that import back with supported workbook semantics', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'Roundtrip Workbook',
        metadata: {
          calculationSettings: { mode: 'manual', compatibilityMode: 'excel-modern' },
          properties: [
            { key: 'locale', value: 'en-US' },
            { key: 'reviewed', value: true },
            { key: 'threshold', value: 0.085 },
          ],
          definedNames: [
            { name: 'SummaryTotal', value: { kind: 'cell-ref', sheetName: 'Summary', address: 'B1' } },
            { name: 'InputRegion', value: { kind: 'range-ref', sheetName: 'Inputs', startAddress: 'A1', endAddress: 'B1' } },
            { name: 'TaxRate', value: { kind: 'scalar', value: 0.085 } },
          ],
          styles: [
            {
              id: 'accent-total',
              fill: { backgroundColor: '#1d3989' },
              font: { family: 'Aptos', size: 12, bold: true, color: '#ffffff' },
              alignment: { horizontal: 'center', vertical: 'middle', wrap: true },
              borders: { bottom: { style: 'solid', weight: 'thin', color: '#000000' } },
            },
          ],
          tables: [
            {
              name: 'Input.Table',
              sheetName: 'Inputs',
              startAddress: 'A1',
              endAddress: 'D4',
              columnNames: ['Region', 'Product', 'Sales', 'Notes'],
              headerRow: true,
              totalsRow: false,
            },
          ],
          charts: [
            {
              id: 'summary-trend',
              sheetName: 'Summary',
              address: 'E1',
              source: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
              chartType: 'line',
              seriesOrientation: 'columns',
              firstRowAsHeaders: true,
              firstColumnAsLabels: true,
              title: 'Summary Trend',
              legendPosition: 'right',
              rows: 12,
              cols: 6,
              anchor: {
                kind: 'twoCell',
                editAs: 'twoCell',
                from: { row: 0, col: 4, rowOffset: 0, colOffset: 0 },
                to: { row: 12, col: 10, rowOffset: 0, colOffset: 0 },
              },
            },
          ],
          pivots: [
            {
              name: 'SalesByRegion',
              sheetName: 'Summary',
              address: 'E15',
              source: { sheetName: 'Inputs', startAddress: 'A1', endAddress: 'D4' },
              groupBy: ['Region'],
              values: [
                { sourceColumn: 'Sales', summarizeBy: 'sum', outputLabel: 'Total Sales' },
                { sourceColumn: 'Product', summarizeBy: 'count', outputLabel: 'Rows' },
              ],
              rows: 4,
              cols: 3,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Summary',
          order: 0,
          metadata: {
            styleRanges: [{ range: { sheetName: 'Summary', startAddress: 'B1', endAddress: 'B1' }, styleId: 'accent-total' }],
            commentThreads: [
              {
                threadId: 'summary-total-note',
                sheetName: 'Summary',
                address: 'B1',
                comments: [{ id: 'summary-total-note-1', body: 'Reviewed total', authorDisplayName: 'Finance' }],
              },
            ],
            columns: [
              { id: 'summary-col-0', index: 0, size: 132 },
              { id: 'summary-col-1', index: 1, size: 96 },
            ],
            rows: [
              { id: 'summary-row-0', index: 0, size: 30 },
              { id: 'summary-row-2', index: 2, size: 24 },
            ],
            freezePane: { rows: 1, cols: 2 },
            merges: [{ sheetName: 'Summary', startAddress: 'A5', endAddress: 'B5' }],
            sheetProtection: { sheetName: 'Summary' },
            protectedRanges: [
              {
                id: 'protect-summary-inputs',
                range: { sheetName: 'Summary', startAddress: 'A2', endAddress: 'B3' },
              },
            ],
            filters: [{ sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' }],
            sorts: [
              {
                range: { sheetName: 'Summary', startAddress: 'A1', endAddress: 'B3' },
                keys: [{ keyAddress: 'B1', direction: 'desc' }],
              },
            ],
            validations: [
              {
                range: { sheetName: 'Summary', startAddress: 'C2', endAddress: 'C4' },
                rule: { kind: 'whole', operator: 'between', values: [0, 100] },
                allowBlank: false,
                errorStyle: 'stop',
                errorTitle: 'Percent required',
                errorMessage: 'Enter a whole number from 0 to 100.',
              },
            ],
            conditionalFormats: [
              {
                id: 'summary-high-total',
                range: { sheetName: 'Summary', startAddress: 'B2', endAddress: 'B3' },
                rule: { kind: 'cellIs', operator: 'greaterThan', values: [1000] },
                style: { fill: { backgroundColor: '#f4cccc' }, font: { bold: true, color: '#990000' } },
                stopIfTrue: true,
                priority: 1,
              },
            ],
          },
          cells: [
            { address: 'A1', value: 'Metric' },
            { address: 'B1', formula: 'SUM(B2:B3)', format: '0.00' },
            { address: 'C1', value: true },
            { address: 'A2', value: 'Revenue' },
            { address: 'B2', value: 1250.5, format: '$#,##0.00' },
            { address: 'A3', value: 'Costs' },
            { address: 'B3', value: 450.25, format: '$#,##0.00' },
          ],
        },
        {
          id: 2,
          name: 'Inputs',
          order: 1,
          metadata: {
            validations: [
              {
                range: { sheetName: 'Inputs', startAddress: 'D2', endAddress: 'D4' },
                rule: { kind: 'list', values: ['Priority', 'Standard'] },
                allowBlank: true,
                showDropdown: true,
                promptTitle: 'Status',
                promptMessage: 'Pick a known priority.',
                errorStyle: 'warning',
                errorTitle: 'Unknown priority',
                errorMessage: 'Use Priority or Standard.',
              },
            ],
          },
          cells: [
            { address: 'A1', value: 'Region' },
            { address: 'B1', value: 'Product' },
            { address: 'C1', value: 'Sales' },
            { address: 'D1', value: 'Notes' },
            { address: 'A2', value: 'East' },
            { address: 'B2', value: 'Widget' },
            { address: 'C2', value: 10 },
            { address: 'D2', value: 'Priority' },
            { address: 'A3', value: 'West' },
            { address: 'B3', value: 'Widget' },
            { address: 'C3', value: 7 },
            { address: 'D3', value: 'Priority' },
            { address: 'A4', value: 'East' },
            { address: 'B4', value: 'Gizmo' },
            { address: 'C4', value: 5 },
            { address: 'D4', value: 'Standard' },
          ],
        },
      ],
    }

    const bytes = exportXlsx(snapshot)
    const imported = importXlsx(bytes, 'roundtrip.xlsx')
    const zip = unzipSync(bytes)

    expect(bytes.byteLength).toBeGreaterThan(0)
    expect(Object.keys(zip)).toEqual(expect.arrayContaining(['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml']))
    expect(Object.keys(zip)).toEqual(
      expect.arrayContaining([
        'xl/pivotTables/pivotTable1.xml',
        'xl/pivotCache/pivotCacheDefinition1.xml',
        'xl/pivotCache/pivotCacheRecords1.xml',
        'xl/tables/table1.xml',
      ]),
    )
    expect(strFromU8(zip['xl/charts/chart1.xml'] ?? new Uint8Array())).toContain('<c:lineChart>')
    expect(strFromU8(zip['xl/drawings/_rels/drawing1.xml.rels'] ?? new Uint8Array())).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    )
    expect(strFromU8(zip['xl/pivotTables/pivotTable1.xml'] ?? new Uint8Array())).toContain('<pivotTableDefinition')
    expect(strFromU8(zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())).toContain(
      '<worksheetSource ref="A1:D4" sheet="Inputs"/>',
    )
    expect(strFromU8(zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())).toContain('refreshOnLoad="1"')
    expect(strFromU8(zip['xl/pivotCache/pivotCacheDefinition1.xml'] ?? new Uint8Array())).toContain('recordCount="3"')
    expect(strFromU8(zip['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] ?? new Uint8Array())).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords',
    )
    expect(strFromU8(zip['xl/pivotCache/pivotCacheRecords1.xml'] ?? new Uint8Array())).toContain(
      '<r><s v="East"/><s v="Widget"/><n v="10"/><s v="Priority"/></r>',
    )
    expect(strFromU8(zip['_rels/.rels'] ?? new Uint8Array())).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
    )
    expect(strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array())).toContain(
      'application/vnd.openxmlformats-officedocument.custom-properties+xml',
    )
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('<property ')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="locale"><vt:lpwstr>en-US</vt:lpwstr>')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="reviewed"><vt:bool>true</vt:bool>')
    expect(strFromU8(zip['docProps/custom.xml'] ?? new Uint8Array())).toContain('name="threshold"><vt:r8>0.085</vt:r8>')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('<table ')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('displayName="Input.Table"')
    expect(strFromU8(zip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('<tableColumn id="3" name="Sales"/>')
    expect(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).toContain('<calcPr calcMode="manual"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<dataValidations count="1">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<dataValidation type="whole" operator="between" allowBlank="0" errorStyle="stop"',
    )
    expect(strFromU8(zip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())).toContain(
      '<dataValidation type="list" allowBlank="1" showDropDown="0"',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<conditionalFormatting sqref="B2:B3">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1"><formula>1000</formula></cfRule>',
    )
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<dxfs count="1">')
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<fgColor rgb="FFF4CCCC"/>')
    expect(strFromU8(zip['xl/styles.xml'] ?? new Uint8Array())).toContain('<color rgb="FF990000"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sheetProtection sheet="1"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<protectedRanges>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain(
      '<protectedRange name="protect-summary-inputs" sqref="A2:B3"/>',
    )
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<autoFilter ref="A1:B3"/>')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sortState ref="A1:B3">')
    expect(strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<sortCondition descending="1" ref="B1:B3"/>')
    expect(imported.snapshot.workbook.metadata?.pivots?.[0]).toMatchObject({
      cacheFields: ['Region', 'Product', 'Sales', 'Notes'],
      cachedRecords: [
        ['East', 'Widget', 10, 'Priority'],
        ['West', 'Widget', 7, 'Priority'],
        ['East', 'Gizmo', 5, 'Standard'],
      ],
    })
    expect(projectSupportedSnapshotSemantics(imported.snapshot)).toEqual(projectSupportedSnapshotSemantics(snapshot))
  })

  it('preserves imported frozen pane scroll targets on export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'frozen-pane-scroll-target' },
      sheets: [
        {
          id: 1,
          name: 'View',
          order: 0,
          metadata: {
            freezePane: { rows: 3, cols: 2, topLeftCell: 'I32', activePane: 'bottomRight' },
          },
          cells: [
            { address: 'A1', value: 'Account' },
            { address: 'B1', value: 'Amount' },
            { address: 'I32', value: 'scroll target' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'frozen-pane-scroll-target.xlsx')
    const freezePane = imported.snapshot.sheets[0]?.metadata?.freezePane
    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const sheetXml = strFromU8(exportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())

    expect(freezePane).toEqual({ rows: 3, cols: 2, topLeftCell: 'I32', activePane: 'bottomRight' })
    expect(sheetXml).toContain('<pane xSplit="2" ySplit="3" topLeftCell="I32" activePane="bottomRight" state="frozen"/>')
  })

  it('preserves unicode table names across XLSX export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: {
        name: 'unicode-table-name',
        metadata: {
          tables: [
            {
              name: 'Données_FR',
              sheetName: 'Données - FR',
              startAddress: 'A1',
              endAddress: 'B2',
              columnNames: ['MESURE', 'CATÉGORIE'],
              headerRow: true,
              totalsRow: false,
            },
          ],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Données - FR',
          order: 0,
          cells: [
            { address: 'A1', value: 'MESURE' },
            { address: 'B1', value: 'CATÉGORIE' },
            { address: 'A2', value: 'Taxes' },
            { address: 'B2', value: 'Fédéral' },
          ],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const exportedZip = unzipSync(exported)
    const imported = importXlsx(exported, 'unicode-table-name.xlsx')

    expect(strFromU8(exportedZip['xl/tables/table1.xml'] ?? new Uint8Array())).toContain('displayName="Données_FR"')
    expect(imported.snapshot.workbook.metadata?.tables?.[0]?.name).toBe('Données_FR')
  })

  it('preserves worksheet tab colors on import and export', () => {
    const snapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'tab-color-roundtrip' },
      sheets: [
        {
          id: 1,
          name: 'Rgb',
          order: 0,
          metadata: { tabColor: { rgb: 'FF0070C0' } },
          cells: [{ address: 'A1', value: 'rgb tab' }],
        },
        {
          id: 2,
          name: 'Theme',
          order: 1,
          metadata: { tabColor: { theme: '8' } },
          cells: [{ address: 'A1', value: 'theme tab' }],
        },
        {
          id: 3,
          name: 'Tint',
          order: 2,
          metadata: { tabColor: { theme: '0', tint: '-0.14999847407452621' } },
          cells: [{ address: 'A1', value: 'tint tab' }],
        },
      ],
    }

    const exported = exportXlsx(snapshot)
    const imported = importXlsx(exported, 'tab-color-roundtrip.xlsx')
    const reexportedZip = unzipSync(exportXlsx(imported.snapshot))

    expect(imported.snapshot.sheets.map((sheet) => sheet.metadata?.tabColor)).toEqual([
      { rgb: 'FF0070C0' },
      { theme: '8' },
      { theme: '0', tint: '-0.14999847407452621' },
    ])
    expect(strFromU8(reexportedZip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<tabColor rgb="FF0070C0"/>')
    expect(strFromU8(reexportedZip['xl/worksheets/sheet2.xml'] ?? new Uint8Array())).toContain('<tabColor theme="8"/>')
    expect(strFromU8(reexportedZip['xl/worksheets/sheet3.xml'] ?? new Uint8Array())).toContain(
      '<tabColor theme="0" tint="-0.14999847407452621"/>',
    )
  })

  it('exports custom number formats on populated and blank cells', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: { id: 'custom-number-format-workbook', name: 'custom-number-format-workbook' },
      sheets: [
        {
          id: 1,
          name: 'Formats',
          order: 0,
          cells: [
            { address: 'A1', value: 0, format: '00' },
            { address: 'A2', format: '00' },
            { address: 'B1', value: 12.34, format: '"$"#,##0.00' },
          ],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'custom-number-format-workbook.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'A1', row: 0, col: 0, value: 0, format: '00' }),
        expect.objectContaining({ address: 'A2', row: 1, col: 0, format: '00' }),
        expect.objectContaining({ address: 'B1', row: 0, col: 1, value: 12.34, format: '"$"#,##0.00' }),
      ]),
    )
  })

  it('roundtrips worksheet AutoFilter criteria and custom filters', () => {
    const sourceSnapshot: WorkbookSnapshot = {
      version: 1,
      workbook: { name: 'autofilter-criteria-source' },
      sheets: [
        {
          name: 'Ledger',
          order: 0,
          metadata: {
            filters: [{ sheetName: 'Ledger', startAddress: 'A1', endAddress: 'D6' }],
          },
          cells: [
            { address: 'A1', value: 'Date' },
            { address: 'B1', value: 'Department' },
            { address: 'C1', value: 'Amount' },
            { address: 'D1', value: 'Status' },
            { address: 'A2', value: '2026-01-01' },
            { address: 'B2', value: 'Finance' },
            { address: 'C2', value: -100 },
            { address: 'D2', value: 'Approved' },
            { address: 'A3', value: '2026-01-02' },
            { address: 'B3', value: 'Operations' },
            { address: 'C3', value: 75 },
            { address: 'D3', value: 'Pending' },
            { address: 'A4', value: '2026-01-03' },
            { address: 'B4', value: 'Finance' },
            { address: 'C4', value: -25 },
            { address: 'D4', value: 'Approved' },
            { address: 'A5', value: '2026-01-04' },
            { address: 'B5', value: 'Sales' },
            { address: 'C5', value: 50 },
            { address: 'D5', value: 'Rejected' },
            { address: 'A6', value: '2026-01-05' },
            { address: 'B6', value: 'Finance' },
            { address: 'C6', value: -1 },
            { address: 'D6', value: 'Approved' },
          ],
        },
      ],
    }
    const sourceZip = unzipSync(exportXlsx(sourceSnapshot))
    const sheetPath = 'xl/worksheets/sheet1.xml'
    const sourceSheetXml = strFromU8(sourceZip[sheetPath] ?? new Uint8Array())
    sourceZip[sheetPath] = strToU8(
      sourceSheetXml.replace(
        '<autoFilter ref="A1:D6"/>',
        [
          '<autoFilter ref="A1:D6">',
          '<filterColumn colId="1"><filters blank="0"><filter val="Finance"/></filters></filterColumn>',
          '<filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="0"/></customFilters></filterColumn>',
          '<filterColumn colId="3"><filters blank="0"><filter val="Approved"/></filters></filterColumn>',
          '</autoFilter>',
        ].join(''),
      ),
    )

    const imported = importXlsx(zipSync(sourceZip), 'autofilter-criteria-source.xlsx')

    expect(imported.snapshot.sheets[0]?.metadata?.filters).toEqual([
      {
        sheetName: 'Ledger',
        startAddress: 'A1',
        endAddress: 'D6',
        criteria: [
          { colId: 1, filters: { blank: false, values: ['Finance'] } },
          { colId: 2, customFilters: { filters: [{ operator: 'lessThan', value: '0' }] } },
          { colId: 3, filters: { blank: false, values: ['Approved'] } },
        ],
      },
    ])
    expect(imported.snapshot.sheets[0]?.metadata?.rowMetadata).toEqual([
      { start: 2, count: 1, filterHidden: true },
      { start: 4, count: 1, filterHidden: true },
    ])

    const exportedZip = unzipSync(exportXlsx(imported.snapshot))
    const exportedSheetXml = strFromU8(exportedZip[sheetPath] ?? new Uint8Array())

    expect(exportedSheetXml).toContain('<autoFilter ref="A1:D6">')
    expect(exportedSheetXml).toContain('<filterColumn colId="1"><filters blank="0"><filter val="Finance"/></filters></filterColumn>')
    expect(exportedSheetXml).toContain(
      '<filterColumn colId="2"><customFilters><customFilter operator="lessThan" val="0"/></customFilters></filterColumn>',
    )
    expect(exportedSheetXml).toContain('<filterColumn colId="3"><filters blank="0"><filter val="Approved"/></filters></filterColumn>')
    expect(exportedSheetXml).toContain('<row r="3" hidden="1"')
    expect(exportedSheetXml).toContain('<row r="5" hidden="1"')
  })

  it('exports range-only number formats on populated and blank cells', () => {
    const snapshot: WorkbookSnapshot = {
      workbook: {
        id: 'range-number-format-workbook',
        name: 'range-number-format-workbook',
        metadata: {
          formats: [{ id: 'format-zip-code', code: '00000', kind: 'number' }],
        },
      },
      sheets: [
        {
          id: 1,
          name: 'Formats',
          order: 0,
          metadata: {
            formatRanges: [
              {
                range: { sheetName: 'Formats', startAddress: 'B2', endAddress: 'C3' },
                formatId: 'format-zip-code',
              },
            ],
          },
          cells: [{ address: 'B2', value: 7 }],
        },
      ],
    }

    const imported = importXlsx(exportXlsx(snapshot), 'range-number-format-workbook.xlsx')

    expect(imported.snapshot.sheets[0]?.cells).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ address: 'B2', row: 1, col: 1, value: 7, format: '00000' }),
        expect.objectContaining({ address: 'C3', row: 2, col: 2, format: '00000' }),
      ]),
    )
  })
})
