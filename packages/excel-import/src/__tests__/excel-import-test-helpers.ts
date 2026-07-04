import { writeSimpleXlsxWorkbook, type SimpleXlsxCell, type SimpleXlsxSheet } from '@bilig/xlsx'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
export { projectWorkbookSemanticSnapshot as projectSupportedSnapshotSemantics, readRuntimeImage, SpreadsheetEngine } from '@bilig/core'
export { ErrorCode, ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
export { writeSimpleXlsxWorkbook, type SimpleXlsxCell, type SimpleXlsxSheet } from '@bilig/xlsx'
export { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
export { describe, expect, it } from 'vitest'
export {
  CSV_CONTENT_TYPE,
  EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES,
  exportXlsx,
  externalWorkbookCompanionAmbiguousMatchWarning,
  externalWorkbookCompanionNoMatchWarning,
  externalWorkbookReferencesWarning,
  importCsv,
  importWorkbookFile,
  importXlsx,
  InvalidXlsxZipContainerError,
  LEGACY_XLS_CONTENT_TYPE,
  readImportedXlsxCellStyle,
  volatileFormulasWarning,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
} from '../index.js'
export { importXlsxFromZipByteSource } from '../xlsx-byte-source-import.js'
export { buildBinaryWorkbook, buildLegacyWorkbook, buildNamespacedFormulaWorkbook } from './sheetjs-legacy-workbook-fixtures.js'

export const relationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const commentsRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
export const vmlDrawingRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
export const commentsContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
export const vmlDrawingContentType = 'application/vnd.openxmlformats-officedocument.vmlDrawing'

export function buildWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      definedNames: [
        { name: 'InputValue', formula: 'Sheet1!$A$1' },
        { name: 'InputBlock', formula: 'Sheet1!$A$1:$B$2' },
      ],
      sheets: [
        {
          name: 'Sheet1',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 1 },
            { address: 'B1', row: 0, col: 1, value: 2 },
            { address: 'C1', row: 0, col: 2, formula: 'A1+B1', numberFormat: '0.00' },
            { address: 'A2', row: 1, col: 0, value: 3 },
          ],
          columns: [
            { index: 0, size: 120 },
            { index: 1, size: 65 },
            { index: 2, size: 80 },
          ],
          rows: [
            { index: 0, size: 30 },
            { index: 1, size: 18 },
          ],
          merges: [{ startAddress: 'A4', endAddress: 'B4' }],
          dimension: { s: { r: 0, c: 0 }, e: { r: 1, c: 2 } },
        },
        {
          name: 'Sheet2',
          cells: [
            { address: 'A1', row: 0, col: 0, value: 'hello' },
            { address: 'A2', row: 1, col: 0, value: true },
          ],
        },
      ],
    }),
  )
  addLegacyCommentToSheet(zip, 2, { ref: 'A1', author: 'Greg', body: 'comment' })
  return zipSync(zip)
}

export function buildExternalLinkCacheWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Report',
          cells: [{ address: 'A1', row: 0, col: 0, formula: "'[1]External Data'!A1+'[1]External Data'!A2", value: 5 }],
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="External Data"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="A1"><v>2</v></cell></row>',
      '<row r="2"><cell r="A2"><v>3</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}

export function buildVolatileFormulaWorkbook(formula: string): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Model',
        cells: [
          { address: 'A1', row: 0, col: 0, value: 1 },
          { address: 'B1', row: 0, col: 1, formula, value: 1 },
        ],
      },
    ],
  })
}

export function buildExternalLinkRangeCacheWorkbook(
  criteriaFormula = "SUMPRODUCT('[1]Rates'!$B$2:$B$4,--('[1]Rates'!$A$2:$A$4=\"C\"))*B1",
  target = 'file:///tmp/rates.xlsx',
): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Model',
          cells: [
            { address: 'B1', row: 0, col: 1, value: 2 },
            { address: 'C1', row: 0, col: 2, formula: "SUM('[1]Rates'!$B$2:$B$4)*B1", value: 120 },
            { address: 'C2', row: 1, col: 2, formula: "_xlfn.XLOOKUP(\"B\",'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1", value: 40 },
            { address: 'C3', row: 2, col: 2, formula: criteriaFormula, value: 60 },
          ],
          dimension: { s: { r: 0, c: 0 }, e: { r: 2, c: 2 } },
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="Rates"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>',
      '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
      '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
      '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
      '</Relationships>',
  )
  return zipSync(zip)
}

export function buildRatesWorkbook(rates: readonly [number, number, number]): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Rates',
        cells: cellsFromRows([
          ['SKU', 'Rate'],
          ['A', rates[0]],
          ['B', rates[1]],
          ['C', rates[2]],
        ]),
      },
    ],
  })
}

export function buildSparseRatesWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Rates',
        cells: [
          ...cellsFromRows([
            ['SKU', 'Rate'],
            ['A', 20],
            ['B', null],
            ['C', 50],
          ]),
          { address: 'A5', row: 4, col: 0, value: 'D' },
          { address: 'B5', row: 4, col: 1, error: '#N/A' },
          { address: 'A6', row: 5, col: 0, value: 'E' },
          { address: 'B6', row: 5, col: 1, error: '#NULL!' },
        ],
      },
    ],
  })
}

export function buildSparseExternalLinkRangeCacheWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Model',
          cells: [
            { address: 'B1', row: 0, col: 1, value: 1 },
            { address: 'C1', row: 0, col: 2, formula: "SUM('[1]Rates'!$B$2:$B$4)*B1", value: 60 },
            { address: 'C2', row: 1, col: 2, formula: "IFERROR(SUM('[1]Rates'!$B$2:$B$5),99)", value: 60 },
            { address: 'C3', row: 2, col: 2, formula: "IFERROR(SUM('[1]Rates'!$B$6),88)", value: 60 },
          ],
          dimension: { s: { r: 0, c: 0 }, e: { r: 2, c: 2 } },
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="Rates"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="2"><cell r="B2"><v>10</v></cell></row>',
      '<row r="3"><cell r="B3"><v>20</v></cell></row>',
      '<row r="4"><cell r="B4"><v>30</v></cell></row>',
      '<row r="5"><cell r="B5"><v>40</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/rates.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}

export function readExternalLinkCacheXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  return strFromU8(zip['xl/externalLinks/externalLink5.xml'] ?? new Uint8Array())
}

export function inflateXlsxForDenseSheetJsParse(bytes: Uint8Array): Uint8Array {
  const zip = unzipSync(bytes)
  const filler = new Uint8Array(1_100_000)
  let state = 0x12345678
  for (let index = 0; index < filler.length; index += 1) {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    filler[index] = state & 0xff
  }
  zip['customXml/dense-parse-filler.bin'] = filler
  return zipSync(zip, { level: 0 })
}

export function byteSourceFor(bytes: Uint8Array): { readonly byteLength: number; readRange(start: number, end: number): Uint8Array } {
  return {
    byteLength: bytes.byteLength,
    readRange(start, end) {
      return bytes.subarray(start, end)
    },
  }
}

export function buildExternalGetPivotDataLinkCacheWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Report',
          cells: [
            {
              address: 'A1',
              row: 0,
              col: 0,
              formula: 'GETPIVOTDATA("Amount",\'[1]External Pivot\'!$G$3,"Region","East")',
              value: 15,
            },
          ],
        },
      ],
    }),
  )
  zip['xl/workbook.xml'] = strToU8(
    strFromU8(zip['xl/workbook.xml'])
      .replace(/<workbook\b([^>]*)>/u, (match) =>
        match.includes('xmlns:r=')
          ? match
          : match.replace('>', ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'),
      )
      .replace('</workbook>', '<externalReferences><externalReference r:id="rId99"/></externalReferences></workbook>'),
  )
  zip['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(zip['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
    ),
  )
  zip['xl/externalLinks/externalLink5.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<externalBook r:id="rId1">',
      '<sheetNames><sheetName val="External Pivot"/></sheetNames>',
      '<sheetDataSet><sheetData sheetId="0">',
      '<row r="3"><cell r="G3" t="str"><v>Row Labels</v></cell></row>',
      '</sheetData></sheetDataSet>',
      '</externalBook>',
      '</externalLink>',
    ].join(''),
  )
  zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///tmp/pivot-source.xlsx" TargetMode="External"/>' +
      '</Relationships>',
  )
  return zipSync(zip)
}

export function buildUnsupportedFunctionCacheWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Model',
        cells: [
          { address: 'A1', row: 0, col: 0, formula: '_xldudf_WISEPRICE(B1,"Shares Outstanding")', value: 14935800000 },
          { address: 'B1', row: 0, col: 1, value: 'AAPL' },
          { address: 'C1', row: 0, col: 2, formula: '_FV(B1,"Ticker symbol",TRUE)', value: 'AAPL' },
        ],
      },
    ],
  })
}

export function buildMacroEnabledWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    macro: {
      vbaProject: new Uint8Array([1, 2, 3, 4]),
      workbookCodeName: 'ThisWorkbook',
      sheetCodeNames: [{ sheetName: 'Sheet1', codeName: 'Sheet1' }],
    },
    sheets: [
      {
        name: 'Sheet1',
        cells: [{ address: 'A1', row: 0, col: 0, value: 'safe value' }],
      },
    ],
  })
}

export function cellsFromRows(rows: readonly (readonly (string | number | boolean | null | undefined)[])[]): SimpleXlsxCell[] {
  return rows.flatMap((row, rowIndex) =>
    row.flatMap((value, colIndex) =>
      value === null || value === undefined
        ? []
        : [
            {
              address: `${columnName(colIndex)}${String(rowIndex + 1)}`,
              row: rowIndex,
              col: colIndex,
              value,
            },
          ],
    ),
  )
}

export function axisSizes(sizes: readonly number[]): SimpleXlsxSheet['columns'] {
  return sizes.map((size, index) => ({ index, size }))
}

export function columnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}

export function addLegacyCommentToSheet(
  zip: Record<string, Uint8Array>,
  sheetNumber: number,
  comment: { readonly ref: string; readonly author: string; readonly body: string },
): void {
  const sheetPath = `xl/worksheets/sheet${String(sheetNumber)}.xml`
  const relsPath = `xl/worksheets/_rels/sheet${String(sheetNumber)}.xml.rels`
  zip[sheetPath] = strToU8(
    strFromU8(zip[sheetPath] ?? new Uint8Array()).replace('</worksheet>', '<legacyDrawing r:id="rIdLegacyCommentVml1"/></worksheet>'),
  )
  zip[relsPath] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      `<Relationships xmlns="${relationshipNamespace}">`,
      `<Relationship Id="rIdLegacyCommentVml1" Type="${vmlDrawingRelationshipType}" Target="../drawings/vmlDrawing1.vml"/>`,
      `<Relationship Id="rIdLegacyComments1" Type="${commentsRelationshipType}" Target="../comments1.xml"/>`,
      '</Relationships>',
    ].join(''),
  )
  zip['xl/comments1.xml'] = strToU8(
    [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<authors><author>${escapeXmlText(comment.author)}</author></authors>`,
      '<commentList>',
      `<comment ref="${escapeXmlAttribute(comment.ref)}" authorId="0"><text><t>${escapeXmlText(comment.body)}</t></text></comment>`,
      '</commentList>',
      '</comments>',
    ].join(''),
  )
  zip['xl/drawings/vmlDrawing1.vml'] = strToU8(legacyCommentVmlXml(comment.ref))
  zip['[Content_Types].xml'] = strToU8(
    addContentTypeOverride(
      addContentTypeDefault(strFromU8(zip['[Content_Types].xml'] ?? new Uint8Array()), 'vml', vmlDrawingContentType),
      '/xl/comments1.xml',
      commentsContentType,
    ),
  )
}

export function legacyCommentVmlXml(ref: string): string {
  const rowIndex = Math.max(0, Number(ref.replace(/^[A-Z]+/iu, '')) - 1)
  const columnIndex =
    ref
      .replace(/[0-9]+$/u, '')
      .toUpperCase()
      .split('')
      .reduce((column, character) => column * 26 + character.charCodeAt(0) - 64, 0) - 1
  return [
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
    '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
    '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
    '</v:shapetype>',
    '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">',
    '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>',
    '<v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"/></v:textbox>',
    '<x:ClientData ObjectType="Note">',
    '<x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor>',
    `<x:Row>${String(rowIndex)}</x:Row>`,
    `<x:Column>${String(columnIndex)}</x:Column>`,
    '</x:ClientData>',
    '</v:shape>',
    '</xml>',
  ].join('')
}

export function buildSingleCellMergeWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Sheet1',
        cells: cellsFromRows([['A', 'B']]),
        merges: [
          { startAddress: 'A1', endAddress: 'A1' },
          { startAddress: 'A1', endAddress: 'B1' },
        ],
      },
    ],
  })
}

export function buildZeroSizeMetadataWorkbook(): Uint8Array {
  const zip = unzipSync(
    writeSimpleXlsxWorkbook({
      sheets: [{ name: 'Sheet1', cells: cellsFromRows([['Value']]) }],
    }),
  )
  zip['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      .replace('<sheetData>', '<cols><col min="1" max="1" width="0" customWidth="1"/></cols><sheetData>')
      .replace('<row r="1">', '<row r="1" ht="0" customHeight="1">'),
  )
  return zipSync(zip)
}

export function buildExternalDefinedNamesWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    definedNames: [
      { name: 'ExternalRange', formula: '[1]Sheet1!$A$1:$A$2' },
      { name: 'ExternalBrokenRef', formula: '[2]Sheet1!#REF!' },
    ],
    sheets: [{ name: 'Sheet1', cells: cellsFromRows([['local']]) }],
  })
}

export function buildScopedDefinedNamesWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    definedNames: [
      { name: 'LocalBonus', formula: 'Global!$A$1' },
      { name: 'LocalBonus', localSheetIndex: 1, formula: 'Local!$A$1' },
      { name: 'LocalRevenue', localSheetIndex: 1, formula: 'Local!$B$1' },
    ],
    sheets: [
      { name: 'Global', cells: cellsFromRows([[100]]) },
      {
        name: 'Local',
        cells: [...cellsFromRows([[7, 10]]), { address: 'C1', row: 0, col: 2, formula: 'LocalBonus*LocalRevenue', value: 70 }],
      },
    ],
  })
}

export function buildWholeColumnDefinedNamesWorkbook(): Uint8Array {
  return writeSimpleXlsxWorkbook({
    definedNames: [
      { name: 'Symbol', formula: 'Projectdata_NYSE!$A:$A' },
      { name: 'Year_num', formula: 'Projectdata_NYSE!$B:$B' },
      { name: 'Total_Revenue', formula: 'Projectdata_NYSE!$C:$C' },
    ],
    sheets: [
      {
        name: 'Projectdata_NYSE',
        cells: cellsFromRows([
          ['Symbol', 'Year', 'Revenue'],
          ['AAA', 2020, 100],
          ['BBB', 2021, 200],
        ]),
      },
    ],
  })
}

export function addContentTypeDefault(contentTypesXml: string, extension: string, contentType: string): string {
  if (new RegExp(`<Default\\b[^>]*\\bExtension=(["'])${escapeRegExp(extension)}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Default Extension="${escapeXmlAttribute(extension)}" ContentType="${escapeXmlAttribute(contentType)}"/></Types>`,
  )
}

export function addContentTypeOverride(contentTypesXml: string, partName: string, contentType: string): string {
  if (new RegExp(`<Override\\b[^>]*\\bPartName=(["'])${escapeRegExp(partName)}\\1`, 'u').test(contentTypesXml)) {
    return contentTypesXml
  }
  return contentTypesXml.replace(
    '</Types>',
    `<Override PartName="${escapeXmlAttribute(partName)}" ContentType="${escapeXmlAttribute(contentType)}"/></Types>`,
  )
}

export function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}

export function escapeXmlText(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

export function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replaceAll('"', '&quot;').replaceAll("'", '&apos;')
}

export function readZipUint16(bytes: Uint8Array, offset: number): number {
  const low = bytes[offset]
  const high = bytes[offset + 1]
  if (low === undefined || high === undefined) {
    throw new Error('Invalid ZIP fixture')
  }
  return low | (high << 8)
}

export function buildCorruptZipBackedWorkbook(): Uint8Array {
  const bytes = zipSync({ 'xl/workbook.xml': strToU8('a'.repeat(1000)) })
  const nameLength = readZipUint16(bytes, 26)
  const extraLength = readZipUint16(bytes, 28)
  const compressedDataStart = 30 + nameLength + extraLength
  const originalByte = bytes[compressedDataStart]
  if (originalByte === undefined) {
    throw new Error('Invalid ZIP fixture')
  }
  const corrupted = new Uint8Array(bytes)
  corrupted[compressedDataStart] = originalByte ^ 0xff
  return corrupted
}

export function buildGenericWorkflowWorkbookFixture(shape: 'multi-sheet-operations' | 'single-sheet-planning'): Uint8Array {
  if (shape === 'multi-sheet-operations') {
    return writeSimpleXlsxWorkbook({
      sheets: [
        {
          name: 'Dashboard',
          cells: [
            ...cellsFromRows([
              ['OPERATIONS DASHBOARD', null, null, null],
              [],
              ['Metric', 'Value'],
              ['Total budget'],
              ['Open balance'],
              ['Completion rate'],
            ]),
            { address: 'B4', row: 3, col: 1, formula: 'SUM(Ledger!F:F)' },
            { address: 'B5', row: 4, col: 1, formula: 'SUMIF(Ledger!H:H,"Open",Ledger!G:G)' },
            { address: 'B6', row: 5, col: 1, formula: 'IF(B4>0,1-B5/B4,0)' },
          ],
          columns: axisSizes([180, 118, 96, 96]),
          rows: [
            { index: 0, size: 30 },
            { index: 2, size: 24 },
          ],
          merges: [{ startAddress: 'A1', endAddress: 'D1' }],
        },
        {
          name: 'Ledger',
          cells: [
            ...cellsFromRows([
              ['OPERATIONS LEDGER', null, null, null, null, null, null, null],
              [],
              ['ID', 'Date', 'Owner', 'Workstream', 'Category', 'Budget', 'Open Balance', 'Status'],
              ['OP001', 45292, 'Facilities', 'Office refresh', 'Capital', 12000, null, 'Open'],
              ['OP002', 45323, 'Engineering', 'Data migration', 'Platform', 18000, null, 'Open'],
            ]),
            { address: 'G4', row: 3, col: 6, formula: 'F4-SUMIF(Rollforward!$B:$B,A4,Rollforward!$E:$E)' },
            { address: 'G5', row: 4, col: 6, formula: 'F5-SUMIF(Rollforward!$B:$B,A5,Rollforward!$E:$E)' },
          ],
          columns: axisSizes([132, 96, 142, 210, 138, 118, 138, 92]),
          rows: [
            { index: 0, size: 30 },
            { index: 2, size: 24 },
          ],
          merges: [{ startAddress: 'A1', endAddress: 'H1' }],
        },
        {
          name: 'Rollforward',
          cells: [
            ...cellsFromRows([
              ['ROLLFORWARD', null, null, null, null],
              [],
              ['Period', 'Item ID', 'Description', 'Monthly Change', 'Cumulative Change'],
              ['Jan 2024', 'OP001', 'Office refresh'],
              ['Feb 2024', 'OP001', 'Office refresh'],
              ['Mar 2024', 'OP002', 'Data migration'],
            ]),
            { address: 'D4', row: 3, col: 3, formula: 'VLOOKUP(B4,Ledger!A:F,6,FALSE())/12' },
            { address: 'E4', row: 3, col: 4, formula: 'D4' },
            { address: 'D5', row: 4, col: 3, formula: 'VLOOKUP(B5,Ledger!A:F,6,FALSE())/12' },
            { address: 'E5', row: 4, col: 4, formula: 'IF(B5=B4,E4+D5,D5)' },
            { address: 'D6', row: 5, col: 3, formula: 'VLOOKUP(B6,Ledger!A:F,6,FALSE())/12' },
            { address: 'E6', row: 5, col: 4, formula: 'IF(B6=B5,E5+D6,D6)' },
          ],
          columns: axisSizes([112, 96, 210, 126, 148]),
          rows: [
            { index: 0, size: 30 },
            { index: 2, size: 24 },
          ],
          merges: [{ startAddress: 'A1', endAddress: 'E1' }],
        },
        {
          name: 'Lookups',
          cells: cellsFromRows([['Category'], ['Capital'], ['Platform']]),
        },
      ],
    })
  }

  return writeSimpleXlsxWorkbook({
    sheets: [
      {
        name: 'Monthly Plan',
        cells: [
          ...cellsFromRows([
            ['Monthly Planning Schedule', null, null, null, null, null, null, null, null],
            ['Owner', 'Workstream', 'Start Date', 'End Date', 'Budget', 'Jan 2026', 'Feb 2026', 'Planned', 'Remaining'],
            ['TenantWorks', 'Facilities platform', 46054, 46234, 6600],
            ['Blue Harbor', 'Insurance binder', 46023, 46388, 12000],
          ]),
          {
            address: 'F3',
            row: 2,
            col: 5,
            formula: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,1,1),0))-MAX($C3,DATE(2026,1,1))+1)/($D3-$C3+1),0),2)',
          },
          {
            address: 'G3',
            row: 2,
            col: 6,
            formula: 'ROUND(IFERROR($E3*MAX(0,MIN($D3,EOMONTH(DATE(2026,2,1),0))-MAX($C3,DATE(2026,2,1))+1)/($D3-$C3+1),0),2)',
          },
          { address: 'H3', row: 2, col: 7, formula: 'ROUND(SUM(F3:G3),2)' },
          { address: 'I3', row: 2, col: 8, formula: 'ROUND(E3-H3,2)' },
          {
            address: 'F4',
            row: 3,
            col: 5,
            formula: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,1,1),0))-MAX($C4,DATE(2026,1,1))+1)/($D4-$C4+1),0),2)',
          },
          {
            address: 'G4',
            row: 3,
            col: 6,
            formula: 'ROUND(IFERROR($E4*MAX(0,MIN($D4,EOMONTH(DATE(2026,2,1),0))-MAX($C4,DATE(2026,2,1))+1)/($D4-$C4+1),0),2)',
          },
          { address: 'H4', row: 3, col: 7, formula: 'ROUND(SUM(F4:G4),2)' },
          { address: 'I4', row: 3, col: 8, formula: 'ROUND(E4-H4,2)' },
        ],
        columns: axisSizes([168, 190, 104, 104, 118, 96, 96, 134, 138]),
        rows: [
          { index: 0, size: 30 },
          { index: 1, size: 24 },
        ],
        merges: [{ startAddress: 'A1', endAddress: 'I1' }],
      },
    ],
  })
}
