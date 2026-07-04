import { readRuntimeImage } from '@bilig/core'
import { strToU8, unzipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { importXlsx, inspectXlsx, XlsxImportSizeLimitExceededError } from '../index.js'
import { importXlsxFromZipByteSource } from '../xlsx-byte-source-import.js'
import { tryImportLargeSimpleXlsx } from '../xlsx-large-simple-import.js'
import { forEachInflatedXlsxZipEntryChunk, readXlsxZipEntriesLazy } from '../xlsx-zip.js'
import {
  buildLargeSimpleWorkbook,
  byteSourceFor,
  countLazyZipEntryStreams,
  deterministicBytes,
} from './xlsx-large-simple-import-test-helpers.js'

describe('large simple XLSX import fast path', () => {
  it('imports simple OpenXML worksheets without SheetJS workbook materialization', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:E5"/>',
        '<sheetFormatPr defaultRowHeight="15"/>',
        '<cols><col min="1" max="2" width="10" customWidth="1"/></cols>',
        '<sheetData>',
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42.5</v></c></row>',
        '<row r="2" ht="20" customHeight="1"><c r="C2" t="b"><v>1</v></c></row>',
        '<row r="3"><c r="D3" t="inlineStr"><is><t>Inline text</t></is></c></row>',
        '<row r="4"><c r="E4" t="e"><v>#N/A</v></c></row>',
        '<row r="5"><c r="A5" t="s"><v>1</v></c></row>',
        '</sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'large-simple.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha &#8211;' },
      { address: 'B1', value: 42.5 },
      { address: 'C2', value: true },
      { address: 'D3', value: 'Inline text' },
      { address: 'E4', value: '#N/A' },
      { address: 'A5', value: 'Line\nBreak' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }])
    expect(imported?.snapshot.sheets[0]?.metadata).toMatchObject({
      columns: [
        { id: 'col:0', index: 0, size: 60 },
        { id: 'col:1', index: 1, size: 60 },
      ],
      columnMetadata: [{ start: 0, count: 2, size: 60, xlsxWidth: 10, customWidth: true }],
      rows: [{ id: 'row:1', index: 1, size: 27 }],
      rowMetadata: [{ start: 1, count: 1, size: 27, xlsxHeight: 20, customHeight: true }],
      sheetFormatPr: { defaultRowHeight: 15 },
    })
    expect(imported?.preview.sheets[0]).toMatchObject({
      rowCount: 5,
      columnCount: 5,
      nonEmptyCellCount: 6,
      previewRows: [
        ['Alpha &#8211;', '42.5', '', '', ''],
        ['', '', 'TRUE', '', ''],
        ['', '', '', 'Inline text', ''],
        ['', '', '', '', '#N/A'],
        ['Line\nBreak', '', '', '', ''],
      ],
    })
  })

  it('imports large value-only worksheets that omit sharedStrings.xml', () => {
    const rows: string[] = []
    for (let row = 1; row <= 2_000; row += 1) {
      rows.push(
        [
          `<row r="${String(row)}">`,
          `<c r="A${String(row)}"><v>${String(row)}</v></c>`,
          `<c r="B${String(row)}" t="inlineStr"><is><t>Row ${String(row)}</t></is></c>`,
          `<c r="C${String(row)}" t="b"><v>${row % 2 === 0 ? '1' : '0'}</v></c>`,
          '</row>',
        ].join(''),
      )
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:C2000"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'numeric-only.xlsx', unzipSync(bytes), { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toHaveLength(6_000)
    expect(imported?.snapshot.sheets[0]?.cells.slice(0, 6)).toEqual([
      { address: 'A1', value: 1 },
      { address: 'B1', value: 'Row 1' },
      { address: 'C1', value: false },
      { address: 'A2', value: 2 },
      { address: 'B2', value: 'Row 2' },
      { address: 'C2', value: true },
    ])
    expect(imported?.preview.sheets[0]).toMatchObject({
      rowCount: 2_000,
      columnCount: 3,
      nonEmptyCellCount: 6_000,
    })
    expect(readRuntimeImage(imported!.snapshot)?.sheetCells?.[0]).toMatchObject({
      sheetName: 'Data',
      coords: [],
      coordinateOrder: 'dense-row-major',
      dimensions: { width: 3, height: 2_000 },
      cellCount: 6_000,
    })
  })

  it('imports small native-only workbooks without SheetJS fallback', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B2"/>',
        '<sheetData>',
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Metric</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>',
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Revenue</t></is></c><c r="B2"><f>1200*5</f><v>6000</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(bytes.byteLength).toBeLessThan(1_000_000)
    const imported = importXlsx(bytes, 'small-native-only.xlsx', { nativeOnly: true })

    expect(imported.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Metric' },
      { address: 'B1', value: 'Value' },
      { address: 'A2', value: 'Revenue' },
      { address: 'B2', formula: '1200*5', value: 6000 },
    ])
    expect(imported.preview.sheets[0]).toMatchObject({
      name: 'Data',
      rowCount: 2,
      columnCount: 2,
      nonEmptyCellCount: 4,
    })
  })

  it('preflights public import materialization limits before building snapshot cell objects', () => {
    const rows: string[] = []
    for (let row = 1; row <= 4; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
    }
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A4"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '</worksheet>',
      ].join(''),
    })

    const inspection = inspectXlsx(bytes, 'limit-preflight.xlsx')
    expect(inspection?.stats.cellCount).toBe(4)
    expect(() =>
      importXlsx(bytes, 'limit-preflight.xlsx', {
        limits: { maxMaterializedCells: 3 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)

    const imported = importXlsx(bytes, 'limit-preflight.xlsx', { limits: { maxMaterializedCells: 4 } })
    expect(imported.snapshot.sheets[0]?.cells).toHaveLength(4)
  })

  it('preflights formula-heavy imports before WorkPaper build can hit evaluation timeout', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(() =>
      importXlsx(bytes, 'formula-limit-preflight.xlsx', {
        limits: { maxMaterializedFormulaCells: 0 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)
  })

  it('imports small simple generated XLSX files through the native path when requested', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B2"/>',
        '<sheetData>',
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Metric</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>',
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Revenue</t></is></c><c r="B2"><f>40*1200</f><v>48000</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    })

    const imported = importXlsx(bytes, 'small-native.xlsx', { preferNativeSimpleImport: true })

    expect(imported.stats?.cellCount).toBe(4)
    expect(imported.stats?.formulaCellCount).toBe(1)
    expect(imported.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Metric' },
      { address: 'B1', value: 'Value' },
      { address: 'A2', value: 'Revenue' },
      { address: 'B2', value: 48000, formula: '40*1200' },
    ])
  })

  it('preflights formula-heavy sheets even when workbook features require SheetJS fallback', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" cm="1"><v>1</v></c><c r="B1"><f>A1+1</f><v>2</v></c></row></sheetData>',
        '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>',
        '</worksheet>',
      ].join(''),
    })

    expect(inspectXlsx(bytes, 'fallback-preflight.xlsx')?.stats).toMatchObject({
      cellCount: 2,
      formulaCellCount: 1,
    })
    expect(() =>
      importXlsx(bytes, 'fallback-preflight.xlsx', {
        limits: { maxMaterializedFormulaCells: 0 },
      }),
    ).toThrow(XlsxImportSizeLimitExceededError)
  })

  it('retries data-only large workbook import before SheetJS fallback when metadata preservation rejects', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
        '<hyperlinks><hyperlink ref="A1:A2000" location="Target"/></hyperlinks>',
        '</worksheet>',
      ].join(''),
      extraEntries: {
        'docProps/padding.bin': deterministicBytes(1_200_000),
      },
    })

    const imported = importXlsxFromZipByteSource(byteSourceFor(bytes), 'data-only-retry.xlsx')

    expect(imported.stats?.phaseTelemetry.map((entry) => entry.phase)).toContain('public-snapshot-materialization')
    expect(imported.snapshot.sheets[0]?.cells).toEqual([{ address: 'A1', value: 1 }])
    expect(imported.snapshot.sheets[0]?.metadata?.hyperlinks).toBeUndefined()
  })

  it('falls back when shared string cells reference a missing sharedStrings part', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'missing-shared-strings.xlsx', unzipSync(bytes), { minByteLength: 0 })).toBeNull()
  })

  it('streams non-materialized large worksheets without inflating worksheet or shared string entries', () => {
    const rows: string[] = []
    for (let row = 1; row <= 128; row += 1) {
      rows.push(
        [
          `<row r="${String(row)}">`,
          `<c r="A${String(row)}" t="s"><v>0</v></c>`,
          `<c r="B${String(row)}"><f>A${String(row)}&amp;"!"</f><v>${String(row)}</v></c>`,
          '</row>',
        ].join(''),
      )
    }
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B128"/>',
        `<sheetData>${rows.join('')}</sheetData>`,
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '<conditionalFormatting sqref="B1:B128"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>64</formula></cfRule></conditionalFormatting>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should not be inflated for non-materialized import')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-headless.xlsx', zip, { materializeCells: false, minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([])
    expect(imported?.stats).toMatchObject({
      sheetCount: 1,
      cellCount: 256,
      formulaCellCount: 128,
      valueCellCount: 256,
      mergeCount: 1,
      conditionalFormatCount: 1,
    })
    expect(imported?.stats.dimensions[0]).toMatchObject({
      sheetName: 'Data',
      rowCount: 128,
      columnCount: 2,
      nonEmptyCellCount: 256,
      usedRange: { startRow: 0, startColumn: 0, endRow: 127, endColumn: 1 },
    })
  })

  it('skips worksheet metadata XML retention in verifier-only headless mode', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B2"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>1</v></c></row></sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '<conditionalFormatting sqref="B1:B2"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should not be inflated for headless verifier import')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'headless-verifier.xlsx', zip, {
      materializeCells: false,
      materializeMetadata: false,
      minByteLength: 0,
    })

    expect(imported?.snapshot.sheets[0]?.metadata).toBeUndefined()
    expect(imported?.stats.tableCount).toBe(1)
    expect(imported?.stats.mergeCount).toBe(1)
    expect(imported?.stats.conditionalFormatCount).toBe(1)
  })

  it('streams compressed worksheet zip entries across multiple compressed chunks', () => {
    const rows: string[] = []
    for (let row = 1; row <= 4_096; row += 1) {
      rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row * 17)}</v></c></row>`)
    }
    const worksheetXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      '<dimension ref="A1:A4096"/>',
      `<sheetData>${rows.join('')}</sheetData>`,
      '</worksheet>',
    ].join('')
    const bytes = buildLargeSimpleWorkbook({ includeSharedStrings: false, worksheetXml })
    const zip = readXlsxZipEntriesLazy(bytes)
    let streamedByteLength = 0

    const streamed = forEachInflatedXlsxZipEntryChunk(
      zip,
      'xl/worksheets/sheet1.xml',
      (chunk) => {
        streamedByteLength += chunk.byteLength
      },
      { chunkSize: 128 },
    )

    expect(streamed).toBe(true)
    expect(streamedByteLength).toBe(strToU8(worksheetXml).byteLength)
  })

  it('streams materialized worksheet cells without inflating the lazy worksheet entry', () => {
    const bytes = buildLargeSimpleWorkbook({
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:C2"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>1+1</f><v>2</v></c></row>',
        '<row r="2"><c r="C2" t="inlineStr"><is><t>Inline</t></is></c></row></sheetData>',
        '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('materialized worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-materialized.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha &#8211;' },
      { address: 'B1', value: 2, formula: '1+1' },
      { address: 'C2', value: 'Inline' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.merges).toEqual([{ sheetName: 'Data', startAddress: 'A1', endAddress: 'B1' }])
  })

  it('falls back before inflating worksheet XML when streamed metadata cannot be typed exactly', () => {
    const bytes = buildLargeSimpleWorkbook({
      includeSharedStrings: false,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '<dimension ref="A1:A1"/>',
        '<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>',
        '<hyperlinks><hyperlink ref="A1:A2000" r:id="rIdHyperlink1"/></hyperlinks>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('unsupported streamed metadata should fall back without inflating worksheet XML')
      },
    })

    expect(tryImportLargeSimpleXlsx(bytes, 'unsupported-hyperlink-range.xlsx', zip, { minByteLength: 0 })).toBeNull()
  })

  it('streams only referenced shared strings for materialized imports', () => {
    const richStringXml = '<si><r><rPr><b/></rPr><t>Rich</t></r><r><t xml:space="preserve"> Value</t></r></si>'
    const bytes = buildLargeSimpleWorkbook({
      sharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>Unused 0</t></si>
  <si><t>Alpha</t></si>
  <si><t>Unused 2</t></si>
  ${richStringXml}
  <si><t>Unused 4</t></si>
</sst>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>3</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'streamed-shared-strings.xlsx', zip, { minByteLength: 0 })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Rich Value' },
    ])
    expect(imported?.snapshot.sheets[0]?.metadata?.richTextArtifacts).toEqual({
      cells: [
        {
          address: 'B1',
          text: 'Rich Value',
          storage: 'sharedString',
          xml: richStringXml,
        },
      ],
    })
  })

  it('collects materialized shared-string references during the real worksheet stream', () => {
    const bytes = buildLargeSimpleWorkbook({
      sharedStringsXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Unused</t></si>
  <si><t>Alpha</t></si>
  <si><t>Beta</t></si>
</sst>`,
      worksheetXml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B1"/>',
        '<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>2</v></c></row></sheetData>',
        '</worksheet>',
      ].join(''),
    })
    const zip = readXlsxZipEntriesLazy(bytes)
    const worksheetStreamCount = countLazyZipEntryStreams(zip, 'xl/worksheets/sheet1.xml')
    Object.defineProperty(zip, 'xl/sharedStrings.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('sharedStrings.xml should be streamed instead of fully inflated')
      },
    })
    Object.defineProperty(zip, 'xl/worksheets/sheet1.xml', {
      configurable: true,
      enumerable: true,
      get() {
        throw new Error('worksheet XML should be streamed instead of inflated')
      },
    })

    const imported = tryImportLargeSimpleXlsx(bytes, 'single-pass-shared-strings.xlsx', zip, {
      minByteLength: 0,
      releaseZipSource: true,
    })

    expect(imported?.snapshot.sheets[0]?.cells).toEqual([
      { address: 'A1', value: 'Alpha' },
      { address: 'B1', value: 'Beta' },
    ])
    expect(worksheetStreamCount()).toBe(1)
    expect(imported?.stats.phaseTelemetry.map((entry) => entry.phase)).toEqual([
      'zip-setup',
      'shared-string-resolution',
      'worksheet-scan',
      'metadata-parsing',
      'style-parsing',
      'zip-source-release',
      'public-snapshot-materialization',
    ])
    expect(imported?.stats.phaseTelemetry.every((entry) => Number.isInteger(entry.elapsedMs) && entry.elapsedMs >= 0)).toBe(true)
    expect(imported?.stats.phaseTelemetry.every((entry) => (entry.rssBytes ?? 0) > 0)).toBe(true)
  })
})
