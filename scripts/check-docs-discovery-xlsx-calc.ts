import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { requireIncludes } from './check-docs-discovery-core.ts'

export async function requireXlsxCalcAlternativeDiscovery(docsRoot: string): Promise<void> {
  const content = await readFile(join(docsRoot, 'xlsx-calc-alternative-node-workbook-recalculation.md'), 'utf8')

  for (const required of [
    'title: xlsx-calc alternative for Node workbook recalculation',
    'canonical_url: https://proompteng.github.io/bilig/xlsx-calc-alternative-node-workbook-recalculation.html',
    'cd bilig/examples/xlsx-recalculation-node',
    '"exportedReimportMatchesAfter": true',
    '"formulasSurvivedXlsxRoundTrip": true',
    'npx --package @bilig/xlsx-formula-recalc xlsx-recalc',
    'Historical `xlsx-calc` timing comparisons are intentionally excluded from the',
    'The product gate here is deterministic workbook',
    'research artifact.',
    'https://github.com/fabiooshiro/xlsx-calc',
    'https://docs.sheetjs.com/docs/csf/features/formulae/',
    'Repository and release notes',
  ] as const) {
    requireIncludes(content, required, 'docs/xlsx-calc-alternative-node-workbook-recalculation.md')
  }
}
