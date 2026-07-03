import { join } from 'node:path'

import { requirePublishedSource } from './check-docs-discovery-core.ts'
import { loadDocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { requireDocsLocalHtmlLinksHaveSources } from './check-docs-discovery-local-links.ts'
import { requireSitemapPublishedSources } from './check-docs-discovery-sitemap.ts'

const { docsRoot, expectedSitemapUrls, sitemap, siteRoot, sourceFilesByUrl } = await loadDocsDiscoveryContext()

const { sourceFilesToVerify } = requireSitemapPublishedSources({
  expectedSitemapUrls,
  sitemap,
  siteRoot,
  sourceFilesByUrl,
})

await Promise.all(sourceFilesToVerify.map((sourceFile) => requirePublishedSource(join(docsRoot, sourceFile))))
requireDocsLocalHtmlLinksHaveSources(docsRoot)

console.log(
  JSON.stringify(
    {
      checked: 'docs-integrity',
      publishedSourceCount: sourceFilesToVerify.length,
    },
    null,
    2,
  ),
)
