# @bilig/benchmarks

Benchmark scenarios, metrics helpers, and workbook generators for bilig.

## Install

```bash
npm install @bilig/benchmarks
```

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`
- Corpus: `./dist/workbook-corpus.js`

## WorkPaper local baseline

The repo tracks a deterministic WorkPaper local baseline at
`packages/benchmarks/baselines/workpaper-baseline.json`.

Refresh or validate it with:

```bash
pnpm workpaper:bench:generate
pnpm workpaper:bench:check
```

## Other Headless Competitor Research

Scoped competitor generators remain available for local performance investigation. These artifacts
are written under `.cache/research-workpaper-benchmarks` and are not checked-in public proof status.

```bash
pnpm research:workpaper:bench:truecalc:generate
pnpm research:workpaper:bench:univer:generate
pnpm research:workpaper:bench:xlsx-calc:generate
```

The Univer generator uses Univer's documented `@univerjs/preset-sheets-node-core` Node preset and
times public Facade API edits through formula recalculation completion.

## Named giant-workbook corpus

The package now ships deterministic named workbook cases for giant-data restore and warm-start
contracts:

- `dense-mixed-100k`
- `dense-mixed-250k`
- `wide-mixed-250k`
- `analysis-multisheet-100k`
- `analysis-multisheet-250k`

Use `buildWorkbookBenchmarkCorpus(...)` to materialize an exact-size workbook snapshot and stable
viewport metadata for CI and perf harnesses.

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
