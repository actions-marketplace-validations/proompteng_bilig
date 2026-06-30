# Headless WorkPaper Benchmark Evidence

Status: public evidence note for `@bilig/headless`

This note keeps the public performance language auditable from a checked-in
benchmark artifact instead of README copy alone.

## Current Source Of Truth

The public benchmark evidence is
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.

It records WorkPaper `100/100` mean-latency wins on scorecard-eligible
comparable workloads against HyperFormula `3.2.0`. The same artifact has
`100/100` workloads winning both mean and p95 latency.

## Artifact Inventory

- Primary workbook-wide artifact:
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`.
- Additional competitor artifacts remain checked in for local performance
  investigation, but they are not rolled up into a public status or broad
  claim:
  - `packages/benchmarks/baselines/workpaper-vs-univer.json`
  - `packages/benchmarks/baselines/workpaper-vs-ironcalc-rust.json`
  - `packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json`
  - `packages/benchmarks/baselines/workpaper-vs-truecalc.json`
- Public generated evidence:
  `docs/public-evidence.json`.

Current checked-in metadata:

- benchmark sampling: `200` measured samples after `2` warmup samples
- comparison engine: HyperFormula `3.2.0`, local checkout commit
  `9a510a2acb97c3d3490f9e3b9e961a1c4a98b9ad`

## Workbook-Wide Lane

The current checked-in WorkPaper-vs-HyperFormula artifact records WorkPaper
`100/100` mean-latency wins:

| Lane    | Comparable Workloads | WorkPaper Mean Wins | HyperFormula Mean Wins |
| ------- | -------------------: | ------------------: | ---------------------: |
| Overall |                `100` |               `100` |                    `0` |
| Public  |                 `73` |                `73` |                    `0` |
| Holdout |                 `27` |                `27` |                    `0` |

The artifact was generated at `2026-05-23T17:51:04.599Z`.

The overall directional mean-ratio geomean is `0.2586071973976171`. The overall
directional p95-ratio geomean is `0.2806672128213908`. Ratios below `1.0` mean
WorkPaper is faster for that metric.

The current worst mean row is `sheet-rename-dependencies`, with a mean ratio of
`0.8056914279903578`. The current worst p95 row is
`sheet-rename-dependencies`, with a p95 ratio of `0.7917355369127405`.

## What Is Measured

Scorecard-eligible families cover:

- workbook build and rebuild paths
- runtime restore from snapshot
- sheet lifecycle and named-expression changes
- cross-sheet scalar and aggregate recalculation
- dirty execution after single edits, chains, fanout, mixed frontiers, and
  formula edits
- batch edits, suspended batches, and undo-including batches
- structural row and column inserts, deletes, and moves
- dense and sparse range reads
- 2D, overlapping, sliding-window, and conditional aggregation
- exact lookup, INDEX/MATCH, INDEX reference, approximate lookup, after-write
  lookup, and text lookup

The artifact excludes the `config-toggle` control family and `dynamic-array`
family from the directly comparable win count.

## What Is Not Claimed

This is not a blanket "fastest at every spreadsheet task" claim.

It does not prove that browser-grid rendering, import/export fidelity,
collaborative sync, or every possible user workbook is faster. This benchmark is
about the headless WorkPaper runtime path against one deterministic
HyperFormula comparison artifact.

## How To Verify

Check that the committed artifact still has the expected workload coverage and
shape:

```bash
pnpm workpaper:bench:competitive:check
pnpm public:evidence:check
```

Regenerate timing evidence only when intentionally refreshing benchmark
artifacts:

```bash
pnpm workpaper:bench:competitive:generate
pnpm public:evidence:generate
```
