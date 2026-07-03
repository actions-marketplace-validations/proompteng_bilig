---
title: HyperFormula alternative for Node WorkPaper automation
published: true
description: Compare Bilig WorkPaper with HyperFormula for backend formula workbooks, persistence, mutation receipts, and agent readback.
tags: hyperformula, spreadsheet, node, formulas, comparison
canonical_url: https://proompteng.github.io/bilig/hyperformula-alternative-headless-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# HyperFormula Alternative For Node WorkPaper Automation

Status: public comparison guide for `@bilig/headless`.

This page is for developers already evaluating HyperFormula or another
headless spreadsheet engine. It is intentionally not a takedown. HyperFormula
is the established TypeScript engine in this category; its official README
positions it as a headless spreadsheet for business web apps with formula
evaluation, CRUD operations, undo/redo, clipboard support, sorting, Node.js
support, and a GPLv3 or commercial license.

`@bilig/headless` is worth evaluating when the workload is closer to a
service-side WorkPaper runtime: formula-backed business logic, structural edits,
agent writeback, persistence, restore, and auditable compatibility evidence from the
same repository.

## Short Version

Use HyperFormula when you want a mature, UI-independent formula calculation
engine with broad spreadsheet-function coverage and commercial support from the
Handsontable team.

Use `@bilig/headless` when you need a MIT-licensed Node package that treats the
workbook as a service object: build it, mutate it, read formulas and values,
persist it, restore it, and verify agent-style edits without opening a browser
grid.

For the broader engine choice, start with the
[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser).

## Comparison Surface

| Question                     | HyperFormula                                                                         | `@bilig/headless`                                                                                                    |
| ---------------------------- | ------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------- |
| Primary shape                | Headless spreadsheet formula engine                                                  | Headless WorkPaper workbook facade                                                                                   |
| Runtime target               | Browser or Node.js                                                                   | Node services, tests, agents, and local runtime automation                                                           |
| License posture              | GPLv3 or commercial license                                                          | MIT                                                                                                                  |
| API orientation              | Spreadsheet-engine instance with formula evaluation and workbook operations          | WorkPaper object with formula evaluation, structural edits, persistence helpers, history, and readback               |
| Agent workflow fit           | Possible, but the project is not specifically packaged around agent writeback proofs | First-class evaluation path includes an agent writeback demo with persistence and restored readback                  |
| Public proof in this repo    | External comparison target                                                           | Deterministic WorkPaper evaluator, formula, import/export, and Excel oracle gates                                   |
| Caveat                       | Strong default engine, with its own licensing and integration model                  | Formula inventory is complete for the tracked Office surface, but arbitrary Excel edge cases remain evidence-scoped  |

## Try The Package

Use the published package when you want a quick local evaluation:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Use the maintained example when you want an end-to-end proof:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run start
pnpm --dir examples/headless-workpaper run agent:verify
```

The agent verifier records the assumption cells changed, checks dependent
formula readback, persists the workbook, restores it, and verifies the restored
values.

## When To Choose bilig First

- You need a workbook object in a Node service, not a browser grid.
- You need formulas plus structural edits, undo/redo, persistence, and restore.
- You are building a coding-agent or workflow-agent loop that must verify
  writes by reading formulas and values back from the same workbook model.
- You want an MIT-licensed package surface.
- You want performance and compatibility claims tied to deterministic fixtures
  and local commands.

## When Not To Choose bilig First

- You need a mature commercial support channel today.
- You need a mature engine with a longer compatibility history before adding
  reduced fixtures for the exact Excel edge cases your workflow uses.
- You need a library already centered around a visual spreadsheet component.
- You need every XLSX feature preserved across import/export right now.

For those cases, start with HyperFormula or a full spreadsheet product, then
use bilig's compatibility notes to decide whether a narrower WorkPaper runtime
fits a later slice.

## Proof Links

- Package README:
  [`packages/headless/README.md`](../packages/headless/README.md)
- Compatibility boundaries:
  [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- Starter issues:
  [`docs/starter-issues.md`](starter-issues.md)
- Official HyperFormula repository:
  <https://github.com/handsontable/hyperformula>
- Official HyperFormula site:
  <https://hyperformula.handsontable.com/>
