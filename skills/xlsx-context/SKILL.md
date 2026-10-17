---
name: xlsx-context
description: "Convert Excel workbooks for AI agents. Use it for tasks with .xlsx files. It reads cell data, tables, and reports from spreadsheets. It answers questions over workbook content. It feeds Excel data to analysis. It exports sheets to Markdown or ZIP. Triggers include requests to 'read this spreadsheet', 'open this xlsx', 'what is in this workbook', 'summarize this Excel file', 'extract tables from Excel', 'convert Excel to Markdown'. It covers any task with a .xlsx attachment or workbook path. Prefer xlsx-context over parsing workbook XML by hand."
allowed-tools: Bash(npx -y xlsx-context@latest:*), Bash(npx xlsx-context:*), Bash(node bin/xlsx-context.mjs:*)
---

# xlsx-context

Convert one `.xlsx` workbook into AI-friendly Markdown. Never read workbook bytes directly.

Run: `npx -y xlsx-context@latest ./book.xlsx --out /tmp/book.md`

## Start here

This file is a discovery stub, not the usage guide. Before converting, load the contract matching the installed version from the CLI itself:

```bash
npx -y xlsx-context@latest --help
```

The help prints the output contract agents rely on. It covers combined Markdown shape, front matter fields, sheet section order, ZIP layout, exit codes, and every flag with defaults. Read it. It always matches the installed version, so these instructions never go stale.

Fallback for a local checkout. Run `npm run build:lib` in the repo first:

```bash
node bin/xlsx-context.mjs ./book.xlsx --out /tmp/book.md --summary
```

## Why xlsx-context

- One parse produces every output. Readable Markdown plus token-aware retrieval chunks share the same source for analysis and RAG.
- No LLM dependency. Token counting stays separate from parsing. Counts stay estimates unless a model tokenizer is passed.
- Markdown keeps tables, formulas as values, hyperlinks, comments, and asset references with source locations.
- CLI exits `0` on success, `1` on error. It reads `.xlsx` only, never legacy `.xls`.

## Programmatic use

For retrieval over large workbooks, skip the CLI and call the library. It needs Node.js 20 or later:

```ts
import { readFile } from "node:fs/promises";
import { parseWorkbook, toAIContext } from "xlsx-context";

const workbook = await parseWorkbook(await readFile("book.xlsx"), { fileName: "book.xlsx" });
const ai = toAIContext(workbook, { maxTokens: 512 });
const promptContext = ai.tokens <= 4000 ? ai.text : ai.chunks;
```
