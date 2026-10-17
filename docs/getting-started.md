# Getting started

Use this page to convert one workbook.

## Requirements

- Node.js 20 or later. Browsers also work (see [Browsers](#browsers)).
- An `.xlsx` file. The library does not read `.xls` files.

## Install

```bash
npm install xlsx-context
```

Two packages are optional:

- `iconv-lite` writes Shift_JIS Markdown.
- `js-tiktoken` counts OpenAI tokens exactly.

## Convert a workbook

```js
import { readFile } from "node:fs/promises";
import { parseWorkbook, toAIContext } from "xlsx-context";

const bytes = await readFile("sales.xlsx");
const workbook = await parseWorkbook(bytes, { fileName: "sales.xlsx" });
const context = toAIContext(workbook, { maxTokens: 512 });

console.log(context.tokens); // token count of the whole text
console.log(context.chunks.length); // number of chunks
```

`toAIContext` returns the IR, the compact text, and the chunks. See [AI context](./ai-context.md).

## Write Markdown

```js
import { convertWorkbookToMarkdown, createCombinedMarkdown } from "xlsx-context";

const files = convertWorkbookToMarkdown(workbook); // one file for each sheet
const document = createCombinedMarkdown(workbook, files); // one document for the workbook
```

## Use the command line

```bash
npx xlsx-context sales.xlsx --out sales.md
```

See [Command line](./cli.md).

## Browsers

Import `xlsx-context` with a bundler. The browser entry uses the `DOMParser` and `DecompressionStream` of the browser.

```js
const bytes = new Uint8Array(await file.arrayBuffer());
const workbook = await parseWorkbook(bytes, { fileName: file.name });
```
