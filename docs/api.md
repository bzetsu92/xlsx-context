# API reference

Import every name from `xlsx-context`. Only the names on this page are public.

## Parse

### `parseWorkbook(input, options?)`

Reads an `.xlsx` file. Returns `Promise<ParsedWorkbook>`.

| Argument                      | Type                        | Default         | Meaning                                                         |
| ----------------------------- | --------------------------- | --------------- | --------------------------------------------------------------- |
| `input`                       | `ArrayBuffer \| Uint8Array` | none            | The bytes of the file. The function accepts a Node.js `Buffer`. |
| `options.fileName`            | `string`                    | `workbook.xlsx` | The name that the outputs show.                                 |
| `options.includeShapeDetails` | `boolean`                   | `true`          | When `false`, the parser skips drawing shapes.                  |
| `options.includeHidden`       | `boolean`                   | `true`          | When `false`, the parser drops hidden sheets, rows, columns.    |

Display values follow the file. `workbook.locale` is the locale that the file shows: locale tags in number formats, or an East Asian default font. It is `null` when the file does not say. Built-in date formats then use ISO dates, and built-in currency formats show no symbol. `workbook.dateSystem` is `"1904"` when the workbook counts dates from 1904-01-01. Hidden sheets, rows, and columns stay in the output and carry `state`, `hiddenRows`, and `hiddenCols`. Declared Excel tables keep their exact range and name.

## Markdown

### `convertWorkbookToMarkdown(workbook, options?)`

Returns `MarkdownFile[]`. The array has one file for each sheet, in sheet order.

### `convertSheetToMarkdown(workbook, sheet, options?)`

Returns one `MarkdownFile` for one sheet.

### `MarkdownOptions`

| Option                  | Values                                | Default    |
| ----------------------- | ------------------------------------- | ---------- |
| `treatFirstRowAsHeader` | `boolean`                             | `true`     |
| `trimText`              | `boolean`                             | `true`     |
| `removeEmptyRows`       | `boolean`                             | `true`     |
| `removeEmptyColumns`    | `boolean`                             | `true`     |
| `includeShapeDetails`   | `boolean`                             | `true`     |
| `outputMode`            | `display`, `raw`, `both`              | `display`  |
| `formattingMode`        | `plain`, `gfm`                        | `plain`    |
| `tableDetectionMode`    | `balanced`, `border`, `planner-aware` | `balanced` |

The command line uses `gfm` (GitHub Flavored Markdown) as the default `formattingMode`. The old name `github` is a deprecated alias of `gfm`.

Charts, images, and shape blocks appear at their anchors, next to the tables and text around them.
A chart names the tables that hold its series (`- Source:`).
When no table on the sheet holds every series, the chart shows its cached data as a table.
A shape block lists each shape with its text. Shapes that connectors join also appear as a Mermaid flowchart.
`includeShapeDetails: false` keeps the list and the flowchart. It drops the raw drawing entries.

### `MarkdownFile`

| Field       | Meaning                                                                   |
| ----------- | ------------------------------------------------------------------------- |
| `fileName`  | The suggested file name.                                                  |
| `sheetName` | The sheet name.                                                           |
| `markdown`  | The Markdown text.                                                        |
| `summary`   | Counts (tables, images, charts, comments, cells) and formula diagnostics. |

### `createCombinedMarkdown(workbook, files, options?)`

Joins the files into one document. Returns `{ fileName, content, data, mimeType }`.
`data` holds the encoded bytes of `content`.

### `createMarkdownArchive(workbook, files, options?)`

Returns a ZIP archive as `Uint8Array`. The archive holds the Markdown and the extracted images.

### `MarkdownExportOptions`

Use these options with `createCombinedMarkdown` and `createMarkdownArchive`.

| Option         | Values                                                               | Default   |
| -------------- | -------------------------------------------------------------------- | --------- |
| `encoding`     | `utf-8`, `shift_jis`, `utf-16le`, `utf-16be`, `utf-32le`, `utf-32be` | `utf-8`   |
| `bom`          | `off`, `on`                                                          | `off`     |
| `frontMatter`  | `include`, `exclude`                                                 | `include` |
| `shapeDetails` | `include`, `exclude`. This value appears in the front matter only.   | `exclude` |
| `toolVersion`  | `string`. The front matter shows this value.                         | `unknown` |

`shift_jis` needs the optional package `iconv-lite`. `shift_jis` does not allow `bom: "on"`.

### Other functions

| Function                     | Returns                                        |
| ---------------------------- | ---------------------------------------------- |
| `createSummaryText(file)`    | A text summary of one `MarkdownFile`.          |
| `getMarkdownOptionChoices()` | The accepted values of the enumerated options. |

## AI context

See [AI context](./ai-context.md) for the concepts.

### `toAIContext(workbook, options?)`

Builds the IR, the compact text, and the chunks in one call. Returns `AiContext`.

| Field    | Meaning                                 |
| -------- | --------------------------------------- |
| `ir`     | The `WorkbookIR`.                       |
| `text`   | The compact text of the whole workbook. |
| `tokens` | The token count of `text`.              |
| `chunks` | The `ContextChunk[]` array.             |

### Options

`toAIContext` accepts `AiIrOptions` and `ChunkOptions` together. The other functions accept the part that applies.

| Option               | Used by               | Default       | Meaning                                            |
| -------------------- | --------------------- | ------------- | -------------------------------------------------- |
| `tableDetectionMode` | `buildWorkbookIR`     | `balanced`    | See `MarkdownOptions`.                             |
| `removeEmptyRows`    | `buildWorkbookIR`     | `true`        | Drops rows without values.                         |
| `removeEmptyColumns` | `buildWorkbookIR`     | `true`        | Drops columns without values.                      |
| `tableFormat`        | render, chunks        | `rows`        | `rows` or `records`.                               |
| `includeComments`    | render, chunks        | `true`        | Adds cell comments.                                |
| `includeAssets`      | render, chunks        | `true`        | Adds charts, shapes, and images.                   |
| `maxTokens`          | chunks                | `512`         | The token budget of one chunk. Must be above zero. |
| `counter`            | chunks, `toAIContext` | char estimate | The `TokenCounter` that counts tokens.             |

### Step functions

| Function                              | Returns          | Meaning                                 |
| ------------------------------------- | ---------------- | --------------------------------------- |
| `buildWorkbookIR(workbook, options?)` | `WorkbookIR`     | Builds the IR from the parsed workbook. |
| `renderCompactMarkdown(ir, options?)` | `string`         | Renders the IR as compact text.         |
| `createContextChunks(ir, options?)`   | `ContextChunk[]` | Splits the IR into chunks.              |

### Counters

| Function                           | Returns                 | Meaning                                                                                                                                                           |
| ---------------------------------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `createCharEstimateCounter()`      | `TokenCounter`          | The default. One token for each CJK character, symbol, and line break, one for four other letters, one for three digits. Reads about 1.1x the `o200k_base` count. |
| `createWhitespaceCounter()`        | `TokenCounter`          | Counts words. Do not use it for Japanese or Chinese text.                                                                                                         |
| `createEncoderCounter(encoder)`    | `TokenCounter`          | Wraps `{ name, encode(text) }`. The count is the length of the result.                                                                                            |
| `createTiktokenCounter(encoding?)` | `Promise<TokenCounter>` | Counts OpenAI tokens. Default encoding: `o200k_base`. Needs `js-tiktoken`.                                                                                        |

A `TokenCounter` is `{ name: string, count(text: string): number }`.

## Types

The package exports the type of each value above. Examples: `ParsedWorkbook`, `ParsedSheet`, `WorkbookIR`,
`IrTable`, `IrCell`, `IrChart`, `ContextChunk`, `AiContext`, `TokenCounter`.
Import them with `import type`.

## Errors

| Cause                                       | Message starts with                                        |
| ------------------------------------------- | ---------------------------------------------------------- |
| The input is empty                          | `Empty workbook input. Expected .xlsx bytes.`              |
| The input is not a ZIP file                 | `ZIP end-of-central-directory record was not found.`       |
| A ZIP entry path leaves the package root    | `OPC part path escapes package root:`                      |
| A package has too many entries              | `Too many ZIP entries:`                                    |
| A package exceeds 512 MiB uncompressed      | `Uncompressed ZIP size is over 512 MiB.`                   |
| `maxTokens` is zero or negative             | `maxTokens must be a positive number:`                     |
| `js-tiktoken` is not installed              | `createTiktokenCounter needs the optional peer dependency` |
| `iconv-lite` is not installed (`shift_jis`) | `Shift_JIS encoding needs the optional peer dependency`    |
