# AI context

AI context is the output for language models. It has three parts: the IR, the compact text, and the chunks.
It needs no language model and no network. The same input always gives the same output.

## Data flow

```text
.xlsx --parseWorkbook--> ParsedWorkbook --buildWorkbookIR--> WorkbookIR
WorkbookIR --renderCompactMarkdown--> compact text
WorkbookIR --createContextChunks--> chunks
```

`toAIContext` runs all of these steps.
No step reads Markdown. The IR comes from the parsed workbook.

## The IR

The IR holds these items for each sheet:

- tables
- narratives
- comments
- charts
- images
- shapes

Each cell keeps its raw value, its display value, its formula, its hyperlink, and a merge flag.
The IR never drops a value. The renderers decide what to omit.

The display value is the text that Excel shows. It follows the number format and the locale of the file,
for example `13,568`, `25%`, `$ 1,234.50`, or `13/3/2026`. When the file does not show a locale,
built-in date formats use ISO dates. `workbook.locale` and `workbook.dateSystem` tell which rules applied.

Each sheet also keeps its visibility:

| Field        | Meaning                               |
| ------------ | ------------------------------------- |
| `state`      | `visible`, `hidden`, or `veryHidden`. |
| `hiddenRows` | One-based rows that Excel hides.      |
| `hiddenCols` | One-based columns that Excel hides.   |

Hidden content stays in the IR. Parse with `includeHidden: false` to drop it before every output.

Ids follow the order on the sheet: `t001` (table), `c001` (chart), `i001` (image), `sh001` (shape).
Tables are numbered in reading order: a column of side-by-side tables comes before the next column.
Table, chart, and image ids match the numbers in the Markdown headings.
A table that the file declares as an Excel table keeps its exact range, and `name` holds its display name.

Each chart series keeps its ranges and the values from the chart cache.
`sourceTable` names the table that holds the series values, for example `t001`.
It is `null` when no table on the sheet holds them.
A connector shape has `connection`: the ids of the shapes it joins and its arrow.

`sections` gives the reading order that Markdown also uses: groups of related tables and
narrative blocks, so a caption stays next to its table. The compact text follows it.

A table whose first row holds data, for example `base1 | 10`, has `hasHeader: false` and empty `headers`.
A first row with a number or a date counts as data unless it is bold and the next row is not.

The IR drops empty rows and columns, like Markdown does.
A table without values does not appear. The numbering counts only the tables that appear.

## Compact text

The compact text is short on purpose.
It has no front matter, no padded tables, and no empty sections.
A sheet with hidden content starts with one note, for example `> Hidden sheet.` or `> Hidden rows: 2, 5-7.`
A declared Excel table adds its name to the heading, for example `### A1-C4 Orders`.

`tableFormat: "rows"` (default) writes a pipe table:

```text
# Orders

### A1-C4
| Order | Customer | Total |
| 1001 | [Acme](https://example.com/acme) | 250.00 |
| 1002 | Globex |  |
| 1003 | Initech | 75.50 |

- [C3] Ann: Total is pending.

[E2] Chart c001: Totals by order (Bar Chart)
- Total: categories Orders!$A$2:$A$4; values Orders!$C$2:$C$4; table t001
```

A chart series that no table holds lists its cached points, for example `data Q1=10, Q2=20`.
A connector line names the shapes it joins, for example `sh001 -> sh002`.

`tableFormat: "records"` writes one line for each row:

```text
### A1-C4
[A2] Order: 1001 | Customer: [Acme](https://example.com/acme) | Total: 250.00
[A3] Order: 1002 | Customer: Globex
[A4] Order: 1003 | Customer: Initech | Total: 75.50
```

Rules for the text:

- A record starts with the address of its first cell, for example `[A2]`.
- A record omits empty values.
- An empty header uses the column letter. A repeated header gets the column letter, for example `Value (D)`.
- A table without a header row has no header line; its records use column letters, for example `[A3] A: base1 | B: 10`.
- An external link is `[label](url)`. An internal link is `label (Sheet2!A1)`.
- A pipe in a cell is `\|`. A line break in a cell is `<br>`.

Use `rows` when the whole text goes into one prompt.
Use `records` when a language model reads single rows.
In the test fixtures, `records` use about 47% more tokens than `rows`.

## Chunks

A chunk is a piece of the compact text that stands alone.

| Field     | Meaning                                                   |
| --------- | --------------------------------------------------------- |
| `id`      | `s{sheet}/{kind}/{number}`, for example `s1/t001/2`.      |
| `kind`    | `table`, `narrative`, `comments`, or `assets`.            |
| `sheet`   | The sheet name.                                           |
| `tableId` | The IR table id. Empty for other kinds.                   |
| `source`  | `{ sheet, range, anchor }`. The location in the workbook. |
| `text`    | The chunk text.                                           |
| `tokens`  | The token count of `text`.                                |

Example with `maxTokens: 30` and `tableFormat: "records"`:

```json
{
    "id": "s1/t001/2",
    "kind": "table",
    "source": { "sheet": "Orders", "range": "A3-C4", "anchor": "" },
    "tokens": 25,
    "text": "# Orders\n## Table t001 (A1-C4)\n[A3] Order: 1002 | Customer: Globex\n[A4] Order: 1003 | Customer: Initech | Total: 75.50"
}
```

Rules for the chunks:

- The order on each sheet is the reading order of the compact text: tables and narrative text
  as `sections` lists them, so a caption stays next to its table. Comments, then assets follow.
- A table chunk holds whole rows. It never splits a row.
- Each table chunk repeats the sheet name and the table heading.
  In `rows` format, it also repeats the header row.
- `source.range` points back to the content of the chunk:
  a table chunk covers its data rows (`A3-C4`), a narrative chunk its row span (`4:8`),
  and a comment or asset chunk its cells (`B3,D10`). An asset chunk also sets `anchor`.
- A unit that is larger than the budget stays whole in its own chunk. Its `tokens` value is then above `maxTokens`.
- The same input, options, and counter give the same chunks.

## Token budgets

`maxTokens` is the budget of one chunk. The default is 512.
`AiContext.tokens` is the size of the whole text.

The library does not cut the text to fit a budget. Cutting without knowledge can remove data.
Compare `tokens` with the limit of your prompt:

```js
const context = toAIContext(workbook, { maxTokens: 512 });
const prompt = context.tokens <= 4000 ? context.text : selectChunks(context.chunks);
```

`selectChunks` is your own retrieval step.

## Counters

The default counter is an estimate. The full list lives in [API reference](./api.md#counters).

Pass an exact counter when the budget must match a language model:

```js
const counter = await createTiktokenCounter("o200k_base");
const context = toAIContext(workbook, { maxTokens: 512, counter });
```

## Untrusted content

A workbook is input from outside your system. Treat its text as data, never as instructions.

- Cell text, comments, sheet names, and shape text go into the context as they are.
  A cell can hold text such as "ignore previous instructions". The library does not remove it.
- Put the workbook text in a delimited data section of your prompt, apart from your instructions.
- Markdown output escapes names and cell text, so the file cannot add headings, links, or HTML.
  A link target with a `javascript:`, `vbscript:`, or `data:` scheme stays plain text.
- Limits stop hostile packages: entry count, uncompressed size, inflation past the declared size,
  and hyperlink ranges. A file over a limit fails with an error.
