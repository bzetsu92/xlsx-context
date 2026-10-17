# xlsx-context

**Turn real-world Excel workbooks into context that a language model can read.**

`xlsx-context` reads `.xlsx` files the way a person reads them. It finds every table on a sheet, keeps the text around them, shows numbers and dates as Excel shows them, and remembers which chart, image, or shape sits where. You get clean Markdown for people and compact, token-aware text and chunks for LLMs and RAG.

No LLM calls. No Python. One small runtime dependency. Node.js 20+ and modern browsers.

## Spreadsheets are not CSV

A CSV export or a raw cell dump loses most of what makes a business workbook readable:

| In the workbook                                      | Raw cell dump           | `xlsx-context`                                              |
| ---------------------------------------------------- | ----------------------- | ----------------------------------------------------------- |
| Three tables and a title on one sheet                | One stream of cells     | Three tables in reading order, with the title as text       |
| `13568` with format `#,##0`, `0.25` with format `0%` | `13568`, `0.25`         | `13,568`, `25%`                                             |
| Date serial `46094`                                  | `46094`                 | `2026/3/13`, `3/13/2026`, or `13/3/2026`, as the file says  |
| `=SUM(B3:B4)`                                        | Formula text or nothing | The cached result, with the formula kept beside it          |
| Merged header cells, "graph paper" layouts           | Empty cells and noise   | One table with merge markers                                |
| A bar chart over the table                           | Lost                    | Title, type, series ranges, and the table it reads          |
| A flowchart drawn with shapes and connectors         | Lost                    | Shape text and `sh001 -> sh002` links, plus a Mermaid graph |
| Hidden sheets and rows                               | Mixed in silently       | Kept and marked, or dropped on request                      |

On the gold questions in this repository, compact output keeps **16 of 16** answers. A naive cell dump keeps 10.

## See it

Input: `tests/fixtures/chart/chart-mixed-sample01.xlsx`, a sheet with a title, a five-row table, and a combined bar and line chart.

```js
import { readFile } from "node:fs/promises";
import { parseWorkbook, toAIContext } from "xlsx-context";

const workbook = await parseWorkbook(await readFile("chart-mixed-sample01.xlsx"), {
    fileName: "chart-mixed-sample01.xlsx",
});
const { text, chunks, tokens } = toAIContext(workbook, { maxTokens: 512 });
```

`text`, ready for a prompt:

```text
# chart-mixed

複合グラフサンプル

### B3-E8
| 項目 | 売上 | 割引額 | 利益率 |
| 2024年 | 13,568 | 1,000 | 25% |
| 2025年 | 19,287 | 1,200 | 27% |
| 2026年 | 25,051 | 1,500 | 32% |
| 2027年 | 28,053 | 2,300 | 25% |
| 2028年 | 31,027 | 2,500 | 18% |

[B10] Chart c001: 棒と折れ線 (Bar Chart + Line Chart (Combined))
- 売上: categories 'chart-mixed'!$B$4:$B$8; values 'chart-mixed'!$C$4:$C$8; table t001
- 割引額: categories 'chart-mixed'!$B$4:$B$8; values 'chart-mixed'!$D$4:$D$8; table t001
- 利益率 (secondary axis): categories 'chart-mixed'!$B$4:$B$8; values 'chart-mixed'!$E$4:$E$8; table t001
```

Each chunk stands alone and points back to its source, so a retrieval step can cite the exact range:

```json
{
    "id": "s1/t001/1",
    "kind": "table",
    "source": { "sheet": "chart-mixed", "range": "B4-E8", "anchor": "" },
    "text": "# chart-mixed\n## Table t001 (B3-E8)\n| 項目 | 売上 | 割引額 | 利益率 |\n| 2024年 | 13,568 | 1,000 | 25% |\n...",
    "tokens": 53
}
```

## What you get

**Structure**

- Several tables on one sheet, side by side or stacked, found by borders, density, and header rows.
- Declared Excel tables (ListObjects) used as they are, with their names.
- Merged cells, "graph paper" forms, and captions kept next to their tables.
- Text outside tables kept as narrative, with indented lists turned into bullets.

**Values**

- Display values from the number format of the file: currency, percent, fractions, scientific, conditions, and custom codes.
- Locale from the file itself: date order, month names, and currency symbols for English, Vietnamese, French, German, Japanese, Chinese, Korean, and any other locale tag. When the file does not say, dates fall back to ISO instead of a guess.
- The 1900 and 1904 date systems.
- Cached formula results first. A built-in evaluator with 60+ common functions fills in when a cache is missing.
- Rich text (bold, italic, strike, underline), hyperlinks, comments, notes, and threaded comments.

**Visuals**

- Charts with title, type, series ranges, secondary axis, and the source table.
- Images extracted with their anchor cell.
- Shapes with their text, connectors resolved to the shapes they join, and flowcharts drawn as Mermaid.

**Ready for AI**

- A typed intermediate representation (IR) with raw and display values for every cell.
- Compact text in `rows` or `records` format.
- Chunks that never split a row, each with a source range and a token count.
- Exact token budgets through `js-tiktoken` or your own tokenizer.

**Ready for people**

- Markdown per sheet or one combined document with YAML front matter.
- ZIP export with images and shape SVGs.
- UTF-8, UTF-16, UTF-32, and Shift_JIS output.

## Install

```bash
npm install xlsx-context
```

Run the CLI without an install:

```bash
npx -y xlsx-context@latest book.xlsx --out book.md
```

Optional peers: `js-tiktoken` for exact OpenAI token counts, `iconv-lite` for Shift_JIS Markdown.

## Use it

**Library**

```js
import {
    parseWorkbook,
    convertWorkbookToMarkdown,
    buildWorkbookIR,
    renderCompactMarkdown,
    createContextChunks,
    createTiktokenCounter,
} from "xlsx-context";

const workbook = await parseWorkbook(bytes, { fileName: "book.xlsx" });

const markdownFiles = convertWorkbookToMarkdown(workbook); // one Markdown file per sheet
const ir = buildWorkbookIR(workbook); // typed structure
const prompt = renderCompactMarkdown(ir); // compact text
const counter = await createTiktokenCounter("o200k_base"); // exact budgets
const chunks = createContextChunks(ir, { maxTokens: 512, counter }); // RAG chunks
```

**Command line**

```bash
npx xlsx-context book.xlsx --out book.md --summary
npx xlsx-context book.xlsx --zip book.zip --hidden exclude
```

**AI agents**

`skills/xlsx-context/SKILL.md` ships in the package. It tells a coding agent to convert a workbook with the CLI and read the Markdown, instead of parsing `.xlsx` bytes by hand.

Copy `skills/xlsx-context/` into the skills folder of your agent, for example `~/.claude/skills/` for Claude Code. For agents that read `AGENTS.md` (Codex, Copilot, OpenCode), add this line to it:

```text
To read an .xlsx file, run `npx -y xlsx-context@latest <file> --out <file>.md --summary` and read the Markdown. Never parse .xlsx bytes by hand.
```

## Documentation

| Page                                         | Read it to                                           |
| -------------------------------------------- | ---------------------------------------------------- |
| [Getting started](./docs/getting-started.md) | Install the package and convert a workbook.          |
| [AI context](./docs/ai-context.md)           | Learn the IR, the text formats, chunks, and budgets. |
| [API reference](./docs/api.md)               | Find every public function, option, and type.        |
| [Command line](./docs/cli.md)                | Convert a workbook from a terminal.                  |
| [Architecture](./docs/architecture.md)       | Learn the layers and the data flow.                  |
| [Development](./docs/development.md)         | Build, test, and contribute.                         |

## Scope

`xlsx-context` reads `.xlsx` files and writes text. It does not:

- read `.xls` files or write `.xlsx` files,
- reproduce the visual layout of Excel,
- cut text to fit a prompt. You select the chunks.

## Quality

- Strict TypeScript. The published declarations are compiled as a consumer in CI.
- Tests check exact outputs on 33 real Excel files, one purpose per file, plus hand-built edge cases.
- A 20,000-row sheet runs against a time and memory budget on every CI run.

## Contributing

Start with [docs/development.md](./docs/development.md). It lists the commands, the checks that CI runs, and the rules for fixtures and code.

## License

Apache License 2.0. See [LICENSE](./LICENSE).
