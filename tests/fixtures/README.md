# Fixtures (test dataset)

Data only: `.xlsx` files. No test code. Tests load these files through
`helpers/paths.js` (`fixture("table", "table-basic-sample01.xlsx")`).
Rules live in `docs/development.md` ("Fixtures") and `tests/README.md`.
`e2e/workbooks/fixture-hygiene.test.js` enforces the metadata rules.

## Naming

- One file has one purpose. Use one sheet when possible.
- End the file name with `sample01`, `sample02`, and so on (no dash before the number).
- Every fixture lives in a topic folder. The root holds no `.xlsx` files.
- History notes: `table-basic-sample04` to `table-basic-sample10` were never issued.
  Continue new table fixtures at `table-basic-sample17`.
  `image-basic-sample02` also holds a chart (coexistence case). The name is history.

## Inventory (33 files)

| File | Sheet(s) | What it proves |
| ---- | -------- | -------------- |
| `table/table-basic-sample01.xlsx` | `table-basic` | Two tables adjacent vertically |
| `table/table-basic-sample02.xlsx` | `table-basic` | Two tables adjacent horizontally |
| `table/table-basic-sample03.xlsx` | `table-basic` | Four tables in a 2x2 grid |
| `table/table-basic-sample11.xlsx` | `table-basic` | Graph-paper style, one table (`B3-T7`) |
| `table/table-basic-sample12.xlsx` | `table-basic` | Graph-paper style, two tables, heavy merges |
| `table/table-basic-sample13.xlsx` | `table-basic` | Graph-paper style, four wide tables (`V-AN`) |
| `table/table-basic-sample14.xlsx` | `table-basic` | Tolerates missing merge cells |
| `table/table-basic-sample15.xlsx` | `table-basic` | Vertical merge inside remarks column |
| `table/table-basic-sample16.xlsx` | `table-basic` | Multi-char merge leak (`Ito`) |
| `table/table-border-priority-sample01.xlsx` | `border-priority` | `balanced` vs `border` detection modes differ |
| `table/table-grid-layout-sample01.xlsx` | `grid-layout` | Graph-paper, 2 tables, 70 merges, 10 formula diagnostics |
| `formula/formula-basic-sample01.xlsx` | `formula` | 9 formulas: ref, arith, IF, SUM, COUNTIF, TEXT, DATE, VALUE x2 |
| `formula/formula-crosssheet-sample01.xlsx` | `Sheet1`, `Sheet2`, `日本語シート` | Cross-sheet and Japanese-sheet references |
| `formula/formula-shared-sample01.xlsx` | `formula` | Shared formulas (`1`-`10`) |
| `formula/formula-spill-sample01.xlsx` | `spill-sample` | Spill (dynamic array) behavior |
| `merge/merge-pattern-sample01.xlsx` | `merge` | Horizontal, vertical, and 2x2 merges (12 ranges, 3 tables) |
| `merge/merge-multiline-sample01.xlsx` | `merge-multiline` | Newline inside a merged range (`B3:C4`) |
| `chart/chart-basic-sample01.xlsx` | `chart-basic` | Bar chart bound to table `B3-D7` |
| `chart/chart-mixed-sample01.xlsx` | `chart-mixed` | Combined bar+line chart, secondary axis |
| `image/image-basic-sample01.xlsx` | `image` | Table plus 2 extracted images |
| `image/image-basic-sample02.xlsx` | `image` | Image and chart on one sheet |
| `hyperlink/hyperlink-basic-sample01.xlsx` | `Summary`, `Other` | External and internal (`Other!A1`) links |
| `named-range/named-range-sample01.xlsx` | `Summary`, `Other` | Defined names incl. local cross-sheet ref |
| `narrative/narrative-vs-table-sample01.xlsx` | `narrative-vs-table` | Narrative text vs bordered table (`B8-F11`) |
| `rich/rich-text-gfm-sample01.xlsx` | `rich-github` | Bold/italic/strike/underline runs, gfm mode |
| `rich/rich-markdown-escape-sample01.xlsx` | `rich+escape` | Markdown escaping of `* _ ~ # - [ ] ( )` `` ` `` `|` `\` |
| `rich/rich-usecase-sample01.xlsx` | `rich+usecase` | Links plus styling plus `<br>` in one table |
| `shape/shape-basic-sample01.xlsx` | `shape-basic` | Text box, arrow connector, rectangle |
| `shape/shape-block-arrow-sample01.xlsx` | `shape-block-arrow` | 5 block-arrow shapes |
| `shape/shape-callout-sample01.xlsx` | `shape-callout` | 4 callout shapes |
| `shape/shape-flowchart-sample01.xlsx` | `shape-flowchart` | Flowchart shapes, connectors, mermaid output |
| `edge/edge-empty-sample01.xlsx` | `edge-empty` | Near-empty sheet: 2 cells, 0 tables |
| `edge/edge-weird-sheetname-sample01.xlsx` | `A B-東京&大阪.01` | Spaces, `&`, `.`, digits, Japanese in sheet name |

## Add a fixture by hand (checklist)

1. Keep one purpose and one sheet. Put a heading cell near the top that says what the sample tests.
2. Use few values. Save the file in Excel as usual, to keep the cached values.
3. Name it `<topic>-<what>-sampleNN.xlsx` inside `fixtures/<topic>/`.
4. Strip the metadata that Excel adds:
    - local path (`x15ac:absPath`) in `xl/workbook.xml`
    - author, last editor, dates, application in `docProps/`
    - comments, threaded comments, persons, custom properties, external links,
      connections, embeddings, VBA, printer settings
5. Check the file by hand before you push:
    ```bash
    unzip -l <file.xlsx>
    unzip -p <file.xlsx> docProps/core.xml
    unzip -p <file.xlsx> docProps/app.xml
    npx -y xlsx-context@latest ./<file.xlsx> --out /tmp/check.md --summary
    ```
6. Reference the file from a test through `fixture("<topic>", "<file>")`.
   Add one row to the inventory table above.
