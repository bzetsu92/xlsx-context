# Tests

Runner: Vitest. `vitest.config.js` has one project per test folder.

| Folder                         | What it checks                                                                                                                                                  | Command                      | Needs a build |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------- | ------------- |
| `e2e/unit/`                    | One source file. The path mirrors `src/`: `e2e/unit/structure/tables/border-grid.test.js` ↔ `src/structure/tables/border-grid.ts`.                              | `npm run test:unit`          | no            |
| `e2e/workbooks/` and `e2e/ai/` | Opens real `.xlsx` files from `fixtures/` and checks the output: parsed data, Markdown, AI context.                                                             | `npm run test:workbooks`     | no            |
| `cli/`                         | Runs the `bin/` command line tool with its flags.                                                                                                               | `npm run test:cli`           | yes           |
| `types/`                       | Compiles `consumer.ts` against the published declarations, as a package user does. Not a Vitest test.                                                           | `npm run typecheck:consumer` | yes           |
| `performance/`                 | Measures time, CPU, and heap on a 20,000-row sheet against a budget. Prints a table and writes `bench-results/performance-report.json`. Not part of `npm test`. | `npm run test:perf`          | no            |
| `fixtures/`                    | Test data only: `.xlsx` files. Inventory: `fixtures/README.md`.                                                                                                 | -                            | -             |
| `helpers/`                     | Code that tests share: `paths.js` (`fixture()`, `BIN`), `setup.js` (Node globals, loaded by `vitest.config.js`), `mode-assertions.js`.                          | -                            | -             |

`npm test` builds the library, then runs `e2e/unit`, `e2e/workbooks` with `e2e/ai`, and `cli` in that order.

## Where a new test goes

- It calls functions of one source file with small hand-made input, and opens no `.xlsx` file → `e2e/unit/<same path as src>.test.js`.
    - A large source file can have more files: `<source-file>.<aspect>.test.js` (`office-drawing.connectors.test.js`).
- It opens a `.xlsx` file from `fixtures/` and checks the result → `e2e/workbooks/`.
    - One file per `fixtures/` folder (`chart`, `table`, `shape`, `formula`, `image`, `merge`, `edge`, `narrative`, `named-range`, `rich`, `hyperlink`), plus `core` (hand-built workbooks), `fast-path` (fast XML reader gives the same result as the DOM path on every fixture), and shared `stored-zip.js`.
    - `e2e/ai/`: IR, compact text, and chunks built from fixtures, one file per function (`compact`, `workbook-ir`, `compact-renderer`, `chunk-planner`) plus shared `helpers.js`. Gold Q&A lives in `QUESTIONS` inside `compact.test.js`.
    - `fixture-hygiene.test.js`: fixture files hold no private metadata.
- It runs the CLI → `cli/`. It measures time → `performance/`.
- `e2e/` tests run in the `node` environment with `@xmldom/xmldom`, like Node users. Do not add `@vitest-environment jsdom`. Do not import `helpers/setup.js` in a test.

## Extra output

```bash
QA_REPORT=1 npx vitest run --project workbooks e2e/ai/compact.test  # print the answer coverage table
```
