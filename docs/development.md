# Development

This page is for contributors. It tells you how to set up the repository, where code goes, and what CI checks.

## Set up

Requirements: Node.js 20 or later.

```bash
npm install
npm test        # builds dist/, then runs unit, workbook, and CLI tests
```

Try the CLI on a fixture:

```bash
npm run cli -- tests/fixtures/chart/chart-mixed-sample01.xlsx --out /tmp/chart.md --summary
```

## Repository map

```text
src/
├── public/        # Public entry. Only names exported from public/index.ts are public API.
├── core.ts        # Wires the layers together.
├── model/         # Shared types: cell, table, worksheet, workbook.
├── extraction/    # Reads the package: ZIP, XML, workbook, styles, locale, drawings.
├── structure/     # Sheet meaning: addresses, tables, layout, number and date formats.
├── formula/       # Formula syntax, references, and evaluation.
└── outputs/
    ├── markdown/  # Markdown, front matter, ZIP export, rich text.
    └── ai-context/# IR, compact text, chunks, token counters.
bin/               # CLI. It uses only the public API.
skills/            # Agent skill that ships in the package.
tests/             # See tests/README.md.
```

Dependencies go one way: `public` → `core` → `outputs`, `formula`, `structure`, `extraction`, all over `model`. Outputs never import each other. Code that both outputs need goes in `structure`. Details: [Architecture](./architecture.md).

## Where a change goes

| You want to                              | Change                                                             |
| ---------------------------------------- | ------------------------------------------------------------------ |
| Read a new part of the `.xlsx` package   | `src/extraction/`, then a field in `src/model/`                    |
| Show a stored value differently          | `src/structure/formatting/` (`number-format.ts`, `date-format.ts`) |
| Support a built-in format id or a locale | `src/extraction/workbook/workbook-locale.ts`                       |
| Change how tables are found              | `src/structure/tables/table-detector.ts`                           |
| Change Markdown                          | `src/outputs/markdown/`                                            |
| Change the IR, compact text, or chunks   | `src/outputs/ai-context/` and `docs/ai-context.md`                 |
| Add a public function, option, or type   | `src/public/`, `docs/api.md`, and `tests/types/consumer.ts`        |
| Add a CLI flag                           | `bin/lib/cli-runner.mjs`, `docs/cli.md`, and `tests/cli/`          |

A change to the public surface needs a review. Keep display values driven by the file: do not hard-code one country.

## Commands

| Command                      | Job                                                                   |
| ---------------------------- | --------------------------------------------------------------------- |
| `npm test`                   | Builds the library, then runs unit, workbook, and CLI tests.          |
| `npm run test:unit`          | Runs `tests/e2e/unit`. No build is necessary.                         |
| `npm run test:workbooks`     | Runs `tests/e2e/workbooks` and `tests/e2e/ai`. No build is necessary. |
| `npm run test:cli`           | Builds, then runs `tests/cli`.                                        |
| `npm run test:perf`          | Runs `tests/performance` and prints the measured times.               |
| `npm run typecheck`          | Checks the types of `src` with `strict` on.                           |
| `npm run typecheck:consumer` | Checks the declarations as a package user. Run `build:lib` first.     |
| `npm run format`             | Formats the files with Prettier.                                      |
| `npm run format:check`       | Fails when a file is not formatted.                                   |
| `npm run build:lib`          | Builds `dist/lib` (JavaScript) and `dist/types` (declarations).       |
| `npm run cli -- <file>`      | Runs the CLI from the repository.                                     |
| `npm run bench:baseline`     | Measures time, size, and token counts on the fixtures.                |

To run one test file: `npx vitest run --project workbooks rich.test`.

`dist/` holds generated files. Do not edit them. Git ignores them.

## Checks before a merge

CI runs these steps in order. Run the same steps before you push a change to `src/`:

1. `npm run format:check`
2. `npm run typecheck`
3. `npm test`
4. `npm run typecheck:consumer`
5. `npm run test:perf`

## Tests

`tests/README.md` explains each folder and where a new test goes. In short:

- `tests/e2e/unit/` mirrors `src/`: `unit/structure/tables/border-grid.test.js` covers `src/structure/tables/border-grid.ts`.
- `tests/e2e/workbooks/` opens real fixtures, one file per fixture folder. `core.test.js` holds hand-built workbooks for cases that need exact XML.
- `tests/e2e/ai/` checks the IR, compact text, and chunks on fixtures.
- Tests import TypeScript from `src` directly. Only `tests/cli` and `tests/types` use the built files.

Write expected values from the workbook, not from the current output. Open the file in Excel or read it with an independent tool, then assert what Excel shows.

## Fixtures

Fixtures are `.xlsx` files in `tests/fixtures/<topic>/`. The inventory is in `tests/fixtures/README.md`.

Rules:

- One file has one purpose. Use one sheet when possible.
- Use few values. Put a heading cell near the top that says what the sample tests.
- Name the file `<topic>-<what>-sampleNN.xlsx`.
- Save the file in Excel as usual, to keep the cached values. Use a hand-built workbook in `core.test.js` for a file without cached values.
- Remove the metadata that Excel adds:
    - the local path (`x15ac:absPath`) in `xl/workbook.xml`
    - the author, the last editor, the dates, and the application in `docProps/`
    - comments, threaded comments, persons, custom properties, external links, connections, embeddings, VBA, and printer settings

`tests/e2e/workbooks/fixture-hygiene.test.js` fails when a fixture keeps this metadata. To check a file by hand, run `unzip -l <file.xlsx>` and `unzip -p <file.xlsx> docProps/core.xml`.

## Code style

- TypeScript with `strict` on. ES modules. Write imports with the `.js` suffix.
- Import deep module paths directly. Group `index.ts` files only re-export.
- Prettier: 4 spaces, 120 columns, double quotes.
- Export plain functions and constants. Do not add `Api`, `Helper`, or `Manager` affixes. Keep one name for one thing.
- Write comments in English under ASD-STE100: short sentences, active voice, one meaning per word. Explain the reason, not the code. Skip comments where the code is clear.

## Benchmark

```bash
npm run bench:baseline -- --out bench-results/baseline.json
```

The report lists, for each fixture, the time of each stage, the memory change, the output size, and the AI context token counts for `rows` and `records`. The token counts are estimates, not the counts of a language model.

## Release

The package is not on the npm registry yet. `npm pack` builds `dist/` and writes the tarball. The tarball holds `bin/`, `dist/lib/`, `dist/types/`, `skills/`, `LICENSE`, and `README.md`.
