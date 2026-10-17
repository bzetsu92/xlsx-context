# Command line

The command `xlsx-context` converts one workbook to Markdown or to a ZIP archive.

The command line does not write AI context. Use the [API](./api.md) for AI context.

## Use

```bash
npx xlsx-context <input.xlsx> [options]
```

From a clone of the repository, run `npm run cli -- <input.xlsx> [options]`.

## Output

| You give       | The command writes                                                        |
| -------------- | ------------------------------------------------------------------------- |
| nothing        | `<workbook>.md` in the current directory.                                 |
| `--out <file>` | The combined Markdown in `<file>`.                                        |
| `--zip <file>` | A ZIP archive in `<file>`. No Markdown file, unless you also use `--out`. |

The Markdown has these parts:

1. YAML front matter, unless you use `--front-matter exclude`.
2. `# Book: <workbook>`.
3. One `## Sheet: <sheet>` section for each sheet, in sheet order.

The ZIP archive holds `output/<workbook>.md` and the extracted assets in `output/assets/`.

## Options

| Option                          | Values                                                               | Default    |
| ------------------------------- | -------------------------------------------------------------------- | ---------- |
| `--out <file>`                  | A file path.                                                         | none       |
| `--zip <file>`                  | A file path.                                                         | none       |
| `--output-mode <mode>`          | `display`, `raw`, `both`                                             | `display`  |
| `--formatting-mode <mode>`      | `plain`, `gfm`                                                       | `gfm`      |
| `--table-detection-mode <mode>` | `balanced`, `border`, `planner-aware`                                | `balanced` |
| `--encoding <value>`            | `utf-8`, `shift_jis`, `utf-16le`, `utf-16be`, `utf-32le`, `utf-32be` | `utf-8`    |
| `--bom <value>`                 | `off`, `on`. `shift_jis` does not allow `on`.                        | `off`      |
| `--shape-details <mode>`        | `include`, `exclude`. `include` adds the raw drawing entries.        | `exclude`  |
| `--front-matter <mode>`         | `include`, `exclude`                                                 | `include`  |
| `--hidden <mode>`               | `include`, `exclude`. `exclude` drops hidden sheets, rows, columns.  | `include`  |
| `--include-shape-details`       | Same as `--shape-details include`.                                   | off        |
| `--no-header-row`               | Do not treat the first row as a header.                              | off        |
| `--no-trim-text`                | Keep the spaces around cell text.                                    | off        |
| `--keep-empty-rows`             | Keep empty rows.                                                     | off        |
| `--keep-empty-columns`          | Keep empty columns.                                                  | off        |
| `--summary`                     | Print a summary of each sheet.                                       | off        |
| `--version`                     | Print the version and stop.                                          | -          |
| `--help`                        | Print the help and stop.                                             | -          |

`--encoding shift_jis` needs the optional package `iconv-lite`.

## Exit codes

| Code | Meaning  |
| ---- | -------- |
| `0`  | Success. |
| `1`  | Error.   |

A failure on a file has the form `[<file>.xlsx] <stage>: <message>`.
The stage is `read failed`, `parse failed`, `convert failed`, `zip write failed`, `markdown write failed`,
or `option validation failed`.

A wrong option gives a short message and exit code `1`:

- `Unknown option: --bogus`
- `Missing value for --out`
- `Invalid output mode: foo`

Without an input file, the command prints the help and exits with code `1`.

## Examples

```bash
npx xlsx-context sales.xlsx --out sales.md
npx xlsx-context sales.xlsx --zip sales.zip --summary
npx xlsx-context sales.xlsx --output-mode both --front-matter exclude --out sales.md
```
