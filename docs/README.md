# Documentation

`xlsx-context` converts `.xlsx` workbooks into Markdown and into context for language models.

## Pages

| Page                                    | Read it to                                        |
| --------------------------------------- | ------------------------------------------------- |
| [Getting started](./getting-started.md) | Install the package and convert a workbook.       |
| [API reference](./api.md)               | Find every public function, option, and type.     |
| [AI context](./ai-context.md)           | Understand the text formats, chunks, and budgets. |
| [Command line](./cli.md)                | Convert a workbook from a terminal.               |
| [Architecture](./architecture.md)       | Learn the layers and the dependency rules.        |
| [Development](./development.md)         | Build, test, and change the code.                 |

## Terms

These pages use one word for one thing.

| Term         | Meaning                                                                    |
| ------------ | -------------------------------------------------------------------------- |
| workbook     | One `.xlsx` file.                                                          |
| sheet        | One worksheet in a workbook.                                               |
| table        | A group of cells that the library detects as a table.                      |
| narrative    | Text on a sheet that is outside any table.                                 |
| IR           | The intermediate representation. A data structure that holds the workbook. |
| compact text | The IR rendered as short text for a language model.                        |
| chunk        | One piece of compact text with a source location and a token count.        |
| token budget | The maximum number of tokens for one chunk.                                |
| counter      | A function that counts the tokens in a text.                               |
| fixture      | A sample `.xlsx` file that the tests use.                                  |

## Writing style

The pages follow the ASD-STE100 rules. Sentences are short. Each term has one meaning.
