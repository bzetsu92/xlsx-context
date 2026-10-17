# xlsx-context — agent instructions

TypeScript library. Converts `.xlsx` workbooks to Markdown and to structured AI context. Node.js 20 or later.

## Verify (run in this order)

```bash
npm run format:check
npm run typecheck
npm test
npm run typecheck:consumer
```

`npm test` builds `dist/` first, then runs `tests/e2e/unit`, `tests/e2e/workbooks` with `tests/e2e/ai`, and `tests/cli`. `tests/README.md` says what each folder checks.

## Conventions

- Plain named exports. No `Api`, `Helper`, or `Manager` affixes. One name for one thing.
- Write English comments under ASD-STE100. Use short sentences and active voice. Give one meaning per word. Comment the reason, not the code. Skip comments where the code is obvious.
- Import deep module paths directly, never through a group `index.ts`. Group barrels only re-export.
- Only names exported from `src/public/index.ts` are public API. Never widen it without review.
- Tests mirror `src/` paths. For example, `tests/e2e/unit/structure/tables/border-grid.test.js` covers `src/structure/tables/border-grid.ts`.
- Dependency direction: `public` → `core` → `outputs`/`formula`/`structure`/`extraction`, all over `model`. Outputs never import each other. Details live in `docs/architecture.md`.

## Read Excel files through conversion

Never parse `.xlsx` bytes by hand. Convert first. Then read the Markdown:

```bash
npx -y xlsx-context@latest ./book.xlsx --out /tmp/book.md --summary
```

Defaults are `display` text, `gfm` formatting, first row as header, and front matter included. The CLI exits `0` on success and `1` on error. It reads `.xlsx` only, never `.xls`. More flags: `npx -y xlsx-context@latest --help`.

Agent skill for consumers lives in `skills/xlsx-context/`. It uses the open SKILL.md format and ships in the npm tarball. It is a stub pointing at CLI `--help`. Keep it free of version-specific details.
