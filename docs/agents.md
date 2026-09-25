# Using google-sheet-cli with coding agents

google-sheet-cli is designed to be driven by coding agents (Claude Code, OpenCode/Oh My Pi, Cursor, or a plain shell script): offline commands, machine-readable output, and writes that fail closed.

## What makes it agent-friendly

- **Offline by default where possible.** `workbook:inspect|read|find|write|names`, `report:run` with local input/output, and `grid:insert|delete|freeze|resize|hide`, `format:cells|merge`, `worksheet:add|remove|rename`, `data:clear` with `--workbook` need no credentials and make zero network calls. An agent can manipulate local Excel files and generate reports in a sandbox with no secrets at all.
- **Machine-readable output.** Every command accepts `--rawOutput`/`-r` (JSON) and table output supports `--csv`, `--columns`, `--sort`, `--filter` — pipe through `jq` instead of parsing prose.
- **Deterministic input.** All data commands accept positional JSON, `-i <file>` (format inferred from extension), or `-i -` stdin.
- **Writes fail closed.** Existing formulas are never overwritten without explicit `--overwriteFormulas`; `--dryRun` previews every mutation (cells, conflicts, managed-range extent) with zero side effects.
- **Locate before writing.** `data:find` returns A1 coordinates for a value/header/regex so an agent can target `data:update` at an exact cell instead of rewriting a whole range.
- **Formatting is a separate, value-safe surface.** `format:cells`/`format:merge` only write `userEnteredFormat` (bold, colors, borders, number formats, merge) — they cannot clobber data. Use them to render a readable report after the data is in place.
- **Structure is mutable.** `grid:insert|delete|hide|resize|freeze` move rows/columns without touching values. `grid:delete --dryRun` previews the exact values about to be lost — always dry-run a delete first.
- **Sharing is Drive, not Sheets.** `spreadsheet:share|permissions|unshare` call the Drive API under `drive.file` — files this app created or has opened only; a pre-existing sheet may need one `spreadsheet:get` first. `--notify` defaults OFF so agents never spam grantees. Tokens issued before the scope was added must re-run `auth:login`.
- **Discovery is scope-honest.** `spreadsheet:list` searches Drive under `drive.file` — it sees only files this app created or opened, so an empty result does not prove absence; address pre-existing sheets by ID. Duplicate titles stay separate rows; the command never auto-selects.
- **CSV export is deterministic and injection-aware.** `data:export-csv` works on both backends (Sheets via `-s`/`-t`, local `.xlsx` via `--workbook`): `--mode raw|formatted|formula`, `--injection safe|preserve` (safe prefixes `= + - @` tab CR with `'`; typed negative numbers exempt), atomic `-o` writes at mode 0600 with `--overwrite` consent, or pure CSV bytes on stdout.
- **Mutation failures are typed.** Multi-request writes (`data:batch-update`, `data:upsert`) attach `error.mutation` to the failure envelope: per-request `acknowledged`/`rejected`/`unknown`/`not-attempted` states with logical A1 ranges, `gridGrowth` tracked separately, and `retryGuidance` (`safe-replay`/`verify-then-replay`/`never-blind-replay`). `unknown` means the write may have been applied — verify before replaying.
- **Diagnostics can be redacted.** `--redacted` (env `GSHEET_REDACTED`) strips cell contents, formulas, incoming values and credentials from error envelopes and dry-run output before serialization — safe for transcripts and logs; coordinates, counts and statuses survive.
- **Self-description is machine-readable.** `capabilities` prints the capability document (schemaVersion 1) with no credentials: per-backend operations, I/O forms, formula write-vs-recalculation semantics, destructive guards, mutation limits, XLSX fidelity exclusions, and the `drive.file` visibility boundary. Listed ≠ authorized — runtime auth still governs.
- **Local workbooks are structurally mutable.** `grid:insert|delete --workbook <f.xlsx>` shift rows/columns and manage merges explicitly (a merge intersecting the splice boundary refuses unless `--force`); `format:cells --workbook` styles ranges (bold, colors, borders, number formats, wrap — `--wrapText` is local-only; `--numberFormatType` and non-WRAP `--wrapStrategy` refuse locally); `grid:freeze|resize|hide --workbook` freeze panes (`--rows`/`--columns`, 0 unfreezes), size rows/columns (`--pixels`/`--auto`), hide/unhide rows/columns; `worksheet:add|remove|rename --workbook` manage sheets (duplicate titles refuse, the last visible sheet cannot be removed, rename collisions refuse; `worksheet:copy` is cloud-only); `workbook:names` lists/adds/removes defined names (local-only, no auth); `format:merge --workbook` merges/unmerges; `data:clear --workbook` clears a bounded range; `workbook:write --cells '{...}'` writes only the listed addresses. Mutations land on `--output` or `--inPlace` (with `.bak` backup), never implicitly on the source. Formula references are rewritten only via `grid:insert|delete --updateRefs` — same-sheet A1 refs and defined names (refs fully inside a deleted span become `#REF!`); cross-sheet refs are deferred — receipts report `refsRewritten`/`refsBroken`, otherwise `formulasAtRisk`. Charts/pivots/macros cannot be preserved and saving a workbook that has them needs `--discardUnsupported`.

## Rules for agents

1. **Credentials via environment or file, never inline.** Set `GSHEET_CLIENT_EMAIL`/`GSHEET_PRIVATE_KEY` or pass `-f service-account.json`. Never put `-p '-----BEGIN PRIVATE KEY-----...'` on a command line — it leaks into shell history and agent transcripts.
2. **OAuth is interactive.** `auth:login` opens a browser; only use it when a human can complete the consent flow. Agents in CI/headless environments must use a service account.
3. **Credential storage is owner-only and relocatable.** OAuth tokens are written atomically at mode 0600 under `~/.config/google-sheet-cli` (dir 0700); `GSHEET_CONFIG_DIR` overrides that directory for tests and sandboxed runs. Never point it at a shared or world-readable path.
4. **Safety ladder for writes:** read (`workbook:read` / `data:get`) → `--dryRun` preview → real write. Add `--overwrite`/`--overwriteFormulas` only after the dry-run shows exactly what will change.
5. **Prefer local over cloud for intermediate artifacts.** Build and verify a report as a local `.xlsx` first; push to Google Sheets only as the final step.
6. **Exact tab titles.** Worksheet titles are case-sensitive and may contain spaces — quote them; list them with `spreadsheet:get -s <id> --rawOutput | jq '.sheets[].properties.title'`.
7. **Report arithmetic is exact-decimal and reconciled.** Finance reports are cash-flow (not P&L); manpower reports are effort in person-days (not schedule duration). Do not reinterpret the outputs.

## Install the bundled skill

The repository ships an agent skill describing every command, flag and workflow: [`skill/SKILL.md`](../skill/SKILL.md). Point your agent at it:

```sh
# Oh My Pi / OpenCode
mkdir -p ~/.omp/skills/google-sheet && cp skill/SKILL.md ~/.omp/skills/google-sheet/SKILL.md

# Claude Code
mkdir -p ~/.claude/skills/google-sheet && cp skill/SKILL.md ~/.claude/skills/google-sheet/SKILL.md
```

`skill/SKILL.md` is the source of truth — re-copy it after pulling changes. Its frontmatter `description` carries the trigger words your agent matches against.
