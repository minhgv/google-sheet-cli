# Using google-sheet-cli with coding agents

google-sheet-cli is designed to be driven by coding agents (Claude Code, OpenCode/Oh My Pi, Cursor, or a plain shell script): offline commands, machine-readable output, and writes that fail closed.

## What makes it agent-friendly

- **Offline by default where possible.** `workbook:inspect|read|write` and `report:run` with local input/output need no credentials and make zero network calls. An agent can manipulate local Excel files and generate reports in a sandbox with no secrets at all.
- **Machine-readable output.** Every command accepts `--rawOutput`/`-r` (JSON) and table output supports `--csv`, `--columns`, `--sort`, `--filter` — pipe through `jq` instead of parsing prose.
- **Deterministic input.** All data commands accept positional JSON, `-i <file>` (format inferred from extension), or `-i -` stdin.
- **Writes fail closed.** Existing formulas are never overwritten without explicit `--overwriteFormulas`; `--dryRun` previews every mutation (cells, conflicts, managed-range extent) with zero side effects.
- **Locate before writing.** `data:find` returns A1 coordinates for a value/header/regex so an agent can target `data:update` at an exact cell instead of rewriting a whole range.
- **Formatting is a separate, value-safe surface.** `format:cells`/`format:merge` only write `userEnteredFormat` (bold, colors, borders, number formats, merge) — they cannot clobber data. Use them to render a readable report after the data is in place.
- **Structure is mutable.** `grid:insert|delete|hide|resize|freeze` move rows/columns without touching values. `grid:delete --dryRun` previews the exact values about to be lost — always dry-run a delete first.
- **Sharing is Drive, not Sheets.** `spreadsheet:share|permissions|unshare` call the Drive API under `drive.file` — files this app created or has opened only; a pre-existing sheet may need one `spreadsheet:get` first. `--notify` defaults OFF so agents never spam grantees. Tokens issued before the scope was added must re-run `auth:login`.

## Rules for agents

1. **Credentials via environment or file, never inline.** Set `GSHEET_CLIENT_EMAIL`/`GSHEET_PRIVATE_KEY` or pass `-f service-account.json`. Never put `-p '-----BEGIN PRIVATE KEY-----...'` on a command line — it leaks into shell history and agent transcripts.
2. **OAuth is interactive.** `auth:login` opens a browser; only use it when a human can complete the consent flow. Agents in CI/headless environments must use a service account.
3. **Safety ladder for writes:** read (`workbook:read` / `data:get`) → `--dryRun` preview → real write. Add `--overwrite`/`--overwriteFormulas` only after the dry-run shows exactly what will change.
4. **Prefer local over cloud for intermediate artifacts.** Build and verify a report as a local `.xlsx` first; push to Google Sheets only as the final step.
5. **Exact tab titles.** Worksheet titles are case-sensitive and may contain spaces — quote them; list them with `spreadsheet:get -s <id> --rawOutput | jq '.sheets[].properties.title'`.
6. **Report arithmetic is exact-decimal and reconciled.** Finance reports are cash-flow (not P&L); manpower reports are effort in person-days (not schedule duration). Do not reinterpret the outputs.

## Install the bundled skill

The repository ships an agent skill describing every command, flag and workflow: [`skill/SKILL.md`](../skill/SKILL.md). Point your agent at it:

```sh
# Oh My Pi / OpenCode
mkdir -p ~/.omp/skills/google-sheet && cp skill/SKILL.md ~/.omp/skills/google-sheet/SKILL.md

# Claude Code
mkdir -p ~/.claude/skills/google-sheet && cp skill/SKILL.md ~/.claude/skills/google-sheet/SKILL.md
```

`skill/SKILL.md` is the source of truth — re-copy it after pulling changes. Its frontmatter `description` carries the trigger words your agent matches against.
