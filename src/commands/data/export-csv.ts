import { Flags } from '@oclif/core';
import { existsSync } from 'fs';
import Command from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { CsvCell, CsvInjectionPolicy, serializeCsv } from '../../lib/csv';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { XlsxReadMode, XlsxWorkbook } from '../../lib/xlsx';
import type { ReportCell } from '../../lib/xlsx-types';
import { saveBufferAtomic } from '../../lib/xlsx-file';

type ExportRenderMode = 'raw' | 'formatted' | 'formula';

const RENDER_MODES: ExportRenderMode[] = ['raw', 'formatted', 'formula'];

const INJECTION_POLICIES: CsvInjectionPolicy[] = ['safe', 'preserve'];

/**
 * The command's render mode translated onto each backend. `src/lib/csv.ts` is the locked
 * serializer: it receives already-rendered cells and owns quoting, escaping and the
 * injection policy - nothing here pre-formats or pre-escapes cell content.
 */
const SHEETS_VALUE_RENDER_OPTION: Record<ExportRenderMode, GoogleSheetCli.ValueRenderOption> = {
  raw: GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE,
  formatted: GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE,
  formula: GoogleSheetCli.ValueRenderOption.FORMULA,
};

// raw dates as serial numbers keep the machine-consumable export deterministic; formatted
// and formula exports carry the display strings a human expects.
const SHEETS_DATE_RENDER_OPTION: Record<ExportRenderMode, GoogleSheetCli.DateTimeRenderOption> = {
  raw: GoogleSheetCli.DateTimeRenderOption.SERIAL_NUMBER,
  formatted: GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING,
  formula: GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING,
};

const XLSX_READ_MODE: Record<ExportRenderMode, XlsxReadMode> = {
  raw: 'unformatted',
  formatted: 'formatted',
  formula: 'formula',
};

/**
 * Render one XlsxWorkbook cell for the serializer. A formula cell becomes its formula text
 * (`=...`), deliberately WITHOUT calling csvEscapeFormulaText first: serializeCsv applies the
 * identical safe/preserve policy to the strings it receives, and pre-escaping here would
 * double-prefix the cell under the safe policy.
 */
const renderXlsxCell = (cell: ReportCell): CsvCell => {
  if (cell !== null && typeof cell === 'object' && 'formula' in cell) {
    return `=${cell.formula}`;
  }
  return cell;
};

/**
 * Export worksheet data as deterministic RFC 4180 CSV from either backend: a Google Sheets
 * worksheet (--spreadsheetId + --worksheetTitle, bounded by the optional --range) or a
 * local XLSX workbook (--workbook + --range). The local source uses the long flag name only:
 * the shared short -f belongs to --credentialsFile, and -o stays free for the output path.
 * With --output the CSV is written atomically (temporary file in the same directory, then
 * renamed, mode 0600) and an existing file is only replaced with --overwrite. Without it the
 * CSV bytes are the only thing stdout carries and no receipt is printed; warnings still go
 * to stderr so the receipt caveats are never lost.
 */
export default class ExportCsv extends Command {
  static description =
    'Export worksheet data as deterministic RFC 4180 CSV, from a Google Sheets worksheet or a local XLSX workbook. ' +
    'Cell rendering is selectable (raw values, formatted display strings, or formula text where the backend stores it), ' +
    'and the injection policy for string cells is explicit: safe (default) prefixes dangerous leading characters with an apostrophe, ' +
    'preserve emits them byte-faithfully and says so in the receipt. ' +
    'Local XLSX formula results are the cached values from the last save by the producing application - they are never recalculated.';

  static examples = [
    `$ gsheet data:export-csv --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --output=./data.csv

Exported 3 row(s) x 2 column(s) from <spreadsheetId>/Sheet1 (Sheet1!A1:B3) to ./data.csv (48 bytes, mode: raw, injection: safe, transformed cells: 0)
`,
    `$ gsheet data:export-csv --workbook=./report.xlsx --range='Sheet1!A1:D10' --mode=formatted --output=./report.csv --overwrite

Exported 10 row(s) x 4 column(s) from ./report.xlsx (Sheet1!A1:D10) to ./report.csv (312 bytes, mode: formatted, injection: safe, transformed cells: 1)
`,
    `$ gsheet data:export-csv --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range=A1:B10 --mode=formula --injection=preserve

=B2*2
=SUM(A1:A10)
`,
  ];

  static flags = {
    ...Command.flags,
    // Optional variants of the shared required flags: a --workbook export must not be forced
    // to name a spreadsheet. Same chars and env bindings as the shared definitions.
    spreadsheetId: Flags.string({
      char: 's',
      description: 'ID of the spreadsheet to export (Google Sheets source)',
      required: false,
      env: 'SPREADSHEET_ID',
    }),
    worksheetTitle: Flags.string({
      char: 't',
      description: 'Title of the worksheet to export (Google Sheets source)',
      required: false,
      env: 'WORKSHEET_TITLE',
    }),
    range: Flags.string({
      description:
        'A1 range to export, e.g. "Sheet1!A1:D10" (optional for the Google Sheets source, where it bounds the used range of --worksheetTitle; required with --workbook)',
      required: false,
    }),
    workbook: Flags.string({
      description:
        'Path to a local .xlsx workbook to export instead of the Google Sheets source. Long name only: the shared short -f belongs to --credentialsFile',
      required: false,
    }),
    mode: Flags.string({
      description:
        'Cell rendering: "raw" machine values (default), "formatted" display strings, "formula" formula text where the backend stores it (XLSX always reports cached formula results)',
      options: RENDER_MODES,
      default: 'raw',
      required: false,
    }),
    injection: Flags.string({
      description:
        'Formula-injection policy for string cells: "safe" prefixes dangerous leading characters (=, +, -, @, tab, CR) with an apostrophe; "preserve" emits them byte-faithfully and warns in the receipt',
      options: INJECTION_POLICIES,
      default: 'safe',
      required: false,
    }),
    output: Flags.string({
      char: 'o',
      description:
        'Path of the CSV file to write atomically (an existing file needs --overwrite). Without it, the CSV goes to stdout and no receipt is printed',
      required: false,
    }),
    overwrite: Flags.boolean({
      description: 'Replace an existing output file (default: refuse)',
      required: false,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, range, workbook, mode, injection, output, overwrite, rawOutput },
    } = await this.parse(ExportCsv);

    const renderMode = mode as ExportRenderMode;
    const injectionPolicy = injection as CsvInjectionPolicy;

    this.start('Exporting CSV');
    const warnings: string[] = [];
    let source: 'sheets' | 'xlsx';
    let sourceLabel: string;
    let exportRange: string;
    let rows: CsvCell[][];

    // Exactly one source. -s/-t belong to the Sheets backend, --workbook to the local one.
    // Validation and the read share one branch each, so the option strings are narrowed
    // exactly where they are used.
    if (workbook) {
      if (spreadsheetId || worksheetTitle) {
        throw new GSheetError(
          GSheetErrorCode.VALIDATION,
          '--workbook exports a local file and cannot be combined with --spreadsheetId or --worksheetTitle. Pass exactly one source.'
        );
      }
      if (!range) {
        throw new GSheetError(GSheetErrorCode.VALIDATION, '--range is required with --workbook, e.g. --range="Sheet1!A1:D10".');
      }

      source = 'xlsx';
      sourceLabel = workbook;
      const loaded = await XlsxWorkbook.load(workbook);
      const read = loaded.read(range, { mode: XLSX_READ_MODE[renderMode] });
      rows = read.values.map((row) => row.map(renderXlsxCell));
      exportRange = read.range;

      // Local formula results are the cache written by the producing application; the read
      // result's freshness counts are the honest source for how much of the export is cache.
      if (renderMode === 'formula') {
        if (read.freshnessSummary.unknownCount > 0) {
          warnings.push(
            `xlsx formula export: ${read.freshnessSummary.unknownCount} formula cell(s) have no cached result, so their computed values are unavailable; the CSV carries the formula text only`
          );
        }
      } else {
        if (read.formulaCount > 0) {
          warnings.push(
            `xlsx ${renderMode} export: ${read.formulaCount} formula cell(s) exported their cached results as saved by the producing application; local XLSX is never recalculated`
          );
        }
        if (read.freshnessSummary.unknownCount > 0) {
          warnings.push(
            `xlsx ${renderMode} export: ${read.freshnessSummary.unknownCount} formula cell(s) had no cached result and exported as empty fields`
          );
        }
      }
    } else if (!spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No source given. Pass --spreadsheetId (with --worksheetTitle) for the Google Sheets source, or --workbook for a local XLSX file.'
      );
    } else if (!worksheetTitle) {
      throw new GSheetError(GSheetErrorCode.VALIDATION, '--worksheetTitle is required for the Google Sheets source.');
    } else {
      source = 'sheets';
      sourceLabel = `${spreadsheetId}/${worksheetTitle}`;
      const data = await this.gsheet.getData(
        {
          worksheetTitle,
          range,
          valueRenderOption: SHEETS_VALUE_RENDER_OPTION[renderMode],
          dateTimeRenderOption: SHEETS_DATE_RENDER_OPTION[renderMode],
        },
        spreadsheetId
      );
      rows = data.rawData;
      // the API echoes the bounded range it actually served; absent only on degenerate responses
      exportRange = data.range ?? 'unknown';
    }

    // serializeCsv refuses ragged input, so short rows are padded with empty fields up to the
    // widest row BEFORE serialization. Both current adapters return rectangular rows; the pad
    // is the documented deterministic behavior if an adapter ever hands over a short row, and
    // the receipt reports it only when it actually happened.
    const columns = rows.reduce((max, row) => Math.max(max, row.length), 0);
    let raggedRows = 0;
    const rectangular = rows.map((row) => {
      if (row.length === columns) return row;
      raggedRows++;
      const padded = row.slice();
      while (padded.length < columns) padded.push(null);
      return padded;
    });
    if (raggedRows > 0) {
      warnings.push(`${raggedRows} row(s) were shorter than the widest row (${columns} column(s)) and were padded with empty fields`);
    }

    const { csv, receipt } = serializeCsv(rectangular, { injectionPolicy });
    warnings.push(...receipt.warnings);

    if (!output) {
      // stdout is the data channel: bytes only, no receipt. Warnings reach the user on stderr
      // so preserve-policy and cache caveats are never lost, without touching the CSV bytes.
      this.stop();
      if (warnings.length > 0) {
        process.stderr.write(`${warnings.map((warning) => `warning: ${warning}`).join('\n')}\n`);
      }
      process.stdout.write(csv);
      return csv;
    }

    // Same consent gate as spreadsheet:export; the atomic save re-checks after serialization,
    // so a file created in between still cannot be clobbered.
    if (!overwrite && existsSync(output)) {
      this.error(`Output file "${output}" already exists. Pass --overwrite to replace it.`, { exit: 1 });
    }
    const saved = await saveBufferAtomic(Buffer.from(csv, 'utf8'), output, { overwrite: Boolean(overwrite) });
    this.stop();

    const result = {
      operation: this.id,
      source,
      range: exportRange,
      rows: receipt.rows,
      columns: receipt.columns,
      mode: renderMode,
      injectionPolicy: receipt.policy,
      transformedCells: receipt.transformedCells,
      savedPath: saved.savedPath,
      bytesWritten: saved.bytesWritten,
      warnings,
    };

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      this.log(
        `Exported ${receipt.rows} row(s) x ${receipt.columns} column(s) from ${sourceLabel} (${exportRange}) to ${saved.savedPath} ` +
          `(${saved.bytesWritten} bytes, mode: ${renderMode}, injection: ${receipt.policy}, transformed cells: ${receipt.transformedCells})`
      );
      for (const warning of warnings) {
        this.log(`  - warning: ${warning}`);
      }
    }
    return result;
  }
}
