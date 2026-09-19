import { Command, Flags } from '@oclif/core';
import { googleAuthFlags, GoogleAuthFlags, resolveGoogleSheetAuth } from '../../lib/base-class';
import { parseRangesFlag } from '../../lib/cli-input';
import { GoogleSheetCli, default as GoogleSheet } from '../../lib/google-sheet';
import { ReportRunOptions, runReport } from '../../lib/report/runner';

export default class ReportRun extends Command {
  static description =
    'Generate and publish business reports (finance, manpower, table) from JSON/CSV/XLSX/Sheets sources into XLSX or Google Sheets';

  static examples = [
    `$ gsheet report:run --template=templates/finance.json --input=data/transactions.csv --output=out/finance-report.xlsx
`,
    `$ gsheet report:run --template=templates/manpower.json --sourceWorkbook=data/tasks.xlsx --ranges='["Tasks!A1:H50"]' --allowCachedFormulaValues --output=out/manpower.xlsx --workbookTemplate=templates/styled-base.xlsx
`,
    `$ gsheet report:run --template=templates/finance.json --input=data/tx.json --spreadsheetId=<spreadsheetId> --dryRun
`,
    `$ gsheet report:run --template=templates/finance.json --sourceSpreadsheet=<sourceId> --ranges='["Sheet1!A1:Z100"]' --spreadsheetId=<targetId> --overwriteFormulas
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
    template: Flags.string({
      description: 'Path to report template JSON file',
      required: true,
    }),
    // Input source flags (mutually exclusive)
    input: Flags.string({
      char: 'i',
      description: 'Path to input data file (JSON or CSV) or "-" for stdin',
      required: false,
    }),
    inputFormat: Flags.string({
      description: 'Format of input file ("json" or "csv")',
      options: ['json', 'csv'],
      required: false,
    }),
    sourceWorkbook: Flags.string({
      description: 'Path to source local XLSX file',
      required: false,
    }),
    sourceSpreadsheet: Flags.string({
      description: 'ID of source Google Spreadsheet to read data from',
      required: false,
    }),
    ranges: Flags.string({
      description: 'JSON array of A1 range strings for source workbook or spreadsheet (e.g. \'["Sheet1!A1:D50"]\')',
      required: false,
    }),
    // Output target flags (mutually exclusive)
    output: Flags.string({
      char: 'o',
      description: 'Path for output destination XLSX file',
      required: false,
    }),
    spreadsheetId: Flags.string({
      char: 's',
      description: 'ID of target Google Spreadsheet to write report into',
      required: false,
    }),
    // Presentation and safety flags
    workbookTemplate: Flags.string({
      description: 'Path to existing XLSX template file to preserve formatting and styles in output file',
      required: false,
    }),
    allowCachedFormulaValues: Flags.boolean({
      description: 'Explicitly allow using cached formula values from source workbook (required when source contains formulas)',
      default: false,
      required: false,
    }),
    dryRun: Flags.boolean({
      description: 'Preview report generation without writing to disk or Google Sheets',
      default: false,
      required: false,
    }),
    overwrite: Flags.boolean({
      description: 'Allow overwriting existing output XLSX file',
      default: false,
      required: false,
    }),
    overwriteFormulas: Flags.boolean({
      description: 'Allow overwriting existing formula cells in target destination',
      default: false,
      required: false,
    }),
    // Reusable Google authentication flags
    ...googleAuthFlags,
  };

  async run() {
    const {
      flags: {
        template,
        input,
        inputFormat,
        sourceWorkbook,
        sourceSpreadsheet,
        ranges,
        output,
        spreadsheetId,
        workbookTemplate,
        allowCachedFormulaValues,
        dryRun,
        overwrite,
        overwriteFormulas,
        rawOutput,
        ...authFlags
      },
    } = await this.parse(ReportRun);

    // 1. Validate input source exclusivity before auth or file reading
    const sourceCount =
      (input !== undefined ? 1 : 0) +
      (sourceWorkbook !== undefined ? 1 : 0) +
      (sourceSpreadsheet !== undefined ? 1 : 0);

    if (sourceCount === 0) {
      throw new Error(
        'No input source specified. Must specify exactly one source: --input (file or "-"), --sourceWorkbook, or --sourceSpreadsheet'
      );
    }
    if (sourceCount > 1) {
      throw new Error(
        'Mutually exclusive input sources: specify only one of --input, --sourceWorkbook, or --sourceSpreadsheet'
      );
    }

    // 2. Validate range requirements for workbook/spreadsheet sources
    let parsedRanges: string[] | undefined;
    if (sourceWorkbook !== undefined || sourceSpreadsheet !== undefined) {
      if (!ranges) {
        throw new Error(
          'The --ranges flag (JSON array of A1 strings) is required when using --sourceWorkbook or --sourceSpreadsheet'
        );
      }
      parsedRanges = parseRangesFlag(ranges);
    }

    // 3. Validate output target exclusivity
    const targetCount =
      (output !== undefined ? 1 : 0) +
      (spreadsheetId !== undefined ? 1 : 0);

    if (targetCount === 0) {
      throw new Error(
        'No output target specified. Must specify exactly one target: --output (for XLSX file) or --spreadsheetId (for Google Spreadsheet)'
      );
    }
    if (targetCount > 1) {
      throw new Error(
        'Mutually exclusive output targets: specify only one of --output or --spreadsheetId'
      );
    }

    // 4. Authenticate Google client ONLY if sourceSpreadsheet or target spreadsheetId is used
    let gsheet: GoogleSheet | undefined;
    const needsGoogle = Boolean(sourceSpreadsheet || spreadsheetId);

    if (needsGoogle) {
      gsheet = await resolveGoogleSheetAuth(authFlags as GoogleAuthFlags, { prompt: false });
    }

    // 5. Construct options and execute runner
    const runOptions: ReportRunOptions = {
      template,
      source: {
        input: input !== undefined ? { file: input, format: inputFormat as 'json' | 'csv' } : undefined,
        sourceWorkbook: sourceWorkbook !== undefined ? { filePath: sourceWorkbook, ranges: parsedRanges!, allowCachedFormulaValues } : undefined,
        sourceSpreadsheet: sourceSpreadsheet !== undefined ? { spreadsheetId: sourceSpreadsheet, ranges: parsedRanges!, gsheet: gsheet! } : undefined,
      },
      target: {
        outputFile: output !== undefined ? { filePath: output, workbookTemplatePath: workbookTemplate, overwrite } : undefined,
        targetSpreadsheet: spreadsheetId !== undefined ? { spreadsheetId, gsheet: gsheet! } : undefined,
      },
      dryRun,
      overwriteFormulas,
    };

    const receipt = await runReport(runOptions);

    const result = {
      operation: this.id,
      ...receipt,
    };

    if (rawOutput) {
      this.log(JSON.stringify(result, null, 2));
    } else if (dryRun) {
      this.log(`Report Run Preview (Dry Run):`);
      this.log(`  - Template ID: ${receipt.templateId} (v${receipt.templateVersion})`);
      this.log(`  - Source: ${receipt.sourceKind} (hash: ${receipt.sourceHash.slice(0, 16)}...)`);
      this.log(`  - Target: ${receipt.targetKind}`);
      this.log(`  - Worksheets to generate (${receipt.sheetsGenerated}):`);
      for (const sheet of receipt.sheets) {
        this.log(`    • "${sheet.name}": ${sheet.rowCount} rows x ${sheet.colCount} cols`);
      }
    } else {
      this.log(`Successfully generated report "${receipt.templateId}" (v${receipt.templateVersion}):`);
      this.log(`  - Source: ${receipt.sourceKind} (hash: ${receipt.sourceHash.slice(0, 16)}...)`);
      this.log(`  - Target: ${receipt.targetKind}`);
      this.log(`  - Sheets generated (${receipt.sheetsGenerated}):`);
      for (const sheet of receipt.sheets) {
        this.log(`    • "${sheet.name}": ${sheet.rowCount} rows x ${sheet.colCount} cols`);
      }
      if (receipt.xlsxSaved) {
        this.log(`  - Saved XLSX: "${receipt.xlsxSaved.savedPath}" (${receipt.xlsxSaved.bytesWritten} bytes, SHA-256: ${receipt.xlsxSaved.sha256Hash})`);
      }
      if (receipt.googleReceipt) {
        this.log(`  - Google Sheets target: "${receipt.googleReceipt.spreadsheetId}"`);
        this.log(`  - Formulas protected: ${receipt.googleReceipt.formulasProtected}, overwritten: ${receipt.googleReceipt.formulasOverwritten}`);
      }
    }

    return result;
  }
}
