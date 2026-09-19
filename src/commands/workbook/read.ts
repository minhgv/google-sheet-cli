import { Command, Flags } from '@oclif/core';
import { XlsxReadMode, XlsxWorkbook } from '../../lib/xlsx';

export default class WorkbookRead extends Command {
  static description = 'Read cell values and formula freshness from a local XLSX workbook range';

  static examples = [
    `$ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10'
`,
    `$ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10' --mode=formatted
`,
    `$ gsheet workbook:read --file=report.xlsx --range='Sheet1!A1:D10' --rawOutput
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    file: Flags.string({
      char: 'f',
      description: 'Path to the local XLSX file to read',
      required: true,
    }),
    range: Flags.string({
      description: 'A1 range to read (e.g. "Sheet1!A1:D10")',
      required: true,
    }),
    mode: Flags.string({
      description: 'Value extraction mode ("unformatted", "formatted", or "formula")',
      options: ['unformatted', 'formatted', 'formula'],
      default: 'unformatted',
      required: false,
    }),
    maxRows: Flags.integer({
      description: 'Maximum rows to read',
      required: false,
    }),
    maxCols: Flags.integer({
      description: 'Maximum columns to read',
      required: false,
    }),
    includeEmpty: Flags.boolean({
      description: 'Include trailing empty cells in output rows',
      default: false,
      required: false,
    }),
    rawOutput: Flags.boolean({
      char: 'r',
      description: 'Get the raw output as a JSON string',
      default: false,
      required: false,
    }),
  };

  async run() {
    const {
      flags: { file, range, mode, maxRows, maxCols, includeEmpty, rawOutput },
    } = await this.parse(WorkbookRead);

    const workbook = await XlsxWorkbook.load(file);
    const readResult = workbook.read(range, {
      mode: mode as XlsxReadMode,
      maxRows,
      maxCols,
      includeEmpty,
    });

    const result = {
      operation: this.id,
      ...readResult,
    };

    if (rawOutput) {
      this.log(JSON.stringify(result, null, 2));
    } else {
      this.log(`Read ${readResult.rowCount} row(s) and ${readResult.colCount} column(s) from "${readResult.range}" (mode: ${readResult.mode}):`);
      this.log(`  - Total formulas: ${readResult.formulaCount}`);
      this.log(`  - Freshness: cached=${readResult.freshnessSummary.cachedCount}, unknown=${readResult.freshnessSummary.unknownCount}, not-formula=${readResult.freshnessSummary.notFormulaCount}`);
      this.log('');
      for (const row of readResult.values) {
        const formattedCells = row.map((cell) => {
          if (cell === null || cell === undefined) return '';
          if (typeof cell === 'object' && 'formula' in cell) {
            return `=${cell.formula}${cell.result !== undefined ? ` [${cell.result}]` : ''}`;
          }
          return String(cell);
        });
        this.log(formattedCells.join('\t'));
      }
    }

    return result;
  }
}
