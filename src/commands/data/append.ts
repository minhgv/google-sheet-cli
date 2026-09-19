import { Flags } from '@oclif/core';
import Command, { optionalData, spreadsheetId, valueInputOption, worksheetTitle } from '../../lib/base-class';
import { resolveDataMatrix } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class AppendData extends Command {
  static description =
    'Append cells with the specified data after the last row in starting col (legacy append semantics; for native table appends use data:append-table)';

  static examples = [
    `$ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'

Data successfully appended to "<worksheetTitle>"
`,
    `$ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=rows.csv

Data successfully appended to "<worksheetTitle>"
`,
    `$ gsheet data:append --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=- --inputFormat=json < rows.json

Data successfully appended to "<worksheetTitle>"
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle,
    spreadsheetId,
    valueInputOption,
    minCol: Flags.integer({ description: 'The optional starting col of the operation', default: 1, required: false }),
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
    dryRun: Flags.boolean({
      description: 'Preview data without modifying worksheet',
      default: false,
      required: false,
    }),
  };

  static args = { data: optionalData };

  async run() {
    const {
      args: { data },
      flags: { minCol, worksheetTitle = '', spreadsheetId, valueInputOption, input, inputFormat, dryRun },
    } = await this.parse(AppendData);

    const parsed = await resolveDataMatrix(data, input, inputFormat);

    if (dryRun) {
      const result = { operation: this.id, worksheetTitle, data: parsed, dryRun: true };
      this.logRaw(`Data append preview (dry run) for "${worksheetTitle}" (${parsed.length} rows)`, result);
      return result;
    }

    this.start('Appending data');
    const options = {
      worksheetTitle,
      minCol,
      valueInputOption: valueInputOption as GoogleSheetCli.ValueInputOption,
    };

    await this.gsheet.appendData(parsed, options, spreadsheetId);
    this.stop();
    const result = { operation: this.id, worksheetTitle, data: parsed };
    this.logRaw(`Data successfully appended to "${worksheetTitle}"`, result);
    return result;
  }
}
