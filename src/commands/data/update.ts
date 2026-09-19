import { Flags } from '@oclif/core';
import Command, { optionalData, spreadsheetId, valueInputOption, worksheetTitle } from '../../lib/base-class';
import { resolveDataMatrix } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class UpdateData extends Command {
  static description = 'Updates cells with the specified data';

  static examples = [
    `$ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'

Data successfully updated in "<worksheetTitle>"
`,
    `$ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=update.csv

Data successfully updated in "<worksheetTitle>"
`,
    `$ gsheet data:update --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=- --inputFormat=json < update.json

Data successfully updated in "<worksheetTitle>"
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle,
    spreadsheetId,
    valueInputOption,
    minRow: Flags.integer({ description: 'The optional starting row of the operation', default: 1, required: false }),
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
      description: 'Preview update without modifying worksheet',
      default: false,
      required: false,
    }),
  };

  static args = { data: optionalData };

  async run() {
    const {
      args: { data },
      flags: { minRow, minCol, worksheetTitle = '', spreadsheetId, valueInputOption, input, inputFormat, dryRun },
    } = await this.parse(UpdateData);

    const parsed = await resolveDataMatrix(data, input, inputFormat);

    if (dryRun) {
      const result = { operation: this.id, worksheetTitle, data: parsed, dryRun: true };
      this.logRaw(`Data update preview (dry run) for "${worksheetTitle}" (${parsed.length} rows)`, result);
      return result;
    }

    this.start('Updating data');
    const options = {
      worksheetTitle,
      minCol,
      minRow,
      valueInputOption: valueInputOption as GoogleSheetCli.ValueInputOption,
    };

    await this.gsheet.updateData(parsed, options, spreadsheetId);
    this.stop();
    const result = { operation: this.id, worksheetTitle, data: parsed };
    this.logRaw(`Data successfully updated in "${worksheetTitle}"`, result);
    return result;
  }
}
