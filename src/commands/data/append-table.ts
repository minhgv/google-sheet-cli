import { Flags } from '@oclif/core';
import Command, { optionalData, spreadsheetId, valueInputOption } from '../../lib/base-class';
import { resolveDataMatrix } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class AppendTableData extends Command {
  static description =
    'Append table data using native Google Sheets values.append API (prevents client-side race conditions)';

  static examples = [
    `$ gsheet data:append-table --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> '[["1", "2", "3"]]'
`,
    `$ gsheet data:append-table --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --input=table.csv
`,
    `$ gsheet data:append-table --spreadsheetId=<spreadsheetId> --range='Sheet1!A1' --input=- --inputFormat=json < table.json
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle: Flags.string({
      char: 't',
      description: 'Title of the worksheet to append table into',
      required: false,
    }),
    range: Flags.string({
      description: 'Specific table range to search for existing data table (e.g. "Sheet1!A1")',
      required: false,
    }),
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
    valueInputOption,
    insertDataOption: Flags.string({
      description: 'How the input data should be inserted: "OVERWRITE" (overwrites blank cells) or "INSERT_ROWS" (inserts new rows)',
      options: ['OVERWRITE', 'INSERT_ROWS'],
      required: false,
    }),
  };

  static args = { data: optionalData };

  async run() {
    const {
      args: { data },
      flags: {
        spreadsheetId,
        worksheetTitle,
        range,
        input,
        inputFormat,
        valueInputOption,
        insertDataOption,
        rawOutput,
      },
    } = await this.parse(AppendTableData);

    const matrix = await resolveDataMatrix(data, input, inputFormat);

    if (matrix.length === 0) {
      throw new Error('Cannot append empty table data (0 rows)');
    }

    const targetDesc = worksheetTitle ? `worksheet "${worksheetTitle}"` : range ? `range "${range}"` : 'spreadsheet';
    this.start(`Appending ${matrix.length} row(s) to ${targetDesc}`);

    const resultReceipt = await this.gsheet.appendTableData(
      matrix,
      {
        worksheetTitle,
        range,
        valueInputOption: valueInputOption as GoogleSheetCli.ValueInputOption,
        insertDataOption: insertDataOption as 'OVERWRITE' | 'INSERT_ROWS' | undefined,
      },
      spreadsheetId
    );
    this.stop();

    const result = {
      operation: this.id,
      ...resultReceipt,
    };

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      this.log(`Successfully appended table data to spreadsheet "${spreadsheetId}":`);
      this.log(`  - Updated range: ${resultReceipt.updatedRange}`);
      this.log(`  - Table range: ${resultReceipt.tableRange}`);
      this.log(`  - Rows appended: ${resultReceipt.updatedRows} (${resultReceipt.updatedCells} cells)`);
    }

    return result;
  }
}
