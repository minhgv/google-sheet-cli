import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';

export default class WorksheetCopy extends Command {
  static description =
    'Copy a worksheet to an explicit destination spreadsheet through the Sheets API. ' +
    'The copy carries the source title; when the destination already has a worksheet with that title, the server assigns a unique one (e.g. "Sheet1 Copy") and the receipt reports it. ' +
    'Values and formulas are carried over; the source worksheet is left untouched.';

  static examples = [
    `$ gsheet worksheet:copy --spreadsheetId=<id> --worksheetTitle=<worksheetTitle> --destinationSpreadsheetId=<destId>

Worksheet "<worksheetTitle>" copied to spreadsheet <destId> as sheet <sheetId>
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    destinationSpreadsheetId: Flags.string({
      description: 'ID of the spreadsheet to copy the worksheet into',
      required: true,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle = '', destinationSpreadsheetId, rawOutput },
    } = await this.parse(WorksheetCopy);

    this.start('Copying worksheet');
    const result = await this.gsheet.copyWorksheet(
      {
        worksheetTitle,
        destinationSpreadsheetId,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', { operation: this.id, ...result });
    } else {
      this.log(`Worksheet "${result.worksheetTitle}" copied to spreadsheet ${result.destinationSpreadsheetId} as sheet ${result.sheetId} ("${result.title}")`);
    }

    return { operation: this.id, ...result };
  }
}
