import { Flags } from '@oclif/core';
import Command, { spreadsheetId } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class SpreadsheetCopy extends Command {
  static description =
    'Copy a spreadsheet into a new one through the Drive API (drive.file scope). ' +
    'Under drive.file only files this app created or has opened are visible. ' +
    'The source spreadsheet is left untouched; sharing settings are not duplicated.';

  static examples = [
    `$ gsheet spreadsheet:copy --spreadsheetId=<id> --title="Copy of <title>"

Spreadsheet "<id>" copied to "<title>" (<newId>) > https://docs.google.com/spreadsheets/d/<newId>/edit
`,
    `$ gsheet spreadsheet:copy --spreadsheetId=<id>

Spreadsheet "<id>" copied to "<sourceTitle>" (<newId>) > https://docs.google.com/spreadsheets/d/<newId>/edit
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    title: Flags.string({
      description: 'Title of the new spreadsheet (the source title is inherited when omitted)',
      required: false,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, title, rawOutput },
    } = await this.parse(SpreadsheetCopy);

    this.start('Copying spreadsheet');
    const result = await this.gsheet.copySpreadsheet(
      {
        title,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', { operation: this.id, ...result });
    } else {
      const link = result.webViewLink ? ` > ${result.webViewLink}` : '';
      this.log(`Spreadsheet "${result.sourceSpreadsheetId}" copied to "${result.title}" (${result.spreadsheetId})${link}`);
    }

    return { operation: this.id, ...result };
  }
}
