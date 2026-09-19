import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { dimensionFlags } from './insert';

export default class GridHide extends Command {
  static description = 'Hide or unhide rows or columns on a worksheet.';

  static examples = [
    `$ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=8 --count=3
`,
    `$ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=2 --unhide
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    ...dimensionFlags,
    unhide: Flags.boolean({ description: 'Unhide instead of hide', required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, dimension, start, count, unhide, dryRun, rawOutput },
    } = await this.parse(GridHide);

    const verb = unhide ? 'Unhiding' : 'Hiding';
    this.start(dryRun ? `Previewing ${verb.toLowerCase()}` : `${verb} ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'hide',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        unhide,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: ${verb.toLowerCase()} ${count} ${dimension.toLowerCase()} at ${start}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`${unhide ? 'Unhid' : 'Hid'} ${count} ${dimension === 'ROWS' ? 'row(s)' : 'column(s)'} at ${start} in "${worksheetTitle}"`);
    }

    return result;
  }
}
