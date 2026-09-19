import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';

export const dimensionFlags = {
  dimension: Flags.string({
    description: 'The dimension to mutate',
    options: ['ROWS', 'COLUMNS'],
    required: true,
  }),
  start: Flags.integer({ description: 'The 1-based first row/column index to affect', required: true }),
  count: Flags.integer({ description: 'How many rows/columns the operation covers', default: 1, required: false }),
  dryRun: Flags.boolean({ description: 'Preview the batchUpdate request without applying it', required: false }),
};

export default class GridInsert extends Command {
  static description =
    'Insert rows or columns into a worksheet at a position. ' +
    'Existing data at and after the position shifts down/right.';

  static examples = [
    `$ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=5 --count=2
`,
    `$ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=3 --inheritFromBefore
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    ...dimensionFlags,
    inheritFromBefore: Flags.boolean({
      description: 'Inherit formatting from the row/column before instead of after (matches "insert above/left")',
      required: false,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, dimension, start, count, inheritFromBefore, dryRun, rawOutput },
    } = await this.parse(GridInsert);

    this.start(dryRun ? 'Previewing insert' : `Inserting ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'insert',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        inheritFromBefore,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: insert ${count} ${dimension.toLowerCase()} at ${start}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`Inserted ${count} ${dimension.toLowerCase()} at ${start} in "${worksheetTitle}"`);
    }

    return result;
  }
}
