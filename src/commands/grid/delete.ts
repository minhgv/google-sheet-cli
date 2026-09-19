import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { dimensionFlags } from './insert';

export default class GridDelete extends Command {
  static description =
    'Delete rows or columns from a worksheet. ' +
    'Data after the deleted range shifts up/left. --dryRun previews the values about to be removed.';

  static examples = [
    `$ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=10 --count=3
`,
    `$ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=5 --dryRun
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    ...dimensionFlags,
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, dimension, start, count, dryRun, rawOutput },
    } = await this.parse(GridDelete);

    this.start(dryRun ? 'Previewing delete' : `Deleting ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'delete',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: would delete ${count} ${dimension.toLowerCase()} at ${start}`);
      if (result.affectedValues && result.affectedValues.length > 0) {
        this.log('Values about to be removed:');
        this.log(JSON.stringify(result.affectedValues, null, 2));
      } else {
        this.log('(no values in the affected range)');
      }
    } else {
      this.log(`Deleted ${count} ${dimension.toLowerCase()} at ${start} from "${worksheetTitle}"`);
    }

    return result;
  }
}
