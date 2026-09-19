import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { dimensionFlags } from './insert';

export default class GridResize extends Command {
  static description =
    'Resize rows or columns to an explicit pixel size, or auto-size them to their content.';

  static examples = [
    `$ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --count=4 --pixels=120
`,
    `$ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --auto
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    ...dimensionFlags,
    pixels: Flags.integer({ description: 'Explicit pixel size', required: false, exclusive: ['auto'] }),
    auto: Flags.boolean({ description: 'Auto-size to content', required: false, exclusive: ['pixels'] }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, dimension, start, count, pixels, auto, dryRun, rawOutput },
    } = await this.parse(GridResize);

    this.start(dryRun ? 'Previewing resize' : `Resizing ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'resize',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        pixels,
        autoResize: auto,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: resize ${count} ${dimension.toLowerCase()} at ${start} (${auto ? 'auto' : `${pixels}px`})`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`Resized ${count} ${dimension.toLowerCase()} at ${start} in "${worksheetTitle}" (${auto ? 'auto' : `${pixels}px`})`);
    }

    return result;
  }
}
