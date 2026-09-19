import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';

export default class GridFreeze extends Command {
  static description =
    'Freeze or unfreeze rows and columns on a worksheet. ' +
    'At least one of --rows/--columns is required; 0 unfreezes that axis.';

  static examples = [
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1
`,
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1 --columns=2
`,
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=0 --columns=0
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    rows: Flags.integer({ description: 'Number of rows to freeze (0 unfreezes)', required: false }),
    columns: Flags.integer({ description: 'Number of columns to freeze (0 unfreezes)', required: false }),
    dryRun: Flags.boolean({ description: 'Preview the batchUpdate request without applying it', required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, rows, columns, dryRun, rawOutput },
    } = await this.parse(GridFreeze);

    this.start(dryRun ? 'Previewing freeze' : 'Freezing panes');
    const result = await this.gsheet.setFrozen({ worksheetTitle, rows, columns, dryRun }, spreadsheetId);
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: freeze rows=${rows ?? '-'} columns=${columns ?? '-'}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      const parts = [rows !== undefined ? `${rows} row(s)` : null, columns !== undefined ? `${columns} column(s)` : null]
        .filter(Boolean)
        .join(', ');
      this.log(`Froze ${parts} in "${worksheetTitle}"`);
    }

    return result;
  }
}
