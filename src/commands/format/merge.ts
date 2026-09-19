import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';

export default class FormatMerge extends Command {
  static description =
    'Merge or unmerge cells over a bounded range. ' +
    'Merging keeps the top-left value; other values in the range are hidden by Google.';

  static examples = [
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1
`,
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:C3 --type=MERGE_ROWS
`,
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --unmerge
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    worksheetTitle,
    range: Flags.string({ description: 'The A1 range to merge or unmerge (e.g. "A1:J1")', required: true }),
    type: Flags.string({
      description: 'Merge type',
      options: ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'],
      default: 'MERGE_ALL',
      required: false,
      exclusive: ['unmerge'],
    }),
    unmerge: Flags.boolean({ description: 'Unmerge previously merged cells in the range', required: false, exclusive: ['type'] }),
    dryRun: Flags.boolean({ description: 'Preview the batchUpdate request without applying it', required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, range, type, unmerge, dryRun, rawOutput },
    } = await this.parse(FormatMerge);

    this.start(dryRun ? 'Previewing merge' : unmerge ? 'Unmerging cells' : 'Merging cells');
    const result = await this.gsheet.setMerge(
      {
        worksheetTitle,
        range,
        mergeType: type as 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS',
        unmerge,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: ${result.fields.join(', ')} over ${range}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`${unmerge ? 'Unmerged' : 'Merged'} ${range}`);
    }

    return result;
  }
}
