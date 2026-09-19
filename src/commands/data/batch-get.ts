import { Flags } from '@oclif/core';
import Command, { spreadsheetId } from '../../lib/base-class';
import { parseRangesFlag } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class BatchGetData extends Command {
  static description = 'Fetch cell data across multiple ranges in a single batch request';

  static examples = [
    `$ gsheet data:batch-get --spreadsheetId=<spreadsheetId> --ranges='["Sheet1!A1:B10", "Sheet2!C1:D5"]'
`,
    `$ gsheet data:batch-get --spreadsheetId=<spreadsheetId> --ranges='["Sheet1!A1:B10"]' --valueRenderOption=UNFORMATTED_VALUE
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    ranges: Flags.string({
      description: 'JSON array of A1 range strings to query (e.g. \'["Sheet1!A1:B10", "Sheet2!C1:D5"]\')',
      required: true,
    }),
    valueRenderOption: Flags.string({
      description: 'Determines how values should be rendered in the output',
      options: [
        GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE,
        GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE,
        GoogleSheetCli.ValueRenderOption.FORMULA,
      ],
      default: GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE,
      required: false,
    }),
    dateTimeRenderOption: Flags.string({
      description: 'Determines how dates, times, and durations should be rendered',
      options: [
        GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING,
        GoogleSheetCli.DateTimeRenderOption.SERIAL_NUMBER,
      ],
      default: GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING,
      required: false,
    }),
    chunkSize: Flags.integer({
      description: 'Maximum number of ranges per batch chunk (default: 50)',
      required: false,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, ranges, valueRenderOption, dateTimeRenderOption, chunkSize, rawOutput },
    } = await this.parse(BatchGetData);

    const parsedRanges = parseRangesFlag(ranges);

    this.start(`Fetching ${parsedRanges.length} range(s) in batch`);
    const results = await this.gsheet.getDataBatch(
      parsedRanges,
      {
        valueRenderOption: valueRenderOption as GoogleSheetCli.ValueRenderOption,
        dateTimeRenderOption: dateTimeRenderOption as GoogleSheetCli.DateTimeRenderOption,
        chunkSize,
      },
      spreadsheetId
    );
    this.stop();

    const totalCells = results.reduce((sum, r) => {
      return sum + r.values.reduce((rowSum, row) => rowSum + row.length, 0);
    }, 0);

    const result = {
      operation: this.id,
      spreadsheetId,
      rangeCount: results.length,
      totalCells,
      ranges: results,
    };

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      this.log(`Successfully fetched ${results.length} range(s) (${totalCells} total cells) from spreadsheet "${spreadsheetId}"`);
      for (const r of results) {
        this.log(`  - ${r.range}: ${r.values.length} row(s)`);
      }
    }

    return result;
  }
}
