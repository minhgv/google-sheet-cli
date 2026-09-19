import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { table, tableFlags } from '../../lib/table';

export default class FindData extends Command {
  static description =
    'Locate cells matching a condition and return their A1 coordinates. ' +
    'Pure read: scans the range once and reports where matches live, so follow-up ' +
    'writes can target exact cells without downloading the whole sheet.';

  static examples = [
    `$ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --equals="NV0123" --column=B --first --rawOutput
`,
    `$ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --contains="invoice" --byRow
`,
    `$ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --regex="^ERR-" --range=A1:K200
`,
    `$ gsheet data:find --spreadsheetId=<id> --worksheetTitle=T1 --equals="Paid" --header="Status"
`,
  ];

  static flags = {
    ...Command.flags,
    ...tableFlags(),
    spreadsheetId,
    worksheetTitle,
    range: Flags.string({ description: 'The A1 range bounding the scan (default: whole worksheet)', required: false }),
    equals: Flags.string({
      description: 'Match cells exactly equal to this value',
      required: false,
      exclusive: ['contains', 'regex'],
    }),
    contains: Flags.string({
      description: 'Match cells containing this substring',
      required: false,
      exclusive: ['equals', 'regex'],
    }),
    regex: Flags.string({
      description: 'Match cells against this regular expression',
      required: false,
      exclusive: ['equals', 'contains'],
    }),
    column: Flags.string({
      description: 'Restrict matches to one column by A1 letter (e.g. "B")',
      required: false,
      exclusive: ['header'],
    }),
    header: Flags.string({
      description: 'Restrict matches to the column whose first scanned row equals this header text',
      required: false,
      exclusive: ['column'],
    }),
    ignoreCase: Flags.boolean({ description: 'Case-insensitive matching', default: true, required: false, allowNo: true }),
    valueRenderOption: Flags.string({
      description: 'Determines how cell values are rendered before matching',
      options: [
        GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE,
        GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE,
        GoogleSheetCli.ValueRenderOption.FORMULA,
      ],
      default: GoogleSheetCli.ValueRenderOption.FORMATTED_VALUE,
      required: false,
    }),
    dateTimeRenderOption: Flags.string({
      description: 'Determines how dates, times, and durations are rendered before matching',
      options: [GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING, GoogleSheetCli.DateTimeRenderOption.SERIAL_NUMBER],
      default: GoogleSheetCli.DateTimeRenderOption.FORMATTED_STRING,
      required: false,
    }),
    limit: Flags.integer({ description: 'Maximum matches to return (matchCount still reports the true total)', default: 100, required: false }),
    first: Flags.boolean({ description: 'Return only the first match (alias for --limit=1)', required: false }),
    byRow: Flags.boolean({ description: 'Collapse matches to unique rows and include full row values', required: false }),
  };

  async run() {
    const {
      flags: {
        spreadsheetId,
        worksheetTitle,
        range,
        equals,
        contains,
        regex,
        column,
        header,
        ignoreCase,
        valueRenderOption,
        dateTimeRenderOption,
        limit,
        first,
        byRow,
        rawOutput,
        ...tableOptions
      },
    } = await this.parse(FindData);

    this.start('Scanning for matches');
    const result = await this.gsheet.findData(
      {
        worksheetTitle,
        range,
        equals,
        contains,
        regex,
        column,
        header,
        ignoreCase,
        valueRenderOption: valueRenderOption as GoogleSheetCli.ValueRenderOption,
        dateTimeRenderOption: dateTimeRenderOption as GoogleSheetCli.DateTimeRenderOption,
        limit: first ? 1 : limit,
        byRow,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
      return result;
    }

    if (result.matches.length === 0) {
      this.log(`No matches in ${result.range}`);
      return result;
    }

    const rows = result.matches.map((match) => ({ ...match }));
    const columns: Record<string, object> = {
      a1: {},
      row: {},
      column: {},
      value: {},
    };
    if (byRow) {
      columns.rowValues = { get: (match: Record<string, unknown>) => JSON.stringify(match.rowValues ?? []) };
    }
    table(rows, columns, tableOptions);
    if (result.truncated) {
      this.log(`... ${result.matchCount - result.matches.length} more match(es) truncated by --limit`);
    }
    return result;
  }
}
