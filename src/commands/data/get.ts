import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { table, tableFlags } from '../../lib/table';
export default class GetData extends Command {
  static description = 'Returns cell data';

  static examples = [
    `$ gsheet data:get --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>

(a)  (b)  (c)
A1   B1   C1
A2   B2   C2
A3   B3   C3
`,
  ];

  static flags = {
    ...Command.flags,
    ...tableFlags(),
    spreadsheetId,
    worksheetTitle,
    hasHeaderRow: Flags.boolean({ char: 'w', description: 'If the first row should be treated as header row', default: false, required: false }),
    range: Flags.string({ description: 'The range to use to query the cells', required: false }),
    minRow: Flags.integer({ description: 'The optional starting row of the operation', default: 1, required: false }),
    minCol: Flags.integer({ description: 'The optional starting col of the operation', default: 1, required: false }),
    maxRow: Flags.integer({ description: 'The optional ending row of the operation', required: false }),
    maxCol: Flags.integer({ description: 'The optional ending col of the operation', required: false }),
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
  };

  async run() {
    const {
      flags: {
        spreadsheetId,
        rawOutput,
        minRow,
        maxRow,
        minCol,
        maxCol,
        range,
        hasHeaderRow,
        worksheetTitle,
        valueRenderOption,
        dateTimeRenderOption,
        ...tableOptions
      },
    } = await this.parse(GetData);

    this.start('Fetching data');
    const res = await this.gsheet.getData(
      {
        minRow,
        maxRow,
        minCol,
        maxCol,
        range,
        hasHeaderRow,
        worksheetTitle,
        valueRenderOption: valueRenderOption as GoogleSheetCli.ValueRenderOption,
        dateTimeRenderOption: dateTimeRenderOption as GoogleSheetCli.DateTimeRenderOption,
      },
      spreadsheetId
    );
    const result = { operation: this.id, ...res };

    if (rawOutput) {
      this.logRaw('', result);
      return result;
    }

    const { header, formatted } = res;
    const columns = header.reduce((red, col) => {
      return { ...red, [col]: {} };
    }, {});
    this.stop();
    table(formatted, columns, tableOptions);
    return result;
  }
}
