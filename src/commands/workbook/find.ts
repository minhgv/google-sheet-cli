import { Command, Flags } from '@oclif/core';
import { table, tableFlags } from '../../lib/table';
import { XlsxWorkbook } from '../../lib/xlsx';

export default class WorkbookFind extends Command {
  static description =
    'Locate cells matching a condition in a local XLSX workbook and return their A1 coordinates. ' +
    'Same match semantics and result shape as data:find, so follow-up writes can target exact ' +
    'cells without reading the whole sheet. Pure read: no network, no auth, no writes.';

  static examples = [
    `$ gsheet workbook:find --file=template.xlsx --equals="Tổng cộng" --first --rawOutput
`,
    `$ gsheet workbook:find --file=template.xlsx --worksheetTitle="Functional effort" --contains="MODULE" --byRow
`,
    `$ gsheet workbook:find --file=report.xlsx --regex="^ERR-" --range="Data!A1:K200"
`,
  ];

  static flags = {
    help: Flags.help({ char: 'h' }),
    ...tableFlags(),
    file: Flags.string({
      char: 'f',
      description: 'Path to the local XLSX file to scan',
      required: true,
    }),
    worksheetTitle: Flags.string({
      char: 't',
      description: 'Title of the worksheet to scan (default: first sheet)',
      required: false,
      env: 'WORKSHEET_TITLE',
    }),
    range: Flags.string({ description: 'The A1 range bounding the scan (default: used range of the sheet)', required: false }),
    equals: Flags.string({ description: 'Match cells whose rendered value equals this string', required: false, exclusive: ['contains', 'regex'] }),
    contains: Flags.string({ description: 'Match cells whose rendered value contains this substring', required: false, exclusive: ['equals', 'regex'] }),
    regex: Flags.string({ description: 'Match cells whose rendered value matches this regular expression', required: false, exclusive: ['equals', 'contains'] }),
    column: Flags.string({ description: 'Restrict the scan to one column letter (e.g. "B")', required: false, exclusive: ['header'] }),
    header: Flags.string({ description: 'Restrict the scan to the column whose first-row header matches this value', required: false, exclusive: ['column'] }),
    ignoreCase: Flags.boolean({ description: 'Case-insensitive matching', default: true, required: false, allowNo: true }),
    limit: Flags.integer({ description: 'Maximum matches to return (matchCount still reports the true total)', default: 100, required: false }),
    first: Flags.boolean({ description: 'Return only the first match (alias for --limit=1)', required: false }),
    byRow: Flags.boolean({ description: 'Collapse matches to unique rows and include full row values', required: false }),
    rawOutput: Flags.boolean({ char: 'r', description: 'Get the raw output as a JSON string', default: false, required: false }),
  };

  async run() {
    const {
      flags: { file, worksheetTitle, range, equals, contains, regex, column, header, ignoreCase, limit, first, byRow, rawOutput, ...tableOptions },
    } = await this.parse(WorkbookFind);

    const workbook = await XlsxWorkbook.load(file);
    const result = workbook.find({
      worksheetTitle,
      range,
      equals,
      contains,
      regex,
      column,
      header,
      ignoreCase,
      limit: first ? 1 : limit,
      byRow,
    });

    const receipt = { operation: this.id, ...result };

    if (rawOutput) {
      this.log(JSON.stringify(receipt, null, 2));
      return receipt;
    }

    if (result.matches.length === 0) {
      this.log(`No matches in ${result.range}`);
      return receipt;
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
    return receipt;
  }
}
