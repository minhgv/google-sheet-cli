import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { discoverSchema } from '../../lib/data-schema';
import { formatBoundedA1Range } from '../../lib/sheet-batch';
import { table, tableFlags } from '../../lib/table';

export default class GetSchema extends Command {
  static description = `Discover the column schema of a worksheet (read-only)

Reports the actual header row with its absolute A1 coordinates, inferred cell types,
formula presence, data-validation rules and named ranges over a bounded sample. The
header row is the first row of the sampled range. Inferred types describe the sample,
not an authoritative schema - Google stores dates as numbers and formatting decides
what they look like.`;

  static examples = [
    `$ gsheet data:schema --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>

HEADER  COLUMN  TYPES               FORMULAS  VALIDATIONS
Name    A       string (3)          0
Amount  B       number (2), string  0
Flag    C       boolean (3)         0         ONE_OF_LIST (3)
`,
  ];

  static flags = {
    ...Command.flags,
    ...tableFlags(),
    spreadsheetId,
    worksheetTitle,
    minRow: Flags.integer({ description: 'The first row of the sampled range (holds the header row)', default: 1, required: false }),
    minCol: Flags.integer({ description: 'The first column of the sampled range', default: 1, required: false }),
    maxRow: Flags.integer({ description: 'The last row of the sampled range', default: 100, required: false }),
    maxCol: Flags.integer({ description: 'The last column of the sampled range', default: 26, required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, rawOutput, worksheetTitle, minRow, minCol, maxRow, maxCol, ...tableOptions },
    } = await this.parse(GetSchema);

    if (minRow < 1 || minCol < 1) throw new GSheetError(GSheetErrorCode.VALIDATION, '--minRow and --minCol must be at least 1');
    if (maxRow < minRow) throw new GSheetError(GSheetErrorCode.VALIDATION, `--maxRow (${maxRow}) must not be smaller than --minRow (${minRow})`);
    if (maxCol < minCol) throw new GSheetError(GSheetErrorCode.VALIDATION, `--maxCol (${maxCol}) must not be smaller than --minCol (${minCol})`);

    this.start('Discovering schema');
    const metadata = await this.gsheet.getWorksheetMetadata(
      { worksheetTitle, range: formatBoundedA1Range(undefined, minCol, minRow, maxCol, maxRow) },
      spreadsheetId
    );
    const discovery = discoverSchema(metadata);
    const result = {
      operation: this.id,
      spreadsheetId: discovery.spreadsheetId,
      worksheetTitle: discovery.worksheetTitle,
      sheetId: discovery.sheetId,
      gridRows: discovery.gridRows,
      gridColumns: discovery.gridColumns,
      sample: discovery.sample,
      columns: discovery.columns,
      warnings: discovery.warnings,
      namedRanges: discovery.namedRanges,
    };

    if (rawOutput) {
      this.logRaw('', result);
      return result;
    }

    this.stop();
    const formatted = discovery.columns.map((column) => ({
      header: column.header,
      column: column.letter,
      headerCell: column.headerA1,
      types: Object.entries(column.inferredTypes)
        .filter(([, count]) => count > 0)
        .map(([kind, count]) => `${kind} (${count})`)
        .join(', ') || '-',
      formulas: String(column.formulaCells),
      validations: column.validations.map(({ condition, count }) => `${condition} (${count})`).join(', ') || '-',
    }));
    table(formatted, { header: {}, column: {}, headerCell: {}, types: {}, formulas: {}, validations: {} }, tableOptions);
    for (const warning of discovery.warnings) this.log(`Warning: ${warning}`);
    for (const namedRange of discovery.namedRanges) this.log(`Named range: ${namedRange.name} = ${namedRange.range}`);
    return result;
  }
}
