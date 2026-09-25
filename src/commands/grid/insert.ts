import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';
import { XlsxDimension } from '../../lib/xlsx-types';

export const dimensionFlags = {
  dimension: Flags.string({
    description: 'The dimension to mutate',
    options: ['ROWS', 'COLUMNS'],
    required: true,
  }),
  start: Flags.integer({ description: 'The 1-based first row/column index to affect', required: true }),
  count: Flags.integer({ description: 'How many rows/columns the operation covers', default: 1, required: false }),
  dryRun: Flags.boolean({ description: 'Preview the mutation without applying it', required: false }),
};

export default class GridInsert extends Command {
  static description =
    'Insert rows or columns into a worksheet at a position. ' +
    'Existing data at and after the position shifts down/right. ' +
    'With --workbook the insert runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=5 --count=2
`,
    `$ gsheet grid:insert --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=3 --inheritFromBefore
`,
    `$ gsheet grid:insert --workbook=template.xlsx --worksheetTitle="Functional effort" --dimension=ROWS --start=8 --count=3 --inPlace
`,
    `$ gsheet grid:insert --workbook=template.xlsx --dimension=ROWS --start=8 --count=2 --updateRefs --inPlace
`,
    `$ gsheet grid:insert --workbook=template.xlsx -t Sheet1 --dimension=ROWS --start=8 --count=3 --dryRun
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    ...dimensionFlags,
    inheritFromBefore: Flags.boolean({
      description: 'Inherit formatting from the row/column before instead of after (matches "insert above/left")',
      required: false,
    }),
    force: Flags.boolean({
      description: 'Local backend only: adjust merged ranges that intersect the insertion boundary instead of refusing',
      required: false,
    }),
    updateRefs: Flags.boolean({
      description:
        'Local backend only: rewrite same-sheet formula references and defined names affected by the insert (refs inside a deleted range would become #REF!); cross-sheet references are left untouched',
      required: false,
    }),
  };

  async run() {
    const {
      flags: {
        spreadsheetId,
        worksheetTitle,
        workbook,
        output,
        inPlace,
        discardUnsupported,
        dimension,
        start,
        count,
        inheritFromBefore,
        force,
        updateRefs,
        dryRun,
        rawOutput,
      },
    } = await this.parse(GridInsert);

    if (updateRefs && spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.USAGE,
        '--updateRefs only applies to the local --workbook backend and cannot be combined with --spreadsheetId.'
      );
    }

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'grid:insert'
    );

    if (target) {
      const result = target.workbook.splice('insert', {
        worksheetTitle: target.worksheetTitle,
        dimension: dimension as XlsxDimension,
        start,
        count,
        inheritFromBefore,
        force,
        updateRefs,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: insert ${count} ${dimension.toLowerCase()} at ${start} in "${result.sheet}" (${workbook})`);
        if (result.mergesAdjusted.length > 0) {
          this.log(`  - Merges adjusted: ${result.mergesAdjusted.map((m) => `${m.before} -> ${m.after ?? 'dropped'}`).join(', ')}`);
        }
      } else {
        this.log(`Inserted ${count} ${dimension.toLowerCase()} at ${start} in "${result.sheet}" -> ${saved?.savedPath}`);
        if (result.mergesAdjusted.length > 0) {
          this.log(`  - Merges adjusted: ${result.mergesAdjusted.map((m) => `${m.before} -> ${m.after ?? 'dropped'}`).join(', ')}`);
        }
        if (updateRefs) {
          this.log(`  - References updated: ${result.refsRewritten} rewritten, ${result.refsBroken} now #REF!`);
        }
      }
      for (const w of result.warnings) this.warn(w);
      return receipt;
    }

    if (!spreadsheetId || !worksheetTitle) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId with --worksheetTitle for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

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
