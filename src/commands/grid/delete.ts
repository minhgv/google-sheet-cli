import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { Flags } from '@oclif/core';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';
import { XlsxDimension } from '../../lib/xlsx-types';
import { dimensionFlags } from './insert';

export default class GridDelete extends Command {
  static description =
    'Delete rows or columns from a worksheet. ' +
    'Data after the deleted range shifts up/left. --dryRun previews the values about to be removed. ' +
    'With --workbook the delete runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=10 --count=3
`,
    `$ gsheet grid:delete --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=5 --dryRun
`,
    `$ gsheet grid:delete --workbook=template.xlsx -t "Functional effort" --dimension=ROWS --start=8 --count=5 --inPlace
`,
    `$ gsheet grid:delete --workbook=template.xlsx --dimension=ROWS --start=8 --count=2 --updateRefs --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    ...dimensionFlags,
    force: Flags.boolean({
      description: 'Local backend only: adjust merged ranges that intersect the deleted span instead of refusing',
      required: false,
    }),
    updateRefs: Flags.boolean({
      description:
        'Local backend only: rewrite same-sheet formula references and defined names affected by the delete (refs fully inside the deleted span become #REF!); cross-sheet references are left untouched',
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
        force,
        updateRefs,
        dryRun,
        rawOutput,
      },
    } = await this.parse(GridDelete);

    if (updateRefs && spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.USAGE,
        '--updateRefs only applies to the local --workbook backend and cannot be combined with --spreadsheetId.'
      );
    }

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'grid:delete'
    );

    if (target) {
      const result = target.workbook.splice('delete', {
        worksheetTitle: target.worksheetTitle,
        dimension: dimension as XlsxDimension,
        start,
        count,
        force,
        updateRefs,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: would delete ${count} ${dimension.toLowerCase()} at ${start} in "${result.sheet}" (${workbook})`);
        if (result.removedValues && result.removedValues.length > 0) {
          this.log('Values about to be removed:');
          this.log(JSON.stringify(result.removedValues, null, 2));
        } else {
          this.log('(no values in the affected range)');
        }
      } else {
        this.log(`Deleted ${count} ${dimension.toLowerCase()} at ${start} from "${result.sheet}" -> ${saved?.savedPath}`);
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
