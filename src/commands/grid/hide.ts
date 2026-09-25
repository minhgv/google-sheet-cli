import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';
import { dimensionFlags } from './insert';

export default class GridHide extends Command {
  static description =
    'Hide or unhide rows or columns on a worksheet. ' +
    'With --workbook the hide runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=ROWS --start=8 --count=3
`,
    `$ gsheet grid:hide --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=2 --unhide
`,
    `$ gsheet grid:hide --workbook=template.xlsx -t Sheet1 --dimension=COLUMNS --start=2 --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    ...dimensionFlags,
    unhide: Flags.boolean({ description: 'Unhide instead of hide', required: false }),
    dryRun: Flags.boolean({ description: 'Preview the mutation without applying it', required: false }),
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
        unhide,
        dryRun,
        rawOutput,
      },
    } = await this.parse(GridHide);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'grid:hide'
    );

    if (target) {
      const result = target.workbook.setGridHidden({
        worksheetTitle: target.worksheetTitle,
        dimension: dimension as 'ROWS' | 'COLUMNS',
        start,
        count,
        unhide,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: ${result.hidden ? 'hide' : 'unhide'} ${result.count} ${result.dimension.toLowerCase()} at ${result.start} in "${result.sheet}" (${workbook}); ${result.model}`);
      } else {
        this.log(
          `${unhide ? 'Unhid' : 'Hid'} ${result.count} ${result.dimension === 'ROWS' ? 'row(s)' : 'column(s)'} at ${result.start} in "${result.sheet}" -> ${saved?.savedPath}`
        );
      }
      return receipt;
    }

    if (!spreadsheetId || !worksheetTitle) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId with --worksheetTitle for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    const verb = unhide ? 'Unhiding' : 'Hiding';
    this.start(dryRun ? `Previewing ${verb.toLowerCase()}` : `${verb} ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'hide',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        unhide,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: ${verb.toLowerCase()} ${count} ${dimension.toLowerCase()} at ${start}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`${unhide ? 'Unhid' : 'Hid'} ${count} ${dimension === 'ROWS' ? 'row(s)' : 'column(s)'} at ${start} in "${worksheetTitle}"`);
    }

    return result;
  }
}
