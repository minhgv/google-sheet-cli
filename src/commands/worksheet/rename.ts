import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class Rename extends Command {
  static description =
    'Rename a worksheet. ' +
    'With --workbook the rename runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet worksheet:rename --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --newWorksheetTitle=<newWorksheetTitle>

Worksheet "<worksheetTitle>" successfully renamed to "<newWorksheetTitle>" 
`,
    `$ gsheet worksheet:rename --workbook=template.xlsx --worksheetTitle=Draft --newWorksheetTitle=Final --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle: optionalWorksheetTitle,
    newWorksheetTitle: Flags.string({
      description: 'New title of the worksheet to use',
      required: true,
    }),
    spreadsheetId: optionalSpreadsheetId,
    ...workbookTargetFlags,
  };

  async run() {
    const {
      flags: {
        worksheetTitle = '',
        newWorksheetTitle,
        spreadsheetId,
        workbook,
        output,
        inPlace,
        discardUnsupported,
        dryRun,
        rawOutput,
      },
    } = await this.parse(Rename);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'worksheet:rename'
    );

    if (target) {
      const result = target.workbook.renameSheet(worksheetTitle, newWorksheetTitle, { dryRun });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: worksheet "${result.from}" would be renamed to "${result.to}" in (${workbook})`);
      } else {
        this.log(`Worksheet "${result.from}" successfully renamed to "${result.to}" -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    if (!spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    this.start('Renaming worksheet');
    await this.gsheet.renameWorksheet(worksheetTitle, newWorksheetTitle, spreadsheetId);
    this.stop();
    const result = { operation: this.id };
    this.logRaw(`Worksheet "${worksheetTitle}" successfully renamed to "${newWorksheetTitle}"`, result);
    return result;
  }
}
