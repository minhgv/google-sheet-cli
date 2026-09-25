import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class Remove extends Command {
  static description =
    'Remove a worksheet with the specified title from the spreadsheet. ' +
    'With --workbook the sheet is removed from a local XLSX file instead; the last visible sheet cannot be removed.';

  static examples = [
    `$ gsheet worksheet:remove --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>

Worksheet "<worksheetTitle>" successfully removed
`,
    `$ gsheet worksheet:remove --workbook=template.xlsx --worksheetTitle=Draft --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle: optionalWorksheetTitle,
    spreadsheetId: optionalSpreadsheetId,
    ...workbookTargetFlags,
  };

  async run() {
    const {
      flags: { worksheetTitle = '', spreadsheetId, workbook, output, inPlace, discardUnsupported, dryRun, rawOutput },
    } = await this.parse(Remove);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'worksheet:remove'
    );

    if (target) {
      const result = target.workbook.removeSheet(worksheetTitle, { dryRun });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: worksheet "${result.sheet}" would be removed from (${workbook})`);
      } else {
        this.log(`Worksheet "${result.sheet}" successfully removed -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    if (!spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    this.start('Removing worksheet');
    await this.gsheet.removeWorksheet(worksheetTitle, spreadsheetId);
    this.stop();
    const result = { operation: this.id, worksheetTitle };
    this.logRaw(`Worksheet "${worksheetTitle}" successfully removed`, result);
    return result;
  }
}
