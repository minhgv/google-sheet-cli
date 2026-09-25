import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class Add extends Command {
  static description =
    'Add a worksheet with the specified title to the spreadsheet. ' +
    'With --workbook the sheet is added to a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet worksheet:add --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle>

Worksheet "<worksheetTitle>" (<id>) successfully created
`,
    `$ gsheet worksheet:add --workbook=template.xlsx --worksheetTitle=Notes --inPlace
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
    } = await this.parse(Add);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'worksheet:add'
    );

    if (target) {
      const result = target.workbook.addSheet(worksheetTitle, { dryRun });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: worksheet "${result.sheet}" would be created in (${workbook})`);
      } else {
        this.log(`Worksheet "${result.sheet}" successfully created -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    if (!spreadsheetId) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    this.start('Adding worksheet');
    const worksheet = await this.gsheet.addWorksheet(worksheetTitle, spreadsheetId);
    if (!worksheet) throw new Error('Worksheet not created');
    const { properties: { title = '', sheetId = '' } = {} } = worksheet;
    this.stop();
    const result = { operation: this.id, ...worksheet };
    this.logRaw(`Worksheet "${title}" (${sheetId}) successfully created`, result);
    return result;
  }
}
