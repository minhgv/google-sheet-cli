import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class ClearData extends Command {
  static description =
    'Clear cell values in an explicitly bounded range, keeping every other cell property (values-only clear). ' +
    'With --workbook the clear runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20"

Successfully cleared 76 cells in "Sheet1" ("Sheet1!A2:D20")
`,
    `$ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20" --dryRun

Clear preview (dry run): 76 cells in "Sheet1!A2:D20" would be cleared
`,
    `$ gsheet data:clear --workbook=template.xlsx --range="Functional effort!A8:J30" --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle: optionalWorksheetTitle,
    spreadsheetId: optionalSpreadsheetId,
    ...workbookTargetFlags,
    range: Flags.string({
      description: 'A1 range to clear; must be bounded on both axes (e.g. "Sheet1!A1:D20")',
      required: true,
    }),
    dryRun: Flags.boolean({
      description: 'Preview the clear without modifying the worksheet',
      default: false,
      required: false,
    }),
    overwriteFormulas: Flags.boolean({
      description: 'Allow the clear to overwrite cells that contain formulas',
      default: false,
      required: false,
    }),
  };

  async run() {
    const {
      flags: { range, dryRun, overwriteFormulas, worksheetTitle, spreadsheetId, workbook, output, inPlace, discardUnsupported, rawOutput },
    } = await this.parse(ClearData);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'data:clear'
    );

    if (target) {
      const result = target.workbook.clearRange({
        worksheetTitle: target.worksheetTitle,
        range,
        overwriteFormulas,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        const formulas = result.formulasOverwritten.length
          ? `; ${result.formulasOverwritten.length} formula cell(s) in the range would refuse the clear without --overwriteFormulas`
          : '';
        this.log(`Clear preview (dry run): ${result.cellsCleared} cells in "${result.range}" would be cleared${formulas}`);
      } else {
        this.log(`Successfully cleared ${result.cellsCleared} cells in "${result.sheet}" ("${result.range}") -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    if (!spreadsheetId || !worksheetTitle) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId with --worksheetTitle for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    this.start(dryRun ? 'Previewing clear' : 'Clearing data');
    const receipt = await this.gsheet.clearData({ range, dryRun, overwriteFormulas, worksheetTitle }, spreadsheetId);
    this.stop();

    const result = { operation: this.id, ...receipt };

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      const formulas = receipt.formulasOverwritten?.length
        ? `; ${receipt.formulasOverwritten.length} formula cell(s) in the range would refuse the clear without --overwriteFormulas`
        : '';
      this.log(`Clear preview (dry run): ${receipt.cellsCleared} cells in "${receipt.range}" would be cleared${formulas}`);
    } else {
      this.log(`Successfully cleared ${receipt.cellsCleared} cells in "${receipt.worksheetTitle}" ("${receipt.range}")`);
    }

    return result;
  }
}
