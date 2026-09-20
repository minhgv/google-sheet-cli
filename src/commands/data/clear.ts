import { Flags } from '@oclif/core';
import Command, { spreadsheetId, worksheetTitle } from '../../lib/base-class';

export default class ClearData extends Command {
  static description = 'Clear cell values in an explicitly bounded range, keeping every other cell property (values-only clear)';

  static examples = [
    `$ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20"

Successfully cleared 76 cells in "Sheet1" ("Sheet1!A2:D20")
`,
    `$ gsheet data:clear --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --range="Sheet1!A2:D20" --dryRun

Clear preview (dry run): 76 cells in "Sheet1!A2:D20" would be cleared
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle,
    spreadsheetId,
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
      flags: { range, dryRun, overwriteFormulas, worksheetTitle, spreadsheetId, rawOutput },
    } = await this.parse(ClearData);

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
