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

export default class GridResize extends Command {
  static description =
    'Resize rows or columns to an explicit pixel size, or auto-size them to their content. ' +
    'With --workbook the resize runs on a local XLSX file: pixels are converted into Excel units ' +
    '(columns: width characters; rows: height points) and --auto sizes columns to their longest cell text (capped) / resets rows to the default height.';

  static examples = [
    `$ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --count=4 --pixels=120
`,
    `$ gsheet grid:resize --spreadsheetId=<id> --worksheetTitle=T1 --dimension=COLUMNS --start=1 --auto
`,
    `$ gsheet grid:resize --workbook=template.xlsx -t Sheet1 --dimension=COLUMNS --start=1 --count=4 --auto --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    ...dimensionFlags,
    pixels: Flags.integer({ description: 'Explicit pixel size', required: false, exclusive: ['auto'] }),
    auto: Flags.boolean({ description: 'Auto-size to content', required: false, exclusive: ['pixels'] }),
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
        pixels,
        auto,
        dryRun,
        rawOutput,
      },
    } = await this.parse(GridResize);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'grid:resize'
    );

    if (target) {
      const result = target.workbook.resizeGrid({
        worksheetTitle: target.worksheetTitle,
        dimension: dimension as 'ROWS' | 'COLUMNS',
        start,
        count,
        pixels,
        auto,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        const planned = result.resized.map((r) => `${r.before ?? '?'} -> ${r.after}`).join(', ');
        this.log(
          `Dry run: resize ${result.count} ${result.dimension.toLowerCase()} at ${result.start} in "${result.sheet}" (${workbook}) [${result.unit}]: ${planned}`
        );
      } else {
        const resized = result.resized.map((r) => `${r.index}: ${r.after}`).join(', ');
        this.log(
          `Resized ${result.count} ${result.dimension.toLowerCase()} at ${result.start} in "${result.sheet}" (${auto ? 'auto' : `${pixels}px`}) [${result.unit}] -> ${saved?.savedPath}: ${resized}`
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

    this.start(dryRun ? 'Previewing resize' : `Resizing ${count} ${dimension.toLowerCase()}`);
    const result = await this.gsheet.mutateDimension(
      'resize',
      {
        worksheetTitle,
        dimension: dimension as GoogleSheetCli.Dimension,
        start,
        count,
        pixels,
        autoResize: auto,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: resize ${count} ${dimension.toLowerCase()} at ${start} (${auto ? 'auto' : `${pixels}px`})`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`Resized ${count} ${dimension.toLowerCase()} at ${start} in "${worksheetTitle}" (${auto ? 'auto' : `${pixels}px`})`);
    }

    return result;
  }
}
