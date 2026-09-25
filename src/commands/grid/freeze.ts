import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class GridFreeze extends Command {
  static description =
    'Freeze or unfreeze rows and columns on a worksheet. ' +
    'At least one of --rows/--columns is required; 0 unfreezes that axis. ' +
    'With --workbook the freeze runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1
`,
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=1 --columns=2
`,
    `$ gsheet grid:freeze --spreadsheetId=<id> --worksheetTitle=Report --rows=0 --columns=0
`,
    `$ gsheet grid:freeze --workbook=template.xlsx -t Sheet1 --rows=1 --columns=2 --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    rows: Flags.integer({ description: 'Number of rows to freeze (0 unfreezes)', required: false }),
    columns: Flags.integer({ description: 'Number of columns to freeze (0 unfreezes)', required: false }),
    dryRun: Flags.boolean({ description: 'Preview the mutation without applying it', required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, workbook, output, inPlace, discardUnsupported, rows, columns, dryRun, rawOutput },
    } = await this.parse(GridFreeze);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'grid:freeze'
    );

    if (target) {
      const result = target.workbook.freezePanes({
        worksheetTitle: target.worksheetTitle,
        rows,
        columns,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: freeze rows=${result.rows} columns=${result.columns} in "${result.sheet}" (${workbook}); ${result.model}`);
      } else {
        const parts = [
          result.rows > 0 ? `${result.rows} row(s)` : null,
          result.columns > 0 ? `${result.columns} column(s)` : null,
        ]
          .filter(Boolean)
          .join(', ');
        this.log(
          `${result.frozen ? `Froze ${parts}` : 'Unfroze panes'} in "${result.sheet}" -> ${saved?.savedPath}`
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

    this.start(dryRun ? 'Previewing freeze' : 'Freezing panes');
    const result = await this.gsheet.setFrozen({ worksheetTitle, rows, columns, dryRun }, spreadsheetId);
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: freeze rows=${rows ?? '-'} columns=${columns ?? '-'}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      const parts = [rows !== undefined ? `${rows} row(s)` : null, columns !== undefined ? `${columns} column(s)` : null]
        .filter(Boolean)
        .join(', ');
      this.log(`Froze ${parts} in "${worksheetTitle}"`);
    }

    return result;
  }
}
