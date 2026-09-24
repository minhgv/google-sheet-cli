import { Flags } from '@oclif/core';
import Command, {
  optionalSpreadsheetId,
  optionalWorksheetTitle,
  workbookTargetFlags,
} from '../../lib/base-class';
import { GSheetError, GSheetErrorCode } from '../../lib/cli-errors';
import { resolveWorkbookTarget, saveWorkbookTarget } from '../../lib/xlsx-target';

export default class FormatMerge extends Command {
  static description =
    'Merge or unmerge cells over a bounded range. ' +
    'Merging keeps the top-left value; other values in the range are hidden. ' +
    'With --workbook the merge runs on a local XLSX file instead of Google Sheets.';

  static examples = [
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1
`,
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:C3 --type=MERGE_ROWS
`,
    `$ gsheet format:merge --spreadsheetId=<id> --worksheetTitle=Report --range=A1:J1 --unmerge
`,
    `$ gsheet format:merge --workbook=template.xlsx -t "Functional effort" --range=B8:B12 --inPlace
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId: optionalSpreadsheetId,
    worksheetTitle: optionalWorksheetTitle,
    ...workbookTargetFlags,
    range: Flags.string({ description: 'The A1 range to merge or unmerge (e.g. "A1:J1")', required: true }),
    type: Flags.string({
      description: 'Merge type',
      options: ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS'],
      default: 'MERGE_ALL',
      required: false,
      exclusive: ['unmerge'],
    }),
    unmerge: Flags.boolean({ description: 'Unmerge previously merged cells in the range', required: false, exclusive: ['type'] }),
    dryRun: Flags.boolean({ description: 'Preview the mutation without applying it', required: false }),
  };

  async run() {
    const {
      flags: { spreadsheetId, worksheetTitle, workbook, output, inPlace, discardUnsupported, range, type, unmerge, dryRun, rawOutput },
    } = await this.parse(FormatMerge);

    const target = await resolveWorkbookTarget(
      { workbook, spreadsheetId, worksheetTitle, output, inPlace, discardUnsupported, dryRun },
      this.id ?? 'format:merge'
    );

    if (target) {
      const result = target.workbook.mergeCellsRange({
        worksheetTitle: target.worksheetTitle,
        range,
        mergeType: type as 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS',
        unmerge,
        dryRun,
      });
      const saved = await saveWorkbookTarget(target, { workbook, output, inPlace, discardUnsupported, dryRun });
      const receipt = { ...result, operation: this.id, saved };

      if (rawOutput) {
        this.logRaw('', receipt);
      } else if (dryRun) {
        this.log(`Dry run: ${unmerge ? 'unmerge' : 'merge'} ${result.affected.join(', ')} in "${result.sheet}" (${workbook})`);
      } else {
        this.log(`${unmerge ? 'Unmerged' : 'Merged'} ${result.affected.join(', ')} in "${result.sheet}" -> ${saved?.savedPath}`);
      }
      return receipt;
    }

    if (!spreadsheetId || !worksheetTitle) {
      throw new GSheetError(
        GSheetErrorCode.VALIDATION,
        'No target given. Pass --spreadsheetId with --worksheetTitle for Google Sheets, or --workbook for a local XLSX file.'
      );
    }

    this.start(dryRun ? 'Previewing merge' : unmerge ? 'Unmerging cells' : 'Merging cells');
    const result = await this.gsheet.setMerge(
      {
        worksheetTitle,
        range,
        mergeType: type as 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS',
        unmerge,
        dryRun,
      },
      spreadsheetId
    );
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Dry run: ${result.fields.join(', ')} over ${range}`);
      this.log(JSON.stringify(result.requests, null, 2));
    } else {
      this.log(`${unmerge ? 'Unmerged' : 'Merged'} ${range}`);
    }

    return result;
  }
}
