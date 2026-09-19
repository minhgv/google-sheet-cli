import { Flags } from '@oclif/core';
import Command, { optionalData, spreadsheetId, valueInputOption } from '../../lib/base-class';
import { resolveBatchUpdates } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class BatchUpdateData extends Command {
  static description = 'Update cell data across multiple ranges in a single batch request';

  static examples = [
    `$ gsheet data:batch-update --spreadsheetId=<spreadsheetId> '[{"range": "Sheet1!A1:B2", "values": [["1", "2"], ["3", "4"]]}]'
`,
    `$ gsheet data:batch-update --spreadsheetId=<spreadsheetId> --input=updates.json
`,
    `$ gsheet data:batch-update --spreadsheetId=<spreadsheetId> --input=- < updates.json
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    valueInputOption,
    input: Flags.string({
      char: 'i',
      description: 'Path to input JSON file or "-" for stdin',
      required: false,
    }),
    inputFormat: Flags.string({
      description: 'Format of input file (currently "json")',
      options: ['json'],
      default: 'json',
      required: false,
    }),
    dryRun: Flags.boolean({
      description: 'Preview batch update changes without modifying spreadsheet',
      default: false,
      required: false,
    }),
    overwriteFormulas: Flags.boolean({
      description: 'Allow overwriting existing formula cells',
      default: false,
      required: false,
    }),
    chunkByteSize: Flags.integer({
      description: 'Maximum byte size per payload chunk',
      required: false,
    }),
    maxRowsPerChunk: Flags.integer({
      description: 'Maximum rows per payload chunk',
      required: false,
    }),
  };

  static args = { data: optionalData };

  async run() {
    const {
      args: { data },
      flags: {
        spreadsheetId,
        valueInputOption,
        input,
        inputFormat,
        dryRun,
        overwriteFormulas,
        chunkByteSize,
        maxRowsPerChunk,
        rawOutput,
      },
    } = await this.parse(BatchUpdateData);

    const updates = await resolveBatchUpdates(data, input, inputFormat);

    this.start(dryRun ? `Previewing batch update (${updates.length} ranges)` : `Executing batch update (${updates.length} ranges)`);
    const receipt = await this.gsheet.updateDataBatch(
      updates,
      {
        valueInputOption: valueInputOption as GoogleSheetCli.ValueInputOption,
        dryRun,
        overwriteFormulas,
        chunkByteSize,
        maxRowsPerChunk,
      },
      spreadsheetId
    );
    this.stop();

    const result = {
      operation: this.id,
      ...receipt,
    };

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      this.log(`Batch update preview (dry run) for spreadsheet "${spreadsheetId}":`);
      this.log(`  - Updated ranges: ${receipt.updatedRanges.join(', ') || 'none'}`);
      this.log(`  - Total cells to update: ${receipt.totalCellsUpdated}`);
      if (receipt.changes && receipt.changes.length > 0) {
        for (const change of receipt.changes) {
          this.log(`  - ${change.range}: ${change.after.length} row(s)`);
          if (change.formulasOverwritten && change.formulasOverwritten.length > 0) {
            this.log(`    ⚠️ Overwriting formulas at: ${change.formulasOverwritten.join(', ')}`);
          }
        }
      }
    } else {
      this.log(`Successfully updated ${receipt.updatedRanges.length} range(s) (${receipt.totalCellsUpdated} cells across ${receipt.totalRowsUpdated} rows) in spreadsheet "${spreadsheetId}"`);
      for (const r of receipt.updatedRanges) {
        this.log(`  - ${r}`);
      }
    }

    return result;
  }
}
