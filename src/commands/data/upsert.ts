import { Flags } from '@oclif/core';
import Command, { optionalData, spreadsheetId, valueInputOption, worksheetTitle } from '../../lib/base-class';
import { resolveDataMatrix } from '../../lib/cli-input';
import { GoogleSheetCli } from '../../lib/google-sheet';

export default class UpsertData extends Command {
  static description =
    'Upsert rows keyed by one column: existing keys update the supplied cells only, new keys append below the table';

  static examples = [
    `$ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id '[["id", "score"], ["001", 99]]'

Upserted into "<worksheetTitle>": 0 added, 1 updated, 0 unchanged
`,
    `$ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id --input=rows.csv

Upserted into "<worksheetTitle>": 3 added, 0 updated, 2 unchanged
`,
    `$ gsheet data:upsert --spreadsheetId=<spreadsheetId> --worksheetTitle=<worksheetTitle> --key=id --input=- --inputFormat=json < rows.json
`,
  ];

  static flags = {
    ...Command.flags,
    worksheetTitle,
    spreadsheetId,
    valueInputOption,
    key: Flags.string({
      description: 'Header name of the single key column used to match input rows against existing rows',
      required: true,
    }),
    range: Flags.string({
      description: 'A1 range of the existing table including its header row (e.g. "Sheet1!A1:F100"); defaults to the whole worksheet grid',
      required: false,
    }),
    input: Flags.string({
      char: 'i',
      description: 'Path to input data file (JSON or CSV) or "-" for stdin',
      required: false,
    }),
    inputFormat: Flags.string({
      description: 'Format of input file ("json" or "csv")',
      options: ['json', 'csv'],
      required: false,
    }),
    dryRun: Flags.boolean({
      description: 'Report added/updated/unchanged rows and planned ranges without writing',
      default: false,
      required: false,
    }),
    overwriteFormulas: Flags.boolean({
      description: 'Allow updates to overwrite existing formulas in changed cells',
      default: false,
      required: false,
    }),
  };

  static args = { data: optionalData };

  async run() {
    const {
      args: { data },
      flags: {
        key,
        range,
        dryRun,
        overwriteFormulas,
        input,
        inputFormat,
        worksheetTitle,
        spreadsheetId,
        valueInputOption,
        rawOutput,
      },
    } = await this.parse(UpsertData);

    const rows = await resolveDataMatrix(data, input, inputFormat);

    this.start(dryRun ? 'Previewing upsert' : 'Upserting data');
    const receipt = await this.gsheet.upsert(
      rows,
      {
        key,
        range,
        dryRun,
        overwriteFormulas,
        worksheetTitle,
        valueInputOption: valueInputOption as GoogleSheetCli.ValueInputOption,
      },
      spreadsheetId
    );
    this.stop();

    const result = { operation: this.id, ...receipt };

    if (rawOutput) {
      this.logRaw('', result);
    } else if (dryRun) {
      const formulas = receipt.formulasOverwritten?.length
        ? `; ${receipt.formulasOverwritten.length} existing formula(s) would be overwritten`
        : '';
      this.log(
        `Upsert preview (dry run) for "${receipt.worksheetTitle}": ${receipt.rowsAdded} added, ${receipt.rowsUpdated} updated, ${receipt.rowsUnchanged} unchanged${formulas}; ${receipt.plannedRanges.length} range(s) planned`
      );
    } else {
      this.log(
        `Upserted into "${receipt.worksheetTitle}": ${receipt.rowsAdded} added, ${receipt.rowsUpdated} updated, ${receipt.rowsUnchanged} unchanged`
      );
    }

    return result;
  }
}
