import { Flags } from '@oclif/core';
import { existsSync } from 'fs';
import Command, { spreadsheetId } from '../../lib/base-class';
import { GoogleSheetCli } from '../../lib/google-sheet';
import { saveBufferAtomic } from '../../lib/xlsx-file';

export default class SpreadsheetExport extends Command {
  static description =
    'Export a spreadsheet to a local PDF or XLSX file through the Drive API (drive.file scope). ' +
    'Under drive.file only files this app created or has opened are visible. ' +
    "The Drive API caps exports at 10 MB; larger spreadsheets fail with Google's own error. " +
    'The output file is written atomically (temporary file in the same directory, then renamed) ' +
    'and an existing file is only replaced with --overwrite.';

  static examples = [
    `$ gsheet spreadsheet:export --spreadsheetId=<id> --format=pdf --output=./report.pdf

Exported spreadsheet <id> to /abs/path/report.pdf (application/pdf, 12345 bytes)
`,
    `$ gsheet spreadsheet:export --spreadsheetId=<id> --format=xlsx --output=./report.xlsx --overwrite

Exported spreadsheet <id> to /abs/path/report.xlsx (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, 23456 bytes)
`,
  ];

  static flags = {
    ...Command.flags,
    spreadsheetId,
    format: Flags.string({
      description: 'Export format',
      options: ['pdf', 'xlsx'],
      required: true,
    }),
    output: Flags.string({
      char: 'o',
      description: 'Path of the local file to write',
      required: true,
    }),
    overwrite: Flags.boolean({
      description: 'Replace an existing output file (default: refuse)',
      required: false,
    }),
  };

  async run() {
    const {
      flags: { spreadsheetId, format, output, overwrite, rawOutput },
    } = await this.parse(SpreadsheetExport);

    // Refuse before the network round-trip when the answer is already no. The atomic save
    // re-checks after the bytes arrive, so a file created in between still cannot be clobbered.
    if (!overwrite && existsSync(output)) {
      this.error(`Output file "${output}" already exists. Pass --overwrite to replace it.`, { exit: 1 });
    }

    this.start('Exporting spreadsheet');
    const exported = await this.gsheet.exportSpreadsheet(
      {
        format: format as GoogleSheetCli.ExportOptions['format'],
      },
      spreadsheetId
    );
    const saved = await saveBufferAtomic(exported.bytes, output, { overwrite: Boolean(overwrite) });
    this.stop();

    const receipt = {
      operation: this.id,
      spreadsheetId: exported.spreadsheetId,
      mimeType: exported.mimeType,
      ...saved,
    };

    if (rawOutput) {
      this.logRaw('', receipt);
    } else {
      this.log(`Exported spreadsheet ${exported.spreadsheetId} to ${saved.savedPath} (${exported.mimeType}, ${saved.bytesWritten} bytes)`);
    }

    return receipt;
  }
}
