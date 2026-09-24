import { Flags } from '@oclif/core';
import Command from '../../lib/base-class';
import { table } from '../../lib/table';
import { DRIVE_FILE_VISIBILITY_NOTE, GoogleSheetCli } from '../../lib/google-sheet';

/** Safety bound for --all: pagination stops here even if Drive keeps handing out tokens. */
const MAX_ALL_RESULTS = 1000;

export default class SpreadsheetList extends Command {
  static description =
    'List spreadsheets visible to this app via the Drive API (discovery). ' +
    'Uses the drive.file scope: only files this app created or has opened are listed — an empty result does not prove a spreadsheet is absent. ' +
    'Duplicate titles stay separate rows; nothing is ever selected automatically, pick an id yourself. ' +
    'OAuth tokens issued before drive.file was added must re-run `gsheet auth:login`.';

  static examples = [
    `$ gsheet spreadsheet:list
`,
    `$ gsheet spreadsheet:list --name "Report 2026" --rawOutput
`,
    `$ gsheet spreadsheet:list --name "Report" --pageSize 100 --all
`,
    `$ gsheet spreadsheet:list --name "Report 2026" --exact
`,
  ];

  static flags = {
    ...Command.flags,
    name: Flags.string({
      description: 'Title fragment to match (partial match); without it every visible spreadsheet is listed',
      required: false,
    }),
    exact: Flags.boolean({
      description: 'Match the whole title exactly instead of a substring (requires --name)',
      required: false,
      dependsOn: ['name'],
    }),
    pageSize: Flags.integer({
      description: 'Files per page (default 50, bounded to 1-100)',
      required: false,
      min: 1,
      max: 100,
    }),
    pageToken: Flags.string({
      description: 'Continuation token from a previous listing (nextPageToken)',
      required: false,
      exclusive: ['all'],
    }),
    all: Flags.boolean({
      description: `Page through every result, aggregated into one listing (safety bound: ${MAX_ALL_RESULTS} files)`,
      required: false,
      exclusive: ['pageToken'],
    }),
  };

  async run(): Promise<GoogleSheetCli.ListSpreadsheetsResult> {
    const {
      flags: { name, exact, pageSize, pageToken, all, rawOutput },
    } = await this.parse(SpreadsheetList);
    const options: GoogleSheetCli.ListSpreadsheetsOptions = { name, exact, pageSize };

    this.start('Listing spreadsheets');
    let result: GoogleSheetCli.ListSpreadsheetsResult;
    if (all) {
      // Aggregate page by page until Drive stops handing out a token or the safety bound is
      // reached. When the bound truncates, the outstanding token stays on the result so a
      // caller can resume honestly instead of trusting an exhausted-looking listing.
      const files: GoogleSheetCli.ListedSpreadsheet[] = [];
      let token: string | undefined;
      do {
        const page = await this.gsheet.listSpreadsheets({ ...options, pageToken: token });
        files.push(...page.files);
        token = page.nextPageToken;
      } while (token && files.length < MAX_ALL_RESULTS);
      result = { files, ...(token ? { nextPageToken: token } : {}), visibilityNote: DRIVE_FILE_VISIBILITY_NOTE };
    } else {
      result = await this.gsheet.listSpreadsheets({ ...options, pageToken });
    }
    this.stop();

    if (rawOutput) {
      this.logRaw('', result);
    } else {
      if (result.files.length > 0) {
        table(
          result.files.map((file) => ({ id: file.id, name: file.name, modifiedTime: file.modifiedTime ?? '-' })),
          { id: {}, name: {}, modifiedTime: {} }
        );
      } else {
        this.log('No spreadsheets found.');
      }
      if (result.nextPageToken) {
        if (all) {
          this.log(`Stopped at the ${MAX_ALL_RESULTS}-file safety bound — continue with --pageToken ${result.nextPageToken}.`);
        } else {
          this.log(`More results available — continue with --pageToken ${result.nextPageToken} (or use --all).`);
        }
      }
      this.log(result.visibilityNote);
    }

    return result;
  }
}
