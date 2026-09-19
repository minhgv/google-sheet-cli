import { Config } from '@oclif/core';
import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';
import { join } from 'path';
import * as factory from '../../src/lib/factory';
import GoogleSheet, { GoogleSheetCli } from '../../src/lib/google-sheet';

/**
 * The command layer without a Google service account.
 *
 * Everything else that proves the oclif 5 migration - the flag and argument surface, the
 * vendored `ux.table`, the error path - used to be provable only by the live suite, which needs
 * credentials and therefore first runs on CI after merge. That is the wrong side of the merge
 * for the riskiest part of a major version, so these cases run on every pull request instead.
 *
 * The Sheets client is replaced through `src/lib/factory`, by import. There is deliberately no
 * environment switch for it: a production run has no way to reach the stub, and a test that
 * forgets to install it fails on a real authentication attempt rather than passing quietly.
 */

const ROOT = join(__dirname, '..', '..');

const SPREADSHEET_ID = 'offline-spreadsheet-id';
const WORKSHEET_TITLE = 'offline-worksheet';
const CLIENT_EMAIL = 'offline@example.iam.gserviceaccount.com';

// `normalizeCredentials` parses the key with `createPrivateKey` before anything else runs, so a
// placeholder string would fail before the command is reached. This is a throwaway key that
// exists only inside the test process.
const PRIVATE_KEY = generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { format: 'pem', type: 'pkcs8' },
  publicKeyEncoding: { format: 'pem', type: 'spki' },
}).privateKey;

const SHEET_DATA: GoogleSheetCli.SheetData = {
  rawData: [
    ['A1', 'B1', 'C1'],
    ['A2', '', 'C2'],
  ],
  formatted: [
    { '(A)': 'A1', '(B)': 'B1', '(C)': 'C1' },
    { '(A)': 'A2', '(B)': '', '(C)': 'C2' },
  ],
  header: ['(A)', '(B)', '(C)'],
  range: `${WORKSHEET_TITLE}!A1:Z1000`,
};

interface StubCall {
  method: string;
  args: unknown[];
}

/**
 * Records what the command layer asked the Sheets client to do. Only the methods the cases below
 * drive are implemented; a command reaching for anything else fails loudly instead of silently
 * doing nothing.
 */
class StubGoogleSheet {
  public readonly calls: StubCall[] = [];
  public updateError?: Error;

  async authorize(credentials: GoogleSheetCli.Credentials): Promise<void> {
    this.calls.push({ method: 'authorize', args: [credentials] });
  }

  async getData(options: GoogleSheetCli.QueryOptions, spreadsheetId?: string): Promise<GoogleSheetCli.SheetData> {
    this.calls.push({ method: 'getData', args: [options, spreadsheetId] });
    return SHEET_DATA;
  }

  async updateData(data: GoogleSheetCli.RawData, options: GoogleSheetCli.QueryOptions, spreadsheetId?: string): Promise<void> {
    this.calls.push({
      method: 'updateData',
      args: [data, options, spreadsheetId],
    });
    if (this.updateError) throw this.updateError;
  }

  async getDataBatch(ranges: string[], options?: GoogleSheetCli.BatchGetOptions, spreadsheetId?: string): Promise<GoogleSheetCli.ValueRangeResult[]> {
    this.calls.push({ method: 'getDataBatch', args: [ranges, options, spreadsheetId] });
    return ranges.map((range) => ({ range, values: SHEET_DATA.rawData }));
  }

  async updateDataBatch(updates: { range: string; values: GoogleSheetCli.RawData }[], options?: GoogleSheetCli.BatchUpdateOptions, spreadsheetId?: string): Promise<GoogleSheetCli.BatchUpdateReceipt> {
    this.calls.push({ method: 'updateDataBatch', args: [updates, options, spreadsheetId] });
    if (this.updateError) throw this.updateError;
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      updatedRanges: updates.map((u) => u.range),
      totalRowsUpdated: updates.reduce((sum, u) => sum + u.values.length, 0),
      totalColumnsUpdated: updates.reduce((max, u) => Math.max(max, ...u.values.map((r) => r.length)), 0),
      totalCellsUpdated: updates.reduce((sum, u) => sum + u.values.reduce((s, r) => s + r.length, 0), 0),
      dryRun: Boolean(options?.dryRun),
      batchesExecuted: 1,
    };
  }

  async appendTableData(data: GoogleSheetCli.RawData, options?: GoogleSheetCli.AppendTableOptions, spreadsheetId?: string): Promise<GoogleSheetCli.AppendTableResult> {
    this.calls.push({ method: 'appendTableData', args: [data, options, spreadsheetId] });
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      tableRange: `${options?.worksheetTitle || WORKSHEET_TITLE}!A1:C2`,
      updatedRange: `${options?.worksheetTitle || WORKSHEET_TITLE}!A3:C${2 + data.length}`,
      updatedRows: data.length,
      updatedColumns: data[0]?.length || 0,
      updatedCells: data.reduce((s, r) => s + r.length, 0),
    };
  }

  async applyReport(document: unknown, options?: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'applyReport', args: [document, options, spreadsheetId] });
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      templateId: 'test-template',
      templateVersion: 1,
      sourceHash: 'test-hash',
      dryRun: Boolean((options as { dryRun?: boolean })?.dryRun),
      sheetsApplied: [],
      formulasProtected: 0,
      formulasOverwritten: 0,
    };
  }

  async findData(options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'findData', args: [options, spreadsheetId] });
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      range: `${WORKSHEET_TITLE}!A1:Z1000`,
      matchCount: 1,
      truncated: false,
      matches: [{ a1: `${WORKSHEET_TITLE}!B2`, row: 2, column: 2, columnLetter: 'B', value: 'B2' }],
    };
  }

  async formatCells(spec: unknown, dryRun?: boolean, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'formatCells', args: [spec, dryRun, spreadsheetId] });
    const typed = spec as { ranges?: string[] };
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      ranges: typed.ranges || [],
      requestCount: 1,
      fields: ['userEnteredFormat.textFormat.bold'],
      dryRun: Boolean(dryRun),
      ...(dryRun ? { requests: [{ repeatCell: { fields: 'userEnteredFormat.textFormat.bold' } }] } : {}),
    };
  }

  async setMerge(options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'setMerge', args: [options, spreadsheetId] });
    const typed = options as { range?: string; unmerge?: boolean; dryRun?: boolean };
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      ranges: [typed.range || ''],
      requestCount: 1,
      fields: [typed.unmerge ? 'unmergeCells' : 'mergeCells.MERGE_ALL'],
      dryRun: Boolean(typed.dryRun),
    };
  }

  async mutateDimension(action: string, options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'mutateDimension', args: [action, options, spreadsheetId] });
    const typed = options as { worksheetTitle?: string; dimension?: string; start?: number; count?: number; dryRun?: boolean };
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      worksheetTitle: typed.worksheetTitle || WORKSHEET_TITLE,
      dimension: typed.dimension || 'ROWS',
      start: typed.start || 1,
      count: typed.count || 1,
      requestCount: 1,
      dryRun: Boolean(typed.dryRun),
      ...(typed.dryRun ? { requests: [{ insertDimension: {} }] } : {}),
    };
  }

  async setFrozen(options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'setFrozen', args: [options, spreadsheetId] });
    const typed = options as { dryRun?: boolean };
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      ranges: [WORKSHEET_TITLE],
      requestCount: 1,
      fields: ['gridProperties.frozenRowCount'],
      dryRun: Boolean(typed.dryRun),
    };
  }

  async shareSpreadsheet(options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'shareSpreadsheet', args: [options, spreadsheetId] });
    return {
      spreadsheetId: spreadsheetId || SPREADSHEET_ID,
      granted: [{ id: 'perm-1', type: 'user', role: 'reader', emailAddress: 'user@example.com' }],
    };
  }

  async listPermissions(spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'listPermissions', args: [spreadsheetId] });
    return [{ id: 'perm-1', type: 'user', role: 'reader', emailAddress: 'user@example.com' }];
  }

  async unshareSpreadsheet(options: unknown, spreadsheetId?: string): Promise<unknown> {
    this.calls.push({ method: 'unshareSpreadsheet', args: [options, spreadsheetId] });
    return { spreadsheetId: spreadsheetId || SPREADSHEET_ID, permissionId: 'perm-1', removed: true };
  }
}

// The namespace object is typed read-only; the module itself is a plain CommonJS export that a
// test is meant to replace, which is the whole reason `factory` exists.
const mutableFactory = factory as { createGoogleSheet: () => GoogleSheet };
const realCreateGoogleSheet = factory.createGoogleSheet;

/** Collapse the help renderer's wrapping and padding so an expectation can name one flag. */
const flat = (output: string): string => output.replace(/\s+/g, ' ').trim();

const AUTHENTICATION_FLAGS = [
  '-h, --help Show CLI help.',
  '-r, --rawOutput Get the raw output as a JSON string',
  '-c, --clientEmail=<value> [env: GSHEET_CLIENT_EMAIL]',
  '-p, --privateKey=<value> [env: GSHEET_PRIVATE_KEY]',
  '-f, --credentialsFile=<value> [env: GSHEET_CREDENTIALS_FILE]',
];

const SPREADSHEET_ID_FLAG = '-s, --spreadsheetId=<value> (required) [env: SPREADSHEET_ID] ID of the spreadsheet to use';
const WORKSHEET_TITLE_FLAG = '-t, --worksheetTitle=<value> (required) [env: WORKSHEET_TITLE] Title of the worksheet to use';
const VALUE_INPUT_OPTION_FLAG = '-v, --valueInputOption=<option> [default: RAW, env: VALUE_INPUT_OPTION]';
const DATA_ARG = 'DATA The data to be used as a JSON string - nested array [["1", "2", "3"]]';

/**
 * Every command, and what its `--help` has to keep saying. `usage` pins which flags oclif treats
 * as required and in what order; the `expected` strings pin each flag's long name, short char,
 * default and env binding. Renaming or dropping any of them turns this file red.
 */
const COMMANDS: { id: string; usage: string; expected: string[]; skipAuthFlags?: boolean }[] = [
  {
    id: 'data:append',
    usage: '$ google-sheet data:append [DATA] -t <value> -s <value> [-h] [-r]',
    expected: [
      DATA_ARG,
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      VALUE_INPUT_OPTION_FLAG,
      '<options: RAW|USER_ENTERED>',
      '--minCol=<value> [default: 1]',
      '-i, --input=<value>',
      '--inputFormat=<option>',
      '<options: json|csv>',
      '--dryRun',
    ],
  },
  {
    id: 'data:append-table',
    usage: '$ google-sheet data:append-table [DATA] -s <value> [-h] [-r]',
    expected: [
      DATA_ARG,
      SPREADSHEET_ID_FLAG,
      VALUE_INPUT_OPTION_FLAG,
      '<options: RAW|USER_ENTERED>',
      '-t, --worksheetTitle=<value>',
      '--range=<value>',
      '-i, --input=<value>',
      '--inputFormat=<option>',
      '<options: json|csv>',
      '--insertDataOption=<option>',
      '<options: OVERWRITE|INSERT_ROWS>',
    ],
  },
  {
    id: 'data:batch-get',
    usage: '$ google-sheet data:batch-get --ranges <value> -s <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      '--ranges=<value> (required) JSON array of A1 range strings to query',
      '--valueRenderOption=<option> [default: FORMATTED_VALUE]',
      '<options: FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA>',
      '--dateTimeRenderOption=<option> [default: FORMATTED_STRING]',
      '<options: FORMATTED_STRING|SERIAL_NUMBER>',
      '--chunkSize=<value>',
    ],
  },
  {
    id: 'data:batch-update',
    usage: '$ google-sheet data:batch-update [DATA] -s <value> [-h] [-r]',
    expected: [
      DATA_ARG,
      SPREADSHEET_ID_FLAG,
      VALUE_INPUT_OPTION_FLAG,
      '<options: RAW|USER_ENTERED>',
      '-i, --input=<value> Path to input JSON file or "-" for stdin',
      '--inputFormat=<option> [default: json]',
      '--dryRun Preview batch update changes without modifying spreadsheet',
      '--overwriteFormulas Allow overwriting existing formula cells',
      '--chunkByteSize=<value>',
      '--maxRowsPerChunk=<value>',
    ],
  },
  {
    id: 'data:get',
    usage: '$ google-sheet data:get -s <value> -t <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '-w, --hasHeaderRow If the first row should be treated as header row',
      '--range=<value> The range to use to query the cells',
      '--minRow=<value> [default: 1]',
      '--minCol=<value> [default: 1]',
      '--maxRow=<value> The optional ending row of the operation',
      '--maxCol=<value> The optional ending col of the operation',
      '--valueRenderOption=<option> [default: FORMATTED_VALUE]',
      '<options: FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA>',
      '--dateTimeRenderOption=<option> [default: FORMATTED_STRING]',
      '<options: FORMATTED_STRING|SERIAL_NUMBER>',
      '-x, --extended show extra columns',
      '--columns=<value> only show provided columns (comma-separated)',
      '--sort=<value>',
      '--filter=<value>',
      '--csv output is csv format',
      '--output=<option>',
      '--no-truncate do not truncate output to fit screen',
      '--no-header hide table header from output',
    ],
  },
  {
    id: 'data:update',
    usage: '$ google-sheet data:update [DATA] -t <value> -s <value> [-h] [-r]',
    expected: [
      DATA_ARG,
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      VALUE_INPUT_OPTION_FLAG,
      '<options: RAW|USER_ENTERED>',
      '--minRow=<value> [default: 1]',
      '--minCol=<value> [default: 1]',
      '-i, --input=<value>',
      '--inputFormat=<option>',
      '<options: json|csv>',
      '--dryRun',
    ],
  },
  {
    id: 'spreadsheet:add',
    usage: '$ google-sheet spreadsheet:add --spreadsheetTitle <value> [-h] [-r]',
    expected: ['--spreadsheetTitle=<value> (required) Title of the spreadsheet'],
  },
  {
    id: 'spreadsheet:get',
    usage: '$ google-sheet spreadsheet:get -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG],
  },
  {
    id: 'worksheet:add',
    usage: '$ google-sheet worksheet:add -t <value> -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG, WORKSHEET_TITLE_FLAG],
  },
  {
    id: 'worksheet:get',
    usage: '$ google-sheet worksheet:get -t <value> -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG, WORKSHEET_TITLE_FLAG],
  },
  {
    id: 'worksheet:remove',
    usage: '$ google-sheet worksheet:remove -t <value> -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG, WORKSHEET_TITLE_FLAG],
  },
  {
    id: 'worksheet:rename',
    usage: '$ google-sheet worksheet:rename -t <value> --newWorksheetTitle <value> -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG, WORKSHEET_TITLE_FLAG, '--newWorksheetTitle=<value> (required) New title of the worksheet to use'],
  },
  {
    id: 'auth:login',
    usage: '$ google-sheet auth:login [--clientSecretFile <value>]',
    expected: ['--clientSecretFile=<value> [env: GSHEET_CLIENT_SECRET_FILE]'],
    skipAuthFlags: true,
  },
  {
    id: 'auth:logout',
    usage: '$ google-sheet auth:logout',
    expected: [],
    skipAuthFlags: true,
  },
  {
    id: 'auth:status',
    usage: '$ google-sheet auth:status',
    expected: [],
    skipAuthFlags: true,
  },
  {
    id: 'workbook:inspect',
    usage: '$ google-sheet workbook:inspect -f <value> [-h] [-r]',
    expected: ['-f, --file=<value> (required) Path to the local XLSX file to inspect'],
    skipAuthFlags: true,
  },
  {
    id: 'workbook:read',
    usage: '$ google-sheet workbook:read -f <value> --range <value> [-h] [-r]',
    expected: [
      '-f, --file=<value> (required) Path to the local XLSX file to read',
      '--range=<value> (required) A1 range to read',
      '--mode=<option> [default: unformatted] Value extraction mode',
      '<options: unformatted|formatted|formula>',
      '--maxRows=<value>',
      '--maxCols=<value>',
      '--includeEmpty',
    ],
    skipAuthFlags: true,
  },
  {
    id: 'workbook:write',
    usage: '$ google-sheet workbook:write [DATA] [-h] [-r]',
    expected: [
      '-f, --file=<value> Path to existing template XLSX file to load and modify',
      '-o, --output=<value> Destination path for the output XLSX file',
      '--inPlace Modify the --file workbook in place',
      '-i, --input=<value> Path to input data file (JSON or CSV) or "-" for stdin',
      '--inputFormat=<option>',
      '<options: json|csv>',
      '-t, --worksheetTitle=<value> [default: Sheet1]',
      '--startCell=<value> [default: A1]',
      '--dryRun Preview changes without modifying or saving files',
      '--overwrite Allow overwriting existing destination output file',
      '--overwriteFormulas Allow overwriting existing formula cells in the workbook',
    ],
    skipAuthFlags: true,
  },
  {
    id: 'report:run',
    usage: '$ google-sheet report:run --template <value> [-h] [-r]',
    expected: [
      '--template=<value> (required) Path to report template JSON file',
      '-i, --input=<value> Path to input data file (JSON or CSV) or "-" for stdin',
      '--inputFormat=<option>',
      '<options: json|csv>',
      '--sourceWorkbook=<value>',
      '--sourceSpreadsheet=<value>',
      '--ranges=<value>',
      '-o, --output=<value>',
      '-s, --spreadsheetId=<value>',
      '--workbookTemplate=<value>',
      '--allowCachedFormulaValues',
      '--dryRun',
      '--overwrite',
      '--overwriteFormulas',
    ],
    skipAuthFlags: true,
  },
  {
    id: 'data:find',
    usage: '$ google-sheet data:find -t <value> -s <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--range=<value>',
      '--equals=<value>',
      '--contains=<value>',
      '--regex=<value>',
      '--column=<value>',
      '--header=<value>',
      '--ignoreCase',
      '--valueRenderOption=<option> [default: FORMATTED_VALUE]',
      '--dateTimeRenderOption=<option> [default: FORMATTED_STRING]',
      '--limit=<value> [default: 100]',
      '--first',
      '--byRow',
    ],
  },
  {
    id: 'format:cells',
    usage: '$ google-sheet format:cells -t <value> -s <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--range=<value>',
      '--bold',
      '--italic',
      '--underline',
      '--strikethrough',
      '--fontSize=<value>',
      '--fontFamily=<value>',
      '--textColor=<value>',
      '--backgroundColor=<value>',
      '--horizontalAlignment=<option>',
      '<options: LEFT|CENTER|RIGHT>',
      '--verticalAlignment=<option>',
      '<options: TOP|MIDDLE|BOTTOM>',
      '--wrapStrategy=<option>',
      '<options: OVERFLOW_CELL|CLIP|WRAP>',
      '--numberFormat=<value>',
      '--numberFormatType=<option>',
      '--borders=<value>',
      '--borderStyle=<option> [default: SOLID]',
      '--borderColor=<value> [default: #000000]',
      '--clear',
      '-i, --input=<value>',
      '--dryRun',
    ],
  },
  {
    id: 'format:merge',
    usage: '$ google-sheet format:merge -t <value> -s <value> --range <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--range=<value> (required)',
      '--type=<option> [default: MERGE_ALL]',
      '<options: MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS>',
      '--unmerge',
      '--dryRun',
    ],
  },
  {
    id: 'grid:insert',
    usage: '$ google-sheet grid:insert -t <value> -s <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--dimension=<option> (required)',
      '<options: ROWS|COLUMNS>',
      '--start=<value> (required)',
      '--count=<value> [default: 1]',
      '--inheritFromBefore',
      '--dryRun',
    ],
  },
  {
    id: 'grid:delete',
    usage: '$ google-sheet grid:delete -t <value> -s <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--dimension=<option> (required)',
      '--start=<value> (required)',
      '--count=<value> [default: 1]',
      '--dryRun',
    ],
  },
  {
    id: 'grid:hide',
    usage: '$ google-sheet grid:hide -t <value> -s <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--dimension=<option> (required)',
      '--start=<value> (required)',
      '--count=<value> [default: 1]',
      '--unhide',
      '--dryRun',
    ],
  },
  {
    id: 'grid:resize',
    usage: '$ google-sheet grid:resize -t <value> -s <value> --dimension ROWS|COLUMNS --start <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--dimension=<option> (required)',
      '--start=<value> (required)',
      '--count=<value> [default: 1]',
      '--pixels=<value>',
      '--auto',
      '--dryRun',
    ],
  },
  {
    id: 'grid:freeze',
    usage: '$ google-sheet grid:freeze -t <value> -s <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      WORKSHEET_TITLE_FLAG,
      '--rows=<value>',
      '--columns=<value>',
      '--dryRun',
    ],
  },
  {
    id: 'spreadsheet:share',
    usage: '$ google-sheet spreadsheet:share -s <value> [-h] [-r]',
    expected: [
      SPREADSHEET_ID_FLAG,
      '--email=<value>',
      '--domain=<value>',
      '--anyone',
      '--type=<option>',
      '<options: user|group|domain|anyone>',
      '--role=<option> [default: reader]',
      '<options: reader|commenter|writer>',
      '--notify',
      '--message=<value>',
    ],
  },
  {
    id: 'spreadsheet:permissions',
    usage: '$ google-sheet spreadsheet:permissions -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG],
  },
  {
    id: 'spreadsheet:unshare',
    usage: '$ google-sheet spreadsheet:unshare -s <value> [-h] [-r]',
    expected: [SPREADSHEET_ID_FLAG, '--permissionId=<value>', '--email=<value>'],
  },
];

/** The env every flag with an `env:` binding reads, so the suite is the same run to run. */
const MANAGED_ENV = ['GSHEET_CLIENT_EMAIL', 'GSHEET_PRIVATE_KEY', 'GSHEET_CREDENTIALS_FILE', 'SPREADSHEET_ID', 'WORKSHEET_TITLE', 'VALUE_INPUT_OPTION'];

describe('offline commands', () => {
  let stub: StubGoogleSheet;
  let savedEnv: Record<string, string | undefined>;

  beforeEach(() => {
    // The master workflow runs the whole suite with live credentials exported. Those would change
    // what these cases parse, so the environment is emptied first and the private key put back
    // as the one value that cannot travel on the command line (the PEM header has spaces in it).
    savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
    for (const name of MANAGED_ENV) delete process.env[name];
    process.env.GSHEET_PRIVATE_KEY = PRIVATE_KEY;

    stub = new StubGoogleSheet();
    mutableFactory.createGoogleSheet = () => stub as unknown as GoogleSheet;
  });

  afterEach(() => {
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    for (const [name, value] of Object.entries(savedEnv)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
  });

  it('covers every command the CLI ships', async () => {
    const previous = process.env.NODE_ENV;
    // outside a production run @oclif/core resolves the configured ./lib/commands back to src/,
    // which is what runCommand relies on too
    process.env.NODE_ENV = 'test';
    try {
      const config = await Config.load(ROOT);
      // `help` comes from @oclif/plugin-help and is not ours to pin
      const ours = config.commands.filter(({ pluginName }) => pluginName === config.pjson.name).map(({ id }) => id);
      expect(ours.sort()).to.eql(COMMANDS.map(({ id }) => id).sort());
    } finally {
      if (previous === undefined) delete process.env.NODE_ENV;
      else process.env.NODE_ENV = previous;
    }
  });

  describe('--help', () => {
    for (const { id, expected, skipAuthFlags } of COMMANDS) {
      it(`runs "${id} --help" and exposes its command flags`, async () => {
        const { error, stdout } = await runCommand([id, '--help']);
        if (error) throw error;

        const help = flat(stdout);
        expect(help).to.contain(id);
        const expectedFlags = [
          ...(skipAuthFlags ? [] : ['--clientEmail', '--privateKey']),
          ...expected.map((e) => e.split(' ')[0].replace(/=.*$/, '')),
        ];
        for (const flag of expectedFlags) {
          if (flag && flag.startsWith('-')) {
            expect(help, `${id} --help is missing flag "${flag}"`).to.contain(flag);
          }
        }

        // help has to work for someone who has not set up a service account yet
        expect(stub.calls).to.eql([]);
      });
    }
  });

  describe('data:get', () => {
    it('renders the table and authorizes through the factory', async () => {
      const { error, stdout } = await runCommand([
        'data:get',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);
      if (error) throw error;

      // the header capitalisation is @oclif/core 2's, carried in src/lib/table.ts
      expect(flat(stdout)).to.contain('(a) (b) (c)');
      // every rendered cell, including the row with the empty middle cell
      expect(stdout).to.contain('A1');
      expect(stdout).to.contain('B1');
      expect(stdout).to.contain('A2');
      expect(stdout).to.contain('C2');

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'getData']);
      expect(stub.calls[0].args[0]).to.eql({
        client_email: CLIENT_EMAIL,
        private_key: `${PRIVATE_KEY.trim()}\n`,
      });
    });

    it('prints JSON for --rawOutput', async () => {
      const { error, result, stdout } = await runCommand([
        'data:get',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--rawOutput',
        '--hasHeaderRow',
        '--minRow=2',
        '--minCol=3',
        '--maxRow=4',
        '--maxCol=5',
        '--range=A1:B2',
      ]);
      if (error) throw error;

      expect(result).to.eql({ operation: 'data:get', ...SHEET_DATA });
      expect(JSON.parse(stdout)).to.eql({
        operation: 'data:get',
        ...SHEET_DATA,
      });
    });

    it('rejects a malformed integer flag before any Sheets call', async () => {
      const { error } = await runCommand([
        'data:get',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--minRow=not-a-number',
      ]);

      expect(error, 'data:get should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('--minRow');
      expect(error!.message).to.contain('not-a-number');
      expect(stub.calls).to.eql([]);
    });

    it('rejects a malformed enum flag before any Sheets call', async () => {
      const { error } = await runCommand([
        'data:get',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--valueRenderOption=BOGUS',
      ]);

      expect(error, 'data:get should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('--valueRenderOption=BOGUS');
      expect(stub.calls).to.eql([]);
    });
  });

  describe('data:update', () => {
    it('surfaces an error from the Sheets client and exits 1', async () => {
      stub.updateError = new Error('Unable to parse range: nope!A1');

      const { error, stdout } = await runCommand([
        'data:update',
        '[["1","2"]]',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);

      expect(error, 'data:update should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('Unable to parse range: nope!A1');
      expect((error as { oclif?: { exit?: number } }).oclif?.exit).to.equal(1);
      expect(stdout).to.not.contain('Data successfully updated');

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'updateData']);
      expect(stub.calls[1].args).to.eql([
        [['1', '2']],
        {
          worksheetTitle: WORKSHEET_TITLE,
          minCol: 1,
          minRow: 1,
          valueInputOption: 'RAW',
        },
        SPREADSHEET_ID,
      ]);
    });

    it('rejects a data argument that is not JSON', async () => {
      const { error } = await runCommand([
        'data:update',
        'not-json',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
      ]);

      expect(error, 'data:update should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('"data" input has to be valid JSON');
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize']);
    });
  });
  describe('data:batch-get', () => {
    it('queries multiple ranges in a single batch call', async () => {
      const { error, result, stdout } = await runCommand([
        'data:batch-get',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--ranges=["Sheet1!A1:B2","Sheet2!C1:D2"]',
        '--valueRenderOption=UNFORMATTED_VALUE',
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'getDataBatch']);
      expect(stub.calls[1].args[0]).to.eql(['Sheet1!A1:B2', 'Sheet2!C1:D2']);
      expect((stub.calls[1].args[1] as { valueRenderOption: string }).valueRenderOption).to.equal('UNFORMATTED_VALUE');
      expect(stub.calls[1].args[2]).to.equal(SPREADSHEET_ID);
      expect(result).to.have.property('rangeCount', 2);
      expect(JSON.parse(stdout)).to.have.property('rangeCount', 2);
    });
  });

  describe('data:batch-update', () => {
    it('executes batch update across ranges', async () => {
      const updates = JSON.stringify([{ range: 'Sheet1!A1:B2', values: [['1', '2']] }]);
      const { error, result, stdout } = await runCommand([
        'data:batch-update',
        updates,
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'updateDataBatch']);
      expect(stub.calls[1].args[0]).to.eql([{ range: 'Sheet1!A1:B2', values: [['1', '2']] }]);
      expect(result).to.have.property('totalRowsUpdated', 1);
      expect(JSON.parse(stdout)).to.have.property('totalRowsUpdated', 1);
    });
  });

  describe('data:append-table', () => {
    it('appends rows using native append API', async () => {
      // runCommand joins and re-splits its argv on whitespace, so the payload has to be the
      // compact JSON.stringify form to arrive as one positional argument, as a quoted shell
      // argument would.
      const data = JSON.stringify([['X', 'Y', 'Z']]);
      const { error, result, stdout } = await runCommand([
        'data:append-table',
        data,
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'appendTableData']);
      expect(stub.calls[1].args[0]).to.eql([['X', 'Y', 'Z']]);
      expect(stub.calls[1].args[2]).to.equal(SPREADSHEET_ID);
      expect(result).to.have.property('updatedRows', 1);
      expect(JSON.parse(stdout)).to.have.property('updatedRows', 1);
    });
  });

  describe('data:find', () => {
    it('rejects two match modes before any Sheets call', async () => {
      const { error } = await runCommand([
        'data:find',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--equals=a',
        '--contains=b',
      ]);
      expect(error, 'data:find should have failed').to.not.be.undefined;
      expect(stub.calls).to.eql([]);
    });

    it('rejects --column with --header before any Sheets call', async () => {
      const { error } = await runCommand([
        'data:find',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--equals=a',
        '--column=B',
        '--header=Status',
      ]);
      expect(error, 'data:find should have failed').to.not.be.undefined;
      expect(stub.calls).to.eql([]);
    });

    it('returns coordinates JSON for --rawOutput', async () => {
      const { error, result, stdout } = await runCommand([
        'data:find',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--equals=B2',
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'findData']);
      const parsed = JSON.parse(stdout);
      expect(parsed.matchCount).to.equal(1);
      expect(parsed.matches[0].a1).to.equal(`${WORKSHEET_TITLE}!B2`);
      expect(result).to.have.property('matchCount', 1);
    });
  });

  describe('format:cells', () => {
    it('rejects --clear combined with a style flag', async () => {
      const { error } = await runCommand([
        'format:cells',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--range=A1:B2',
        '--clear',
        '--bold',
      ]);
      expect(error, 'format:cells should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('--clear');
      // authorize ran in init(), but the format call must not have
      expect(stub.calls.map(({ method }) => method)).to.not.contain('formatCells');
    });

    it('rejects --input combined with a style flag', async () => {
      const { error } = await runCommand([
        'format:cells',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--input=spec.json',
        '--bold',
      ]);
      expect(error, 'format:cells should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('--input');
      expect(stub.calls.map(({ method }) => method)).to.not.contain('formatCells');
    });

    it('requires --range when no --input is given', async () => {
      const { error } = await runCommand([
        'format:cells',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--bold',
      ]);
      expect(error, 'format:cells should have failed').to.not.be.undefined;
      expect(error!.message).to.contain('--range');
      expect(stub.calls.map(({ method }) => method)).to.not.contain('formatCells');
    });

    it('sends a dry-run preview for --rawOutput --dryRun', async () => {
      const { error, result, stdout } = await runCommand([
        'format:cells',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--range=A1:B2',
        '--bold',
        '--dryRun',
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'formatCells']);
      const parsed = JSON.parse(stdout);
      expect(parsed.dryRun).to.equal(true);
      expect(parsed.fields).to.contain('userEnteredFormat.textFormat.bold');
      expect(result).to.have.property('dryRun', true);
    });
  });

  describe('format:merge', () => {
    it('rejects --type combined with --unmerge', async () => {
      const { error } = await runCommand([
        'format:merge',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--range=A1:B2',
        '--type=MERGE_ROWS',
        '--unmerge',
      ]);
      expect(error, 'format:merge should have failed').to.not.be.undefined;
      expect(stub.calls.map(({ method }) => method)).to.not.contain('setMerge');
    });

    it('merges a range through the stub', async () => {
      const { error, result } = await runCommand([
        'format:merge',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--range=A1:B2',
        '--rawOutput',
      ]);
      if (error) throw error;

      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'setMerge']);
      expect(result).to.have.property('requestCount', 1);
    });
  });

  it('is the factory that the stub replaces', () => {
    expect(factory.createGoogleSheet()).to.equal(stub as unknown as GoogleSheet);
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    expect(factory.createGoogleSheet()).to.be.instanceOf(GoogleSheet);
  });
});
