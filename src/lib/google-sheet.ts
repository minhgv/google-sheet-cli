import { auth, sheets, sheets_v4 } from '@googleapis/sheets';
import { OAuth2Client } from 'google-auth-library';
import { CredentialsInput, normalizeCredentials } from './credentials';
import { getAuthenticatedClient } from './oauth';
import { log } from './log';
import { colToA, getLongestArray, getRange, parseRange, rangeWorksheet, requiredGrid } from './utils';
import {
  a1ToCol,
  colToA1,
  DEFAULT_BATCH_GET_CHUNK_SIZE,
  DEFAULT_CHUNK_BYTE_SIZE,
  escapeWorksheetTitle,
  extractFormulaText,
  findFormulaOverwrites,
  FormulaConflict,
  formatA1Cell,
  formatBoundedA1Range,
  isFormula,
  METADATA_KEY_REPORT_MANAGED,
  packUpdateBatches,
  parseA1Cell,
  parseStrictA1Range,
  SheetManagedMetadata,
  toGoogleExtendedValue,
  toGoogleNumberFormat,
} from './sheet-batch';
import type { ReportCell, ReportDocument, ReportFormula, ReportNumberFormat, ReportSheet } from './report/types';

export namespace GoogleSheetCli {
  export interface Credentials {
    client_email: string;
    private_key: string;
  }

  export type RawData = (string | number | boolean | null)[][];

  export enum ValueInputOption {
    USER_ENTERED = 'USER_ENTERED',
    RAW = 'RAW',
  }

  export enum ValueRenderOption {
    FORMATTED_VALUE = 'FORMATTED_VALUE',
    UNFORMATTED_VALUE = 'UNFORMATTED_VALUE',
    FORMULA = 'FORMULA',
  }

  export enum DateTimeRenderOption {
    SERIAL_NUMBER = 'SERIAL_NUMBER',
    FORMATTED_STRING = 'FORMATTED_STRING',
  }

  export interface QueryOptions {
    minCol?: number;
    maxCol?: number;
    minRow?: number;
    maxRow?: number;
    range?: string;
    valueInputOption?: ValueInputOption;
    worksheetTitle?: string | null;
    hasHeaderRow?: boolean;
    valueRenderOption?: ValueRenderOption;
    dateTimeRenderOption?: DateTimeRenderOption;
    rawOnly?: boolean;
  }

  export interface FormattedData {
    [name: string]: string;
  }

  export interface SheetData {
    rawData: RawData;
    formatted: FormattedData[];
    header: string[];
    range?: string | null;
  }

  export interface BatchGetOptions {
    valueRenderOption?: ValueRenderOption;
    dateTimeRenderOption?: DateTimeRenderOption;
    chunkSize?: number;
  }

  export interface ValueRangeResult {
    range: string;
    values: RawData;
  }

  export interface BatchUpdateOptions {
    valueInputOption?: ValueInputOption;
    dryRun?: boolean;
    overwriteFormulas?: boolean;
    chunkByteSize?: number;
    maxRowsPerChunk?: number;
  }

  export interface BatchUpdateChange {
    range: string;
    before?: RawData;
    after: RawData;
    formulasOverwritten?: string[];
  }

  export interface BatchUpdateReceipt {
    spreadsheetId: string;
    updatedRanges: string[];
    totalRowsUpdated: number;
    totalColumnsUpdated: number;
    totalCellsUpdated: number;
    dryRun: boolean;
    changes?: BatchUpdateChange[];
    batchesExecuted: number;
  }

  export interface AppendTableOptions {
    worksheetTitle?: string | null;
    range?: string;
    valueInputOption?: ValueInputOption;
    insertDataOption?: 'OVERWRITE' | 'INSERT_ROWS';
  }

  export interface AppendTableResult {
    spreadsheetId: string;
    tableRange: string;
    updatedRange: string;
    updatedRows: number;
    updatedColumns: number;
    updatedCells: number;
  }

  export interface ApplyReportOptions {
    dryRun?: boolean;
    overwriteFormulas?: boolean;
    overwrite?: boolean;
  }
  export interface ApplyReportReceipt {
    spreadsheetId: string;
    templateId: string;
    templateVersion: string | number;
    sourceHash: string;
    dryRun: boolean;
    sheetsApplied: {
      name: string;
      writtenRange: string;
      rowsCount: number;
      colsCount: number;
      clearedRange?: string;
    }[];
    formulasProtected: number;
    formulasOverwritten: number;
  }

  export interface WorksheetFormatting {
    freezeRows?: number;
    columnWidths?: { column: number; width: number }[];
    numberFormats?: { column: number; format: string; startRow?: number; endRow?: number }[];
  }
}

// The Sheets API scope. 2.x asked for the retired Sheets v3 feed scope, which Google still
// accepted for v4 calls; this is the scope the v4 API actually documents, and it covers every
// call this class makes, `spreadsheets.create` included. Service account JWTs carry their scope
// in the assertion rather than in a consent screen, so nothing has to be re-granted.
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

const LOG_NAMESPACE = 'gsheet:sheets';

/**
 * Say on stderr why a call did nothing. Not gated behind DEBUG, unlike the credentials
 * output: a silent no-op is the thing worth warning about, so the caller has to see it
 * without knowing to ask.
 */
const warn = (message: string): void => log(LOG_NAMESPACE, message);

/**
 * How long to wait before each retry of a quota rejection, in ms.
 *
 * Google answers a spent per-minute quota with `429 RESOURCE_EXHAUSTED` and asks for exponential
 * backoff (https://developers.google.com/workspace/sheets/api/limits). Retrying that is already
 * happening: `googleapis-common` sets `retry: true` on every request it builds, and 429 is in
 * gaxios' default `statusCodesToRetry`. What is not happening is waiting long enough. gaxios'
 * default schedule is 100ms, 500ms, 1500ms, so - measured against the fake in test/fake-sheets.ts
 * - all four attempts land inside 2.1s of a limit that is counted over a whole minute, and each
 * of them spends another unit of the quota they are queueing for.
 *
 * Three waits of 3s, 12s and 48s put the last attempt 63s after the first rejection, so the
 * minute that rejection was counted in has rolled over by the time it is made. A bucket still
 * empty then is not transient, and the call fails carrying Google's own message.
 */
const QUOTA_BACKOFF_MS = [3_000, 12_000, 48_000];

/**
 * Retry settings handed to every request the client makes.
 *
 * `retryBackoff` is the only thing set, and it replaces nothing but the length of the wait.
 * gaxios keeps deciding *whether* to retry - the same statuses, the same methods, the same three
 * retries - so no request that fails today starts being retried, and a request that succeeds
 * never reaches any of this. A rejection that is not a 429 keeps gaxios' own delay.
 *
 * That leaves one gap on purpose: gaxios retries GET, HEAD, PUT, OPTIONS and DELETE, so
 * `spreadsheets.get` and `values.update` are covered but the POSTs - `spreadsheets.batchUpdate`
 * and `values.append` - are not. A 429 is a rejection rather than a half-done write, so those
 * would be safe to replay, but the method list is also what stops a non-idempotent write from
 * being replayed after a 5xx, and narrowing it to the quota case is a bigger change than this
 * one is meant to be.
 */
const RETRY_CONFIG = {
  retryBackoff: (error: unknown, defaultBackoffMs: number): Promise<void> => {
    const errObj = error && typeof error === 'object' ? (error as { config?: { retryConfig?: { currentRetryAttempt?: number } }; response?: { status?: number } }) : undefined;
    // gaxios has already counted this attempt when it calls us, so the first retry is 1.
    const attempt: number = errObj?.config?.retryConfig?.currentRetryAttempt || 1;
    const quota = errObj?.response?.status === 429;
    const ms = quota ? QUOTA_BACKOFF_MS[attempt - 1] ?? QUOTA_BACKOFF_MS[QUOTA_BACKOFF_MS.length - 1] : defaultBackoffMs;
    // A silent 48 second pause is indistinguishable from a hang, so say what is being waited for.
    if (quota) warn(`quota exceeded, retrying in ${Math.round(ms / 1000)}s (attempt ${attempt} of ${QUOTA_BACKOFF_MS.length})`);
    return new Promise((resolve) => setTimeout(resolve, ms));
  },
};

/**
 * GoogleSheet helper class for CRUD and Batch operations
 *
 * @export
 * @class GoogleSheet
 */
export default class GoogleSheet {
  private sheets!: sheets_v4.Sheets;

  /**
   * Creates an instance of GoogleSheet.
   * @param {string} [spreadsheetId]
   * @param {string} [worksheetTitle]
   * @memberof GoogleSheet
   */
  constructor(private spreadsheetId?: string, private worksheetTitle?: string | null) {}

  /**
   * Authorize with credentials, either passed directly or read from a service account JSON file
   *
   * @param {CredentialsInput} credentials
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async authorize(credentials: CredentialsInput): Promise<void> {
    const { client_email, private_key } = normalizeCredentials(credentials);
    if (!client_email) throw new Error('client_email is required to authorize');
    if (!private_key) throw new Error('private_key is required to authorize');
    // Create the JWT client
    const client = new auth.JWT({ email: client_email, key: private_key, scopes: [SHEETS_SCOPE] });
    this.sheets = sheets({ version: 'v4', auth: client, retryConfig: RETRY_CONFIG });
  }

  /**
   * Authorize with OAuth 2.0 user credentials (personal Google account).
   * Loads tokens from ~/.config/google-sheet-cli/token.json and refreshes if expired.
   *
   * @param {string} [clientSecretPath] - Path to client_secret.json (Desktop App type)
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async authorizeOAuth(clientSecretPath?: string): Promise<void> {
    const client = await getAuthenticatedClient(clientSecretPath);
    // google-auth-library v11 (root) vs v10 (nested in @googleapis/sheets) have incompatible private fields.
    // The runtime objects are identical; the cast resolves the dual-package type mismatch.
    this.sheets = sheets({ version: 'v4', auth: client as unknown as InstanceType<typeof auth.JWT>, retryConfig: RETRY_CONFIG });
  }

  /**
   * Get information about the spreadsheet.
   *
   * @param {string} [spreadsheetId]
   * @returns {Promise<sheets_v4.Schema$Spreadsheet>}
   * @memberof GoogleSheet
   */
  async getSpreadsheet(spreadsheetId?: string): Promise<sheets_v4.Schema$Spreadsheet> {
    const { data: sheet } = await this.sheets.spreadsheets.get({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
    });

    if (!sheet) throw new Error(`Spreadsheet "${spreadsheetId || this.spreadsheetId}" not found`);
    return sheet;
  }

  /**
   * Get information about the worksheet.
   *
   * @param {string} title
   * @param {string} [spreadsheetId]
   * @returns {Promise<sheets_v4.Schema$Sheet>}
   * @memberof GoogleSheet
   */
  async getWorksheet(title: string, spreadsheetId?: string): Promise<sheets_v4.Schema$Sheet> {
    const { sheets = [], properties: { title: ssTitle = '' } = {} } = await this.getSpreadsheet(spreadsheetId);

    const sheet = sheets.find(({ properties: { title: ws } = {} }) => ws === title);
    if (!sheet) throw new Error(`Sheet "${title}" not found in "${ssTitle}"`);

    this.worksheetTitle = sheet?.properties?.title;
    return sheet;
  }

  /**
   * Get the data of the specified cells (or every available cell data)
   *
   * @param {GoogleSheetCli.QueryOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.SheetData>}
   * @memberof GoogleSheet
   */
  async getData(options: GoogleSheetCli.QueryOptions = {}, spreadsheetId?: string): Promise<GoogleSheetCli.SheetData> {
    options.worksheetTitle = options.worksheetTitle || this.worksheetTitle;
    if (options.range) {
      const parsedOptions = parseRange(options.range);
      // A quoted title inside the range overwrites worksheetTitle; an unquoted one does not.
      // That is not a preference, it is what 2.2.0 did - its regex only ever recognised the
      // quoted form - and the choice is remembered on the instance, so it steers every later
      // command in the run as well. Both halves have to stay.
      if (parsedOptions.worksheetTitle && rangeWorksheet(options.range).quoted) {
        options.worksheetTitle = parsedOptions.worksheetTitle;
      }
      if (parsedOptions.minCol) {
        options.minCol = parsedOptions.minCol;
      }
      if (parsedOptions.maxCol) {
        options.maxCol = parsedOptions.maxCol;
      }
      if (parsedOptions.minRow) {
        options.minRow = parsedOptions.minRow;
      }
      if (parsedOptions.maxRow) {
        options.maxRow = parsedOptions.maxRow;
      }
    }

    if (!options.worksheetTitle) {
      throw new Error('Option property "worksheetTitle" is required');
    }

    const sheet = await this.getWorksheet(options.worksheetTitle, spreadsheetId);
    const { rowCount = 0, columnCount = 0 } = sheet?.properties?.gridProperties || {};

    const sanitizedOptions: GoogleSheetCli.QueryOptions = {
      ...options,
      maxCol: options.maxCol || columnCount || 0,
      maxRow: options.maxRow || rowCount || 0,
    };

    const res = await this.sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      range: getRange(sanitizedOptions),
      valueRenderOption: options.valueRenderOption,
      dateTimeRenderOption: options.dateTimeRenderOption,
    });

    const range = res.data.range;
    let values = res.data.values;

    let header: string[] = [];
    if (sanitizedOptions.hasHeaderRow) {
      if (!sanitizedOptions.minRow || sanitizedOptions.minRow <= 1) {
        if (values && values.length > 0) {
          header = (values[0] as string[]) || [];
          values = values.slice(1);
        } else {
          header = [];
          values = [];
        }
      } else {
        const res = await this.sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId || this.spreadsheetId,
          range: getRange({
            ...sanitizedOptions,
            worksheetTitle: options.worksheetTitle,
            minRow: 1,
            maxRow: 1,
            range: undefined,
          }),
          valueRenderOption: options.valueRenderOption,
          dateTimeRenderOption: options.dateTimeRenderOption,
        });
        [header] = res.data.values ?? [[]];
        if (!header.length) throw new Error('No header row exists');
      }
    }

    // Where the generated `(A)`, `(B)` labels below start counting, and how many of them there
    // are. With an explicit minCol it is minCol, as it always was. With minCol absent it depends
    // on whether 2.2.x got this far at all, and the two cases are deliberately different.
    const returnedOn22x = Boolean(header && header[0]);
    const labelOrigin = sanitizedOptions.minCol || (returnedOn22x ? 0 : 1);

    let maxCol = (sanitizedOptions.maxCol ? sanitizedOptions.maxCol + 1 : 0) - labelOrigin;
    let maxRow = 0;
    if (values) {
      maxCol = getLongestArray(values).length;
      maxRow = values.length;
    }

    // fill missing headings
    for (let c = 0; c < maxCol; c++) {
      header[c] = header[c] || `(${colToA(c + labelOrigin)})`;
    }

    const rawOnly = Boolean(options.rawOnly);
    const formatted: GoogleSheetCli.FormattedData[] = rawOnly ? [] : new Array(maxRow);
    const rawData: GoogleSheetCli.RawData = new Array(maxRow);

    for (let r = 0; r < maxRow; r++) {
      const row = values?.[r] || [];
      const rawRow: (string | number | boolean | null)[] = new Array(maxCol);
      let set: GoogleSheetCli.FormattedData | undefined;
      if (!rawOnly) {
        set = {};
      }

      for (let c = 0; c < maxCol; c++) {
        const cell = row[c] !== undefined && row[c] !== null ? row[c] : '';
        rawRow[c] = cell;
        if (!rawOnly && set) {
          const heading = header[c];
          set[heading] = typeof cell === 'string' ? cell : String(cell);
        }
      }

      if (!rawOnly && set) {
        formatted[r] = set;
      }
      rawData[r] = rawRow;
    }

    return { rawData, formatted, header, range };
  }

  /**
   * Batch get multiple ranges in a single or bounded set of requests without per-range metadata calls.
   * Returns results in exact input order.
   *
   * @param {string[]} ranges
   * @param {GoogleSheetCli.BatchGetOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.ValueRangeResult[]>}
   * @memberof GoogleSheet
   */
  async getDataBatch(
    ranges: string[],
    options: GoogleSheetCli.BatchGetOptions = {},
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.ValueRangeResult[]> {
    if (!ranges || !Array.isArray(ranges) || ranges.length === 0) {
      return [];
    }

    // Strict prevalidation of all ranges
    for (const r of ranges) {
      parseStrictA1Range(r);
    }

    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId;
    const chunkSize = options.chunkSize || DEFAULT_BATCH_GET_CHUNK_SIZE;
    const results: GoogleSheetCli.ValueRangeResult[] = [];

    for (let i = 0; i < ranges.length; i += chunkSize) {
      const chunkRanges = ranges.slice(i, i + chunkSize);
      const response = await this.sheets.spreadsheets.values.batchGet({
        spreadsheetId: targetSpreadsheetId,
        ranges: chunkRanges,
        valueRenderOption: options.valueRenderOption,
        dateTimeRenderOption: options.dateTimeRenderOption,
      });

      const valueRanges = response.data.valueRanges || [];
      for (let idx = 0; idx < chunkRanges.length; idx++) {
        const requestedRange = chunkRanges[idx];
        const vr = valueRanges[idx];
        results.push({
          range: vr?.range || requestedRange,
          values: vr?.values || [],
        });
      }
    }

    return results;
  }

  /**
   * Append row data to a worksheet, starting after the last row in a specific column (Historical 2.x/3.x contract)
   *
   * @param {GoogleSheetCli.RawData} data
   * @param {GoogleSheetCli.QueryOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async appendData(data: GoogleSheetCli.RawData, options: GoogleSheetCli.QueryOptions, spreadsheetId?: string): Promise<void> {
    const { rawData }: GoogleSheetCli.SheetData = await this.getData({ ...options }, spreadsheetId);
    options.minRow = rawData.length + 1;
    await this.updateData(data, options, spreadsheetId);
  }

  /**
   * Native contiguous table append using Sheets API spreadsheets.values.append without reading entire table.
   *
   * @param {GoogleSheetCli.RawData} data
   * @param {GoogleSheetCli.AppendTableOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.AppendTableResult>}
   * @memberof GoogleSheet
   */
  async appendTableData(
    data: GoogleSheetCli.RawData,
    options: GoogleSheetCli.AppendTableOptions = {},
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.AppendTableResult> {
    if (!Array.isArray(data) || !data.every(Array.isArray)) {
      throw new Error('Check "data" property - has to be supplied as nested array ([["1", "2"], ["3", "4"]])');
    }

    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId;
    const worksheetTitle = options.worksheetTitle || this.worksheetTitle;

    let targetRange = options.range;
    if (!targetRange) {
      if (!worksheetTitle) {
        throw new Error('Option property "worksheetTitle" or "range" is required');
      }
      targetRange = `${escapeWorksheetTitle(worksheetTitle)}!A1`;
    }

    if (!data.length) {
      warn('no rows to append, nothing was sent to the spreadsheet');
      return {
        spreadsheetId: targetSpreadsheetId || '',
        tableRange: targetRange,
        updatedRange: targetRange,
        updatedRows: 0,
        updatedColumns: 0,
        updatedCells: 0,
      };
    }

    const response = await this.sheets.spreadsheets.values.append({
      spreadsheetId: targetSpreadsheetId,
      range: targetRange,
      valueInputOption: options.valueInputOption || GoogleSheetCli.ValueInputOption.RAW,
      insertDataOption: options.insertDataOption || 'INSERT_ROWS',
      requestBody: {
        values: data,
      },
    });

    const updates = response.data.updates;
    const updatedRange = updates?.updatedRange || targetRange;
    const updatedRows = updates?.updatedRows ?? 0;
    const updatedColumns = updates?.updatedColumns ?? 0;
    const updatedCells = updates?.updatedCells ?? 0;

    return {
      spreadsheetId: targetSpreadsheetId || '',
      tableRange: response.data.tableRange || targetRange,
      updatedRange,
      updatedRows,
      updatedColumns,
      updatedCells,
    };
  }

  /**
   * Update the data starting at a specific row and column (Historical 2.x/3.x contract)
   *
   * @param {GoogleSheetCli.RawData} data
   * @param {GoogleSheetCli.QueryOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async updateData(data: GoogleSheetCli.RawData, options: GoogleSheetCli.QueryOptions, spreadsheetId?: string): Promise<void> {
    const namedTitle = options.worksheetTitle;
    options.worksheetTitle = options.worksheetTitle || this.worksheetTitle;

    const { worksheetTitle: rangeTitle, quoted } = options.range ? rangeWorksheet(options.range) : { worksheetTitle: undefined, quoted: false };

    const contradicted = Boolean(rangeTitle && namedTitle && rangeTitle !== namedTitle);
    if (contradicted) {
      warn(`range "${options.range}" targets worksheet "${rangeTitle}" but worksheetTitle is "${namedTitle}"; writing to "${rangeTitle}", as 2.2.x did`);
    }

    const targetTitle = (quoted || contradicted ? rangeTitle : undefined) || options.worksheetTitle;
    if (!targetTitle) throw new Error('Specify worksheetTitle');
    if (!Array.isArray(data) || !data.every(Array.isArray)) {
      throw new Error('Check "data" property - has to be supplied as nested array ([["1", "2"], ["3", "4"]])');
    }

    if (!data.length) {
      warn('no rows to write, nothing was sent to the spreadsheet');
      return;
    }

    const { rows, cols } = requiredGrid(data, options);
    if (!rangeTitle || rangeTitle === targetTitle) {
      const sheet = await this.getWorksheet(targetTitle, spreadsheetId);
      await this.ensureGridSize(sheet, rows, cols, spreadsheetId);
    }

    const range = getRange(options);
    await this.sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      valueInputOption: options.valueInputOption || GoogleSheetCli.ValueInputOption.RAW,
      range,
      requestBody: {
        values: data,
      },
    });
  }

  /**
   * Batch update multiple ranges with prevalidation, formula overwrite protection,
   * dry-run preview, bounded payload chunking (<=~2MB), and single grid-growth metadata check.
   *
   * @param {{ range: string; values: GoogleSheetCli.RawData }[]} updates
   * @param {GoogleSheetCli.BatchUpdateOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.BatchUpdateReceipt>}
   * @memberof GoogleSheet
   */
  async updateDataBatch(
    updates: { range: string; values: GoogleSheetCli.RawData }[],
    options: GoogleSheetCli.BatchUpdateOptions = {},
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.BatchUpdateReceipt> {
    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId || '';
    if (!updates || !Array.isArray(updates) || updates.length === 0) {
      return {
        spreadsheetId: targetSpreadsheetId,
        updatedRanges: [],
        totalRowsUpdated: 0,
        totalColumnsUpdated: 0,
        totalCellsUpdated: 0,
        dryRun: Boolean(options.dryRun),
        batchesExecuted: 0,
      };
    }

    // Step 1: Strict prevalidation of all updates
    let totalRows = 0;
    let totalCells = 0;
    let maxCols = 0;

    for (const update of updates) {
      if (!update.range || typeof update.range !== 'string') {
        throw new Error('Each update must have a valid non-empty range string');
      }
      const parsed = parseStrictA1Range(update.range);
      if (!Array.isArray(update.values) || !update.values.every(Array.isArray)) {
        throw new Error(`Update values for range "${update.range}" must be a 2D array`);
      }

      const rowCount = update.values.length;
      totalRows += rowCount;
      const updateMaxCols = update.values.reduce((max, row) => Math.max(max, row?.length || 0), 0);
      if (updateMaxCols > maxCols) maxCols = updateMaxCols;
      for (const row of update.values) {
        totalCells += row.length;
      }

      // Check bounding box fit if explicit end range given
      if (parsed.isBounded && parsed.startRow && parsed.endRow && parsed.startCol && parsed.endCol) {
        const allowedRows = parsed.endRow - parsed.startRow + 1;
        const allowedCols = parsed.endCol - parsed.startCol + 1;
        if (rowCount > allowedRows || updateMaxCols > allowedCols) {
          throw new Error(
            `Data (${rowCount}x${updateMaxCols}) exceeds explicit bounded range "${update.range}" (${allowedRows}x${allowedCols})`
          );
        }
      }
    }

    // Step 2: Formula Overwrite Protection & Dry Run Inspection
    const shouldInspectFormulas = options.dryRun || !options.overwriteFormulas;
    let existingValues: GoogleSheetCli.ValueRangeResult[] = [];
    let conflicts: FormulaConflict[] = [];

    if (shouldInspectFormulas) {
      const rangesToInspect = updates.map((u) => u.range);
      existingValues = await this.getDataBatch(
        rangesToInspect,
        { valueRenderOption: GoogleSheetCli.ValueRenderOption.FORMULA },
        targetSpreadsheetId
      );
      conflicts = findFormulaOverwrites(updates, existingValues);

      if (conflicts.length > 0 && !options.overwriteFormulas && !options.dryRun) {
        const conflictDetails = conflicts
          .slice(0, 5)
          .map((c) => `${c.cell} (existing: "${c.existingFormula}", incoming: ${JSON.stringify(c.incomingValue)})`)
          .join('; ');
        throw new Error(
          `Cannot overwrite existing formula(s) without overwriteFormulas=true. Conflicts detected: ${conflictDetails}`
        );
      }
    }

    // Step 3: Dry-Run Preview
    if (options.dryRun) {
      const changes: GoogleSheetCli.BatchUpdateChange[] = updates.map((u, i) => {
        const conflictsForUpdate = conflicts.filter((c) => c.updateIndex === i);
        return {
          range: u.range,
          before: existingValues[i]?.values,
          after: u.values,
          formulasOverwritten: conflictsForUpdate.map((c) => `${c.cell}: ${c.existingFormula}`),
        };
      });

      return {
        spreadsheetId: targetSpreadsheetId,
        updatedRanges: updates.map((u) => u.range),
        totalRowsUpdated: totalRows,
        totalColumnsUpdated: maxCols,
        totalCellsUpdated: totalCells,
        dryRun: true,
        changes,
        batchesExecuted: 0,
      };
    }

    // Step 4: Grid Growth Check once per job across all affected worksheets
    const spreadsheet = await this.getSpreadsheet(targetSpreadsheetId);
    const sheetMap = new Map<string, { sheetId: number; rowCount: number; columnCount: number }>();
    for (const s of spreadsheet.sheets || []) {
      const title = s.properties?.title;
      if (title) {
        sheetMap.set(title.toLowerCase(), {
          sheetId: s.properties?.sheetId || 0,
          rowCount: s.properties?.gridProperties?.rowCount || 0,
          columnCount: s.properties?.gridProperties?.columnCount || 0,
        });
      }
    }

    const gridGrowthRequests: sheets_v4.Schema$Request[] = [];
    const neededDimensions = new Map<string, { neededRows: number; neededCols: number }>();

    for (const update of updates) {
      const parsed = parseStrictA1Range(update.range);
      const sheetName = (parsed.worksheetTitle || this.worksheetTitle || '').toLowerCase();
      const startRow = parsed.startRow || 1;
      const startCol = parsed.startCol || 1;
      const longest = update.values.reduce((max, row) => Math.max(max, row?.length || 0), 0);
      const reqRows = startRow + update.values.length - 1;
      const reqCols = startCol + longest - 1;

      const current = neededDimensions.get(sheetName) || { neededRows: 0, neededCols: 0 };
      neededDimensions.set(sheetName, {
        neededRows: Math.max(current.neededRows, reqRows),
        neededCols: Math.max(current.neededCols, reqCols),
      });
    }

    for (const [sheetName, needed] of neededDimensions.entries()) {
      const info = sheetMap.get(sheetName);
      if (info) {
        if (needed.neededRows > info.rowCount) {
          gridGrowthRequests.push({
            appendDimension: {
              sheetId: info.sheetId,
              dimension: 'ROWS',
              length: needed.neededRows - info.rowCount,
            },
          });
        }
        if (needed.neededCols > info.columnCount) {
          gridGrowthRequests.push({
            appendDimension: {
              sheetId: info.sheetId,
              dimension: 'COLUMNS',
              length: needed.neededCols - info.columnCount,
            },
          });
        }
      }
    }

    if (gridGrowthRequests.length > 0) {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: targetSpreadsheetId,
        requestBody: { requests: gridGrowthRequests },
      });
    }

    // Step 5: Chunk and Execute Writes
    const maxChunkBytes = options.chunkByteSize || DEFAULT_CHUNK_BYTE_SIZE;
    const packedBatches = packUpdateBatches(updates, maxChunkBytes);
    const completedRanges: string[] = [];
    let batchesExecuted = 0;

    for (let b = 0; b < packedBatches.length; b++) {
      const batch = packedBatches[b];
      const batchData: sheets_v4.Schema$ValueRange[] = batch.map((item) => ({
        range: item.range,
        values: item.values,
      }));

      try {
        await this.sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: targetSpreadsheetId,
          requestBody: {
            valueInputOption: options.valueInputOption || GoogleSheetCli.ValueInputOption.RAW,
            data: batchData,
          },
        });
        batchesExecuted++;
        for (const item of batch) {
          completedRanges.push(item.range);
        }
      } catch (error: unknown) {
        const failedRanges = batch.map((item) => item.range).join(', ');
        const msg = error instanceof Error ? error.message : String(error);
        throw new Error(
          `Batch update failed at batch ${b + 1}/${packedBatches.length} [${failedRanges}]: ${msg}. Successfully completed ranges: ${completedRanges.length > 0 ? completedRanges.join(', ') : 'none'}`
        );
      }
    }

    return {
      spreadsheetId: targetSpreadsheetId,
      updatedRanges: completedRanges,
      totalRowsUpdated: totalRows,
      totalColumnsUpdated: maxCols,
      totalCellsUpdated: totalCells,
      dryRun: false,
      batchesExecuted,
    };
  }

  /**
   * Apply ReportDocument to Google Sheets with presentation formatting, typed cell values,
   * formula safety, and managed extent clearing for idempotent shorter reruns.
   *
   * @param {ReportDocument} document
   * @param {GoogleSheetCli.ApplyReportOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.ApplyReportReceipt>}
   * @memberof GoogleSheet
   */
  async applyReport(
    document: ReportDocument,
    options: GoogleSheetCli.ApplyReportOptions = {},
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.ApplyReportReceipt> {
    if (!document || !document.sheets || !Array.isArray(document.sheets)) {
      throw new Error('Invalid ReportDocument: sheets array is required');
    }

    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId || '';
    const spreadsheet = await this.getSpreadsheet(targetSpreadsheetId);
    const existingSheets = new Map<string, sheets_v4.Schema$Sheet>();
    const usedSheetIds = new Set<number>();

    for (const s of spreadsheet.sheets || []) {
      if (s.properties?.title) {
        existingSheets.set(s.properties.title.toLowerCase(), s);
      }
      if (typeof s.properties?.sheetId === 'number') {
        usedSheetIds.add(s.properties.sheetId);
      }
    }

    const templateId = document.provenance?.templateId || '';
    const templateVersion = document.provenance?.templateVersion || '1';
    const sourceHash = document.provenance?.sourceHash || '';

    // Step 1: Pre-calculate sheet IDs and extract developer metadata for all sheets
    interface SheetApplyPlan {
      sheet: ReportSheet;
      sheetName: string;
      sheetId: number;
      isNewSheet: boolean;
      start: { col: number; row: number };
      startRowIdx: number;
      startColIdx: number;
      rowCount: number;
      colCount: number;
      endRowIdx: number;
      endColIdx: number;
      writtenRange: string;
      inspectRange: string;
      prevMetadata?: SheetManagedMetadata;
      existingMetadataId?: number;
      clearedRange?: string;
      clearBounds?: { startRow: number; endRow: number; startCol: number; endCol: number };
    }

    const plans: SheetApplyPlan[] = [];
    const rangesToInspect: string[] = [];
    const inspectPlanIndices: number[] = [];

    let nextAllocatedId = 100000;
    const allocateSheetId = (): number => {
      while (usedSheetIds.has(nextAllocatedId)) {
        nextAllocatedId += Math.floor(Math.random() * 1000) + 1;
      }
      usedSheetIds.add(nextAllocatedId);
      return nextAllocatedId;
    };

    for (const sheet of document.sheets) {
      const sheetName = sheet.name;
      const sheetObj = existingSheets.get(sheetName.toLowerCase());
      const isNewSheet = !sheetObj;
      const sheetId = sheetObj?.properties?.sheetId ?? allocateSheetId();

      const start = sheet.startCell ? parseA1Cell(sheet.startCell) : { col: 1, row: 1 };
      const startRowIdx = start.row - 1;
      const startColIdx = start.col - 1;
      const rowCount = sheet.rows.length;
      const colCount = sheet.rows.reduce((max, row) => Math.max(max, row?.length || 0), 0);
      const endRowIdx = startRowIdx + rowCount;
      const endColIdx = startColIdx + colCount;

      const writtenRange = formatBoundedA1Range(
        sheetName,
        start.col,
        start.row,
        Math.max(start.col, start.col + colCount - 1),
        Math.max(start.row, start.row + rowCount - 1)
      );

      // Extract existing managed developer metadata from sheet
      let prevMetadata: SheetManagedMetadata | undefined;
      let existingMetadataId: number | undefined;

      const metadataList = [
        ...(sheetObj?.developerMetadata || []),
        ...(spreadsheet.developerMetadata || []).filter((dm) => dm.location?.sheetId === sheetId),
      ];

      for (const dm of metadataList) {
        if (dm.metadataKey === METADATA_KEY_REPORT_MANAGED && dm.metadataValue) {
          try {
            const parsed = JSON.parse(dm.metadataValue) as SheetManagedMetadata;
            if (parsed.templateId) {
              prevMetadata = parsed;
              existingMetadataId = dm.metadataId ?? undefined;
              break;
            }
          } catch {
            // Ignore malformed metadata
          }
        }
      }

      // Calculate bounded clear region
      let clearedRange: string | undefined;
      let clearBounds: { startRow: number; endRow: number; startCol: number; endCol: number } | undefined;

      if (sheet.clearManagedRange || (prevMetadata && (prevMetadata.rowCount > rowCount || prevMetadata.colCount > colCount))) {
        if (prevMetadata && prevMetadata.rowCount > 0 && prevMetadata.colCount > 0) {
          const prevStartRowIdx = prevMetadata.startRow - 1;
          const prevStartColIdx = prevMetadata.startCol - 1;
          const prevEndRowIdx = prevStartRowIdx + prevMetadata.rowCount;
          const prevEndColIdx = prevStartColIdx + prevMetadata.colCount;

          const clearStartRow = Math.min(prevStartRowIdx, startRowIdx);
          const clearEndRow = Math.max(prevEndRowIdx, endRowIdx);
          const clearStartCol = Math.min(prevStartColIdx, startColIdx);
          const clearEndCol = Math.max(prevEndColIdx, endColIdx);

          clearBounds = {
            startRow: clearStartRow,
            endRow: clearEndRow,
            startCol: clearStartCol,
            endCol: clearEndCol,
          };
          clearedRange = formatBoundedA1Range(sheetName, clearStartCol + 1, clearStartRow + 1, clearEndCol, clearEndRow);
        } else {
          clearBounds = {
            startRow: startRowIdx,
            endRow: endRowIdx,
            startCol: startColIdx,
            endCol: endColIdx,
          };
          clearedRange = formatBoundedA1Range(sheetName, start.col, start.row, Math.max(start.col, start.col + colCount - 1), Math.max(start.row, start.row + rowCount - 1));
        }
      }

      // Determine range to inspect for preflight
      const inspectStartCol = clearBounds ? clearBounds.startCol + 1 : start.col;
      const inspectStartRow = clearBounds ? clearBounds.startRow + 1 : start.row;
      const inspectEndCol = clearBounds ? clearBounds.endCol : Math.max(start.col, start.col + colCount - 1);
      const inspectEndRow = clearBounds ? clearBounds.endRow : Math.max(start.row, start.row + rowCount - 1);
      const inspectRange = formatBoundedA1Range(sheetName, inspectStartCol, inspectStartRow, inspectEndCol, inspectEndRow);

      const plan: SheetApplyPlan = {
        sheet,
        sheetName,
        sheetId,
        isNewSheet,
        start,
        startRowIdx,
        startColIdx,
        rowCount,
        colCount,
        endRowIdx,
        endColIdx,
        writtenRange,
        inspectRange,
        prevMetadata,
        existingMetadataId,
        clearedRange,
        clearBounds,
      };

      plans.push(plan);

      if (!isNewSheet) {
        rangesToInspect.push(inspectRange);
        inspectPlanIndices.push(plans.length - 1);
      }
    }

    // Step 2: Preflight Inspection (Formulas and Unowned Data) across ALL sheets BEFORE any mutations
    const existingValuesMap = new Map<number, GoogleSheetCli.RawData>();
    if (rangesToInspect.length > 0) {
      const fetched = await this.getDataBatch(
        rangesToInspect,
        { valueRenderOption: GoogleSheetCli.ValueRenderOption.FORMULA },
        targetSpreadsheetId
      );
      for (let i = 0; i < fetched.length; i++) {
        const planIdx = inspectPlanIndices[i];
        existingValuesMap.set(planIdx, fetched[i].values || []);
      }
    }

    let formulasProtectedCount = 0;
    let formulasOverwrittenCount = 0;

    for (let pIdx = 0; pIdx < plans.length; pIdx++) {
      const plan = plans[pIdx];
      if (plan.isNewSheet) continue;

      const existingData = existingValuesMap.get(pIdx) || [];
      const parsedInspect = parseStrictA1Range(plan.inspectRange);
      const inspectBaseCol = parsedInspect.startCol || 1;
      const inspectBaseRow = parsedInspect.startRow || 1;

      const isOwnedCell = (col: number, row: number): boolean => {
        if (!plan.prevMetadata || plan.prevMetadata.templateId !== templateId) {
          return false;
        }
        const pStartRow = plan.prevMetadata.startRow;
        const pStartCol = plan.prevMetadata.startCol;
        const pEndRow = pStartRow + plan.prevMetadata.rowCount - 1;
        const pEndCol = pStartCol + plan.prevMetadata.colCount - 1;
        return row >= pStartRow && row <= pEndRow && col >= pStartCol && col <= pEndCol;
      };

      // Check for unowned non-empty cells
      if (!options.overwrite) {
        for (let r = 0; r < plan.rowCount; r++) {
          const row = plan.sheet.rows[r] || [];
          for (let c = 0; c < row.length; c++) {
            if (row[c] === undefined) continue; // ragged row: missing cell untouched
            const colNum = plan.start.col + c;
            const rowNum = plan.start.row + r;

            const existingR = rowNum - inspectBaseRow;
            const existingC = colNum - inspectBaseCol;
            const existingVal = existingData[existingR]?.[existingC];

            const hasExistingValue = existingVal !== undefined && existingVal !== null && existingVal !== '';
            if (hasExistingValue && !isOwnedCell(colNum, rowNum)) {
              if (!options.dryRun) {
                throw new Error(
                  `Cannot overwrite unowned populated cell at "${plan.sheetName}!${formatA1Cell(colNum, rowNum)}" (existing: ${JSON.stringify(existingVal)}) without overwrite=true`
                );
              }
            }
          }
        }
      }

      // Check formula collisions and count protected / overwritten formulas
      for (let r = 0; r < plan.rowCount; r++) {
        const row = plan.sheet.rows[r] || [];
        for (let c = 0; c < row.length; c++) {
          const incomingCell = row[c];
          if (incomingCell === undefined) continue; // ragged row: skipped, does not collide

          const colNum = plan.start.col + c;
          const rowNum = plan.start.row + r;
          const existingR = rowNum - inspectBaseRow;
          const existingC = colNum - inspectBaseCol;
          const existingVal = existingData[existingR]?.[existingC];

          if (typeof existingVal === 'string' && existingVal.startsWith('=')) {
            const incomingFormula = extractFormulaText(incomingCell);
            if (incomingFormula === existingVal) {
              // Identical formula: safe refresh
              continue;
            }

            if (!options.overwriteFormulas) {
              formulasProtectedCount++;
              if (!options.dryRun) {
                throw new Error(
                  `Formula collision at "${plan.sheetName}!${formatA1Cell(colNum, rowNum)}": existing "${existingVal}" protected. Set overwriteFormulas=true to proceed.`
                );
              }
            } else {
              formulasOverwrittenCount++;
            }
          }
        }
      }
    }

    // Step 3: Build Atomic Batch Requests for All Sheets
    const batchRequests: sheets_v4.Schema$Request[] = [];
    const sheetsApplied: GoogleSheetCli.ApplyReportReceipt['sheetsApplied'] = [];

    // Missing sheets creation requests
    for (const plan of plans) {
      if (plan.isNewSheet) {
        batchRequests.push({
          addSheet: {
            properties: {
              sheetId: plan.sheetId,
              title: plan.sheetName,
            },
          },
        });
      }
    }

    // Per-sheet mutations
    for (const plan of plans) {
      const { sheet, sheetName, sheetId, start, startRowIdx, startColIdx, rowCount, colCount, endRowIdx, endColIdx } = plan;
      const sheetObj = existingSheets.get(sheetName.toLowerCase());

      const gridRows = sheetObj?.properties?.gridProperties?.rowCount || 1000;
      const gridCols = sheetObj?.properties?.gridProperties?.columnCount || 26;

      // Ensure grid size
      if (!plan.isNewSheet && (endRowIdx > gridRows || endColIdx > gridCols)) {
        const addRows = Math.max(0, endRowIdx - gridRows);
        const addCols = Math.max(0, endColIdx - gridCols);
        if (addRows > 0) {
          batchRequests.push({ appendDimension: { sheetId, dimension: 'ROWS', length: addRows } });
        }
        if (addCols > 0) {
          batchRequests.push({ appendDimension: { sheetId, dimension: 'COLUMNS', length: addCols } });
        }
      }

      // Bounded clearing
      if (plan.clearBounds) {
        batchRequests.push({
          updateCells: {
            range: {
              sheetId,
              startRowIndex: plan.clearBounds.startRow,
              endRowIndex: plan.clearBounds.endRow,
              startColumnIndex: plan.clearBounds.startCol,
              endColumnIndex: plan.clearBounds.endCol,
            },
            fields: 'userEnteredValue',
          },
        });
      }

      // Build typed cell data
      const rowData: sheets_v4.Schema$RowData[] = [];
      for (const row of sheet.rows) {
        const cellData: sheets_v4.Schema$CellData[] = [];
        for (const cell of row) {
          cellData.push({
            userEnteredValue: toGoogleExtendedValue(cell),
          });
        }
        rowData.push({ values: cellData });
      }

      batchRequests.push({
        updateCells: {
          rows: rowData,
          start: {
            sheetId,
            rowIndex: startRowIdx,
            columnIndex: startColIdx,
          },
          fields: 'userEnteredValue',
        },
      });

      // Presentation: Freeze rows
      if (typeof sheet.freezeRows === 'number' && sheet.freezeRows > 0) {
        batchRequests.push({
          updateSheetProperties: {
            properties: {
              sheetId,
              gridProperties: {
                frozenRowCount: sheet.freezeRows,
              },
            },
            fields: 'gridProperties.frozenRowCount',
          },
        });
      }

      // Presentation: Column widths
      if (sheet.columnWidths && Array.isArray(sheet.columnWidths)) {
        for (let colIdx = 0; colIdx < sheet.columnWidths.length; colIdx++) {
          const charWidth = sheet.columnWidths[colIdx];
          if (typeof charWidth === 'number' && charWidth > 0) {
            batchRequests.push({
              updateDimensionProperties: {
                range: {
                  sheetId,
                  dimension: 'COLUMNS',
                  startIndex: startColIdx + colIdx,
                  endIndex: startColIdx + colIdx + 1,
                },
                properties: {
                  pixelSize: Math.round(charWidth * 8.5),
                },
                fields: 'pixelSize',
              },
            });
          }
        }
      }

      // Presentation: Number formats
      if (sheet.numberFormats && Array.isArray(sheet.numberFormats)) {
        for (const nf of sheet.numberFormats) {
          const col0 = startColIdx + (nf.column - 1);
          batchRequests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: startRowIdx + (sheet.freezeRows || 0),
                endRowIndex: endRowIdx,
                startColumnIndex: col0,
                endColumnIndex: col0 + 1,
              },
              cell: {
                userEnteredFormat: {
                  numberFormat: toGoogleNumberFormat(nf.format),
                },
              },
              fields: 'userEnteredFormat.numberFormat',
            },
          });
        }
      }

      // Provenance metadata tracking marker (Update if existing, Create if new)
      const metadataPayload: SheetManagedMetadata = {
        templateId,
        templateVersion,
        sourceHash,
        startRow: start.row,
        startCol: start.col,
        rowCount,
        colCount,
        updatedAt: new Date().toISOString(),
      };

      if (plan.existingMetadataId) {
        batchRequests.push({
          updateDeveloperMetadata: {
            dataFilters: [
              {
                developerMetadataLookup: {
                  metadataId: plan.existingMetadataId,
                },
              },
            ],
            developerMetadata: {
              metadataId: plan.existingMetadataId,
              metadataKey: METADATA_KEY_REPORT_MANAGED,
              metadataValue: JSON.stringify(metadataPayload),
              location: {
                sheetId,
                locationType: 'SHEET',
              },
              visibility: 'DOCUMENT',
            },
            fields: 'metadataValue',
          },
        });
      } else {
        batchRequests.push({
          createDeveloperMetadata: {
            developerMetadata: {
              metadataKey: METADATA_KEY_REPORT_MANAGED,
              metadataValue: JSON.stringify(metadataPayload),
              location: {
                sheetId,
                locationType: 'SHEET',
              },
              visibility: 'DOCUMENT',
            },
          },
        });
      }

      sheetsApplied.push({
        name: sheetName,
        writtenRange: plan.writtenRange,
        rowsCount: rowCount,
        colsCount: colCount,
        clearedRange: plan.clearedRange,
      });
    }

    // Step 4: Execute atomic batchUpdate (dryRun sends NO mutation)
    if (!options.dryRun && batchRequests.length > 0) {
      await this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: targetSpreadsheetId,
        requestBody: { requests: batchRequests },
      });
    }

    return {
      spreadsheetId: targetSpreadsheetId,
      templateId,
      templateVersion,
      sourceHash,
      dryRun: Boolean(options.dryRun),
      sheetsApplied,
      formulasProtected: formulasProtectedCount,
      formulasOverwritten: formulasOverwrittenCount,
    };
  }
  /**
   * Narrow execution of structural/presentation batchUpdate requests on the spreadsheet
   *
   * @param {sheets_v4.Schema$Request[]} requests
   * @param {string} [spreadsheetId]
   * @returns {Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse>}
   * @memberof GoogleSheet
   */
  async batchUpdateSpreadsheet(
    requests: sheets_v4.Schema$Request[],
    spreadsheetId?: string
  ): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
    const response = await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      requestBody: { requests },
    });
    return response.data;
  }

  /**
   * Format a worksheet with freeze rows, column widths, and number formats
   *
   * @param {string} worksheetTitle
   * @param {GoogleSheetCli.WorksheetFormatting} formatting
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async formatWorksheet(
    worksheetTitle: string,
    formatting: GoogleSheetCli.WorksheetFormatting,
    spreadsheetId?: string
  ): Promise<void> {
    const sheet = await this.getWorksheet(worksheetTitle, spreadsheetId);
    const sheetId = sheet.properties?.sheetId ?? 0;
    const requests: sheets_v4.Schema$Request[] = [];

    if (typeof formatting.freezeRows === 'number') {
      requests.push({
        updateSheetProperties: {
          properties: {
            sheetId,
            gridProperties: {
              frozenRowCount: formatting.freezeRows,
            },
          },
          fields: 'gridProperties.frozenRowCount',
        },
      });
    }

    if (formatting.columnWidths && Array.isArray(formatting.columnWidths)) {
      for (const col of formatting.columnWidths) {
        requests.push({
          updateDimensionProperties: {
            range: {
              sheetId,
              dimension: 'COLUMNS',
              startIndex: col.column - 1,
              endIndex: col.column,
            },
            properties: {
              pixelSize: Math.round(col.width * 8.5),
            },
            fields: 'pixelSize',
          },
        });
      }
    }

    if (formatting.numberFormats && Array.isArray(formatting.numberFormats)) {
      for (const nf of formatting.numberFormats) {
        requests.push({
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: (nf.startRow ? nf.startRow - 1 : 0),
              endRowIndex: nf.endRow,
              startColumnIndex: nf.column - 1,
              endColumnIndex: nf.column,
            },
            cell: {
              userEnteredFormat: {
                numberFormat: toGoogleNumberFormat(nf.format),
              },
            },
            fields: 'userEnteredFormat.numberFormat',
          },
        });
      }
    }

    if (requests.length > 0) {
      await this.batchUpdateSpreadsheet(requests, spreadsheetId);
    }
  }

  /**
   * Grow the worksheet grid so that it holds at least the requested number of rows and columns.
   *
   * @param {sheets_v4.Schema$Sheet} sheet
   * @param {number} neededRows
   * @param {number} neededCols
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  private async ensureGridSize(sheet: sheets_v4.Schema$Sheet, neededRows: number, neededCols: number, spreadsheetId?: string): Promise<void> {
    const { rowCount, columnCount } = sheet.properties?.gridProperties || {};
    const rows = rowCount ?? 0;
    const cols = columnCount ?? 0;
    const sheetId = sheet.properties?.sheetId;

    const requests: sheets_v4.Schema$Request[] = [];
    if (neededRows > rows) requests.push({ appendDimension: { sheetId, dimension: 'ROWS', length: neededRows - rows } });
    if (neededCols > cols) requests.push({ appendDimension: { sheetId, dimension: 'COLUMNS', length: neededCols - cols } });
    if (!requests.length) return;

    await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      requestBody: { requests },
    });
  }

  /**
   * Add a worksheet with title
   *
   * @param {string} title
   * @param {string} [spreadsheetId]
   * @returns {Promise<sheets_v4.Schema$Sheet>}
   * @memberof GoogleSheet
   */
  async addWorksheet(title: string, spreadsheetId?: string): Promise<sheets_v4.Schema$Sheet | undefined> {
    const response = await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title,
              },
            },
          },
        ],
      },
    });
    const sheet = response.data.replies?.[0]?.addSheet;
    this.worksheetTitle = sheet?.properties?.title;
    return sheet;
  }

  /**
   * Remove a worksheet by title
   *
   * @param {string} title
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async removeWorksheet(title: string, spreadsheetId?: string): Promise<void> {
    const sheet = await this.getWorksheet(title, spreadsheetId);
    await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      requestBody: {
        requests: [
          {
            deleteSheet: {
              sheetId: sheet.properties?.sheetId || -1,
            },
          },
        ],
      },
    });
    this.worksheetTitle = '';
  }

  /**
   * Rename a worksheet by title
   *
   * @param {string} title
   * @param {string} newTitle
   * @param {string} [spreadsheetId]
   * @returns {Promise<void>}
   * @memberof GoogleSheet
   */
  async renameWorksheet(title: string, newTitle: string, spreadsheetId?: string): Promise<void> {
    const worksheet = await this.getWorksheet(title, spreadsheetId);
    await this.sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId || this.spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: worksheet.properties?.sheetId || -1,
                title: newTitle,
              },
              fields: 'title',
            },
          },
        ],
      },
    });
    this.worksheetTitle = newTitle;
  }

  /**
   * Add a spreadsheet with title
   *
   * @param {string} title
   * @returns {Promise<sheets_v4.Schema$Spreadsheet>}
   * @memberof GoogleSheet
   */
  async addSpreadsheet(title: string): Promise<sheets_v4.Schema$Spreadsheet> {
    const { data: sheet } = await this.sheets.spreadsheets.create({
      requestBody: {
        properties: {
          title,
        },
      },
    });
    if (sheet.spreadsheetId) {
      this.spreadsheetId = sheet.spreadsheetId;
    }
    return sheet;
  }
}
