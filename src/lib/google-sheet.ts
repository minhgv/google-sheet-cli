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
  ParsedA1Range,
  parseStrictA1Range,
  SheetManagedMetadata,
  toGoogleExtendedValue,
  toGoogleNumberFormat,
} from './sheet-batch';
import type { ReportCell, ReportDocument, ReportFormula, ReportNumberFormat, ReportSheet } from './report/types';
import { buildFormatRequests, buildMergeRequest, FormatSpec } from './sheet-format';
import { buildTableFrame, firstBelowBoundCell, parseUpsertInput, planUpsert } from './upsert-plan';
import { ValidationError } from './validation-error';
// Type-only: the value side of cli-errors is loaded lazily on first failure, so this module's
// static import graph - and with it the oclif-free `google-sheet-cli/sheet` subpath - stays clean.
import type { classifyError, GSheetError, GSheetErrorCode } from './cli-errors';
import { MutationOutcomeBuilder, MutationOutcomeReport, MutationPhase, unknownOutcomeGuidance } from './mutation-outcome';

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
    /**
     * Operation label recorded in a failure's mutation-outcome report; defaults to
     * 'updateDataBatch'. Internal callers that wrap this method (e.g. upsert) pass their own
     * name so agents see the operation they actually invoked.
     */
    operation?: string;
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

  export interface ClearOptions {
    /** A1 range to clear; must be explicitly bounded on both axes, e.g. "Sheet1!A1:D20" */
    range: string;
    worksheetTitle?: string | null;
    dryRun?: boolean;
    overwriteFormulas?: boolean;
  }

  export interface ClearReceipt {
    spreadsheetId: string;
    worksheetTitle: string;
    /** the canonical bounded range that was, or on dryRun would be, cleared */
    range: string;
    rowsCleared: number;
    columnsCleared: number;
    cellsCleared: number;
    dryRun: boolean;
    batchesExecuted: number;
    /** dryRun only: formulas the clear would overwrite */
    formulasOverwritten?: string[];
  }

  export interface UpsertOptions {
    /** header name of the single key column used to match input rows against existing rows */
    key: string;
    /** A1 range of the existing table including its header row; defaults to the whole worksheet grid */
    range?: string;
    worksheetTitle?: string | null;
    dryRun?: boolean;
    overwriteFormulas?: boolean;
    /** defaults to RAW so values land exactly as given ('001' stays text) */
    valueInputOption?: ValueInputOption;
  }

  export interface UpsertReceipt {
    spreadsheetId: string;
    worksheetTitle: string;
    keyColumn: string;
    existingRows: number;
    rowsAdded: number;
    rowsUpdated: number;
    rowsUnchanged: number;
    addedRanges: string[];
    updatedRanges: string[];
    /** every range this run planned to write, in execution order */
    plannedRanges: string[];
    dryRun: boolean;
    batchesExecuted: number;
    /** dryRun only: formulas the planned writes would overwrite */
    formulasOverwritten?: string[];
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

  export interface FindOptions {
    worksheetTitle?: string | null;
    /** A1 range bounding the scan; defaults to the whole worksheet */
    range?: string;
    /** exactly one match mode is required */
    equals?: string;
    contains?: string;
    regex?: string;
    /** restrict matches to one column by A1 letter (e.g. "B") */
    column?: string;
    /** restrict matches to the column whose first scanned row equals this header */
    header?: string;
    ignoreCase?: boolean;
    valueRenderOption?: ValueRenderOption;
    dateTimeRenderOption?: DateTimeRenderOption;
    /** max matches returned; matchCount still reports the true total */
    limit?: number;
    /** collapse matches to unique rows and include full row values */
    byRow?: boolean;
  }

  export interface FindMatch {
    a1: string;
    row: number;
    column: number;
    columnLetter: string;
    value: string;
    /** present only when byRow is set: the full scanned row the match sits in */
    rowValues?: (string | number | boolean | null)[];
  }

  export interface FindResult {
    spreadsheetId: string;
    range: string;
    matchCount: number;
    truncated: boolean;
    matches: FindMatch[];
  }

  export interface FormatReceipt {
    spreadsheetId: string;
    ranges: string[];
    requestCount: number;
    fields: string[];
    dryRun: boolean;
    /** the exact batchUpdate request bodies; populated on dryRun only */
    requests?: sheets_v4.Schema$Request[];
  }

  export interface MergeOptions {
    worksheetTitle?: string | null;
    range: string;
    mergeType?: 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS';
    unmerge?: boolean;
    dryRun?: boolean;
  }

  export type Dimension = 'ROWS' | 'COLUMNS';

  export interface DimensionOptions {
    worksheetTitle?: string | null;
    dimension: Dimension;
    /** 1-based first index of the affected dimension */
    start: number;
    /** how many rows/columns the operation covers; default 1 */
    count?: number;
    /** insertDimension only: inherit formatting from the row/column before instead of after */
    inheritFromBefore?: boolean;
    /** hide only: unhide instead of hide */
    unhide?: boolean;
    /** resize only: explicit pixel size; mutually exclusive with autoResize */
    pixels?: number;
    /** resize only: let Google auto-size to content; mutually exclusive with pixels */
    autoResize?: boolean;
    dryRun?: boolean;
  }

  export interface DimensionReceipt {
    spreadsheetId: string;
    worksheetTitle: string;
    dimension: Dimension;
    start: number;
    count: number;
    requestCount: number;
    dryRun: boolean;
    /** the exact batchUpdate request bodies; populated on dryRun only */
    requests?: sheets_v4.Schema$Request[];
    /** delete dryRun only: the values about to be removed */
    affectedValues?: RawData;
  }

  export interface FreezeOptions {
    worksheetTitle?: string | null;
    /** frozen row count; 0 unfreezes rows */
    rows?: number;
    /** frozen column count; 0 unfreezes columns */
    columns?: number;
    dryRun?: boolean;
  }

  export interface ShareOptions {
    /** grant to these addresses; inferred type "user" unless --type overrides to "group" */
    emails?: string[];
    /** grant to a whole Google Workspace domain */
    domain?: string;
    /** grant to anyone with the link */
    anyone?: boolean;
    /** explicit grantee type; inferred from emails/domain/anyone when omitted */
    type?: 'user' | 'group' | 'domain' | 'anyone';
    role?: 'reader' | 'commenter' | 'writer';
    /** send Google's notification email; default false — agents should not spam */
    notify?: boolean;
    /** message attached to the notification email (requires notify) */
    message?: string;
  }

  export interface Permission {
    id: string;
    type: string;
    role: string;
    emailAddress?: string;
    domain?: string;
    displayName?: string;
  }

  export interface ShareResult {
    spreadsheetId: string;
    granted: Permission[];
  }

  export interface UnshareOptions {
    /** permission id from listPermissions */
    permissionId?: string;
    /** resolve the permission id by grantee email */
    email?: string;
  }

  export interface WorksheetMetadataOptions {
    worksheetTitle?: string | null;
    /** optional bounded A1 range WITHOUT a worksheet title (e.g. "A1:Z100"); the title is prepended */
    range?: string;
  }

  export interface WorksheetMetadata {
    spreadsheetId: string;
    worksheetTitle: string;
    properties: sheets_v4.Schema$SheetProperties;
    /** grid data for the requested range; each entry's startRow/startColumn are 0-based */
    gridData: sheets_v4.Schema$GridData[];
    /** named ranges scoped to this worksheet */
    namedRanges: sheets_v4.Schema$NamedRange[];
  }

  export interface CopySpreadsheetOptions {
    /** title of the new spreadsheet; Drive names the copy after the source when omitted */
    title?: string;
  }

  /** The subset of the Drive File resource a copy reports. */
  export interface DriveFileMetadata {
    id?: string | null;
    name?: string | null;
    mimeType?: string | null;
    webViewLink?: string | null;
  }

  export interface CopySpreadsheetResult {
    /** the spreadsheet that was copied; it is left untouched */
    sourceSpreadsheetId: string;
    /** the copy's file id; pass it as --spreadsheetId to every other command */
    spreadsheetId: string;
    title: string;
    mimeType: string;
    webViewLink?: string;
  }

  export interface ListSpreadsheetsOptions {
    /** title fragment matched with Drive's `contains` operator; omit to list everything visible */
    name?: string;
    /** match the whole title with Drive's `=` operator instead of a substring */
    exact?: boolean;
    /** page size per request; defaults to 50 and is bounded to 1..100 */
    pageSize?: number;
    /** opaque continuation token from a previous page's nextPageToken */
    pageToken?: string;
  }

  /** The subset of the Drive File resource a listing reports. */
  export interface ListedSpreadsheet {
    id: string;
    name: string;
    mimeType: string;
    modifiedTime?: string;
  }

  export interface ListSpreadsheetsResult {
    /** one row per Drive file, in Drive's order; duplicate titles stay separate rows — a title is never an id */
    files: ListedSpreadsheet[];
    /** present when more pages remain; pass it back as pageToken to continue */
    nextPageToken?: string;
    /** fixed disclosure of the drive.file visibility boundary; never drop it from the output */
    visibilityNote: string;
  }

  export interface CopyWorksheetOptions {
    worksheetTitle: string;
    /** copyTo never guesses the destination; it has to be named explicitly */
    destinationSpreadsheetId: string;
  }

  export interface CopyWorksheetResult {
    spreadsheetId: string;
    worksheetTitle: string;
    destinationSpreadsheetId: string;
    /** the new sheet's id inside the destination spreadsheet */
    sheetId: number;
    /** the title the copy carries in the destination; copyTo keeps the source title */
    title: string;
    /** the copy's position in the destination; present when the API reports it */
    index?: number;
  }

  export interface ExportOptions {
    format: 'pdf' | 'xlsx';
  }

  export interface ExportResult {
    spreadsheetId: string;
    mimeType: string;
    /** raw export bytes: write them to disk, never log or JSON-encode them */
    bytes: Buffer;
    byteLength: number;
  }
}

// The Sheets API scope. 2.x asked for the retired Sheets v3 feed scope, which Google still
// accepted for v4 calls; this is the scope the v4 API actually documents, and it covers every
// call this class makes, `spreadsheets.create` included. Service account JWTs carry their scope
// in the assertion rather than in a consent screen, so nothing has to be re-granted.
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

// Sharing is Drive API, not Sheets. `drive.file` covers files this app created or opened —
// the agent report-handoff flow — without jumping to the restricted full-drive scope.
// Existing OAuth tokens predate it: `auth:login` again before `spreadsheet:share`.
const DRIVE_FILE_SCOPE = 'https://www.googleapis.com/auth/drive.file';

const DRIVE_API_BASE = 'https://www.googleapis.com/drive/v3';

// Drive files.export serves several formats; the CLI offers exactly these two, and test/fake-sheets.ts
// mirrors the map. Keys are CLI formats, values the MIME types Drive expects on the query string.
const EXPORT_MIME_TYPES: Record<'pdf' | 'xlsx', string> = {
  pdf: 'application/pdf',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

// The only mime type discovery looks for. Kept in step with test/fake-sheets.ts, which seeds
// and filters on the same literal.
const DRIVE_FILES_SPREADSHEET_MIME = 'application/vnd.google-apps.spreadsheet';

// files.list page size: 50 by default, hard-bounded to Drive's documented maximum of 100 so an
// unbounded caller cannot turn one listing into a bulk export.
const DEFAULT_LIST_PAGE_SIZE = 50;
const MAX_LIST_PAGE_SIZE = 100;

/**
 * Disclosure that rides on every discovery result: under drive.file the listing only covers
 * spreadsheets this application created or has opened, so "no rows" never means "no file".
 */
export const DRIVE_FILE_VISIBILITY_NOTE =
  'Listing uses the drive.file OAuth scope: only spreadsheets this application created or has opened are visible. ' +
  'An empty result does not prove a spreadsheet is absent — address such files by ID.';

/**
 * Shape of the Drive files.list response this class reads; the fields asked for are exactly the
 * ones the result maps onto.
 */
type DriveFilesListResponse = {
  files?: { id?: string | null; name?: string | null; mimeType?: string | null; modifiedTime?: string | null }[];
  nextPageToken?: string;
};

/**
 * Escape a literal for the Drive query language: backslash first, then the single quote, so a
 * title like `Bob's Plan` travels as `Bob\'s Plan` and can never close the quoted string early.
 */
const escapeDriveQueryLiteral = (value: string): string => value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");

/**
 * Build the files.list q parameter from fixed parts and escaped literals only — caller input
 * is a value, never a query fragment.
 */
const driveSpreadsheetsQuery = (name: string | undefined, exact: boolean): string => {
  let query = `mimeType='${DRIVE_FILES_SPREADSHEET_MIME}'`;
  if (name !== undefined) {
    const literal = escapeDriveQueryLiteral(name);
    query += exact ? ` and name='${literal}'` : ` and name contains '${literal}'`;
  }
  return query;
};

/**
 * Resolve the files.list page size: undefined means the 50 default, anything finite is clamped
 * into 1..100 so the listing stays bounded on both ends.
 */
const resolveListPageSize = (pageSize?: number): number => {
  if (pageSize === undefined) return DEFAULT_LIST_PAGE_SIZE;
  if (typeof pageSize !== 'number' || !Number.isFinite(pageSize)) {
    throw new Error('listSpreadsheets requires a finite "pageSize" when one is provided');
  }
  return Math.min(MAX_LIST_PAGE_SIZE, Math.max(1, Math.round(pageSize)));
};

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
 * Physical request kind recorded for one packed values chunk of an `updateDataBatch` run.
 */
const VALUES_BATCH_UPDATE_REQUEST_KIND = 'values.batchUpdate';

/**
 * HTTP status carried by a thrown Google API/transport error, if any, read from the usual
 * response/status/code shapes and validated to the same 100-599 range `asHttpStatus` enforces
 * in cli-errors. cli-errors is deliberately outside this module's static import graph (the
 * `google-sheet-cli/sheet` subpath must stay oclif-free), so only its error shapes are read here.
 */
const outcomeHttpStatus = (err: unknown): number | undefined => {
  if (!err || typeof err !== 'object') return undefined;
  const e = err as Record<string, unknown>;
  const response = e.response;
  const responseStatus = response && typeof response === 'object' ? (response as Record<string, unknown>).status : undefined;
  const candidate = responseStatus ?? e.status ?? e.code;
  return typeof candidate === 'number' && Number.isInteger(candidate) && candidate >= 100 && candidate <= 599
    ? candidate
    : undefined;
};

/** Shape of the cli-errors module the failure paths need. */
interface CliErrorsModule {
  GSheetError: typeof GSheetError;
  GSheetErrorCode: typeof GSheetErrorCode;
  classifyError: typeof classifyError;
}

let cliErrorsModule: CliErrorsModule | undefined;
const loadCliErrors = (): CliErrorsModule => {
  // Lazy require, deliberately not a static import: cli-errors carries @oclif/core, and this
  // module's import graph - the oclif-free `google-sheet-cli/sheet` subpath - must stay free
  // of it until a failure actually needs the typed error contract. The extensionless CJS
  // specifier resolves exactly like this file's static imports, compiled and under ts-node.
  if (!cliErrorsModule) {
    // untyped CJS require; the shape is pinned by CliErrorsModule
    cliErrorsModule = require('./cli-errors') as CliErrorsModule;
  }
  return cliErrorsModule;
};

/**
 * A preflight refusal of `updateDataBatch`: nothing was dispatched, so the outcome carries an
 * empty request list, the validate phase and the derived safe-replay guidance. The message
 * text is byte-identical to the plain `Error` this used to throw; only the typed code
 * (VALIDATION, the established code for local input refusals) and the outcome are added.
 */
const batchPreflightRefusal = (message: string, redactedMessage?: string): GSheetError => {
  const { GSheetError, GSheetErrorCode } = loadCliErrors();
  return new GSheetError(GSheetErrorCode.VALIDATION, message, {
    mutationOutcome: buildBatchUpdateMutationOutcome({ batches: [], batchesExecuted: 0, cause: undefined, phase: 'validate' }),
    ...(redactedMessage !== undefined ? { redactedMessage } : {}),
  });
};

/** Input for {@link buildBatchUpdateMutationOutcome}. */
export interface BatchUpdateOutcomeInput {
  /** The full packed chunk plan, in dispatch order. */
  batches: readonly (readonly { range: string }[])[];
  /** Number of chunks acknowledged before the failure; the chunk at this index is the failing one. */
  batchesExecuted: number;
  /** The thrown transport/API error; inspected for a usable HTTP status. */
  cause: unknown;
  /** Grid-growth effect observed before the values write, when the job attempted any. */
  gridGrowth?: { attempted: boolean; completed: boolean; details?: string };
  /** Phase the failure occurred in; defaults to 'values-write'. */
  phase?: MutationPhase;
  /** False when the failure happened before any values chunk was dispatched (grid-growth phase). */
  failedChunkDispatched?: boolean;
  /**
   * True only for append-like operations reusing this builder: a lost response then locks
   * never-blind-replay via `unknownOutcomeGuidance(true)`. values.update chunks are
   * deterministic overwrites, so updateDataBatch itself stays idempotent and keeps the
   * derived verify-then-replay.
   */
  nonIdempotent?: boolean;
  /** Operation label recorded in the report; defaults to 'updateDataBatch'. */
  operation?: string;
}

/**
 * Build the typed mutation-outcome report for a failed `updateDataBatch` run. Pure mapping of
 * the chunk plan and the thrown cause onto the locked vocabulary: the dispatched prefix is
 * `acknowledged`; the failing chunk is `rejected` only when a usable response proves
 * non-application (4xx, including a 429 refused before processing), `unknown` for a
 * 5xx-after-send or a lost response; the undispatched tail is `not-attempted`. Grid growth
 * rides in `gridGrowth`, separately from the value writes: the grid may have grown even when
 * the value write failed.
 */
export function buildBatchUpdateMutationOutcome(input: BatchUpdateOutcomeInput): MutationOutcomeReport {
  const status = outcomeHttpStatus(input.cause);
  const failedRejected = status !== undefined && status >= 400 && status < 500;
  const failingChunkDispatched = input.failedChunkDispatched !== false && input.batchesExecuted < input.batches.length;
  const builder = new MutationOutcomeBuilder(input.operation ?? 'updateDataBatch');
  input.batches.forEach((batch, index) => {
    const a1Ranges = batch.map((item) => item.range);
    if (index < input.batchesExecuted) {
      builder.request({ requestIndex: index, kind: VALUES_BATCH_UPDATE_REQUEST_KIND, a1Ranges, state: 'acknowledged' });
    } else if (index === input.batchesExecuted && failingChunkDispatched) {
      builder.request({
        requestIndex: index,
        kind: VALUES_BATCH_UPDATE_REQUEST_KIND,
        a1Ranges,
        state: failedRejected ? 'rejected' : 'unknown',
        ...(status !== undefined ? { httpStatus: status } : {}),
        causeSummary: failedRejected
          ? `rejected before application with HTTP ${status}`
          : status !== undefined
            ? `server error after send (HTTP ${status}); application state uncertain`
            : 'dispatched but no usable response came back',
      });
    } else {
      builder.request({
        requestIndex: index,
        kind: VALUES_BATCH_UPDATE_REQUEST_KIND,
        a1Ranges,
        state: 'not-attempted',
        causeSummary: 'not dispatched after the earlier failure',
      });
    }
  });
  if (input.gridGrowth) {
    builder.gridGrowth(input.gridGrowth.attempted, input.gridGrowth.completed, input.gridGrowth.details);
  }
  builder.phase(input.phase ?? 'values-write');
  if (input.gridGrowth?.attempted && !input.gridGrowth.completed) {
    // The derivation cannot see gridGrowth: an unconfirmed growth leaves a replay's own
    // growth effect uncertain, so be more conservative than the request-only derivation.
    builder.retryGuidance('verify-then-replay', 'The grid growth request was not confirmed before the failure, so a replay may grow the grid again. Check the worksheet dimensions before replaying.');
  }
  if (input.nonIdempotent && failingChunkDispatched && !failedRejected) {
    // unknownOutcomeGuidance locks never-blind-replay for non-idempotent operations: a hidden
    // application would be duplicated by an unconditional replay.
    builder.retryGuidance(unknownOutcomeGuidance(true), 'A request was dispatched without a usable acknowledgement and the operation is not idempotent, so replaying could duplicate a hidden application. Verify remote state before any replay.');
  }
  return builder.build();
}

/**
 * GoogleSheet helper class for CRUD and Batch operations
 *
 * @export
 * @class GoogleSheet
 */
export default class GoogleSheet {
  private sheets!: sheets_v4.Sheets;

  // Kept for raw Drive API calls (permissions, file copy/export). Structural type: both auth.JWT and
  // OAuth2Client satisfy it, and it sidesteps the v10/v11 dual-package mismatch. responseType is
  // what keeps files.export binary-safe instead of decoded into a string.
  private authClient?: { request<T = unknown>(opts: { url: string; method?: string; params?: Record<string, unknown>; data?: unknown; responseType?: 'arraybuffer' | 'blob' | 'json' | 'text' | 'stream' | 'unknown' }): Promise<{ data: T }> };

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
    const client = new auth.JWT({ email: client_email, key: private_key, scopes: [SHEETS_SCOPE, DRIVE_FILE_SCOPE] });
    this.authClient = client;
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
    this.authClient = client;
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
   * Fetch scoped worksheet metadata: the worksheet's properties, bounded grid data carrying
   * per-cell user-entered values, effective values, formatted values and data validation, and
   * the spreadsheet's named ranges for this sheet - in a single `spreadsheets.get` request
   * limited to the requested worksheet and (optionally) range.
   *
   * `getSpreadsheet()` deliberately does not return grid data; callers needing cell-level
   * formula or validation detail must go through here and should bound the request with
   * `range` (e.g. "A1:Z100") so a large sheet cannot blow up the response.
   *
   * @param {GoogleSheetCli.WorksheetMetadataOptions} [options={}]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.WorksheetMetadata>}
   * @memberof GoogleSheet
   */
  async getWorksheetMetadata(options: GoogleSheetCli.WorksheetMetadataOptions = {}, spreadsheetId?: string): Promise<GoogleSheetCli.WorksheetMetadata> {
    const title = options.worksheetTitle ?? this.worksheetTitle;
    if (!title) throw new Error('Option property "worksheetTitle" is required');
    const id = spreadsheetId || this.spreadsheetId;
    if (!id) throw new Error('Option property "spreadsheetId" is required');
    if (options.range && options.range.includes('!')) {
      throw new Error(`Range "${options.range}" must be plain A1 notation; the worksheet title is prepended automatically`);
    }

    const escaped = escapeWorksheetTitle(title);
    const requestedRange = options.range ? `${escaped}!${options.range}` : escaped;
    const response = await this.sheets.spreadsheets.get({
      spreadsheetId: id,
      ranges: [requestedRange],
      // Scoped field mask: worksheet properties, the bounded grid values with formula and
      // validation detail, and the spreadsheet-level named ranges. This is the only call in
      // the class that returns grid data.
      fields:
        'sheets.properties(sheetId,title,index,gridProperties),sheets.data(startRow,startColumn,rowData(values(userEnteredValue,effectiveValue,formattedValue,dataValidation))),namedRanges',
    });

    const sheet = response.data.sheets?.find((entry) => entry.properties?.title === title) ?? response.data.sheets?.[0];
    if (!sheet?.properties) throw new Error(`Worksheet "${title}" not found`);
    const sheetId = sheet.properties.sheetId;
    const namedRanges = (response.data.namedRanges ?? []).filter((namedRange) => namedRange.range?.sheetId === sheetId);

    return {
      spreadsheetId: id,
      worksheetTitle: title,
      properties: sheet.properties,
      gridData: sheet.data ?? [],
      namedRanges,
    };
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
   * Failures after dispatch throw a `GSheetError` whose message text is byte-identical to the
   * plain `Error` this used to throw, carrying the original error as `cause` and a typed
   * `mutationOutcome` report (see `./mutation-outcome`): dispatched chunks `acknowledged`, a
   * chunk refused by a usable 4xx response `rejected`, a 5xx-after-send or lost response
   * `unknown`, undispatched chunks `not-attempted`; grid growth is reported separately in
   * `gridGrowth` - the grid may have grown even when the value write failed. Local preflight
   * refusals (bad range, non-2D values, bounded range overflow, formula-overwrite conflicts)
   * throw `GSheetError` with code VALIDATION and a validate-phase outcome. A failure of the
   * `getSpreadsheet` planning read propagates unchanged.
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
        throw batchPreflightRefusal('Each update must have a valid non-empty range string');
      }
      const parsed = parseStrictA1Range(update.range);
      if (!Array.isArray(update.values) || !update.values.every(Array.isArray)) {
        throw batchPreflightRefusal(`Update values for range "${update.range}" must be a 2D array`);
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
          throw batchPreflightRefusal(
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
        throw batchPreflightRefusal(
          `Cannot overwrite existing formula(s) without overwriteFormulas=true. Conflicts detected: ${conflictDetails}`,
          `Cannot overwrite existing formula(s) without overwriteFormulas=true. Conflicts detected: ${conflicts.length} cell(s) at ${conflicts.slice(0, 5).map((c) => c.cell).join(', ')}`
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

    // The chunk plan is pure computation, so it happens before the growth dispatch: the
    // grid-growth failure report below can then list every values chunk as not-attempted.
    const maxChunkBytes = options.chunkByteSize || DEFAULT_CHUNK_BYTE_SIZE;
    const packedBatches = packUpdateBatches(updates, maxChunkBytes);
    const completedRanges: string[] = [];
    let batchesExecuted = 0;

    // Step 5: Grid growth in one batchUpdate, tracked separately from the value writes - the
    // grid may grow even when the value write that follows fails.
    let gridGrowthOutcome: { attempted: boolean; completed: boolean; details?: string } = {
      attempted: false,
      completed: false,
    };
    if (gridGrowthRequests.length > 0) {
      const { GSheetError, classifyError } = loadCliErrors();
      try {
        await this.sheets.spreadsheets.batchUpdate({
          spreadsheetId: targetSpreadsheetId,
          requestBody: { requests: gridGrowthRequests },
        });
        gridGrowthOutcome = {
          attempted: true,
          completed: true,
          details: `${gridGrowthRequests.length} appendDimension request(s) applied`,
        };
      } catch (error: unknown) {
        // The raw upstream message is kept verbatim - no wrapper prose is added on this path.
        const meta = classifyError(error);
        throw new GSheetError(meta.code, error instanceof Error ? error.message : String(error), {
          retryable: meta.retryable,
          retryAfterMs: meta.retryAfterMs,
          mutationOutcome: buildBatchUpdateMutationOutcome({
            batches: packedBatches,
            batchesExecuted: 0,
            cause: error,
            gridGrowth: {
              attempted: true,
              completed: false,
              details: `${gridGrowthRequests.length} appendDimension request(s) not confirmed`,
            },
            phase: 'grid-growth',
            failedChunkDispatched: false,
            operation: options.operation,
          }),
          cause: error,
        });
      }
    }

    // Step 6: Chunk and Execute Writes
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
        const { GSheetError, classifyError } = loadCliErrors();
        const meta = classifyError(error);
        throw new GSheetError(
          meta.code,
          `Batch update failed at batch ${b + 1}/${packedBatches.length} [${failedRanges}]: ${msg}. Successfully completed ranges: ${completedRanges.length > 0 ? completedRanges.join(', ') : 'none'}`,
          {
            retryable: meta.retryable,
            retryAfterMs: meta.retryAfterMs,
            mutationOutcome: buildBatchUpdateMutationOutcome({
              batches: packedBatches,
              batchesExecuted,
              cause: error,
              gridGrowth: gridGrowthOutcome,
              phase: 'values-write',
              operation: options.operation,
            }),
            cause: error,
          }
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
   * Clear cell values in an explicitly bounded range. Values-only: formatting, data
   * validation and every other cell property stay untouched. Cells that contain formulas
   * refuse to clear unless `overwriteFormulas` is set, and `dryRun` reports the plan without
   * touching a single cell.
   *
   * @param {GoogleSheetCli.ClearOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.ClearReceipt>}
   * @memberof GoogleSheet
   */
  async clearData(options: GoogleSheetCli.ClearOptions, spreadsheetId?: string): Promise<GoogleSheetCli.ClearReceipt> {
    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId || '';
    if (!targetSpreadsheetId) throw new ValidationError('Option property "spreadsheetId" is required');
    if (!options?.range || typeof options.range !== 'string') {
      throw new ValidationError('clearData requires an explicit "range"');
    }

    const parsed = parseStrictA1Range(options.range);
    // Whole columns, whole rows and bare worksheet names are refused: only a box with both
    // corners spelled out can be cleared.
    if (
      parsed.startCol === undefined ||
      parsed.startRow === undefined ||
      parsed.endCol === undefined ||
      parsed.endRow === undefined
    ) {
      throw new ValidationError(
        `clearData requires an explicitly bounded A1 range with both corners (e.g. "Sheet1!A1:D20"), got "${options.range}"`
      );
    }

    const { worksheetTitle } = this.resolveTargetWorksheet(options.range, options.worksheetTitle);

    // A clear never grows the grid, so a range reaching past it would fail mid-write; check
    // the bounds while there is still nothing to undo.
    const sheet = await this.getWorksheet(worksheetTitle, targetSpreadsheetId);
    const rowCount = sheet.properties?.gridProperties?.rowCount ?? 0;
    const columnCount = sheet.properties?.gridProperties?.columnCount ?? 0;
    if (parsed.endRow > rowCount || parsed.endCol > columnCount) {
      throw new ValidationError(`Clear range "${options.range}" reaches past the worksheet grid (${rowCount}x${columnCount})`);
    }

    const rows = parsed.endRow - parsed.startRow + 1;
    const columns = parsed.endCol - parsed.startCol + 1;
    const range = formatBoundedA1Range(worksheetTitle, parsed.startCol, parsed.startRow, parsed.endCol, parsed.endRow);

    // Formula guard: read the exact target once, rendered as formulas, so a formula is seen
    // even when it evaluates to empty. The guard's input is one single-cell witness per
    // formula cell - the clear conceptually writes '' everywhere, so every formula cell is a
    // conflict - which keeps the guard O(formulas) instead of materializing the O(cells)
    // empty matrix the write used to carry. A malformed read refuses the clear outright:
    // without a verified look at the range, no cell in it may be wiped.
    const shouldInspectFormulas = Boolean(options.dryRun) || !options.overwriteFormulas;
    let conflicts: FormulaConflict[] = [];
    if (shouldInspectFormulas) {
      const [existing] = await this.getDataBatch(
        [range],
        { valueRenderOption: GoogleSheetCli.ValueRenderOption.FORMULA },
        targetSpreadsheetId
      );
      if (existing?.values !== undefined && !Array.isArray(existing.values)) {
        throw new ValidationError(
          `Clear preflight read of "${range}" returned malformed data; refusing to clear a range it could not inspect for formulas`
        );
      }
      const grid: GoogleSheetCli.RawData = existing?.values || [];
      const witnesses: { range: string; values: GoogleSheetCli.RawData }[] = [];
      for (let r = 0; r < grid.length; r++) {
        const row = grid[r] || [];
        for (let c = 0; c < row.length; c++) {
          if (extractFormulaText(row[c])) {
            witnesses.push({
              range: formatBoundedA1Range(
                worksheetTitle,
                parsed.startCol + c,
                parsed.startRow + r,
                parsed.startCol + c,
                parsed.startRow + r
              ),
              // The incoming clear value at that cell: an empty string.
              values: [['']],
            });
          }
        }
      }
      conflicts = findFormulaOverwrites(witnesses, [{ range, values: grid }]);
    }

    if (conflicts.length > 0 && !options.overwriteFormulas && !options.dryRun) {
      const conflictDetails = conflicts
        .slice(0, 5)
        .map((c) => `${c.cell} (existing: "${c.existingFormula}", incoming: ${JSON.stringify(c.incomingValue)})`)
        .join('; ');
      throw new ValidationError(
        `Cannot overwrite existing formula(s) without overwriteFormulas=true. Conflicts detected: ${conflictDetails}`,
        `Cannot overwrite existing formula(s) without overwriteFormulas=true. Conflicts detected: ${conflicts.length} cell(s) at ${conflicts.slice(0, 5).map((c) => c.cell).join(', ')}`
      );
    }

    // values.clear empties the values of the whole bounded range in one request while every
    // other cell property stays untouched; no payload matrix is ever built or chunked.
    if (!options.dryRun) {
      await this.sheets.spreadsheets.values.clear({ spreadsheetId: targetSpreadsheetId, range });
    }

    return {
      spreadsheetId: targetSpreadsheetId,
      worksheetTitle,
      range,
      rowsCleared: rows,
      columnsCleared: columns,
      cellsCleared: rows * columns,
      dryRun: Boolean(options.dryRun),
      batchesExecuted: options.dryRun ? 0 : 1,
      ...(options.dryRun
        ? { formulasOverwritten: Array.from(new Set(conflicts.map((c) => `${c.cell}: ${c.existingFormula}`))) }
        : {}),
    };
  }

  /**
   * Key-based upsert of header-mapped rows. A row whose key matches an existing row updates
   * the cells the input supplies - and only those; a row with an unseen key appends below the
   * table. Columns the input does not supply and cells outside the plan stay exactly as they
   * are, formulas included. Keys are type-aware: the string "001" never matches the number 1.
   *
   * The whole plan is computed before the first write, and one updateDataBatch call carries
   * every range, so its formula guard inspects the complete set up front and a conflict
   * anywhere refuses everything before anything is sent. Payloads beyond the chunk boundary
   * split into several requests; a failure part way through is thrown with the ranges already
   * written and is never replayed automatically.
   *
   * @param {GoogleSheetCli.RawData} rows - header-first input matrix as produced by resolveDataMatrix
   * @param {GoogleSheetCli.UpsertOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.UpsertReceipt>}
   * @memberof GoogleSheet
   */
  async upsert(
    rows: GoogleSheetCli.RawData,
    options: GoogleSheetCli.UpsertOptions,
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.UpsertReceipt> {
    const targetSpreadsheetId = spreadsheetId || this.spreadsheetId || '';
    if (!targetSpreadsheetId) throw new ValidationError('Option property "spreadsheetId" is required');
    if (!options || typeof options.key !== 'string' || options.key === '') {
      throw new ValidationError('upsert requires a "key" option naming exactly one existing header column');
    }

    const { parsed, worksheetTitle } = this.resolveTargetWorksheet(options.range, options.worksheetTitle);

    // An empty input has nothing to place anywhere: a no-op receipt, no reads, no writes.
    if (!Array.isArray(rows) || rows.length === 0 || (rows.length === 1 && (!rows[0] || rows[0].length === 0))) {
      return {
        spreadsheetId: targetSpreadsheetId,
        worksheetTitle,
        keyColumn: options.key,
        existingRows: 0,
        rowsAdded: 0,
        rowsUpdated: 0,
        rowsUnchanged: 0,
        addedRanges: [],
        updatedRanges: [],
        plannedRanges: [],
        dryRun: Boolean(options.dryRun),
        batchesExecuted: 0,
      };
    }

    const sheet = await this.getWorksheet(worksheetTitle, targetSpreadsheetId);
    const rowCount = sheet.properties?.gridProperties?.rowCount ?? 0;
    const columnCount = sheet.properties?.gridProperties?.columnCount ?? 0;
    const startRow = parsed?.startRow || 1;
    const startCol = parsed?.startCol || 1;
    // Reads clamp to the grid exactly the way values.batchGet clamps server side.
    const endRow = Math.min(parsed?.endRow || rowCount, rowCount);
    const endCol = Math.min(parsed?.endCol || columnCount, columnCount);
    if (startRow > rowCount || startCol > columnCount) {
      throw new ValidationError(
        `Upsert range "${options.range}" lies entirely outside the worksheet grid (${rowCount}x${columnCount})`
      );
    }

    const tableRange = formatBoundedA1Range(worksheetTitle, startCol, startRow, endCol, endRow);
    const [tableRead] = await this.getDataBatch(
      [tableRange],
      { valueRenderOption: GoogleSheetCli.ValueRenderOption.UNFORMATTED_VALUE },
      targetSpreadsheetId
    );

    // Pure planning: every bad input, ambiguous table and duplicate key refuses here, before
    // a single write is planned, let alone sent.
    const frame = buildTableFrame(tableRead?.values, { worksheetTitle, startRow, startCol, keyColumn: options.key });
    const input = parseUpsertInput(rows, frame, options.key);
    const plan = planUpsert(frame, input);

    // Below-bound preflight: a bounded read cannot see rows under its end row, so an append
    // there could overwrite unseen data, and a matching key farther below would go unseen as
    // a duplicate. Scan the table columns from just past the bound to the grid bottom,
    // rendered as formulas so even a formula evaluating to empty is caught, and refuse the
    // whole upsert - matched-row updates included - before anything is written. A range
    // ending at the grid bottom, bounded or not, has nothing unseen below it and skips this.
    if (endRow < rowCount) {
      const belowRange = formatBoundedA1Range(worksheetTitle, startCol, endRow + 1, endCol, rowCount);
      const [belowRead] = await this.getDataBatch(
        [belowRange],
        { valueRenderOption: GoogleSheetCli.ValueRenderOption.FORMULA },
        targetSpreadsheetId
      );
      if (belowRead?.values !== undefined && !Array.isArray(belowRead.values)) {
        throw new ValidationError(
          `Upsert preflight read of "${belowRange}" returned malformed data; refusing to upsert without a verified look below row ${endRow}`
        );
      }
      const belowCell = firstBelowBoundCell(belowRead?.values, { startRow: endRow + 1, startCol });
      if (belowCell) {
        // The refusal names the offending cell's address, the bound and the remedy - never
        // the cell's value, so error text does not disclose spreadsheet contents.
        throw new ValidationError(
          `Upsert refused: range "${options.range}" stops at row ${endRow} but "${worksheetTitle}" holds content below it, ` +
            `first at ${belowCell} inside the table columns. An appended row could overwrite that data, ` +
            'and a matching key farther below would be invisible to key matching, so nothing was written. ' +
            `Pass a range covering the whole table through row ${rowCount}, or remove the content below row ${endRow} first.`
        );
      }
    }

    const base = {
      spreadsheetId: targetSpreadsheetId,
      worksheetTitle,
      keyColumn: options.key,
      existingRows: frame.rows.length,
      rowsAdded: plan.addedRows,
      rowsUpdated: plan.updatedRows,
      rowsUnchanged: plan.unchangedRows,
      addedRanges: plan.addedRanges,
      updatedRanges: plan.updatedRanges,
      plannedRanges: plan.plannedRanges,
      dryRun: Boolean(options.dryRun),
    };

    if (plan.updates.length === 0) {
      return { ...base, batchesExecuted: 0 };
    }

    // One call carries the whole plan so the formula guard inside updateDataBatch inspects
    // every target before the first write; a conflict anywhere refuses everything.
    const receipt = await this.updateDataBatch(
      plan.updates,
      {
        dryRun: options.dryRun,
        overwriteFormulas: options.overwriteFormulas,
        valueInputOption: options.valueInputOption || GoogleSheetCli.ValueInputOption.RAW,
        operation: 'upsert',
      },
      targetSpreadsheetId
    );

    return {
      ...base,
      dryRun: receipt.dryRun,
      batchesExecuted: receipt.batchesExecuted,
      ...(options.dryRun
        ? {
            formulasOverwritten: Array.from(
              new Set((receipt.changes || []).flatMap((change) => change.formulasOverwritten || []))
            ),
          }
        : {}),
    };
  }

  /**
   * Resolve the worksheet a range-carrying operation targets and refuse a range whose
   * worksheet name contradicts the worksheetTitle flag.
   */
  private resolveTargetWorksheet(
    range: string | undefined,
    worksheetTitleFlag: string | null | undefined
  ): { parsed?: ParsedA1Range; worksheetTitle: string } {
    const flagTitle = worksheetTitleFlag || this.worksheetTitle || '';
    let parsed: ParsedA1Range | undefined;
    if (range !== undefined && range !== '') {
      parsed = parseStrictA1Range(range);
      if (parsed.worksheetTitle && flagTitle && parsed.worksheetTitle !== flagTitle) {
        throw new Error(
          `Conflicting worksheet: range "${range}" names "${parsed.worksheetTitle}" but worksheetTitle is "${flagTitle}"`
        );
      }
    }
    const worksheetTitle = parsed?.worksheetTitle || flagTitle;
    if (!worksheetTitle) throw new Error('Option property "worksheetTitle" is required');
    return { parsed, worksheetTitle };
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
   * Locate cells matching a condition and return their A1 coordinates.
   * Pure read: one values.get over the bounded range, then a client-side scan.
   * `matchCount` is the true total; `matches` is capped at `limit`.
   *
   * @param {GoogleSheetCli.FindOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.FindResult>}
   * @memberof GoogleSheet
   */
  async findData(options: GoogleSheetCli.FindOptions, spreadsheetId?: string): Promise<GoogleSheetCli.FindResult> {
    const modes = [options.equals !== undefined, options.contains !== undefined, options.regex !== undefined].filter(Boolean).length;
    if (modes !== 1) {
      throw new Error('Exactly one match mode is required: equals, contains or regex');
    }
    if (options.column !== undefined && options.header !== undefined) {
      throw new Error('Options "column" and "header" are mutually exclusive');
    }

    const ignoreCase = options.ignoreCase !== false;
    // An empty contains/regex needle would match every cell in the range, which
    // is almost never the intent and floods the result; reject it. An empty
    // equals needle is meaningful - it locates empty cells - so it stays legal.
    if (options.contains !== undefined && String(options.contains) === '') {
      throw new Error('Option "contains" requires a non-empty value');
    }
    if (options.regex !== undefined && String(options.regex) === '') {
      throw new Error('Option "regex" requires a non-empty pattern');
    }
    let regex: RegExp | undefined;
    if (options.regex !== undefined) {
      try {
        regex = new RegExp(options.regex, ignoreCase ? 'i' : '');
      } catch (err) {
        throw new Error(`Invalid regex "${options.regex}": ${(err as Error).message}`);
      }
    }

    const needle = options.equals ?? options.contains;
    const needleText = needle === undefined ? undefined : ignoreCase ? String(needle).toLowerCase() : String(needle);
    const needleNumber = options.equals !== undefined && String(needle).trim() !== '' && Number.isFinite(Number(needle)) ? Number(needle) : undefined;

    const matchesCell = (cell: string | number | boolean | null): boolean => {
      const empty = cell === null || cell === undefined || cell === '';
      if (empty) return options.equals === '';
      if (regex) return regex.test(String(cell));
      if (options.equals !== undefined) {
        if (needleNumber !== undefined && typeof cell === 'number') return cell === needleNumber;
        const cellText = ignoreCase ? String(cell).toLowerCase() : String(cell);
        return cellText === needleText;
      }
      const cellText = ignoreCase ? String(cell).toLowerCase() : String(cell);
      return cellText.includes(needleText as string);
    };

    const { rawData, range } = await this.getData(
      {
        worksheetTitle: options.worksheetTitle,
        range: options.range,
        valueRenderOption: options.valueRenderOption,
        dateTimeRenderOption: options.dateTimeRenderOption,
        rawOnly: true,
      },
      spreadsheetId
    );

    // The API reports the range it actually returned; coordinates are anchored to it.
    const parsed = parseStrictA1Range(range || options.range || 'A1');
    const baseRow = parsed.startRow ?? 1;
    const baseCol = parsed.startCol ?? 1;

    let restrictCol: number | undefined;
    let headerRowOffset = 0;
    if (options.column !== undefined) {
      restrictCol = a1ToCol(options.column);
    } else if (options.header !== undefined) {
      const headerText = ignoreCase ? options.header.toLowerCase() : options.header;
      const firstRow = rawData[0] || [];
      const idx = firstRow.findIndex((cell) => {
        if (cell === null || cell === undefined) return false;
        const text = ignoreCase ? String(cell).toLowerCase() : String(cell);
        return text === headerText;
      });
      if (idx < 0) {
        throw new Error(`Header "${options.header}" not found in the first scanned row of ${range}`);
      }
      restrictCol = baseCol + idx;
      headerRowOffset = 1;
    }

    // An explicit limit is honored as given, including 0 ("count only, return
    // nothing"); only an absent or negative limit falls back to the default.
    const limit = options.limit !== undefined && options.limit >= 0 ? options.limit : 100;
    const all: GoogleSheetCli.FindMatch[] = [];
    const seenRows = new Set<number>();

    for (let r = headerRowOffset; r < rawData.length; r++) {
      const row = rawData[r] || [];
      for (let c = 0; c < row.length; c++) {
        const absCol = baseCol + c;
        if (restrictCol !== undefined && absCol !== restrictCol) continue;
        const cell = row[c];
        if (!matchesCell(cell)) continue;
        const absRow = baseRow + r;
        if (options.byRow) {
          if (seenRows.has(absRow)) continue;
          seenRows.add(absRow);
        }
        all.push({
          a1: formatBoundedA1Range(parsed.worksheetTitle || options.worksheetTitle || undefined, absCol, absRow, absCol, absRow),
          row: absRow,
          column: absCol,
          columnLetter: colToA1(absCol),
          value: String(cell),
          ...(options.byRow ? { rowValues: row } : {}),
        });
      }
    }

    return {
      spreadsheetId: spreadsheetId || this.spreadsheetId || '',
      range: range || options.range || '',
      matchCount: all.length,
      truncated: all.length > limit,
      matches: all.slice(0, limit),
    };
  }

  /**
   * Apply cell formatting (text style, colors, alignment, wrap, number format,
   * borders) or clear formatting over one or more ranges. The generated field
   * mask is confined to userEnteredFormat.*, so values cannot be overwritten.
   * dryRun returns the exact request bodies without sending them.
   *
   * @param {FormatSpec} spec
   * @param {boolean} [dryRun]
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.FormatReceipt>}
   * @memberof GoogleSheet
   */
  async formatCells(spec: FormatSpec, dryRun = false, spreadsheetId?: string): Promise<GoogleSheetCli.FormatReceipt> {
    if (!spec || !Array.isArray(spec.ranges) || spec.ranges.length === 0) {
      throw new Error('formatCells requires a non-empty "ranges" array');
    }

    // Resolve every worksheet title the ranges touch to a sheetId. Ranges without
    // an explicit title fall back to spec.worksheetTitle, then to the instance's.
    const defaultTitle = spec.worksheetTitle || this.worksheetTitle || undefined;
    const titles = new Set<string>();
    for (const a1 of spec.ranges) {
      const parsed = parseStrictA1Range(a1);
      const title = parsed.worksheetTitle ?? defaultTitle;
      if (!title) {
        throw new Error(`Range "${a1}" has no worksheet title and no worksheetTitle was provided`);
      }
      titles.add(title);
    }

    const sheetIds = new Map<string, number>();
    for (const title of titles) {
      const sheet = await this.getWorksheet(title, spreadsheetId);
      sheetIds.set(title, sheet.properties?.sheetId ?? 0);
    }
    const sheetIdFor = (title: string | undefined): number => {
      const resolved = title ?? defaultTitle;
      const id = resolved ? sheetIds.get(resolved) : undefined;
      if (id === undefined) throw new Error(`Could not resolve worksheet "${resolved}" to a sheetId`);
      return id;
    };

    const { requests, fields } = buildFormatRequests(spec, sheetIdFor);

    if (!dryRun) {
      await this.batchUpdateSpreadsheet(requests, spreadsheetId);
    }

    return {
      spreadsheetId: spreadsheetId || this.spreadsheetId || '',
      ranges: spec.ranges,
      requestCount: requests.length,
      fields,
      dryRun,
      ...(dryRun ? { requests } : {}),
    };
  }

  /**
   * Merge or unmerge cells over a bounded range.
   * Google keeps the top-left value on merge; other values in the range are hidden.
   *
   * @param {GoogleSheetCli.MergeOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.FormatReceipt>}
   * @memberof GoogleSheet
   */
  async setMerge(options: GoogleSheetCli.MergeOptions, spreadsheetId?: string): Promise<GoogleSheetCli.FormatReceipt> {
    if (!options || !options.range) {
      throw new Error('setMerge requires a "range"');
    }
    const defaultTitle = options.worksheetTitle || this.worksheetTitle || undefined;
    const parsed = parseStrictA1Range(options.range);
    const title = parsed.worksheetTitle ?? defaultTitle;
    if (!title) {
      throw new Error(`Range "${options.range}" has no worksheet title and no worksheetTitle was provided`);
    }
    const sheet = await this.getWorksheet(title, spreadsheetId);
    const sheetId = sheet.properties?.sheetId ?? 0;

    const request = buildMergeRequest(options.range, options.mergeType, Boolean(options.unmerge), () => sheetId);

    if (!options.dryRun) {
      await this.batchUpdateSpreadsheet([request], spreadsheetId);
    }

    return {
      spreadsheetId: spreadsheetId || this.spreadsheetId || '',
      ranges: [options.range],
      requestCount: 1,
      fields: [options.unmerge ? 'unmergeCells' : `mergeCells.${options.mergeType || 'MERGE_ALL'}`],
      dryRun: Boolean(options.dryRun),
      ...(options.dryRun ? { requests: [request] } : {}),
    };
  }

  /**
   * Structural dimension mutation: insert, delete, hide/unhide, or resize rows/columns.
   * One batchUpdate request per call. `start`/`count` are 1-based inclusive on the
   * CLI and converted to the API's 0-based, end-exclusive DimensionRange here.
   *
   * @param {'insert'|'delete'|'hide'|'resize'} action
   * @param {GoogleSheetCli.DimensionOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.DimensionReceipt>}
   * @memberof GoogleSheet
   */
  async mutateDimension(
    action: 'insert' | 'delete' | 'hide' | 'resize',
    options: GoogleSheetCli.DimensionOptions,
    spreadsheetId?: string
  ): Promise<GoogleSheetCli.DimensionReceipt> {
    if (!options || (options.dimension !== 'ROWS' && options.dimension !== 'COLUMNS')) {
      throw new Error('mutateDimension requires dimension "ROWS" or "COLUMNS"');
    }
    const count = options.count ?? 1;
    if (!Number.isInteger(options.start) || options.start < 1) {
      throw new Error(`mutateDimension requires a 1-based start >= 1, got "${options.start}"`);
    }
    if (!Number.isInteger(count) || count < 1) {
      throw new Error(`mutateDimension requires a count >= 1, got "${count}"`);
    }
    if (action === 'resize') {
      if (options.pixels !== undefined && options.autoResize) {
        throw new Error('resize accepts either "pixels" or "autoResize", not both');
      }
      if (options.pixels === undefined && !options.autoResize) {
        throw new Error('resize requires either "pixels" or "autoResize"');
      }
      if (options.pixels !== undefined && (!Number.isFinite(options.pixels) || options.pixels <= 0)) {
        throw new Error(`resize requires a positive pixel size, got "${options.pixels}"`);
      }
    }

    const title = options.worksheetTitle || this.worksheetTitle || undefined;
    if (!title) {
      throw new Error('Option property "worksheetTitle" is required');
    }
    const sheet = await this.getWorksheet(title, spreadsheetId);
    const sheetId = sheet.properties?.sheetId ?? 0;

    // 1-based inclusive CLI -> 0-based, end-exclusive API
    const dimensionRange: sheets_v4.Schema$DimensionRange = {
      sheetId,
      dimension: options.dimension,
      startIndex: options.start - 1,
      endIndex: options.start - 1 + count,
    };

    let request: sheets_v4.Schema$Request;
    if (action === 'insert') {
      request = { insertDimension: { range: dimensionRange, inheritFromBefore: Boolean(options.inheritFromBefore) } };
    } else if (action === 'delete') {
      request = { deleteDimension: { range: dimensionRange } };
    } else if (action === 'hide') {
      request = {
        updateDimensionProperties: {
          range: dimensionRange,
          properties: { hiddenByUser: !options.unhide },
          fields: 'hiddenByUser',
        },
      };
    } else if (options.autoResize) {
      request = { autoResizeDimensions: { dimensions: dimensionRange } };
    } else {
      request = {
        updateDimensionProperties: {
          range: dimensionRange,
          properties: { pixelSize: options.pixels },
          fields: 'pixelSize',
        },
      };
    }

    // A delete dry-run previews the values about to be lost - the only preview
    // that matters for a destructive op.
    let affectedValues: GoogleSheetCli.RawData | undefined;
    if (action === 'delete' && options.dryRun) {
      const dataOptions: GoogleSheetCli.QueryOptions =
        options.dimension === 'ROWS'
          ? { worksheetTitle: title, minRow: options.start, maxRow: options.start + count - 1, rawOnly: true }
          : { worksheetTitle: title, minCol: options.start, maxCol: options.start + count - 1, rawOnly: true };
      const { rawData } = await this.getData(dataOptions, spreadsheetId);
      affectedValues = rawData;
    }

    if (!options.dryRun) {
      await this.batchUpdateSpreadsheet([request], spreadsheetId);
    }

    return {
      spreadsheetId: spreadsheetId || this.spreadsheetId || '',
      worksheetTitle: title,
      dimension: options.dimension,
      start: options.start,
      count,
      requestCount: 1,
      dryRun: Boolean(options.dryRun),
      ...(options.dryRun ? { requests: [request] } : {}),
      ...(affectedValues !== undefined ? { affectedValues } : {}),
    };
  }

  /**
   * Freeze or unfreeze rows/columns on a worksheet.
   * At least one of rows/columns is required; 0 unfreezes that axis.
   *
   * @param {GoogleSheetCli.FreezeOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.FormatReceipt>}
   * @memberof GoogleSheet
   */
  async setFrozen(options: GoogleSheetCli.FreezeOptions, spreadsheetId?: string): Promise<GoogleSheetCli.FormatReceipt> {
    const hasRows = options?.rows !== undefined;
    const hasCols = options?.columns !== undefined;
    if (!hasRows && !hasCols) {
      throw new Error('setFrozen requires at least one of "rows" or "columns"');
    }
    for (const [name, value] of [['rows', options.rows], ['columns', options.columns]] as const) {
      if (value !== undefined && (!Number.isInteger(value) || value < 0)) {
        throw new Error(`setFrozen "${name}" must be an integer >= 0, got "${value}"`);
      }
    }

    const title = options.worksheetTitle || this.worksheetTitle || undefined;
    if (!title) {
      throw new Error('Option property "worksheetTitle" is required');
    }
    const sheet = await this.getWorksheet(title, spreadsheetId);
    const sheetId = sheet.properties?.sheetId ?? 0;

    const gridProperties: sheets_v4.Schema$GridProperties = {};
    const fields: string[] = [];
    if (hasRows) {
      gridProperties.frozenRowCount = options.rows;
      fields.push('gridProperties.frozenRowCount');
    }
    if (hasCols) {
      gridProperties.frozenColumnCount = options.columns;
      fields.push('gridProperties.frozenColumnCount');
    }

    const request: sheets_v4.Schema$Request = {
      updateSheetProperties: {
        properties: { sheetId, gridProperties },
        fields: fields.join(','),
      },
    };

    if (!options.dryRun) {
      await this.batchUpdateSpreadsheet([request], spreadsheetId);
    }

    return {
      spreadsheetId: spreadsheetId || this.spreadsheetId || '',
      ranges: [title],
      requestCount: 1,
      fields,
      dryRun: Boolean(options.dryRun),
      ...(options.dryRun ? { requests: [request] } : {}),
    };
  }

  /**
   * Share the spreadsheet with users, a domain, or anyone with the link.
   * Drive API permissions.create — requires the drive.file scope (re-run `auth:login`
   * on tokens issued before it was added).
   *
   * @param {GoogleSheetCli.ShareOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.ShareResult>}
   * @memberof GoogleSheet
   */
  async shareSpreadsheet(options: GoogleSheetCli.ShareOptions, spreadsheetId?: string): Promise<GoogleSheetCli.ShareResult> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');

    const role = options.role || 'reader';
    const grantees: { type: string; emailAddress?: string; domain?: string }[] = [];
    if (options.anyone) {
      grantees.push({ type: options.type || 'anyone' });
    } else if (options.domain) {
      grantees.push({ type: options.type || 'domain', domain: options.domain });
    } else if (options.emails && options.emails.length > 0) {
      const type = options.type || 'user';
      for (const emailAddress of options.emails) grantees.push({ type, emailAddress });
    } else {
      throw new Error('shareSpreadsheet requires one of "emails", "domain", or "anyone"');
    }

    const granted: GoogleSheetCli.Permission[] = [];
    for (const grantee of grantees) {
      const response = await this.driveRequest<{ id: string; type: string; role: string; emailAddress?: string; domain?: string; displayName?: string }>({
        method: 'POST',
        url: `${DRIVE_API_BASE}/files/${fileId}/permissions`,
        params: {
          sendNotificationEmail: Boolean(options.notify),
          ...(options.message ? { emailMessage: options.message } : {}),
          fields: 'id,type,role,emailAddress,domain,displayName',
        },
        data: { role, ...grantee },
      });
      granted.push(response);
    }

    return { spreadsheetId: fileId, granted };
  }

  /**
   * List the sharing permissions on the spreadsheet.
   *
   * @param {string} [spreadsheetId]
   * @returns {Promise<GoogleSheetCli.Permission[]>}
   * @memberof GoogleSheet
   */
  async listPermissions(spreadsheetId?: string): Promise<GoogleSheetCli.Permission[]> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');
    const response = await this.driveRequest<{ permissions?: GoogleSheetCli.Permission[] }>({
      method: 'GET',
      url: `${DRIVE_API_BASE}/files/${fileId}/permissions`,
      params: { fields: 'permissions(id,type,role,emailAddress,domain,displayName)' },
    });
    return response.permissions || [];
  }

  /**
   * Remove a sharing permission, by id or by grantee email.
   *
   * @param {GoogleSheetCli.UnshareOptions} options
   * @param {string} [spreadsheetId]
   * @returns {Promise<{ spreadsheetId: string; permissionId: string; removed: boolean }>}
   * @memberof GoogleSheet
   */
  async unshareSpreadsheet(
    options: GoogleSheetCli.UnshareOptions,
    spreadsheetId?: string
  ): Promise<{ spreadsheetId: string; permissionId: string; removed: boolean }> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');

    let permissionId = options.permissionId;
    if (!permissionId && options.email) {
      const permissions = await this.listPermissions(fileId);
      const match = permissions.find((p) => p.emailAddress?.toLowerCase() === options.email!.toLowerCase());
      if (!match) throw new Error(`No permission found for "${options.email}" on this spreadsheet`);
      permissionId = match.id;
    }
    if (!permissionId) throw new Error('unshareSpreadsheet requires "permissionId" or "email"');

    await this.driveRequest<void>({
      method: 'DELETE',
      url: `${DRIVE_API_BASE}/files/${fileId}/permissions/${permissionId}`,
    });
    return { spreadsheetId: fileId, permissionId, removed: true };
  }

  /**
   * Copy a whole spreadsheet into a new one through the Drive API files.copy endpoint.
   * Requires the drive.file scope: only files this app created or has opened are visible.
   * The source is left untouched and sharing grants are not duplicated; a POST that fails
   * ambiguously is never replayed, because a copy that actually went through twice would
   * leave two spreadsheets behind.
   *
   * @param {GoogleSheetCli.CopySpreadsheetOptions} [options={}] - new title; Drive keeps the source title when omitted
   * @param {string} [spreadsheetId] - the spreadsheet to copy; defaults to the instance id
   * @returns {Promise<GoogleSheetCli.CopySpreadsheetResult>} identifiers and metadata of the copy Drive actually created
   * @memberof GoogleSheet
   */
  async copySpreadsheet(options: GoogleSheetCli.CopySpreadsheetOptions = {}, spreadsheetId?: string): Promise<GoogleSheetCli.CopySpreadsheetResult> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');
    if (options.title !== undefined && (typeof options.title !== 'string' || !options.title.trim())) {
      throw new Error('copySpreadsheet requires a non-empty "title" when one is provided');
    }

    // No request body when no title is asked for: Drive then names the copy after the source.
    const file = await this.driveRequest<GoogleSheetCli.DriveFileMetadata>({
      method: 'POST',
      url: `${DRIVE_API_BASE}/files/${encodeURIComponent(fileId)}/copy`,
      params: { fields: 'id,name,mimeType,webViewLink' },
      ...(options.title !== undefined ? { data: { name: options.title } } : {}),
    });
    if (!file?.id) throw new Error(`Drive API returned no id for the copy of "${fileId}"`);

    return {
      sourceSpreadsheetId: fileId,
      spreadsheetId: file.id,
      title: file.name || '',
      mimeType: file.mimeType || '',
      ...(file.webViewLink ? { webViewLink: file.webViewLink } : {}),
    };
  }

  /**
   * Copy one worksheet into an explicit destination spreadsheet through Sheets
   * spreadsheets.sheets.copyTo. The copy carries the source title; when the destination —
   * the same spreadsheet included — already uses it, Google assigns a unique
   * "<title> Copy" title and the receipt reports the title the server actually chose.
   * Values and formulas are carried over; the source worksheet is left untouched.
   *
   * @param {GoogleSheetCli.CopyWorksheetOptions} options - source worksheet title and the destination spreadsheet id
   * @param {string} [spreadsheetId] - the source spreadsheet; defaults to the instance id
   * @returns {Promise<GoogleSheetCli.CopyWorksheetResult>} the SheetProperties the destination reports for the copy
   * @memberof GoogleSheet
   */
  async copyWorksheet(options: GoogleSheetCli.CopyWorksheetOptions, spreadsheetId?: string): Promise<GoogleSheetCli.CopyWorksheetResult> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');
    if (typeof options.worksheetTitle !== 'string' || !options.worksheetTitle.trim()) {
      throw new Error('copyWorksheet requires a non-empty "worksheetTitle"');
    }
    if (typeof options.destinationSpreadsheetId !== 'string' || !options.destinationSpreadsheetId.trim()) {
      throw new Error('copyWorksheet requires a non-empty "destinationSpreadsheetId"');
    }

    // Resolves the title to a sheetId and proves the source worksheet exists before anything
    // is written anywhere.
    const sheet = await this.getWorksheet(options.worksheetTitle, fileId);
    const sheetId = sheet.properties?.sheetId;
    if (typeof sheetId !== 'number') throw new Error(`Worksheet "${options.worksheetTitle}" has no sheet id to copy`);

    const response = await this.sheets.spreadsheets.sheets.copyTo({
      spreadsheetId: fileId,
      sheetId,
      requestBody: { destinationSpreadsheetId: options.destinationSpreadsheetId },
    });
    const properties = response.data;
    if (typeof properties?.sheetId !== 'number') {
      throw new Error(`Sheets API returned no sheet id for the copy of "${options.worksheetTitle}"`);
    }

    return {
      spreadsheetId: fileId,
      worksheetTitle: sheet.properties?.title || options.worksheetTitle,
      destinationSpreadsheetId: options.destinationSpreadsheetId,
      sheetId: properties.sheetId,
      title: properties.title || options.worksheetTitle,
      ...(typeof properties.index === 'number' ? { index: properties.index } : {}),
    };
  }

  /**
   * Export a spreadsheet as PDF or XLSX bytes through the Drive API files.export endpoint.
   * Requires the drive.file scope, and Drive refuses to export more than 10 MB. The bytes
   * come back raw — responseType arraybuffer, no text decoding — so what a caller writes
   * to disk is exactly what Google sent.
   *
   * @param {GoogleSheetCli.ExportOptions} options - the export format
   * @param {string} [spreadsheetId] - the spreadsheet to export; defaults to the instance id
   * @returns {Promise<GoogleSheetCli.ExportResult>} mime type and raw bytes; never a decoded string
   * @memberof GoogleSheet
   */
  async exportSpreadsheet(options: GoogleSheetCli.ExportOptions, spreadsheetId?: string): Promise<GoogleSheetCli.ExportResult> {
    const fileId = spreadsheetId || this.spreadsheetId;
    if (!fileId) throw new Error('Option property "spreadsheetId" is required');
    const mimeType = EXPORT_MIME_TYPES[options.format];
    if (!mimeType) throw new Error('exportSpreadsheet requires format "pdf" or "xlsx"');

    // gaxios 7 hands back an ArrayBuffer for responseType arraybuffer; older shapes may give
    // a Buffer. Both convert losslessly, anything else is not an export payload.
    const raw = await this.driveRequest<Buffer | ArrayBuffer>({
      method: 'GET',
      url: `${DRIVE_API_BASE}/files/${encodeURIComponent(fileId)}/export`,
      params: { mimeType },
      responseType: 'arraybuffer',
    });
    const bytes = Buffer.isBuffer(raw) ? raw : raw instanceof ArrayBuffer ? Buffer.from(raw) : undefined;
    if (!bytes) throw new Error('Drive API returned an unexpected export payload');

    return { spreadsheetId: fileId, mimeType, bytes, byteLength: bytes.length };
  }

  /**
   * List the spreadsheets visible to this app through the Drive API files.list endpoint.
   * Requires the drive.file scope: only files this application created or has opened are
   * visible, and the result always carries a visibilityNote saying so, because an empty
   * listing does not prove a spreadsheet is absent. Query values are escaped into the Drive
   * query language, never interpolated raw. Duplicate titles are kept as separate rows and
   * nothing here ever selects a spreadsheet by itself — discovery hands back ids, a caller
   * picks one.
   *
   * @param {GoogleSheetCli.ListSpreadsheetsOptions} [options={}] - name filter (partial by default, whole-title with exact), page size and continuation token
   * @returns {Promise<GoogleSheetCli.ListSpreadsheetsResult>} one page of files, the next token when more pages remain, and the visibility note
   * @memberof GoogleSheet
   */
  async listSpreadsheets(options: GoogleSheetCli.ListSpreadsheetsOptions = {}): Promise<GoogleSheetCli.ListSpreadsheetsResult> {
    const { name, exact = false, pageSize, pageToken } = options;
    if (name !== undefined && (typeof name !== 'string' || name.length === 0)) {
      throw new Error('listSpreadsheets requires a non-empty "name" when one is provided');
    }
    if (pageToken !== undefined && (typeof pageToken !== 'string' || pageToken.length === 0)) {
      throw new Error('listSpreadsheets requires a non-empty "pageToken" when one is provided');
    }

    const params: Record<string, unknown> = {
      q: driveSpreadsheetsQuery(name, exact),
      pageSize: resolveListPageSize(pageSize),
      fields: 'files(id,name,mimeType,modifiedTime),nextPageToken',
      spaces: 'drive',
      ...(pageToken !== undefined ? { pageToken } : {}),
    };

    const response = await this.driveRequest<DriveFilesListResponse>({
      method: 'GET',
      url: `${DRIVE_API_BASE}/files`,
      params,
    });

    type DriveFileRow = NonNullable<DriveFilesListResponse['files']>[number];
    const files = (response.files || [])
      .filter((file): file is DriveFileRow & { id: string } => typeof file?.id === 'string')
      .map((file) => ({
        id: file.id,
        name: typeof file.name === 'string' ? file.name : '',
        mimeType: typeof file.mimeType === 'string' ? file.mimeType : '',
        ...(typeof file.modifiedTime === 'string' ? { modifiedTime: file.modifiedTime } : {}),
      }));

    return {
      files,
      ...(response.nextPageToken ? { nextPageToken: response.nextPageToken } : {}),
      visibilityNote: DRIVE_FILE_VISIBILITY_NOTE,
    };
  }

  /**
   * One raw Drive API call through the stored auth client, with errors translated
   * into the two messages a caller can act on.
   */
  private async driveRequest<T>(opts: {
    url: string;
    method?: string;
    params?: Record<string, unknown>;
    data?: unknown;
    /** 'arraybuffer' keeps binary answers (files.export) as raw bytes instead of decoded text */
    responseType?: 'json' | 'arraybuffer';
  }): Promise<T> {
    if (!this.authClient) throw new Error('authorize() or authorizeOAuth() must run before Drive calls');
    try {
      const response = await this.authClient.request<T>(opts);
      return response.data;
    } catch (error) {
      // GaxiosError carries the HTTP status on error.response.status; narrow by shape
      // rather than instanceof so no transitive gaxios import is needed.
      const rawStatus =
        error && typeof error === 'object' && 'response' in error && error.response && typeof error.response === 'object' && 'status' in error.response
          ? error.response.status
          : undefined;
      const status = typeof rawStatus === 'number' ? rawStatus : undefined;
      const message = error instanceof Error ? error.message : 'Drive API request failed';
      if (status === 403) {
        throw this.driveStatusError(
          status,
          `Drive API denied the request (${message}). ` +
            'If this is an OAuth token issued before drive.file was added, run `gsheet auth:login` again to re-grant scopes.',
          error
        );
      }
      if (status === 404) {
        throw this.driveStatusError(
          status,
          `Drive API could not find the file (${message}). ` +
            'Under the drive.file scope only files this app created or has opened are visible — ' +
            'a pre-existing spreadsheet may need to be opened once through this app first.',
          error
        );
      }
      throw error;
    }
  }

  /**
   * Wrap a Drive HTTP failure without losing its shape: the HTTP status stays available as an
   * own enumerable `status` property (the CLI serializer maps it to the right exit envelope),
   * and the original error rides along as a non-enumerable `cause` so error chains survive
   * without leaking internals into JSON output. Plain Error subclassing only - no oclif, no
   * gaxios import.
   */
  private driveStatusError(status: number, message: string, cause: unknown): Error {
    const wrapped: Error & { status: number } = Object.assign(new Error(message), { status });
    Object.defineProperty(wrapped, 'cause', { value: cause, enumerable: false, writable: true, configurable: true });
    return wrapped;
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
