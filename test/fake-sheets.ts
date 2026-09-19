import { generateKeyPairSync } from 'crypto';
import { PassThrough } from 'stream';
import { URL } from 'url';

// The live builtin module objects, not the namespace copies an `import * as` would
// produce, so that swapping `.request` is visible to node-fetch inside gaxios.
const httpsModule = require('https');
const httpModule = require('http');

/**
 * An in-memory fake of the Google Sheets v4 REST API.
 *
 * It is installed over `https.request`, so everything above the socket runs for real:
 * the googleapis client, google-auth-library, gtoken, gaxios and node-fetch all build
 * their requests and parse the responses exactly as they do against Google. That makes
 * these tests a behavioral contract for `GoogleSheet` rather than an assertion about stubs.
 *
 * The A1 handling below is written from scratch instead of reusing `src/lib/utils.ts`,
 * so a bug in the production parser cannot hide itself inside the fake.
 */

// Both OAuth token endpoints Google has served service-account assertions from: gtoken 6
// (googleapis 118) posted to the first, gtoken 8 (@googleapis/sheets 14) posts to the second.
// Accepting both keeps the fake honest about which one the client under test actually used.
const TOKEN_URLS = ['https://www.googleapis.com/oauth2/v4/token', 'https://oauth2.googleapis.com/token'];
const SHEETS_HOST = 'sheets.googleapis.com';
const DRIVE_HOST = 'www.googleapis.com';
const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

export interface FakeWorksheet {
  sheetId: number;
  title: string;
  index: number;
  rowCount: number;
  columnCount: number;
  frozenRowCount?: number;
  frozenColumnCount?: number;
  /** merged regions as reported by spreadsheets.get, 0-based end-exclusive */
  merges?: { startRowIndex: number; endRowIndex: number; startColumnIndex: number; endColumnIndex: number }[];
  /** per-dimension properties keyed by 1-based row/column index */
  dimensionProps?: {
    rows?: Record<number, { hiddenByUser?: boolean; pixelSize?: number }>;
    columns?: Record<number, { hiddenByUser?: boolean; pixelSize?: number }>;
  };
  /** cell values keyed by `${row}:${col}`, both 1 based */
  cells: Map<string, string>;
}

export interface FakeSpreadsheet {
  spreadsheetId: string;
  title: string;
  sheets: FakeWorksheet[];
  developerMetadata?: Record<string, unknown>[];
}

export interface RecordedRequest {
  method: string;
  url: string;
  body?: any;
}

export interface FakeRange {
  title?: string;
  startRow?: number;
  startCol?: number;
  endRow?: number;
  endCol?: number;
  /** true when the range spelled out an end ("A1:B2"), false for a bare anchor cell ("A1") */
  bounded: boolean;
}

interface FakeResponse {
  status: number;
  body: any;
}

interface ResolvedRange {
  sheet: FakeWorksheet;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
  bounded: boolean;
}

/**
 * Convert a column label ("A", "AA") to a 1 based column number
 *
 * @param {string} label
 * @returns {number}
 */
export const letterToCol = (label: string): number =>
  label
    .toUpperCase()
    .split('')
    .reduce((acc, char) => acc * ALPHABET.length + (char.charCodeAt(0) - 64), 0);

/**
 * Convert a 1 based column number to a column label
 *
 * @param {number} col
 * @returns {string}
 */
export const colToLetter = (col: number): string => {
  let div = col;
  let label = '';
  while (div > 0) {
    const mod = (div - 1) % ALPHABET.length;
    label = `${ALPHABET[mod]}${label}`;
    div = Math.floor((div - mod - 1) / ALPHABET.length);
  }
  return label;
};

/**
 * Parse one A1 cell reference into row and column, either of which may be missing
 *
 * @param {string} ref
 * @returns {{ col?: number; row?: number }}
 */
const parseCell = (ref: string): { col?: number; row?: number } => {
  const match = ref.match(/^\$?([A-Za-z]+)?\$?(\d+)?$/);
  if (!match || (!match[1] && !match[2])) throw new Error(`Unable to parse range: ${ref}`);
  return {
    col: match[1] ? letterToCol(match[1]) : undefined,
    row: match[2] ? parseInt(match[2], 10) : undefined,
  };
};

/**
 * Parse an A1 range the way the Sheets API does, independently of the production parser
 *
 * @param {string} range
 * @returns {FakeRange}
 */
export const parseFakeRange = (range: string): FakeRange => {
  let title: string | undefined;
  let rest = range;

  if (range.startsWith("'")) {
    let i = 1;
    let unquoted = '';
    for (; i < range.length; i++) {
      if (range[i] === "'") {
        if (range[i + 1] === "'") {
          unquoted += "'";
          i++;
          continue;
        }
        break;
      }
      unquoted += range[i];
    }
    if (range[i] !== "'") throw new Error(`Unable to parse range: ${range}`);
    title = unquoted;
    rest = range.slice(i + 1);
    if (rest.startsWith('!')) rest = rest.slice(1);
    else if (rest.length) throw new Error(`Unable to parse range: ${range}`);
  } else if (range.startsWith('"')) {
    throw new Error(`Unable to parse range: ${range}`);
  } else {
    const bang = range.lastIndexOf('!');
    if (bang >= 0) {
      title = range.slice(0, bang);
      rest = range.slice(bang + 1);
    } else if (/^[A-Za-z]+\d*(:[A-Za-z]*\d*)?$/.test(range) || /^\d+:\d+$/.test(range)) {
      rest = range;
    } else {
      title = range;
      rest = '';
    }
  }

  if (!rest) return { title, bounded: false };

  const [from, to] = rest.split(':');
  const start = parseCell(from);
  const end = to === undefined ? undefined : parseCell(to);

  return {
    title,
    startRow: start.row,
    startCol: start.col,
    endRow: end ? end.row : start.row,
    endCol: end ? end.col : start.col,
    bounded: end !== undefined,
  };
};

/**
 * Whether A1 notation has to put a worksheet title in quotes to name it.
 *
 * Sheets used to quote the title of every range it echoed. Since September 2026 it quotes only
 * the titles that cannot be written bare, which is what this reproduces: anything outside
 * letters, digits and underscores or not starting with a letter (a space, punctuation, a
 * leading digit), a title that would read as a cell reference instead - a column label is at
 * most three letters, so `A1` and `ZZZ100` need quotes while `Sheet1` does not - and the two
 * boolean literals.
 *
 * Measured against the live API on 2026-09-07: a request for `'worksheet_remove_168...'!A1:Z1000`
 * came back as `worksheet_remove_168...!A1:Z1000`.
 *
 * @param {string} title
 * @returns {boolean}
 */
const needsQuoting = (title: string): boolean =>
  !/^[A-Za-z][A-Za-z0-9_]*$/.test(title) || /^[A-Za-z]{1,3}[0-9]+$/.test(title) || /^(TRUE|FALSE)$/i.test(title);

/**
 * Render a range back to the A1 notation the Sheets API echoes in its responses
 *
 * @param {string} title
 * @param {number} startRow
 * @param {number} startCol
 * @param {number} [endRow]
 * @param {number} [endCol]
 * @returns {string}
 */
const formatRange = (title: string, startRow: number, startCol: number, endRow?: number, endCol?: number): string => {
  const named = needsQuoting(title) ? `'${title.replace(/'/g, "''")}'` : title;
  const start = `${colToLetter(startCol)}${startRow}`;
  if (endRow === undefined && endCol === undefined) return `${named}!${start}`;
  return `${named}!${start}:${colToLetter(endCol as number)}${endRow}`;
};

/**
 * The rejection Google sends once a per-minute quota is spent. The wording is the one the live
 * suite reported on 2026-09-07; the 429 and `RESOURCE_EXHAUSTED` are what the API documents for
 * it (https://developers.google.com/workspace/sheets/api/limits). `error.message` is what the
 * client turns into the error a caller sees, so a test asserting on the message is asserting on
 * Google's.
 */
const QUOTA_REJECTION: FakeResponse = {
  status: 429,
  body: {
    error: {
      code: 429,
      message:
        "Quota exceeded for quota metric 'Read requests' and limit 'Read requests per minute per user' of service 'sheets.googleapis.com' for consumer 'project_number:000000000000'.",
      status: 'RESOURCE_EXHAUSTED',
    },
  },
};

export class FakeSheets {
  readonly spreadsheets = new Map<string, FakeSpreadsheet>();
  readonly requests: RecordedRequest[] = [];
  readonly credentials: { client_email: string; private_key: string };

  private nextSheetId = 100;
  private nextSpreadsheetId = 1;
  private nextPermissionId = 1;
  private quotaRejectionsLeft = 0;
  /** Drive permissions keyed by fileId */
  private permissions = new Map<string, { id: string; type: string; role: string; emailAddress?: string; domain?: string }[]>();
  private originalHttpsRequest?: Function;
  private originalHttpRequest?: Function;

  constructor() {
    // gtoken really signs the assertion, so the fake needs a real key
    const { privateKey } = generateKeyPairSync('rsa', {
      modulusLength: 2048,
      privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
      publicKeyEncoding: { type: 'spki', format: 'pem' },
    });
    this.credentials = { client_email: 'fake@fake-project.iam.gserviceaccount.com', private_key: privateKey as unknown as string };
  }

  /**
   * Route every https request through the fake API
   *
   * @returns {void}
   * @memberof FakeSheets
   */
  install(): void {
    if (this.originalHttpsRequest) return;
    this.originalHttpsRequest = httpsModule.request;
    this.originalHttpRequest = httpModule.request;
    httpsModule.request = this.request.bind(this);
    httpModule.request = this.request.bind(this);
  }

  /**
   * Restore the real https module
   *
   * @returns {void}
   * @memberof FakeSheets
   */
  uninstall(): void {
    if (!this.originalHttpsRequest) return;
    httpsModule.request = this.originalHttpsRequest;
    httpModule.request = this.originalHttpRequest;
    this.originalHttpsRequest = undefined;
    this.originalHttpRequest = undefined;
  }

  /**
   * Drop every spreadsheet and every recorded request
   *
   * @returns {void}
   * @memberof FakeSheets
   */
  reset(): void {
    this.spreadsheets.clear();
    this.requests.length = 0;
    this.quotaRejectionsLeft = 0;
    this.permissions.clear();
    this.nextPermissionId = 1;
  }

  /**
   * Answer the next `times` Sheets requests with the rejection Google sends once a per-minute
   * quota is spent - HTTP 429, status `RESOURCE_EXHAUSTED` - and serve the real handler again
   * after that. `Infinity` is a bucket that never refills.
   *
   * The token endpoint is deliberately left alone: a client that cannot authorize would never
   * reach the call under test.
   *
   * @param {number} [times=1]
   * @returns {void}
   * @memberof FakeSheets
   */
  rejectWithQuota(times = 1): void {
    this.quotaRejectionsLeft = times;
  }

  /**
   * Create a spreadsheet with the given worksheets
   *
   * @param {string} spreadsheetId
   * @param {string} title
   * @param {{ title: string; rowCount?: number; columnCount?: number }[]} [sheets=[]]
   * @returns {FakeSpreadsheet}
   * @memberof FakeSheets
   */
  addSpreadsheet(spreadsheetId: string, title: string, sheets: { title: string; rowCount?: number; columnCount?: number }[] = []): FakeSpreadsheet {
    const spreadsheet: FakeSpreadsheet = { spreadsheetId, title, sheets: [] };
    sheets.forEach((sheet) => {
      spreadsheet.sheets.push({
        sheetId: this.nextSheetId++,
        title: sheet.title,
        index: spreadsheet.sheets.length,
        rowCount: sheet.rowCount ?? 1000,
        columnCount: sheet.columnCount ?? 26,
        cells: new Map(),
      });
    });
    this.spreadsheets.set(spreadsheetId, spreadsheet);
    return spreadsheet;
  }

  /**
   * Look up a worksheet by title
   *
   * @param {string} spreadsheetId
   * @param {string} title
   * @returns {FakeWorksheet}
   * @memberof FakeSheets
   */
  worksheet(spreadsheetId: string, title: string): FakeWorksheet {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) throw new Error(`No fake spreadsheet "${spreadsheetId}"`);
    const sheet = spreadsheet.sheets.find((s) => s.title === title);
    if (!sheet) throw new Error(`No fake worksheet "${title}" in "${spreadsheetId}"`);
    return sheet;
  }

  /**
   * Read a single cell straight out of the fake, bypassing the API
   *
   * @param {string} spreadsheetId
   * @param {string} title
   * @param {string} a1
   * @returns {string}
   * @memberof FakeSheets
   */
  cell(spreadsheetId: string, title: string, a1: string): string {
    const { startRow, startCol } = parseFakeRange(a1);
    return this.worksheet(spreadsheetId, title).cells.get(`${startRow}:${startCol}`) ?? '';
  }

  /**
   * Fill a block of cells straight into the fake, bypassing the API
   *
   * @param {string} spreadsheetId
   * @param {string} title
   * @param {string} a1
   * @param {any[][]} values
   * @returns {void}
   * @memberof FakeSheets
   */
  setCells(spreadsheetId: string, title: string, a1: string, values: any[][]): void {
    const { startRow = 1, startCol = 1 } = parseFakeRange(a1);
    const sheet = this.worksheet(spreadsheetId, title);
    values.forEach((row, r) => {
      row.forEach((value, c) => {
        if (value === null || value === undefined) return;
        sheet.cells.set(`${startRow + r}:${startCol + c}`, String(value));
      });
    });
  }

  /**
   * Stand-in for `https.request`, answering out of the in-memory state
   *
   * @param {*} options
   * @param {*} [callback]
   * @returns {*}
   * @memberof FakeSheets
   */
  private request(urlOrOptions: any, optionsOrCallback?: any, maybeCallback?: any): any {
    // `http.request` has two documented signatures, `(options[, callback])` and
    // `(url[, options][, callback])`. gaxios 5 (googleapis 118) used the first, gaxios 7
    // (@googleapis/sheets 14) uses the second, so the fake has to take both or it silently
    // stops intercepting whichever client it was not written against.
    const fromUrl = typeof urlOrOptions === 'string' || urlOrOptions instanceof URL;
    const base = fromUrl ? new URL(String(urlOrOptions)) : undefined;
    const options: any = (fromUrl ? (typeof optionsOrCallback === 'function' ? undefined : optionsOrCallback) : urlOrOptions) || {};
    const callback = [optionsOrCallback, maybeCallback].find((argument) => typeof argument === 'function');

    const req: any = new PassThrough();
    const chunks: Buffer[] = [];
    req.abort = () => undefined;
    req.setTimeout = () => req;
    req.setNoDelay = () => req;
    req.setSocketKeepAlive = () => req;
    req.flushHeaders = () => undefined;

    // Node merges the two: the URL supplies the defaults, an explicit option overrides it.
    const protocol = options.protocol || base?.protocol || 'https:';
    const host = options.hostname || options.host || base?.host || 'localhost';
    const path = options.path || (base ? `${base.pathname}${base.search}` : '/');
    const url = `${protocol}//${host}${path}`;
    const method = (options.method || 'GET').toUpperCase();

    req.on('data', (chunk: Buffer) => chunks.push(Buffer.from(chunk)));
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      let answer: FakeResponse;
      try {
        answer = this.handle(method, url, raw);
      } catch (error) {
        answer = { status: 500, body: { error: { code: 500, message: (error as Error).message, status: 'INTERNAL' } } };
      }
      const payload = Buffer.from(JSON.stringify(answer.body), 'utf8');
      const res: any = new PassThrough();
      res.statusCode = answer.status;
      res.statusMessage = answer.status === 200 ? 'OK' : 'Bad Request';
      res.headers = { 'content-type': 'application/json; charset=UTF-8', 'content-length': String(payload.length) };
      res.rawHeaders = ['content-type', 'application/json; charset=UTF-8', 'content-length', String(payload.length)];
      if (callback) callback(res);
      req.emit('response', res);
      res.end(payload);
    });

    return req;
  }

  /**
   * Route one request to the matching Sheets API handler
   *
   * @param {string} method
   * @param {string} url
   * @param {string} raw
   * @returns {FakeResponse}
   * @memberof FakeSheets
   */
  private handle(method: string, url: string, raw: string): FakeResponse {
    const parsed = new URL(url);
    let body: any;
    if (raw && raw.trim().startsWith('{')) body = JSON.parse(raw);
    else if (raw) body = raw;
    this.requests.push({ method, url, body });

    if (TOKEN_URLS.some((tokenUrl) => url.startsWith(tokenUrl))) {
      return { status: 200, body: { access_token: 'fake-access-token', expires_in: 3600, token_type: 'Bearer' } };
    }

    if (parsed.hostname === DRIVE_HOST && parsed.pathname.startsWith('/drive/v3/')) {
      return this.driveHandle(method, parsed, body);
    }

    if (parsed.hostname !== SHEETS_HOST) {
      return { status: 404, body: { error: { code: 404, message: `Unexpected host ${parsed.hostname}`, status: 'NOT_FOUND' } } };
    }

    if (this.quotaRejectionsLeft > 0) {
      this.quotaRejectionsLeft--;
      return QUOTA_REJECTION;
    }

    const path = decodeURIComponent(parsed.pathname);

    if (method === 'POST' && path === '/v4/spreadsheets') return this.createSpreadsheet(body);
    const batch = path.match(/^\/v4\/spreadsheets\/([^/]+):batchUpdate$/);
    if (method === 'POST' && batch) return this.batchUpdate(batch[1], body);

    const valuesBatchGet = path.match(/^\/v4\/spreadsheets\/([^/]+)\/values:batchGet$/);
    if (method === 'GET' && valuesBatchGet) return this.valuesBatchGet(valuesBatchGet[1], parsed);

    const valuesBatchUpdate = path.match(/^\/v4\/spreadsheets\/([^/]+)\/values:batchUpdate$/);
    if (method === 'POST' && valuesBatchUpdate) return this.valuesBatchUpdate(valuesBatchUpdate[1], body);

    const values = path.match(/^\/v4\/spreadsheets\/([^/]+)\/values\/(.+?)(:append)?$/);
    if (values) {
      const [, spreadsheetId, range, append] = values;
      if (method === 'GET') return this.valuesGet(spreadsheetId, range);
      if (method === 'PUT') return this.valuesUpdate(spreadsheetId, range, body);
      if (method === 'POST' && append) return this.valuesAppend(spreadsheetId, range, body, parsed);
    }
    const spreadsheet = path.match(/^\/v4\/spreadsheets\/([^/]+)$/);
    if (method === 'GET' && spreadsheet) return this.spreadsheetGet(spreadsheet[1]);

    return { status: 404, body: { error: { code: 404, message: `Unhandled ${method} ${path}`, status: 'NOT_FOUND' } } };
  }

  /**
   * Drive v3 permissions: create/list/delete against an in-memory map keyed by fileId.
   * The fake does not model drive.file's created-or-opened visibility rule — it answers
   * for any known spreadsheet id, which is what the tests exercise.
   */
  private driveHandle(method: string, parsed: URL, body: any): FakeResponse {
    const path = decodeURIComponent(parsed.pathname);

    const create = path.match(/^\/drive\/v3\/files\/([^/]+)\/permissions$/);
    if (method === 'POST' && create) {
      const fileId = create[1];
      if (!this.spreadsheets.has(fileId)) return this.notFound();
      const permission = {
        id: `perm-${this.nextPermissionId++}`,
        type: body?.type ?? 'user',
        role: body?.role ?? 'reader',
        ...(body?.emailAddress ? { emailAddress: body.emailAddress } : {}),
        ...(body?.domain ? { domain: body.domain } : {}),
      };
      const list = this.permissions.get(fileId) || [];
      list.push(permission);
      this.permissions.set(fileId, list);
      return { status: 200, body: permission };
    }

    if (method === 'GET' && create) {
      const fileId = create[1];
      if (!this.spreadsheets.has(fileId)) return this.notFound();
      return { status: 200, body: { permissions: this.permissions.get(fileId) || [] } };
    }

    const remove = path.match(/^\/drive\/v3\/files\/([^/]+)\/permissions\/([^/]+)$/);
    if (method === 'DELETE' && remove) {
      const [, fileId, permissionId] = remove;
      const list = this.permissions.get(fileId) || [];
      const remaining = list.filter((p) => p.id !== permissionId);
      if (remaining.length === list.length) return this.notFound();
      this.permissions.set(fileId, remaining);
      return { status: 200, body: {} };
    }

    return { status: 404, body: { error: { code: 404, message: `Unhandled Drive ${method} ${path}`, status: 'NOT_FOUND' } } };
  }

  /**
   * `spreadsheets.create`
   */
  private createSpreadsheet(body: any): FakeResponse {
    const spreadsheetId = `fake-spreadsheet-${this.nextSpreadsheetId++}`;
    const title = body?.properties?.title ?? 'Untitled spreadsheet';
    const spreadsheet = this.addSpreadsheet(spreadsheetId, title, [{ title: 'Sheet1' }]);
    return { status: 200, body: this.renderSpreadsheet(spreadsheet) };
  }

  /**
   * `spreadsheets.get`
   */
  private spreadsheetGet(spreadsheetId: string): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();
    return { status: 200, body: this.renderSpreadsheet(spreadsheet) };
  }

  /**
   * `spreadsheets.batchUpdate` for addSheet, deleteSheet, updateSheetProperties and appendDimension
   */
  private batchUpdate(spreadsheetId: string, body: any): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    // Google applies a batch atomically and in order: a later subrequest may reference a
    // sheet an earlier one created, and any invalid request rejects the whole batch without
    // leaving partial state behind. The fake therefore applies each request to a throwaway
    // copy and commits the copy only after every request has validated.
    const working: FakeSpreadsheet = {
      spreadsheetId: spreadsheet.spreadsheetId,
      title: spreadsheet.title,
      sheets: spreadsheet.sheets.map((sheet) => ({ ...sheet, cells: new Map(sheet.cells) })),
      developerMetadata: spreadsheet.developerMetadata?.map((meta) => ({ ...meta })),
    };
    const sheetIdInUse = (id: number): boolean => working.sheets.some((s) => s.sheetId === id);

    const replies: Record<string, unknown>[] = [];
    const requests = body?.requests ?? [];
    for (let index = 0; index < requests.length; index++) {
      const request = requests[index];
      if (request.addSheet) {
        const properties = request.addSheet.properties ?? {};
        const title = properties.title;
        if (working.sheets.some((s) => s.title === title)) {
          return this.badRequest(`Invalid requests[${index}].addSheet: A sheet with the name "${title}" already exists. Please enter another name.`);
        }
        // The caller may pin the id (0 is a valid Google sheet id); only invent one when it
        // is absent, and never hand out an id the spreadsheet already uses.
        let sheetId: number;
        if (properties.sheetId !== undefined && properties.sheetId !== null) {
          if (!Number.isInteger(properties.sheetId) || properties.sheetId < 0) {
            return this.badRequest(`Invalid requests[${index}].addSheet: The provided sheetId ${properties.sheetId} is invalid.`);
          }
          if (sheetIdInUse(properties.sheetId)) {
            return this.badRequest(`Invalid requests[${index}].addSheet: A sheet with the id ${properties.sheetId} already exists.`);
          }
          sheetId = properties.sheetId;
        } else {
          sheetId = this.nextSheetId++;
          while (sheetIdInUse(sheetId)) sheetId = this.nextSheetId++;
        }
        const grid = properties.gridProperties ?? {};
        const sheet: FakeWorksheet = {
          sheetId,
          title,
          index: working.sheets.length,
          rowCount: grid.rowCount ?? 1000,
          columnCount: grid.columnCount ?? 26,
          cells: new Map(),
        };
        if (grid.frozenRowCount !== undefined) sheet.frozenRowCount = grid.frozenRowCount;
        if (grid.frozenColumnCount !== undefined) sheet.frozenColumnCount = grid.frozenColumnCount;
        working.sheets.push(sheet);
        replies.push({ addSheet: { properties: this.renderSheetProperties(sheet) } });
        continue;
      }

      if (request.deleteSheet) {
        const sheetIndex = working.sheets.findIndex((s) => s.sheetId === request.deleteSheet.sheetId);
        if (sheetIndex < 0) return this.badRequest(`Invalid requests[${index}].deleteSheet: No sheet with id: ${request.deleteSheet.sheetId}`);
        working.sheets.splice(sheetIndex, 1);
        working.sheets.forEach((s, i) => (s.index = i));
        replies.push({});
        continue;
      }

      if (request.updateSheetProperties) {
        const properties = request.updateSheetProperties.properties ?? {};
        const sheet = working.sheets.find((s) => s.sheetId === properties.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].updateSheetProperties: No sheet with id: ${properties.sheetId}`);
        const fields: string[] = String(request.updateSheetProperties.fields ?? '').split(',');
        if (fields.includes('title') || fields.includes('*')) sheet.title = properties.title;
        if (properties.gridProperties?.rowCount) sheet.rowCount = properties.gridProperties.rowCount;
        if (properties.gridProperties?.columnCount) sheet.columnCount = properties.gridProperties.columnCount;
        if (properties.gridProperties?.frozenRowCount !== undefined) sheet.frozenRowCount = properties.gridProperties.frozenRowCount;
        if (properties.gridProperties?.frozenColumnCount !== undefined) sheet.frozenColumnCount = properties.gridProperties.frozenColumnCount;
        replies.push({});
        continue;
      }

      if (request.appendDimension) {
        const { sheetId, dimension, length } = request.appendDimension;
        const sheet = working.sheets.find((s) => s.sheetId === sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].appendDimension: No sheet with id: ${sheetId}`);
        if (!length || length < 1) return this.badRequest(`Invalid requests[${index}].appendDimension: length must be positive.`);
        if (dimension === 'ROWS') sheet.rowCount += length;
        else if (dimension === 'COLUMNS') sheet.columnCount += length;
        else return this.badRequest(`Invalid requests[${index}].appendDimension: unknown dimension ${dimension}`);
        replies.push({});
        continue;
      }

      if (request.insertDimension) {
        const { range } = request.insertDimension;
        const sheet = working.sheets.find((s) => s.sheetId === range?.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].insertDimension: No sheet with id: ${range?.sheetId}`);
        const startIndex = range.startIndex ?? 0;
        const endIndex = range.endIndex ?? startIndex + 1;
        const length = endIndex - startIndex;
        if (length < 1) return this.badRequest(`Invalid requests[${index}].insertDimension: endIndex must exceed startIndex.`);
        if (range.dimension === 'ROWS') {
          const shifted = new Map<string, string>();
          for (const [key, value] of sheet.cells) {
            const [r, c] = key.split(':').map(Number);
            shifted.set(r > startIndex ? `${r + length}:${c}` : key, value);
          }
          sheet.cells = shifted;
          sheet.rowCount += length;
        } else if (range.dimension === 'COLUMNS') {
          const shifted = new Map<string, string>();
          for (const [key, value] of sheet.cells) {
            const [r, c] = key.split(':').map(Number);
            shifted.set(c > startIndex ? `${r}:${c + length}` : key, value);
          }
          sheet.cells = shifted;
          sheet.columnCount += length;
        } else {
          return this.badRequest(`Invalid requests[${index}].insertDimension: unknown dimension ${range.dimension}`);
        }
        replies.push({});
        continue;
      }

      if (request.deleteDimension) {
        const { range } = request.deleteDimension;
        const sheet = working.sheets.find((s) => s.sheetId === range?.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].deleteDimension: No sheet with id: ${range?.sheetId}`);
        const startIndex = range.startIndex ?? 0;
        const endIndex = range.endIndex ?? startIndex + 1;
        const length = endIndex - startIndex;
        if (length < 1) return this.badRequest(`Invalid requests[${index}].deleteDimension: endIndex must exceed startIndex.`);
        if (range.dimension === 'ROWS') {
          const shifted = new Map<string, string>();
          for (const [key, value] of sheet.cells) {
            const [r, c] = key.split(':').map(Number);
            if (r > startIndex && r <= endIndex) continue; // inside the deleted band
            shifted.set(r > endIndex ? `${r - length}:${c}` : key, value);
          }
          sheet.cells = shifted;
          sheet.rowCount = Math.max(1, sheet.rowCount - length);
        } else if (range.dimension === 'COLUMNS') {
          const shifted = new Map<string, string>();
          for (const [key, value] of sheet.cells) {
            const [r, c] = key.split(':').map(Number);
            if (c > startIndex && c <= endIndex) continue;
            shifted.set(c > endIndex ? `${r}:${c - length}` : key, value);
          }
          sheet.cells = shifted;
          sheet.columnCount = Math.max(1, sheet.columnCount - length);
        } else {
          return this.badRequest(`Invalid requests[${index}].deleteDimension: unknown dimension ${range.dimension}`);
        }
        replies.push({});
        continue;
      }

      if (request.updateCells) {
        const { rows, start, range, fields } = request.updateCells;
        if (start && rows) {
          const sheet = working.sheets.find((s) => s.sheetId === start.sheetId);
          if (!sheet) return this.badRequest(`Invalid requests[${index}].updateCells: No sheet with id: ${start.sheetId}`);
          const startRowIdx = start.rowIndex ?? 0;
          const startColIdx = start.columnIndex ?? 0;
          rows.forEach((row: { values?: Array<{ userEnteredValue?: { stringValue?: string; numberValue?: number; boolValue?: boolean; formulaValue?: string } }> }, r: number) => {
            const cellValues = row?.values ?? [];
            cellValues.forEach((cellData, c: number) => {
              const rPos = startRowIdx + r + 1;
              const cPos = startColIdx + c + 1;
              const key = `${rPos}:${cPos}`;
              const uev = cellData?.userEnteredValue;
              if (!uev) {
                if (fields === 'userEnteredValue' || fields === '*') {
                  sheet.cells.delete(key);
                }
                return;
              }
              if (uev.formulaValue !== undefined) {
                const f = String(uev.formulaValue);
                sheet.cells.set(key, f.startsWith('=') ? f : `=${f}`);
              } else if (uev.numberValue !== undefined) {
                sheet.cells.set(key, String(uev.numberValue));
              } else if (uev.stringValue !== undefined) {
                sheet.cells.set(key, String(uev.stringValue));
              } else if (uev.boolValue !== undefined) {
                sheet.cells.set(key, String(uev.boolValue));
              }
            });
          });
          replies.push({});
          continue;
        }

        if (range && !rows) {
          const sheet = working.sheets.find((s) => s.sheetId === range.sheetId);
          if (!sheet) return this.badRequest(`Invalid requests[${index}].updateCells: No sheet with id: ${range.sheetId}`);
          const startRow = (range.startRowIndex ?? 0) + 1;
          const endRow = range.endRowIndex !== undefined ? range.endRowIndex : sheet.rowCount;
          const startCol = (range.startColumnIndex ?? 0) + 1;
          const endCol = range.endColumnIndex !== undefined ? range.endColumnIndex : sheet.columnCount;
          for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
              sheet.cells.delete(`${r}:${c}`);
            }
          }
          replies.push({});
          continue;
        }

        replies.push({});
        continue;
      }

      if (request.updateDimensionProperties) {
        const { range, properties, fields } = request.updateDimensionProperties;
        const sheet = working.sheets.find((s) => s.sheetId === range?.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].updateDimensionProperties: No sheet with id: ${range?.sheetId}`);
        const fieldList = String(fields ?? '').split(',');
        const startIndex = range.startIndex ?? 0;
        const endIndex = range.endIndex ?? startIndex + 1;
        const bucket = range.dimension === 'ROWS' ? 'rows' : 'columns';
        if (!sheet.dimensionProps) sheet.dimensionProps = {};
        if (!sheet.dimensionProps[bucket]) sheet.dimensionProps[bucket] = {};
        const props = sheet.dimensionProps[bucket] as Record<number, { hiddenByUser?: boolean; pixelSize?: number }>;
        for (let i = startIndex; i < endIndex; i++) {
          const key = i + 1; // dimensionProps is keyed 1-based
          if (!props[key]) props[key] = {};
          if (fieldList.includes('hiddenByUser') || fieldList.includes('*')) props[key].hiddenByUser = properties?.hiddenByUser;
          if (fieldList.includes('pixelSize') || fieldList.includes('*')) props[key].pixelSize = properties?.pixelSize;
        }
        replies.push({});
        continue;
      }

      if (request.autoResizeDimensions) {
        replies.push({});
        continue;
      }

      if (request.repeatCell) {
        replies.push({});
        continue;
      }

      if (request.mergeCells) {
        const { range } = request.mergeCells;
        const sheet = working.sheets.find((s) => s.sheetId === range?.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].mergeCells: No sheet with id: ${range?.sheetId}`);
        const merge = {
          startRowIndex: range.startRowIndex ?? 0,
          endRowIndex: range.endRowIndex ?? sheet.rowCount,
          startColumnIndex: range.startColumnIndex ?? 0,
          endColumnIndex: range.endColumnIndex ?? sheet.columnCount,
        };
        sheet.merges = (sheet.merges || []).filter(
          (m) =>
            m.endRowIndex <= merge.startRowIndex ||
            m.startRowIndex >= merge.endRowIndex ||
            m.endColumnIndex <= merge.startColumnIndex ||
            m.startColumnIndex >= merge.endColumnIndex
        );
        sheet.merges.push(merge);
        replies.push({});
        continue;
      }

      if (request.unmergeCells) {
        const { range } = request.unmergeCells;
        const sheet = working.sheets.find((s) => s.sheetId === range?.sheetId);
        if (!sheet) return this.badRequest(`Invalid requests[${index}].unmergeCells: No sheet with id: ${range?.sheetId}`);
        const bounds = {
          startRowIndex: range.startRowIndex ?? 0,
          endRowIndex: range.endRowIndex ?? sheet.rowCount,
          startColumnIndex: range.startColumnIndex ?? 0,
          endColumnIndex: range.endColumnIndex ?? sheet.columnCount,
        };
        sheet.merges = (sheet.merges || []).filter(
          (m) =>
            m.endRowIndex <= bounds.startRowIndex ||
            m.startRowIndex >= bounds.endRowIndex ||
            m.endColumnIndex <= bounds.startColumnIndex ||
            m.startColumnIndex >= bounds.endColumnIndex
        );
        replies.push({});
        continue;
      }

      if (request.updateBorders) {
        replies.push({});
        continue;
      }

      if (request.createDeveloperMetadata) {
        const meta = request.createDeveloperMetadata.developerMetadata || {};
        const metadataId = this.nextSheetId++;
        const fullMeta = { developerMetadataId: metadataId, ...meta };
        if (!working.developerMetadata) working.developerMetadata = [];
        working.developerMetadata.push(fullMeta);
        replies.push({
          createDeveloperMetadata: {
            developerMetadata: fullMeta,
          },
        });
        continue;
      }

      if (request.updateDeveloperMetadata) {
        const meta = request.updateDeveloperMetadata.developerMetadata || {};
        if (!working.developerMetadata) working.developerMetadata = [];
        const existing = working.developerMetadata.find((m: Record<string, unknown>) => {
          const mLocation = m.location as { sheetId?: number } | undefined;
          const metaLocation = meta.location as { sheetId?: number } | undefined;
          return (meta.metadataId && m.developerMetadataId === meta.metadataId) ||
            (meta.metadataKey && m.metadataKey === meta.metadataKey &&
              (!metaLocation?.sheetId || mLocation?.sheetId === metaLocation.sheetId));
        });
        if (existing) {
          Object.assign(existing, meta);
          replies.push({
            updateDeveloperMetadata: {
              developerMetadata: [existing],
            },
          });
        } else {
          const metadataId = this.nextSheetId++;
          const fullMeta = { developerMetadataId: metadataId, ...meta };
          working.developerMetadata.push(fullMeta);
          replies.push({
            updateDeveloperMetadata: {
              developerMetadata: [fullMeta],
            },
          });
        }
        continue;
      }

      if (request.addNamedRange) {
        replies.push({
          addNamedRange: {
            namedRange: {
              namedRangeId: `nr-${this.nextSheetId++}`,
              ...request.addNamedRange.namedRange,
            },
          },
        });
        continue;
      }

      if (request.updateNamedRange) {
        replies.push({
          updateNamedRange: {
            namedRange: request.updateNamedRange.namedRange,
          },
        });
        continue;
      }

      return this.badRequest(`Invalid requests[${index}]: unsupported request ${Object.keys(request).join(',')}`);
    }

    // Every request validated: publish the copy in one swap so readers never see
    // a half-applied batch, and failures above leave the live spreadsheet untouched.
    spreadsheet.sheets = working.sheets;
    spreadsheet.developerMetadata = working.developerMetadata;
    return { status: 200, body: { spreadsheetId, replies } };
  }

  /**
   * `spreadsheets.values.batchGet`
   */
  private valuesBatchGet(spreadsheetId: string, parsed: URL): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    let ranges = parsed.searchParams.getAll('ranges');
    if (ranges.length === 0) {
      const single = parsed.searchParams.get('ranges');
      if (single) ranges = [single];
    }

    const valueRenderOption = parsed.searchParams.get('valueRenderOption') || 'FORMATTED_VALUE';

    const valueRanges = ranges.map((rangeStr) => {
      let target: ResolvedRange;
      try {
        target = this.resolve(spreadsheet, rangeStr);
      } catch {
        return {
          range: rangeStr,
          majorDimension: 'ROWS',
          values: [],
        };
      }

      const { sheet } = target;
      const startRow = Math.min(target.startRow, sheet.rowCount);
      const startCol = Math.min(target.startCol, sheet.columnCount);
      const endRow = Math.min(target.endRow, sheet.rowCount);
      const endCol = Math.min(target.endCol, sheet.columnCount);

      const rows: (string | number | boolean)[][] = [];
      for (let r = startRow; r <= endRow; r++) {
        const row: (string | number | boolean)[] = [];
        for (let c = startCol; c <= endCol; c++) {
          const val = sheet.cells.get(`${r}:${c}`) ?? '';
          if (valueRenderOption === 'FORMULA') {
            row.push(val);
          } else if (valueRenderOption === 'UNFORMATTED_VALUE') {
            if (val === 'TRUE' || val === 'true') row.push(true);
            else if (val === 'FALSE' || val === 'false') row.push(false);
            else if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(val)) {
              row.push(Number(val));
            } else {
              row.push(val);
            }
          } else {
          }
        }
        while (row.length && row[row.length - 1] === '') row.pop();
        rows.push(row);
      }
      while (rows.length && rows[rows.length - 1].length === 0) rows.pop();

      return {
        range: formatRange(sheet.title, target.startRow, target.startCol, target.endRow, target.endCol),
        majorDimension: 'ROWS',
        values: rows,
      };
    });

    return {
      status: 200,
      body: {
        spreadsheetId,
        valueRanges,
      },
    };
  }

  /**
   * `spreadsheets.values.batchUpdate`
   */
  private valuesBatchUpdate(spreadsheetId: string, body: { valueInputOption?: string; data?: Array<{ range: string; values: unknown[][] }> }): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    const data = body?.data ?? [];
    let totalUpdatedRows = 0;
    let totalUpdatedColumns = 0;
    let totalUpdatedCells = 0;
    const responses: unknown[] = [];

    for (const item of data) {
      let target: ResolvedRange;
      try {
        target = this.resolve(spreadsheet, item.range);
      } catch (error) {
        return this.badRequest((error as Error).message);
      }

      const { sheet, startRow, startCol } = target;
      const values = item.values ?? [];

      const gridError = this.checkGrid(sheet, item.range, target, values);
      if (gridError) return gridError;

      let updatedRows = 0;
      let updatedColumns = 0;
      let updatedCells = 0;
      values.forEach((row, r) => {
        if (row.length) updatedRows++;
        updatedColumns = Math.max(updatedColumns, row.length);
        row.forEach((value, c) => {
          if (value === null || value === undefined) return;
          sheet.cells.set(`${startRow + r}:${startCol + c}`, String(value));
          updatedCells++;
        });
      });

      const endRow = startRow + Math.max(values.length, 1) - 1;
      const endCol = startCol + Math.max(updatedColumns, 1) - 1;

      totalUpdatedRows += updatedRows;
      totalUpdatedColumns = Math.max(totalUpdatedColumns, updatedColumns);
      totalUpdatedCells += updatedCells;

      responses.push({
        spreadsheetId,
        updatedRange: formatRange(sheet.title, startRow, startCol, endRow, endCol),
        updatedRows,
        updatedColumns,
        updatedCells,
      });
    }

    return {
      status: 200,
      body: {
        spreadsheetId,
        totalUpdatedRows,
        totalUpdatedColumns,
        totalUpdatedCells,
        totalUpdatedSheets: new Set(responses.map((r) => ((r as { updatedRange: string }).updatedRange.split('!')[0]))).size,
        responses,
      },
    };
  }
  /**
   * `spreadsheets.values.get`. A read whose range reaches past the grid is refused exactly the
   * way a write in the same position is refused. That is the strict reading, and the library
   * agrees with it: getData clamps maxRow and maxCol to gridProperties before every read, which
   * is only worth doing if an unclamped read fails.
   */
  private valuesGet(spreadsheetId: string, range: string): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    let target: ResolvedRange;
    try {
      target = this.resolve(spreadsheet, range);
    } catch (error) {
      return this.badRequest((error as Error).message);
    }

    const gridError = this.checkGrid(target.sheet, range, target, []);
    if (gridError) return gridError;

    const { sheet } = target;
    const startRow = Math.min(target.startRow, sheet.rowCount);
    const startCol = Math.min(target.startCol, sheet.columnCount);
    const endRow = Math.min(target.endRow, sheet.rowCount);
    const endCol = Math.min(target.endCol, sheet.columnCount);

    const rows: string[][] = [];
    for (let r = startRow; r <= endRow; r++) {
      const row: string[] = [];
      for (let c = startCol; c <= endCol; c++) row.push(sheet.cells.get(`${r}:${c}`) ?? '');
      while (row.length && row[row.length - 1] === '') row.pop();
      rows.push(row);
    }
    while (rows.length && rows[rows.length - 1].length === 0) rows.pop();

    const body: any = { range: formatRange(sheet.title, target.startRow, target.startCol, target.endRow, target.endCol), majorDimension: 'ROWS' };
    if (rows.length) body.values = rows;
    return { status: 200, body };
  }

  /**
   * `spreadsheets.values.update`. Writes are strict: the grid never grows on its own.
   */
  private valuesUpdate(spreadsheetId: string, range: string, body: any): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    let target: ResolvedRange;
    try {
      target = this.resolve(spreadsheet, range);
    } catch (error) {
      return this.badRequest((error as Error).message);
    }

    const { sheet, startRow, startCol } = target;
    const values: any[][] = body?.values ?? [];

    const gridError = this.checkGrid(sheet, range, target, values);
    if (gridError) return gridError;

    let updatedRows = 0;
    let updatedColumns = 0;
    let updatedCells = 0;
    values.forEach((row, r) => {
      if (row.length) updatedRows++;
      updatedColumns = Math.max(updatedColumns, row.length);
      row.forEach((value, c) => {
        // Google ignores null cells rather than clearing them
        if (value === null || value === undefined) return;
        sheet.cells.set(`${startRow + r}:${startCol + c}`, String(value));
        updatedCells++;
      });
    });

    const endRow = startRow + Math.max(values.length, 1) - 1;
    const endCol = startCol + Math.max(updatedColumns, 1) - 1;
    return {
      status: 200,
      body: {
        spreadsheetId,
        updatedRange: formatRange(sheet.title, startRow, startCol, endRow, endCol),
        updatedRows,
        updatedColumns,
        updatedCells,
      },
    };
  }

  /**
   * `spreadsheets.values.append` with `insertDataOption: OVERWRITE`. The grid grows server side.
   */
  private valuesAppend(spreadsheetId: string, range: string, body: any, parsed: URL): FakeResponse {
    const spreadsheet = this.spreadsheets.get(spreadsheetId);
    if (!spreadsheet) return this.notFound();

    let target: ResolvedRange;
    try {
      target = this.resolve(spreadsheet, range);
    } catch (error) {
      return this.badRequest((error as Error).message);
    }

    const insertDataOption = parsed.searchParams.get('insertDataOption') ?? 'OVERWRITE';

    const { sheet, startCol, endCol } = target;
    // the API looks for the last row of the table inside the searched columns
    let lastRow = target.startRow - 1;
    sheet.cells.forEach((value, key) => {
      if (value === '') return;
      const [r, c] = key.split(':').map(Number);
      if (c < startCol || c > endCol) return;
      if (r < target.startRow || r > target.endRow) return;
      if (r > lastRow) lastRow = r;
    });

    const values: any[][] = body?.values ?? [];
    const writeStart = lastRow + 1;
    let updatedColumns = 0;
    values.forEach((row) => (updatedColumns = Math.max(updatedColumns, row.length)));

    if (insertDataOption === 'INSERT_ROWS') {
      const shiftRows = values.length;
      if (shiftRows > 0) {
        const toShift: { r: number; c: number; val: string }[] = [];
        sheet.cells.forEach((value, key) => {
          const [r, c] = key.split(':').map(Number);
          if (r >= writeStart) {
            toShift.push({ r, c, val: value });
          }
        });
        toShift.forEach(({ r, c }) => {
          sheet.cells.delete(`${r}:${c}`);
        });
        toShift.forEach(({ r, c, val }) => {
          sheet.cells.set(`${r + shiftRows}:${c}`, val);
        });
        sheet.rowCount += shiftRows;
      }
    }

    // append grows the grid instead of failing
    const neededRows = writeStart + values.length - 1;
    const neededCols = startCol + updatedColumns - 1;
    if (neededRows > sheet.rowCount) sheet.rowCount = neededRows;
    if (neededCols > sheet.columnCount) sheet.columnCount = neededCols;

    let updatedCells = 0;
    values.forEach((row, r) => {
      row.forEach((value, c) => {
        if (value === null || value === undefined) return;
        sheet.cells.set(`${writeStart + r}:${startCol + c}`, String(value));
        updatedCells++;
      });
    });
    return {
      status: 200,
      body: {
        spreadsheetId,
        tableRange: formatRange(sheet.title, target.startRow, startCol, lastRow, endCol),
        updates: {
          spreadsheetId,
          updatedRange: formatRange(sheet.title, writeStart, startCol, writeStart + values.length - 1, startCol + updatedColumns - 1),
          updatedRows: values.length,
          updatedColumns,
          updatedCells,
        },
      },
    };
  }

  /**
   * Resolve a request range against a spreadsheet, defaulting to the whole grid
   */
  private resolve(spreadsheet: FakeSpreadsheet, range: string): ResolvedRange {
    const parsed = parseFakeRange(range);
    const sheet = parsed.title ? spreadsheet.sheets.find((s) => s.title === parsed.title) : spreadsheet.sheets[0];
    if (!sheet) throw new Error(`Unable to parse range: ${range}`);

    return {
      sheet,
      startRow: parsed.startRow ?? 1,
      startCol: parsed.startCol ?? 1,
      endRow: parsed.endRow ?? sheet.rowCount,
      endCol: parsed.endCol ?? sheet.columnCount,
      bounded: parsed.bounded,
    };
  }

  /**
   * Reproduce the ways the API refuses a write that does not fit the grid
   */
  private checkGrid(sheet: FakeWorksheet, range: string, target: ResolvedRange, values: any[][]): FakeResponse | undefined {
    const parsed = parseFakeRange(range);
    const limits = `Max rows: ${sheet.rowCount}, max columns: ${sheet.columnCount}`;

    if (target.startRow > sheet.rowCount || target.startCol > sheet.columnCount) {
      return this.badRequest(`Range (${sheet.title}!${colToLetter(target.startCol)}${target.startRow}) exceeds grid limits. ${limits}`);
    }
    if (parsed.bounded && parsed.endRow !== undefined && (parsed.endRow > sheet.rowCount || (parsed.endCol ?? 1) > sheet.columnCount)) {
      return this.badRequest(`Range (${sheet.title}!${colToLetter(parsed.endCol ?? 1)}${parsed.endRow}) exceeds grid limits. ${limits}`);
    }

    let longest = 0;
    values.forEach((row) => (longest = Math.max(longest, row.length)));
    const lastRow = target.startRow + values.length - 1;
    const lastCol = target.startCol + longest - 1;

    if (lastRow > sheet.rowCount) return this.badRequest(`Requested writing within range [${range}], but tried writing to row [${lastRow}]`);
    if (lastCol > sheet.columnCount) return this.badRequest(`Requested writing within range [${range}], but tried writing to column [${colToLetter(lastCol)}]`);
    if (target.bounded && lastRow > target.endRow) return this.badRequest(`Requested writing within range [${range}], but tried writing to row [${lastRow}]`);
    if (target.bounded && lastCol > target.endCol) {
      return this.badRequest(`Requested writing within range [${range}], but tried writing to column [${colToLetter(lastCol)}]`);
    }
    return undefined;
  }

  private renderSheetProperties(sheet: FakeWorksheet): any {
    const gridProperties: any = { rowCount: sheet.rowCount, columnCount: sheet.columnCount };
    if (sheet.frozenRowCount !== undefined) gridProperties.frozenRowCount = sheet.frozenRowCount;
    if (sheet.frozenColumnCount !== undefined) gridProperties.frozenColumnCount = sheet.frozenColumnCount;
    const properties: any = {
      sheetId: sheet.sheetId,
      title: sheet.title,
      index: sheet.index,
      sheetType: 'GRID',
      gridProperties,
    };
    if (sheet.merges && sheet.merges.length > 0) properties.merges = sheet.merges;
    return properties;
  }

  private renderSpreadsheet(spreadsheet: FakeSpreadsheet): any {
    return {
      spreadsheetId: spreadsheet.spreadsheetId,
      properties: { title: spreadsheet.title, locale: 'en_US', timeZone: 'Etc/GMT' },
      sheets: spreadsheet.sheets.map((sheet) => ({ properties: this.renderSheetProperties(sheet) })),
      developerMetadata: spreadsheet.developerMetadata || [],
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheet.spreadsheetId}/edit`,
    };
  }

  private notFound(): FakeResponse {
    return { status: 404, body: { error: { code: 404, message: 'Requested entity was not found.', status: 'NOT_FOUND' } } };
  }

  private badRequest(message: string): FakeResponse {
    return { status: 400, body: { error: { code: 400, message, status: 'INVALID_ARGUMENT' } } };
  }
}
