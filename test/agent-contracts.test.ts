import { Errors } from '@oclif/core';
import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';
import { join } from 'path';

import { GSheetError, GSheetErrorCode, classifyError, sanitizeMessage, toErrorEnvelope } from '../src/lib/cli-errors';
import type { GoogleSheetCli } from '../src/lib/google-sheet';
import { discoverSchema, findUniqueViolations, parseValidationSchema } from '../src/lib/data-schema';
import * as factory from '../src/lib/factory';
import GoogleSheet from '../src/lib/google-sheet';
import { refreshTokens } from '../src/lib/oauth';
import type { sheets_v4 } from '@googleapis/sheets';
import type { OAuth2Client } from 'google-auth-library';

/**
 * The shared CLI error contract plus the read-only data:schema / data:validate slice.
 *
 * Command cases run through `runCommand` with the factory swapped for a recording stub that
 * implements only what these commands may call - any attempt at a write reaches for a method
 * that does not exist and fails loudly, which is the zero-write proof. Error classification
 * cases inject gaxios-shaped failures (`{ response: { status } }`) because the shape, not the
 * transport, is what the envelope maps; pure cases cover the classifier directly.
 */

const ROOT = join(__dirname, '..');

const SPREADSHEET_ID = 'offline-spreadsheet-id';
const WORKSHEET_TITLE = 'offline-worksheet';
const CLIENT_EMAIL = 'offline@example.iam.gserviceaccount.com';

const PRIVATE_KEY = generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { format: 'pem', type: 'pkcs8' },
  publicKeyEncoding: { format: 'pem', type: 'spki' },
}).privateKey;

const MANAGED_ENV = ['GSHEET_CLIENT_EMAIL', 'GSHEET_PRIVATE_KEY', 'GSHEET_CREDENTIALS_FILE', 'SPREADSHEET_ID', 'WORKSHEET_TITLE'];

const AUTH_FLAGS = [`--spreadsheetId=${SPREADSHEET_ID}`, `--worksheetTitle=${WORKSHEET_TITLE}`, `--clientEmail=${CLIENT_EMAIL}`];

// ============================================================================
// Fixtures
// ============================================================================

const textCell = (value: string): sheets_v4.Schema$CellData => ({
  userEnteredValue: { stringValue: value },
  effectiveValue: { stringValue: value },
  formattedValue: value,
});

const numberCell = (value: number): sheets_v4.Schema$CellData => ({
  userEnteredValue: { numberValue: value },
  effectiveValue: { numberValue: value },
  formattedValue: String(value),
});

const formulaCell = (formula: string, computed: number): sheets_v4.Schema$CellData => ({
  userEnteredValue: { formulaValue: formula },
  effectiveValue: { numberValue: computed },
  formattedValue: String(computed),
});

const headerCell = (value: string): sheets_v4.Schema$CellData => ({
  userEnteredValue: { stringValue: value },
  formattedValue: value,
});

const rowOf = (...cells: sheets_v4.Schema$CellData[]): sheets_v4.Schema$RowData => ({ values: cells });

/**
 * A sample starting at sheet row 5, column C (0-based startRow 4 / startColumn 2):
 *
 *        C        D         E
 *   5    Name     Amount    Code     (header row)
 *   6    Ada      10        X1
 *   7    Grace    text      X2       ("text" violates the decimal Amount)
 *   8    (empty)  12        X3       (missing required Name)
 *   9    Linus    14        X1       (duplicate unique Code)
 */
const validationFixture = (): GoogleSheetCli.WorksheetMetadata => ({
  spreadsheetId: SPREADSHEET_ID,
  worksheetTitle: WORKSHEET_TITLE,
  properties: { sheetId: 42, title: WORKSHEET_TITLE, gridProperties: { rowCount: 1000, columnCount: 26 } },
  gridData: [
    {
      startRow: 4,
      startColumn: 2,
      rowData: [
        rowOf(headerCell('Name'), headerCell('Amount'), headerCell('Code')),
        rowOf(textCell('Ada'), numberCell(10), textCell('X1')),
        rowOf(textCell('Grace'), textCell('text'), textCell('X2')),
        rowOf({ formattedValue: '' }, numberCell(12), textCell('X3')),
        rowOf(textCell('Linus'), numberCell(14), textCell('X1')),
      ],
    },
  ],
  namedRanges: [
    { name: 'Mine', range: { sheetId: 42, startRowIndex: 4, endRowIndex: 9, startColumnIndex: 2, endColumnIndex: 5 } },
    { name: 'Elsewhere', range: { sheetId: 99, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 } },
  ],
});

/** A sample with mixed kinds, formulas, per-cell validation and duplicate headers. */
const discoveryFixture = (): GoogleSheetCli.WorksheetMetadata => ({
  spreadsheetId: SPREADSHEET_ID,
  worksheetTitle: WORKSHEET_TITLE,
  properties: { sheetId: 7, title: WORKSHEET_TITLE, gridProperties: { rowCount: 500, columnCount: 20 } },
  gridData: [
    {
      startRow: 4,
      startColumn: 2,
      rowData: [
        rowOf(headerCell('amount'), headerCell(''), headerCell('amount')),
        rowOf(numberCell(10), formulaCell('=B6*2', 20), textCell('ok')),
        rowOf(textCell('ten'), numberCell(1), textCell('ok')),
        rowOf(numberCell(12), numberCell(2), textCell('ok')),
      ],
    },
  ],
  namedRanges: [{ name: 'Block', range: { sheetId: 7, startRowIndex: 0, endRowIndex: 8, startColumnIndex: 0, endColumnIndex: 3 } }],
});

const httpError = (status: number, message = `Request failed with status ${status}`, headers?: Record<string, string>): Error => {
  const err = new Error(message);
  (err as unknown as Record<string, unknown>).response = { status, ...(headers ? { headers } : {}) };
  return err;
};

// ============================================================================
// Stub
// ============================================================================

class SchemaStubGoogleSheet {
  calls: { method: string; args: unknown[] }[] = [];
  metadataResponse: GoogleSheetCli.WorksheetMetadata = validationFixture();
  authorizeError: unknown;
  authorizeOAuthError: unknown;
  metadataError: unknown;

  async authorize(credentials: unknown): Promise<void> {
    this.calls.push({ method: 'authorize', args: [credentials] });
    if (this.authorizeError) throw this.authorizeError;
  }

  async authorizeOAuth(): Promise<void> {
    this.calls.push({ method: 'authorizeOAuth', args: [] });
    throw this.authorizeOAuthError ?? new Error('authorizeOAuth is not expected in these cases');
  }

  async getWorksheetMetadata(options: unknown, spreadsheetId?: string): Promise<GoogleSheetCli.WorksheetMetadata> {
    this.calls.push({ method: 'getWorksheetMetadata', args: [options, spreadsheetId] });
    if (this.metadataError) throw this.metadataError;
    return this.metadataResponse;
  }
}

const mutableFactory = factory as { createGoogleSheet: () => GoogleSheet };
const realCreateGoogleSheet = factory.createGoogleSheet;

const parseEnvelope = (stderr: string): { error: { code: string; message: string; retryable: boolean; retryAfterMs?: number; issues?: unknown[] } } =>
  JSON.parse(stderr);

describe('cli error contract and read-only data commands', () => {
  let stub: SchemaStubGoogleSheet;
  let savedEnv: Record<string, string | undefined>;

  beforeEach(() => {
    savedEnv = Object.fromEntries(MANAGED_ENV.map((name) => [name, process.env[name]]));
    for (const name of MANAGED_ENV) delete process.env[name];
    process.env.GSHEET_PRIVATE_KEY = PRIVATE_KEY;
    process.env.NODE_ENV = 'test';

    stub = new SchemaStubGoogleSheet();
    mutableFactory.createGoogleSheet = () => stub as unknown as GoogleSheet;
  });

  afterEach(() => {
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    for (const [name, value] of Object.entries(savedEnv)) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
  });

  describe('json failure envelope', () => {
    it('reports a parser failure as USAGE when --json is in argv before any flag parsing', async () => {
      const { error, stdout, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('USAGE');
      expect(envelope.error.retryable).to.equal(false);
      expect(envelope.error.message).to.be.a('string').and.to.have.length.above(0);
      expect(envelope.error).to.not.have.property('stack');
      expect(stdout).to.equal('');
      expect(error).to.exist;
    });

    it('reports a local credential failure as AUTH_REQUIRED without leaking the key', async () => {
      const { error, stderr } = await runCommand([
        'data:schema',
        `--spreadsheetId=${SPREADSHEET_ID}`,
        `--worksheetTitle=${WORKSHEET_TITLE}`,
        `--clientEmail=${CLIENT_EMAIL}`,
        '--privateKey=not-a-real-key',
        '--json',
      ]);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('AUTH_REQUIRED');
      expect(stderr).to.not.contain('PRIVATE KEY');
      expect(error).to.exist;
      // the key is rejected while credentials are normalized, before any client call
      expect(stub.calls.map(({ method }) => method)).to.eql([]);
    });

    it('maps a 403 to FORBIDDEN and carries nothing but the envelope fields', async () => {
      stub.metadataError = httpError(403);
      const { error, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('FORBIDDEN');
      expect(envelope.error.retryable).to.equal(false);
      expect(Object.keys(envelope.error).sort()).to.eql(['code', 'message', 'retryable']);
      expect(error).to.exist;
    });

    it('maps a 429 to retryable RATE_LIMITED with the Retry-After hint in milliseconds', async () => {
      stub.metadataError = httpError(429, 'Quota exceeded', { 'retry-after': '7' });
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('RATE_LIMITED');
      expect(envelope.error.retryable).to.equal(true);
      expect(envelope.error.retryAfterMs).to.equal(7000);
    });

    it('maps a transport failure to retryable NETWORK', async () => {
      const networkError = new Error('getaddrinfo ENOTFOUND sheets.googleapis.com');
      (networkError as unknown as Record<string, unknown>).code = 'ENOTFOUND';
      stub.metadataError = networkError;
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('NETWORK');
      expect(envelope.error.retryable).to.equal(true);
    });

    it('keeps the human error path when no json trigger is present', async () => {
      stub.metadataError = httpError(403);
      const { error, stderr } = await runCommand(['data:schema', ...AUTH_FLAGS]);
      expect(() => JSON.parse(stderr)).to.throw();
      expect(error).to.exist;
    });

    it('treats the rawOutput short flag as a json trigger even when parsing fails', async () => {
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '-r']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('USAGE');
    });

    it('treats the json short flag as a json trigger even when parsing fails', async () => {
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '-j']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('USAGE');
    });

    it('treats a short cluster combining both json flags as a trigger', async () => {
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '-rj']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('USAGE');
    });

    it('does not trigger json mode when a short cluster char is an option value', async () => {
      // `-tr` parses as --worksheetTitle r (an option char owns the rest of its cluster); the
      // later unknown flag fails the invocation, but the "r" was data, so the human path stays
      const { stderr } = await runCommand(['data:schema', `--spreadsheetId=${SPREADSHEET_ID}`, '-tr', '--nope']);
      expect(() => JSON.parse(stderr)).to.throw();
    });

    it('does not trigger json mode for a short flag occupying an option value slot', async () => {
      // --worksheetTitle expects a value, so a bare `-j` after it is the failed value slot,
      // not a request for JSON failure mode
      const { stderr } = await runCommand(['data:schema', `--spreadsheetId=${SPREADSHEET_ID}`, '--worksheetTitle', '-j', '--nope']);
      expect(() => JSON.parse(stderr)).to.throw();
    });

    it('does not trigger json mode for an inline option value containing the flag', async () => {
      const { stderr } = await runCommand(['data:schema', `--spreadsheetId=${SPREADSHEET_ID}`, '--worksheetTitle=-j', '--nope']);
      expect(() => JSON.parse(stderr)).to.throw();
    });

    it('does not trigger json mode for a short flag after the passthrough marker', async () => {
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '--', '-j']);
      expect(() => JSON.parse(stderr)).to.throw();
    });

    it('redacts credential material from the serialized message end to end', async () => {
      stub.metadataError = new Error(
        'request rejected: {"access_token":"ya29.topsecrettoken","refresh_token":"1//refreshsecret"} key: -----BEGIN PRIVATE KEY-----\\nMIIEvQIBADANBgkq\\n-----END PRIVATE KEY-----'
      );
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      expect(stderr).to.not.contain('MIIEvQ');
      expect(stderr).to.not.contain('topsecrettoken');
      expect(stderr).to.not.contain('refreshsecret');
      expect(stderr).to.contain('[REDACTED_KEY]');
    });

    it('carries nothing but the envelope fields for a raw provider error object', async () => {
      stub.metadataError = Object.assign(new Error('upstream rejected the request'), {
        config: { data: { refresh_token: '1//provider-refresh-secret' }, url: 'https://sheets.googleapis.com/v4/spreadsheets' },
        response: { status: 400, data: { error: { message: 'bad' } } },
      });
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('REQUEST_INVALID');
      expect(Object.keys(envelope.error).sort()).to.eql(['code', 'message', 'retryable']);
      expect(stderr).to.not.contain('provider-refresh-secret');
      expect(stderr).to.not.contain('sheets.googleapis.com');
    });

    it('classifies local OAuth failures as AUTH_REQUIRED and upstream ones by their own code', async () => {
      const OAUTH_FLAGS = [`--spreadsheetId=${SPREADSHEET_ID}`, `--worksheetTitle=${WORKSHEET_TITLE}`, '--useOauth'];

      stub.authorizeOAuthError = new Error('No OAuth tokens found. Please run "google-sheet auth:login" first to authenticate with your Google account.');
      const local = parseEnvelope((await runCommand(['data:schema', ...OAUTH_FLAGS, '--json'])).stderr);
      expect(local.error.code).to.equal('AUTH_REQUIRED');
      expect(local.error.retryable).to.equal(false);
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorizeOAuth']);

      stub.authorizeOAuthError = httpError(403, 'The caller does not have permission');
      expect(parseEnvelope((await runCommand(['data:schema', ...OAUTH_FLAGS, '--json'])).stderr).error.code).to.equal('FORBIDDEN');

      stub.authorizeOAuthError = httpError(429, 'Quota exceeded', { 'retry-after': '5' });
      const limited = parseEnvelope((await runCommand(['data:schema', ...OAUTH_FLAGS, '--json'])).stderr);
      expect(limited.error.code).to.equal('RATE_LIMITED');
      expect(limited.error.retryable).to.equal(true);
      expect(limited.error.retryAfterMs).to.equal(5000);

      stub.authorizeOAuthError = Object.assign(new Error('getaddrinfo ENOTFOUND oauth2.googleapis.com'), { code: 'ENOTFOUND' });
      const dead = parseEnvelope((await runCommand(['data:schema', ...OAUTH_FLAGS, '--json'])).stderr);
      expect(dead.error.code).to.equal('NETWORK');
      expect(dead.error.retryable).to.equal(true);

      // oauth.ts flattens refresh failures into a plain message; network evidence inside it
      // must still win over the auth classification
      stub.authorizeOAuthError = new Error('Failed to refresh access token: fetch failed');
      expect(parseEnvelope((await runCommand(['data:schema', ...OAUTH_FLAGS, '--json'])).stderr).error.code).to.equal('NETWORK');
    });

    it('ignores json triggers after the -- passthrough marker', async () => {
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--nope', '--', '--json']);
      expect(() => JSON.parse(stderr)).to.throw();
    });

    it('envelopes a thrown non-Error as INTERNAL without stringifying it', async () => {
      stub.metadataError = 'plain string failure';
      const { stderr } = await runCommand(['data:schema', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('INTERNAL');
      expect(envelope.error.message).to.equal('Unexpected error');
      expect(stderr).to.not.contain('plain string failure');
    });
  });

  describe('oauth token refresh wrapper', () => {
    const catchRefresh = async (client: OAuth2Client): Promise<unknown> => {
      try {
        await refreshTokens(client);
      } catch (err) {
        return err;
      }
      throw new Error('refreshTokens resolved; expected a rejection');
    };

    it('keeps transport status, cause and message on the wrapped refresh error', async () => {
      const cause = Object.assign(httpError(403, 'permission denied'), {
        config: { data: { refresh_token: '1//cause-secret-value' } },
      });
      const failingClient = {
        refreshAccessToken: async () => {
          throw cause;
        },
      } as unknown as OAuth2Client; // drives the actual refreshTokens wrapper; never touches the network
      const wrapped = (await catchRefresh(failingClient)) as Error & { status?: number; cause?: unknown };
      expect(wrapped.message).to.equal('Failed to refresh access token: permission denied');
      expect(wrapped.status).to.equal(403);
      expect(wrapped.cause).to.equal(cause);
      // non-enumerable cause: spreads and JSON serialization never pick it up
      expect(Object.keys(wrapped)).to.not.include('cause');
      expect(JSON.stringify(wrapped)).to.not.contain('cause-secret-value');
      expect(classifyError(wrapped).code).to.equal('FORBIDDEN');
    });

    it('maps a 429 refresh failure to retryable RATE_LIMITED and a network code to NETWORK', async () => {
      const failingClient = {
        refreshAccessToken: async () => {
          throw httpError(429);
        },
      } as unknown as OAuth2Client; // drives the actual refreshTokens wrapper; never touches the network
      const limited = (await catchRefresh(failingClient)) as { status?: number };
      expect(classifyError(limited)).to.eql({ code: 'RATE_LIMITED', retryable: true });

      const dead = Object.assign(new Error('getaddrinfo ENOTFOUND oauth2.googleapis.com'), { code: 'ENOTFOUND' });
      const net = (await catchRefresh({
        refreshAccessToken: async () => {
          throw dead;
        },
      } as unknown as OAuth2Client)) as { code?: string };
      expect(net.code).to.equal('ENOTFOUND');
      expect(classifyError(net).code).to.equal('NETWORK');
    });

    it('wraps an evidence-free refresh failure with the message only', async () => {
      const failingClient = {
        refreshAccessToken: async () => {
          throw new Error('refresh endpoint shut down');
        },
      } as unknown as OAuth2Client; // drives the actual refreshTokens wrapper; never touches the network
      const wrapped = (await catchRefresh(failingClient)) as Record<string, unknown>;
      expect(wrapped.message).to.equal('Failed to refresh access token: refresh endpoint shut down');
      expect(wrapped).to.not.have.property('status');
      expect(wrapped).to.not.have.property('code');
      // no transport evidence at the pure layer; resolveGoogleSheetAuth classifies this as AUTH_REQUIRED
      expect(classifyError(wrapped).code).to.equal('INTERNAL');
    });
  });

  describe('error classification (pure)', () => {
    it('keeps a GSheetError code and its issue details in the envelope', () => {
      const err = new GSheetError(GSheetErrorCode.DATA_INVALID, 'Data validation failed with 1 issue(s)', {
        details: { issues: [{ row: 7, a1: 'D7', field: 'Amount', code: 'TYPE_MISMATCH', message: 'not a number', value: 'text' }] },
      });
      const envelope = toErrorEnvelope(err);
      expect(envelope.error.code).to.equal('DATA_INVALID');
      expect(envelope.error.issues).to.have.lengthOf(1);
      expect((envelope.error.issues as { a1?: string }[])[0].a1).to.equal('D7');
    });

    it('never serializes the cause of a GSheetError', () => {
      const cause = new Error('inner with private_key: REDACTED-SECRET');
      const err = new GSheetError(GSheetErrorCode.AUTH_REQUIRED, 'outer failure', { cause });
      const envelope = toErrorEnvelope(err);
      expect(JSON.stringify(envelope)).to.not.contain('inner with private_key');
    });

    it('classifies numeric-string status codes and client errors', () => {
      expect(classifyError({ code: '403' }).code).to.equal('FORBIDDEN');
      expect(classifyError({ response: { status: 400 } }).code).to.equal('REQUEST_INVALID');
      expect(classifyError({ response: { status: 500 } })).to.eql({ code: 'UPSTREAM', retryable: true });
      expect(classifyError(new Errors.CLIError('bad flag')).code).to.equal('USAGE');
      expect(classifyError(new Error('something else')).code).to.equal('INTERNAL');
    });

    it('classifies dependency-free errors carrying a structural validation code', () => {
      // the shape the local input/validation layers throw (own enumerable `code`, no
      // dependency on this module): read structurally, message kept, nothing else copied
      const err = Object.assign(new Error('Input data must be a 2-dimensional matrix'), { code: 'VALIDATION' });
      expect(classifyError(err)).to.eql({ code: 'VALIDATION', retryable: false });
      const envelope = toErrorEnvelope(err);
      expect(envelope.error.code).to.equal('VALIDATION');
      expect(envelope.error.message).to.equal('Input data must be a 2-dimensional matrix');
    });

    it('classifies the other structural local codes without an INTERNAL fallback', () => {
      expect(classifyError({ code: 'SCHEMA_INVALID' }).code).to.equal('SCHEMA_INVALID');
      expect(classifyError({ code: 'DATA_INVALID' }).code).to.equal('DATA_INVALID');
      expect(classifyError({ code: 'USAGE' }).code).to.equal('USAGE');
      expect(classifyError({ code: 'AUTH_REQUIRED' }).code).to.equal('AUTH_REQUIRED');
    });

    it('puts a structural code ahead of transport evidence and keeps unknown errors INTERNAL', () => {
      expect(classifyError(Object.assign(new Error('nope'), { code: 'VALIDATION', response: { status: 500 } })).code).to.equal('VALIDATION');
      expect(classifyError(new Error('something else')).code).to.equal('INTERNAL');
    });

    it('redacts before truncating so no partial token survives the message cap', () => {
      const token = `ya29.${'a'.repeat(40)}`;
      const sanitized = sanitizeMessage(`${'x'.repeat(390)} ${token}`);
      expect(sanitized).to.have.lengthOf(403);
      expect(sanitized).to.not.contain('aaaa');
      expect(sanitized).to.contain('ya29.[RED');
    });

    it('redacts token material and collapses messages to one bounded line', () => {
      const sanitized = sanitizeMessage('line one\nya29.abcdefghijklmnop line two\n  Bearer abcdefghijklmno');
      expect(sanitized).to.not.contain('\n');
      expect(sanitized).to.contain('ya29.[REDACTED]');
      expect(sanitized).to.contain('Bearer [REDACTED]');
      expect(sanitizeMessage('x'.repeat(600))).to.have.lengthOf(403);
    });
  });

  describe('data:schema discovery', () => {
    it('reports absolute coordinates, duplicates, mixed kinds, formulas and named ranges', () => {
      const discovery = discoverSchema(discoveryFixture());
      expect(discovery.sheetId).to.equal(7);
      expect(discovery.sample).to.eql({ range: `${WORKSHEET_TITLE}!C5:E8`, startRow: 5, startColumn: 3, rows: 4, columns: 3 });

      const [first, second, third] = discovery.columns;
      expect(first.headerA1).to.equal('C5');
      expect(first.letter).to.equal('C');
      expect(second.letter).to.equal('D');
      expect(second.headerA1).to.equal('D5');
      expect(third.letter).to.equal('E');

      expect(first.duplicate).to.equal(true);
      expect(first.emptyHeader).to.equal(false);
      expect(second.emptyHeader).to.equal(true);
      expect(third.duplicate).to.equal(true);

      expect(first.inferredTypes).to.eql({ string: 1, number: 2, boolean: 0, error: 0 });
      expect(first.mixed).to.equal(true);
      expect(second.mixed).to.equal(false);
      expect(second.formulaCells).to.equal(1);

      expect(third.validations).to.eql([]);
      expect(discovery.namedRanges).to.eql([{ name: 'Block', sheetId: 7, range: 'A1:C8' }]);

      expect(discovery.warnings.some((w) => w.includes('Duplicate header "amount"'))).to.equal(true);
      expect(discovery.warnings.some((w) => w.includes('Empty header at D5'))).to.equal(true);
      expect(discovery.warnings.some((w) => w.includes('mixed types'))).to.equal(true);
    });

    it('defaults to A1-based coordinates when the API reports no offsets', () => {
      const meta = discoveryFixture();
      delete (meta.gridData[0] as { startRow?: number }).startRow;
      delete (meta.gridData[0] as { startColumn?: number }).startColumn;
      const discovery = discoverSchema(meta);
      expect(discovery.columns[0].headerA1).to.equal('A1');
      expect(discovery.columns[2].letter).to.equal('C');
    });

    it('returns schema JSON with provenance over --rawOutput and never writes', async () => {
      stub.metadataResponse = discoveryFixture();
      const { error, result } = await runCommand(['data:schema', ...AUTH_FLAGS, '--rawOutput']);
      if (error) throw error;
      expect((result as { columns: { headerA1: string }[] }).columns[0].headerA1).to.equal('C5');
      expect((result as { warnings: string[] }).warnings.length).to.be.above(0);
      expect((result as { namedRanges: unknown[] }).namedRanges).to.have.lengthOf(1);
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'getWorksheetMetadata']);
    });
  });

  describe('validation schema parsing', () => {
    it('accepts a well-formed schema and extracts unique fields', () => {
      const { schema, uniqueFields } = parseValidationSchema({
        fields: [
          { name: 'Name', type: 'string', required: true },
          { name: 'Amount', type: 'decimal', min: 0 },
          { name: 'Code', type: 'string', unique: true },
        ],
        allowExtraColumns: false,
      });
      expect(schema.fields).to.have.lengthOf(3);
      expect([...uniqueFields]).to.eql(['Code']);
    });

    it('collects every problem of a malformed schema into one SCHEMA_INVALID error', () => {
      let caught: unknown;
      try {
        parseValidationSchema({
          fields: [
            { name: 'A', type: 'nope', requird: true },
            { name: 'A', type: 'string', enum: [] },
            { name: 'B', type: 'string', decimalPlaces: -1 },
          ],
          bogus: true,
        });
      } catch (err) {
        caught = err;
      }
      expect(caught).to.be.instanceOf(GSheetError);
      const err = caught as GSheetError;
      expect(err.code).to.equal('SCHEMA_INVALID');
      const messages = (err.details?.issues ?? []).map((issue) => issue.message).join(' | ');
      expect(messages).to.contain('"requird"');
      expect(messages).to.contain('type');
      expect(messages).to.contain('Duplicate field name');
      expect(messages).to.contain('enum');
      expect(messages).to.contain('decimalPlaces');
      expect(messages).to.contain('bogus');
    });
  });

  describe('data:validate', () => {
    const VALIDATION_SCHEMA = JSON.stringify({
      fields: [
        { name: 'Name', type: 'string', required: true },
        { name: 'Amount', type: 'decimal' },
        { name: 'Code', type: 'string', unique: true },
      ],
    });

    it('reports violations at exact sheet coordinates and exits nonzero without writing', async () => {
      const { error, stderr } = await runCommand(['data:validate', ...AUTH_FLAGS, `--schema=${VALIDATION_SCHEMA}`, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('DATA_INVALID');
      expect(envelope.error.message).to.contain('issue');

      const issues = (envelope.error.issues ?? []) as { a1?: string; code?: string }[];
      const byA1 = (a1: string): { a1?: string; code?: string } | undefined => issues.find((issue) => issue.a1 === a1);
      expect(byA1('D7')).to.exist; // "text" in the decimal Amount column
      expect(byA1('C8')).to.exist; // missing required Name
      expect(byA1('E9')?.code).to.equal('DUPLICATE_VALUE'); // second X1
      expect(error).to.exist;
      // one read, zero writes: the stub has no write method at all
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'getWorksheetMetadata']);
    });

    it('passes valid data with a zero-issue receipt and no writes', async () => {
      stub.metadataResponse = validationFixture();
      stub.metadataResponse.gridData[0].rowData = stub.metadataResponse.gridData[0].rowData!.slice(0, 2);
      const { error, result } = await runCommand(['data:validate', ...AUTH_FLAGS, `--schema=${VALIDATION_SCHEMA}`, '--rawOutput']);
      expect(error).to.not.exist;
      expect((result as { valid: boolean; issues: unknown[]; rowsChecked: number })).to.eql({
        valid: true,
        issues: [],
        rowsChecked: 1,
        operation: 'data:validate',
        spreadsheetId: SPREADSHEET_ID,
        worksheetTitle: WORKSHEET_TITLE,
        range: `${WORKSHEET_TITLE}!C5:E6`,
        columnsChecked: 3,
        warnings: [],
      });
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize', 'getWorksheetMetadata']);
    });

    it('refuses a sheet whose header row does not carry the schema fields', async () => {
      const meta = validationFixture();
      meta.gridData[0].rowData![0].values = [headerCell('Wrong'), headerCell('Header'), headerCell('Row')];
      stub.metadataResponse = meta;
      const { stderr } = await runCommand(['data:validate', ...AUTH_FLAGS, `--schema=${VALIDATION_SCHEMA}`, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('DATA_INVALID');
      const issues = envelope.error.issues as { code?: string }[];
      expect(issues.filter(({ code }) => code === 'MISSING_HEADER')).to.have.lengthOf(3);
    });

    it('fails a malformed schema fast, before any Sheets call', async () => {
      const { stderr } = await runCommand([
        'data:validate',
        ...AUTH_FLAGS,
        '--schema={"fields":[{"name":"X","type":"nope"}]}',
        '--json',
      ]);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('SCHEMA_INVALID');
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize']);
    });

    it('classifies a missing schema as VALIDATION', async () => {
      const { stderr } = await runCommand(['data:validate', ...AUTH_FLAGS, '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('VALIDATION');
      expect(envelope.error.message).to.contain('No schema provided');
    });

    it('classifies an unreadable schema file as VALIDATION, not INTERNAL', async () => {
      const { stderr } = await runCommand(['data:validate', ...AUTH_FLAGS, '--schemaFile=no-such-schema-file.json', '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('VALIDATION');
      // the schema is read locally, so nothing beyond authorization reaches the client stub
      expect(stub.calls.map(({ method }) => method)).to.eql(['authorize']);
    });

    it('classifies malformed inline schema JSON as VALIDATION', async () => {
      const { stderr } = await runCommand(['data:validate', ...AUTH_FLAGS, '--schema={oops', '--json']);
      const envelope = parseEnvelope(stderr);
      expect(envelope.error.code).to.equal('VALIDATION');
    });

    it('finds uniqueness violations through the pure helper', () => {
      const { schema, uniqueFields } = parseValidationSchema({
        fields: [
          { name: 'Name', type: 'string' },
          { name: 'Amount', type: 'decimal' },
          { name: 'Code', type: 'string', unique: true },
        ],
      });
      const discovery = discoverSchema(validationFixture());
      const normalized = [
        { Name: 'Ada', Amount: 10, Code: 'X1' },
        { Name: 'Grace', Amount: null, Code: 'X2' },
        { Name: null, Amount: 12, Code: 'X3' },
        { Name: 'Linus', Amount: 14, Code: 'X1' },
      ];
      const issues = findUniqueViolations(normalized, uniqueFields, discovery);
      expect(issues).to.have.lengthOf(1);
      expect(issues[0].a1).to.equal('E9');
      expect(issues[0].row).to.equal(9);
    });
  });
});
