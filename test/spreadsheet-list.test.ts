import { Config } from '@oclif/core';
import { runCommand } from '@oclif/test';
import { expect } from 'chai';
import { generateKeyPairSync } from 'crypto';
import { join } from 'path';
import * as factory from '../src/lib/factory';
import GoogleSheet, { DRIVE_FILE_VISIBILITY_NOTE, GoogleSheetCli } from '../src/lib/google-sheet';
import { classifyError, GSheetErrorCode } from '../src/lib/cli-errors';
import { FakeSheets, SPREADSHEET_MIME_TYPE } from './fake-sheets';

/**
 * Scope-honest spreadsheet discovery: the listSpreadsheets library operation over Drive
 * files.list (query building, escaping, pagination, Drive error translation) and the
 * spreadsheet:list command layer on top of it (flag passthrough, --all aggregation, human
 * and --rawOutput rendering).
 */

const LIST_ID = 'list-instance-id';

const seededFiles: GoogleSheetCli.ListedSpreadsheet[] = [
  { id: 'list-1', name: 'Report 2026', mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-01T10:00:00.000Z' },
  { id: 'list-2', name: 'Q1 Report', mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-02T10:00:00.000Z' },
  { id: 'list-3', name: 'Other Notes', mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-03T10:00:00.000Z' },
  { id: 'list-4', name: "Bob's Plan", mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-04T10:00:00.000Z' },
  { id: 'list-5', name: 'Back\\slash Plan', mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-05T10:00:00.000Z' },
  // Same title as list-1 on purpose: a title is not a unique key.
  { id: 'list-6', name: 'Report 2026', mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-06T10:00:00.000Z' },
];

describe('spreadsheet discovery (Drive files.list)', () => {
  const fake = new FakeSheets();
  let gsheet: GoogleSheet;

  before(() => fake.install());
  after(() => fake.uninstall());

  beforeEach(async () => {
    fake.reset();
    for (const file of seededFiles) {
      fake.seedDriveFile(file);
    }
    gsheet = new GoogleSheet(LIST_ID);
    await gsheet.authorize(fake.credentials);
  });

  /** Only the files.list requests; the trailing `?` keeps them apart from /files/<id>/... endpoints. */
  const listRequests = (): { method: string; url: string }[] => fake.requests.filter((r) => r.url.includes('/drive/v3/files?'));
  const lastQuery = (): string => {
    const requests = listRequests();
    expect(requests, 'expected a files.list request').to.have.length.of.at.least(1);
    const q = new URL(requests[requests.length - 1].url).searchParams.get('q');
    expect(q, 'expected a q parameter').to.be.a('string');
    return q as string;
  };

  it('lists every visible spreadsheet with id, name, mimeType and modifiedTime', async () => {
    const result = await gsheet.listSpreadsheets();
    expect(result.files).to.eql(seededFiles);
    expect(result.visibilityNote).to.equal(DRIVE_FILE_VISIBILITY_NOTE);
    expect(result.nextPageToken).to.equal(undefined);
  });

  it('sends the documented request shape with the default page size', async () => {
    await gsheet.listSpreadsheets();
    const request = listRequests()[0];
    expect(request.method).to.equal('GET');
    const params = new URL(request.url).searchParams;
    expect(new URL(request.url).pathname).to.equal('/drive/v3/files');
    expect(params.get('spaces')).to.equal('drive');
    expect(params.get('fields')).to.equal('files(id,name,mimeType,modifiedTime),nextPageToken');
    expect(params.get('pageSize')).to.equal('50');
    expect(params.get('q')).to.equal(`mimeType='${SPREADSHEET_MIME_TYPE}'`);
    expect(params.get('pageToken')).to.equal(null);
  });

  it('filters by partial title with the contains operator', async () => {
    const result = await gsheet.listSpreadsheets({ name: 'Report' });
    expect(result.files.map((f) => f.id)).to.eql(['list-1', 'list-2', 'list-6']);
    expect(lastQuery()).to.equal(`mimeType='${SPREADSHEET_MIME_TYPE}' and name contains 'Report'`);
  });

  it('filters by whole title on exact', async () => {
    const result = await gsheet.listSpreadsheets({ name: 'Q1 Report', exact: true });
    expect(result.files.map((f) => f.id)).to.eql(['list-2']);
    expect(lastQuery()).to.equal(`mimeType='${SPREADSHEET_MIME_TYPE}' and name='Q1 Report'`);
  });

  it('keeps duplicate titles as separate rows', async () => {
    const result = await gsheet.listSpreadsheets({ name: 'Report 2026', exact: true });
    expect(result.files.map((f) => f.id)).to.eql(['list-1', 'list-6']);
    expect(result.files.map((f) => f.name)).to.eql(['Report 2026', 'Report 2026']);
  });

  it('escapes a single quote in the query and still matches the title', async () => {
    const result = await gsheet.listSpreadsheets({ name: "Bob's Plan", exact: true });
    expect(result.files.map((f) => f.id)).to.eql(['list-4']);
    expect(lastQuery()).to.equal(`mimeType='${SPREADSHEET_MIME_TYPE}' and name='Bob\\'s Plan'`);
  });

  it('escapes a backslash in the query and still matches the title', async () => {
    const result = await gsheet.listSpreadsheets({ name: 'Back\\slash Plan', exact: true });
    expect(result.files.map((f) => f.id)).to.eql(['list-5']);
    expect(lastQuery()).to.equal(`mimeType='${SPREADSHEET_MIME_TYPE}' and name='Back\\\\slash Plan'`);
  });

  it('paginates with opaque tokens until Drive stops handing them out', async () => {
    fake.reset();
    for (let i = 1; i <= 5; i++) {
      fake.seedDriveFile({ id: `page-${i}`, name: `Page ${i}` });
    }

    const first = await gsheet.listSpreadsheets({ pageSize: 2 });
    expect(first.files.map((f) => f.id)).to.eql(['page-1', 'page-2']);
    expect(first.nextPageToken).to.be.a('string').that.is.not.empty;

    const second = await gsheet.listSpreadsheets({ pageSize: 2, pageToken: first.nextPageToken });
    expect(second.files.map((f) => f.id)).to.eql(['page-3', 'page-4']);
    expect(second.nextPageToken).to.be.a('string').that.is.not.empty;

    const third = await gsheet.listSpreadsheets({ pageSize: 2, pageToken: second.nextPageToken });
    expect(third.files.map((f) => f.id)).to.eql(['page-5']);
    expect(third.nextPageToken).to.equal(undefined);
  });

  it('bounds the page size on both ends', async () => {
    await gsheet.listSpreadsheets({ pageSize: 1000 });
    expect(new URL(listRequests()[0].url).searchParams.get('pageSize')).to.equal('100');

    await gsheet.listSpreadsheets({ pageSize: 0 });
    expect(new URL(listRequests()[1].url).searchParams.get('pageSize')).to.equal('1');
  });

  it('returns an empty page with the visibility note for an unmatched title', async () => {
    const result = await gsheet.listSpreadsheets({ name: 'no-such-title-anywhere' });
    expect(result.files).to.eql([]);
    expect(result.nextPageToken).to.equal(undefined);
    expect(result.visibilityNote).to.contain('does not prove');
    expect(result.visibilityNote).to.contain('drive.file');
  });

  it('refuses an empty name instead of sending a bare clause', async () => {
    try {
      await gsheet.listSpreadsheets({ name: '' });
      expect.fail('should have thrown');
    } catch (error) {
      expect((error as Error).message).to.contain('non-empty');
    }
  });

  it('turns a Drive 403 into an actionable FORBIDDEN error naming drive.file', async () => {
    fake.failRequests('/drive/v3/files', 403);
    try {
      await gsheet.listSpreadsheets();
      expect.fail('should have thrown');
    } catch (error) {
      expect((error as Error).message).to.contain('drive.file');
      expect((error as Error).message).to.contain('auth:login');
      expect((error as { status?: number }).status).to.equal(403);
      const classified = classifyError(error);
      expect(classified.code).to.equal(GSheetErrorCode.FORBIDDEN);
      expect(classified.retryable).to.equal(false);
    }
  });

  it('turns a Drive 500 into a retryable UPSTREAM error', async () => {
    fake.failRequests('/drive/v3/files', 500);
    try {
      await gsheet.listSpreadsheets();
      expect.fail('should have thrown');
    } catch (error) {
      const classified = classifyError(error);
      expect(classified.code).to.equal(GSheetErrorCode.UPSTREAM);
      expect(classified.retryable).to.equal(true);
    }
  });
});

/**
 * Command layer: the factory seam is swapped for a stub that scripts pages, so the flag
 * plumbing and the --all aggregation loop are proven without any credential or network.
 */
const ROOT = join(__dirname, '..');

const CLIENT_EMAIL = 'list-command@example.iam.gserviceaccount.com';
// `normalizeCredentials` parses the key with `createPrivateKey`, so a placeholder would fail
// before the command is reached; this throwaway key exists only inside the test process. The
// key travels through the GSHEET_PRIVATE_KEY env binding, never the command line - the PEM
// header has spaces in it and would be torn apart by argv parsing.
const PRIVATE_KEY = generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { format: 'pem', type: 'pkcs8' },
  publicKeyEncoding: { format: 'pem', type: 'spki' },
}).privateKey;

const AUTH_FLAGS = [`--clientEmail=${CLIENT_EMAIL}`];

/** The env every auth flag can bind to, emptied so the master workflow's live credentials cannot leak in. */
const MANAGED_ENV = ['GSHEET_CLIENT_EMAIL', 'GSHEET_PRIVATE_KEY', 'GSHEET_CREDENTIALS_FILE'];

class StubListSheets {
  public readonly calls: { method: string; args: unknown[] }[] = [];
  public pages: GoogleSheetCli.ListSpreadsheetsResult[] = [];
  private pageIndex = 0;

  async authorize(credentials: GoogleSheetCli.Credentials): Promise<void> {
    this.calls.push({ method: 'authorize', args: [credentials] });
  }

  async listSpreadsheets(options: GoogleSheetCli.ListSpreadsheetsOptions = {}): Promise<GoogleSheetCli.ListSpreadsheetsResult> {
    this.calls.push({ method: 'listSpreadsheets', args: [options] });
    const page = this.pages[this.pageIndex++];
    if (!page) throw new Error(`Unexpected extra listSpreadsheets call (options: ${JSON.stringify(options)})`);
    return page;
  }
}

describe('spreadsheet:list command', () => {
  const mutableFactory = factory as { createGoogleSheet: () => GoogleSheet };
  const realCreateGoogleSheet = factory.createGoogleSheet;
  const savedNodeEnv = { value: undefined as string | undefined };
  let savedEnv: (readonly [string, string | undefined])[] = [];
  let stub: StubListSheets;

  const file = (id: string): GoogleSheetCli.ListedSpreadsheet => ({ id, name: `File ${id}`, mimeType: SPREADSHEET_MIME_TYPE, modifiedTime: '2026-09-20T00:00:00.000Z' });
  const page = (ids: string[], nextPageToken?: string): GoogleSheetCli.ListSpreadsheetsResult => ({
    files: ids.map(file),
    ...(nextPageToken ? { nextPageToken } : {}),
    visibilityNote: DRIVE_FILE_VISIBILITY_NOTE,
  });

  beforeEach(() => {
    savedNodeEnv.value = process.env.NODE_ENV;
    process.env.NODE_ENV = 'test';
    // The master workflow runs the whole suite with live credentials exported; those would sit
    // behind the env-bound auth flags, so the environment is emptied and the throwaway key
    // travels on the command line instead.
    savedEnv = MANAGED_ENV.map((name) => [name, process.env[name]] as const);
    for (const [name] of savedEnv) delete process.env[name];
    // the one value that cannot travel on the command line (the PEM header has spaces in it)
    process.env.GSHEET_PRIVATE_KEY = PRIVATE_KEY;

    stub = new StubListSheets();
    mutableFactory.createGoogleSheet = () => stub as unknown as GoogleSheet;
  });

  afterEach(() => {
    mutableFactory.createGoogleSheet = realCreateGoogleSheet;
    for (const [name, value] of savedEnv) {
      if (value === undefined) delete process.env[name];
      else process.env[name] = value;
    }
    if (savedNodeEnv.value === undefined) delete process.env.NODE_ENV;
    else process.env.NODE_ENV = savedNodeEnv.value;
  });

  it('aggregates every page under --all and re-feeds the continuation tokens', async () => {
    stub.pages = [page(['a', 'b'], 'token-1'), page(['c'], 'token-2'), page(['d'])];

    const { error, result, stdout } = await runCommand<GoogleSheetCli.ListSpreadsheetsResult>(['spreadsheet:list', '--all', '--rawOutput', ...AUTH_FLAGS]);
    if (error) throw error;
    if (!result) throw new Error('expected a command result');

    const listCalls = stub.calls.filter((c) => c.method === 'listSpreadsheets');
    expect(listCalls.map((c) => (c.args[0] as GoogleSheetCli.ListSpreadsheetsOptions).pageToken)).to.eql([undefined, 'token-1', 'token-2']);
    expect(result.files.map((f) => f.id)).to.eql(['a', 'b', 'c', 'd']);
    expect(result.nextPageToken).to.equal(undefined);
    expect(result.visibilityNote).to.contain('drive.file');
    expect(JSON.parse(stdout).files).to.have.length(4);
  });

  it('stops --all at the 1000-file safety bound and keeps the outstanding token', async () => {
    // 11 pages of 100 ids = 1100 files; the loop must stop at 1000 and hand the
    // resume token back instead of silently dropping the tail.
    stub.pages = Array.from({ length: 11 }, (_, i) =>
      page(Array.from({ length: 100 }, (_, j) => `f${i * 100 + j}`), `token-${i + 1}`)
    );

    const { error, result } = await runCommand<GoogleSheetCli.ListSpreadsheetsResult>(['spreadsheet:list', '--all', '--rawOutput', ...AUTH_FLAGS]);
    if (error) throw error;
    if (!result) throw new Error('expected a command result');

    expect(result.files).to.have.length(1000);
    expect(result.nextPageToken).to.equal('token-10');
    expect(stub.calls.filter((c) => c.method === 'listSpreadsheets')).to.have.length(10);
  });

  it('passes the discovery flags through on a single page', async () => {
    stub.pages = [page(['a'], 'more-token')];

    const { error, result } = await runCommand<GoogleSheetCli.ListSpreadsheetsResult>([
      'spreadsheet:list',
      '--name=rep',
      '--exact',
      '--pageSize=10',
      '--pageToken=resume-here',
      '--rawOutput',
      ...AUTH_FLAGS,
    ]);
    if (error) throw error;
    if (!result) throw new Error('expected a command result');

    expect(stub.calls.filter((c) => c.method === 'listSpreadsheets')).to.have.length(1);
    const options = stub.calls[1].args[0] as GoogleSheetCli.ListSpreadsheetsOptions;
    expect(options.name).to.equal('rep');
    expect(options.exact).to.equal(true);
    expect(options.pageSize).to.equal(10);
    expect(options.pageToken).to.equal('resume-here');
    expect(result.files.map((f) => f.id)).to.eql(['a']);
    expect(result.nextPageToken).to.equal('more-token');
  });

  it('renders rows, the continuation hint and the visibility note in human mode', async () => {
    stub.pages = [page(['human-1', 'human-2'], 'next-here')];

    const { error, stdout } = await runCommand<GoogleSheetCli.ListSpreadsheetsResult>(['spreadsheet:list', ...AUTH_FLAGS]);
    if (error) throw error;

    expect(stdout).to.contain('human-1');
    expect(stdout).to.contain('human-2');
    expect(stdout).to.contain('File human-1');
    expect(stdout).to.contain('More results available');
    expect(stdout).to.contain('--pageToken next-here');
    expect(stdout).to.contain('drive.file');
    expect(stdout).to.contain('does not prove');
  });
});
